import os
import re
import time
import html
from io import BytesIO
from urllib.parse import unquote, urlparse, urljoin

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup
from PIL import Image, UnidentifiedImageError
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from deep_translator import GoogleTranslator
    HAS_TRANSLATOR = True
except Exception:
    HAS_TRANSLATOR = False


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

VERIFY_SSL = False
if not VERIFY_SSL:
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SIZE_PATTERN = re.compile(r"\b(?:XS|S|M|L|XL|XXL|XXXL|2XL|3XL)\b", re.IGNORECASE)
SIZE_RANGE_PATTERN = re.compile(
    r"\b(?:XS|S|M|L|XL|XXL|XXXL|2XL|3XL)\s*-\s*(?:XS|S|M|L|XL|XXL|XXXL|2XL|3XL)\b",
    re.IGNORECASE
)

BAD_URL_HINTS = [
    "logo", "icon", "sprite", "placeholder", "avatar", "banner",
    "favicon", "thumb", "thumbnail", "default", "blank", "spacer"
]

NEGATIVE_TERMS = "-logo -icon -placeholder -banner -avatar -favicon"

MIN_WIDTH = 250
MIN_HEIGHT = 250
MAX_BING_RESULTS_PER_QUERY = 12
MAX_TOTAL_CANDIDATES = 20
MAX_CANDIDATES_TO_VALIDATE = 8
MAX_SITE_PAGES_TO_CHECK = 5
MAX_SITE_IMAGES_TO_VALIDATE = 12


def make_session():
    session = requests.Session()
    session.headers.update(HEADERS)

    retries = Retry(
        total=2,
        backoff_factor=0.4,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD"],
    )

    adapter = HTTPAdapter(
        max_retries=retries,
        pool_connections=20,
        pool_maxsize=20
    )

    session.mount("http://", adapter)
    session.mount("https://", adapter)

    return session


def normalize_spaces(text):
    return re.sub(r"\s+", " ", str(text)).strip()


def remove_sizes_only(text):
    t = str(text)
    t = SIZE_RANGE_PATTERN.sub("", t)
    t = SIZE_PATTERN.sub("", t)
    return normalize_spaces(t)


def clean_product_identity(text):
    t = remove_sizes_only(text)
    t = re.sub(r"[-,;/]+", " ", t)
    t = normalize_spaces(t)
    return t


def build_cache_key(query):
    return clean_product_identity(query).lower()


def normalize_url(url):
    if not url:
        return None

    url = html.unescape(str(url).strip())
    url = url.replace("\\/", "/")
    url = unquote(url)
    url = url.strip('\'" ,;<>')

    if not url.startswith(("http://", "https://")):
        return None

    return url


def normalize_site_filter(site_text):
    if not site_text:
        return ""

    site = str(site_text).strip().lower()
    site = re.sub(r"^https?://", "", site)
    site = re.sub(r"^www\.", "", site)
    site = site.strip("/")
    site = site.strip()
    site = site.split("/")[0]
    return site


def translate_query_if_needed(text, listing_language, translate_to_english):
    if not translate_to_english:
        return None

    lang_map = {
        "english": "en",
        "spanish": "es",
        "italian": "it",
    }

    source_lang = lang_map.get(str(listing_language).lower(), "en")
    if source_lang == "en":
        return None

    if not HAS_TRANSLATOR:
        return None

    try:
        translated = GoogleTranslator(source=source_lang, target="en").translate(text)
        return normalize_spaces(translated)
    except Exception:
        return None


def build_query_variants(
    query,
    listing_language="english",
    translate_to_english=False,
    brand_filter=""
):
    exact_query = normalize_spaces(query)
    no_size_query = remove_sizes_only(query)
    cleaned_query = clean_product_identity(query)
    translated_query = translate_query_if_needed(cleaned_query, listing_language, translate_to_english)

    brand_filter = normalize_spaces(brand_filter)

    base_queries = [exact_query, no_size_query, cleaned_query]
    if translated_query:
        base_queries.append(translated_query)

    deduped_base = []
    for q in base_queries:
        q = normalize_spaces(q)
        if q and q not in deduped_base:
            deduped_base.append(q)

    variants = []

    for q in deduped_base:
        local_variants = [
            q,
            f'"{q}"',
            f"{q} {NEGATIVE_TERMS}",
        ]

        if brand_filter:
            local_variants.extend([
                f"{brand_filter} {q}",
                f'"{brand_filter} {q}"',
                f"{brand_filter} {q} {NEGATIVE_TERMS}",
            ])

        for candidate in local_variants:
            candidate = normalize_spaces(candidate)
            if candidate and candidate not in variants:
                variants.append(candidate)

    return variants[:10]


def search_bing_images(query, session, max_results=12):
    try:
        r = session.get(
            "https://www.bing.com/images/search",
            params={"q": query},
            timeout=20,
            allow_redirects=True,
            headers=HEADERS,
            verify=VERIFY_SSL,
        )

        if r.status_code != 200:
            return []

        matches = re.findall(r'murl&quot;:&quot;(.*?)&quot;', r.text)
        if not matches:
            matches = re.findall(r'"murl":"(.*?)"', r.text)

        urls = []
        seen = set()

        for m in matches:
            u = normalize_url(m)
            if u and u not in seen:
                seen.add(u)
                urls.append(u)
            if len(urls) >= max_results:
                break

        return urls

    except Exception:
        return []


def search_bing_web_pages(query, session, max_results=5):
    """
    Bing normal web search, used to find product pages on a specific site.
    """
    try:
        r = session.get(
            "https://www.bing.com/search",
            params={"q": query},
            timeout=20,
            allow_redirects=True,
            headers=HEADERS,
            verify=VERIFY_SSL,
        )

        if r.status_code != 200:
            return []

        page_urls = []
        seen = set()

        # extract normal result urls
        matches = re.findall(r'<a href="(https?://[^"]+)"', r.text)
        for url in matches:
            u = normalize_url(url)
            if not u:
                continue

            # skip Bing internal links / cache / misc junk
            if "bing.com" in urlparse(u).netloc.lower():
                continue

            if u not in seen:
                seen.add(u)
                page_urls.append(u)

            if len(page_urls) >= max_results:
                break

        return page_urls

    except Exception:
        return []


def is_url_on_domain(url, domain):
    try:
        netloc = urlparse(url).netloc.lower()
        domain = normalize_site_filter(domain)
        return domain in netloc
    except Exception:
        return False


def extract_images_from_html(page_url, html_text):
    soup = BeautifulSoup(html_text, "html.parser")
    image_urls = []

    def add_url(raw):
        if not raw:
            return
        abs_url = urljoin(page_url, raw)
        abs_url = normalize_url(abs_url)
        if abs_url and abs_url not in image_urls:
            image_urls.append(abs_url)

    # highest priority: og:image / twitter:image
    for tag in soup.find_all("meta"):
        prop = (tag.get("property") or "").lower()
        name = (tag.get("name") or "").lower()
        content = tag.get("content")

        if prop in {"og:image", "og:image:url", "og:image:secure_url"}:
            add_url(content)
        if name in {"twitter:image", "twitter:image:src"}:
            add_url(content)

    # link rel preload/image_src
    for link in soup.find_all("link"):
        rel = " ".join(link.get("rel", [])).lower()
        href = link.get("href")
        if "image_src" in rel or "preload" in rel:
            add_url(href)

    # normal img tags
    for img in soup.find_all("img"):
        for attr in ["src", "data-src", "data-lazy-src", "data-original", "srcset"]:
            value = img.get(attr)
            if not value:
                continue

            if attr == "srcset":
                first = value.split(",")[0].strip().split(" ")[0]
                add_url(first)
            else:
                add_url(value)

    return image_urls


def search_site_product_images(
    query,
    site_filter,
    session,
    listing_language="english",
    translate_to_english=False,
    brand_filter="",
    log_func=None
):
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    domain = normalize_site_filter(site_filter)
    if not domain:
        return []

    query_variants = build_query_variants(
        query=query,
        listing_language=listing_language,
        translate_to_english=translate_to_english,
        brand_filter=brand_filter
    )

    all_page_urls = []
    seen_pages = set()

    # 1) find likely product pages on that domain
    for q in query_variants[:5]:
        page_query = f"site:{domain} {q}"
        log(f"  [Site pages] Trying query: {page_query}")

        page_urls = search_bing_web_pages(page_query, session=session, max_results=MAX_SITE_PAGES_TO_CHECK)

        if not page_urls:
            log("  [Site pages] No page results")
            continue

        for page_url in page_urls:
            if is_url_on_domain(page_url, domain) and page_url not in seen_pages:
                seen_pages.add(page_url)
                all_page_urls.append(page_url)

        if len(all_page_urls) >= MAX_SITE_PAGES_TO_CHECK:
            break

        time.sleep(0.15)

    if not all_page_urls:
        return []

    # 2) scrape images from those pages
    image_candidates = []
    seen_images = set()

    for idx, page_url in enumerate(all_page_urls[:MAX_SITE_PAGES_TO_CHECK], start=1):
        log(f"  [Site scrape] Checking page {idx}: {page_url}")

        try:
            r = session.get(
                page_url,
                timeout=20,
                allow_redirects=True,
                headers=HEADERS,
                verify=VERIFY_SSL
            )

            if r.status_code != 200:
                log(f"    Page HTTP {r.status_code}")
                continue

            extracted = extract_images_from_html(page_url, r.text)
            log(f"    Extracted {len(extracted)} image candidate(s) from page")

            for img_url in extracted:
                if not is_url_on_domain(img_url, domain):
                    continue
                if img_url not in seen_images:
                    seen_images.add(img_url)
                    image_candidates.append(img_url)

            if len(image_candidates) >= MAX_SITE_IMAGES_TO_VALIDATE:
                break

        except Exception as e:
            log(f"    Site scrape error: {e}")

        time.sleep(0.10)

    return image_candidates[:MAX_SITE_IMAGES_TO_VALIDATE]


def find_candidate_urls_bing(
    query,
    session,
    listing_language="english",
    translate_to_english=False,
    brand_filter="",
    log_func=None
):
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    all_candidates = []
    seen = set()

    variants = build_query_variants(
        query=query,
        listing_language=listing_language,
        translate_to_english=translate_to_english,
        brand_filter=brand_filter
    )

    for q in variants:
        log(f"  [Bing images] Trying query: {q}")

        candidates = search_bing_images(
            query=q,
            session=session,
            max_results=MAX_BING_RESULTS_PER_QUERY
        )

        if not candidates:
            log("  [Bing images] No results")
            continue

        for idx, url in enumerate(candidates, start=1):
            short_url = url[:120] + ("..." if len(url) > 120 else "")
            log(f"    Candidate {idx}: {short_url}")

            if url not in seen:
                seen.add(url)
                all_candidates.append(url)

            if len(all_candidates) >= MAX_TOTAL_CANDIDATES:
                return all_candidates[:MAX_TOTAL_CANDIDATES]

        time.sleep(0.15)

    return all_candidates[:MAX_TOTAL_CANDIDATES]


def build_request_headers_for_url(url):
    headers = dict(HEADERS)
    try:
        parsed = urlparse(url)
        origin = f"{parsed.scheme}://{parsed.netloc}"
        headers["Referer"] = origin + "/"
    except Exception:
        pass
    return headers


def tokenise_query(text):
    text = clean_product_identity(text).lower()
    parts = re.findall(r"[a-z0-9]+", text)
    return [p for p in parts if len(p) >= 3]


def score_candidate(url, width, height, query, brand_filter="", site_filter=""):
    score = 0
    url_l = url.lower()

    if width >= 500 and height >= 500:
        score += 40
    elif width >= 350 and height >= 350:
        score += 20
    else:
        score -= 50

    ratio = width / max(height, 1)
    if ratio > 3.0 or ratio < 0.25:
        score -= 35
    elif 0.5 <= ratio <= 1.8:
        score += 10

    for hint in BAD_URL_HINTS:
        if hint in url_l:
            score -= 40

    query_tokens = tokenise_query(query)
    for token in query_tokens:
        if token in url_l:
            score += 10

    if brand_filter:
        brand_tokens = tokenise_query(brand_filter)
        for token in brand_tokens:
            if token in url_l:
                score += 14

    normalized_site = normalize_site_filter(site_filter)
    if normalized_site and normalized_site in url_l:
        score += 30

    return score


def fetch_image_and_score(url, session, query, brand_filter="", site_filter="", log_func=None):
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    try:
        headers = build_request_headers_for_url(url)
        r = session.get(
            url,
            timeout=25,
            allow_redirects=True,
            headers=headers,
            verify=VERIFY_SSL
        )

        ctype = (r.headers.get("Content-Type") or "").lower()
        log(f"      HTTP {r.status_code} | content-type={r.headers.get('Content-Type')}")

        if r.status_code != 200:
            return None

        if "image" not in ctype and "octet-stream" not in ctype:
            return None

        img = Image.open(BytesIO(r.content))
        img.load()

        width, height = img.size
        if width < MIN_WIDTH or height < MIN_HEIGHT:
            log(f"      Rejected: too small ({width}x{height})")
            return None

        img = img.convert("RGB")
        png_buffer = BytesIO()
        img.save(png_buffer, format="PNG")
        png_bytes = png_buffer.getvalue()

        score = score_candidate(url, width, height, query, brand_filter=brand_filter, site_filter=site_filter)

        return {
            "url": url,
            "png_bytes": png_bytes,
            "width": width,
            "height": height,
            "score": score,
        }

    except UnidentifiedImageError as e:
        log(f"      PIL could not identify image: {e}")
        return None
    except requests.exceptions.SSLError as e:
        log(f"      SSL error: {e}")
        return None
    except requests.exceptions.RequestException as e:
        log(f"      Request error: {e}")
        return None
    except Exception as e:
        log(f"      Other error: {type(e).__name__}: {e}")
        return None


def choose_best_candidate(candidates, query, session, brand_filter="", site_filter="", log_func=None):
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    scored = []

    for idx, url in enumerate(candidates[:MAX_CANDIDATES_TO_VALIDATE], start=1):
        short_url = url[:120] + ("..." if len(url) > 120 else "")
        log(f"    Checking candidate {idx}: {short_url}")

        result = fetch_image_and_score(
            url=url,
            session=session,
            query=query,
            brand_filter=brand_filter,
            site_filter=site_filter,
            log_func=log
        )

        if result:
            log(
                f"      Accepted for scoring | "
                f"score={result['score']} | "
                f"size={result['width']}x{result['height']}"
            )
            scored.append(result)
        else:
            log("      Rejected")

    if not scored:
        return None

    scored.sort(key=lambda x: (x["score"], x["width"] * x["height"]), reverse=True)
    return scored[0]


def resize_image_keep_ratio(img_path, max_w=144, max_h=135):
    with Image.open(img_path) as img:
        w, h = img.size
        ratio = min(max_w / w, max_h / h)

        if ratio < 1:
            new_w = max(1, int(w * ratio))
            new_h = max(1, int(h * ratio))
            img = img.resize((new_w, new_h))

        img.save(img_path)


def write_excel_with_embedded_images(excel_path, rows_data, log_func=None):
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    CELL_WIDTH_PX = 160
    CELL_HEIGHT_PX = 150
    IMG_WIDTH_PX = int(CELL_WIDTH_PX * 0.9)
    IMG_HEIGHT_PX = int(CELL_HEIGHT_PX * 0.9)

    ws["A1"] = "Reference"
    ws["B1"] = "Link"
    ws["C1"] = "Image"

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 23

    for i, row in enumerate(rows_data, start=2):
        ref = row.get("Reference", "")
        chosen_url = row.get("ChosenURL", "")
        local_img_path = row.get("LocalImagePath", "")

        ws[f"A{i}"] = ref
        ws[f"B{i}"] = chosen_url

        ws[f"A{i}"].alignment = Alignment(vertical="top")
        ws[f"B{i}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws[f"C{i}"].alignment = Alignment(vertical="top")
        ws.row_dimensions[i].height = CELL_HEIGHT_PX * 0.75

        if not local_img_path or not os.path.exists(local_img_path):
            continue

        try:
            resize_image_keep_ratio(local_img_path, max_w=IMG_WIDTH_PX, max_h=IMG_HEIGHT_PX)

            with Image.open(local_img_path) as pil_img:
                actual_w, actual_h = pil_img.size

            xl_img = XLImage(local_img_path)
            xl_img.width = actual_w
            xl_img.height = actual_h

            offset_x = max(0, int((CELL_WIDTH_PX - actual_w) / 2))
            offset_y = max(0, int((CELL_HEIGHT_PX - actual_h) / 2))

            marker = AnchorMarker(
                col=2,
                colOff=offset_x * 9525,
                row=i - 1,
                rowOff=offset_y * 9525
            )

            size = XDRPositiveSize2D(
                cx=actual_w * 9525,
                cy=actual_h * 9525
            )

            xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
            ws.add_image(xl_img)
            log(f"Embedded image in row {i}")

        except Exception as e:
            log(f"Could not embed image in row {i}: {e}")

    wb.save(excel_path)


def process_excel(
    input_file,
    output_dir,
    listing_language="english",
    translate_to_english=False,
    brand_filter="",
    site_filter="",
    log_func=None
):
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_excel = os.path.join(output_dir, f"{base_name}_with_images.xlsx")
    images_dir = os.path.join(output_dir, f"{base_name}_images")
    os.makedirs(images_dir, exist_ok=True)

    log("Reading Excel file...")
    df_raw = pd.read_excel(input_file, header=None)

    if df_raw.empty:
        raise ValueError("The Excel file is empty.")

    references = df_raw.iloc[:, 0].fillna("").astype(str).str.strip()
    references = references[references != ""]

    if references.empty:
        raise ValueError("Column A does not contain any valid values.")

    refs_list = references.tolist()
    total = len(refs_list)

    brand_filter = normalize_spaces(brand_filter)
    site_filter = normalize_site_filter(site_filter)

    log(f"Processing {total} article(s)...")

    session = make_session()
    success_cache = {}
    rows_data = []

    for i, ref in enumerate(refs_list, start=1):
        log(f"[{i}/{total}] Searching: {ref}")

        cache_key = build_cache_key(f"{ref} | brand={brand_filter} | site={site_filter}")

        if cache_key in success_cache:
            cached = success_cache[cache_key]
            log(f"  Reused successful cached image for: {cache_key}")
            rows_data.append({
                "Reference": ref,
                "ChosenURL": cached["url"],
                "LocalImagePath": cached["local_path"],
            })
            continue

        chosen_url = ""
        local_img_path = ""

        # 1) Site scraping first, if site provided
        site_candidates = []
        if site_filter:
            log(f"  Site filter detected -> trying site scraping first: {site_filter}")
            site_candidates = search_site_product_images(
                query=ref,
                site_filter=site_filter,
                session=session,
                listing_language=listing_language,
                translate_to_english=translate_to_english,
                brand_filter=brand_filter,
                log_func=log
            )

        candidates = []
        if site_candidates:
            log(f"  Found {len(site_candidates)} site image candidate(s)")
            candidates = site_candidates
        else:
            if site_filter:
                log("  Site scraping found nothing usable -> fallback to Bing images")

            candidates = find_candidate_urls_bing(
                query=ref,
                session=session,
                listing_language=listing_language,
                translate_to_english=translate_to_english,
                brand_filter=brand_filter,
                log_func=log
            )

        if candidates:
            log("  Validating and scoring candidates...")
            best = choose_best_candidate(
                candidates=candidates,
                query=ref,
                session=session,
                brand_filter=brand_filter,
                site_filter=site_filter,
                log_func=log
            )

            if best:
                chosen_url = best["url"]
                local_img_path = os.path.join(images_dir, f"row_{i}.png")

                with open(local_img_path, "wb") as f:
                    f.write(best["png_bytes"])

                log(
                    f"  Accepted best candidate | "
                    f"score={best['score']} | "
                    f"size={best['width']}x{best['height']}"
                )

                success_cache[cache_key] = {
                    "url": chosen_url,
                    "local_path": local_img_path,
                }
            else:
                log("  No valid image after scoring")
        else:
            log("  No candidates found")

        rows_data.append({
            "Reference": ref,
            "ChosenURL": chosen_url,
            "LocalImagePath": local_img_path,
        })

        time.sleep(0.10)

    log("Writing Excel file and embedding images...")
    write_excel_with_embedded_images(output_excel, rows_data, log_func=log)

    log(f"Done: {output_excel}")
    return output_excel
