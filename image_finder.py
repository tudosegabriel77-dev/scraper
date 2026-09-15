import os
import re
import time
import html
from io import BytesIO
from urllib.parse import unquote, urlparse

import pandas as pd
import requests
import urllib3

from PIL import Image, UnidentifiedImageError

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# ============================================================
# OPTIONAL TRANSLATION
# ============================================================

try:
    from deep_translator import GoogleTranslator
    TRANSLATION_AVAILABLE = True
except ImportError:
    TRANSLATION_AVAILABLE = False


LANGUAGE_CODE_MAP = {
    "english": "en",
    "german": "de",
    "french": "fr",
    "spanish": "es",
    "italian": "it",
    "portuguese": "pt",
    "dutch": "nl",
    "polish": "pl",
    "russian": "ru",
    "turkish": "tr",
    "czech": "cs",
    "swedish": "sv",
    "danish": "da",
    "finnish": "fi",
    "norwegian": "no",
    "greek": "el",
    "hungarian": "hu",
    "romanian": "ro",
    "ukrainian": "uk",
    "japanese": "ja",
    "chinese": "zh-CN",
    "korean": "ko",
    "arabic": "ar",
}


# ============================================================
# SETTINGS
# ============================================================

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
    urllib3.disable_warnings(
        urllib3.exceptions.InsecureRequestWarning
    )


SIZE_PATTERN = re.compile(
    r"\b(?:XS|S|M|L|XL|XXL|XXXL|2XL|3XL)\b",
    re.IGNORECASE
)

SIZE_RANGE_PATTERN = re.compile(
    r"\b(?:XS|S|M|L|XL|XXL|XXXL|2XL|3XL)"
    r"\s*-\s*"
    r"(?:XS|S|M|L|XL|XXL|XXXL|2XL|3XL)\b",
    re.IGNORECASE
)


# ============================================================
# SESSION
# ============================================================

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


# ============================================================
# TEXT / QUERY HELPERS
# ============================================================

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


def build_query_variants(query):
    exact_query = normalize_spaces(query)
    no_size_query = remove_sizes_only(query)
    cleaned_query = clean_product_identity(query)

    variants = []

    for q in [exact_query, no_size_query, cleaned_query]:
        q = normalize_spaces(q)

        if q and q not in variants:
            variants.append(q)

    return variants


# ============================================================
# TRANSLATION
# ============================================================

def translate_text_to_english(
    text,
    source_language="auto",
    log_func=None
):
    def log(msg):
        if log_func:
            log_func(msg)

    if not TRANSLATION_AVAILABLE:
        log(
            " Translation skipped: "
            "deep-translator is not installed."
        )
        return text

    text = normalize_spaces(text)

    if not text:
        return text

    src = LANGUAGE_CODE_MAP.get(
        str(source_language).lower(),
        "auto"
    )

    if src == "en":
        return text

    try:
        translated = GoogleTranslator(
            source=src,
            target="en"
        ).translate(text)

        if translated and translated.strip():
            return translated.strip()

    except Exception as e:
        log(f" Translation failed: {e}")

    return text


# ============================================================
# URL HELPERS
# ============================================================

def normalize_url(url):
    if not url:
        return None

    url = html.unescape(str(url).strip())

    url = url.replace("\\/", "/")

    url = unquote(url)

    url = url.strip("'\" ,;<>")

    if not url.startswith(("http://", "https://")):
        return None

    return url


def build_request_headers_for_url(url):
    headers = dict(HEADERS)

    try:
        parsed = urlparse(url)

        origin = f"{parsed.scheme}://{parsed.netloc}"

        headers["Referer"] = origin + "/"

    except Exception:
        pass

    return headers


# ============================================================
# QUERY BUILDING (brand / site filters)
# ============================================================

def apply_filters_to_query(
    query,
    brand_filter="",
    site_filter=""
):
    q = normalize_spaces(query)

    brand_filter = normalize_spaces(brand_filter)
    site_filter = normalize_spaces(site_filter)

    if brand_filter and brand_filter.lower() not in q.lower():
        q = f"{brand_filter} {q}"

    if site_filter:
        # Accept either "example.com" or "site:example.com"
        site_term = (
            site_filter
            if site_filter.lower().startswith("site:")
            else f"site:{site_filter}"
        )

        if site_term.lower() not in q.lower():
            q = f"{q} {site_term}"

    return normalize_spaces(q)


# ============================================================
# BING IMAGE SEARCH
# ============================================================

def search_bing_images(
    query,
    session,
    max_results=12
):
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

        matches = re.findall(
            r'murl&quot;:&quot;(.*?)&quot;',
            r.text
        )

        if not matches:
            matches = re.findall(
                r'"murl":"(.*?)"',
                r.text
            )

        urls = []

        for m in matches:
            u = normalize_url(m)

            if u and u not in urls:
                urls.append(u)

            if len(urls) >= max_results:
                break

        return urls

    except Exception:
        return []


def find_candidate_urls_bing(
    query,
    session,
    brand_filter="",
    site_filter="",
    log_func=None
):
    def log(msg):
        if log_func:
            log_func(msg)
        else:
            print(msg)

    all_candidates = []

    variants = build_query_variants(query)

    for q in variants:

        q = apply_filters_to_query(
            q,
            brand_filter=brand_filter,
            site_filter=site_filter
        )

        log(f" Trying query: {q}")

        candidates = search_bing_images(
            q,
            session=session,
            max_results=12
        )

        if not candidates:
            log(" No results")
            continue

        for idx, url in enumerate(candidates, start=1):

            short_url = (
                url[:120]
                + ("..." if len(url) > 120 else "")
            )

            log(f" Candidate {idx}: {short_url}")

            if url not in all_candidates:
                all_candidates.append(url)

    return all_candidates


# ============================================================
# IMAGE DOWNLOAD
# ============================================================

def download_image_to_png(
    url,
    out_path,
    session,
    log_func=None
):
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

        log(
            f" HTTP {r.status_code} | "
            f"content-type={r.headers.get('Content-Type')}"
        )

        if r.status_code != 200:
            return False

        img = Image.open(BytesIO(r.content))
        img.load()
        img = img.convert("RGB")
        img.save(out_path, format="PNG")

        return True

    except UnidentifiedImageError as e:
        log(f" PIL could not identify image: {e}")
        return False

    except requests.exceptions.SSLError as e:
        log(f" SSL error: {e}")
        return False

    except requests.exceptions.RequestException as e:
        log(f" Request error: {e}")
        return False

    except Exception as e:
        log(f" Other error: {type(e).__name__}: {e}")
        return False


# ============================================================
# IMAGE RESIZE
# ============================================================

def resize_image_keep_ratio(
    img_path,
    max_w=144,
    max_h=135
):
    with Image.open(img_path) as img:

        w, h = img.size

        ratio = min(max_w / w, max_h / h)

        if ratio < 1:

            new_w = max(1, int(w * ratio))
            new_h = max(1, int(h * ratio))

            img = img.resize((new_w, new_h))
            img.save(img_path)


# ============================================================
# WRITE EXCEL
# ============================================================

def write_excel_with_embedded_images(
    excel_path,
    rows_data,
    log_func=None
):
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
        ws[f"B{i}"].alignment = Alignment(
            wrap_text=True,
            vertical="top"
        )
        ws[f"C{i}"].alignment = Alignment(vertical="top")

        ws.row_dimensions[i].height = CELL_HEIGHT_PX * 0.75

        if (
            not local_img_path
            or not os.path.exists(local_img_path)
        ):
            continue

        try:

            resize_image_keep_ratio(
                local_img_path,
                max_w=IMG_WIDTH_PX,
                max_h=IMG_HEIGHT_PX
            )

            with Image.open(local_img_path) as pil_img:
                actual_w, actual_h = pil_img.size

            xl_img = XLImage(local_img_path)
            xl_img.width = actual_w
            xl_img.height = actual_h

            offset_x = max(
                0,
                int((CELL_WIDTH_PX - actual_w) / 2)
            )
            offset_y = max(
                0,
                int((CELL_HEIGHT_PX - actual_h) / 2)
            )

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

            xl_img.anchor = OneCellAnchor(
                _from=marker,
                ext=size
            )

            ws.add_image(xl_img)

            log(f"Embedded image in row {i}")

        except Exception as e:
            log(f"Could not embed image in row {i}: {e}")

    wb.save(excel_path)


# ============================================================
# MAIN EXCEL PROCESSING
# ============================================================

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
        raise FileNotFoundError(
            f"Input file not found: {input_file}"
        )

    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(
        os.path.basename(input_file)
    )[0]

    output_excel = os.path.join(
        output_dir,
        f"{base_name}_with_images.xlsx"
    )

    images_dir = os.path.join(
        output_dir,
        f"{base_name}_images"
    )

    os.makedirs(images_dir, exist_ok=True)

    log("Reading Excel file...")

    df_raw = pd.read_excel(input_file, header=None)

    if df_raw.empty:
        raise ValueError("The Excel file is empty.")

    references = (
        df_raw.iloc[:, 0]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    references = references[references != ""]

    if references.empty:
        raise ValueError(
            "Column A does not contain any valid values."
        )

    refs_list = references.tolist()
    total = len(refs_list)

    log(f"Processing {total} article(s)...")

    session = make_session()
    success_cache = {}
    rows_data = []

    for i, ref in enumerate(refs_list, start=1):

        log(f"[{i}/{total}] Searching: {ref}")

        # ----------------------------------------------------
        # Optional translation
        # ----------------------------------------------------
        search_reference = ref

        if (
            translate_to_english
            and listing_language
            and listing_language.lower() != "english"
        ):
            log(
                f" Translating from "
                f"{listing_language} to English..."
            )

            translated = translate_text_to_english(
                ref,
                source_language=listing_language,
                log_func=log
            )

            if translated and translated != ref:
                log(f" Translated -> {translated}")
                search_reference = translated
            else:
                log(" No translation applied.")

        cache_key = build_cache_key(search_reference)

        if cache_key in success_cache:

            cached = success_cache[cache_key]

            log(
                f" Reused successful cached "
                f"image for: {cache_key}"
            )

            rows_data.append({
                "Reference": ref,
                "ChosenURL": cached["url"],
                "LocalImagePath": cached["local_path"],
            })

            continue

        candidates = find_candidate_urls_bing(
            search_reference,
            session=session,
            brand_filter=brand_filter,
            site_filter=site_filter,
            log_func=log
        )

        chosen_url = ""
        local_img_path = ""

        if candidates:

            log(" Trying candidate downloads...")

            for idx, url in enumerate(candidates, start=1):

                short_url = (
                    url[:120]
                    + ("..." if len(url) > 120 else "")
                )

                log(
                    f" Download candidate {idx}: "
                    f"{short_url}"
                )

                temp_img_path = os.path.join(
                    images_dir,
                    f"row_{i}.png"
                )

                if download_image_to_png(
                    url,
                    temp_img_path,
                    session,
                    log_func=log
                ):
                    chosen_url = url
                    local_img_path = temp_img_path

                    log(" Accepted and downloaded")

                    success_cache[cache_key] = {
                        "url": chosen_url,
                        "local_path": local_img_path,
                    }

                    break

                else:
                    log(" Failed")

        else:
            log(" No candidates found")

        rows_data.append({
            "Reference": ref,
            "ChosenURL": chosen_url,
            "LocalImagePath": local_img_path,
        })

        time.sleep(0.05)

    log("Writing Excel file and embedding images...")

    write_excel_with_embedded_images(
        output_excel,
        rows_data,
        log_func=log
    )

    log(f"Done: {output_excel}")

    return output_excel
