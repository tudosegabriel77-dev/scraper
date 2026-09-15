import os
import tempfile
from flask import Flask, request, send_file, render_template_string
from waitress import serve
from image_finder import process_excel

app = Flask(__name__)

HTML_FORM = """
<!doctype html>
<html>
<head><title>Image Finder</title></head>
<body>
  <h1>Upload Excel File</h1>
  <form action="/process" method="post" enctype="multipart/form-data">
    <input type="file" name="file" accept=".xlsx,.xls" required>
    <input type="submit" value="Process">
  </form>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_FORM)

@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    if file:
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, file.filename)
        output_dir = os.path.join(temp_dir, "output")
        os.makedirs(output_dir, exist_ok=True)
        file.save(input_path)

        try:
            output_excel = process_excel(input_path, output_dir, log_func=print)
            return send_file(
                output_excel,
                as_attachment=True,
                download_name=os.path.basename(output_excel)
            )
        except Exception as e:
            return f"Error: {str(e)}", 500

if __name__ == '__main__':
    # For local development use app.run(debug=True)
    # For Render, use Waitress
    serve(app, host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))
