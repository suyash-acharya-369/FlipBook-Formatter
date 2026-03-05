import os
import uuid
import re
import html as html_mod
from glob import glob
from flask import Flask, request, send_file, render_template, jsonify, Response
from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException
from flask_cors import CORS
from docx import Document
from docx.oxml.ns import qn

# Import our formatting logic
from formatter import format_document

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__name__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__name__)), 'outputs')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.errorhandler(Exception)
def handle_exception(e):
    if isinstance(e, HTTPException):
        return jsonify({'error': e.description}), e.code
    return jsonify({'error': str(e)}), 500


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request.'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected for uploading.'}), 400

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            job_id = str(uuid.uuid4())

            input_filename = f"{job_id}_{filename}"
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
            file.save(input_path)

            output_filename = f"Formatted_{filename}"
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{job_id}_{output_filename}")

            format_document(input_path, output_path)

            try:
                os.remove(input_path)
            except:
                pass

            return jsonify({
                'success': True,
                'job_id': job_id,
                'filename': output_filename,
                'message': 'File formatted successfully!'
            })
        else:
            return jsonify({'error': 'Allowed file type is .docx only.'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/download/<job_id>/<filename>')
def download_file(job_id, filename):
    target_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{job_id}_{filename}")
    if os.path.exists(target_path):
        return send_file(target_path, as_attachment=True, download_name=filename)
    else:
        return jsonify({'error': 'File not found or has expired.'}), 404


# ────────────────────────────────────────────
#  PREVIEW endpoint – renders DOCX → HTML
# ────────────────────────────────────────────
def _para_to_html(para):
    """Convert a python-docx paragraph to a minimal, safe HTML snippet."""
    style_name = para.style.name if para.style else ''

    # Collect runs with inline bold/italic
    inner = []
    for run in para.runs:
        txt = html_mod.escape(run.text)
        if not txt:
            continue
        if run.bold and run.italic:
            txt = f'<strong><em>{txt}</em></strong>'
        elif run.bold:
            txt = f'<strong>{txt}</strong>'
        elif run.italic:
            txt = f'<em>{txt}</em>'
        inner.append(txt)
    content = ''.join(inner) or html_mod.escape(para.text)

    if not content.strip():
        return '<p class="pv-blank">&nbsp;</p>'

    # Detect heading level
    h_map = {
        'Heading 1': 'h1', 'Heading 2': 'h2', 'Heading 3': 'h3', 'Heading 4': 'h4'
    }
    if style_name in h_map:
        tag = h_map[style_name]
        return f'<{tag} class="pv-{tag}">{content}</{tag}>'

    # Bullet list
    numPr = para._element.find(qn('w:pPr') + '/' + qn('w:numPr'))
    if numPr is not None:
        return f'<li class="pv-li">{content}</li>'

    # Caption
    if 'caption' in style_name.lower():
        return f'<p class="pv-caption">{content}</p>'

    return f'<p class="pv-p">{content}</p>'


def _docx_to_preview_html(docx_path, max_chars=4000):
    """Extract the first ~4000 chars of content from a DOCX as styled HTML."""
    doc = Document(docx_path)
    chunks = []
    total = 0
    in_list = False

    for para in doc.paragraphs:
        html_bit = _para_to_html(para)
        is_li = html_bit.startswith('<li')

        if is_li and not in_list:
            chunks.append('<ul class="pv-ul">')
            in_list = True
        elif not is_li and in_list:
            chunks.append('</ul>')
            in_list = False

        chunks.append(html_bit)
        total += len(para.text)
        if total > max_chars:
            break

    if in_list:
        chunks.append('</ul>')

    return '\n'.join(chunks)


@app.route('/preview/<job_id>/<filename>')
def preview_file(job_id, filename):
    target_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{job_id}_{filename}")
    if not os.path.exists(target_path):
        return jsonify({'error': 'Preview not available.'}), 404

    try:
        body_html = _docx_to_preview_html(target_path)
        return jsonify({'success': True, 'html': body_html})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5055)
