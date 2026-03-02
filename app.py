import os
import uuid
import threading
from glob import glob
from flask import Flask, request, send_file, render_template, jsonify
from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException

# Import our formatting logic
from formatter import format_document

app = Flask(__name__)

# Configure upload and output directories
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__name__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__name__)), 'outputs')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB max size

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.errorhandler(Exception)
def handle_exception(e):
    # Pass through HTTP errors as JSON
    if isinstance(e, HTTPException):
        return jsonify({'error': e.description}), e.code
    # Return 500 for unhandled exceptions
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
    # Find the output file
    target_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{job_id}_{filename}")
    if os.path.exists(target_path):
        return send_file(target_path, as_attachment=True, download_name=filename)
    else:
        return jsonify({'error': 'File not found or has expired.'}), 404

# Background task to clean up old files after some time could be added here
# For now, it's a simple local app.

if __name__ == '__main__':
    app.run(debug=True, port=5000)
