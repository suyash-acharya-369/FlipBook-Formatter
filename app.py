import os
import uuid
import threading
from glob import glob
from flask import Flask, request, send_file, render_template, jsonify
from werkzeug.utils import secure_filename

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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in the request.'}), 400
        
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected for uploading.'}), 400
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        
        # Create a unique job id for this conversion
        job_id = str(uuid.uuid4())
        
        # Save original file
        input_filename = f"{job_id}_{filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(input_path)
        
        # Define output path
        output_filename = f"Formatted_{filename}"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{job_id}_{output_filename}")
        
        try:
            # Run the formatter synchronously
            format_document(input_path, output_path)
            
            # Clean up the input file to save space
            try:
                os.remove(input_path)
            except:
                pass
                
            # Return the file ID and filename so the client can download it
            return jsonify({
                'success': True,
                'job_id': job_id,
                'filename': output_filename,
                'message': 'File formatted successfully!'
            })
            
        except Exception as e:
            # Clean up on failure
            try:
                os.remove(input_path)
            except:
                pass
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Allowed file type is .docx only.'}), 400

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
