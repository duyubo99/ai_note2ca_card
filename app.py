import os
import shutil
import asyncio
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, jsonify
from werkzeug.utils import secure_filename

# Import the core processing functions
from core.ai_core import process_files
from core.flatten_aijson import JsonFlattener
from core.excel_generator import generate_excel
from core.ppt_generator import generate_ppt

app = Flask(__name__)
app.secret_key = 'ai_note2ca_card_secret_key'

# Configure upload and output directories
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'data/input/temp_upload/docx')
JSON_OUTPUT_FOLDER = os.path.join(BASE_DIR, 'data/input/json/uploads')
FINAL_OUTPUT_FOLDER = os.path.join(BASE_DIR, 'data/output')

# Ensure directories exist
Path(UPLOAD_FOLDER).mkdir(parents=True, exist_ok=True)
Path(JSON_OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)
Path(FINAL_OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)

# Configure allowed file extensions
ALLOWED_EXTENSIONS = {'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['JSON_OUTPUT_FOLDER'] = JSON_OUTPUT_FOLDER
app.config['FINAL_OUTPUT_FOLDER'] = FINAL_OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'),
                               'favicon.ico', mimetype='image/vnd.microsoft.icon')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Check if any files were uploaded
    if 'files[]' not in request.files:
        flash('没有选择文件', 'error')
        return redirect(url_for('index'))
    
    files = request.files.getlist('files[]')
    
    # Check if any files were selected
    if not files or files[0].filename == '':
        flash('没有选择文件', 'error')
        return redirect(url_for('index'))
    
    # Clear upload directory before processing new files
    for file in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, file)
        if os.path.isfile(file_path):
            os.unlink(file_path)
    
    # Save uploaded files
    filenames = []
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)
            filenames.append(filename)
    
    if not filenames:
        flash('没有有效的文件被上传', 'error')
        return redirect(url_for('index'))
    
    # Get output format preferences
    generate_excel_output = 'excel' in request.form
    generate_ppt_output = 'ppt' in request.form
    
    if not (generate_excel_output or generate_ppt_output):
        flash('请至少选择一种输出格式', 'error')
        return redirect(url_for('index'))
    
    # Process the files
    try:
        # Get full paths of uploaded files
        file_paths = [os.path.join(UPLOAD_FOLDER, filename) for filename in filenames]
        
        # Process DOCX files to generate JSON
        json_output_filename = 'ai_json.json'
        json_output_path = os.path.join(JSON_OUTPUT_FOLDER, json_output_filename)
        
        # Process the files using the existing functionality
        process_files(file_paths, JSON_OUTPUT_FOLDER, json_output_filename)
        
        # Generate outputs based on user preferences
        output_files = []
        
        if generate_excel_output:
            # Flatten JSON and generate Excel
            json_processor = JsonFlattener(
                input_pattern=os.path.join(JSON_OUTPUT_FOLDER, '*.json'),
                output_dir=os.path.join(BASE_DIR, 'data/input/json/temp')
            )
            processed_data = json_processor.process_files()
            
            # Generate Excel file
            excel_output_path = os.path.join(FINAL_OUTPUT_FOLDER, 'ai_transcript.xlsx')
            generate_excel(processed_data, excel_output_path)
            output_files.append(('Excel文件', 'ai_transcript.xlsx'))
        
        if generate_ppt_output:
            # Generate PPT files
            ppt_files = generate_ppt(input_dir=JSON_OUTPUT_FOLDER, output_dir=FINAL_OUTPUT_FOLDER)
            for ppt_file in ppt_files:
                ppt_filename = os.path.basename(ppt_file)
                output_files.append(('PPT文件', ppt_filename))
        
        return render_template('results.html', output_files=output_files)
    
    except Exception as e:
        flash(f'处理文件时出错: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(FINAL_OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
