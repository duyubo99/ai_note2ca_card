import os
import json
import uuid
import asyncio
from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session
from werkzeug.utils import secure_filename

from core.flatten_aijson import JsonFlattener
from core.excel_generator import generate_excel
from core.ppt_generator import generate_ppt
from core.ai_core import process_docx

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data/input/json/uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload size

# 确保上传目录存在
Path(app.config['UPLOAD_FOLDER']).mkdir(parents=True, exist_ok=True)

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # 检查是否有文件被上传
    if 'files[]' not in request.files:
        flash('没有选择文件', 'error')
        return redirect(request.url)
    
    files = request.files.getlist('files[]')
    
    # 检查是否至少有一个文件
    if not files or files[0].filename == '':
        flash('没有选择文件', 'error')
        return redirect(request.url)
    
    # 检查文件类型
    for file in files:
        if not allowed_file(file.filename):
            flash(f'不支持的文件类型: {file.filename}', 'error')
            return redirect(request.url)
    
    # 创建唯一的会话ID
    session_id = str(uuid.uuid4())
    session['session_id'] = session_id
    
    # 创建会话目录
    session_dir = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
    Path(session_dir).mkdir(parents=True, exist_ok=True)
    
    # 保存文件
    filenames = []
    for file in files:
        filename = secure_filename(file.filename)
        file_path = os.path.join(session_dir, filename)
        file.save(file_path)
        filenames.append(filename)
    
    # 保存文件名到会话
    session['filenames'] = filenames
    
    # 获取用户选择的输出格式
    output_formats = request.form.getlist('output_format')
    session['output_formats'] = output_formats
    
    # 重定向到处理页面
    return redirect(url_for('process_files_route'))

@app.route('/process')
def process_files_route():
    # 检查会话是否存在
    if 'session_id' not in session:
        flash('会话已过期，请重新上传文件', 'error')
        return redirect(url_for('index'))
    
    session_id = session['session_id']
    filenames = session.get('filenames', [])
    output_formats = session.get('output_formats', [])
    
    # 设置处理路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    session_dir = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
    ai_output_dir = os.path.join(base_dir, f"data/input/json/ai/{session_id}")
    temp_output_dir = os.path.join(base_dir, f"data/input/json/temp/{session_id}")
    final_output_dir = os.path.join(base_dir, f"data/output/{session_id}")
    
    # 创建必要的目录
    Path(ai_output_dir).mkdir(parents=True, exist_ok=True)
    Path(temp_output_dir).mkdir(parents=True, exist_ok=True)
    Path(final_output_dir).mkdir(parents=True, exist_ok=True)
    
    try:
        # 处理上传的文件
        file_paths = [os.path.join(session_dir, filename) for filename in filenames]
        
        # 第一步：处理Word文档并生成JSON数据
        asyncio.run(process_docx(
            file_paths=file_paths,
            output_dir=ai_output_dir,
            output_filename='ai_json.json'
        ))
        
        # 第二步：执行JSON展平处理
        json_processor = JsonFlattener(
            input_pattern=os.path.join(ai_output_dir, '*.json'),
            output_dir=temp_output_dir
        )
        processed_data = json_processor.process_files()
        
        # 生成选定的输出格式
        output_files = {}
        
        if 'excel' in output_formats:
            # 生成Excel文件
            excel_output_path = os.path.join(final_output_dir, 'ai_transcript.xlsx')
            generate_excel(processed_data, excel_output_path)
            output_files['excel'] = 'ai_transcript.xlsx'
        
        if 'ppt' in output_formats:
            # 生成PPT文件
            ppt_files = generate_ppt(
                input_dir=ai_output_dir,
                output_dir=final_output_dir
            )
            if ppt_files:
                output_files['ppt'] = [os.path.basename(file) for file in ppt_files]
        
        # 保存输出文件信息到会话
        session['output_files'] = output_files
        session['output_dir'] = final_output_dir
        
        return render_template('download.html', output_files=output_files)
    
    except Exception as e:
        flash(f'处理文件时出错: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/download/<file_type>/<filename>')
def download_file(file_type, filename):
    # 检查会话是否存在
    if 'session_id' not in session or 'output_dir' not in session:
        flash('会话已过期，请重新上传文件', 'error')
        return redirect(url_for('index'))
    
    output_dir = session['output_dir']
    file_path = os.path.join(output_dir, filename)
    
    # 检查文件是否存在
    if not os.path.exists(file_path):
        flash(f'文件不存在: {filename}', 'error')
        return redirect(url_for('index'))
    
    # 设置下载文件的MIME类型
    mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if file_type == 'excel' else 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    
    return send_file(file_path, as_attachment=True, mimetype=mime_type)

@app.route('/reset')
def reset():
    # 清除会话
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
