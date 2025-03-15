import os
import json
import shutil
from pathlib import Path
from flask import Flask, request, render_template, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import sys

# 添加当前目录到系统路径，确保可以导入本地模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入处理模块
from core.flatten_aijson import JsonFlattener
from core.excel_generator import generate_excel
from core.ppt_generator import generate_ppt

app = Flask(__name__)

# 配置上传文件存储路径
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data/input/json/uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data/output')
ALLOWED_EXTENSIONS = {'json'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制上传文件大小为16MB

# 确保上传和输出目录存在
Path(UPLOAD_FOLDER).mkdir(parents=True, exist_ok=True)
Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'files' not in request.files:
        return jsonify({'error': '没有文件部分'}), 400
    
    files = request.files.getlist('files')
    
    if not files or files[0].filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    # 获取生成类型选项
    generate_excel = request.form.get('generate_excel', 'false').lower() == 'true'
    generate_ppt = request.form.get('generate_ppt', 'false').lower() == 'true'
    
    if not generate_excel and not generate_ppt:
        return jsonify({'error': '请至少选择一种生成类型（Excel或PPT）'}), 400
    
    # 清空上传目录，确保只处理当前上传的文件
    for f in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, f)
        if os.path.isfile(file_path):
            os.unlink(file_path)
    
    saved_files = []
    
    # 保存所有上传的文件
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            saved_files.append(file_path)
        else:
            return jsonify({'error': f'不支持的文件类型: {file.filename}'}), 400
    
    try:
        # 处理上传的文件
        result = process_files(saved_files, generate_excel, generate_ppt)
        return jsonify(result)
    except Exception as e:
        import traceback
        error_traceback = traceback.format_exc()
        print(f"处理文件时出错: {str(e)}")
        print(f"错误详情: {error_traceback}")
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

def process_files(file_paths, generate_excel=True, generate_ppt=True):
    """处理上传的多个文件，并根据选择生成Excel和PPT文件"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.dirname(file_paths[0])  # 所有文件都在同一目录
    
    # 配置处理参数
    config = {
        "json_input_pattern": os.path.join(input_path, "*.json"),
        "temp_output_dir": os.path.join(base_dir, "data/input/json/temp/"),
        "output_dir": os.path.join(base_dir, "data/output/"),
        "excel_output_file": "ai_transcript.xlsx"
    }
    
    # 创建输出目录
    Path(config['temp_output_dir']).mkdir(parents=True, exist_ok=True)
    Path(config['output_dir']).mkdir(parents=True, exist_ok=True)
    
    # 处理结果
    result = {
        'status': 'success',
        'messages': [],
        'files': []
    }
    
    try:
        # 第一步：执行JSON展平处理
        json_processor = JsonFlattener(
            input_pattern=config['json_input_pattern'],
            output_dir=config['temp_output_dir']
        )
        processed_data = json_processor.process_files()
        result['messages'].append(f"JSON数据处理完成，共处理 {len(file_paths)} 个文件")
        
        # 第二步：根据选择生成Excel文件
        if generate_excel:
            output_excel_path = os.path.join(config['output_dir'], config['excel_output_file'])
            from core.excel_generator import generate_excel as gen_excel
            gen_excel(processed_data, output_excel_path)
            result['messages'].append(f"共生成一个Excel文件")
            result['files'].append({
                'name': config['excel_output_file'],
                'path': f"/download/{config['excel_output_file']}",
                'type': 'excel'
            })
        
        # 第三步：根据选择生成PPT文件
        if generate_ppt:
            # 复制文件到ai目录以便PPT生成器处理
            ai_dir = os.path.join(base_dir, "data/input/json/ai")
            Path(ai_dir).mkdir(parents=True, exist_ok=True)
            
            # 清空ai目录
            for f in os.listdir(ai_dir):
                ai_file_path = os.path.join(ai_dir, f)
                if os.path.isfile(ai_file_path):
                    os.unlink(ai_file_path)
            
            # 复制上传的文件到ai目录
            for file_path in file_paths:
                filename = os.path.basename(file_path)
                dst = os.path.join(ai_dir, filename)
                shutil.copy2(file_path, dst)
            
            # 生成PPT
            from core.ppt_generator import generate_ppt as gen_ppt
            generated_files = gen_ppt(input_dir="data/input/json/ai", output_dir=config['output_dir'])
            
            result['messages'].append(f"共生成 {len(generated_files)} 个PPT文件")
            
            # 添加生成的PPT文件到结果
            for file_path in generated_files:
                file_name = os.path.basename(file_path)
                result['files'].append({
                    'name': file_name,
                    'path': f"/download/{file_name}",
                    'type': 'ppt'
                })
        
        return result
    
    except Exception as e:
        raise Exception(f"处理文件时出错: {str(e)}")

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

@app.route('/output_files')
def list_output_files():
    files = []
    for filename in os.listdir(OUTPUT_FOLDER):
        if os.path.isfile(os.path.join(OUTPUT_FOLDER, filename)):
            file_type = 'excel' if filename.endswith('.xlsx') else 'ppt' if filename.endswith('.pptx') else 'other'
            files.append({
                'name': filename,
                'path': f"/download/{filename}",
                'type': file_type
            })
    return jsonify({'files': files})

@app.route('/delete_file/<filename>', methods=['DELETE'])
def delete_file(filename):
    """删除指定的输出文件"""
    try:
        # 安全检查：确保文件名不包含路径分隔符
        if '/' in filename or '\\' in filename:
            return jsonify({'error': '无效的文件名'}), 400
        
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        
        # 检查文件是否存在
        if not os.path.exists(file_path) or not os.path.isfile(file_path):
            return jsonify({'error': '文件不存在'}), 404
        
        # 删除文件
        os.remove(file_path)
        return jsonify({'success': True, 'message': f'文件 {filename} 已成功删除'})
    
    except Exception as e:
        return jsonify({'error': f'删除文件时出错: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
