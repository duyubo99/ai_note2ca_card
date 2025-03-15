import os
import json
from pathlib import Path
from core.flatten_aijson import JsonFlattener
from core.excel_generator import generate_excel

def main():
    # 配置路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 输入输出配置
    config = {
        "json_input_pattern": os.path.join(base_dir, "data/input/json/ai/*.json"),
        "temp_output_dir": os.path.join(base_dir, "data/input/json/temp/"), 
        "excel_output_dir": os.path.join(base_dir, "data/output/"),
        "excel_output_file": "ai_transcript.xlsx"
    }

    try:
        # 创建输出目录
        Path(config['temp_output_dir']).mkdir(parents=True, exist_ok=True)
        Path(config['excel_output_dir']).mkdir(parents=True, exist_ok=True)

        # 第一步：执行JSON展平处理
        json_processor = JsonFlattener(
            input_pattern=config['json_input_pattern'],
            output_dir=config['temp_output_dir']
        )
        processed_data = json_processor.process_files()
        print(f"✓ JSON数据处理完成，临时文件保存在：{json_processor.output_dir}")

        # 第二步：生成Excel文件
        output_excel_path = os.path.join(config['excel_output_dir'], config['excel_output_file'])
        
        # 直接从内存数据生成Excel（避免重复读取文件）
        generate_excel(processed_data, output_excel_path)
        print(f"✓ Excel文件生成完成，保存路径：{output_excel_path}")

    except FileNotFoundError as e:
        print(f"× 文件未找到错误：{str(e)}")
    except json.JSONDecodeError as e:
        print(f"× JSON解析错误：{str(e)}")
    except PermissionError as e:
        print(f"× 文件权限错误：{str(e)}\n请检查：\n1. 文件是否被其他程序占用\n2. 是否有写入权限")
    except Exception as e:
        print(f"× 处理过程中发生未预期错误：{str(e)}")
        print("请检查：")
        print("1. JSON文件格式是否正确")
        print("2. 输入文件路径是否存在")
        print("3. 输出目录是否有写入权限")

if __name__ == "__main__":
    main()
