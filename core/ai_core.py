'''
@Project ：code 
@File    ：ai_core.py
@Author  ：Sito
@Date    ：2025/3/4 11:17 
@Description    ：AI核心处理模块，提供文档处理和AI分析功能
'''
import copy
import glob
import json
import os
import asyncio
import traceback
from pathlib import Path
from core.config import *
from docx import Document
from core.genText import genText
from core.prompt_manager import *
from tqdm.asyncio import tqdm_asyncio


def extract_word_tables(file_path=None, input_dir='data/input/docx'):
    """
    从Word文档中提取表格内容
    
    Args:
        file_path: 单个文档路径，如果提供则只处理该文档
        input_dir: 文档目录，当file_path为None时使用
        
    Returns:
        包含(文档路径, 表格内容)元组的列表
    """
    result = []
    
    # 确定要处理的文档列表
    if file_path:
        docxs = [file_path]
    else:
        docxs = glob.glob(f'{input_dir}/*.docx')
    
    # 处理每个文档
    for doc_path in docxs:
        doc = Document(doc_path)
        all_tables = []

        # 遍历所有表格
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)
            all_tables.append(table_data)

        # 写入 result
        for idx, table in enumerate(all_tables):
            tmp_text = f"Table {idx + 1}:\n"
            for row in table:
                tmp_text += '\t|\t'.join(row).replace('\n', '') + '\n'
            tmp_text += '\n' + '=' * 50 + '\n'
            for i in range(0, len(tmp_text), CHUNK_SIZE):
                tmp_text_part = tmp_text[i:i + CHUNK_SIZE]
                result.append((doc_path, tmp_text_part))
    
    return result
async def process(doc_path, text):
    """
    处理文档内容，调用AI生成结构化数据
    
    Args:
        doc_path: 文档路径
        text: 文档内容
        
    Returns:
        (文档路径, 处理结果)元组
    """
    query = get_extract_prompt(text)
    j = await genText(query)
    try:
        j = json.loads(j)
    except:
        print(f'[process] parse json fail, json : {j}, error : {traceback.format_exc()}')
    return doc_path, j


def postprocess(results, doc_path, j):
    """
    后处理AI生成的结果
    
    Args:
        results: 结果字典
        doc_path: 文档路径
        j: AI生成的JSON数据
    """
    try:
        for k in j:
            for kk in j[k]:
                for kkk in j[k][kk]:
                    if j[k][kk][kkk]:
                        if j[k][kk][kkk] in deny_words:
                            continue
                        if k == '总结标签':
                            results[doc_path][k][kk][kkk] += j[k][kk][kkk] + '  '
                        else:
                            results[doc_path][k][kk][kkk] = j[k][kk][kkk]
    except:
        print(f'[postprocess] fail, json : {j}, error : {traceback.format_exc()}')


async def process_docx(file_path=None, input_dir='data/input/docx', output_dir='data/input/json/temp', output_filename='ai_json.json'):
    """
    处理Word文档并生成结构化JSON数据
    
    Args:
        file_path: 单个文档路径，如果提供则只处理该文档
        input_dir: 文档目录，当file_path为None时使用
        output_dir: 输出目录
        output_filename: 输出文件名
        
    Returns:
        处理结果字典
    """
    # 确保输出目录存在
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    # 提取表格内容
    info = extract_word_tables(file_path, input_dir)
    
    # 处理每个表格内容
    process_task = [process(doc_path, txt) for (doc_path, txt) in info]
    llm_result = await tqdm_asyncio.gather(*process_task)

    # 整合结果
    results = {}
    format = copy.deepcopy(level_label)
    for (doc_path, j) in llm_result:
        if doc_path not in results: 
            results[doc_path] = copy.deepcopy(format)
        postprocess(results, doc_path, j)

    # 保存结果
    output_path = os.path.join(output_dir, output_filename)
    json_results = json.dumps(results, indent=4, ensure_ascii=False)
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(json_results)
    
    return results


def process_file(file_path, output_dir='data/input/json/temp', output_filename='ai_json.json'):
    """
    处理单个Word文档并生成结构化JSON数据（外部调用接口）
    
    Args:
        file_path: 文档路径
        output_dir: 输出目录
        output_filename: 输出文件名
        
    Returns:
        处理结果字典
    """
    return asyncio.run(process_docx(file_path=file_path, output_dir=output_dir, output_filename=output_filename))


async def main():
    """原始主函数，处理input_dir目录下的所有文档"""
    return await process_docx()


if __name__ == '__main__':
    asyncio.run(main())
