'''
@Project ：code 
@File    ：main.py
@Author  ：Sito
@Date    ：2025/3/4 11:17 
@Description    ：
'''
import copy
import glob
import json
import asyncio
import traceback
from config import *
from docx import Document
from genText import genText
from prompt_manager import *
from tqdm.asyncio import tqdm_asyncio

logger = logger_configuration()
folders = ['年轻人波轮项目笔录', '成都8位用户笔录', '衡阳7位用户笔录']


def extract_word_tables():
    result = []
    for foler in folders:
        docxs = glob.glob(f'{foler}/*.docx')
        # 调用示例
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
    query = get_extract_prompt(text)
    j = await genText(query)
    try:
        j = json.loads(j)
    except:
        logger.error(f'[process] parse json fail, json : {j}, error : {traceback.format_exc()}')
    return doc_path, j


def postprocess(results, doc_path, j):
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
        logger.error(f'[postprocess] fail, json : {j}, error : {traceback.format_exc()}')


async def main():
    info = extract_word_tables()
    process_task = [process(doc_path, txt) for (doc_path, txt) in info]
    llm_result = await tqdm_asyncio.gather(*process_task)

    for folder in folders:
        results = {}
        format = copy.deepcopy(level_label)
        for (doc_path, j) in llm_result:
            if folder in doc_path:
                if doc_path not in results: results[doc_path] = format
                postprocess(results, doc_path, j)

        json_results = json.dumps(results, indent=4, ensure_ascii=False)
        with open(f'{folder}.txt', 'w', encoding='utf-8') as file:
            file.write(json_results)
        logger.info(f'全部处理完成，已经写入到{folder}.txt中')


if __name__ == '__main__':
    asyncio.run(main())
