'''
@Project ：code 
@File    ：genText.py
@Author  ：Sito
@Date    ：2024/10/9 16:29 
@Description    ：

apikey：bf15a351-8e5e-4818-a299-646001cb675d
v3 endpoint:ep-20250219164708-nnn2f

apikey：0cd86c81-8714-4c30-9485-d755b8c8cdce
v3 endpoint:ep-20250225145337-hkpkn
r1 endpoint:ep-20250225145427-8546n

embedding:ep-20250226124533-ncq9f

'''
import json
import asyncio
import requests
import traceback
import logging
import sys

# 配置日志
def logger_configuration(task='server'):
    '''
        配置日志 handler
    Returns: logger
    '''
    logger = logging.getLogger(f'{task}-log')
    logger.setLevel(logging.INFO)

    # 禁用父级日志记录器的传播
    logger.propagate = False

    # stream handler
    rf_handler = logging.StreamHandler(sys.stderr)
    rf_handler.setLevel(logging.INFO)
    rf_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(message)s"))

    # 添加处理器前检查是否已经存在
    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        logger.addHandler(rf_handler)

    return logger

logger = logger_configuration('genText')

apikeys = [
    'bf15a351-8e5e-4818-a299-646001cb675d',
    '0cd86c81-8714-4c30-9485-d755b8c8cdce',
    'bf15a351-8e5e-4818-a299-646001cb675d'  # doubao 32k pro
]
endpoints = [
    'ep-20250219164708-nnn2f',  # v3
    'ep-20250225145337-hkpkn',  # v3
    'ep-20250225145427-8546n',  # r1
    'ep-20250310163026-cfksl'  # doubao 32k pro
]


async def genText(prompt):
    url = 'https://ark.cn-beijing.volces.com/api/v3/chat/completions'
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {apikeys[1]}"
    }

    data = {
        "model": endpoints[1],
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7
    }
    
    # 使用requests而不是aiohttp
    loop = asyncio.get_event_loop()
    response = await loop.run_in_executor(
        None, 
        lambda: requests.post(url, headers=headers, json=data)
    )
    
    if response.status_code == 200:
        try:
            data = response.json()
            return data['choices'][0]['message']['content'].replace('```json', '').replace('```', '')
        except Exception as e:
            logger.error(f"Request failed to parse: {str(e)}")
            return ""
    else:
        logger.error(f"Request failed with status code {response.status_code}: {response.text}")
        return ""


if __name__ == '__main__':
    text = asyncio.run(genText('''你是谁'''))
    print(text)
