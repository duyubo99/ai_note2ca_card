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
import aiohttp
import asyncio
import requests
import traceback
from server_utils import logger_configuration

logger = logger_configuration('server')

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
    headers = {"Authorization": f"Bearer {apikeys[1]}"}

    async with aiohttp.ClientSession() as session:
        async with session.post(
                url,
                headers=headers,
                json={
                    "model": endpoints[1],
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.7
                }
        ) as response:
            if response.status == 200:
                data = await response.json()
                return data['choices'][0]['message']['content'].replace('```json', '').replace('```', '')
            else:
                logger.error(f"Request failed: {await response.text()}")
                return ""

# deepseek v3(内部测试) 
# async def genText(prompt):
#     url = 'https://ark.cn-beijing.volces.com/api/v3/chat/completions'
#     headers = {
#         "Content-Type": "application/json",
#         "Authorization": f"Bearer {apikeys[1]}"
#     }
#
#     data = {
#         "model": endpoints[1],
#         "messages": [
#             {"role": "system", "content": "You are a helpful assistant."},
#             {"role": "user", "content": prompt}
#         ],
#         "temperature": 0.7,
#         "top_p": 0.8,
#         "repetition_penalty": 1.05,
#         "max_tokens": 16384
#     }
#     response = requests.post(url, headers=headers, json=data)
#     if response.status_code == 200:
#         try:
#             text = response.text.replace('```json', '').replace('```', '')
#             text = json.loads(text)
#             text = text['choices'][0]['message']['content']
#             # 正文内容
#             # text = json.loads(text)
#             return text
#         except:
#             logger.error(f'[genText] error in parse json result, result : {text}')
#             return {}
#     else:
#         # 否则，打印错误信息
#         logger.info(f"Request failed with status code {response.status_code}, response text : {response.text}")
#         return {}


if __name__ == '__main__':
    text = asyncio.run(genText('''你是谁'''))
    print(text)