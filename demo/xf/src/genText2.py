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
import requests
import time
import random
import sys


apikeys = [
    'bf15a351-8e5e-4818-a299-646001cb675d',
    '0cd86c81-8714-4c30-9485-d755b8c8cdce',
    'bf15a351-8e5e-4818-a299-646001cb675d',  # doubao 32k pro
    'sk-pumqmsnwjxgqbbtrbpqkpnibxdskabpqyftbrpsnsnalnqrv'
]
endpoints = [
    'ep-20250219164708-nnn2f',  # v3
    'ep-20250225145337-hkpkn',  # v3
    'ep-20250225145427-8546n',  # r1
    'ep-20250310163026-cfksl',  # doubao 32k pro
    'Qwen/QwQ-32B'
]
# deepseek v3(内部测试) 
def genText(prompt, max_retries=2, retry_delay=2):
    url = 'https://ark.cn-beijing.volces.com/api/v3/chat/completions'
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {apikeys[-2]}"
    }

    data = {
        "model": endpoints[-2],
        "messages": [
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7,
        "top_p": 0.8,
        "repetition_penalty": 1.05,
        "max_tokens": 12000  # Reduced from 16384 to be within the API limit of 12288
    }
    
    retries = 0
    while retries <= max_retries:
        try:
            print(f"Attempt {retries+1}/{max_retries+1}: Sending request to {url} with model {endpoints[-2]}")
            response = requests.post(url, headers=headers, json=data, timeout=30)
            
            if response.status_code == 200:
                try:
                    print("Received 200 response, processing...")
                    text = response.text.replace('```json', '').replace('```', '')
                    json_response = json.loads(text)
                    
                    if 'choices' in json_response and len(json_response['choices']) > 0:
                        content = json_response['choices'][0]['message']['content']
                        print(f"Successfully extracted content: {content[:50]}...")
                        return content
                    else:
                        print(f"Error: 'choices' not found in response: {json_response}")
                        if retries < max_retries:
                            retries += 1
                            wait_time = retry_delay * (1 + random.random())
                            print(f"Retrying in {wait_time:.2f} seconds...")
                            time.sleep(wait_time)
                            continue
                    print("API返回格式错误，无法提取内容。")
                    sys.exit(1)
                except json.JSONDecodeError as e:
                    print(f"JSON decode error: {e}")
                    print(f"Raw response: {response.text[:200]}...")
                    if retries < max_retries:
                        retries += 1
                        wait_time = retry_delay * (1 + random.random())
                        print(f"Retrying in {wait_time:.2f} seconds...")
                        time.sleep(wait_time)
                        continue
                    print(f"JSON解析错误: {str(e)}")
                    sys.exit(1)
                except Exception as e:
                    print(f"Error processing response: {str(e)}")
                    if retries < max_retries:
                        retries += 1
                        wait_time = retry_delay * (1 + random.random())
                        print(f"Retrying in {wait_time:.2f} seconds...")
                        time.sleep(wait_time)
                        continue
                    print(f"处理响应时出错: {str(e)}")
                    sys.exit(1)
            else:
                # 打印错误信息
                print(f"Error: API returned status code {response.status_code}")
                print(f"Response: {response.text[:200]}...")
                if retries < max_retries:
                    retries += 1
                    wait_time = retry_delay * (1 + random.random())
                    print(f"Retrying in {wait_time:.2f} seconds...")
                    time.sleep(wait_time)
                    continue
                print(f"API请求失败，状态码: {response.status_code}")
                sys.exit(1)
        except requests.exceptions.Timeout:
            print("Error: Request timed out")
            if retries < max_retries:
                retries += 1
                wait_time = retry_delay * (1 + random.random())
                print(f"Retrying in {wait_time:.2f} seconds...")
                time.sleep(wait_time)
                continue
            print("API请求超时")
            sys.exit(1)
        except requests.exceptions.ConnectionError:
            print("Error: Connection error")
            if retries < max_retries:
                retries += 1
                wait_time = retry_delay * (1 + random.random())
                print(f"Retrying in {wait_time:.2f} seconds...")
                time.sleep(wait_time)
                continue
            print("API连接错误")
            sys.exit(1)
        except Exception as e:
            print(f"Unexpected error: {str(e)}")
            if retries < max_retries:
                retries += 1
                wait_time = retry_delay * (1 + random.random())
                print(f"Retrying in {wait_time:.2f} seconds...")
                time.sleep(wait_time)
                continue
            print(f"未预期的错误: {str(e)}")
            sys.exit(1)


if __name__ == '__main__':
    text = genText('''你是谁''')
    print(text)
