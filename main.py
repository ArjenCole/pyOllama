import requests
import json

tUrl = "http://127.0.0.1:11434/api/generate"
tData = {
    # "model": "llama3",
    "model": "qwen:7b",
    "prompt": "为什么天空是蓝色的？",
    "stream": False
}

tResponse = requests.post(tUrl, json=tData)

# 检查请求是否成功
if tResponse.status_code == 200:
    tDict = tResponse.json()
    # print("请求成功，响应数据：", tDict)
    print(tDict['response'])
else:
    print(f"请求失败，状态码：{tResponse.status_code}, 响应内容：{tResponse.text}")
