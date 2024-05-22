from flask import Flask, request, jsonify
import requests

app = Flask(__name__)


@app.route('/ai', methods=['POST'])
def chat_ai():
    print("11111")

    # 获取前端发送的 JSON 数据
    user_input = request.json.get('prompt', '')

    # 构建请求体
    tData = {
        "model": "qwen:7b",
        "prompt": user_input,
        "stream": False
    }
    print("2")

    # 发送请求到 AI 模型服务
    try:
        tResponse = requests.post("http://127.0.0.1:11434/api/generate", json=tData)
        tResponse.raise_for_status()  # 如果响应状态码不是 200，将抛出异常
    except requests.exceptions.RequestException as e:
        return jsonify({'error': 'An error occurred while communicating with the AI model service.'}), 500
    print("3")

    # 检查请求是否成功
    if tResponse.status_code == 200:
        print("4")

        tDict = tResponse.json()
        print(tDict['response'])
        return jsonify({'response': tDict['response']})
    else:
        return jsonify({'error': f"Request failed with status code: {tResponse.status_code}"}), tResponse.status_code


def excel_ai(pList):

    # 获取前端发送的 JSON 数据
    user_input = "下面是一组字符串列表，请告诉我这些字符串中，字面意思最接近建筑工程的字符串出现在第几个:"
    for feStr in pList:
        user_input = user_input + "\n" + str(feStr)

    # 构建请求体
    tData = {
        "model": "qwen:7b",
        "prompt": user_input,
        "stream": False
    }

    # 发送请求到 AI 模型服务
    try:
        tResponse = requests.post("http://127.0.0.1:11434/api/generate", json=tData)
        tResponse.raise_for_status()  # 如果响应状态码不是 200，将抛出异常
    except requests.exceptions.RequestException as e:
        return jsonify({'error': 'An error occurred while communicating with the AI model service.'}), 500

    # 检查请求是否成功
    if tResponse.status_code == 200:
        print("4")

        tDict = tResponse.json()
        print(tDict['response'])
        return jsonify({'response': tDict['response']})
    else:
        return jsonify({'error': f"Request failed with status code: {tResponse.status_code}"}), tResponse.status_code

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
