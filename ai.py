from flask import Flask, request, jsonify
import requests

app = Flask(__name__)


@app.route('/ai', methods=['POST'])
def chat_ai():
    # 获取前端发送的 JSON 数据
    _userInput = request.json.get('prompt', '')

    return _send_to_ai(_userInput)


def excel_ai(p_list):
    # 获取前端发送的 JSON 数据
    _userInput = "你是一个字符串匹配器，可以对字符串的字面意思进行匹配；下面是一组字符串列表，每一行是一个字符串， " + p_list[0] + " 是第一行的字符串，请告诉我这些字符串中，建筑工程 出现在第几行:"
    for feStr in p_list:
        _userInput = _userInput + "\n" + str(feStr)

    return _send_to_ai(_userInput)


def _send_to_ai(p_prompt):
    # 构建请求体
    _data = {
        "model": "qwen:7b",
        "prompt": p_prompt,
        "stream": False
    }

    print(p_prompt)
    # 发送请求到 AI 模型服务
    try:
        _response = requests.post("http://127.16.104.58:11434/api/generate", json=_data)
        _response.raise_for_status()  # 如果响应状态码不是 200，将抛出异常
    except requests.exceptions.RequestException as e:
        print(e)
        return jsonify({'error': 'An error occurred while communicating with the AI model service.'}), 500

    # 检查请求是否成功
    if _response.status_code == 200:

        _dict = _response.json()
        print(_dict['response'])
        return jsonify({'response': _dict['response']})
    else:
        return jsonify({'error': f"Request failed with status code: {_response.status_code}"}), _response.status_code


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
