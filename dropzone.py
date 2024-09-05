from flask import Flask, render_template_string, request, jsonify, render_template
from flask_cors import CORS
import os

import ai
import pyExcel
import pyFCM
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)  # 允许跨源请求


@app.route('/dropzone')
def dropzone():
    return render_template('dropzone.html')


@app.route('/upload', methods=['POST'])
def dropzone_upload():
    print('进来了！', request.files)
    if 'file0' not in request.files:
        print('出去了！')
        return jsonify({'error': '没有文件部分'})
    # file = request.files.get('file')
    _file = request.files['file0']
    print('filename', _file.filename)
    if _file.filename == '':
        return jsonify({'error': '没有选择文件'})
    if _file:
        if not os.path.exists('F:/GithubRepos/ArjenCole/pyOllama/uploads'):
            os.makedirs('F:/GithubRepos/ArjenCole/pyOllama/uploads')  # 确保目录存在
        # filename = secure_filename(file.filename)  # 使用 Werkzeug 库提供的 secure_filename 函数
        _fileName = _file.filename
        _filePath = os.path.join('F:/GithubRepos/ArjenCole/pyOllama/uploads', _fileName)

        print(_filePath)
        _file.save(_filePath)

        # 读取表头单元格字符串内容
        # headers = pyExcel.extract_headers(_filePath)
        # print(headers)

        # ai.excel_ai(headers)
        raw_word = '单价\r\n（元）'
        matched_word = pyFCM.fuzzy_match(raw_word)
        print(f"识别的结论: '{raw_word}' 与 '{matched_word}'匹配")

        return jsonify({'message': '文件上传成功'})
    return jsonify({'error': '上传失败，未知错误'})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
