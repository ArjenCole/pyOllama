from flask import Flask, render_template_string, request, jsonify, render_template
from flask_cors import CORS
import os
import pyExcel

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
    if 'file0' not in request.files:
        return jsonify({'error': '没有文件部分'})
    # file = request.files.get('file')
    _file = request.files['file0']
    # print('filename', _file.filename)
    _dict = file_process(_file)
    if 'message' in _dict.keys():
        _raw_word = '单价\r\n（元）'
        _matched_word = pyFCM.fuzzy_match(_raw_word)
        print(f"识别的结论: '{_raw_word}' 与 '{_matched_word}'匹配")
        return jsonify({'message': f"识别的结论: '{_raw_word}' 与 '{_matched_word}'匹配"})
    else:
        return jsonify({'error': '上传失败，未知错误'})


def file_process(p_file):
    if p_file.filename == '':
        return {'error': '没有选择文件'}
    if p_file:
        if not os.path.exists('F:/GithubRepos/ArjenCole/pyOllama/uploads'):
            os.makedirs('F:/GithubRepos/ArjenCole/pyOllama/uploads')  # 确保目录存在
        # filename = secure_filename(file.filename)  # 使用 Werkzeug 库提供的 secure_filename 函数
        _file_name = p_file.filename
        _file_path = os.path.join('F:/GithubRepos/ArjenCole/pyOllama/uploads', _file_name)
        p_file.save(_file_path)
        print(_file_path)
        pyExcel.get_workbook(_file_path)
        return {'message': '文件上传成功！~'}


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
