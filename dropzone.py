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
UPLOADS_DIR = 'F:/GithubRepos/ArjenCole/pyOllama/uploads'  #上传文件存储地址


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
    _dict = file_save(_file)
    if 'DIR' in _dict.keys():
        # _work_book_similarity = workbook_similarity(_dict['DIR'])

        _raw_word = '建  筑\r\n工  程'
        _matched_word, similarity_score, = pyFCM.fuzzy_match(_raw_word)
        print(f"识别的结论: '{_raw_word}' 与 '{_matched_word}'匹配，匹配度'{similarity_score}")
        return jsonify({'message': f"识别的结论: '{_raw_word}' 与 '{_matched_word}'匹配"})
    else:
        return jsonify({'error': '上传失败，未知错误'})


def file_save(p_file):
    if p_file.filename == '':
        return {'error': '没有选择文件'}
    if p_file:
        if not os.path.exists(UPLOADS_DIR):
            os.makedirs(UPLOADS_DIR)  # 确保目录存在
        # filename = secure_filename(file.filename)  # 使用 Werkzeug 库提供的 secure_filename 函数
        _file_name = p_file.filename
        _file_path = os.path.join(UPLOADS_DIR, _file_name)
        p_file.save(_file_path)
        print(_file_path)
        return {'DIR': _file_path}


def workbook_similarity(p_dir):
    _work_book = pyExcel.get_workbook(p_dir)
    rt_work_book = pyExcel.new_workbook
    print(_work_book,rt_work_book)

    for fe_sheet_name in _work_book.sheetnames:
        _worksheet = _work_book[fe_sheet_name]

        for fe_row in range(1, _worksheet.max_row + 1):
            for fe_col in range(1, _worksheet.max_column + 1):
                rt_work_book[fe_sheet_name][fe_row][fe_col] = match_f4(_worksheet[fe_row][fe_col])
    return rt_work_book


def match_f4(p_raw_word):
    _target_words = ['建筑工程', '安装工程', '设备及工器具购置费', '其他费用']

    _matched_word, rt_similarity_score, = pyFCM.fuzzy_match(p_raw_word, _target_words)
    return rt_similarity_score


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
