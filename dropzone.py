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
        _match_sheet_name, _match_sheet_row = workbook_similarity(_dict['DIR'])
        print('匹配sheet', _match_sheet_name, _match_sheet_row)
        # _work_book = pyExcel.get_workbook(_dict['DIR'])

        '''
        _raw_word = '建  筑\r\n工  程'
        _matched_word, similarity_score, = pyFCM.fuzzy_match(_raw_word)
        print(f"识别的结论: '{_raw_word}' 与 '{_matched_word}'匹配，匹配度'{similarity_score}")        
        return jsonify({'message': f"识别的结论: '{_raw_word}' 与 '{_matched_word}'匹配"})
        '''
        return None
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

    _max_similarity = 0
    rt_match_sheet_name = None
    rt_match_sheet_row = 0

    for fe_sheet_name in _work_book.keys():
        _work_sheet = _work_book[fe_sheet_name]
        _sheet_match_row, _sheet_similarity = worksheet_similarity(_work_sheet)
        if _sheet_similarity > _max_similarity:
            _max_similarity = _sheet_similarity
            rt_match_sheet_name = fe_sheet_name
            rt_match_sheet_row = _sheet_match_row

        '''
        for fe_row in range(0, _work_sheet.shape[0]):
            for fe_col in range(0, _work_sheet.shape[1]):
                # print(fe_row, fe_col, _work_sheet.iloc[fe_row][fe_col])
                _similarity = match_f4(_work_sheet.iloc[fe_row][fe_col])
                rt_work_book[fe_sheet_name].iloc[fe_row][fe_col] = _similarity
                print(fe_row, fe_col, _work_sheet.iloc[fe_row][fe_col], rt_work_book[fe_sheet_name].iloc[fe_row][fe_col], _similarity)
        '''
    return rt_match_sheet_name, rt_match_sheet_row


def worksheet_similarity(p_sheet):
    print("表单", p_sheet)

    rt_match_row = 0
    rt_max_similarity = 0
    for fe_row in range(0, min(20, p_sheet.shape[0])):
        _f4_similarity_array = list(range(21))
        for fe_col in range(0, min(20, p_sheet.shape[1])):
            print("行列", fe_row, fe_col, p_sheet.iloc[fe_row][fe_col])
            f4_similarity = match_f4(p_sheet.iloc[fe_row][fe_col])
            _f4_similarity_array[fe_col] = f4_similarity
            _row_similarity = 0
            if fe_col >= 3:
                _row_similarity = (_f4_similarity_array[fe_col] + _f4_similarity_array[fe_col - 1]
                                   + _f4_similarity_array[fe_col - 2] + _f4_similarity_array[fe_col - 3])
            if _row_similarity > rt_max_similarity:
                rt_max_similarity = _row_similarity
                rt_match_row = fe_row
    return rt_match_row, rt_max_similarity


def match_f4(p_raw_word):
    # ('match_f4', p_raw_word)
    _target_words = ['建筑工程', '安装工程', '设备及工器具购置费', '其他费用']

    _matched_word, rt_similarity_score, = pyFCM.fuzzy_match(p_raw_word, _target_words)
    if rt_similarity_score is not None:
        return rt_similarity_score[0][0]
    else:
        return 0


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
