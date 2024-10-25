from flask import Flask, render_template_string, request, jsonify, render_template
from flask_cors import CORS
import os

import openpyxl

import ai
import pyExcel
import pyFCM
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)  # 允许跨源请求
UPLOADS_DIR = 'F:/GithubRepos/ArjenCole/pyOllama/uploads'  #上传文件存储地址
TARGET_WORDS_F8 = ['建筑工程', '安装工程', '设备及工器具购置费', '其他费用', '合计', '单位', '数量', '单位价值元']


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
        # _match_sheet_name, _match_sheet_row, _match_sheet_col = workbook_similarity(_dict['DIR'])
        # return jsonify({'message': f"匹配的sheet名称: '{_match_sheet_name}' 匹配的行：'{_match_sheet_row + 1}'"})
        _match_dict = workbook_similarity(_dict['DIR'])
        return jsonify({
            'message': f"检测到匹配的表单：《{_match_dict['表单名称']}》，"
                       f"建筑工程 坐标：({_match_dict['建筑工程']['row']},{_match_dict['建筑工程']['col']})，"
                       f"安装工程 坐标：({_match_dict['安装工程']['row']},{_match_dict['安装工程']['col']})，"
                       f"设备及工器具购置费 坐标：({_match_dict['设备及工器具购置费']['row']},{_match_dict['设备及工器具购置费']['col']})，"
                       f"其他费用 坐标：({_match_dict['其他费用']['row']},{_match_dict['其他费用']['col']})"
                       f"合计 坐标：({_match_dict['合计']['row']},{_match_dict['合计']['col']})"
                       f"单位 坐标：({_match_dict['单位']['row']},{_match_dict['单位']['col']})"
                       f"数量 坐标：({_match_dict['数量']['row']},{_match_dict['数量']['col']})"
                       f"单位价值（元） 坐标：({_match_dict['单位价值元']['row']},{_match_dict['单位价值元']['col']})"

        })
    else:
        return jsonify({'error': '文件保存失败'})


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
    _work_book = pyExcel.get_workbook(p_dir)  # 用pandas加载文件用于处理
    _work_book_openpyxl = openpyxl.load_workbook(p_dir)  # 用openpyxl加载文件用于识别隐藏表单

    _max_similarity = 0
    rt_match_sheet_name = None
    rt_match_sheet_row = 0
    rt_match_sheet_col = 0

    for fe_sheet_name in _work_book.keys():
        if _work_book_openpyxl[fe_sheet_name].sheet_state == 'hidden':
            continue
        _work_sheet = _work_book[fe_sheet_name]

        print('正在处理表单：', fe_sheet_name)
        # print(_work_sheet)
        _sheet_match_row, _sheet_match_col, _sheet_similarity = worksheet_similarity(_work_sheet)
        if _sheet_similarity > _max_similarity:
            _max_similarity = _sheet_similarity
            rt_match_sheet_name = fe_sheet_name
            rt_match_sheet_row = _sheet_match_row
            rt_match_sheet_col = _sheet_match_col

    # print('匹配的表单：', rt_match_sheet_name, '匹配单元格坐标位置在第', rt_match_sheet_row + 1, '行，第', rt_match_sheet_col + 1, '列')
    rt_dict = sort_f8(_work_book, rt_match_sheet_name, rt_match_sheet_row, rt_match_sheet_col)
    return rt_dict


def worksheet_similarity(p_sheet):
    rt_match_row = 0
    rt_match_col = 0
    rt_max_similarity = 0
    for fe_row in range(0, min(20, p_sheet.shape[0])):
        _f4_similarity_array = list(range(21))
        for fe_col in range(0, min(20, p_sheet.shape[1])):
            # print("行列", fe_row, fe_col)
            _str = str(p_sheet.iloc[fe_row][fe_col])
            if _str == 'nan':
                _f4_similarity = 0
            else:
                # print(_str)
                _f4_similarity = match_f4(_str)
            _f4_similarity_array[fe_col] = _f4_similarity
            _row_similarity = 0
            if fe_col > 8:
                for fe_i in range(8):
                    _row_similarity += _f4_similarity_array[fe_col - fe_i]
                # _row_similarity = (_f4_similarity_array[fe_col] + _f4_similarity_array[fe_col - 1] + _f4_similarity_array[fe_col - 2] + _f4_similarity_array[fe_col - 3])

            if _row_similarity > rt_max_similarity:
                rt_max_similarity = _row_similarity
                rt_match_row = fe_row
                rt_match_col = fe_col - 7
    return rt_match_row, rt_match_col, rt_max_similarity


def match_f4(p_raw_word):
    # ('match_f4', p_raw_word)
    _target_words = TARGET_WORDS_F8

    _matched_word, rt_similarity_score, = pyFCM.fuzzy_match(p_raw_word, _target_words)

    # print('_matched_word', type(_matched_word), _matched_word)
    # print('sim', type(rt_similarity_score), rt_similarity_score)
    if rt_similarity_score is not None:
        # print('rt_similarity_score', rt_similarity_score)
        return rt_similarity_score[0]
    else:
        return 0


def sort_f8(p_work_book, p_sheet_name, p_row, p_col):
    _target_words = TARGET_WORDS_F8
    rt_dict = {'表单名称': p_sheet_name}
    for fe_target_word in _target_words:
        _max_similarity = 0
        _max_fe_i = 0
        for fe_i in range(8):
            _raw_word = p_work_book[p_sheet_name].iloc[p_row][p_col + fe_i]
            _matched_word, _similarity_score, = pyFCM.fuzzy_match(_raw_word, [fe_target_word])
            # 加上识别单元格下方单元格一起识别，以防两个文字被拆分到两个单元格里
            _raw_word = str(_raw_word) + str(p_work_book[p_sheet_name].iloc[p_row + 1][p_col + fe_i])
            _matched_word1, _similarity_score1, = pyFCM.fuzzy_match(_raw_word, [fe_target_word])
            _score = max(_similarity_score[0], _similarity_score1[0])
            if _score > _max_similarity:
                _max_similarity = _score
                _max_fe_i = fe_i
        rt_dict[fe_target_word] = {'row': p_row, 'col': p_col + _max_fe_i}
    print(rt_dict)
    return rt_dict


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
