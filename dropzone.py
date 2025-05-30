import copy
import time

from flask import Flask, render_template_string, request, jsonify, render_template
from flask_cors import CORS
import os

import openpyxl

import ai
import pyExcel
import pyFCM
from werkzeug.utils import secure_filename

import pyFormat

app = Flask(__name__)
CORS(app)  # 允许跨源请求

UPLOADS_DIR = 'F:/GithubRepos/ArjenCole/pyOllama/uploads'  # 上传文件存储地址
OUTPUT_DIR = 'F:/GithubRepos/ArjenCole/pyOllama/output'  # 输出文件存储地址
STANDARD_DIR = 'F:/GithubRepos/ArjenCole/pyOllama/templates/standard1.xlsx'  # 标准文件存储地址

TARGET_WORDS_F4 = ['建筑工程', '安装工程', '设备及工器具购置费', '其他费用']
TARGET_WORDS_F8 = ['建筑工程', '安装工程', '设备及工器具购置费', '其他费用', '合计', '单位', '数量', '单位价值元', '备注']
TARGET_WORDS_NO = ['序号', '项', '目', '节', '细目', '工程或费用名称']
MAPPING_TABLE = {'建筑工程': ['建筑工程'],
                 '安装工程': ['安装工程', '管件材料及设备安装工程'],
                 '设备及工器具购置费': ['设备及工器具购置费', '设备购置', '工器具购置', '设备费'],
                 '其他费用': ['其他费用', '独立费用', '工程建设其他费用'],
                 '合计': ['合计'],
                 '单位': ['单位'],
                 '数量': ['数量'],
                 '单位价值元': ['单位价值元', '单位指标元'],
                 '备注': ['备注'],
                 '序号': ['序号'], '项': ['项'], '目': ['目'], '节': ['节'], '细目': ['细目'],
                 '工程或费用名称': ['工程或费用名称', '工程及费用名称'],
                 }
# 表头识别时，《建筑工程》单元格的检索范围
ROW_RANGE = 10
COL_RANGE = 20
KEYWORDS_NUM = 8
THRESHOLD = 0.5  # 相似度阈值


@app.route('/dropzone')
def dropzone():
    return render_template('dropzone.html')


def dropzone_upload(p_socketio, p_client_sessions):
    if 'file0' not in request.files:
        return jsonify({'error': '没有文件部分'})
    _file = request.files['file0']
    _dir_dict = _file_save(_file)

    # 从请求头中获取会话ID
    session_id = request.headers.get('X-Session-ID')

    if 'DIR' in _dir_dict.keys():
        _stage_update(p_socketio, 10, '文件上传成功！开始解析工作簿……', session_id)
        try:
            _dict = _parse_workbook(_dir_dict['DIR'], p_socketio, session_id)
        except Exception as e:
            _stage_update(p_socketio, 100, f'文件识别异常！发生错误：{e}', session_id)
            return
        else:
            _stage_update(p_socketio, 80, '文件解析成功！正在输出结果', session_id)

        pyFormat.table_format(_dir_dict['DIR'], _dict)
        _stage_update(p_socketio, 100, '文件标准化成功！\r\n' + _beautify(_dict), session_id)
    else:
        _stage_update(p_socketio, 100, '文件上传失败！', session_id)

    return _dir_dict


'''
def dropzone_parse(p_dict):
    if 'DIR' in p_dict.keys():
        # _match_sheet_name, _match_sheet_row, _match_sheet_col = workbook_similarity(_dict['DIR'])
        # return jsonify({'message': f"匹配的sheet名称: '{_match_sheet_name}' 匹配的行：'{_match_sheet_row + 1}'"})
        _match_dict = _parse_workbook(p_dict['DIR'])
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
'''


def _file_save(p_file):
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
        _, _file_extension = os.path.splitext(_file_path)
        if _file_extension.lower() == '.xls':
            _file_path = pyExcel.trans_to_xlsx(_file_path)
            print(_file_path)
        return {'DIR': _file_path}


# 对于对工作簿进行处理
def _parse_workbook(p_dir, p_socketio, p_session_id):
    isXLS = False
    try:
        rt_work_book = pyExcel.get_workbook(p_dir)  # 用pandas加载文件用于处理
        _work_book_openpyxl = openpyxl.load_workbook(p_dir)  # 用openpyxl加载文件用于识别隐藏表单
    except openpyxl.utils.exceptions.InvalidFileException as e:
        isXLS = True  # 如果是xls格式的openpyxl库无法解析，因此不使用此库，导致后续隐藏表单无法跳过
    except Exception as e:
        raise

    _max_similarity = 0
    _match_sheet_name = None
    _match_sheet_row = 0
    _match_sheet_col = 0

    _sheet_count = len(rt_work_book)
    _progress = 10
    _progress_step = int((80 - _progress) / (_sheet_count + 1))

    for fe_sheet_name in rt_work_book.keys():
        _progress += _progress_step
        _stage_update(p_socketio, _progress, '正在处理表单：' + fe_sheet_name, p_session_id)
        print('正在处理表单：', fe_sheet_name)
        if not isXLS:
            if _work_book_openpyxl[fe_sheet_name].sheet_state == 'hidden':
                continue
        _work_sheet = rt_work_book[fe_sheet_name]

        _sheet_match_row, _sheet_match_col, _sheet_similarity = _worksheet_similarity(_work_sheet)
        # print('max', _sheet_match_row, _sheet_match_col, _sheet_similarity)
        if _sheet_similarity > _max_similarity:
            _max_similarity = _sheet_similarity
            _match_sheet_name = fe_sheet_name
            _match_sheet_row = _sheet_match_row
            _match_sheet_col = _sheet_match_col
    # print('表单：', _match_sheet_name, ' 第', _match_sheet_row, '行，第', _match_sheet_col, '列', '*计数从0开始')
    rt_dict = {'表单名称': _match_sheet_name}
    rt_dict.update(
        _sort_words(rt_work_book, _match_sheet_name, _match_sheet_row, _match_sheet_col,
                    TARGET_WORDS_F8))
    rt_dict.update(
        _sort_words(rt_work_book, _match_sheet_name, max(_match_sheet_row - 1, 0), _match_sheet_col,
                    TARGET_WORDS_F8, rt_dict))
    rt_dict.update(
        _sort_words(rt_work_book, _match_sheet_name, max(_match_sheet_row - 1, 0), max(_match_sheet_col - 6, 0),
                    TARGET_WORDS_NO, rt_dict, _match_sheet_col - max(_match_sheet_col - 6, 0)))
    rt_dict.update(
        _sort_words(rt_work_book, _match_sheet_name, max(_match_sheet_row, 0), max(_match_sheet_col - 6, 0),
                    TARGET_WORDS_NO, rt_dict, _match_sheet_col - max(_match_sheet_col - 6, 0)))
    # print(rt_dict)
    rt_dict = _parse_no(rt_dict)
    # print(rt_dict)
    rt_dict = _parse_low_sim(rt_dict)
    print(rt_dict)
    return rt_dict


# 对于单个表单前20行、20列进行f8识别
def _worksheet_similarity(p_sheet):
    rt_match_row = 0
    rt_match_col = 0
    rt_max_similarity = 0
    for fe_row in range(0, min(ROW_RANGE, p_sheet.shape[0])):
        _f4_similarity_array = [0.00 for _ in range(COL_RANGE+KEYWORDS_NUM)]
        _row_similarity = 0.00
        for fe_col in range(0, min(COL_RANGE+KEYWORDS_NUM, p_sheet.shape[1])):
            # print("行列", _col_name(fe_col), fe_row+1)
            _str = str(p_sheet.iloc[fe_row][fe_col])
            if _str == 'nan':
                _f4_similarity_array[fe_col] = 0.00
            else:
                _match = _match_f8(_str)
                if _match <= THRESHOLD:
                    _f4_similarity_array[fe_col] = 0.00
                else:
                    _f4_similarity_array[fe_col] = _match_f8(_str)

            _row_similarity += _f4_similarity_array[fe_col]
            if fe_col >= KEYWORDS_NUM:
                _row_similarity -= _f4_similarity_array[fe_col-KEYWORDS_NUM]
            # print(_col_name(fe_col), fe_row+1, round(_row_similarity, 3), '|', _f4_similarity_array)
            '''
            _row_similarity = 0
            if fe_col >= KEYWORDS_NUM:
                for fe_i in range(KEYWORDS_NUM):
                    _row_similarity += _f4_similarity_array[fe_col - fe_i]
            '''
            if _row_similarity >= rt_max_similarity:
                # print('?', fe_col, _col_name(fe_col), fe_row+1)
                rt_max_similarity = _row_similarity
                rt_match_row = fe_row
                rt_match_col = fe_col - KEYWORDS_NUM + 1
                if rt_match_col < 0:  # 如果相似度求和最大列出现在8列之前，识别的起始列定义为从第0列开始第一个相似度大于0的列
                    for fe_i in range(fe_col):
                        if _f4_similarity_array[fe_i] > 0:
                            rt_match_col = fe_i
                            break
    # print('return', _col_name(rt_match_col), rt_match_row+1, rt_max_similarity)
    return rt_match_row, rt_match_col, rt_max_similarity


# 识别从《建筑工程》到《单位价值（元）》的八个关键词
def _match_f8(p_raw_word):
    _target_words = TARGET_WORDS_F8
    _matched_word, rt_similarity_score, = pyFCM.fuzzy_match(p_raw_word, _target_words)
    if rt_similarity_score is not None:
        if p_raw_word in TARGET_WORDS_F4:
            rt_similarity_score[0] = rt_similarity_score[0] * 3
        return rt_similarity_score[0]
    else:
        return 0


# 将匹配F8的单元格与F8关键字进行匹配对应
def _sort_words(p_work_book, p_sheet_name, p_row, p_col, p_target_words, p_dict=None, p_max_col=10):
    # print('sort', p_col, p_max_col)
    if p_dict is None:
        rt_dict = {}
    else:
        rt_dict = copy.deepcopy(p_dict)
    for fe_i in range(p_max_col):
        if p_col + fe_i >= p_work_book[p_sheet_name].shape[1]:
            continue
        _cell_word = str(p_work_book[p_sheet_name].iloc[p_row][p_col + fe_i])
        _cell_word1 = str(p_work_book[p_sheet_name].iloc[p_row + 1][p_col + fe_i])
        if str(_cell_word) == 'nan' or '%' in _cell_word or '百分比' in _cell_word or '比例' in _cell_word:
            continue
        _max_similarity = 0.00
        _match_target_word = None
        for fe_target_word in p_target_words:
            _raw_word = _cell_word
            _matched_word, _similarity_score, = pyFCM.fuzzy_match(_raw_word, MAPPING_TABLE[fe_target_word])
            # print(_raw_word, _matched_word, _similarity_score)
            # 加上识别单元格下方单元格一起识别，以防两个文字被拆分到两个单元格里
            _raw_word = _cell_word + _cell_word1
            _matched_word1, _similarity_score1, = pyFCM.fuzzy_match(_raw_word, MAPPING_TABLE[fe_target_word])
            # print(_raw_word, _matched_word, _similarity_score)
            _score = max(_similarity_score[0], _similarity_score1[0])
            if _score >= _max_similarity:
                _max_similarity = _score
                _match_target_word = fe_target_word
        # print('cell', _cell_word, _max_similarity, _match_target_word)
        if _match_target_word not in rt_dict:
            rt_dict[_match_target_word] = [{'row': p_row, 'col': p_col + fe_i,
                                           'sim': str(round(_max_similarity, 3))}]
        else:
            for fe_value in rt_dict[_match_target_word]:
                if fe_value['col'] == p_col + fe_i:
                    if _max_similarity > float(fe_value['sim']):
                        fe_value['row'] = p_row
                        fe_value['sim'] = str(round(_max_similarity, 3))
                else:
                    rt_dict[_match_target_word].append({'row': p_row, 'col': p_col + fe_i,
                                                        'sim': str(round(_max_similarity, 3))})

    # print('=', rt_dict)
    return rt_dict


# 判断序号模式是"项目节还是序号"
def _parse_no(p_dict):
    rt_dict = p_dict.copy()
    _target_words = ['序号', '项', '目', '节', '细目']
    if '项' in p_dict or '目' in p_dict or '节' in p_dict or '细目' in p_dict:
        _No_xmjx = 0.00
        _No_Num = 0.00
        for fe_target_word in _target_words:
            if fe_target_word in p_dict:
                _max_similarity = 0.00
                _max_coordinate = None
                for fe_coordinate in rt_dict[fe_target_word]:
                    if float(fe_coordinate['sim']) > _max_similarity:
                        _max_similarity = float(fe_coordinate['sim'])
                        _max_coordinate = fe_coordinate
                rt_dict[fe_target_word] = [_max_coordinate]
                if fe_target_word == '序号':
                    _No_Num = max(_No_Num, float(_max_coordinate['sim']))
                else:
                    _No_xmjx = max(_No_xmjx, float(_max_coordinate['sim']))

        if _No_xmjx >= _No_Num:
            if '序号' in rt_dict: del rt_dict['序号']
        else:
            if '项' in rt_dict: del rt_dict['项']
            if '目' in rt_dict: del rt_dict['目']
            if '节' in rt_dict: del rt_dict['节']
            if '细目' in rt_dict: del rt_dict['细目']
    return rt_dict


def _parse_low_sim(p_dict):
    rt_dict = copy.deepcopy(p_dict)
    for fe_key in p_dict.keys():
        if type(p_dict[fe_key]) is not str:
            rt_dict[fe_key].clear()
            for fe_value in p_dict[fe_key]:
                if float(fe_value['sim']) > THRESHOLD:
                    rt_dict[fe_key].append(copy.deepcopy(fe_value))
            if len(rt_dict[fe_key]) == 0:
                del rt_dict[fe_key]
    return rt_dict


# 更新前端进度条
def _stage_update(p_socketio, p_percent, p_stage, p_session_id=None):
    if p_session_id:
        p_socketio.emit('progress', {'progress': p_percent, 'stage': p_stage, 'sessionId': p_session_id})
    else:
        p_socketio.emit('progress', {'progress': p_percent, 'stage': p_stage})
    if p_percent < 100:
        time.sleep(0.1)


def _beautify(p_dict):
    rt_str = '该文件中匹配的总表表单是《' + p_dict['表单名称'] + '》\n 其中：'
    _key_words = TARGET_WORDS_NO + TARGET_WORDS_F8
    for fe_key in _key_words:
        if fe_key in p_dict.keys():
            _str = fe_key + ' 坐标:'
            for fe_value in p_dict[fe_key]:
                _str += '(' + _col_name(fe_value['col']) + ',' + str(fe_value['row'] + 1) + ') '
            if fe_key in TARGET_WORDS_F8 or fe_key == '工程或费用名称':
                rt_str = rt_str + '\n' + _str
            else:
                rt_str = rt_str + _str
    return rt_str


def _col_name(p_int):
    _alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    rt_str = ''
    result_int = int(p_int / 26)
    result_mod = p_int % 26
    if result_int > 0:
        rt_str = _alphabet[result_int-1]
    rt_str += _alphabet[result_mod]
    return rt_str


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
    # socketio.run(app, debug=True, host='0.0.0.0', port=8080)
