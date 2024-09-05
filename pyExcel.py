import openpyxl
import pandas as pd


def extract_headers(filename, sheetname=None):
    # 加载Excel工作簿
    wb = openpyxl.load_workbook(filename)

    # 选择工作表
    if sheetname:
        ws = wb[sheetname]
    else:
        ws = wb.active

    # 初始化表头数据
    header_data = []

    # 假设表头在前几行，检测包含文字的单元格
    for row in ws.iter_rows(min_row=1, max_row=4):  # 假设表头在前4行之内
        for cell in row:
            if cell.value is not None:
                header_data.append(cell.value)

    # 打印表头数据
    return header_data


def get_workbook(p_file_path):
    _data_frame_dict = pd.read_excel(p_file_path, sheet_name=None)
    return _data_frame_dict
