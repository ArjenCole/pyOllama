import openpyxl
import pandas as pd


def get_workbook(p_file_path):
    _data_frame_dict = pd.read_excel(p_file_path, sheet_name=None)
    return _data_frame_dict


def table_format(p_dict,p_data_frame):
    _json = {"row": 5, "column": 6}
    print(type(_json))