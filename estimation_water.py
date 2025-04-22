import os
import time
import pandas as pd
from dataclasses import dataclass
from typing import List, Dict
from flask import jsonify, request, render_template, send_from_directory
from openpyxl.styles import Alignment
from werkzeug.utils import secure_filename
from flask_socketio import SocketIO
from openpyxl.utils import get_column_letter

from pyFCM import fuzzy_match_EM, extract_specifications, init_atlas, fuzzy_match

# 配置上传文件存储路径
UPLOAD_FOLDER = 'uploads/estimation_water'
OUTPUT_FOLDER = 'output/estimation_water'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
UNIT_MAPPING_LEN_MM = {"米": 1000, "m": 1000, "dm": 100, "cm": 10, "mm": 1}

# 确保上传目录和输出目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# socketio = SocketIO() 这里似乎没必要单独初始化这个定义，def init_routes(app, socketio): 这里作为参数传入了一个socketio


@dataclass
class EquipmentMaterial:
    """设备材料信息类"""
    name: str  # 名称
    specification: str  # 规格
    material: str  # 材料
    unit: str  # 单位
    quantity: float  # 数量
    remarks: str  # 备注


Atlas_PipeFittingsQ235A = {}  # 钢管配件重量表
Atlas_PipeFittingsDuctileIron = {}  # 球墨铸铁管配件
Atlas_Valve = {}  # 阀门价格表
Atlas_Equipment = {}


# 读取重量表
def atlas():
    # 读取 Excel 文件中的所有工作表
    tXls = pd.ExcelFile("templates/250407管配件重量表.xlsx")
    for feSheetName in tXls.sheet_names:
        df = pd.read_excel(tXls, sheet_name=feSheetName, header=0, index_col=0)
        # 创建一个字典，用于存储每种管配件的每米重量
        # 遍历表格，提取每米重量
        for _, feRow in df.iterrows():
            # 遍历 feRow 中的每一个单元格
            dn1, dn2 = 0, 0
            for column_name, value in feRow.items():
                if column_name == "管径1":
                    dn1 = int(value)
                elif column_name == "管径2":
                    dn2 = int(value)
                else:
                    if pd.notna(value):  # 使用 pd.notna() 判断值是否不是 NaN
                        if feSheetName == "Q235A":
                            # 如果 column_name 不在 Atlas_PipeFittingsQ235A 中，初始化一个空字典
                            if column_name not in Atlas_PipeFittingsQ235A:
                                Atlas_PipeFittingsQ235A[column_name] = {}
                            # 如果 dn1 不在 Atlas_PipeFittingsQ235A[column_name] 中，初始化一个空字典
                            if dn1 not in Atlas_PipeFittingsQ235A[column_name]:
                                Atlas_PipeFittingsQ235A[column_name][dn1] = {}
                            # 存储 dn2 和对应的重量
                            Atlas_PipeFittingsQ235A[column_name][dn1][dn2] = value
                        elif feSheetName == "球铁":
                            # 如果 column_name 不在 Atlas_PipeFittingsDuctileIron 中，初始化一个空字典
                            if column_name not in Atlas_PipeFittingsDuctileIron:
                                Atlas_PipeFittingsDuctileIron[column_name] = {}
                            # 如果 dn1 不在 Atlas_PipeFittingsDuctileIron[column_name] 中，初始化一个空字典
                            if dn1 not in Atlas_PipeFittingsDuctileIron[column_name]:
                                Atlas_PipeFittingsDuctileIron[column_name][dn1] = {}
                            # 存储 dn2 和对应的重量
                            Atlas_PipeFittingsDuctileIron[column_name][dn1][dn2] = value

    # 读取阀门价格表
    df = pd.read_excel("templates/250410阀门.xlsx", header=0, index_col=0)
    for _, feRow in df.iterrows():
        dn1 = 0
        for column_name, value in feRow.items():
            if column_name == "管径1":
                dn1 = int(value)
            else:
                if pd.notna(value):
                    Atlas_Valve[column_name] = {dn1: value}
    # 读取设备价格表
    df = pd.read_excel("templates/250410设备.xlsx", header=0, index_col=0)
    for _, feRow in df.iterrows():
        tName = ""
        for column_name, value in feRow.items():
            if pd.notna(value):
                if column_name == "设备名称":
                    tName = value
                    Atlas_Equipment[tName] = {}
                else:
                    Atlas_Equipment[tName][column_name] = value
    # print(Atlas_Equipment.keys())
    init_atlas(Atlas_PipeFittingsQ235A, Atlas_PipeFittingsDuctileIron, Atlas_Valve, Atlas_Equipment)


def allowed_file(filename):
    """检查文件类型是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_filename(filename: str) -> str:
    name, ext = os.path.splitext(filename)  # 获取文件扩展名
    safe_name = name.replace('/', '_').replace('\\', '_')  # 替换不安全的字符，但保留中文
    return safe_name + ext  # 重新组合文件名和扩展名


def _stage_update(p_percent, p_stage, p_session_id=None, socketio=None):
    if socketio:
        socketio.emit('progress', {'progress': p_percent, 'stage': p_stage, 'sessionId': p_session_id})
    time.sleep(0)


def process_excel_file(file_path: str, session_id: str, socketio=None) -> Dict[str, List[EquipmentMaterial]]:
    """
    处理Excel文件，提取设备材料表信息
    Args:
        file_path: Excel文件路径
        session_id: 会话ID，用于更新进度
        socketio: SocketIO实例，用于进度更新
    Returns:
        以所属单体为key，设备材料列表为value的字典
    """
    _Target_Words = {"序号": ["序号", "编号"],
                     "所属单体": ["单体", "所属单体", "构筑物", "位置", "设备位置", "安装位置", "安装地点"],
                     "名称": ["名称", "设备名称"],
                     "规格": ["规格", "规格尺寸", "规格参数", "型号规格"],
                     "材料": ["材料", "材质"],
                     "单位": ["单位"],
                     "数量": ["数量"],
                     "备注": ["备注"]}
    # 假设表头包含这些字段
    # _Required_Columns = ["序号", "所属单体", "名称", "规格", "材料", "单位", "数量", "备注"]
    _Required_Columns = ["名称", "规格", "单位", "数量"]

    def _match_row(p_row):
        _target_col = {}
        _target_sim = {}
        for feKey in _Target_Words:
            _target_sim[feKey] = 0.0
        for feCol in range(0, len(p_row)):

            feCellValue = p_row[feCol]
            if pd.isna(feCellValue):
                continue
            # print(feCol, feCellValue)
            _match_key = None
            _match_sim = 0.0
            for feKey in _Target_Words.keys():  # 与目标字段逐个匹配
                _matched_word, _similarity_score = fuzzy_match(feCellValue, _Target_Words[feKey])
                # print("match", _matched_word, _similarity_score)
                if _similarity_score[0] > _match_sim:
                    _match_sim = _similarity_score[0]
                    _match_key = feKey
            # print("ks", feCellValue, _match_key, _match_sim)
            if _match_key is not None:
                if _match_sim > max(_target_sim[_match_key], 0.8):
                    _target_col[_match_key] = feCol
                    _target_sim[_match_key] = _match_sim

        return _target_col

    try:

        # 使用 openpyxl 引擎读取 Excel 文件
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')

        result_dict = {}
        total_sheets = len(excel_file.sheet_names)

        # 遍历所有工作表
        for feSheetIndex, feSheetName in enumerate(excel_file.sheet_names):
            progress = 10 + (feSheetIndex / total_sheets) * 60  # 10-70% 的进度
            _stage_update(progress, f'正在处理工作表: {feSheetName}……', session_id, socketio)

            df_sheet = pd.read_excel(excel_file, sheet_name=feSheetName, engine='openpyxl', header=None)
            _match_head_row = 0
            _match_head_sim = 0.0
            current_row = 0
            _key_exchange = {}
            for feRowIndex, feRow in df_sheet.iterrows():
                t_target_col = _match_row(feRow)
                _match_head_row = feRowIndex
                _key_exchange = t_target_col
                current_row = current_row + 1
                if all(col in _key_exchange.keys() for col in _Required_Columns) or current_row > 10:
                    break

            print("_match_head_row", _match_head_row)
            print(_key_exchange)

            df_sheet = pd.read_excel(excel_file, sheet_name=feSheetName, engine='openpyxl', header=_match_head_row)
            if all(col in _key_exchange.keys() for col in _Required_Columns):  # 判断要求的字段是否都在识别结果中能找到
                # 找到目标表格
                last_individual = ""
                last_EM = None
                total_rows = len(df_sheet)

                for feRowIndex, (_, feRow) in enumerate(df_sheet.iterrows()):
                    if "所属单体" not in _key_exchange.keys():
                        individual = feSheetName
                    else:
                        # 处理单体单元格合并情况
                        individual = feRow.iloc[_key_exchange["所属单体"]]
                    if pd.isna(individual):
                        if last_individual == "":
                            continue
                        else:
                            individual = last_individual
                    else:
                        last_individual = individual
                    # 处理名称单元格合并情况

                    tEMname = feRow.iloc[_key_exchange["名称"]]
                    if pd.isna(tEMname):
                        if pd.isna(feRow.iloc[_key_exchange["数量"]]):
                            if last_EM is not None:
                                last_EM.specification += str(feRow.iloc[_key_exchange["规格"]])
                            continue
                        else:
                            if last_EM is not None:
                                tEMname = last_EM.name
                            else:
                                continue

                    tQList = str(feRow.iloc[_key_exchange["数量"]]).split('/')  # 如果有多个规格写在同一行的，例：墙管 DN500/DN300 个 40/91
                    tSpStr = str(feRow.iloc[_key_exchange["规格"]])
                    tSpList = []
                    if len(tQList) > 1:
                        tSpList = tSpStr.split('/')
                    else:
                        tSpList.append(str(feRow.iloc[_key_exchange["规格"]]))

                    for i in range(0, len(tQList)):
                        tQ = tQList[i]
                        if i <= len(tSpList) - 1:
                            tSp = tSpList[i]
                        else:
                            tSp = tSpList[len(tSpList) - 1]
                        tMaterial = ""
                        if "材料" in _key_exchange.keys():
                            if pd.notna(feRow.iloc[_key_exchange["材料"]]):
                                tMaterial = str(feRow.iloc[_key_exchange["材料"]])
                        tRemark = ""
                        if "备注" in _key_exchange.keys():
                            if pd.notna(feRow.iloc[_key_exchange["备注"]]):
                                tRemark = str(feRow.iloc[_key_exchange["备注"]])
                        # 创建设备材料对象
                        tEM = EquipmentMaterial(
                            name=str(tEMname),
                            specification=str(tSp),
                            material=tMaterial,
                            unit=str(feRow.iloc[_key_exchange["单位"]]),
                            quantity=tQ if pd.notna(tQ) else 0.0,
                            remarks=tRemark
                        )
                        # 将设备材料添加到对应单体的列表中
                        if individual not in result_dict:
                            result_dict[individual] = []
                        result_dict[individual].append(tEM)
                        last_EM = tEM

        return result_dict

        # raise ValueError("未找到符合要求的设备材料表")

    except Exception as e:
        _stage_update(0, f"处理Excel文件时出错: {str(e)}", session_id, socketio)
        raise Exception(f"处理Excel文件时出错: {str(e)}")


def write_to_excel(equipment_dict: Dict[str, List[EquipmentMaterial]], original_filename: str) -> str:
    _individual_sum_row = {}
    # 写入总表
    def write_to_excle_summary():
        def cell_format(pWorksheet, pTemplate_ws, pCurrent_row, name_col_idx, sum_col_idx, pValue):
            pCell = pWorksheet.cell(row=pCurrent_row, column=name_col_idx)
            pCell.value = pValue
            pCell.alignment = Alignment(horizontal='right', vertical='center')
            cell = pWorksheet.cell(row=pCurrent_row, column=sum_col_idx)
            cell.value = f"=SUM(D{pCurrent_row}:G{pCurrent_row})"
            row_formate(pWorksheet, pTemplate_ws, pCurrent_row)

        def row_formate(pWorksheet, pTemplate_ws, pRow):
            pWorksheet.row_dimensions[pRow].height = pTemplate_ws.row_dimensions[7].height  # 调整行高
            for feCol in range(1, 13):
                tCell = pWorksheet.cell(row=pRow, column=feCol)
                tCell.border = pTemplate_ws.cell(row=8, column=feCol).border.copy()  # 使用模板的边框样式
                if feCol != 3:
                    tCell.alignment = pTemplate_ws.cell(row=8, column=feCol).alignment.copy()  # 使用模板的对齐方式

        worksheet = writer.sheets['总表']  # 获取工作表对象
        # 复制模板的前7行格式
        for feRow in range(1, 8):  # 复制前7行
            for feCol in range(1, len(template_ws[1]) + 1):
                # 获取模板单元格
                template_cell = template_ws.cell(row=feRow, column=feCol)
                # 获取目标单元格
                target_cell = worksheet.cell(row=feRow, column=feCol)

                # 复制单元格值
                target_cell.value = template_cell.value

                # 复制单元格格式
                if template_cell.has_style:
                    # 复制字体
                    target_cell.font = template_cell.font.copy()
                    # 复制边框
                    target_cell.border = template_cell.border.copy()
                    # 复制填充
                    target_cell.fill = template_cell.fill.copy()
                    # 复制数字格式
                    target_cell.number_format = template_cell.number_format
                    # 复制保护
                    target_cell.protection = template_cell.protection.copy()
                    # 复制对齐方式
                    target_cell.alignment = template_cell.alignment.copy()
        # 复制合并单元格
        for merged_range in template_ws.merged_cells.ranges:
            if merged_range.min_row <= 7:  # 只复制前7行的合并单元格
                worksheet.merge_cells(str(merged_range))
        # 复制列宽
        for feCol in range(1, len(template_ws[1]) + 1):
            column_letter = get_column_letter(feCol)
            worksheet.column_dimensions[column_letter].width = template_ws.column_dimensions[column_letter].width
            # print(column_letter, template_ws.column_dimensions[column_letter].width)
        # 复制行高
        for feRow in range(1, 8):  # 复制前7行的行高
            worksheet.row_dimensions[feRow].height = template_ws.row_dimensions[feRow].height
        # 找到"工程或费用名称"列的索引
        name_col_idx = 3
        sum_col_idx = 8
        # 从第8行开始写入数据
        current_row = 8

        # 写入每个key，并在key之间添加3个空行
        for feIndivName in equipment_dict.keys():
            # 写入key
            cell = worksheet.cell(row=current_row, column=name_col_idx)
            cell.value = feIndivName
            cell = worksheet.cell(row=current_row, column=sum_col_idx)
            cell.value = f"=SUM(D{current_row}:G{current_row})"
            row_formate(worksheet, template_ws, current_row)
            current_row += 1
            cell_format(worksheet, template_ws, current_row, name_col_idx, sum_col_idx, "土建")
            current_row += 1
            cell_format(worksheet, template_ws, current_row, name_col_idx, sum_col_idx, "管配件")
            cell = worksheet.cell(row=current_row, column=5)
            cell.value = f"=ROUND('{feIndivName}gpj'!H{_individual_sum_row[feIndivName + 'gpj']}/10000,2)"
            current_row += 1
            cell_format(worksheet, template_ws, current_row, name_col_idx, sum_col_idx, "设备")
            cell = worksheet.cell(row=current_row, column=5)
            cell.value = f"=ROUND('{feIndivName}sb'!H{_individual_sum_row[feIndivName + 'sb']}/10000,2)"
            for feCol in range(4, 8):
                cell = worksheet.cell(row=current_row - 3, column=feCol)
                cell.value = f"=SUM({get_column_letter(feCol)}{current_row - 2}:{get_column_letter(feCol)}{current_row})"

            current_row += 1

        for feCol in range(4, 8):
            cell = worksheet.cell(row=7, column=feCol)
            cell.value = f"=SUM({get_column_letter(feCol)}{7 + 1}:{get_column_letter(feCol)}{current_row})"
        cell = worksheet.cell(row=7, column=8)
        cell.value = "=SUM(D7:G7)"

    # 写入单项概算
    def write_to_excel_individual():
        def find_closest_key(random_number, dictionary):
            """
            找到与随机整数差值最小的字典键。
            参数:
                random_number (int): 随机整数
                dictionary (dict): 字典，键是整数
            返回:
                int: 与随机整数差值最小的键
            """
            # 初始化最小差值和最接近的键
            min_diff = float('inf')  # 设置为无穷大
            closest_key = None

            # 遍历字典的键
            for feKey in dictionary.keys():
                # 计算差值
                diff = abs(feKey - random_number)
                # 更新最小差值和最接近的键
                if diff < min_diff:
                    min_diff = diff
                    closest_key = feKey

            return closest_key

        # =============================================设备材料表================================================================
        category = {"gpj": ["管配件", "材料"], "sb": ["设备"]}
        for feIndivName in equipment_dict.keys():
            for feSuffix in category.keys():
                # 创建第二个工作表（设备材料表）
                template_ws2 = template_wb['Sheet2']  # 获取模板的第二个工作表
                worksheet2 = workbook.create_sheet(feIndivName + feSuffix)  # 创建新的工作表
                # 复制Sheet2的前7行格式
                for feRow in range(1, 8):
                    for feCol in range(1, len(template_ws2[1]) + 1):
                        template_cell = template_ws2.cell(row=feRow, column=feCol)  # 获取模板单元格
                        target_cell = worksheet2.cell(row=feRow, column=feCol)  # 获取目标单元格

                        # 复制单元格值
                        if feRow == 3 and feCol == 2:
                            target_cell.value = feIndivName + " " + category[feSuffix][0]
                        else:
                            target_cell.value = template_cell.value

                        # 复制单元格格式
                        if template_cell.has_style:
                            target_cell.font = template_cell.font.copy()  # 复制字体
                            target_cell.border = template_cell.border.copy()  # 复制边框
                            target_cell.fill = template_cell.fill.copy()  # 复制填充
                            target_cell.number_format = template_cell.number_format  # 复制数字格式
                            target_cell.protection = template_cell.protection.copy()  # 复制保护
                            target_cell.alignment = template_cell.alignment.copy()  # 复制对齐方式

                # 复制Sheet2的合并单元格
                for feMergedRange in template_ws2.merged_cells.ranges:
                    if feMergedRange.min_row <= 7:  # 只复制前7行的合并单元格
                        worksheet2.merge_cells(str(feMergedRange))

                # 复制Sheet2的列宽
                for feCol in range(1, len(template_ws2[1]) + 1):
                    column_letter = get_column_letter(feCol)
                    if column_letter in template_ws2.column_dimensions:
                        worksheet2.column_dimensions[column_letter].width = template_ws2.column_dimensions[
                            column_letter].width

                # 复制Sheet2的行高
                for feRow in range(1, 8):  # 复制前7行的行高
                    worksheet2.row_dimensions[feRow].height = template_ws2.row_dimensions[feRow].height

                current_row = 8

                for feEM in equipment_dict[feIndivName]:
                    tBM, tFlange, tMaterial, tScore, tType = fuzzy_match_EM(feEM)
                    tResult = extract_specifications(feEM.specification)
                    dn1 = 0
                    dn2 = 0
                    tValue = ""
                    tPrice = 1
                    tPriceFlange = 1
                    tDensity = ""  # 与铁的容重比
                    tAtlas = Atlas_PipeFittingsQ235A
                    if tMaterial in ["Q235A", "Q235B", "Q235C", "Q235D", "Q235E"]:
                        tPrice = 1
                        tPriceFlange = 1
                        tDensity = ""
                        tAtlas = Atlas_PipeFittingsQ235A
                    elif tMaterial in ["SS304"]:
                        tPrice = 3
                        tPriceFlange = 1
                        tDensity = "*7.93/7.85"
                        tAtlas = Atlas_PipeFittingsQ235A
                    elif tMaterial in ["SS316"]:
                        tPrice = 5
                        tPriceFlange = 1
                        tDensity = "*8.0/7.85"
                        tAtlas = Atlas_PipeFittingsQ235A
                    elif tMaterial in ["球铁"]:
                        tPrice = 7
                        tPriceFlange = 0
                        tAtlas = Atlas_PipeFittingsDuctileIron

                    if len(tResult["管径"]) > 0:
                        dn1 = tResult["管径"][0]
                        dn2 = dn1
                        if len(tResult["管径"]) > 1:
                            dn2 = tResult["管径"][1]
                    if tScore > 0:
                        if tBM in tAtlas.keys():
                            tDic = tAtlas[tBM][find_closest_key(dn1, tAtlas[tBM])]
                            tFlangeDn1 = find_closest_key(dn1, tAtlas["法兰"])
                            tFlangeWeight = tAtlas["法兰"][tFlangeDn1][tFlangeDn1]
                            tCircleStr = ""
                            tLengthStr = ""

                            if tBM in ["直管", "套管",
                                       "穿墙套管", "柔性防水套管A型Ⅰ型", "柔性防水套管A型Ⅱ型", "柔性防水套管B型Ⅰ型",
                                       "柔性防水套管B型Ⅱ型", "法兰套管A型Ⅰ型", "法兰套管A型Ⅱ型", "法兰套管B型Ⅰ型",
                                       "法兰套管B型Ⅱ型", "刚性防水套管A型", "刚性防水套管B型", "刚性防水套管C型"]:
                                if feEM.unit not in UNIT_MAPPING_LEN_MM.keys():
                                    tLengthStr = "*" + str(tResult["长度"]) + "/1000"
                                if tBM == "穿墙套管":
                                    tCircle = Atlas_PipeFittingsQ235A["穿墙套管-配件"][
                                        find_closest_key(dn1, Atlas_PipeFittingsQ235A["穿墙套管-配件"])]
                                    tCircleStr = "+" + str(tCircle[find_closest_key(dn2, tCircle)])

                            tValue = (
                                f"=({tDic[find_closest_key(dn2, tDic)]}{tLengthStr}{tCircleStr})/1000{tDensity}*K{tPrice}"
                                f"+{tFlange}*{tFlangeWeight}/1000{tDensity}*K{tPrice + tPriceFlange}")

                        if tType == "阀门" and tResult["功率"] == 0.0:
                            if len(tResult["管径"]) > 0:
                                if tResult["管径"][0] >= 600:
                                    tType = "设备"
                                else:
                                    tType = "材料"
                            else:
                                tType = "材料"
                    if tResult["功率"] > 0:
                        tType = "设备"
                    if tType == "":
                        tType = "材料"
                    if tType not in category[feSuffix]:
                        continue
                    Cell = worksheet2.cell(row=current_row, column=2)
                    Cell.value = current_row - 7
                    Cell = worksheet2.cell(row=current_row, column=3)
                    Cell.value = feEM.name
                    Cell = worksheet2.cell(row=current_row, column=4)
                    if feEM.material != "nan":
                        Cell.value = f"{feEM.specification} {feEM.material}"
                    else:
                        Cell.value = feEM.specification
                    Cell = worksheet2.cell(row=current_row, column=5)
                    Cell.value = feEM.unit
                    Cell = worksheet2.cell(row=current_row, column=6)
                    Cell.value = feEM.quantity
                    Cell = worksheet2.cell(row=current_row, column=7)
                    Cell.value = str(tValue)
                    Cell = worksheet2.cell(row=current_row, column=8)
                    Cell.value = f"=F{current_row}*G{current_row}"

                    Cell = worksheet2.cell(row=current_row, column=13)
                    Cell.value = tBM
                    Cell = worksheet2.cell(row=current_row, column=12)
                    Cell.value = tScore
                    Cell = worksheet2.cell(row=current_row, column=14)
                    if tResult["长度"] != 0:
                        Cell.value = f"DN{dn1}×DN{dn2} L=" + str(tResult["长度"]) + str(tResult["长度单位"])
                    else:
                        Cell.value = f"DN{dn1}×DN{dn2}"

                    Cell = worksheet2.cell(row=current_row, column=17)
                    Cell.value = str(tType)

                    worksheet2.row_dimensions[current_row].height = template_ws2.row_dimensions[8].height
                    for feCol in range(1, 9):
                        Cell = worksheet2.cell(row=current_row, column=feCol)
                        template_cell = template_ws2.cell(row=8, column=feCol)
                        Cell.font = template_cell.font.copy()  # 复制字体
                        Cell.border = template_cell.border.copy()  # 复制边框
                        Cell.fill = template_cell.fill.copy()  # 复制填充
                        Cell.number_format = template_cell.number_format  # 复制数字格式
                        Cell.protection = template_cell.protection.copy()  # 复制保护
                        Cell.alignment = template_cell.alignment.copy()  # 复制对齐方式

                    current_row += 1

                for feRow in range(current_row, current_row + 7):
                    worksheet2.row_dimensions[feRow].height = template_ws2.row_dimensions[23].height
                    for feCol in range(1, 9):
                        Cell = worksheet2.cell(row=feRow, column=feCol)
                        template_cell = template_ws2.cell(row=feRow - current_row + 23, column=feCol)
                        Cell.font = template_cell.font.copy()  # 复制字体
                        Cell.border = template_cell.border.copy()  # 复制边框
                        Cell.fill = template_cell.fill.copy()  # 复制填充
                        Cell.number_format = template_cell.number_format  # 复制数字格式
                        Cell.protection = template_cell.protection.copy()  # 复制保护
                        Cell.alignment = template_cell.alignment.copy()  # 复制对齐方式
                        if feCol in [3, 4]:
                            Cell.value = template_cell.value
                        elif feCol == 5:
                            if feRow - current_row not in [0, 6]:
                                Cell.value = "元"
                        elif feCol == 8:
                            if feRow - current_row == 1:
                                Cell.value = f"=SUM(H8:H{feRow - 1})"
                            elif feRow - current_row in [2, 4]:
                                Cell.value = f"=H{feRow - 1}*D{feRow}"
                            elif feRow - current_row in [3, 5]:
                                Cell.value = f"=SUM(H{feRow - 2}:H{feRow - 1})"
                Cell = worksheet2.cell(row=4, column=2)
                Cell.value = f'="估算价值(元)："&ROUND(H{current_row + 5},0)'
                _individual_sum_row[feIndivName + feSuffix] = current_row + 5
    #  函数本体从这里开始执行
    try:
        timestamp = time.strftime("%y%m%d%H%M%S")  # 生成时间戳格式的文件名 (YYMMDDHHMMSS)
        safe_original_filename = safe_filename(original_filename)  # 使用安全的文件名处理函数，但保留原始字符
        new_filename = f"{timestamp}_{safe_original_filename}"  # 构建新文件名
        output_path = os.path.join(OUTPUT_FOLDER, new_filename)

        template_path = os.path.join('templates', 'standard.xlsx')  # 读取标准模板
        # 使用openpyxl直接读取模板文件以保留格式
        from openpyxl import load_workbook
        template_wb = load_workbook(template_path)
        template_ws = template_wb['Sheet1']

        df = pd.DataFrame()  # 创建一个空的DataFrame作为初始工作表
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:  # 创建一个新的Excel写入器
            df.to_excel(writer, sheet_name='总表', index=False)  # 先写入空的DataFrame以创建工作表
            workbook = writer.book  # 获取工作簿
            write_to_excel_individual()
            write_to_excle_summary()
        return output_path

    except Exception as e:
        raise Exception(f"写入Excel文件时出错: {str(e)}")


def init_routes(app, socketio):
    """初始化路由和Socket.IO事件处理"""

    @app.route('/estimation/water')
    def estimation_water():
        """水厂智能估算页面路由"""
        atlas()
        return render_template('estimation_water.html')

    @app.route('/estimation/water/upload', methods=['POST'])
    def upload_water_file():
        """处理文件上传"""
        if 'file0' not in request.files:
            return jsonify({'error': '没有文件被上传'}), 400

        file = request.files['file0']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': '不支持的文件类型'}), 400

        # 获取会话ID
        session_id = request.headers.get('X-Session-ID')

        try:
            # 使用安全的文件名处理函数，但保留原始字符
            filename = safe_filename(file.filename)
            _stage_update(5, f'文件 {file.filename} 上传成功！', session_id, socketio)

            # 添加时间戳确保文件名唯一
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            new_filename = f"{timestamp}_{filename}"
            file_path = os.path.join(UPLOAD_FOLDER, new_filename)

            file.save(file_path)  # 保存文件
            _stage_update(10, f'文件 {file.filename} 保存成功！，数据读取中……', session_id, socketio)  # 发送进度更新

            equipment_dict = process_excel_file(file_path, session_id, socketio)  # 处理Excel文件，传入session_id
            _stage_update(70, f'文件 {file.filename} 读取完成，正在生成结果文件……', session_id, socketio)  # 发送进度更新

            output_path = write_to_excel(equipment_dict, file.filename)  # 使用原始文件名，写入新的Excel文件

            _stage_update(100, f'处理完成，已输出文件: {os.path.basename(output_path)}', session_id, socketio)  # 发送进度更新

            return jsonify({
                'message': '文件上传成功',
                'filename': new_filename,
                'progress': 100,
                'output_file': os.path.basename(output_path),
                'equipment_data': {
                    unit: [
                        {
                            'name': eq.name,
                            'specification': eq.specification,
                            'material': eq.material,
                            'unit': eq.unit,
                            'quantity': eq.quantity,
                            'remarks': eq.remarks
                        }
                        for eq in equipment_list
                    ]
                    for unit, equipment_list in equipment_dict.items()
                }
            })

        except Exception as e:
            _stage_update(0, f'上传失败: {str(e)}', session_id, socketio)  # 发送错误信息
            return jsonify({'error': str(e)}), 500

    @app.route('/estimation/water/download/<filename>')
    def download_file(filename):
        """处理文件下载"""
        try:
            return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)
        except Exception as e:
            return jsonify({'error': str(e)}), 500

    @socketio.on('upload')
    def handle_upload(data):
        """处理上传开始事件"""
        filename = data.get('filename')
        session_id = data.get('sessionId')
        print(f'Upload started for file: {filename} (Session: {session_id})')
