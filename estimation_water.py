import os
import time
import pandas as pd
from dataclasses import dataclass
from typing import List, Dict
from flask import jsonify, request, render_template
from openpyxl.styles import Alignment
from werkzeug.utils import secure_filename
from flask_socketio import SocketIO
from openpyxl.utils import get_column_letter

from pyFCM import fuzzy_match_EM, extract_specifications, init_atlas

# 配置上传文件存储路径
UPLOAD_FOLDER = 'uploads/estimation_water'
OUTPUT_FOLDER = 'output/estimation_water'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 确保上传目录和输出目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# 初始化 SocketIO
socketio = SocketIO()


@dataclass
class EquipmentMaterial:
    """设备材料信息类"""
    name: str  # 名称
    specification: str  # 规格
    material: str  # 材料
    unit: str  # 单位
    quantity: float  # 数量
    remarks: str  # 备注


Atlas_PipeFittingsQ235A = {}  # 管配件重量表
Atlas_Valve = {}  # 阀门价格表

# 读取重量表
def atlas():
    df = pd.read_excel("templates/250407管配件重量表.xlsx", header=0, index_col=0)
    # 创建一个字典，用于存储每种管配件的每米重量
    # 遍历表格，提取每米重量
    for index, row in df.iterrows():
        # 遍历 row 中的每一个单元格
        dn1 = 0
        dn2 = 0
        for column_name, value in row.items():
            if column_name == "管径1":
                dn1 = int(value)
            elif column_name == "管径2":
                dn2 = int(value)
            else:
                if pd.notna(value):  # 使用 pd.notna() 判断值是否不是 NaN
                    # 如果 column_name 不在 Atlas_PipeFittingsQ235A 中，初始化一个空字典
                    if column_name not in Atlas_PipeFittingsQ235A:
                        Atlas_PipeFittingsQ235A[column_name] = {}
                    # 如果 dn1 不在 Atlas_PipeFittingsQ235A[column_name] 中，初始化一个空字典
                    if dn1 not in Atlas_PipeFittingsQ235A[column_name]:
                        Atlas_PipeFittingsQ235A[column_name][dn1] = {}
                    # 存储 dn2 和对应的重量
                    Atlas_PipeFittingsQ235A[column_name][dn1][dn2] = value

    # 读取阀门价格表
    df = pd.read_excel("templates/250410阀门.xlsx", header=0, index_col=0)
    for index, row in df.iterrows():
        dn1 = 0
        for column_name, value in row.items():
            if column_name == "管径1":
                dn1 = int(value)
            else:
                if pd.notna(value):
                    Atlas_Valve[column_name] = {dn1: value}

    init_atlas(Atlas_PipeFittingsQ235A, Atlas_Valve)


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
    for key in dictionary.keys():
        # 计算差值
        diff = abs(key - random_number)
        # 更新最小差值和最接近的键
        if diff < min_diff:
            min_diff = diff
            closest_key = key

    return closest_key


def allowed_file(filename):
    """检查文件类型是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_filename(filename: str) -> str:
    """
    生成安全的文件名，但保留原始字符（包括中文）
    
    Args:
        filename: 原始文件名
        
    Returns:
        安全的文件名
    """
    # 获取文件扩展名
    name, ext = os.path.splitext(filename)

    # 替换不安全的字符，但保留中文
    safe_name = name.replace('/', '_').replace('\\', '_')

    # 重新组合文件名和扩展名
    return safe_name + ext


def process_excel_file(file_path: str) -> Dict[str, List[EquipmentMaterial]]:
    """
    处理Excel文件，提取设备材料表信息
    
    Args:
        file_path: Excel文件路径
        
    Returns:
        以所属单体为key，设备材料列表为value的字典
    """
    try:
        # 使用openpyxl引擎读取Excel文件
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')

        # 查找设备材料表
        # 假设表头包含这些字段
        required_columns = ["序号", "所属单体", "名称", "规格", "材料", "单位", "数量", "备注"]

        # 遍历所有工作表
        for sheet_name in excel_file.sheet_names:
            df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
            if all(col in df_sheet.columns for col in required_columns):
                # 找到目标表格
                result_dict = {}

                # 遍历每一行
                for _, row in df_sheet.iterrows():
                    unit = row["所属单体"]
                    if pd.isna(unit):  # 跳过空行
                        continue

                    # 创建设备材料对象
                    equipment = EquipmentMaterial(
                        name=str(row["名称"]),
                        specification=str(row["规格"]),
                        material=str(row["材料"]),
                        unit=str(row["单位"]),
                        quantity=float(row["数量"]) if pd.notna(row["数量"]) else 0.0,
                        remarks=str(row["备注"]) if pd.notna(row["备注"]) else ""
                    )

                    # 将设备材料添加到对应单体的列表中
                    if unit not in result_dict:
                        result_dict[unit] = []
                    result_dict[unit].append(equipment)

                return result_dict

        raise ValueError("未找到符合要求的设备材料表")

    except Exception as e:
        raise Exception(f"处理Excel文件时出错: {str(e)}")


def write_to_excel(equipment_dict: Dict[str, List[EquipmentMaterial]], original_filename: str) -> str:
    """
    将设备材料字典数据写入新的Excel文件，使用标准模板格式
    
    Args:
        equipment_dict: 设备材料字典数据
        original_filename: 原始文件名
        
    Returns:
        新生成的Excel文件路径
    """
    try:
        # 生成时间戳格式的文件名 (YYMMDDHHMMSS)
        timestamp = time.strftime("%y%m%d%H%M%S")

        # 使用安全的文件名处理函数，但保留原始字符
        safe_original_filename = safe_filename(original_filename)

        # 构建新文件名
        new_filename = f"{timestamp}_{safe_original_filename}"
        output_path = os.path.join(OUTPUT_FOLDER, new_filename)

        # 读取标准模板
        template_path = os.path.join('templates', 'standard.xlsx')

        # 使用openpyxl直接读取模板文件以保留格式
        from openpyxl import load_workbook
        template_wb = load_workbook(template_path)
        template_ws = template_wb['Sheet1']

        # 创建一个空的DataFrame作为初始工作表
        df = pd.DataFrame()

        # 创建一个新的Excel写入器
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 先写入空的DataFrame以创建工作表
            df.to_excel(writer, sheet_name='总表', index=False)
            # 获取工作簿和工作表对象
            workbook = writer.book
            worksheet = writer.sheets['总表']
            # 复制模板的前7行格式
            for row in range(1, 8):  # 复制前7行
                for col in range(1, len(template_ws[1]) + 1):
                    # 获取模板单元格
                    template_cell = template_ws.cell(row=row, column=col)
                    # 获取目标单元格
                    target_cell = worksheet.cell(row=row, column=col)

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
            for col in range(1, len(template_ws[1]) + 1):
                column_letter = get_column_letter(col)
                worksheet.column_dimensions[column_letter].width = template_ws.column_dimensions[column_letter].width
                # print(column_letter, template_ws.column_dimensions[column_letter].width)
            # 复制行高
            for row in range(1, 8):  # 复制前7行的行高
                worksheet.row_dimensions[row].height = template_ws.row_dimensions[row].height
            # 找到"工程或费用名称"列的索引
            name_col_idx = 3
            sum_col_idx = 8
            # 从第8行开始写入数据
            current_row = 8

            # 写入每个key，并在key之间添加3个空行
            for key in equipment_dict.keys():
                # 写入key
                cell = worksheet.cell(row=current_row, column=name_col_idx)
                cell.value = key
                cell = worksheet.cell(row=current_row, column=sum_col_idx)
                cell.value = f"=SUM(D{current_row}:G{current_row})"
                row_formate(worksheet, template_ws, current_row)
                current_row += 1
                cell_format(worksheet, template_ws, current_row, name_col_idx, sum_col_idx, "土建")
                current_row += 1
                cell_format(worksheet, template_ws, current_row, name_col_idx, sum_col_idx, "管配件")
                current_row += 1
                cell_format(worksheet, template_ws, current_row, name_col_idx, sum_col_idx, "设备")

                for feCol in range(4, 8):
                    cell = worksheet.cell(row=current_row - 3, column=feCol)
                    cell.value = f"=SUM({get_column_letter(feCol)}{current_row - 2}:{get_column_letter(feCol)}{current_row})"

                current_row += 1

            for feCol in range(4, 8):
                cell = worksheet.cell(row=7, column=feCol)
                cell.value = f"=SUM({get_column_letter(feCol)}{7 + 1}:{get_column_letter(feCol)}{current_row})"
            cell = worksheet.cell(row=7, column=8)
            cell.value = "=SUM(D7:G7)"
# =============================================设备材料表================================================================
            category = {"gpj": ["管配件", "材料"], "sb": ["设备"]}
            for key in equipment_dict.keys():
                for feSheetname in category.keys():
                    # 创建第二个工作表（设备材料表）
                    template_ws2 = template_wb['Sheet2']  # 获取模板的第二个工作表
                    worksheet2 = workbook.create_sheet(key + feSheetname)  # 创建新的工作表
                    for row in range(1, 8):  # 复制Sheet2的前7行格式
                        for col in range(1, len(template_ws2[1]) + 1):
                            template_cell = template_ws2.cell(row=row, column=col)  # 获取模板单元格
                            target_cell = worksheet2.cell(row=row, column=col)  # 获取目标单元格

                            # 复制单元格值
                            if row == 3 and col == 2:
                                target_cell.value = key + " " + category[feSheetname][0]
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
                    for merged_range in template_ws2.merged_cells.ranges:
                        if merged_range.min_row <= 7:  # 只复制前7行的合并单元格
                            worksheet2.merge_cells(str(merged_range))

                    # 复制Sheet2的列宽
                    for col in range(1, len(template_ws2[1]) + 1):
                        column_letter = get_column_letter(col)
                        if column_letter in template_ws2.column_dimensions:
                            worksheet2.column_dimensions[column_letter].width = template_ws2.column_dimensions[
                                column_letter].width

                    # 复制Sheet2的行高
                    for row in range(1, 8):  # 复制前7行的行高
                        worksheet2.row_dimensions[row].height = template_ws2.row_dimensions[row].height

                    current_row = 8

                    for feEM in equipment_dict[key]:
                        tBM, tFlange, tMaterial, tScore, tType = fuzzy_match_EM(feEM)
                        tResult = extract_specifications(feEM.specification)
                        dn1 = 0
                        dn2 = 0
                        tValue = ""
                        tPrice = 1
                        if tMaterial in ["Q235A", "Q235B", "Q235C", "Q235D", "Q235E"]:
                            tPrice = 1
                        elif tMaterial in ["SS304"]:
                            tPrice = 3
                        elif tMaterial in ["SS316"]:
                            tPrice = 5

                        if tScore > 0:
                            if tBM in Atlas_PipeFittingsQ235A.keys():
                                if len(tResult["管径"]) == 0:
                                    continue
                                dn1 = tResult["管径"][0]
                                dn2 = dn1
                                if len(tResult["管径"]) > 1:
                                    dn2 = tResult["管径"][1]
                                tDic = Atlas_PipeFittingsQ235A[tBM][find_closest_key(dn1, Atlas_PipeFittingsQ235A[tBM])]
                                tFlangeDn1 = find_closest_key(dn1, Atlas_PipeFittingsQ235A["法兰"])
                                tFlangeWeight = Atlas_PipeFittingsQ235A["法兰"][tFlangeDn1][tFlangeDn1]
                                tValue = (f"={tDic[find_closest_key(dn2, tDic)]}/1000*K{tPrice}"
                                          f"+{tFlange}*{tFlangeWeight}/1000*K{tPrice+1}")


                            if tType == "阀门" and tResult["功率"] == 0.0:
                                if len(tResult["管径"]) > 0:
                                    if tResult["管径"][0] >= 600:
                                        tType = "设备"
                                    else:
                                        tType = "管配件"
                                else:
                                    tType = "材料"
                        if tResult["功率"] > 0:
                            tType = "设备"
                        if tType == "":
                            tType = "材料"
                        if tType not in category[feSheetname]:
                            continue
                        Cell = worksheet2.cell(row=current_row, column=2)
                        Cell.value = current_row - 7
                        Cell = worksheet2.cell(row=current_row, column=3)
                        Cell.value = feEM.name
                        Cell = worksheet2.cell(row=current_row, column=4)
                        Cell.value = feEM.specification
                        Cell = worksheet2.cell(row=current_row, column=5)
                        Cell.value = feEM.unit
                        Cell = worksheet2.cell(row=current_row, column=6)
                        Cell.value = feEM.quantity
                        Cell = worksheet2.cell(row=current_row, column=8)
                        Cell.value = f"=F{current_row}*G{current_row}"
                        Cell = worksheet2.cell(row=current_row, column=10)
                        Cell.value = feEM.material
                        Cell = worksheet2.cell(row=current_row, column=11)
                        Cell.value = tBM
                        Cell = worksheet2.cell(row=current_row, column=12)
                        Cell.value = tScore
                        Cell = worksheet2.cell(row=current_row, column=13)
                        for feDN in tResult["管径"]:
                            if Cell.value is None:
                                Cell.value = ""
                            Cell.value = str(Cell.value) + " " + str(feDN)
                        Cell = worksheet2.cell(row=current_row, column=14)
                        Cell.value = str(tResult["长度"]) + " " + str(tResult["长度单位"])
                        Cell = worksheet2.cell(row=current_row, column=15)
                        Cell.value = f"{tBM} DN{dn1}×DN{dn2}"
                        Cell = worksheet2.cell(row=current_row, column=7)
                        Cell.value = str(tValue)
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
        return output_path

    except Exception as e:
        raise Exception(f"写入Excel文件时出错: {str(e)}")


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
            # 发送进度更新
            socketio.emit('progress', {
                'progress': 5,
                'stage': f'文件 {file.filename} 上传成功！',
                'sessionId': session_id
            })
            socketio.sleep(0)

            # 添加时间戳确保文件名唯一
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            new_filename = f"{timestamp}_{filename}"
            file_path = os.path.join(UPLOAD_FOLDER, new_filename)

            # 保存文件
            file.save(file_path)
            # 发送进度更新
            socketio.emit('progress', {
                'progress': 10,
                'stage': f'文件 {file.filename} 保存成功！',
                'sessionId': session_id
            })
            socketio.sleep(0)

            # 处理Excel文件
            equipment_dict = process_excel_file(file_path)
            # 发送进度更新
            socketio.emit('progress', {
                'progress': 10,
                'stage': f'文件 {file.filename} 读取完成！',
                'sessionId': session_id
            })
            socketio.sleep(0)

            # 写入新的Excel文件
            output_path = write_to_excel(equipment_dict, file.filename)  # 使用原始文件名

            # 发送进度更新
            socketio.emit('progress', {
                'progress': 100,
                'stage': f'文件 {file.filename} 上传成功！\n保存为: {new_filename}\n成功读取设备材料表\n已生成输出文件: {os.path.basename(output_path)}',
                'sessionId': session_id
            })
            socketio.sleep(0)

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
            # 发送错误信息
            socketio.emit('progress', {
                'progress': 0,
                'stage': f'上传失败: {str(e)}',
                'sessionId': session_id
            })
            socketio.sleep(0)
            return jsonify({'error': str(e)}), 500

    @socketio.on('upload')
    def handle_upload(data):
        """处理上传开始事件"""
        filename = data.get('filename')
        session_id = data.get('sessionId')
        print(f'Upload started for file: {filename} (Session: {session_id})')
