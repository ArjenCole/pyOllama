import os
import time
import pandas as pd
from dataclasses import dataclass
from typing import List, Dict
from flask import jsonify, request, render_template
from werkzeug.utils import secure_filename
from flask_socketio import SocketIO

# 配置上传文件存储路径
UPLOAD_FOLDER = 'uploads/estimation_water'
OUTPUT_FOLDER = 'output/estimation_water'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 确保上传目录和输出目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@dataclass
class EquipmentMaterial:
    """设备材料信息类"""
    name: str  # 名称
    specification: str  # 规格
    material: str  # 材料
    unit: str  # 单位
    quantity: float  # 数量
    remarks: str  # 备注

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
        # 读取Excel文件的所有工作表
        excel_file = pd.ExcelFile(file_path)
        
        # 查找设备材料表
        # 假设表头包含这些字段
        required_columns = ["序号", "所属单体", "名称", "规格", "材料", "单位", "数量", "备注"]
        
        # 遍历所有工作表
        for sheet_name in excel_file.sheet_names:
            df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
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
        template_wb = pd.ExcelFile(template_path)
        template_df = pd.read_excel(template_wb, sheet_name='Sheet1')
        
        # 创建一个新的Excel写入器
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 复制模板的格式
            template_df.to_excel(writer, sheet_name='Sheet1', index=False)
            
            # 获取工作簿和工作表对象
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # 从模板中复制样式
            from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
            from openpyxl.utils import get_column_letter
            
            # 设置边框样式
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # 设置标题行样式
            header_font = Font(bold=True)
            header_fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')
            
            # 应用标题行样式
            for col in range(1, len(template_df.columns) + 1):
                cell = worksheet.cell(row=1, column=col)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # 找到"工程或费用名称"列的索引
            name_col_idx = 3

            # 从第2行开始写入数据（第1行是标题）
            current_row = 7
            
            # 写入每个key，并在key之间添加3个空行
            for key in equipment_dict.keys():
                # 写入key
                cell = worksheet.cell(row=current_row, column=name_col_idx)
                cell.value = key
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='left', vertical='center')
                
                # 添加3个空行
                current_row += 4
            
            # 调整列宽
            for col in range(1, len(template_df.columns) + 1):
                column_letter = get_column_letter(col)
                worksheet.column_dimensions[column_letter].width = 15
            
            # 设置行高
            for row in range(1, current_row):
                worksheet.row_dimensions[row].height = 20
        
        return output_path
        
    except Exception as e:
        raise Exception(f"写入Excel文件时出错: {str(e)}")

def init_routes(app, socketio):
    """初始化路由和Socket.IO事件处理"""
    
    @app.route('/estimation/water')
    def estimation_water():
        """水厂智能估算页面路由"""
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
            
            # 添加时间戳确保文件名唯一
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            new_filename = f"{timestamp}_{filename}"
            file_path = os.path.join(UPLOAD_FOLDER, new_filename)
            
            # 保存文件
            file.save(file_path)
            
            # 处理Excel文件
            equipment_dict = process_excel_file(file_path)
            
            # 写入新的Excel文件
            output_path = write_to_excel(equipment_dict, file.filename)  # 使用原始文件名
            
            # 发送进度更新
            socketio.emit('progress', {
                'progress': 100,
                'stage': f'文件 {file.filename} 上传成功！\n保存为: {new_filename}\n成功读取设备材料表\n已生成输出文件: {os.path.basename(output_path)}',
                'sessionId': session_id
            })
            
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
            return jsonify({'error': str(e)}), 500

    @socketio.on('upload')
    def handle_upload(data):
        """处理上传开始事件"""
        filename = data.get('filename')
        session_id = data.get('sessionId')
        print(f'Upload started for file: {filename} (Session: {session_id})') 