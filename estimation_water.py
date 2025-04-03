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
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 确保上传目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

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
            # 生成安全的文件名
            filename = secure_filename(file.filename)
            # 添加时间戳确保文件名唯一
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            new_filename = f"{timestamp}_{filename}"
            file_path = os.path.join(UPLOAD_FOLDER, new_filename)
            
            # 保存文件
            file.save(file_path)
            
            # 处理Excel文件
            equipment_dict = process_excel_file(file_path)
            
            # 发送进度更新
            socketio.emit('progress', {
                'progress': 100,
                'stage': f'文件 {filename} 上传成功！\n保存为: {new_filename}\n成功读取设备材料表',
                'sessionId': session_id
            })
            
            return jsonify({
                'message': '文件上传成功',
                'filename': new_filename,
                'progress': 100,
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