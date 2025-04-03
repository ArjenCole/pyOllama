import time
import sys
from flask import Flask, render_template, request
from flask_socketio import SocketIO
import os

from dropzone import dropzone_upload  # 假设dropzone.py在同一包内，使用相对导入
from ai import chat_ai
from estimation_water import init_routes as init_water_routes

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'
socketio = SocketIO(app)

# 存储客户端会话ID的字典
client_sessions = {}

@app.route('/')
def index():
    # 首页路由，渲染首页的 HTML
    return render_template('index.html')

@app.route('/dropzone')
def dropzone():
    # 文件上传页面路由
    return render_template('dropzone.html')

@app.route('/upload', methods=['POST'])
def upload():
    _stage_update(5, '开始上传文件')
    _dir_dict = dropzone_upload(socketio, client_sessions)
    return {'??': 100}

@app.route('/chat')
def chat():
    # 聊天页面路由
    return render_template('chat.html')

@app.route('/ai', methods=['POST'])
def ai():
    return chat_ai()

@app.route('/knowledge')
def knowledge():
    # 知识库页面路由
    return render_template('knowledge.html')

@app.route('/estimation')
def estimation():
    # 智能估算页面路由
    return render_template('estimation.html')

def _stage_update(p_percent, p_stage, session_id=None):
    if session_id:
        # 只向特定客户端发送进度更新
        socketio.emit('progress', {'progress': p_percent, 'stage': p_stage, 'sessionId': session_id})
    else:
        # 如果没有指定session_id，则广播给所有客户端
        socketio.emit('progress', {'progress': p_percent, 'stage': p_stage})
    if p_percent < 100:
        time.sleep(1)

@socketio.on('connect')
def handle_connect():
    print('Client connected')

@socketio.on('disconnect')
def handle_disconnect():
    # 客户端断开连接时清理会话ID
    if request.sid in client_sessions:
        del client_sessions[request.sid]
    print('Client disconnected')

@socketio.on('init')
def handle_init(data):
    # 处理客户端初始化，保存会话ID
    session_id = data.get('sessionId')
    if session_id:
        client_sessions[request.sid] = session_id

# 初始化水厂估算相关路由
init_water_routes(app, socketio)

if __name__ == '__main__':
    try:
        socketio.run(app, debug=True, host='0.0.0.0', port=8080)
    except KeyboardInterrupt:
        print('正在关闭服务器...')
        socketio.stop()
        sys.exit(0)
