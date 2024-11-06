import time

from flask import Flask, render_template, request
from flask_socketio import SocketIO
import os

from dropzone import dropzone_upload  # 假设dropzone.py在同一包内，使用相对导入
from ai import chat_ai

app = Flask(__name__)
socketio = SocketIO(app)


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
    _dir_dict = dropzone_upload(socketio)

    return {'??': 100}
    # 使用从dropzone.py导入的upload_file逻辑处理上传
    # return dropzone_upload()


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


def _stage_update(p_percent, p_stage):
    socketio.emit('progress', {'progress': p_percent, 'stage': p_stage})
    if p_percent < 100:
        time.sleep(1)


if __name__ == '__main__':
    # app.run(debug=True, host='0.0.0.0', port=8080)
    socketio.run(app, debug=True, host='0.0.0.0', port=8080)
