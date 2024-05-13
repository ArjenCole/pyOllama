from flask import Flask, render_template
from dropzone import upload_file  # 假设dropzone.py在同一包内，使用相对导入
from chat import commAI

app = Flask(__name__)


@app.route('/')
def index():
    # 首页路由，渲染首页的 HTML
    return render_template('index.html')


@app.route('/dropzone')
def dropzone():
    # 文件上传页面路由
    return render_template('dropzone.html')


@app.route('/upload', methods=['POST'])
def handle_upload():
    # 使用从dropzone.py导入的upload_file逻辑处理上传
    return upload_file()


@app.route('/chat')
def chat():
    # 聊天页面路由
    return render_template('chat.html')


@app.route('/ai', methods=['POST'])
def handle_commAI():
    return commAI()


@app.route('/knowledge')
def knowledge():
    # 知识库页面路由
    return render_template('knowledge.html')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
