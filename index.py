from flask import Flask, render_template

app = Flask(__name__)


@app.route('/')
def index():
    # 首页路由，渲染首页的 HTML
    return render_template('index.html')


@app.route('/dropzone')
def dropzone():
    # 文件上传页面路由
    return render_template('dropzone.html')


@app.route('/chat')
def chat():
    # 聊天页面路由
    return render_template('chat.html')


@app.route('/knowledge')
def knowledge():
    # 知识库页面路由
    return render_template('knowledge.html')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)
