from flask import Flask, render_template_string, request, jsonify
from flask_cors import CORS
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)  # 允许跨源请求

# HTML模板
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>File Upload</title>
    <style>
        .dropzone {
            border: 2px dashed #ccc;
            padding: 20px;
            text-align: center;
            color: #ccc;
            margin: 100px auto;
            width: 70%;
            cursor: pointer;
        }
        .hidden {
            display: none;
        }
    </style>
</head>
<body>

<div class="dropzone" id="dropzone">
    <p>将文件拖入这里，或点击选择文件上传</p>
    <input type="file" id="fileInput" class="hidden" multiple>
</div>

<script>
    // 获取dropzone和fileInput DOM元素
    var dropzone = document.getElementById('dropzone');
    var fileInput = document.getElementById('fileInput');

    // 拖拽时显示为复制状态
    dropzone.addEventListener('dragover', function(event) {
        event.stopPropagation();
        event.preventDefault();
        event.dataTransfer.dropEffect = 'copy';
    });

    // 放开时触发文件选择
    dropzone.addEventListener('drop', function(event) {
        event.stopPropagation();
        event.preventDefault();
        if (event.dataTransfer.files.length) {
            fileInput.files = event.dataTransfer.files;
            uploadFiles(fileInput.files);
        }
    });

    // 点击dropzone时触发文件选择对话框
    dropzone.addEventListener('click', function() {
        fileInput.click();
    });

    // 监听文件输入控件的变化
    fileInput.addEventListener('change', function(event) {
        uploadFiles(event.target.files);
    });

    // 上传文件的函数
    function uploadFiles(files) {
        var formData = new FormData();
        for (var i = 0; i < files.length; i++) {
            formData.append('file' + i, files[i]);
        }

        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/upload', true); // 打开一个新连接，使用POST请求访问服务器上的/upload路径
        xhr.onload = function () {
            if (xhr.status === 200) {
                alert('文件上传成功');
            } else {
                alert('文件上传失败');
            }
            console.log('Server response:', xhr.responseText);
        };
        xhr.onerror = function () {
            alert('文件上传发生错误');
        };
        xhr.send(formData); // 发送请求
    }
</script>

</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_file():
    print('进来了！', request.files)
    if 'file0' not in request.files:
        print('出去了！')
        return jsonify({'error': '没有文件部分'})
    # file = request.files.get('file')
    file = request.files['file0']
    print('filename', file.filename)
    if file.filename == '':
        return jsonify({'error': '没有选择文件'})
    if file:
        if not os.path.exists('G:/code/uploads'):
            os.makedirs('G:/code/uploads')  # 确保目录存在
        # filename = secure_filename(file.filename)  # 使用 Werkzeug 库提供的 secure_filename 函数
        filename = file.filename
        file_path = os.path.join('G:/code/uploads', filename)

        print(file_path)
        file.save(file_path)
        return jsonify({'message': '文件上传成功'})
    return jsonify({'error': '上传失败，未知错误'})

if __name__ == '__main__':
    app.run(debug=True)