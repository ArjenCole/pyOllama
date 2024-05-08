from flask import Flask, render_template_string, request

app = Flask(__name__)

# HTML模板
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>File Upload</title>
    <style>
        .dropzone {
            border: 2px dashed #ccc; /* 虚线边框 */
            padding: 20px;
            text-align: center;
            color: #ccc;
            margin: 50px auto;
            width: 70%;
        }
        .dropzone.dragover {
            border-color: #000;
            color: #000;
        }
        .hidden {
            display: none;
        }
    </style>
</head>
<body>

<div class="dropzone" id="dropzone" ondrop="drop(event)" ondragover="allowDrop(event)">
    <p>将文件拖入这里...</p>
    <input type="file" id="fileInput" class="hidden" multiple>
</div>

<script>
    function allowDrop(event) {
        event.preventDefault();
    }

    function drop(event) {
        event.preventDefault();
        document.getElementById('fileInput').click();
    }

    document.getElementById('fileInput').addEventListener('change', handleFileSelect, false);

    function handleFileSelect(event) {
        var files = event.target.files;
        for (var i = 0, f; f = files[i]; i++) {
            var reader = new FileReader();
            reader.onload = function(e) {
                console.log('File contents:', e.target.result);
                // 这里可以将文件内容发送到服务器，或者进行其他处理
            };
            reader.readAsText(f);
        }
    }
    document.getElementById('dropzone').onclick = function() {
        document.getElementById('fileInput').click();
    }
</script>

</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

if __name__ == '__main__':
    app.run(debug=True)