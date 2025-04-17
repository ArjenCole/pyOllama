// 获取dropzone和fileInput DOM元素
var dropzone = document.getElementById('dropzone');
var fileInput = document.getElementById('fileInput');
var progressLable = document.getElementById('progressLabel');
var estimationButton = document.getElementById('estimationButton');
var downloadButton = document.getElementById('downloadButton');
const socket = io();
const sessionId = Math.random().toString(36).substring(7); // 生成唯一的会话ID
var uploadedFilePath = null;

// 连接时发送会话ID
socket.emit('init', { sessionId: sessionId });

// 拖拽时显示为复制状态
dropzone.addEventListener('dragover', function(event) {
    event.preventDefault();
    event.stopPropagation();
    event.dataTransfer.dropEffect = 'copy';
    dropzone.classList.add('dragover');
});

// 拖拽离开时移除样式
dropzone.addEventListener('dragleave', function(event) {
    event.preventDefault();
    event.stopPropagation();
    dropzone.classList.remove('dragover');
});

// 放开时触发文件选择
dropzone.addEventListener('drop', function(event) {
    event.preventDefault();
    event.stopPropagation();
    dropzone.classList.remove('dragover');
    
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
    socket.emit('upload', {filename: files[0].name, sessionId: sessionId});
    fetch('/estimation/water/upload', {
        method: 'POST',
        headers: {
            'X-Session-ID': sessionId
        },
        body: formData
    }).then(response => response.json())
        .then(data => {
            socket.emit('progress', {progress: data.progress, sessionId: sessionId});
            if (data.filename) {
                uploadedFilePath = data.filename;
                // 显示估算按钮
                estimationButton.style.display = 'block';
                // 设置估算按钮的点击事件
                estimationButton.onclick = function() {
                    startEstimation(uploadedFilePath);
                };
            }
        });
}

// 开始估算的函数
function startEstimation(filePath) {
    fetch('/estimation/water/process', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'X-Session-ID': sessionId
        },
        body: JSON.stringify({filename: filePath})
    }).then(response => response.json())
        .then(data => {
            socket.emit('progress', {progress: data.progress, sessionId: sessionId});
            if (data.output_file) {
                // 隐藏估算按钮
                estimationButton.style.display = 'none';
                // 显示下载按钮
                downloadButton.style.display = 'block';
                // 设置下载按钮的点击事件
                downloadButton.onclick = function() {
                    window.location.href = `/estimation/water/download/${data.output_file}`;
                };
            }
        });
}

socket.on('progress', function (data) {
    // 只处理属于当前会话的进度更新
    if (data.sessionId === sessionId) {
        const progressBar = document.getElementById('progressBar');
        progressBar.style.width = data.progress + '%';
        var inputString = data.stage;
        // 将换行符替换为HTML的换行标签
        progressLable.innerHTML = inputString.replace(/\n/g, '<br>')
    }
}); 