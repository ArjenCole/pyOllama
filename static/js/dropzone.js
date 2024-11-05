
    // 获取dropzone和fileInput DOM元素
    var dropzone = document.getElementById('dropzone');
    var fileInput = document.getElementById('fileInput');
    var progressLable = document.getElementById('progressLabel')
    const socket = io();

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
        socket.emit('upload', {filename: files[0].name});
        fetch('/upload', {
            method: 'POST',
            body: formData
        }).then(response => response.json())
            .then(data => {
                socket.emit('progress', {progress: data.progress});
            });
    }

    socket.on('progress', function (data) {
        const progressBar = document.getElementById('progressBar');
        progressBar.style.width = data.progress + '%';
        progressLable.textContent = data.stage
    });