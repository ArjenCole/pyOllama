
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