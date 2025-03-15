document.addEventListener('DOMContentLoaded', function() {
    // 获取DOM元素
    const uploadForm = document.getElementById('upload-form');
    const fileInput = document.getElementById('file-input');
    const uploadBtn = document.getElementById('upload-btn');
    const processingCard = document.getElementById('processing-card');
    const resultCard = document.getElementById('result-card');
    const resultMessages = document.getElementById('result-messages');
    const resultFiles = document.getElementById('result-files');
    const existingFilesCard = document.getElementById('existing-files-card');
    const existingFiles = document.getElementById('existing-files');
    const noFilesMessage = document.getElementById('no-files-message');
    const refreshFilesBtn = document.getElementById('refresh-files-btn');
    
    // 初始化折叠按钮
    initCollapseButtons();

    // 页面加载时获取已有文件
    loadExistingFiles();

    // 刷新文件列表按钮事件
    refreshFilesBtn.addEventListener('click', loadExistingFiles);

    // 文件上传表单提交事件
    uploadForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        // 检查是否选择了文件
        if (!fileInput.files.length) {
            alert('请选择要上传的文件');
            return;
        }
        
        // 检查是否至少选择了一种生成类型
        const generateExcel = document.getElementById('generate-excel').checked;
        const generatePpt = document.getElementById('generate-ppt').checked;
        
        if (!generateExcel && !generatePpt) {
            alert('请至少选择一种生成类型（Excel或PPT）');
            return;
        }
        
        // 检查文件类型
        const files = fileInput.files;
        let allValid = true;
        
        for (let i = 0; i < files.length; i++) {
            if (!files[i].name.toLowerCase().endsWith('.json')) {
                alert(`文件 "${files[i].name}" 不是JSON格式，请上传JSON文件`);
                allValid = false;
                break;
            }
        }
        
        if (!allValid) {
            return;
        }
        
        // 显示处理中状态
        uploadBtn.disabled = true;
        processingCard.classList.remove('d-none');
        resultCard.classList.add('d-none');
        
        // 创建FormData对象
        const formData = new FormData();
        
        // 添加所有文件
        for (let i = 0; i < files.length; i++) {
            formData.append('files', files[i]);
        }
        
        // 添加生成类型选项
        formData.append('generate_excel', generateExcel);
        formData.append('generate_ppt', generatePpt);
        
        // 发送上传请求
        fetch('/upload', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                return response.json().then(data => {
                    throw new Error(data.error || '上传失败');
                });
            }
            return response.json();
        })
        .then(data => {
            // 处理成功
            displayResults(data);
            // 刷新文件列表
            loadExistingFiles();
        })
        .catch(error => {
            // 处理错误
            alert('处理文件时出错: ' + error.message);
            console.error('Error:', error);
        })
        .finally(() => {
            // 恢复上传按钮状态
            uploadBtn.disabled = false;
            processingCard.classList.add('d-none');
        });
    });

    // 显示处理结果
    function displayResults(data) {
        // 清空之前的结果
        resultMessages.innerHTML = '';
        
        // 添加处理消息
        if (data.messages && data.messages.length) {
            data.messages.forEach(message => {
                const messageEl = document.createElement('div');
                messageEl.className = 'result-message';
                messageEl.textContent = message;
                resultMessages.appendChild(messageEl);
            });
        }
        
        // 显示结果卡片
        resultCard.classList.remove('d-none');
    }

    // 加载已有文件
    function loadExistingFiles() {
        fetch('/output_files')
            .then(response => response.json())
            .then(data => {
                existingFiles.innerHTML = '';
                
                if (data.files && data.files.length) {
                    noFilesMessage.style.display = 'none';
                    
                    data.files.forEach(file => {
                        existingFiles.appendChild(createFileItem(file));
                    });
                } else {
                    noFilesMessage.style.display = 'block';
                }
            })
            .catch(error => {
                console.error('Error loading existing files:', error);
                existingFiles.innerHTML = '<p class="text-center text-danger">加载文件列表失败</p>';
            });
    }

    // 创建文件列表项
    function createFileItem(file) {
        const item = document.createElement('div');
        item.className = 'list-group-item file-item';
        
        // 根据文件类型设置图标
        let iconClass = 'bi';
        if (file.type === 'excel') {
            iconClass += ' bi-file-earmark-excel excel-icon';
        } else if (file.type === 'ppt') {
            iconClass += ' bi-file-earmark-slides ppt-icon';
        } else {
            iconClass += ' bi-file-earmark';
        }
        
        // 创建文件链接、下载按钮和删除按钮
        item.innerHTML = `
            <div class="d-flex justify-content-between align-items-center">
                <a href="#" class="file-link">
                    <i class="${iconClass} file-icon"></i>
                    <span>${file.name}</span>
                </a>
                <div class="file-actions">
                    <button class="btn btn-sm btn-outline-success download-file-btn me-2" data-filepath="${file.path}" data-filename="${file.name}">
                        <i class="bi bi-download"></i>
                    </button>
                    <button class="btn btn-sm btn-outline-danger delete-file-btn" data-filename="${file.name}">
                        <i class="bi bi-trash"></i>
                    </button>
                </div>
            </div>
        `;
        
        // 添加下载按钮事件
        const downloadBtn = item.querySelector('.download-file-btn');
        downloadBtn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            const filepath = this.getAttribute('data-filepath');
            const filename = this.getAttribute('data-filename');
            
            // 创建一个临时链接并触发下载
            const link = document.createElement('a');
            link.href = filepath;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });
        
        // 添加删除按钮事件
        const deleteBtn = item.querySelector('.delete-file-btn');
        deleteBtn.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            const filename = this.getAttribute('data-filename');
            deleteFile(filename);
        });
        
        return item;
    }
    
    // 初始化折叠按钮功能
    function initCollapseButtons() {
        const collapseButtons = document.querySelectorAll('.collapse-btn');
        
        collapseButtons.forEach(button => {
            const targetId = button.getAttribute('data-target');
            const targetElement = document.getElementById(targetId);
            
            button.addEventListener('click', function() {
                // 切换内容区域的显示状态
                if (targetElement.style.display === 'none') {
                    targetElement.style.display = 'block';
                    button.querySelector('i').className = 'bi bi-chevron-up';
                } else {
                    targetElement.style.display = 'none';
                    button.querySelector('i').className = 'bi bi-chevron-down';
                }
            });
        });
    }
    
    // 删除文件
    function deleteFile(filename) {
        fetch(`/delete_file/${encodeURIComponent(filename)}`, {
            method: 'DELETE'
        })
        .then(response => {
            if (!response.ok) {
                return response.json().then(data => {
                    throw new Error(data.error || '删除失败');
                });
            }
            return response.json();
        })
        .then(data => {
            // 删除成功，刷新文件列表
            loadExistingFiles();
        })
        .catch(error => {
            alert('删除文件时出错: ' + error.message);
            console.error('Error:', error);
        });
    }
});
