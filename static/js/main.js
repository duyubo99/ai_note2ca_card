// Main JavaScript for AI笔录解析系统

document.addEventListener('DOMContentLoaded', function() {
    // Get DOM elements
    const fileUpload = document.getElementById('file-upload');
    const fileList = document.getElementById('file-list');
    const selectedFiles = document.getElementById('selected-files');
    const uploadForm = document.getElementById('upload-form');
    const submitBtn = document.getElementById('submit-btn');
    const processingIndicator = document.getElementById('processing-indicator');
    const excelCheckbox = document.getElementById('excel-checkbox');
    const pptCheckbox = document.getElementById('ppt-checkbox');
    
    // Handle file selection
    if (fileUpload) {
        fileUpload.addEventListener('change', function() {
            fileList.innerHTML = '';
            
            if (this.files.length > 0) {
                selectedFiles.style.display = 'block';
                
                for (let i = 0; i < this.files.length; i++) {
                    const file = this.files[i];
                    const listItem = document.createElement('li');
                    listItem.className = 'list-group-item';
                    
                    // Add file icon based on extension
                    const fileIcon = document.createElement('i');
                    fileIcon.className = 'bi bi-file-earmark-word me-2';
                    fileIcon.style.color = '#2b579a';
                    
                    // Add file name and size
                    const fileName = document.createElement('span');
                    fileName.textContent = file.name;
                    
                    const fileSize = document.createElement('span');
                    fileSize.className = 'text-muted ms-2';
                    fileSize.textContent = `(${formatFileSize(file.size)})`;
                    
                    listItem.appendChild(fileIcon);
                    listItem.appendChild(fileName);
                    listItem.appendChild(fileSize);
                    
                    fileList.appendChild(listItem);
                }
            } else {
                selectedFiles.style.display = 'none';
            }
        });
    }
    
    // Handle form submission
    if (uploadForm) {
        uploadForm.addEventListener('submit', function(e) {
            // Validate file selection
            if (fileUpload && fileUpload.files.length === 0) {
                e.preventDefault();
                alert('请选择至少一个文件进行上传');
                return false;
            }
            
            // Validate output format selection
            if (excelCheckbox && pptCheckbox && !excelCheckbox.checked && !pptCheckbox.checked) {
                e.preventDefault();
                alert('请至少选择一种输出格式');
                return false;
            }
            
            // Show processing indicator
            if (submitBtn && processingIndicator) {
                submitBtn.disabled = true;
                processingIndicator.style.display = 'block';
            }
            
            return true;
        });
    }
    
    // Format file size to human-readable format
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
    
    // Enable tooltips if Bootstrap is available
    if (typeof bootstrap !== 'undefined' && bootstrap.Tooltip) {
        const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
        tooltipTriggerList.map(function (tooltipTriggerEl) {
            return new bootstrap.Tooltip(tooltipTriggerEl);
        });
    }
});
