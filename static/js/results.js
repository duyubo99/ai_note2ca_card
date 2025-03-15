// JavaScript for results page

document.addEventListener('DOMContentLoaded', function() {
    // Add download tracking
    const downloadButtons = document.querySelectorAll('.btn-success');
    
    downloadButtons.forEach(button => {
        button.addEventListener('click', function() {
            const filename = this.getAttribute('href').split('/').pop();
            console.log(`Downloading file: ${filename}`);
            
            // Optional: Add visual feedback when download starts
            const originalText = this.innerHTML;
            this.innerHTML = '<i class="bi bi-check-circle me-1"></i> 下载中...';
            
            setTimeout(() => {
                this.innerHTML = originalText;
            }, 2000);
        });
    });
    
    // Enable tooltips if Bootstrap is available
    if (typeof bootstrap !== 'undefined' && bootstrap.Tooltip) {
        const tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
        tooltipTriggerList.map(function (tooltipTriggerEl) {
            return new bootstrap.Tooltip(tooltipTriggerEl);
        });
    }
});
