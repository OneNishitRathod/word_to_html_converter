// Handle file selection display
document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('document');
    const fileNameSpan = document.getElementById('file-name');
    
    if (fileInput && fileNameSpan) {
        fileInput.addEventListener('change', function() {
            if (fileInput.files.length > 0) {
                fileNameSpan.textContent = fileInput.files[0].name;
            } else {
                fileNameSpan.textContent = 'No file chosen';
            }
        });
    }
});