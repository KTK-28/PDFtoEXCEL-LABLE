// JavaScript for file drop feature
const dropContainer = document.getElementById('dropContainer');
const fileInput = document.getElementById('fileInput');
const dropMessage = document.getElementById('dropMessage');

dropContainer.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropContainer.classList.add('drag-over');
});

dropContainer.addEventListener('dragleave', () => {
    dropContainer.classList.remove('drag-over');
});

dropContainer.addEventListener('drop', (e) => {
    e.preventDefault();
    dropContainer.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    fileInput.files = e.dataTransfer.files;
    dropMessage.textContent = file.name;
});

function uploadFile() {
    document.getElementById('uploadForm').submit();
}