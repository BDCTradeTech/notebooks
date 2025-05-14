document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    const uploadForm = document.getElementById('uploadForm');
    const fileInputLabel = document.querySelector('.file-input-label');

    fileInput.addEventListener('change', function(e) {
        const fileName = e.target.files[0]?.name || 'Seleccionar archivo Excel';
        fileInputLabel.textContent = fileName;
    });

    uploadForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        const formData = new FormData(uploadForm);
        
        fetch('/', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (response.ok) {
                return response.blob();
            }
            throw new Error('Error en el procesamiento');
        })
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Notebooks BDC ' + new Date().toLocaleDateString('es-ES') + '.xlsx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();
        })
        .catch(error => {
            console.error('Error:', error);
            alert('Error al procesar el archivo. Por favor, intente nuevamente.');
        });
    });
}); 