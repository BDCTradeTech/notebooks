document.addEventListener('DOMContentLoaded', () => {
    const uploadForm = document.getElementById('uploadForm');
    const fileInput = document.getElementById('fileInput');
    const processButton = document.getElementById('processButton');
    const downloadSection = document.getElementById('downloadSection');
    const downloadButton = document.getElementById('downloadButton');

    // Actualizar el texto del label cuando se selecciona un archivo
    fileInput.addEventListener('change', (e) => {
        const fileName = e.target.files[0]?.name || 'Seleccionar archivo Excel';
        document.querySelector('.file-input-label').textContent = fileName;
    });

    // Manejar el envío del formulario
    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const file = fileInput.files[0];
        if (!file) {
            alert('Por favor, selecciona un archivo Excel');
            return;
        }

        // Verificar que el archivo sea de Excel
        if (!file.name.match(/\.(xlsx|xls)$/)) {
            alert('Por favor, selecciona un archivo Excel válido (.xlsx o .xls)');
            return;
        }

        // Aquí irá la lógica de procesamiento del archivo
        // Por ahora, solo simularemos el procesamiento
        processButton.disabled = true;
        processButton.textContent = 'Procesando...';

        try {
            // Simular un procesamiento
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Mostrar el botón de descarga
            downloadSection.style.display = 'block';
        } catch (error) {
            alert('Error al procesar el archivo');
        } finally {
            processButton.disabled = false;
            processButton.textContent = 'Procesar';
        }
    });

    // Manejar la descarga
    downloadButton.addEventListener('click', () => {
        // Aquí irá la lógica de descarga
        alert('Funcionalidad de descarga pendiente de implementar');
    });
}); 