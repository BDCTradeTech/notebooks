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

    let generatedWorkbook = null;

    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const file = fileInput.files[0];
        if (!file) {
            alert('Por favor, selecciona un archivo Excel');
            return;
        }

        if (!file.name.match(/\.(xlsx|xls)$/)) {
            alert('Por favor, selecciona un archivo Excel válido (.xlsx o .xls)');
            return;
        }

        processButton.disabled = true;
        processButton.textContent = 'Procesando...';
        downloadSection.style.display = 'none';
        generatedWorkbook = null;

        try {
            // Leer el archivo Excel subido
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const originalSheet = workbook.Sheets[firstSheetName];
            const originalData = XLSX.utils.sheet_to_json(originalSheet, { header: 1 });

            // Crear nuevo libro y hoja
            const newHeader = [
                'SKU', 'Marca', 'Descripción', 'Familia', 'Pantalla', 'Memoria', 'Disco', 'Qty', 'Price', 'ETA', 'MOQ'
            ];
            const newData = [newHeader];

            // Copiar los datos del archivo original (ignorando la cabecera original)
            for (let i = 1; i < originalData.length; i++) {
                newData.push(originalData[i]);
            }

            const newSheet = XLSX.utils.aoa_to_sheet(newData);
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Notebooks');

            generatedWorkbook = newWorkbook;
            downloadSection.style.display = 'block';
        } catch (error) {
            alert('Error al procesar el archivo: ' + error.message);
        } finally {
            processButton.disabled = false;
            processButton.textContent = 'Procesar';
        }
    });

    downloadButton.addEventListener('click', () => {
        if (!generatedWorkbook) {
            alert('No hay archivo generado para descargar.');
            return;
        }
        const wbout = XLSX.write(generatedWorkbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Notebooks.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });
}); 