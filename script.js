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

            // Buscar la fila de cabecera correcta
            let headerRowIndex = -1;
            let headerMap = {};
            for (let i = 0; i < originalData.length; i++) {
                const row = originalData[i].map(cell => (cell || '').toString().trim().toLowerCase());
                if (
                    row.includes('sku') &&
                    row.includes('marca') &&
                    (row.includes('descripcion') || row.includes('descripción')) &&
                    row.includes('qty') &&
                    row.includes('price') &&
                    row.includes('eta') &&
                    row.includes('moq')
                ) {
                    headerRowIndex = i;
                    // Mapear nombre de columna a índice
                    row.forEach((cell, idx) => {
                        headerMap[cell] = idx;
                    });
                    break;
                }
            }
            if (headerRowIndex === -1) {
                throw new Error('No se encontró la fila de cabecera esperada en el archivo.');
            }

            // Crear nuevo libro y hoja
            const newHeader = [
                'SKU', 'Marca', 'Descripción', 'Familia', 'Pantalla', 'Memoria', 'Disco', 'Qty', 'Price', 'ETA', 'MOQ'
            ];
            const newData = [newHeader];

            // Copiar solo filas válidas (ignorando títulos de grupo y filas vacías)
            for (let i = headerRowIndex + 1; i < originalData.length; i++) {
                const row = originalData[i];
                // Considerar fila válida si tiene SKU, Marca, Qty, Price, ETA y MOQ
                const sku = row[headerMap['sku']] || '';
                const marca = row[headerMap['marca']] || '';
                const qty = row[headerMap['qty']] || '';
                const price = row[headerMap['price']] || '';
                const eta = row[headerMap['eta']] || '';
                const moq = row[headerMap['moq']] || '';
                if (
                    sku.toString().trim() === '' ||
                    marca.toString().trim() === '' ||
                    qty.toString().trim() === '' ||
                    price.toString().trim() === '' ||
                    eta.toString().trim() === '' ||
                    moq.toString().trim() === ''
                ) {
                    continue; // Saltar filas no válidas
                }
                // Mapear columnas según la cabecera encontrada
                const descripcionIdx = headerMap['descripcion'] !== undefined ? headerMap['descripcion'] : headerMap['descripción'];
                const familiaIdx = headerMap['familia'];
                const pantallaIdx = headerMap['pantalla'];
                const memoriaIdx = headerMap['memoria'];
                const discoIdx = headerMap['disco'];
                const newRow = [
                    row[headerMap['sku']] || '',
                    row[headerMap['marca']] || '',
                    row[descripcionIdx] || '',
                    familiaIdx !== undefined ? row[familiaIdx] || '' : '',
                    pantallaIdx !== undefined ? row[pantallaIdx] || '' : '',
                    memoriaIdx !== undefined ? row[memoriaIdx] || '' : '',
                    discoIdx !== undefined ? row[discoIdx] || '' : '',
                    row[headerMap['qty']] || '',
                    row[headerMap['price']] || '',
                    row[headerMap['eta']] || '',
                    row[headerMap['moq']] || ''
                ];
                newData.push(newRow);
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
        a.download = 'Notebooks BDC.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });
}); 