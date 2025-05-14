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
                'SKU', 'Marca', 'Descripción', 'Familia', 'Pantalla', 'Memoria', 'Disco', 'Qty', 'Precio', 'ETA', 'MOQ'
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
                const descripcion = row[descripcionIdx] || '';

                // Generar Familia según la descripción
                let familia = '';
                const familias = [
                    'Intel Celeron', 'Intel Pentium', 'Intel Core 5', 'Intel Core 7', 'Intel Core i3', 'Intel Core i5', 'Intel Core i7', 'Intel Core i9', 'Intel Core Ultra5', 'Intel Core Ultra7', 'Intel Core Ultra9', 'AMD Ryzen 3', 'AMD Ryzen 5', 'AMD Ryzen 7', 'AMD Ryzen 9', 'Apple'
                ];
                for (const f of familias) {
                    if (descripcion.toLowerCase().includes(f.toLowerCase())) {
                        familia = f;
                        break;
                    }
                }

                // Generar Pantalla según la descripción
                let pantalla = '';
                const pantallas = [
                    { pattern: '11.6', value: '11.6"' },
                    { pattern: '14.1', value: '14.1"' },
                    { pattern: '13"', value: '13"' },
                    { pattern: '15.6', value: '15.6"' },
                    { pattern: '16 inch', value: '16"' },
                    { pattern: '13.3 inch', value: '13.3"' },
                    { pattern: '13IN', value: '13"' }
                ];
                for (const p of pantallas) {
                    if (descripcion.includes(p.pattern)) {
                        pantalla = p.value;
                        break;
                    }
                }

                // Generar Memoria según la descripción
                let memoria = '';
                const memorias = [
                    { pattern: '8G', value: '8GB' },
                    { pattern: '4GB', value: '4GB' },
                    { pattern: '12GB', value: '12GB' },
                    { pattern: '16GB', value: '16GB' },
                    { pattern: '32GB', value: '32GB' },
                    { pattern: '24GB', value: '24GB' }
                ];
                for (const m of memorias) {
                    if (descripcion.includes(m.pattern)) {
                        memoria = m.value;
                        break;
                    }
                }

                // Generar Disco según la descripción
                let disco = '';
                const discos = [
                    { pattern: '512G', value: '512GB' },
                    { pattern: '128G', value: '128GB' },
                    { pattern: '64GB', value: '64GB' },
                    { pattern: '256GB', value: '256GB' },
                    { pattern: '127GB', value: '127GB' },
                    { pattern: '1TB', value: '1TB' }
                ];
                for (const d of discos) {
                    if (descripcion.includes(d.pattern)) {
                        disco = d.value;
                        break;
                    }
                }

                // Aplicar margen BDC al precio
                const margenBDC = parseFloat(document.getElementById('margenBDC').value) || 5.0;
                const precioOriginal = parseFloat(price) || 0;
                const precioConMargen = precioOriginal * (1 + margenBDC / 100);

                const newRow = [
                    row[headerMap['sku']] || '',
                    row[headerMap['marca']] || '',
                    descripcion,
                    familia,
                    pantalla,
                    memoria,
                    disco,
                    row[headerMap['qty']] || '',
                    precioConMargen.toFixed(2),
                    row[headerMap['eta']] || '',
                    row[headerMap['moq']] || ''
                ];
                newData.push(newRow);
            }

            const newSheet = XLSX.utils.aoa_to_sheet(newData);
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Notebooks');

            // Aplicar formato al Excel generado
            const range = XLSX.utils.decode_range(newSheet['!ref']);
            for (let R = range.s.r; R <= range.e.r; R++) {
                for (let C = range.s.c; C <= range.e.c; C++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!newSheet[cellAddress]) newSheet[cellAddress] = {};
                    newSheet[cellAddress].s = {
                        font: { name: 'Verdana', sz: 11 },
                        fill: { fgColor: { rgb: 'FFFFFF' } },
                        border: {
                            top: { style: 'thin' },
                            bottom: { style: 'thin' },
                            left: { style: 'thin' },
                            right: { style: 'thin' }
                        }
                    };
                }
            }
            // Aplicar formato a la cabecera
            for (let C = range.s.c; C <= range.e.c; C++) {
                const cellAddress = XLSX.utils.encode_cell({ r: 0, c: C });
                if (!newSheet[cellAddress]) newSheet[cellAddress] = {};
                newSheet[cellAddress].s = {
                    font: { name: 'Verdana', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
                    fill: { fgColor: { rgb: '000000' } },
                    border: {
                        top: { style: 'thin' },
                        bottom: { style: 'thin' },
                        left: { style: 'thin' },
                        right: { style: 'thin' }
                    }
                };
            }
            // Aplicar borde exterior grueso
            for (let C = range.s.c; C <= range.e.c; C++) {
                const cellAddress = XLSX.utils.encode_cell({ r: 0, c: C });
                if (!newSheet[cellAddress]) newSheet[cellAddress] = {};
                newSheet[cellAddress].s.border.top = { style: 'thick' };
                newSheet[cellAddress].s.border.bottom = { style: 'thick' };
                newSheet[cellAddress].s.border.left = { style: 'thick' };
                newSheet[cellAddress].s.border.right = { style: 'thick' };
            }

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
        const today = new Date();
        const dd = String(today.getDate()).padStart(2, '0');
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const yy = String(today.getFullYear()).slice(-2);
        const fileName = `Notebooks BDC ${dd}-${mm}-${yy}.xlsx`;
        const wbout = XLSX.write(generatedWorkbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });
}); 