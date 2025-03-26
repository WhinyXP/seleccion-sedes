document.addEventListener('DOMContentLoaded', function() {
    // Elementos del DOM
    const excelFileInput = document.getElementById('excel-file');
    const sedeFileInput = document.getElementById('sede-file');
    const exportBtn = document.getElementById('export-btn');
    const clearBtn = document.getElementById('clear-btn');
    const exportExcelBtn = document.getElementById('export-excel-btn');
    const searchInput = document.getElementById('search-input');
    const tableBody = document.querySelector('#students-table tbody');
    const totalStudentsSpan = document.getElementById('total-students');
    const sedeModal = document.getElementById('sede-modal');
    const closeBtn = document.querySelector('.close-btn');
    const sedeList = document.getElementById('sede-list');
    const modalSearchInput = document.getElementById('modal-search-input');
    
    // Variables de estado
    let studentsData = [];
    let currentStudentIndex = null;
    let sedesData = [];

    // 1. Cargar archivo de alumnos
    excelFileInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        
        // En el evento change del excelFileInput, reemplazamos esta parte:
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            let headerRowIndex = 7;
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row && row.length >= 9 && 
                    (row[0] === 'No.' || row[0] === 'No' || row[0] === 'Número')) {
                    headerRowIndex = i;
                    break;
                }
            }
            
            // Obtener los encabezados reales del archivo
            const headers = jsonData[headerRowIndex].map(h => String(h).trim());
            
            // Determinar qué columnas existen y tienen datos
            const existingColumns = {};
            for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row) {
                    headers.forEach((header, index) => {
                        if (row[index] !== undefined && row[index] !== '') {
                            existingColumns[header] = true;
                        }
                    });
                }
            }
            // Después de determinar los headers y existingColumns, añade esto:
            window.tableColumns = headers
            .filter(header => existingColumns[header] && 
                !['no.', 'no', 'número'].includes(header.toLowerCase()))
            .concat(['Sede', 'Acción']);
            // Dentro del reader.onload del excelFileInput, después de determinar existingColumns:
            window.originalHeaders = headers.filter(header => existingColumns[header]);
            
            studentsData = [];
            for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row && row.length > 0) {
                    const student = {
                        no: row[headers.indexOf('No.')] || row[headers.indexOf('No')] || '',
                        sede: '',
                        sedeId: '',
                        plazasSede: 0
                    };
                    
                    // Añadir dinámicamente las propiedades basadas en los encabezados existentes
                    headers.forEach((header, index) => {
                        if (existingColumns[header] && header !== 'No.' && header !== 'No') {
                            student[header] = row[index] || '';
                        }
                    });
                    
                    studentsData.push(student);
                }
            }
            
            // Guardar información de columnas para usar al generar la tabla
            window.tableColumns = Object.keys(existingColumns)
                .filter(col => col !== 'No.' && col !== 'No')
                .concat(['Sede', 'Acción']);
            
            exportBtn.disabled = false;
            clearBtn.disabled = false;
            totalStudentsSpan.textContent = studentsData.length;
        };
        
        reader.readAsArrayBuffer(file);
    });

    // 2. Cargar archivo de sedes
    sedeFileInput.addEventListener('change', async function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        try {
            const data = new Uint8Array(await file.arrayBuffer());
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            const headerRowIndex = findHeaderRowIndex(jsonData);
            if (headerRowIndex === -1) throw new Error('No se encontraron encabezados válidos');
            
            const headerRow = jsonData[headerRowIndex].map(cell => String(cell).toLowerCase().trim());
            const institucionIndex = headerRow.findIndex(h => h.includes('institución') || h.includes('institucion'));
            const nombreSedeIndex = headerRow.findIndex(h => h.includes('nombre de sede') || h.includes('sede') || h.includes('nombre sede'));
            const plazasIndex = headerRow.findIndex(h => h.includes('número de plazas') || h.includes('numero de plazas') || h.includes('plazas'));

            if (institucionIndex === -1 || nombreSedeIndex === -1) {
                throw new Error('Columnas requeridas no encontradas');
            }

            sedesData = [];
            for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row && row[institucionIndex] && row[nombreSedeIndex]) {
                    const institucion = String(row[institucionIndex]).trim();
                    const sede = String(row[nombreSedeIndex]).trim();
                    const plazas = plazasIndex !== -1 ? parseInt(row[plazasIndex]) || 0 : 0;
                    
                    if (institucion && sede) {
                        sedesData.push({
                            id: `${institucion}-${sede}`,
                            displayText: `${institucion} - ${sede}`,
                            institucion,
                            sede,
                            plazasDisponibles: plazas
                        });
                    }
                }
            }
            
            alert(`Se cargaron ${sedesData.length} sedes correctamente`);
            
        } catch (error) {
            console.error('Error:', error);
            alert(`Error al cargar sedes: ${error.message}`);
        }
    });

    // 3. Exportar a tabla HTML
    // Reemplazamos el exportBtn.addEventListener con esto:
    exportBtn.addEventListener('click', function() {
    // Muestra la tabla solo cuando se exporta
    const tableContainer = document.getElementById('table-container');
    const studentsTable = document.getElementById('students-table');

    tableContainer.style.display = 'block';
    studentsTable.style.display = 'table';

    // Limpia cualquier contenido previo
    document.querySelector('#students-table thead').innerHTML = '';
    document.querySelector('#students-table tbody').innerHTML = '';

        if (studentsData.length === 0) {
            alert('No hay datos para mostrar');
            return;
        }
        
        tableBody.innerHTML = '';
        
        // Generar los encabezados de la tabla
        const thead = document.querySelector('#students-table thead');
        thead.innerHTML = '<tr></tr>';
        const headerRow = thead.querySelector('tr');
        
        // Siempre mostrar No.
        headerRow.innerHTML += '<th>No.</th>';
        
        // Mostrar las columnas dinámicas
        window.tableColumns.forEach(column => {
            headerRow.innerHTML += `<th>${column}</th>`;
        });
        
        // Generar las filas de datos
        studentsData.forEach((student, index) => {
            const row = document.createElement('tr');
            
            // Siempre mostrar el número
            row.innerHTML = `<td>${student.no || ''}</td>`;
            
            // Mostrar las columnas dinámicas
            // Reemplaza esta parte (si usas innerHTML):
            window.tableColumns.forEach(column => {
                if (column === 'Sede') {
                    row.innerHTML += `<td class="sede-cell">${student.sede || ''}</td>`;
                } 
                else if (column === 'Acción') {
                    row.innerHTML += `
                        <td class="action-buttons">
                            <button class="add-sede-btn" data-index="${index}">Añadir sede</button>
                            <button class="clear-sede-btn" data-index="${index}" 
                                ${!student.sede ? 'disabled' : ''}>Limpiar sede</button>
                        </td>
                    `;
                } 
                else {
                    // Marcar la celda de número de cuenta si coincide con patrones comunes
                    const isAccountColumn = [
                        'no. cuenta', 'número de cuenta', 'numero de cuenta', 'cuenta', 'número', 'numero'
                    ].some(term => column.toLowerCase().includes(term));
                    
                    if (isAccountColumn) {
                        row.innerHTML += `<td data-account>${student[column] || ''}</td>`;
                    } else {
                        row.innerHTML += `<td>${student[column] || ''}</td>`;
                    }
                }
            });
            
            tableBody.appendChild(row);
        });
        
        // El resto del código para los event listeners de los botones permanece igual
        document.querySelectorAll('.add-sede-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                if (sedesData.length === 0) {
                    alert('Primero carga el archivo de sedes');
                    return;
                }
                
                currentStudentIndex = this.getAttribute('data-index');
                displaySedes(removeDuplicateSedes(sedesData));
                sedeModal.style.display = 'block';
            });
        });

        document.querySelectorAll('.clear-sede-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                const index = this.getAttribute('data-index');
                studentsData[index].sede = '';
                studentsData[index].sedeId = '';
                studentsData[index].plazasSede = 0;
                
                const rows = tableBody.querySelectorAll('tr');
                if (rows[index]) {
                    const sedeCell = rows[index].querySelector('.sede-cell');
                    if (sedeCell) {
                        sedeCell.textContent = '';
                        sedeCell.removeAttribute('data-plazas');
                    }
                }
                
                this.disabled = true;
                
                if (sedesData.length > 0) {
                    displaySedes(removeDuplicateSedes(sedesData));
                }
            });
        });

        exportExcelBtn.disabled = false;
    });

    // 4. Funciones para el modal de sedes
    function displaySedes(sedes) {
        sedeList.innerHTML = '';
        
        if (sedes.length === 0) {
            sedeList.innerHTML = '<li>No hay sedes disponibles</li>';
            return;
        }
        
        sedes.forEach(sede => {
            const plazasRestantes = sede.plazasDisponibles - getAsignacionesSede(sede.id);
            
            // Solo mostrar sedes con plazas disponibles
            if (plazasRestantes > 0) {
                const li = document.createElement('li');
                li.innerHTML = `
                    <strong>${sede.displayText}</strong>
                    <span class="plazas-badge">Plazas: ${plazasRestantes}</span>
                `;
                
                li.addEventListener('click', () => selectSede(sede));
                li.setAttribute('data-plazas', plazasRestantes);
                sedeList.appendChild(li);
            }
        });
    }

    function selectSede(sede) {
        if (currentStudentIndex === null || !studentsData[currentStudentIndex]) return;
        
        const plazasRestantes = sede.plazasDisponibles - getAsignacionesSede(sede.id);
        
        if (plazasRestantes <= 0) {
            alert('No hay plazas disponibles en esta sede');
            return;
        }
        
        // Asignar la sede al alumno
        studentsData[currentStudentIndex].sede = sede.displayText;
        studentsData[currentStudentIndex].sedeId = sede.id;
        studentsData[currentStudentIndex].plazasSede = sede.plazasDisponibles;
        
        // Actualizar la tabla
        const rows = tableBody.querySelectorAll('tr');
        if (rows[currentStudentIndex]) {
            const sedeCell = rows[currentStudentIndex].querySelector('.sede-cell');
            const clearBtn = rows[currentStudentIndex].querySelector('.clear-sede-btn');
            
            if (sedeCell) {
                sedeCell.textContent = sede.displayText;
                sedeCell.setAttribute('data-plazas', plazasRestantes - 1);
            }
            
            // Habilitar el botón "Limpiar sede" para ESTE alumno
            if (clearBtn) {
                clearBtn.disabled = false;
            }
        }
        
        // Actualizar lista de sedes
        displaySedes(removeDuplicateSedes(sedesData));
        sedeModal.style.display = 'none';
    }

    // 5. Exportar a Excel
    function exportToExcel() {
        if (studentsData.length === 0) {
            alert('No hay datos para exportar');
            return;
        }
    
        const excelData = studentsData.map(student => {
            const rowData = {};
            
            // Usar los nombres de columna originales
            window.originalHeaders.forEach(header => {
                rowData[header] = student[header] || '';
            });
            
            // Añadir la sede si existe
            if (student.sede) {
                rowData['Sede'] = student.sede;
            }
            
            return rowData;
        });
    
        const worksheet = XLSX.utils.json_to_sheet(excelData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Alumnos');
    
        const fecha = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(workbook, `Alumnos_${fecha}.xlsx`);
    }

    // Funciones auxiliares
    function findHeaderRowIndex(data) {
        for (let i = 0; i < Math.min(5, data.length); i++) {
            const row = data[i];
            if (row && row.some(cell => 
                String(cell).toLowerCase().includes('institución') || 
                String(cell).toLowerCase().includes('sede'))) {
                return i;
            }
        }
        return -1;
    }

    function removeDuplicateSedes(sedes) {
        const unique = [];
        const seen = new Set();
        
        sedes.forEach(sede => {
            if (!seen.has(sede.id)) {
                seen.add(sede.id);
                unique.push(sede);
            }
        });
        
        return unique.sort((a, b) => a.displayText.localeCompare(b.displayText));
    }

    function getAsignacionesSede(sedeId) {
        return studentsData.filter(student => {
            const studentSedeId = student.sede?.split(' - ').join('-');
            return studentSedeId === sedeId;
        }).length;
    }

    function filterStudents(searchTerm) {
        searchTerm = searchTerm.toLowerCase().trim();
        const rows = tableBody.querySelectorAll('tr');
        
        if (!searchTerm) {
            rows.forEach(row => row.style.display = '');
            return;
        }
    
        const searchWords = searchTerm.split(/\s+/).filter(word => word.length > 0);
        const isAccountNumberSearch = /^\d+$/.test(searchTerm);
        
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            let shouldShow = false;
            
            // Caso 1: Búsqueda por número de cuenta (solo dígitos)
            if (isAccountNumberSearch) {
                const accountCell = Array.from(cells).find(cell => 
                    cell.getAttribute('data-account') !== null
                );
                shouldShow = accountCell?.textContent.toLowerCase().includes(searchTerm);
            }
            // Caso 2: Búsqueda por nombre completo (combinación de nombre + apellidos)
            else if (searchWords.length >= 2) {
                // Extraer texto de todas las celdas (excepto acción)
                const fullText = Array.from(cells)
                    .slice(0, -1)
                    .map(cell => cell.textContent.toLowerCase())
                    .join('|'); // Separador único
                    
                // Verificar si TODAS las palabras existen en cualquier orden
                shouldShow = searchWords.every(word => 
                    fullText.includes(word)
                );
                
                // Opcional: Priorizar coincidencias exactas de nombre completo
                if (!shouldShow) {
                    const combinedText = Array.from(cells)
                        .slice(0, -1)
                        .map(cell => cell.textContent.toLowerCase())
                        .join(' ');
                    shouldShow = combinedText.includes(searchTerm);
                }
            }
            // Caso 3: Búsqueda simple (1 palabra)
            else {
                shouldShow = Array.from(cells)
                    .slice(0, -1)
                    .some(cell => 
                        cell.textContent.toLowerCase().includes(searchWords[0])
                    );
            }
            
            row.style.display = shouldShow ? '' : 'none';
        });
    }

    function filterSedes(searchTerm) {
        searchTerm = searchTerm.toLowerCase();
        const items = sedeList.querySelectorAll('li');
        
        items.forEach(item => {
            const text = item.textContent.toLowerCase();
            const plazas = parseInt(item.getAttribute('data-plazas')) || 0;
            
            const matchesSearch = text.includes(searchTerm);
            const hasPlazas = plazas > 0 || searchTerm.includes('agotad');
            
            item.style.display = (matchesSearch && hasPlazas) ? '' : 'none';
        });
    }

    // Event listeners
    searchInput.addEventListener('input', function() {
        filterStudents(this.value);
    });
    
    modalSearchInput.addEventListener('input', function() {
        filterSedes(this.value);
    });
    
    closeBtn.addEventListener('click', function() {
        sedeModal.style.display = 'none';
    });
    
    window.addEventListener('click', function(event) {
        if (event.target === sedeModal) {
            sedeModal.style.display = 'none';
        }
    });
    
    exportExcelBtn.addEventListener('click', exportToExcel);
    
    clearBtn.addEventListener('click', function() {
        excelFileInput.value = '';
        sedeFileInput.value = '';
        searchInput.value = '';
        tableBody.innerHTML = '';
        studentsData = [];
        sedesData = [];
        totalStudentsSpan.textContent = '0';
        exportBtn.disabled = true;
        exportExcelBtn.disabled = true;
        clearBtn.disabled = true;
        // Oculta la tabla
        document.getElementById('table-container').style.display = 'none';
        document.getElementById('students-table').style.display = 'none';
    });
});