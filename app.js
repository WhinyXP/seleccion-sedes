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
            
            studentsData = [];
            for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row && row.length >= 9) {
                    studentsData.push({
                        no: row[0],
                        grupo: row[1],
                        planEstudio: row[2],
                        apellidoPaterno: row[3],
                        apellidoMaterno: row[4],
                        nombres: row[5],
                        numeroCuenta: row[6],
                        promedio: row[7],
                        correo: row[8],
                        sede: '',
                        sedeId: '',
                        plazasSede: 0
                    });
                }
            }
            
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
    exportBtn.addEventListener('click', function() {
        if (studentsData.length === 0) {
            alert('No hay datos para mostrar');
            return;
        }
        
        tableBody.innerHTML = '';
        
        studentsData.forEach((student, index) => {
            const row = document.createElement('tr');
            
            row.innerHTML = `
                <td>${student.no || ''}</td>
                <td>${student.grupo || ''}</td>
                <td>${student.planEstudio || ''}</td>
                <td>${student.apellidoPaterno || ''}</td>
                <td>${student.apellidoMaterno || ''}</td>
                <td>${student.nombres || ''}</td>
                <td>${student.numeroCuenta || ''}</td>
                <td>${student.promedio || ''}</td>
                <td>${student.correo || ''}</td>
                <td class="sede-cell">${student.sede || ''}</td>
                <td class="action-buttons">
                    <button class="add-sede-btn" data-index="${index}">Añadir sede</button>
                    <button class="clear-sede-btn" data-index="${index}" 
                        ${!student.sede ? 'disabled' : ''}>Limpiar sede</button>
                </td>
            `;
            
            tableBody.appendChild(row);
        });;
        
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

        // Dentro del exportBtn.addEventListener, después de asignar el event listener para add-sede-btn:
        document.querySelectorAll('.clear-sede-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                const index = this.getAttribute('data-index');
                studentsData[index].sede = '';
                studentsData[index].sedeId = '';
                studentsData[index].plazasSede = 0;
                
                // Actualizar la tabla
                const rows = tableBody.querySelectorAll('tr');
                if (rows[index]) {
                    const sedeCell = rows[index].querySelector('.sede-cell');
                    if (sedeCell) {
                        sedeCell.textContent = '';
                        sedeCell.removeAttribute('data-plazas');
                    }
                }
                
                // Deshabilitar el botón "Limpiar sede" después de limpiar
                this.disabled = true;
                
                // Volver a mostrar las sedes (para actualizar disponibilidad)
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

        const excelData = studentsData.map(student => ({
            'No.': student.no,
            'Plan de estudio': student.planEstudio,
            'Apellido Paterno': student.apellidoPaterno,
            'Apellido Materno': student.apellidoMaterno,
            'Nombres': student.nombres,
            'Número de cuenta': student.numeroCuenta,
            'Promedio': student.promedio,
            'Correo': student.correo,
            'Sede': student.sede
        }));

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
        searchTerm = searchTerm.toLowerCase();
        const rows = tableBody.querySelectorAll('tr');
        
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            const name = cells[5].textContent.toLowerCase();
            const lastName = cells[3].textContent.toLowerCase();
            const studentId = cells[6].textContent.toLowerCase();
            const sede = cells[9].textContent.toLowerCase(); // Nueva columna de sede
            
            if (
                name.includes(searchTerm) || 
                lastName.includes(searchTerm) || 
                studentId.includes(searchTerm) || 
                sede.includes(searchTerm) // Ahora también busca por sede
            ) {
                row.style.display = '';
            } else {
                row.style.display = 'none';
            }
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
    });
});