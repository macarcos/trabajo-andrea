// Variables globales
let data1 = [];
let data2 = [];
let headers1 = [];
let headers2 = [];
let columnMapping = {
    cedula1: '',
    cedula2: '',
    nombre1: '',
    nombre2: '',
    lugar1: '',
    lugar2: ''
};
let errors = {
    duplicados: [],
    nombresEntreverados: [],
    cedulasIncorrectas: [],
    nombresIncorrectos: [],
    lugaresIncorrectos: [],
    personasInexistentes: []
};

// Función para normalizar texto
function normalizeText(text) {
    if (!text) return '';
    return text.toString().trim().toUpperCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^\w\s]/g, ' ')
        .replace(/\s+/g, ' ');
}

// Función para extraer palabras de un nombre
function extractWords(name) {
    const normalized = normalizeText(name);
    return normalized.split(' ').filter(word => word.length > 1);
}

// Función para comparar nombres entreverados
function compareScrambledNames(name1, name2) {
    const words1 = extractWords(name1);
    const words2 = extractWords(name2);
    
    if (words1.length === 0 || words2.length === 0) return false;
    if (words1.length !== words2.length) return false;
    
    // Crear copias para no modificar los originales
    const sortedWords1 = [...words1].sort();
    const sortedWords2 = [...words2].sort();
    
    // Comparar si tienen las mismas palabras
    return JSON.stringify(sortedWords1) === JSON.stringify(sortedWords2);
}

// Función para manejar la carga de archivos
function handleFileUpload(fileNumber) {
    const fileInput = document.getElementById(`file${fileNumber}`);
    const file = fileInput.files[0];
    
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            if (fileNumber === 1) {
                data1 = jsonData;
                headers1 = jsonData.length > 0 ? Object.keys(jsonData[0]) : [];
                updateFileInfo(1, file.name, jsonData.length);
                populateSelects(1, headers1);
            } else {
                data2 = jsonData;
                headers2 = jsonData.length > 0 ? Object.keys(jsonData[0]) : [];
                updateFileInfo(2, file.name, jsonData.length);
                populateSelects(2, headers2);
            }
            
            // Mostrar sección de mapeo si ambos archivos están cargados
            if (data1.length > 0 && data2.length > 0) {
                document.getElementById('mappingSection').classList.add('show');
            }
            
            resetResults();
            
        } catch (error) {
            alert(`Error al leer el archivo ${fileNumber}: ${error.message}`);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Función para actualizar información del archivo
function updateFileInfo(fileNumber, fileName, recordCount) {
    const fileInfo = document.getElementById(`fileInfo${fileNumber}`);
    const uploadBox = document.getElementById(`uploadBox${fileNumber}`);
    
    fileInfo.innerHTML = `✅ ${fileName} (${recordCount} registros)`;
    fileInfo.style.display = 'flex';
    uploadBox.classList.add('active');
}

// Función para poblar los selects con las columnas
function populateSelects(docNumber, headers) {
    const selects = ['cedula', 'nombre', 'lugar'];
    
    selects.forEach(type => {
        const select = document.getElementById(`${type}${docNumber}`);
        select.innerHTML = `<option value="">Seleccionar columna de ${type}...</option>`;
        
        headers.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            select.appendChild(option);
        });
    });
}

// Función para actualizar el mapeo
function updateMapping() {
    columnMapping = {
        cedula1: document.getElementById('cedula1').value,
        cedula2: document.getElementById('cedula2').value,
        nombre1: document.getElementById('nombre1').value,
        nombre2: document.getElementById('nombre2').value,
        lugar1: document.getElementById('lugar1').value,
        lugar2: document.getElementById('lugar2').value
    };
    
    const validateBtn = document.getElementById('validateBtn');
    const warning = document.getElementById('validationWarning');
    
    if (columnMapping.cedula1 && columnMapping.cedula2 && 
        columnMapping.nombre1 && columnMapping.nombre2) {
        validateBtn.disabled = false;
        warning.style.display = 'none';
    } else {
        validateBtn.disabled = true;
        warning.style.display = 'block';
    }
}

// Función para alternar la visibilidad del mapeo
function toggleMapping() {
    const mappingSection = document.getElementById('mappingSection');
    const toggleBtn = document.querySelector('.toggle-mapping');
    
    if (mappingSection.classList.contains('show')) {
        mappingSection.classList.remove('show');
        toggleBtn.textContent = 'Mostrar Mapeo';
    } else {
        mappingSection.classList.add('show');
        toggleBtn.textContent = 'Ocultar Mapeo';
    }
}

// Función para detectar duplicados MEJORADA
function detectDuplicates(data, mappings, docNumber) {
    const duplicates = [];
    const seen = new Map();
    
    data.forEach((row, index) => {
        const cedula = normalizeText(row[mappings.cedula] || '');
        const nombre = normalizeText(row[mappings.nombre] || '');
        const lugar = normalizeText(row[mappings.lugar] || '');
        
        if (!cedula) return;
        
        // Primera verificación: duplicado exacto
        const exactKey = `${cedula}|${nombre}|${lugar}`;
        
        if (seen.has(exactKey)) {
            const firstOccurrence = seen.get(exactKey);
            duplicates.push({
                tipo: 'DUPLICADO_EXACTO',
                cedula: cedula,
                nombre: nombre,
                lugar: lugar,
                filaActual: index + 2,
                filaPrimera: firstOccurrence,
                documento: docNumber,
                datosOriginales: row
            });
            return; // No seguir procesando este registro
        }
        
        // Segunda verificación: misma cédula y lugar pero nombre entreverado
        let encontradoDuplicadoEntreverado = false;
        for (let [existingKey, filaAnterior] of seen.entries()) {
            const [existingCedula, existingNombre, existingLugar] = existingKey.split('|');
            
            // Si es diferente cédula o lugar, continuar
            if (existingCedula !== cedula || existingLugar !== lugar) continue;
            
            // Si ya lo procesamos como duplicado exacto, continuar
            if (existingNombre === nombre) continue;
            
            // Verificar si los nombres están entreverados
            if (compareScrambledNames(existingNombre, nombre)) {
                duplicates.push({
                    tipo: 'DUPLICADO_NOMBRE_ENTREVERADO',
                    cedula: cedula,
                    nombre: `${nombre} (Original: ${existingNombre})`,
                    lugar: lugar,
                    filaActual: index + 2,
                    filaPrimera: filaAnterior,
                    documento: docNumber,
                    datosOriginales: row
                });
                encontradoDuplicadoEntreverado = true;
                break;
            }
        }
        
        // Solo agregar al mapa si no es un duplicado
        if (!encontradoDuplicadoEntreverado) {
            seen.set(exactKey, index + 2);
        }
    });
    
    return duplicates;
}

// Función principal de validación MEJORADA
function validateDocuments() {
    if (data1.length === 0 || data2.length === 0) {
        alert('Por favor, suba ambos archivos Excel antes de validar.');
        return;
    }

    if (!columnMapping.cedula1 || !columnMapping.cedula2 || 
        !columnMapping.nombre1 || !columnMapping.nombre2) {
        alert('Por favor, seleccione las columnas obligatorias (cédula y nombre) para ambos documentos.');
        return;
    }

    // Mostrar spinner
    const validateBtn = document.getElementById('validateBtn');
    const validateText = document.getElementById('validateText');
    const spinner = document.getElementById('spinner');
    
    validateBtn.disabled = true;
    validateText.textContent = 'Validando...';
    spinner.style.display = 'block';

    // Resetear errores
    errors = {
        duplicados: [],
        nombresEntreverados: [],
        cedulasIncorrectas: [],
        nombresIncorrectos: [],
        lugaresIncorrectos: [],
        personasInexistentes: []
    };

    // 1. Detectar duplicados en ambos documentos (incluyendo nombres entreverados como duplicados)
    const duplicados1 = detectDuplicates(data1, {
        cedula: columnMapping.cedula1,
        nombre: columnMapping.nombre1,
        lugar: columnMapping.lugar1
    }, 1);
    
    const duplicados2 = detectDuplicates(data2, {
        cedula: columnMapping.cedula2,
        nombre: columnMapping.nombre2,
        lugar: columnMapping.lugar2
    }, 2);

    errors.duplicados = [...duplicados1, ...duplicados2];

    // 2. Crear mapa del documento maestro
    const masterData = new Map();
    data1.forEach((row, index) => {
        const cedula = normalizeText(row[columnMapping.cedula1] || '');
        const nombre = normalizeText(row[columnMapping.nombre1] || '');
        const lugar = normalizeText(row[columnMapping.lugar1] || '');
        
        if (cedula && nombre) {
            masterData.set(cedula, {
                nombre: nombre,
                lugar: lugar,
                rowIndex: index + 2,
                originalRow: row
            });
        }
    });

    // 3. Validar documento 2 contra documento 1
    data2.forEach((row, index) => {
        const cedula = normalizeText(row[columnMapping.cedula2] || '');
        const nombre = normalizeText(row[columnMapping.nombre2] || '');
        const lugar = normalizeText(row[columnMapping.lugar2] || '');
        
        if (!cedula || !nombre) return;

        const masterRecord = masterData.get(cedula);
        
        if (!masterRecord) {
            // Buscar si existe la persona con nombre similar pero cédula diferente
            let encontradoPorNombre = false;
            
            for (let [masterCedula, masterInfo] of masterData.entries()) {
                // Comparación exacta de nombres
                if (masterInfo.nombre === nombre && masterCedula !== cedula) {
                    errors.cedulasIncorrectas.push({
                        tipo: 'CEDULA_INCORRECTA',
                        cedulaIncorrecta: cedula,
                        cedulaCorrecta: masterCedula,
                        nombre: nombre,
                        nombreCorrecto: nombre,
                        nombreIncorrecto: '',
                        lugarDoc1: masterInfo.lugar,
                        lugarDoc2: lugar,
                        filaDoc1: masterInfo.rowIndex,
                        filaDoc2: index + 2,
                        datosOriginales: { doc1: masterInfo.originalRow, doc2: row }
                    });
                    encontradoPorNombre = true;
                    break;
                }
                
                // Comparación de nombres entreverados (SOLO si no están en el mismo documento como duplicados)
                if (compareScrambledNames(masterInfo.nombre, nombre) && masterCedula !== cedula) {
                    // Verificar si ya fueron marcados como duplicados en el mismo documento
                    const yaEsDuplicadoEnDoc2 = errors.duplicados.some(dup => 
                        dup.documento === 2 && 
                        dup.cedula === cedula && 
                        (dup.tipo === 'DUPLICADO_NOMBRE_ENTREVERADO' || dup.tipo === 'DUPLICADO_EXACTO')
                    );
                    
                    if (!yaEsDuplicadoEnDoc2) {
                        errors.nombresEntreverados.push({
                            tipo: 'NOMBRE_ENTREVERADO',
                            cedula: cedula,
                            nombreCorrecto: masterInfo.nombre,
                            nombreEntreverado: nombre,
                            lugarDoc1: masterInfo.lugar,
                            lugarDoc2: lugar,
                            filaDoc1: masterInfo.rowIndex,
                            filaDoc2: index + 2,
                            datosOriginales: { doc1: masterInfo.originalRow, doc2: row }
                        });
                    }
                    encontradoPorNombre = true;
                    break;
                }
            }
            
            // Si no se encontró por nombre, es una persona inexistente
            if (!encontradoPorNombre) {
                errors.personasInexistentes.push({
                    tipo: 'PERSONA_INEXISTENTE',
                    cedula: cedula,
                    nombre: nombre,
                    lugar: lugar,
                    filaDoc2: index + 2,
                    datosOriginales: row
                });
            }
        } else {
            // La cédula existe, validar otros campos
            
            // Verificar si ya fue marcado como duplicado dentro del mismo documento
            const yaEsDuplicadoEnDoc2 = errors.duplicados.some(dup => 
                dup.documento === 2 && 
                dup.cedula === cedula && 
                (dup.tipo === 'DUPLICADO_NOMBRE_ENTREVERADO' || dup.tipo === 'DUPLICADO_EXACTO') &&
                dup.filaActual === index + 2
            );
            
            if (!yaEsDuplicadoEnDoc2) {
                // Validar nombres (exacto)
                if (masterRecord.nombre !== nombre) {
                    // Verificar si es nombre entreverado
                    if (compareScrambledNames(masterRecord.nombre, nombre)) {
                        errors.nombresEntreverados.push({
                            tipo: 'NOMBRE_ENTREVERADO',
                            cedula: cedula,
                            nombreCorrecto: masterRecord.nombre,
                            nombreEntreverado: nombre,
                            lugarDoc1: masterRecord.lugar,
                            lugarDoc2: lugar,
                            filaDoc1: masterRecord.rowIndex,
                            filaDoc2: index + 2,
                            datosOriginales: { doc1: masterRecord.originalRow, doc2: row }
                        });
                    } else {
                        // Nombre completamente diferente
                        errors.nombresIncorrectos.push({
                            tipo: 'NOMBRE_INCORRECTO',
                            cedula: cedula,
                            nombreCorrecto: masterRecord.nombre,
                            nombreIncorrecto: nombre,
                            lugarDoc1: masterRecord.lugar,
                            lugarDoc2: lugar,
                            filaDoc1: masterRecord.rowIndex,
                            filaDoc2: index + 2,
                            datosOriginales: { doc1: masterRecord.originalRow, doc2: row }
                        });
                    }
                }

                // Validar lugar de trabajo (solo si ambos documentos tienen esta columna)
                if (columnMapping.lugar1 && columnMapping.lugar2 &&
                    masterRecord.lugar && lugar && masterRecord.lugar !== lugar) {
                    errors.lugaresIncorrectos.push({
                        tipo: 'LUGAR_INCORRECTO',
                        cedula: cedula,
                        nombre: masterRecord.nombre,
                        lugarCorrecto: masterRecord.lugar,
                        lugarIncorrecto: lugar,
                        filaDoc1: masterRecord.rowIndex,
                        filaDoc2: index + 2,
                        datosOriginales: { doc1: masterRecord.originalRow, doc2: row }
                    });
                }
            }
        }
    });

    // Mostrar resultados
    displayResults();
    
    // Restaurar botón
    validateBtn.disabled = false;
    validateText.textContent = 'Validar Documentos';
    spinner.style.display = 'none';
}

// Función para mostrar los resultados
function displayResults() {
    const resultsSection = document.getElementById('resultsSection');
    resultsSection.style.display = 'block';

    const totalErrores = errors.duplicados.length + errors.nombresEntreverados.length + 
                        errors.cedulasIncorrectas.length + errors.nombresIncorrectos.length + errors.lugaresIncorrectos.length + 
                        errors.personasInexistentes.length;

    if (totalErrores === 0) {
        document.getElementById('successMessage').style.display = 'block';
        document.getElementById('exportSection').style.display = 'none';
        document.getElementById('errorTablesContainer').innerHTML = '';
        document.getElementById('errorSummary').innerHTML = '';
        return;
    }

    // Mostrar resumen de errores
    displayErrorSummary();
    
    // Mostrar tablas detalladas
    displayErrorTables();
    
    // Mostrar sección de exportación
    document.getElementById('exportSection').style.display = 'block';
    document.getElementById('successMessage').style.display = 'none';
    
    // Scroll hacia los resultados
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Función para mostrar resumen de errores
function displayErrorSummary() {
    const summaryContainer = document.getElementById('errorSummary');
    let summaryHTML = '';

    const errorTypes = [
        { key: 'duplicados', title: 'Datos Duplicados', icon: '📝', class: 'duplicates' },
        { key: 'nombresEntreverados', title: 'Nombres Entreverados', icon: '🔄', class: 'scrambled' },
        { key: 'cedulasIncorrectas', title: 'Cédulas Incorrectas', icon: '🆔', class: 'cedula' },
        { key: 'nombresIncorrectos', title: 'Nombres Incorrectos', icon: '❌', class: 'nombres' },
        { key: 'lugaresIncorrectos', title: 'Lugares Incorrectos', icon: '🏢', class: 'lugar' },
        { key: 'personasInexistentes', title: 'Personas Inexistentes', icon: '👤', class: 'missing' }
    ];

    errorTypes.forEach(errorType => {
        const count = errors[errorType.key].length;
        if (count > 0) {
            summaryHTML += `
                <div class="summary-card ${errorType.class}">
                    <div class="summary-card-header">
                        <h3>${errorType.title}</h3>
                        <span class="icon">${errorType.icon}</span>
                    </div>
                    <div class="error-count">${count}</div>
                    <div class="error-description">registros con inconsistencias</div>
                </div>
            `;
        }
    });

    summaryContainer.innerHTML = summaryHTML;
}

// Función para mostrar tablas detalladas de errores
function displayErrorTables() {
    const tablesContainer = document.getElementById('errorTablesContainer');
    let tablesHTML = '';

    // Tabla de duplicados
    if (errors.duplicados.length > 0) {
        tablesHTML += createErrorTable(
            'duplicados',
            '📝 Datos Duplicados',
            'duplicates',
            [
                { key: 'documento', label: 'Documento' },
                { key: 'cedula', label: 'Cédula' },
                { key: 'nombre', label: 'Nombre' },
                { key: 'lugar', label: 'Lugar' },
                { key: 'filaActual', label: 'Fila Actual' },
                { key: 'filaPrimera', label: 'Primera Aparición' },
                { key: 'tipo', label: 'Observación' }
            ],
            errors.duplicados.map(error => ({
                ...error,
                documento: `Documento ${error.documento}`
            }))
        );
    }

    // Tabla de nombres entreverados
    if (errors.nombresEntreverados.length > 0) {
        tablesHTML += createErrorTable(
            'entreverados',
            '🔄 Nombres Entreverados',
            'scrambled',
            [
                { key: 'cedula', label: 'Cédula' },
                { key: 'nombreCorrecto', label: 'Nombre Correcto' },
                { key: 'nombreEntreverado', label: 'Nombre Entreverado' },
                { key: 'lugarDoc1', label: 'Lugar Doc. Maestro' },
                { key: 'lugarDoc2', label: 'Lugar Doc. Validación' },
                { key: 'filaDoc1', label: 'Fila Doc. Maestro' },
                { key: 'filaDoc2', label: 'Fila Doc. Validación' },
                { key: 'tipo', label: 'Observación' }
            ],
            errors.nombresEntreverados
        );
    }

    // Tabla de cédulas incorrectas
    if (errors.cedulasIncorrectas.length > 0) {
        tablesHTML += createErrorTable(
            'cedulas',
            '🆔 Cédulas Incorrectas',
            'cedula',
            [
                { key: 'cedulaIncorrecta', label: 'Cédula Incorrecta' },
                { key: 'cedulaCorrecta', label: 'Cédula Correcta' },
                { key: 'nombre', label: 'Nombre' },
                { key: 'lugarDoc1', label: 'Lugar Doc. Maestro' },
                { key: 'lugarDoc2', label: 'Lugar Doc. Validación' },
                { key: 'filaDoc1', label: 'Fila Doc. Maestro' },
                { key: 'filaDoc2', label: 'Fila Doc. Validación' },
                { key: 'tipo', label: 'Observación' }
            ],
            errors.cedulasIncorrectas
        );
    }

    // Tabla de nombres incorrectos
    if (errors.nombresIncorrectos.length > 0) {
        tablesHTML += createErrorTable(
            'nombres',
            '❌ Nombres Incorrectos',
            'nombres',
            [
                { key: 'cedula', label: 'Cédula' },
                { key: 'nombreCorrecto', label: 'Nombre Correcto' },
                { key: 'nombreIncorrecto', label: 'Nombre Incorrecto' },
                { key: 'lugarDoc1', label: 'Lugar Doc. Maestro' },
                { key: 'lugarDoc2', label: 'Lugar Doc. Validación' },
                { key: 'filaDoc1', label: 'Fila Doc. Maestro' },
                { key: 'filaDoc2', label: 'Fila Doc. Validación' },
                { key: 'tipo', label: 'Observación' }
            ],
            errors.nombresIncorrectos
        );
    }

    // Tabla de lugares incorrectos
    if (errors.lugaresIncorrectos.length > 0) {
        tablesHTML += createErrorTable(
            'lugares',
            '🏢 Lugares de Trabajo Incorrectos',
            'lugar',
            [
                { key: 'cedula', label: 'Cédula' },
                { key: 'nombre', label: 'Nombre' },
                { key: 'lugarCorrecto', label: 'Lugar Correcto' },
                { key: 'lugarIncorrecto', label: 'Lugar Incorrecto' },
                { key: 'filaDoc1', label: 'Fila Doc. Maestro' },
                { key: 'filaDoc2', label: 'Fila Doc. Validación' },
                { key: 'tipo', label: 'Observación' }
            ],
            errors.lugaresIncorrectos
        );
    }

    // Tabla de personas inexistentes
    if (errors.personasInexistentes.length > 0) {
        tablesHTML += createErrorTable(
            'inexistentes',
            '👤 Personas Inexistentes',
            'missing',
            [
                { key: 'cedula', label: 'Cédula' },
                { key: 'nombre', label: 'Nombre' },
                { key: 'lugar', label: 'Lugar' },
                { key: 'filaDoc2', label: 'Fila' },
                { key: 'tipo', label: 'Observación' }
            ],
            errors.personasInexistentes
        );
    }

    tablesContainer.innerHTML = tablesHTML;
}

// Función auxiliar para crear tablas de errores
function createErrorTable(id, title, cssClass, columns, data) {
    let tableHTML = `
        <div class="error-table-section ${cssClass}">
            <h3 class="error-table-title">${title} (${data.length} registros)</h3>
            <div class="table-container">
                <table class="error-table">
                    <thead>
                        <tr>
    `;
    
    columns.forEach(col => {
        tableHTML += `<th>${col.label}</th>`;
    });
    
    tableHTML += `
                        </tr>
                    </thead>
                    <tbody>
    `;
    
    data.forEach(row => {
        tableHTML += '<tr>';
        columns.forEach(col => {
            const value = row[col.key] || '';
            const cellClass = col.key === 'tipo' ? 'error-cell' : '';
            tableHTML += `<td class="${cellClass}">${value}</td>`;
        });
        tableHTML += '</tr>';
    });
    
    tableHTML += `
                    </tbody>
                </table>
            </div>
        </div>
    `;
    
    return tableHTML;
}

// Función para exportar errores a Excel MEJORADA
function exportErrors() {
    const totalErrores = errors.duplicados.length + errors.nombresEntreverados.length + 
                        errors.cedulasIncorrectas.length + errors.nombresIncorrectos.length + errors.lugaresIncorrectos.length + 
                        errors.personasInexistentes.length;
                        
    if (totalErrores === 0) {
        alert('No hay errores para exportar.');
        return;
    }

    const wb = XLSX.utils.book_new();

    // Hoja de duplicados
    if (errors.duplicados.length > 0) {
        const duplicadosData = [
            ['CEDULA', 'NOMBRE', 'LUGAR', 'DOCUMENTO', 'FILA_ACTUAL', 'PRIMERA_APARICION', 'TIPO_ERROR', 'CAUSA'],
            ...errors.duplicados.map(error => [
                error.cedula || '',
                error.nombre || '',
                error.lugar || '',
                `Documento ${error.documento}`,
                error.filaActual || '',
                error.filaPrimera || '',
                error.tipo || '',
                error.tipo === 'DUPLICADO_EXACTO' ? 'Datos duplicados exactos' : 
                error.tipo === 'DUPLICADO_NOMBRE_ENTREVERADO' ? 'Misma persona, nombre entreverado' : 'Datos duplicados'
            ])
        ];
        const wsDuplicados = XLSX.utils.aoa_to_sheet(duplicadosData);
        XLSX.utils.book_append_sheet(wb, wsDuplicados, 'Duplicados');
    }

    // Hoja de nombres entreverados
    if (errors.nombresEntreverados.length > 0) {
        const entreveradosData = [
            ['CEDULA', 'NOMBRE_CORRECTO', 'NOMBRE_ENTREVERADO', 'LUGAR_DOC1', 'LUGAR_DOC2', 'FILA_DOC1', 'FILA_DOC2', 'CAUSA'],
            ...errors.nombresEntreverados.map(error => [
                error.cedula || '',
                error.nombreCorrecto || '',
                error.nombreEntreverado || '',
                error.lugarDoc1 || '',
                error.lugarDoc2 || '',
                error.filaDoc1 || '',
                error.filaDoc2 || '',
                'Nombre con palabras en diferente orden'
            ])
        ];
        const wsEntreverados = XLSX.utils.aoa_to_sheet(entreveradosData);
        XLSX.utils.book_append_sheet(wb, wsEntreverados, 'Nombres_Entreverados');
    }

    // Hoja de cédulas incorrectas
    if (errors.cedulasIncorrectas.length > 0) {
        const cedulasData = [
            ['CEDULA_INCORRECTA', 'CEDULA_CORRECTA', 'NOMBRE', 'LUGAR_DOC1', 'LUGAR_DOC2', 'FILA_DOC1', 'FILA_DOC2', 'CAUSA'],
            ...errors.cedulasIncorrectas.map(error => [
                error.cedulaIncorrecta || '',
                error.cedulaCorrecta || '',
                error.nombre || '',
                error.lugarDoc1 || '',
                error.lugarDoc2 || '',
                error.filaDoc1 || '',
                error.filaDoc2 || '',
                'Cédula no coincide con el nombre registrado'
            ])
        ];
        const wsCedulas = XLSX.utils.aoa_to_sheet(cedulasData);
        XLSX.utils.book_append_sheet(wb, wsCedulas, 'Cedulas_Incorrectas');
    }

    // Hoja de nombres incorrectos
    if (errors.nombresIncorrectos.length > 0) {
        const nombresData = [
            ['CEDULA', 'NOMBRE_CORRECTO', 'NOMBRE_INCORRECTO', 'LUGAR_DOC1', 'LUGAR_DOC2', 'FILA_DOC1', 'FILA_DOC2', 'CAUSA'],
            ...errors.nombresIncorrectos.map(error => [
                error.cedula || '',
                error.nombreCorrecto || '',
                error.nombreIncorrecto || '',
                error.lugarDoc1 || '',
                error.lugarDoc2 || '',
                error.filaDoc1 || '',
                error.filaDoc2 || '',
                'Nombre no coincide con la cédula registrada'
            ])
        ];
        const wsNombres = XLSX.utils.aoa_to_sheet(nombresData);
        XLSX.utils.book_append_sheet(wb, wsNombres, 'Nombres_Incorrectos');
    }

    // Hoja de lugares incorrectos
    if (errors.lugaresIncorrectos.length > 0) {
        const lugaresData = [
            ['CEDULA', 'NOMBRE', 'LUGAR_CORRECTO', 'LUGAR_INCORRECTO', 'FILA_DOC1', 'FILA_DOC2', 'CAUSA'],
            ...errors.lugaresIncorrectos.map(error => [
                error.cedula || '',
                error.nombre || '',
                error.lugarCorrecto || '',
                error.lugarIncorrecto || '',
                error.filaDoc1 || '',
                error.filaDoc2 || '',
                'Lugar de trabajo no coincide'
            ])
        ];
        const wsLugares = XLSX.utils.aoa_to_sheet(lugaresData);
        XLSX.utils.book_append_sheet(wb, wsLugares, 'Lugares_Incorrectos');
    }

    // Hoja de personas inexistentes
    if (errors.personasInexistentes.length > 0) {
        const inexistentesData = [
            ['CEDULA', 'NOMBRE', 'LUGAR', 'FILA_DOC2', 'CAUSA'],
            ...errors.personasInexistentes.map(error => [
                error.cedula || '',
                error.nombre || '',
                error.lugar || '',
                error.filaDoc2 || '',
                'Persona no existe en la base de datos principal'
            ])
        ];
        const wsInexistentes = XLSX.utils.aoa_to_sheet(inexistentesData);
        XLSX.utils.book_append_sheet(wb, wsInexistentes, 'Personas_Inexistentes');
    }

    // Generar y descargar archivo
    const fecha = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Inconsistencias_RRHH_${fecha}.xlsx`);
}

// Función para resetear resultados
function resetResults() {
    errors = {
        duplicados: [],
        nombresEntreverados: [],
        cedulasIncorrectas: [],
        nombresIncorrectos: [],
        lugaresIncorrectos: [],
        personasInexistentes: []
    };
    
    const resultsSection = document.getElementById('resultsSection');
    if (resultsSection) {
        resultsSection.style.display = 'none';
    }
    
    updateMapping();
}

// Event listeners
document.getElementById('file1').addEventListener('change', () => handleFileUpload(1));
document.getElementById('file2').addEventListener('change', () => handleFileUpload(2));

// Inicializar la página
document.addEventListener('DOMContentLoaded', function() {
    console.log('Sistema de Validación RRHH Mejorado - Iniciado');
    
    document.getElementById('mappingSection').classList.remove('show');
    document.getElementById('resultsSection').style.display = 'none';
});