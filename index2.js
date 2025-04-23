// ========== CONFIGURACIÓN ==========
const workHours = {
    1: [[8, 12], [14, 18]],  // Lunes (horas)
    2: [[8, 12], [14, 18]],   // Martes
    3: [[8, 12], [14, 18]],   // Miércoles
    4: [[8, 12], [14, 18]],   // Jueves
    5: [[8, 12], [14, 18]],   // Viernes
    6: [[8, 12]]              // Sábado
};

const fixedHolidays = [
    '01-01', // Año Nuevo
    '05-01', // Día del Trabajo
    '07-20', // Independencia Colombia
    '08-07', // Batalla de Boyacá
    '12-25'  // Navidad
];

// ========== FUNCIONES AUXILIARES ==========
function parseExcelDate(dateValue) {
    // Si el valor está vacío
    if (dateValue === null || dateValue === undefined || dateValue === '') {
        return null;
    }

    // 1. Si es número (fecha serial de Excel)
    if (typeof dateValue === 'number') {
        const date = new Date((dateValue - 25569) * 86400 * 1000);
        date.setHours(date.getHours() + 5); // Ajuste zona horaria
        return date;
    }

    // 2. Si es texto, probar múltiples formatos
    if (typeof dateValue === 'string') {
        // Intentar formato m/d/yyyy h:mm
        let match = dateValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{1,2})$/);
        if (match) {
            return new Date(match[3], match[1]-1, match[2], match[4], match[5]);
        }

        // Intentar formato d/m/yyyy h:mm
        match = dateValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{1,2})$/);
        if (match) {
            return new Date(match[3], match[2]-1, match[1], match[4], match[5]);
        }

        // Intentar formato yyyy-mm-dd hh:mm
        match = dateValue.match(/^(\d{4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{1,2})$/);
        if (match) {
            return new Date(match[1], match[2]-1, match[3], match[4], match[5]);
        }

        // Intentar formato sin hora (asume 00:00)
        match = dateValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (match) {
            return new Date(match[3], match[1]-1, match[2]);
        }
    }

    // 3. Si ya es un objeto Date
    if (dateValue instanceof Date) {
        return dateValue;
    }

    console.warn('Formato de fecha no reconocido:', dateValue);
    return null;
}
function isHoliday(date) {
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return fixedHolidays.includes(`${month}-${day}`);
}

function isWeekend(date) {
    return date.getDay() === 0; // Domingo
}

// ========== CÁLCULO DE HORAS HÁBILES ==========
function calculateWorkingTime(start, end) {
    let totalMinutes = 0;
    let current = new Date(start);
    
    while (current <= end) {
        const dayOfWeek = current.getDay();
        
        // Saltar domingos y festivos
        if (isWeekend(current) || isHoliday(current)) {
            current.setDate(current.getDate() + 1);
            current.setHours(0, 0, 0, 0);
            continue;
        }
        
        // Obtener bloques laborales del día
        const dayBlocks = workHours[dayOfWeek] || [];
        
        for (const [startHour, endHour] of dayBlocks) {
            // Crear fechas para los límites del bloque
            const blockStart = new Date(current);
            blockStart.setHours(startHour, 0, 0, 0);
            
            const blockEnd = new Date(current);
            blockEnd.setHours(endHour, 0, 0, 0);
            
            // Determinar el período efectivo de trabajo
            const effectiveStart = new Date(Math.max(current, blockStart));
            const effectiveEnd = new Date(Math.min(end, blockEnd));
            
            if (effectiveStart < effectiveEnd) {
                totalMinutes += (effectiveEnd - effectiveStart) / 60000; // milisegundos a minutos
            }
        }
        
        // Avanzar al siguiente día
        current.setDate(current.getDate() + 1);
        current.setHours(0, 0, 0, 0);
    }
    
    return {
        hours: Math.floor(totalMinutes / 60),
        minutes: Math.round(totalMinutes % 60)
    };
}

// ========== PROCESAMIENTO DEL ARCHIVO ==========
function processExcel() {
    const fileInput = document.getElementById('file_upload');
    if (!fileInput.files.length) {
        alert('Por favor seleccione un archivo Excel');
        return;
    }
    
    const file = fileInput.files[0];
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            const headers = jsonData[0];
            const startCol = headers.findIndex(h => 
                h.toString().toLowerCase().includes('hora de creación'));
            const endCol = headers.findIndex(h => 
                h.toString().toLowerCase().includes('tiempo terminado'));
            
            if (startCol === -1 || endCol === -1) {
                alert('El archivo debe contener columnas "Hora de creación" y "Tiempo terminado"');
                return;
            }
            
            const tbody = document.querySelector('#tabla-resultados tbody');
            tbody.innerHTML = '';
            
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (!row[startCol] || !row[endCol]) continue;
                
                const startDate = parseExcelDate(row[startCol]);
                const endDate = parseExcelDate(row[endCol]);
                
                if (!startDate || !endDate) {
                    tbody.innerHTML += `
                    <tr>
                        <td>${row[startCol] || 'N/A'}</td>
                        <td>${row[endCol] || 'N/A'}</td>
                        <td colspan="2">Fecha inválida</td>
                    </tr>`;
                    continue;
                }
                
                
                if (startDate && endDate === null) {
                    tbody.innerHTML += `
                    <tr>
                        <td>${'N/A'}</td>
                        <td>${'N/A'}</td>
                        <td colspan="2">Fechas y horas vacías </td>
                    </tr>`;
                    continue;
                }

                const { hours, minutes } = calculateWorkingTime(startDate, endDate);
                
                // Resaltar filas que incluyen fines de semana o festivos
                let rowClass = '';
                if (isWeekend(startDate) || isWeekend(endDate)) {
                    rowClass = 'weekend';
                } else if (isHoliday(startDate) || isHoliday(endDate)) {
                    rowClass = 'holiday';
                }
                
                tbody.innerHTML += `
                <tr class="${rowClass}">
                    <td>${startDate.toLocaleString()}</td>
                    <td>${endDate.toLocaleString()}</td>
                    <td>${hours}</td>
                    <td>${minutes}</td>
                </tr>`;
            }
            
            document.getElementById('results-section').style.display = 'block';
            
        } catch (error) {
            console.error('Error procesando archivo:', error);
            alert('Ocurrió un error al procesar el archivo: ' + error.message);
        }
    };
    
    reader.readAsArrayBuffer(file);
}