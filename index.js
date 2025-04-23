// ========== CONFIGURACIÓN HORARIO LABORAL Y FESTIVOS ==========
const workHours = {
    1: [[480, 720], [840, 1080]],  // Lunes: 8:00-12:00, 14:00-18:00 (en minutos)
    2: [[480, 720], [840, 1080]],   // Martes
    3: [[480, 720], [840, 1080]],   // Miércoles
    4: [[480, 720], [840, 1080]],   // Jueves
    5: [[480, 720], [840, 1080]],   // Viernes
    6: [[480, 720]]                 // Sábado solo mañana
};

const fixedHolidays = [
    '01-01', '05-01', '07-20', '08-07', '12-25'
];

const movableHolidays = [
    '01-06', '03-19', '06-29', '08-15', '10-12', '11-01', '11-11'
];

// ========== FUNCIONES PARA CÁLCULO DE TIEMPO ==========
function excelDateToJSDate(serial) {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    const total_seconds = Math.floor(86400 * fractional_day);
    const seconds = total_seconds % 60;
    const minutes = Math.floor(total_seconds / 60) % 60;
    const hours = Math.floor(total_seconds / 3600);

    return new Date(
        date_info.getFullYear(),
        date_info.getMonth(),
        date_info.getDate(),
        hours,
        minutes,
        seconds
    );
}

function calculateEaster(year) {
    const a = year % 19;
    const b = Math.floor(year / 100);
    const c = year % 100;
    const d = Math.floor(b / 4);
    const e = b % 4;
    const f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3);
    const h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4);
    const k = c % 4;
    const l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const month = Math.floor((h + l - 7 * m + 114) / 31);
    const day = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(year, month - 1, day);
}

function getHolyWeekDates(year) {
    const easter = calculateEaster(year);
    const dates = {
        juevesSanto: new Date(easter),
        viernesSanto: new Date(easter)
    };
    dates.juevesSanto.setDate(easter.getDate() - 3);
    dates.viernesSanto.setDate(easter.getDate() - 2);
    return dates;
}

function getActualHolidayDate(year, monthDay) {
    const [month, day] = monthDay.split('-').map(Number);
    const date = new Date(year, month - 1, day);
    if (date.getDay() === 1) return date;
    const daysToAdd = (8 - date.getDay()) % 7;
    const movedDate = new Date(date);
    movedDate.setDate(date.getDate() + daysToAdd);
    return movedDate;
}

function isHoliday(date) {
    const year = date.getFullYear();
    const monthDay = `${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;

    if (fixedHolidays.includes(monthDay)) return true;

    for (const holiday of movableHolidays) {
        const actualHolidayDate = getActualHolidayDate(year, holiday);
        if (actualHolidayDate.toDateString() === date.toDateString()) {
            return true;
        }
    }

    const holyWeek = getHolyWeekDates(year);
    if (
        date.toDateString() === holyWeek.juevesSanto.toDateString() ||
        date.toDateString() === holyWeek.viernesSanto.toDateString()
    ) {
        return true;
    }

    return false;
}

function calculateWorkedTime(startDateTime, endDateTime) {
    let totalMinutes = 0;
    let workDays = 0;
    let currentDateTime = new Date(startDateTime);

    while (currentDateTime <= endDateTime) {
        const dayOfWeek = currentDateTime.getDay();

        if (dayOfWeek !== 0 && !isHoliday(currentDateTime) && workHours[dayOfWeek]) {
            let dayHasWork = false;
            
            workHours[dayOfWeek].forEach(period => {
                const periodStart = period[0];
                const periodEnd = period[1];

                const currentDayStart = new Date(currentDateTime);
                currentDayStart.setHours(0, 0, 0, 0);

                const periodStartDateTime = new Date(currentDayStart);
                periodStartDateTime.setMinutes(periodStart);

                const periodEndDateTime = new Date(currentDayStart);
                periodEndDateTime.setMinutes(periodEnd);

                const effectiveStart = new Date(Math.max(currentDateTime, periodStartDateTime));
                const effectiveEnd = new Date(Math.min(endDateTime, periodEndDateTime));

                if (effectiveEnd > effectiveStart) {
                    const minutesWorked = (effectiveEnd - effectiveStart) / 60000;
                    totalMinutes += minutesWorked;
                    dayHasWork = true;
                }
            });

            if (dayHasWork) workDays++;
        }

        currentDateTime.setDate(currentDateTime.getDate() + 1);
        currentDateTime.setHours(0, 0, 0, 0);
    }

    const hours = Math.floor(totalMinutes / 60);
    const minutes = Math.floor(totalMinutes % 60);

    return { hours, minutes, workDays };
}

// ========== PROCESAMIENTO DEL EXCEL ==========
function processExcel() {
    const files = document.getElementById('file_upload').files;
    if (files.length === 0) {
        alert("Por favor seleccione un archivo Excel.");
        return;
    }

    const filename = files[0].name;
    const extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension !== '.XLS' && extension !== '.XLSX') {
        alert("Por favor seleccione un archivo Excel válido (.xls o .xlsx).");
        return;
    }

    const reader = new FileReader();
    reader.readAsBinaryString(files[0]);
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { raw: true });

        const tabla = document.getElementById("tabla-resultados");
        const tbody = tabla.querySelector("tbody");
        tbody.innerHTML = "";
        console.log(tbody +'1');
        console.log(tabla +'2');
        jsonData.forEach(row => {
            const inicio = row["Hora de creación (Ticket)"] || row["Inicio"] || row["Fecha Inicio"];
            const fin = row["Ticket Tiempo terminado"] || row["Fin"] || row["Fecha Fin"];

            if (typeof inicio === "number" && typeof fin === "number") {
                const fechaInicio = excelDateToJSDate(inicio);
                const fechaFin = excelDateToJSDate(fin);
                
                const { hours, minutes, workDays } = calculateWorkedTime(fechaInicio, fechaFin);

                const rowElement = document.createElement("tr");
                rowElement.innerHTML = `
                    <td>${fechaInicio.toLocaleString()}</td>
                    <td>${fechaFin.toLocaleString()}</td>
                    <td>${hours}</td>
                    <td>${minutes}</td>
                    <td>${workDays}</td>
                `;
                tbody.appendChild(rowElement);
                console.log(rowElement.innerHTML +'3');
                console.log(rowElement +'4');
            }
        });

        document.getElementById("results-section").style.display = "block";
    };
}