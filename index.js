function calculateWorkedTime() {
    const startDate = document.getElementById('start-date').value;
    const startTime = document.getElementById('start-time').value;
    const endDate = document.getElementById('end-date').value;
    const endTime = document.getElementById('end-time').value;

    if (!startDate || !startTime || !endDate || !endTime) {
        alert('Por favor, ingrese ambas fechas y horas.');
        return;
    }

    const startDateTime = new Date(`${startDate}T${startTime}`);
    const endDateTime = new Date(`${endDate}T${endTime}`);

    if (endDateTime < startDateTime) {
        alert('La fecha y hora final no pueden ser anteriores a la fecha y hora de inicio.');
        return;
    }

    const workHours = {
        1: [[480, 720], [840, 1080]],  // Lunes: 8:00-12:00, 14:00-18:00
        2: [[480, 720], [840, 1080]],  // Martes: 8:00-12:00, 14:00-18:00
        3: [[480, 720], [840, 1080]],  // Miércoles: 8:00-12:00, 14:00-18:00
        4: [[480, 720], [840, 1080]],  // Jueves: 8:00-12:00, 14:00-18:00
        5: [[480, 720], [840, 1080]],  // Viernes: 8:00-12:00, 14:00-18:00
        6: [[480, 720]]                // Sábado: 8:00-12:00
    };

    let totalMinutes = 0;
    let currentDateTime = new Date(startDateTime);

    while (currentDateTime <= endDateTime) {
        const dayOfWeek = currentDateTime.getDay();
        
        if (dayOfWeek !== 0 && workHours[dayOfWeek]) { // Excluir domingos (dayOfWeek 0)
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
                    totalMinutes += (effectiveEnd - effectiveStart) / 60000;
                }
            });
        }

        currentDateTime.setDate(currentDateTime.getDate() + 1);
        currentDateTime.setHours(0, 0, 0, 0); // Reset time to the beginning of the next day
    }

    const hours = Math.floor(totalMinutes / 60);
    const minutes = Math.floor(totalMinutes % 60);

    document.getElementById('result').textContent = `Tiempo Transcurrido: ${hours} horas y ${minutes} minutos.`;
}

function convertToMinutes(time) {
    const [hours, minutes] = time.split(':').map(Number);
    return hours * 60 + minutes;
}


