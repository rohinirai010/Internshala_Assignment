const XLSX = require('xlsx');
const moment = require('moment');

const excelFilePath = 'employee_records.xlsx';

const workbook = XLSX.readFile(excelFilePath);

const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const data = XLSX.utils.sheet_to_json(worksheet);

function areDatesConsecutive(date1, date2) {
    try {
        const day1 = moment(date1);
        const day2 = moment(date2);
        return day2.diff(day1, 'days') === 1;
    } catch (error) {
        console.log(error)
    }
}

const consecutiveWorkdays = new Map();

const dateFormat = 'MM-DD-YYYY hh:mm A';

for (const entry of data) {
    const employeeName = entry['Employee Name'];
    const timeStr = entry['Time'];

    if (timeStr) {
        const time = moment(timeStr, dateFormat);

        if (time.isValid()) {
            const previousTime = consecutiveWorkdays.get(employeeName);

            if (previousTime && areDatesConsecutive(previousTime, time)) {

                consecutiveWorkdays.set(employeeName, time);
            } else {

                consecutiveWorkdays.set(employeeName, time);
            }
        } else {
            // Handle invalid date
            console.log(`Invalid date for employee: ${employeeName}, Date: ${timeStr}`);
        }
    }
}
// Check which employees have worked 7 consecutive days
for (const [employeeName, lastWorkday] of consecutiveWorkdays.entries()) {
    const today = moment();
    const consecutiveDays = today.diff(lastWorkday, 'days');

    console.log(`Employee: ${employeeName}, Last Workday: ${lastWorkday}, Consecutive Days: ${consecutiveDays}`);

    if (consecutiveDays >= 7) {
        console.log(`Employee: ${employeeName}, Consecutive Workdays: ${consecutiveDays}`);
    }
}

console.log("end")