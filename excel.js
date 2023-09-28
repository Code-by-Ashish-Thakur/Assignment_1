const XLSX = require('xlsx');
const moment = require('moment');

// Load the Excel file (update the file path)
const workbook = XLSX.readFile('Assignment_Timecard.xlsx');

// Assuming the data is in the first sheet, you can change it if needed
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Initialize arrays to store the results
const consecutiveDaysEmployees = [];
const lessThan10HoursEmployees = [];
const moreThan14HoursEmployees = [];

// Helper function to calculate the time difference in hours between two timestamps
function calculateHourDifference(start, end) {
  const startTime = moment(start, 'MM/DD/YYYY HH:mm:ss');
  const endTime = moment(end, 'MM/DD/YYYY HH:mm:ss');
  const duration = moment.duration(endTime.diff(startTime));
  return duration.asHours();
}

// Iterate over the records
const range = XLSX.utils.decode_range(sheet['!ref']);
for (let row = range.s.r + 1; row <= range.e.r; row++) {
  const cellPosition = 'A' + (row + 1);
  const cellStatus = 'B' + (row + 1);
  const cellTimeIn = 'C' + (row + 1);
  const cellTimeOut = 'D' + (row + 1);
  const cellTimecardHours = 'E' + (row + 1);
  const cellStartDate = 'G' + (row + 1);
  const cellEndDate = 'H' + (row + 1);
  const cellName = 'I' + (row + 1);

  const position = sheet[cellPosition] ? sheet[cellPosition].v : '';
  const status = sheet[cellStatus] ? sheet[cellStatus].v : '';
  const timeIn = sheet[cellTimeIn] ? sheet[cellTimeIn].w : ''; // Use .w for date/time cells
  const timeOut = sheet[cellTimeOut] ? sheet[cellTimeOut].w : ''; // Use .w for date/time cells
  const timecardHours = sheet[cellTimecardHours] ? sheet[cellTimecardHours].v : '';
  const startDate = sheet[cellStartDate] ? sheet[cellStartDate].w : ''; // Use .w for date cells
  const endDate = sheet[cellEndDate] ? sheet[cellEndDate].w : ''; // Use .w for date cells
  const name = sheet[cellName] ? sheet[cellName].v : '';

  // Check for consecutive days
  const current = moment(startDate, 'MM/DD/YYYY');
  const end = moment(endDate, 'MM/DD/YYYY');
  let consecutiveDays = 1;

  while (current.isBefore(end)) {
    const nextDay = current.clone().add(1, 'days');
    if (moment(status, 'HH:mm:ss').isBefore(moment('00:00:00', 'HH:mm:ss')) &&
        moment(timeIn, 'HH:mm:ss').isBefore(moment('00:00:00', 'HH:mm:ss')) &&
        moment(timeOut, 'HH:mm:ss').isBefore(moment('00:00:00', 'HH:mm:ss'))) {
      consecutiveDays++;
    } else {
      break;
    }
    current.add(1, 'days');
  }

  if (consecutiveDays === 7) {
    consecutiveDaysEmployees.push({ name, position });
  }

  // Check for less than 10 hours between shifts
  if (moment(timeIn, 'HH:mm:ss').diff(moment(timeOut, 'HH:mm:ss'), 'hours') > 1 && moment(timeIn, 'HH:mm:ss').diff(moment(timeOut, 'HH:mm:ss'), 'hours') < 10) {
    lessThan10HoursEmployees.push({ name, position });
  }

  // Check for more than 14 hours in a single shift
  if (parseFloat(timecardHours) > 14) {
    moreThan14HoursEmployees.push({ name, position });
  }
}

// Print the results
console.log('Employees who have worked for 7 consecutive days:');
console.log(consecutiveDaysEmployees);

console.log('Employees who have less than 10 hours between shifts but greater than 1 hour:');
console.log(lessThan10HoursEmployees);

console.log('Employees who have worked for more than 14 hours in a single shift:');
console.log(moreThan14HoursEmployees);
