# GoogleWorkspace
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Customize Format')
    .addItem('Format Form Responses', 'formatFormResponses')
    .addItem('Format Shortlist', 'formatShortlist')
    .addItem('Employee data', 'tableFormat')
    .addToUi();
}

function formatFormResponses() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let header = sheet.getRange('A1:N1'); // Adjusted range
  let table = sheet.getDataRange();

  header.setFontWeight('bold');
  header.setFontColor('black');
  header.setBackground('#ADD8E6');

  table.setFontFamily('Roboto');
  table.setHorizontalAlignment('center');
  table.setVerticalAlignment('middle');
  table.setBorder(true, true, true, true, false, true, '#ADD8E6', SpreadsheetApp.BorderStyle.SOLID);

  let position = sheet.getRange('K1'); // Adjusted column
  position.setValue('Position Applied');

  let salary = sheet.getRange('L1'); // Adjusted column
  salary.setValue('Expected Salary');

  let resume = sheet.getRange('M1'); // Adjusted column
  resume.setValue('Cover Letter / Resume');

  for (let col = 1; col <= sheet.getLastColumn(); col++) {
    sheet.autoResizeColumn(col);
  }

  // Call salaryRange function
  salaryRange('L2:L');

  // Call formatPositions function
  formatPositions('K2:K');
}


function salaryRange(salaryColRange) { // format salary bg colour according to salary range
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let salaryColumn = sheet.getRange(salaryColRange);  // Assuming the salary column is 'G' and starts from row 2
  let salaries = salaryColumn.getValues();

  for (let i = 0; i < salaries.length; i++) {  // Start from index 0 for correct offset
    let salary = salaries[i][0];
    let cell = salaryColumn.offset(i, 0, 1, 1);  // Use offset to get the correct cell

    if (salary && !isNaN(salary)) {  // Check if the cell is not empty and is a valid number
      if (salary < 3000) {
        cell.setBackground('#9ACD32');  // YellowGreen
      } else if (salary >= 3000 && salary <= 5000) {
        cell.setBackground('#FFD700');  // Gold
      } else if (salary > 5000) {
        cell.setBackground('#FF6347');  // Tomato
      }
    } else {
      cell.setBackground('');  // Clear background if not a valid number
    }
  }
}

function formatPositions(positionColRange) { // format applied position to numbering
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet.getRange(positionColRange);  // Adjust the range to the cells containing the positions
  let positions = range.getValues();

  for (let i = 0; i < positions.length; i++) {  // Start from index 0 for correct offset
    let position = positions[i][0];
    if (position) {  // Check if the cell is not empty
      let lines = position.split('\n');
      let isNumbered = lines.every(line => /^\d+\.\s/.test(line));  // Check if all lines are already numbered

      let formattedPositions;
      if (!isNumbered) {  // Only format if not already numbered
        formattedPositions = position.split(',').map((pos, index) => `${index + 1}. ${pos.trim()}`).join('\n');
      } else {
        formattedPositions = position;  // Keep the existing numbering
      }

      let cell = range.getCell(i + 1, 1);
      cell.setValue(formattedPositions);
      cell.setHorizontalAlignment('left');  // Set text alignment to left
    }
  }
}

function formatShortlist() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let header = sheet.getRange('A1:J1'); // Adjusted range
  let table = sheet.getDataRange();

  header.setFontWeight('bold');
  header.setFontColor('black');
  header.setBackground('#ADD8E6');

  table.setFontFamily('Roboto');
  table.setHorizontalAlignment('center');
  table.setVerticalAlignment('middle');
  table.setBorder(true, true, true, true, false, true, '#ADD8E6', SpreadsheetApp.BorderStyle.SOLID);

  for (let col = 1; col <= sheet.getLastColumn(); col++) {
    sheet.autoResizeColumn(col);
  }

  addDatePickers();
  addDropdownsToColumns();
  applyGreyBackgroundToNonApproved();
}


function applyGreyBackgroundToNonApproved() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  let dataValues = dataRange.getValues();

  for (let i = 0; i < dataValues.length; i++) {
    let statusCell = sheet.getRange(i + 2, 8); // Column H (8th column)
    let statusValue = statusCell.getValue();
    let notificationCell = sheet.getRange(i + 2, 9); // Column I
    let decisionCell = sheet.getRange(i + 2, 10); // Column J
    let decisionValue = decisionCell.getValue();

    if (statusValue !== 'Approved') {
      notificationCell.setBackground('#CCCCCC'); // Grey
      if (decisionValue !== 'Accepted' && decisionValue !== 'Rejected') {
        decisionCell.setBackground('#CCCCCC'); // Grey
      }
    } else {
      notificationCell.setBackground(null); // Reset background
      if (decisionValue !== 'Accepted' && decisionValue !== 'Rejected') {
        decisionCell.setBackground(null); // Reset background
      }
    }
  }
}



function shortList() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sourceSheet = spreadsheet.getSheetByName('Form Responses');
  let destinationSheet = spreadsheet.getSheetByName('Shortlisted Candidates');

  if (!sourceSheet || !destinationSheet) {
    SpreadsheetApp.getUi().alert('One or both of the sheets do not exist.');
    return;
  }

  let selectedRange = sourceSheet.getActiveRangeList();
  if (!selectedRange) {
    SpreadsheetApp.getUi().alert('Please select rows to copy.');
    return;
  }

  let columns = [2, 3, 7, 8]; // Adjusted column indices
  let numCols = columns.length;

  let destinationData = destinationSheet.getRange(2, 1, destinationSheet.getLastRow(), numCols).getValues();

  let destinationSet = new Set(destinationData.map(row => row.join()));

  let rowsCopied = 0;

  let ranges = selectedRange.getRanges();
  for (let range of ranges) {
    let rowStart = range.getRow();
    let rowEnd = rowStart + range.getNumRows() - 1;

    for (let row = rowStart; row <= rowEnd; row++) {
      let rowData = columns.map(col => sourceSheet.getRange(row, col).getValue());

      if (rowData.some(cell => cell !== "")) {
        let rowString = rowData.join();

        if (!destinationSet.has(rowString)) {
          let lastRow = destinationSheet.getLastRow();
          let nextRow = lastRow + 1;

          destinationSheet.getRange(nextRow, 1, 1, numCols).setValues([rowData]);
          rowsCopied++;

          destinationSet.add(rowString);
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert(rowsCopied + ' Candidate(s) shortlisted.');
}

function addDatePickers() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  let dataValues = dataRange.getValues();

  for (let i = 0; i < dataValues.length; i++) {
    let rowData = dataValues[i];
    let hasData = rowData.some(cell => cell !== '');

    if (hasData) {
      let dateCell = sheet.getRange(i + 2, 5); // Adjusted column index
      let dateRule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(false)
        .build();
      dateCell.setDataValidation(dateRule);
    }
  }
}

function addDropdownsToColumns() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  let statusOptions = ['In progress', 'Approved', 'Rejected'];
  let hourOptions = Array.from({ length: 24 }, (_, i) => (i < 10 ? '0' : '') + i.toString());
  let minuteOptions = ['00', '15', '30', '45'];
  let notificationOptions = ['-', 'Sent', 'Not sent'];
  let decisionOptions = ['-', 'Accepted', 'Rejected'];

  let colors = {
    'In progress': '#FFD700',
    'Approved': '#9ACD32',
    'Rejected': '#FF6347',
    'Accepted': '#9ACD32', // Green
    'Rejected': '#FF6347'  // Red
  };

  let statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions)
    .setAllowInvalid(false)
    .build();

  let hourRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(hourOptions)
    .setAllowInvalid(false)
    .build();

  let minuteRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(minuteOptions)
    .setAllowInvalid(false)
    .build();
    
  let notificationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(notificationOptions)
    .setAllowInvalid(false)
    .build();

  let decisionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(decisionOptions)
    .setAllowInvalid(false)
    .build();

  let dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  let dataValues = dataRange.getValues();

  for (let i = 0; i < dataValues.length; i++) {
    let rowData = dataValues[i];
    let hasData = rowData.slice(0, 5).concat(rowData.slice(8)).some(cell => cell !== ''); // Adjusted indices

    if (hasData) {
      let statusCell = sheet.getRange(i + 2, 8); // Adjusted column index
      statusCell.setDataValidation(statusRule);

      let cellValue = statusCell.getValue();
      if (!statusOptions.includes(cellValue)) {
        statusCell.setValue('In progress');
        cellValue = 'In progress';
      }

      statusCell.setBackground(colors[cellValue]);

      let hourCell = sheet.getRange(i + 2, 6); // Adjusted column index
      hourCell.setDataValidation(hourRule);
      hourCell.setNumberFormat('00');

      let minuteCell = sheet.getRange(i + 2, 7); // Adjusted column index
      minuteCell.setDataValidation(minuteRule);
      minuteCell.setNumberFormat('00');

      let notificationCell = sheet.getRange(i + 2, 9); // Adjusted column index
      notificationCell.setDataValidation(notificationRule);
      if (!notificationOptions.includes(notificationCell.getValue())) {
        notificationCell.setValue('-');
      }

      let decisionCell = sheet.getRange(i + 2, 10); // Adjusted column index
      decisionCell.setDataValidation(decisionRule);
      if (!decisionOptions.includes(decisionCell.getValue())) {
        decisionCell.setValue('-');
      }
    }
  }

  addConditionalFormatting(sheet, statusOptions, colors);
}

function addConditionalFormatting(sheet, statusOptions, colors) {
  let range = sheet.getRange(2, 8, sheet.getLastRow() - 1); // Adjusted column index
  let rules = [];

  statusOptions.forEach(option => {
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(option)
      .setBackground(colors[option])
      .setRanges([range])
      .build();
    rules.push(rule);
  });

  sheet.getRange(2, 8, sheet.getLastRow() - 1).clearFormat(); // Adjusted column index
  sheet.setConditionalFormatRules(sheet.getConditionalFormatRules().concat(rules));
}

function onEditHandler(e) {
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let row = range.getRow();
  let col = range.getColumn();

  if (col === 8 && row > 1) { // Column H (8th column)
    let cellValue = range.getValue();
    let notificationCell = sheet.getRange(row, 9); // Column I (9th column)
    let decisionCell = sheet.getRange(row, 10); // Column J (10th column)
    
    if (cellValue === 'Approved') {
      notificationCell.setValue('Not sent').setBackground(null);
      decisionCell.setBackground(null);
    } else {
      notificationCell.setValue('-').setBackground('#CCCCCC');
      decisionCell.setValue('-').setBackground('#CCCCCC');
    }
  }

  if (col === 9 && row > 1) { // Column I (9th column)
    let cellValue = range.getValue();
    if (!['-', 'Sent', 'Not sent'].includes(cellValue)) {
      range.setValue('-');
    }
  }

  if (col === 10 && row > 1) { // Column J (10th column)
    let cellValue = range.getValue();
    if (!['-', 'Accepted', 'Rejected'].includes(cellValue)) {
      range.setValue('-');
    } else {
      if (cellValue === 'Accepted') {
        range.setBackground('#9ACD32');
      } else if (cellValue === 'Rejected') {
        range.setBackground('#FF6347');
      } else {
        range.setBackground(null);
      }
    }
  }

  if (col === 6 || col === 7) { // Columns F (6th) and G (7th)
    let value = range.getValue();
    if (value.length === 1) {
      value = '0' + value;
    } else if (col === 7 && value !== '00' && value !== '15' && value !== '30' && value !== '45') {
      range.setValue('00');
      return;
    }
    range.setValue(value);
    range.setNumberFormat('00');
  }
}




function sendOfferLetters() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Shortlisted Candidates');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('The "Shortlisted Candidates" sheet does not exist.');
    return;
  }

  let selectedRange = sheet.getActiveRangeList();
  if (!selectedRange) {
    SpreadsheetApp.getUi().alert('Please select rows to send offer letters.');
    return;
  }

  let emailSentCount = 0;

  selectedRange.getRanges().forEach(range => {
    let rowStart = range.getRow();
    let rowEnd = rowStart + range.getNumRows() - 1;

    for (let row = rowStart; row <= rowEnd; row++) {
      let email = sheet.getRange(row, 3).getValue(); //  email is in the 3rd column (C)
      let status = sheet.getRange(row, 8).getValue(); //  status is in the 8th column (H)
      
      if (status === 'Approved') {
        let candidateName = sheet.getRange(row, 1).getValue(); //  name is in the 1st column (A)

        let subject = 'Job Offer from Handsome Boi';
        let body = `
          Dear ${candidateName},

          We are pleased to offer you the position you desired at our company. 

          Looking forward to your reply.

          Best regards,
          Handsome Boi
        `;

        MailApp.sendEmail(email, subject, body);
        sheet.getRange(row, 9).setValue('Sent'); // Set notification status to "Sent"
        emailSentCount++;
      }
    }
  });

  SpreadsheetApp.getUi().alert(emailSentCount + ' offer letter(s) sent.');
}


function employButton() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName('Shortlisted Candidates');
  const formResponsesSheet = spreadsheet.getSheetByName('Form Responses');
  const destinationSheet = spreadsheet.getSheetByName('Sheet1');

  if (!sourceSheet || !formResponsesSheet || !destinationSheet) {
    SpreadsheetApp.getUi().alert('One or more sheets do not exist.');
    return;
  }

  const selectedRange = sourceSheet.getActiveRangeList();
  if (!selectedRange) {
    SpreadsheetApp.getUi().alert('Please select rows to process.');
    return;
  }

  const columnsToCopy = [1, 2, 8, 9, 4, 5, 6]; // Columns B, C, I, J, E, F, G in Form Responses
  const destinationColumns = [2, 3, 4, 5, 6, 7, 8]; // Columns B, C, D, E, F, G, H in Sheet1
  const numCols = columnsToCopy.length;

  let rowsCopied = 0;
  const destinationData = destinationSheet.getRange(2, 2, destinationSheet.getLastRow() - 1, numCols).getValues();
  const destinationSet = new Set(destinationData.map(row => row.slice(1).join())); // Excluding the ID

  const ranges = selectedRange.getRanges();
  for (let range of ranges) {
    const rowStart = range.getRow();
    const rowEnd = rowStart + range.getNumRows() - 1;

    for (let row = rowStart; row <= rowEnd; row++) {
      const email = sourceSheet.getRange(row, 3).getValue(); // Column C in Shortlisted Candidates
      if (email) {
        const formResponsesData = formResponsesSheet.getDataRange().getValues();
        const formRowIndex = formResponsesData.findIndex(row => row[6] === email); // Column C in Form Responses

        if (formRowIndex !== -1) {
          const formRow = formResponsesData[formRowIndex];
          const valuesToCopy = columnsToCopy.map(col => formRow[col]);

          const rowString = valuesToCopy.join();
          if (!destinationSet.has(rowString)) {
            // Find the next available ID
            const lastRow = destinationSheet.getLastRow();
            const lastId = lastRow > 1 ? destinationSheet.getRange(lastRow, 1).getValue() : 0; // Get the last ID or 0 if no data
            const newId = lastId + 1;

            // Insert new data with the new ID
            destinationSheet.getRange(lastRow + 1, 1, 1, numCols + 1).setValues([[newId, ...valuesToCopy]]);
            rowsCopied++;
            destinationSet.add(rowString);
          }
        }
      }
    }
  }

  SpreadsheetApp.getUi().alert(rowsCopied + ' candidate(s) is employed.');
}





/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    // Enable Gmail service
var selectedEmployee = null;

function tableFormat() {
  // Get the active spreadsheet
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get the header
  let headers = sheet.getRange('A1:I1');

  // Get the table
  let table = sheet.getDataRange();

  // Set the characteristics of the head
  headers.setFontWeight('bold');
  headers.setFontColor('black');
  headers.setBackground('#00ffff');

  // Set the characteristics of the table
  table.setFontFamily('Roboto');
  table.setHorizontalAlignment('center');
  table.setBorder(true, true, true, true, false, true, '#00ffff', SpreadsheetApp.BorderStyle.SOLID);
}


// ADDED ABOVE IN LINE 6
// function onOpen() {
//   // Get the UI
//   let ui = SpreadsheetApp.getUi();

//   // Create a menu option at the top navigation bar in Google Sheet
//   ui.createMenu('Employee data formatting').addItem('Employee data', 'tableFormat').addToUi();
// }

function onChange(e) {
  var changeType = e.changeType;

  // Only proceed if the change type is an edit or other relevant change type
  if (changeType === 'EDIT' || changeType === 'INSERT_ROW' || changeType === 'INSERT_COLUMN' || changeType === 'REMOVE_ROW' || changeType === 'REMOVE_COLUMN') {
    handleEdit(e);
  }
}

function handleEdit(e) {
  var leaveRequestsSheetName = "Leave request status"; // The actual name of the leave requests sheet
  var leaveRequestsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(leaveRequestsSheetName);
  
  if (!leaveRequestsSheet) {
    Logger.log("Target sheet not found: " + leaveRequestsSheetName);
    return;
  }

  var leaveRequestsRange = leaveRequestsSheet.getActiveRange();
  var leaveRequestsRow = leaveRequestsRange.getRow();
  var leaveRequestsColumn = leaveRequestsRange.getColumn();
  
  var statusColumn = 6; // Change this to the column number of your "Status" column in the Leave request status sheet
  var employeeIdColumn = 2; // Change this to the column number of the employee ID in the Leave request status sheet

  Logger.log("Edited cell - Row: " + leaveRequestsRow + ", Column: " + leaveRequestsColumn);

  // Check if the edited column is the "Status" column
  if (leaveRequestsColumn === statusColumn) {
    var status = leaveRequestsSheet.getRange(leaveRequestsRow, leaveRequestsColumn).getValue();
    var employeeId = leaveRequestsSheet.getRange(leaveRequestsRow, employeeIdColumn).getValue();

    Logger.log("Status: " + status + ", Employee ID: " + employeeId);
    
    // Fetch employee info from EmployeeInfo sheet
    var employeeInfoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    if (!employeeInfoSheet) {
      Logger.log("EmployeeInfo sheet not found.");
      return;
    }
    
    var employeeData = getEmployeeData(employeeInfoSheet, employeeId);
    Logger.log("Employee Data: " + JSON.stringify(employeeData));
    
    if (employeeData) {
      var employeeName = employeeData.name;
      var employeeEmail = employeeData.email;
      var subject, body;
      
      if (status === 'Approved') {
        subject = 'Leave Request Approved';
        body = 'Dear ' + employeeName + ',\n\nYour leave request has been approved.\n\nBest regards,\nHR Department';
      } else if (status === 'Rejected') {
        subject = 'Leave Request Rejected';
        body = 'Dear ' + employeeName + ',\n\nYour leave request has been rejected.\n\nBest regards,\nHR Department';
      }
      
      Logger.log("Email subject: " + subject);
      Logger.log("Email body: " + body);
      Logger.log("Employee email: " + employeeEmail);
      
      if (subject && body) {
        sendEmail(employeeEmail, subject, body);
      } else {
        Logger.log("Email subject or body is missing.");
      }
    } else {
      Logger.log("Employee data not found for ID: " + employeeId);
    }
  }
}

function getEmployeeData(sheet, employeeId) {
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) { // Start from 1 to skip header row
    if (data[i][0] == employeeId) { // Assuming employee ID is in the first column
      return {
        id: data[i][0],
        name: data[i][1], // Assuming employee name is in the second column
        email: data[i][7] // Assuming employee email is in the eighth column
      };
    }
  }
  return null; // Return null if employee not found
}

function sendEmail(email, subject, body) {
  MailApp.sendEmail(email, subject, body);
  Logger.log("Email sent to: " + email);
}

function authorize() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Authorization Test", "This is a test email to authorize the script.");
}

function generateAndSendReport() {
  Logger.log("Report generation started...");

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leave request status');
    if (!sheet) {
      Logger.log("Sheet not found");
      return;
    }

    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();

    Logger.log("Data retrieved from sheet");

    var report = "<h1>Employee Leave Requests Report</h1>";
    report += "<table border='1'><tr><th>ID</th><th>Employee Name</th><th>Email</th><th>Leave Type</th><th>Start Date</th><th>End Date</th><th>Status</th></tr>";

    for (var i = 1; i < data.length; i++) {
      report += "<tr>";
      for (var j = 0; j < data[i].length; j++) {
        report += "<td>" + data[i][j] + "</td>";
      }
      report += "</tr>";
    }

    report += "</table>";

    var emailAddress = "czefenglim@gmail.com";  // Change this to the manager's email address
    var subject = "Employee Leave Requests Report";
    var body = "Please find the employee leave requests report below:<br><br>" + report;

    Logger.log("Sending email to: " + emailAddress);

    MailApp.sendEmail({
      to: emailAddress,
      subject: subject,
      htmlBody: body
    });

    Logger.log("Report sent successfully");

  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

function calculatePayroll() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Payroll');
  if (!sheet) {
    Logger.log("Sheet 'Payroll' not found");
    return;
  }

  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  Logger.log("Data range obtained: " + data.length + " rows");

  for (var i = 1; i < data.length; i++) {
    var hoursWorked = data[i][1];
    var hourlyRate = data[i][2];
    var deductions = data[i][3];
    var bonuses = data[i][4];

    var grossPay = hoursWorked * hourlyRate;
    var netPay = grossPay - deductions + bonuses;

    Logger.log("Row " + (i + 1) + ": Hours Worked=" + hoursWorked + ", Hourly Rate=" + hourlyRate + ", Deductions=" + deductions + ", Bonuses=" + bonuses + ", Net Pay=" + netPay);

    sheet.getRange(i + 1, 7).setValue(netPay); 
  }

  Logger.log("Payroll calculated successfully");
}

function copyEmployeesToScans() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Sheet1');
  var targetSheet = ss.getSheetByName('Scans');

  // Get data from source sheet
  var sourceData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 1).getValues();
  // Get the last row of the target sheet
  var targetLastRow = targetSheet.getLastRow();

  // Prepare new data array
  var newData = [];
  for (var i = 0; i < sourceData.length; i++) {
    var scanID = generateScanID();
    var id = sourceData[i][0]; 
    newData.push([scanID, id, '', '']); // Columns: Scan_ID, ID, DateTime, Status
  }

  // Append new data to target sheet
  targetSheet.getRange(targetLastRow + 1, 1, newData.length, newData[0].length).setValues(newData);
}

function generateScanID() {
  var characters = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var scanID = '';
  for (var i = 0; i < 8; i++) {
    var randomIndex = Math.floor(Math.random() * characters.length);
    scanID += characters.charAt(randomIndex);
  }
  return scanID;
}
