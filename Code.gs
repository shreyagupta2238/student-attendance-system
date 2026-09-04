function doGet() {
  return HtmlService.createTemplateFromFile('listofstudent')
    .evaluate()
    .setTitle('Student Attendance System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Fetch all students dynamically from Sheet1
function getStudentListFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var students = [];

  // Starts from Row 2 (index 1) to skip header row
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) { // Checks if Roll No exists in Col B
      students.push({
        sNo: data[i][0],                      // Col A: S.NO
        rollNo: data[i][1].toString().trim(), // Col B: ROLL NO
        name: data[i][2]                       // Col C: STUDENT NAME
      });
    }
  }
  return students;
}

// Fetch available subjects
function getAvailableSubjects() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Teacher_Auth");
  if (!sheet) {
    sheet = ss.insertSheet("Teacher_Auth");
    sheet.appendRow(["Subject", "PIN"]);
    sheet.appendRow(["ADMIN", "9999"]);
  }

  var data = sheet.getDataRange().getValues();
  var subjects = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0] !== "ADMIN") {
      subjects.push(data[i][0].toString().trim());
    }
  }
  return subjects;
}

// Verify Teacher PIN
function verifyTeacherPin(subject, enteredPin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Teacher_Auth");
  if (!sheet || !enteredPin) return false;

  var data = sheet.getDataRange().getValues();
  var adminPin = "";

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "ADMIN") adminPin = data[i][1].toString().trim();
    if (data[i][0] === subject && data[i][1].toString().trim() === enteredPin.toString().trim()) {
      return true;
    }
  }
  return enteredPin.toString().trim() === adminPin;
}

// Allow teachers to add new subjects and set PINs
function registerOrUpdateSubject(subject, newPin, masterPin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Teacher_Auth");
  if (!sheet) {
    sheet = ss.insertSheet("Teacher_Auth");
    sheet.appendRow(["Subject", "PIN"]);
    sheet.appendRow(["ADMIN", "9999"]);
  }

  var data = sheet.getDataRange().getValues();
  var isValidMaster = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "ADMIN" && data[i][1].toString().trim() === masterPin.toString().trim()) {
      isValidMaster = true;
      break;
    }
  }

  if (!isValidMaster) {
    return { success: false, message: "Invalid Master Authorization PIN." };
  }

  for (var j = 1; j < data.length; j++) {
    if (data[j][0].toString().toUpperCase() === subject.toString().toUpperCase()) {
      sheet.getRange(j + 1, 2).setValue(newPin);
      return { success: true, message: "PIN updated for " + subject };
    }
  }

  sheet.appendRow([subject.toUpperCase(), newPin]);
  return { success: true, message: "New Subject '" + subject.toUpperCase() + "' created successfully!" };
}

// Fetch Attendance Data matching your exact sheet columns
function getAttendanceData(subject, date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  if (!sheet) {
    sheet = ss.insertSheet("Attendance");
    sheet.appendRow(["S.NO", "Roll No.", "Student Name", "Status", "Subject", "Date", "Logged Timestamp"]);
  }

  var data = sheet.getDataRange().getValues();
  var records = {};

  for (var i = 1; i < data.length; i++) {
    var rowSub = data[i][4]; // Subject is Col E
    var rowDate = formatDate(data[i][5]); // Date is Col F
    
    if (rowSub === subject && rowDate === date) {
      records[data[i][1].toString()] = { // Roll No is Col B
        status: data[i][3], // Status is Col D
        timestamp: data[i][6] ? formatDateString(data[i][6]) : "" // Timestamp is Col G
      };
    }
  }

  return { records: records };
}

// Save or Update Attendance (Teacher Action)
function saveTeacherAttendance(subject, date, rollNo, studentName, status, pin) {
  if (!verifyTeacherPin(subject, pin)) {
    return { success: false, message: "Unauthorized. Valid Teacher PIN required." };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  if (!sheet) {
    sheet = ss.insertSheet("Attendance");
    sheet.appendRow(["S.NO", "Roll No.", "Student Name", "Status", "Subject", "Date", "Logged Timestamp"]);
  }

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowRoll = data[i][1].toString(); // Col B
    var rowSub = data[i][4];             // Col E
    var rowDate = formatDate(data[i][5]); // Col F

    if (rowSub === subject && rowDate === date && rowRoll === rollNo.toString()) {
      sheet.getRange(i + 1, 4).setValue(status); // Update Status in Col D
      sheet.getRange(i + 1, 7).setValue(new Date()); // Update Timestamp in Col G
      return { success: true };
    }
  }

  // Append new row in exact column order
  var newSNo = data.length; 
  sheet.appendRow([newSNo, rollNo, studentName, status, subject, date, new Date()]);
  return { success: true };
}

// Delete Entire Day's Attendance for Subject
function deleteAllDataForDate(subject, date, pin) {
  if (!verifyTeacherPin(subject, pin)) {
    return { success: false, message: "Unauthorized. Valid Teacher PIN required." };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  if (!sheet) return { success: false, message: "Attendance sheet not found." };

  var data = sheet.getDataRange().getValues();
  var count = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    var rowSub = data[i][4]; // Col E
    var rowDate = formatDate(data[i][5]); // Col F
    
    if (rowSub === subject && rowDate === date) {
      sheet.deleteRow(i + 1);
      count++;
    }
  }

  return { success: true, message: "Deleted " + count + " record(s) for " + subject + " on " + date };
}

// Delete Single Student Record
function deleteSingleStudentForDate(subject, date, rollNo, pin) {
  if (!verifyTeacherPin(subject, pin)) {
    return { success: false, message: "Unauthorized. Valid Teacher PIN required." };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  if (!sheet) return { success: false, message: "Attendance sheet not found." };

  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    var rowRoll = data[i][1].toString(); // Col B
    var rowSub = data[i][4];             // Col E
    var rowDate = formatDate(data[i][5]); // Col F
    
    if (rowSub === subject && rowDate === date && rowRoll === rollNo.toString()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Removed record for Roll No: " + rollNo };
    }
  }

  return { success: false, message: "Record not found for Roll No: " + rollNo };
}

// Helpers
function formatDate(dateObj) {
  if (!dateObj) return "";
  return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatDateString(dateObj) {
  if (!dateObj) return "";
  return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}