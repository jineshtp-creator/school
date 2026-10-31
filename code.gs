function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Student Marks Entry')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheets() {
  var files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
  var sheetList = [];
  var count = 0;
  
  while (files.hasNext() && count < 100) {
    var file = files.next();
    sheetList.push({ id: file.getId(), name: file.getName() });
    count++;
  }
  return sheetList;
}

function submitMarks(formObject) {
  var fileId = formObject.spreadsheetId;
  var sheetName = formObject.className;
  var subject = formObject.subject;
  var marksText = formObject.marks;
  
  if (!fileId) return "Error: Please select a Spreadsheet.";
  
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return "Error: Class sheet '" + sheetName + "' not found in the selected file.";
  
  var marksArray = marksText.split('\n');
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  var subjectColIndex = -1;
  
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim().toUpperCase() === subject.toUpperCase()) {
      subjectColIndex = i + 1;
      break;
    }
  }
  
  if (subjectColIndex === -1) return "Error: Subject '" + subject + "' not found.";
  
  var startRow = 3;
  var numRows = marksArray.length;
  var values = [];
  
  for (var i = 0; i < numRows; i++) {
    values.push([marksArray[i].trim()]);
  }
  
  sheet.getRange(startRow, subjectColIndex, numRows, 1).setValues(values);
  return "Success: Marks successfully submitted to " + sheetName + " for " + subject + "!";
}

function updateCellMark(fileId, sheetName, rowIndex, colIndex, newMark) {
  if (!fileId) return "Error: Please select a Spreadsheet.";
  
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "Error: Class sheet not found.";

  sheet.getRange(rowIndex, colIndex).setValue(newMark.trim());
  return "Success: Mark updated!";
}

function getTopRanks(fileId, sheetName) {
  if (!fileId) return { error: "Please select a Spreadsheet." };
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return { error: "Class sheet not found in the selected file." };
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 3) return { error: "Not enough data to calculate ranks." };
  
  var headers = data[1];
  var nameColIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().toUpperCase().indexOf("NAME") !== -1) {
      nameColIndex = i; break;
    }
  }
  if (nameColIndex === -1) nameColIndex = 3;
  
  var students = [];
  for (var r = 2; r < data.length; r++) {
    var row = data[r];
    var name = row[nameColIndex];
    if (!name || name.toString().trim() === "") continue;
    
    var total = 0;
    var hasMarks = false;
    for (var c = nameColIndex + 1; c < row.length; c++) {
      if (row[c] !== "") {
        var val = parseFloat(row[c]);
        if (!isNaN(val)) { total += val; hasMarks = true; }
      }
    }
    if (hasMarks) students.push({ name: name, total: total });
  }
  
  students.sort(function(a, b) { return b.total - a.total; });
  
  var rankedStudents = [];
  var currentRank = 1;
  var previousTotal = -1;
  for (var i = 0; i < students.length; i++) {
    if (previousTotal !== -1 && students[i].total < previousTotal) currentRank++;
    if (currentRank > 3) break;
    rankedStudents.push({ rank: currentRank, name: students[i].name, total: students[i].total });
    previousTotal = students[i].total;
  }
  return { success: true, data: rankedStudents };
}

function getAllStudentsMarks(fileId, sheetName) {
  if (!fileId) return { error: "Please select a Spreadsheet." };
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return { error: "Class sheet not found in the selected file." };
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 3) return { error: "Not enough data to display." };
  
  var headers = data[1]; // Row 2 contains headers
  
  // Locate Falls No, Roll No, and Name Column Indices dynamically
  var fallsColIndex = -1;
  var rollColIndex = -1;
  var nameColIndex = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var headerStr = headers[i].toString().toUpperCase().trim();
    if (headerStr.indexOf("FALL") !== -1 || headerStr.indexOf("FALLS") !== -1) {
      if (fallsColIndex === -1) fallsColIndex = i;
    }
    if (headerStr.indexOf("ROLL") !== -1 || headerStr.indexOf("ADM") !== -1) {
      if (rollColIndex === -1) rollColIndex = i;
    }
    if (headerStr.indexOf("NAME") !== -1) {
      if (nameColIndex === -1) nameColIndex = i;
    }
  }
  
  // Fallback defaults if headers aren't explicitly named
  if (fallsColIndex === -1 && rollColIndex !== -1) fallsColIndex = rollColIndex;
  else if (fallsColIndex === -1) fallsColIndex = 0; // Default to 1st column
  
  if (rollColIndex === -1) rollColIndex = 0;
  if (nameColIndex === -1) nameColIndex = 3; // Default to 4th column
  
  // Identify Subject Columns (columns after Name that have a non-empty header)
  var subjectHeaders = [];
  for (var c = nameColIndex + 1; c < headers.length; c++) {
    if (headers[c].toString().trim() !== "") {
      subjectHeaders.push({ name: headers[c], index: c + 1 });
    }
  }
  
  var students = [];
  for (var r = 2; r < data.length; r++) {
    var row = data[r];
    var fallsNo = row[fallsColIndex];
    var rollNo = row[rollColIndex];
    var name = row[nameColIndex];
    
    // Skip row if student name is empty
    if (!name || name.toString().trim() === "") continue;
    
    var studentData = { 
      fallsNo: fallsNo !== undefined && fallsNo !== null ? fallsNo : "",
      rollNo: rollNo !== undefined && rollNo !== null ? rollNo : "", 
      name: name, 
      marks: [], 
      total: 0, 
      hasMarks: false 
    };

    for (var i = 0; i < subjectHeaders.length; i++) {
      var colIdx = subjectHeaders[i].index;
      var mark = row[colIdx - 1];
      
      studentData.marks.push({
        value: mark,
        rowIndex: r + 1,
        colIndex: colIdx
      });

      if (mark !== "" && mark !== null && mark !== undefined) {
        var val = parseFloat(mark);
        if (!isNaN(val)) { 
          studentData.total += val; 
          studentData.hasMarks = true; 
        }
      }
    }
    
    if (studentData.hasMarks) {
      students.push(studentData);
    }
  }
  
  // Sort students in ASCENDING order of Falls No. / Roll No.
  //students.sort(function(a, b) {
  //  var numA = parseFloat(a.fallsNo || a.rollNo);
   // var numB = parseFloat(b.fallsNo || b.rollNo);
    
   // if (!isNaN(numA) && !isNaN(numB)) {
   //   return numA - numB;
   // }
   // return (a.fallsNo || a.rollNo).toString().localeCompare((b.fallsNo || b.rollNo).toString(), undefined, { numeric: true, sensitivity: 'base' });
  //});

  students.sort(function(a, b) { return b.total - a.total; });
  
  // Construct Table Headers
  var headerRow = ["Falls No.", "Roll No.", "Student Name"];
  for (var i = 0; i < subjectHeaders.length; i++) {
    headerRow.push(subjectHeaders[i].name);
  }
  headerRow.push("Total");
  
  return { success: true, headers: headerRow, students: students };
}

function getAllStudentsMarkso(fileId, sheetName) {
  if (!fileId) return { error: "Please select a Spreadsheet." };
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return { error: "Class sheet not found in the selected file." };
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 3) return { error: "Not enough data to display." };
  
  var headers = data[1];
  var nameColIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().toUpperCase().indexOf("NAME") !== -1) {
      nameColIndex = i; break;
    }
  }
  if (nameColIndex === -1) nameColIndex = 3;
  
  var subjectHeaders = [];
  for (var c = nameColIndex + 1; c < headers.length; c++) {
    if (headers[c].toString().trim() !== "") subjectHeaders.push({ name: headers[c], index: c + 1 });
  }
  
  var students = [];
  for (var r = 2; r < data.length; r++) {
    var row = data[r];
    var name = row[nameColIndex];
    if (!name || name.toString().trim() === "") continue;
    
    var studentData = { name: name, marks: [], total: 0, hasMarks: false };
    for (var i = 0; i < subjectHeaders.length; i++) {
      var colIdx = subjectHeaders[i].index;
      var mark = row[colIdx - 1];
      
      studentData.marks.push({
        value: mark,
        rowIndex: r + 1,
        colIndex: colIdx
      });

      if (mark !== "") {
        var val = parseFloat(mark);
        if (!isNaN(val)) { studentData.total += val; studentData.hasMarks = true; }
      }
    }
    if (studentData.hasMarks) students.push(studentData);
  }
  
  students.sort(function(a, b) { return b.total - a.total; });
  
  var headerRow = ["Student Name"];
  for (var i = 0; i < subjectHeaders.length; i++) headerRow.push(subjectHeaders[i].name);
  headerRow.push("Total");
  
  return { success: true, headers: headerRow, students: students };
}

function getAllStudentsMarks_2(fileId, sheetName) {
  if (!fileId) return { error: "Please select a Spreadsheet." };
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return { error: "Class sheet not found in the selected file." };
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 3) return { error: "Not enough data to display." };
  
  var headers = data[1]; // Row 2 contains headers
  
  // Locate Roll No and Name Column Indices dynamically
  var rollColIndex = -1;
  var nameColIndex = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var headerStr = headers[i].toString().toUpperCase().trim();
    if (headerStr.indexOf("ROLL") !== -1 || headerStr.indexOf("ADM") !== -1) {
      if (rollColIndex === -1) rollColIndex = i;
    }
    if (headerStr.indexOf("NAME") !== -1) {
      if (nameColIndex === -1) nameColIndex = i;
    }
  }
  
  // Fallback defaults if headers aren't explicitly named
  if (rollColIndex === -1) rollColIndex = 0; // Default to 1st column
  if (nameColIndex === -1) nameColIndex = 3; // Default to 4th column
  
  // Identify Subject Columns (Any column after Name that has a non-empty header)
  var subjectHeaders = [];
  for (var c = nameColIndex + 1; c < headers.length; c++) {
    if (headers[c].toString().trim() !== "") {
      subjectHeaders.push({ name: headers[c], index: c + 1 });
    }
  }
  
  var students = [];
  for (var r = 2; r < data.length; r++) {
    var row = data[r];
    var rollNo = row[rollColIndex];
    var name = row[nameColIndex];
    
    // Skip row if student name is empty
    if (!name || name.toString().trim() === "") continue;
    
    var studentData = { 
      rollNo: rollNo !== undefined && rollNo !== null ? rollNo : "", 
      name: name, 
      marks: [], 
      total: 0, 
      hasMarks: false 
    };

    for (var i = 0; i < subjectHeaders.length; i++) {
      var colIdx = subjectHeaders[i].index;
      var mark = row[colIdx - 1];
      
      studentData.marks.push({
        value: mark,
        rowIndex: r + 1,
        colIndex: colIdx
      });

      if (mark !== "" && mark !== null && mark !== undefined) {
        var val = parseFloat(mark);
        if (!isNaN(val)) { 
          studentData.total += val; 
          studentData.hasMarks = true; 
        }
      }
    }
    
    if (studentData.hasMarks) {
      students.push(studentData);
    }
  }
  
  // Sort students in ASCENDING order of Roll No.
  students.sort(function(a, b) {
    var numA = parseFloat(a.rollNo);
    var numB = parseFloat(b.rollNo);
    
    // If both are numbers, compare numerically
    if (!isNaN(numA) && !isNaN(numB)) {
      return numA - numB;
    }
    // Fallback natural alphanumeric sorting
    return a.rollNo.toString().localeCompare(b.rollNo.toString(), undefined, { numeric: true, sensitivity: 'base' });
  });
  
  // Construct Table Headers
  var headerRow = ["Roll No.", "Student Name"];
  for (var i = 0; i < subjectHeaders.length; i++) {
    headerRow.push(subjectHeaders[i].name);
  }
  headerRow.push("Total");
  
  return { success: true, headers: headerRow, students: students };
}
// NEW FUNCTION: Categorizes students into percentage bands for a specific subject
function getSubjectPerformance(fileId, sheetName, subject, maxMarks) {
  if (!fileId) return { error: "Please select a Spreadsheet." };
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return { error: "Class sheet not found in the selected file." };
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 3) return { error: "Not enough data to analyze." };
  
  var headers = data[1];
  var nameColIndex = -1;
  var subjectColIndex = -1;

  for (var i = 0; i < headers.length; i++) {
    var headerStr = headers[i].toString().trim().toUpperCase();
    if (headerStr.indexOf("NAME") !== -1 && nameColIndex === -1) {
      nameColIndex = i;
    }
    if (headerStr === subject.toUpperCase()) {
      subjectColIndex = i;
    }
  }

  if (nameColIndex === -1) nameColIndex = 3;
  if (subjectColIndex === -1) return { error: "Subject '" + subject + "' not found in the sheet headers." };

  var low = [];   // <= 30%
  var mid = [];   // 31% - 89%
  var high = [];  // >= 90%

  for (var r = 2; r < data.length; r++) {
    var row = data[r];
    var name = row[nameColIndex];
    var markVal = row[subjectColIndex];

    if (!name || name.toString().trim() === "" || markVal === "" || markVal === null || markVal === undefined) {
      continue;
    }

    var score = parseFloat(markVal);
    if (isNaN(score)) continue;

    var pct = Math.round((score / maxMarks) * 100 * 10) / 10; // Rounded to 1 decimal

    var studentInfo = { name: name, mark: score, percentage: pct };

    if (pct <= 30) {
      low.push(studentInfo);
    } else if (pct >= 90) {
      high.push(studentInfo);
    } else {
      mid.push(studentInfo);
    }
  }

  // Sort each group descending by percentage
  var sortDesc = function(a, b) { return b.percentage - a.percentage; };
  low.sort(sortDesc);
  mid.sort(sortDesc);
  high.sort(sortDesc);

  return { success: true, low: low, mid: mid, high: high };
}

