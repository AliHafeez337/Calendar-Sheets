function createSimpleCalendar() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentYear = 2026;
  
  // Check if "Calendar 2026" sheet exists, if not create it
  var sheet = spreadsheet.getSheetByName('Calendar 2026');
  if (sheet) {
    // Clear existing sheet
    sheet.clear();
  } else {
    // Create new sheet
    sheet = spreadsheet.insertSheet('Calendar 2026');
  }
  
  // Create first page (months 1-6) starting from January 2026
  createMonthGrid(sheet, 2026, 0); // 0 = January
  
  // Create second sheet for months 7-12
  var sheet2 = spreadsheet.getSheetByName('Calendar 2026 (Jul-Dec)');
  if (sheet2) {
    sheet2.clear();
  } else {
    sheet2 = spreadsheet.insertSheet('Calendar 2026 (Jul-Dec)');
  }
  
  createMonthGrid(sheet2, 2026, 6); // 6 = July
  
  // Format both sheets
  formatCalendar(sheet);
  formatCalendar(sheet2);
  
  // Setup print formatting
  setupPrintFormatting(sheet);
  setupPrintFormatting(sheet2);
  
  // Set active sheet to first page
  sheet.activate();
  
  // Show success message
  var ui = SpreadsheetApp.getUi();
  ui.alert('✅ Calendar 2026 Created!', 
    'Your calendar has been created in:\n' +
    '• "Calendar 2026" sheet (Jan-Jun)\n' +
    '• "Calendar 2026 (Jul-Dec)" sheet (Jul-Dec)\n\n' +
    'Layout: 2 rows × 3 columns\n' +
    '• 2 column spaces between months\n' +
    '• Empty rows between each week\n' +
    '• Font size 14\n' +
    '• No borders - clean and simple!\n' +
    'Ready to print on 2 A4 pages!', 
    ui.ButtonSet.OK);
}

function setupPrintFormatting(sheet) {
  // Set column widths
  for (var i = 1; i <= 27; i++) {
    sheet.setColumnWidth(i, 45);
  }
  
  // Set row heights
  for (var i = 1; i <= 40; i++) {
    sheet.setRowHeight(i, 18);
  }
}

function createMonthGrid(sheet, year, startMonth) {
  var months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  
  var days = ['S', 'M', 'T', 'W', 'T', 'F', 'S'];
  var monthCount = 0;
  
  // Clear the sheet first
  sheet.clear();
  
  // Create 6 months in a 2x3 grid (2 rows, 3 columns)
  for (var row = 0; row < 2; row++) {
    for (var col = 0; col < 3; col++) {
      if (monthCount >= 6) break;
      
      var monthIndex = startMonth + monthCount;
      var monthName = months[monthIndex];
      
      // Calculate position for this month
      // Each month takes: title row + day headers row + (6 weeks * 2 rows each) = 1 + 1 + 12 = 14 rows
      var monthStartRow = 1 + (row * 16); // 16 rows per month block (14 rows calendar + 2 empty rows between months)
      var monthStartCol = 1 + (col * 9);   // 9 columns per month (7 days + 2 spacer columns)
      
      // Add month title
      var titleCell = sheet.getRange(monthStartRow, monthStartCol);
      titleCell.setValue(monthName + ' ' + year);
      titleCell.setFontWeight('bold');
      titleCell.setFontSize(14);
      
      // Merge title cells (spanning 7 days)
      var titleRange = sheet.getRange(monthStartRow, monthStartCol, 1, 7);
      titleRange.merge();
      titleRange.setHorizontalAlignment('center');
      titleRange.setBackground('#e6f2ff');
      titleRange.setFontWeight('bold');
      titleRange.setFontSize(14);
      titleRange.setBorder(false, false, false, false, false, false);
      
      // Add day headers (row after title)
      for (var d = 0; d < 7; d++) {
        var dayCell = sheet.getRange(monthStartRow + 1, monthStartCol + d);
        dayCell.setValue(days[d]);
        dayCell.setFontWeight('bold');
        dayCell.setHorizontalAlignment('center');
        dayCell.setBackground('#f0f0f0');
        dayCell.setFontSize(14);
        dayCell.setBorder(false, false, false, false, false, false);
      }
      
      // Add empty row after day headers
      var emptyRow1 = monthStartRow + 2;
      for (var emptyCol = monthStartCol; emptyCol < monthStartCol + 7; emptyCol++) {
        var emptyCell = sheet.getRange(emptyRow1, emptyCol);
        emptyCell.setValue('');
        emptyCell.setBackground('white');
        emptyCell.setBorder(false, false, false, false, false, false);
      }
      
      // Fill in the days with empty rows between weeks
      var firstDay = new Date(year, monthIndex, 1);
      var startingDay = firstDay.getDay(); // 0 = Sunday
      var daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
      
      var dayCounter = 1;
      for (var week = 0; week < 6; week++) {
        // Week row
        var weekRow = monthStartRow + 3 + (week * 2);
        
        for (var day = 0; day < 7; day++) {
          var cellRow = weekRow;
          var cellCol = monthStartCol + day;
          var cell = sheet.getRange(cellRow, cellCol);
          
          if ((week === 0 && day < startingDay) || dayCounter > daysInMonth) {
            // Empty cell
            cell.setValue('');
            cell.setBackground('white');
          } else {
            // Fill with day number
            cell.setValue(dayCounter);
            cell.setBackground('white');
            dayCounter++;
          }
          
          cell.setHorizontalAlignment('center');
          cell.setFontSize(14);
          cell.setVerticalAlignment('middle');
          cell.setBorder(false, false, false, false, false, false);
        }
        
        // Add empty row after each week (except last week)
        if (week < 5) {
          var emptyWeekRow = weekRow + 1;
          for (var emptyCol = monthStartCol; emptyCol < monthStartCol + 7; emptyCol++) {
            var emptyCell = sheet.getRange(emptyWeekRow, emptyCol);
            emptyCell.setValue('');
            emptyCell.setBackground('white');
            emptyCell.setBorder(false, false, false, false, false, false);
          }
        }
      }
      
      // Add two spacer columns after each month (except last column)
      if (col < 2) {
        // First spacer column
        var spacerCol1 = monthStartCol + 7;
        var spacerRange1 = sheet.getRange(monthStartRow, spacerCol1, 15, 1);
        spacerRange1.setValue('');
        spacerRange1.setBackground('white');
        spacerRange1.setBorder(false, false, false, false, false, false);
        
        // Second spacer column
        var spacerCol2 = monthStartCol + 8;
        var spacerRange2 = sheet.getRange(monthStartRow, spacerCol2, 15, 1);
        spacerRange2.setValue('');
        spacerRange2.setBackground('white');
        spacerRange2.setBorder(false, false, false, false, false, false);
      }
      
      monthCount++;
    }
    
    // Add two empty rows between month rows
    if (row === 0) {
      var emptyRowStart = 16; // After first month block (14 rows + 2)
      for (var r = 0; r < 2; r++) {
        var emptyRow = emptyRowStart + r;
        for (var emptyCol = 1; emptyCol <= 27; emptyCol++) {
          var emptyCell = sheet.getRange(emptyRow, emptyCol);
          emptyCell.setValue('');
          emptyCell.setBackground('white');
          emptyCell.setBorder(false, false, false, false, false, false);
        }
        sheet.setRowHeight(emptyRow, 12);
      }
    }
  }
}

function formatCalendar(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  // Limit to columns we care about (max 27)
  if (lastCol > 27) lastCol = 27;
  
  if (lastRow > 0 && lastCol > 0) {
    var allRange = sheet.getRange(1, 1, lastRow, lastCol);
    allRange.setFontSize(14);
    allRange.setBorder(false, false, false, false, false, false);
  }
  
  // Adjust row heights
  for (var row = 1; row <= lastRow; row++) {
    // Title rows (1, 17, etc.)
    if (row === 1 || row === 17) {
      sheet.setRowHeight(row, 26);
    }
    // Day header rows (2, 18, etc.)
    else if (row === 2 || row === 18) {
      sheet.setRowHeight(row, 22);
    }
    // Empty row after day headers (3, 19, etc.)
    else if (row === 3 || row === 19) {
      sheet.setRowHeight(row, 10);
    }
    // Week rows (4,6,8,10,12,14 and 20,22,24,26,28,30)
    else if ((row >= 4 && row <= 14 && row % 2 === 0) || 
             (row >= 20 && row <= 30 && row % 2 === 0)) {
      sheet.setRowHeight(row, 24);
    }
    // Empty rows between weeks (5,7,9,11,13 and 21,23,25,27,29)
    else if ((row >= 5 && row <= 13 && row % 2 === 1) || 
             (row >= 21 && row <= 29 && row % 2 === 1)) {
      sheet.setRowHeight(row, 8);
    }
    // Empty rows between month rows (15-16)
    else if (row === 15 || row === 16) {
      sheet.setRowHeight(row, 12);
    }
    // Any other rows
    else {
      sheet.setRowHeight(row, 15);
    }
  }
  
  // Adjust column widths
  for (var col = 1; col <= 27; col++) {
    // Check if this is a spacer column (columns 8-9, 17-18, 26-27)
    if (col % 9 === 8 || col % 9 === 0) {
      sheet.setColumnWidth(col, 10); // Narrow spacer columns
    } else {
      sheet.setColumnWidth(col, 50); // Wider for day columns
    }
  }
}

// Add menu item
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 Calendar Tools')
    .addItem('Create Calendar 2026', 'createSimpleCalendar')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

function showAbout() {
  var ui = SpreadsheetApp.getUi();
  ui.alert('📅 Calendar 2026',
    'This tool creates a simple 2026 calendar:\n\n' +
    '• Page 1: January - June 2026\n' +
    '• Page 2: July - December 2026\n' +
    '• Layout: 2 rows × 3 columns\n' +
    '• 2 column spaces between months\n' +
    '• EMPTY ROWS BETWEEN EACH WEEK\n' +
    '• Font size 14\n' +
    '• NO BORDERS - clean and minimal\n' +
    '• Month names in bold with light blue background\n' +
    '• Day initials in bold with gray background\n' +
    '• Plain day numbers\n' +
    '• Optimized for A4 printing\n\n' +
    'Created by: Simple Calendar Generator',
    ui.ButtonSet.OK);
}