function createTaskSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheetName = "MARCH SUMMARY";
  
  var summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(summarySheetName);
  } else {
    summarySheet.clearContents(); 
  }
  
  var dataSheets = ss.getSheets().filter(function(sheet) {
    return sheet.getName() !== summarySheetName;
  });
  
  var totalsByDateAndMember = {};
  var allMembers = {};
  var allDates = new Set();
  
  dataSheets.forEach(function(sheet) {
    var dateName = sheet.getName();  
    allDates.add(dateName);
    
    if (!totalsByDateAndMember[dateName]) {
      totalsByDateAndMember[dateName] = {};
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Start at row index 2 (third row) to ignore constant title rows.
    for (var rowIndex = 2; rowIndex < data.length; rowIndex++) { 
      var row = data[rowIndex];
      
      var member1 = row[2];  
      var member2 = row[3];  
      var total   = row[4];  
      
      if (member1) {
        if (!totalsByDateAndMember[dateName][member1]) {
          totalsByDateAndMember[dateName][member1] = 0;
        }
        totalsByDateAndMember[dateName][member1] += total;
        allMembers[member1] = true;
      }
      
      if (member2) {
        if (!totalsByDateAndMember[dateName][member2]) {
          totalsByDateAndMember[dateName][member2] = 0;
        }
        totalsByDateAndMember[dateName][member2] += total;
        allMembers[member2] = true;
      }
    }
  });
  
  var memberList = Object.keys(allMembers).sort();
  var dateList = Array.from(allDates).sort();
  
  summarySheet.getRange(1, 1).setValue("Members");
  for (var col = 0; col < dateList.length; col++) {
    summarySheet.getRange(1, col + 2).setValue(dateList[col]);
  }
  
  for (var i = 0; i < memberList.length; i++) {
    var memberName = memberList[i];
    var rowIndex = i + 2; 
    summarySheet.getRange(rowIndex, 1).setValue(memberName);
    
    for (var col = 0; col < dateList.length; col++) {
      var dateName = dateList[col];
      var total = 0;
      if (totalsByDateAndMember[dateName] && totalsByDateAndMember[dateName][memberName]) {
        total = totalsByDateAndMember[dateName][memberName];
      }
      summarySheet.getRange(rowIndex, col + 2).setValue(total);
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Summarize Now")
    .addItem("Run Summary", "createTaskSummary")
    .addToUi();
}
