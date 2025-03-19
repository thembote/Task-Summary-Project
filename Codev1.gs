function createTaskSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheetName = "MARCH SUMMARY";
  
  var summarySheet = ss.getSheetByName(summarySheetName) || ss.insertSheet(summarySheetName);
  summarySheet.clear();
  
  var dataSheets = ss.getSheets().filter(sheet => sheet.getName() !== summarySheetName);
  var totalsByDateAndMember = {};
  var allMembers = new Set();
  var allDates = new Set();
  
  dataSheets.forEach(sheet => {
    var dateName = sheet.getName();
    allDates.add(dateName);
    totalsByDateAndMember[dateName] = totalsByDateAndMember[dateName] || {};
    
    var data = sheet.getDataRange().getValues().slice(2); // Skip header rows
    
    data.forEach(row => {
      var [,, member1, member2, total] = row;
      total = parseFloat(total) || 0;
      
      [member1, member2].forEach(member => {
        if (member) {
          member = member.toString().trim();
          totalsByDateAndMember[dateName][member] = (totalsByDateAndMember[dateName][member] || 0) + total;
          allMembers.add(member);
        }
      });
    });
  });
  
  var memberList = [...allMembers].sort();
  var dateList = [...allDates].sort((a, b) => new Date(a) - new Date(b));
  
  var lastColumn = dateList.length + 1;
  summarySheet.getRange(1, 1, 1, lastColumn).setValues([
    ["Members", ...dateList]
  ]).setFontWeight("bold").setFontColor("white").setBackground("#356854");
  
  var output = memberList.map(member => [
    member,
    ...dateList.map(date => totalsByDateAndMember[date]?.[member] || 0)
  ]);
  
  if (output.length > 0) {
    summarySheet.getRange(2, 1, output.length, lastColumn).setValues(output);
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Summarize Now")
    .addItem("Run Summary", "createTaskSummary")
    .addToUi();
}
