# 14vimal2.Ek-pahal-tech
//You can see how the code works and use it in your way and we will love to get suggestions and feedback to improve this
/**
 * The event handler triggered when installing the add-on.
 * @param {Event} e The onInstall event.
 */
function onInstall(e) {
  onOpen(e);
  setProfile();
  setAreas();
}
var timeArr = ['12:00 AM', '1:00 AM', '2:00 AM', '3:00 AM', '4:00 AM', '5:00 AM', '6:00 AM', '7:00 AM', '8:00 AM', '9:00 AM', '10:00 AM', '11:00 AM', '12:00 PM', '1:00 PM', '2:00 PM', '3:00 PM', '4:00 PM', '5:00 PM', '6:00 PM', '7:00 PM', '8:00 PM', '9:00 PM', '10:00 PM', '11:00 AM']
var timeIndex = 5;
var startTime = timeArr[timeIndex];

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tracker')
    .addItem('Today tracker', 'TodaySheetGen')
    .addSeparator()
    .addItem('Profile', 'setProfile')
    .addItem('Area and Activities', 'setAreas')
    .addToUi();
}

function setProfile() {
  var Tracker = SpreadsheetApp.getActiveSpreadsheet();
  var profilesheet = Tracker.getSheetByName('Profile');

  // check if today's sheet already exists 
  if (profilesheet != null) {
    SpreadsheetApp.setActiveSheet(Tracker.getSheets()[profilesheet.getIndex() - 1]);
    return
  }
  else {
    profilesheet = Tracker.insertSheet();
    profilesheet.setName('Profile');
    profilesheet.getRange(1, 1, 2).setValues([['Name'], ['Start time']]);
    var startcell = profilesheet.getRange(2, 2);
    var startTimeRule = SpreadsheetApp.newDataValidation().requireValueInList(timeArr, true).setAllowInvalid(false).build();
    startcell.setDataValidation(startTimeRule);
    startcell.setValue(timeArr[timeIndex]);
    profilesheet.getRange(6, 1).setValue('Records').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
    var areasheetname = 'Area and Activities'
    profilesheet.getRange('C7').setFormula('=TRANSPOSE(' + '\'' + areasheetname + '\'' + '!A2:A25)');
    profilesheet.setColumnWidths(3, 24, 75);
    profilesheet.autoResizeColumns(3, 24);
    return
  }

}

function setAreas() {
  var Tracker = SpreadsheetApp.getActiveSpreadsheet();
  var areaSheet = Tracker.getSheetByName('Area and Activities');

  // check if today's sheet already exists 
  if (areaSheet != null) {
    SpreadsheetApp.setActiveSheet(Tracker.getSheets()[areaSheet.getIndex() - 1]);
    return
  }
  else {
    areaSheet = Tracker.insertSheet();
    areaSheet.setName('Area and Activities');
    areaSheet.getRange(1, 1, 6, 2).setValues([['Area', 'Activities'], ['Area1', 'Sleeping'], ['Area2', 'Running, excercise, etc'], ['Area3', 'activities of area 3'], ['Area4', 'activities of area 4'], ['Area5', 'activities of area 5']]).setFontColors([['#ffffff', '#ffffff'], ['#ffffff', '#ffffff'], ['#ffffff', '#ffffff'], ['#ffffff', '#ffffff'], ['#000000', '#000000'], ['#000000', '#000000']]).setBackgrounds([['#000000', '#000000'], ['#980000', '#980000'], ['#ff0000', '#ff0000'], ['#ff9900', '#ff9900'], ['#ffff00', '#ffff00'], ['#00ff00', '#00ff00']]);
    areaSheet.setFrozenRows(1);
    areaSheet.setColumnWidth(2, 300);
    areaSheet.getRangeList(['A1:A25', 'B1']).setHorizontalAlignment('center').setVerticalAlignment('middle');
    areaSheet.deleteColumns(3, 24);
    areaSheet.deleteRows(26, 975);
    return
  }

}
function onEdit() {
  //e.getSheetName()
  var ssName = SpreadsheetApp.getActiveSheet().getSheetName().toString();
  var date = Utilities.formatDate(new Date(), "GMT + 5:30", "dd/MM/yyyy").toString();
  if (ssName == date) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ssName);
    // to get lastrow in column 1 (duration column)
    var rowIndex = 0;
    var cell_1value = sheet.getRange(2 + rowIndex, 1).getValue();
    var cell_2value = sheet.getRange(2 + rowIndex + 1, 1).getValue();
    while (cell_1value != "" || cell_2value != "") {
      rowIndex = rowIndex + 1;
      cell_1value = sheet.getRange(2 + rowIndex, 1).getValue();
      cell_2value = sheet.getRange(2 + rowIndex + 1, 1).getValue();
    };
    rowIndex = rowIndex + 2;
    var i = rowIndex - 1;
    var cellstrValue = sheet.getRange(i, 1).getDisplayValue();
    var durHour = 0;
    while (cellstrValue != "") {
      cellstrValue = sheet.getRange(i, 1).getDisplayValue();
      if (isNaN(parseInt(cellstrValue))) {
        break
      }
      durHour = durHour + parseInt(cellstrValue);
      i--;
    }
    var val = 60 - durHour
    if (sheet.getRange(rowIndex - 1, 3).getDisplayValue() != "") {
      if (val != 0) {
        if (val > 0) {
          sheet.getRange(rowIndex, 1).setValue(val.toString());
        } else if (val < 0) {
          var ui = SpreadsheetApp.getUi();
          ui.alert('Total number of minutes in an hour must not be greater than 60 minutes');
          val = val + parseInt(sheet.getRange(rowIndex - 1, 1).getDisplayValue());
          sheet.getRange(rowIndex - 1, 1).setValue(val.toString());
          sheet.getRange(rowIndex - 1, 3).clearContent();
        }
        rowIndex++;
      }
    }
    if (sheet.getRange(rowIndex - 1, 3).getDisplayValue() != "") {
      var prevTime = sheet.getRange(i, 2).getDisplayValue();
      var prevTimeindex = timeArr.indexOf(prevTime);
      sheet.getRange(rowIndex, 2).setValue(timeArr[(prevTimeindex + 1) % 24]).setHorizontalAlignment('center').setFontWeight('bold').setVerticalAlignment('middle');
      sheet.getRange(rowIndex, 3).setDataValidation(null);
      sheet.getRange(rowIndex + 1, 1).setValue('60');
    }

  } else if (ssName == 'Area and Activities') {
    SpreadsheetApp.flush();
    condColorFormat();
  };
  formattingChart();
}
/*
Description://Creates Today's tracker if not created
//elif asks to delete today's tracker //else go to today's tracker.
*/

function TodaySheetGen() {
  var ui = SpreadsheetApp.getUi();
  var Tracker = SpreadsheetApp.getActiveSpreadsheet();
  var date = Utilities.formatDate(new Date(), "GMT + 5:30", "dd/MM/yyyy");
  var todayTracker = Tracker.getSheetByName(date);
  // check if today's sheet already exists 
  if (todayTracker != null) {
    var result = ui.alert(todayTracker + " already created.", "Do you want to delete it?", ui.ButtonSet.YES_NO);
    if (result == ui.Button.YES) {
      Tracker.deleteSheet(todayTracker);
      Tracker.getSheetByName('Profile').deleteRow(8);
      Browser.msgBox(todayTracker + " deleted.");
      return
    }
    return
  }
  else {
    todayTracker = Tracker.insertSheet();
    todayTracker.deleteRows(201, 800);
    todayTracker.autoResizeRows(2, 199);
    todayTracker.setName(date);
    // for formatting of todayTracker(Daily sheets)
    todayTracker.appendRow(['Duration', 'Activity', 'Area']);
    todayTracker.getRange('A1:C1').setHorizontalAlignment('center').setFontWeight('bold').setVerticalAlignment('middle');
    var areaTargetRange = todayTracker.getRange('C2:C');
    // for areas Name validation
    var area_activitySource = Tracker.getSheetByName('Area and Activities');
    var areasSourceRange = area_activitySource.getRange('A2:A25');
    var areaRule = SpreadsheetApp.newDataValidation().requireValueInRange(areasSourceRange).setAllowInvalid(false).build();
    areaTargetRange.setDataValidation(areaRule);

    // for duration validation
    var durationTargetRange = todayTracker.getRange('A3:A');
    var durationRule = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 60).setHelpText('Time duration in minutes less than 60').setAllowInvalid(false).build();
    durationTargetRange.setDataValidation(durationRule);

    todayTracker.setColumnWidth(1, 70);
    todayTracker.setColumnWidth(2, 300);
    todayTracker.setColumnWidth(3, 100);
    todayTracker.insertRowAfter(1);
    var startTime = Tracker.getSheetByName('Profile').getRange(2, 2).getDisplayValue();
    todayTracker.getRange(2, 1, 1, 3).setValues([['', startTime, '']])
    todayTracker.getRange(3, 1).setValue(60);
    // Setting Clock
    todayTracker.getRange('E1').setValue('Time').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('bold');
    todayTracker.getRange('E2:E6').merge().setFormula('=SPARKLINE( ArrayFormula({ QUERY(ArrayFormula({ 0, 0, 1; 0, 0, 0.8; SEQUENCE(37,1,0,10), SIN(RADIANS(SEQUENCE(37,1,0,10))), COS(RADIANS(SEQUENCE(37,1,0,10))); SEQUENCE(12,1,30,30), 0.9 * SIN(RADIANS(SEQUENCE(12,1,30,30))), 0.9 * COS(RADIANS(SEQUENCE(12,1,30,30))); SEQUENCE(12,1,30,30), SIN(RADIANS(SEQUENCE(12,1,30,30))), COS(RADIANS(SEQUENCE(12,1,30,30))); SEQUENCE(4,1,90,90), 0.8 * SIN(RADIANS(SEQUENCE(4,1,90,90))), 0.8 * COS(RADIANS(SEQUENCE(4,1,90,90))); SEQUENCE(4,1,90,90), SIN(RADIANS(SEQUENCE(4,1,90,90))), COS(RADIANS(SEQUENCE(4,1,90,90)))}), "SELECT Col2, Col3 ORDER BY Col1", 0) ; IF( MINUTE(NOW()) = 0, 0, SIN(RADIANS(SEQUENCE(MINUTE(NOW())/60*360,1,1,1))) ), IF( MINUTE(NOW())=0, 1, COS(RADIANS(SEQUENCE(MINUTE(NOW())/60*360,1,1,1))) ) ; 0, 0 ; 0.75 * SIN(RADIANS((MOD(HOUR(NOW()),12)/12 * 360) + MINUTE(NOW())/60 * 30)), 0.75 * COS(RADIANS((MOD(HOUR(NOW()),12)/12 * 360) + MINUTE(NOW())/60 * 30)) }), {"linewidth",2 } )');
    todayTracker.setColumnWidth(5, 105);
    todayTracker.getRange('E7').setFormula('=NOW()-TODAY()').setNumberFormat('hh:mm am/pm').setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('bold');
    todayTracker.getRange('E9:F9').setValues([['Area', 'Duration']]).setVerticalAlignment('middle').setHorizontalAlignment('center').setFontWeight('bold');
    todayTracker.getRange('E10').setFormula('=SORT(UNIQUE(C3:C))');
    todayTracker.getRange('F10').setFormula('=IF(not(isblank(E10)),arrayformula(SUM(IF(C$3:C = E10,A$3:A))),"")').copyTo(todayTracker.getRange('F11:F33'));
    // inserting Chart
    var chartRange = todayTracker.getRange("E10:F33");
    var chart = todayTracker.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(chartRange)
      .setPosition(1, 7, 0, 0)
      .build();
    todayTracker.insertChart(chart);
    todayTracker.setFrozenRows(1);
    var profilesheet = Tracker.getSheetByName('Profile');
    profilesheet.insertRowAfter(7);
    profilesheet.getRange(8, 1).setValue(date);
    profilesheet.getRange(8, 2).setFormula('=SPARKLINE(TRANSPOSE(C8:Z8),{"charttype","column"})');
    profilesheet.getRange(8, 3).setFormula('=if(NOT(ISBLANK(C7)),IFNA(VLOOKUP(C7,' + '\'' + date + '\'' + '!$E$10:$F$33,2,FALSE),0),"")').copyTo(profilesheet.getRange('D8:Z8'));
    condColorFormat();
    // var formattingRules = SpreadsheetApp.getActiveSheet().getConditionalFormatRules();
    // // getting areasName colors
    // var areasFontColors = areasSourceRange.getFontColors();
    // var areas = areasSourceRange.getDisplayValues(); // Problems solved using getDisplayValues() after adding toString() method - this line of code is showing problem it is returning only first character of the string not whole as expected without toString() it took three index to give the same result
    // var areasCellColors = areasSourceRange.getBackgrounds();
    // for (var i = 0; i < 24; i++) {
    //   formattingRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(areas[i][0]).setFontColor(areasFontColors[i][0]).setBackground(areasCellColors[i][0]).setRanges([areaTargetRange, todayTracker.getRange('E9:E32')]).build());
    // }
    // SpreadsheetApp.getActiveSheet().setConditionalFormatRules(formattingRules);
    return
  }
}
function condColorFormat() {
  SpreadsheetApp.flush();
  var Tracker = SpreadsheetApp.getActiveSpreadsheet();
  var profilesheet = Tracker.getSheetByName('Profile');
  var date = Utilities.formatDate(new Date(), "GMT + 5:30", "dd/MM/yyyy");
  var todayTracker = Tracker.getSheetByName(date);
  var areaSheet = Tracker.getSheetByName('Area and Activities');
  var areasSourceRange = areaSheet.getRange('A2:A25');
  var areaTargetRange = todayTracker.getRange('C2:C');


  areaTargetRange.clearFormat();
  profilesheet.getRange('C7:Z7').clearFormat();
  todayTracker.getRange('E9:E32').clearFormat();

  // conditional formatting in todaytracker
  var formattingRules = todayTracker.getConditionalFormatRules();
  // getting areasName colors
  var areasFontColors = areasSourceRange.getFontColors();
  var areas = areasSourceRange.getDisplayValues(); // Problems solved using getDisplayValues() after adding toString() method - this line of code is showing problem it is returning only first character of the string not whole as expected without toString() it took three index to give the same result
  var areasCellColors = areasSourceRange.getBackgrounds();
  for (var i = 0; i < 24; i++) {
    formattingRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(areas[i][0]).setFontColor(areasFontColors[i][0]).setBackground(areasCellColors[i][0]).setRanges([areaTargetRange, todayTracker.getRange('E9:E32')]).build());
  }
  todayTracker.clearConditionalFormatRules();
  todayTracker.setConditionalFormatRules(formattingRules);

  // conditional formatting in profilesheet
  var formattingRules = profilesheet.getConditionalFormatRules();
  // getting areasName colors
  var areasFontColors = areasSourceRange.getFontColors();
  var areas = areasSourceRange.getDisplayValues(); // Problems solved using getDisplayValues() after adding toString() method - this line of code is showing problem it is returning only first character of the string not whole as expected without toString() it took three index to give the same result
  var areasCellColors = areasSourceRange.getBackgrounds();
  for (var i = 0; i < 24; i++) {
    formattingRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(areas[i][0]).setFontColor(areasFontColors[i][0]).setBackground(areasCellColors[i][0]).setRanges([profilesheet.getRange('C7:Z7')]).build());
  }
  profilesheet.clearConditionalFormatRules();
  profilesheet.setConditionalFormatRules(formattingRules);
  formattingChart();
  return
}

function formattingChart() {
  var Tracker = SpreadsheetApp.getActiveSpreadsheet();
  var date = Utilities.formatDate(new Date(), "GMT + 5:30", "dd/MM/yyyy");
  var todayTracker = Tracker.getSheetByName(date);
  var colorOfCells = [];
  var i = 10;
  var cell_value = todayTracker.getRange(i, 5).getDisplayValue();
  while (cell_value != "") {
    colorOfCells.push(todayTracker.getRange(i, 5).getBackground().toString());
    i++;
    cell_value = todayTracker.getRange(i, 5).getDisplayValue();
  }
  var chart = todayTracker.getCharts()[0];
  chart = chart.modify().setOption('backgroundColor', '#424949').setOption('colors',  colorOfCells ).build();
  todayTracker.updateChart(chart);
}
