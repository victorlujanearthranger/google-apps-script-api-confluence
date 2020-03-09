//some details were removed for privacy

function runDevHistorical() {
  try {
    addToHistorical(DevConfig)
  } catch (err) {
    var error = catchToString(err)
    Logger.log(error);
    throw (error);
  }
}

function runProdHistorical() { //Warning: this function has a trigger, if you rename it ensure that the trigger is still there.
  try {
    addToHistorical(ProdConfig)
  } catch (err) {
    var error = catchToString(err)
    Logger.log(error);
    throw (error);
  }
}


function addToHistorical(config) {
  var arr = [];
  weekDataSpreadSheet = SpreadsheetApp.openById(config.sheetId);
  historicSpreadSheet = SpreadsheetApp.openById(config.historicSheetId);

  weekDataSheet = weekDataSpreadSheet.getSheetByName(strings.mainTable);
  lastRow = weekDataSheet.getLastRow();
  lastColumn = weekDataSheet.getLastColumn();
  var header = weekDataSheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  copySheet = historicSpreadSheet.getSheetByName("Local Copy");
  copySheet.clear();

  //desired order
  //Project	Date	Overall Status	Account	Portfolio	URL Health Report
  var rows1Column = weekDataSheet.getRange(1, header.indexOf(strings.prj) + 1, lastRow, 1).getValues();
  copySheet.getRange(1, 1, lastRow, 1).setValues(rows1Column);
  rows1Column = weekDataSheet.getRange(1, header.indexOf(strings.date) + 1, lastRow, 1).getValues();
  copySheet.getRange(1, 2, lastRow, 1).setValues(rows1Column);
  rows1Column = weekDataSheet.getRange(1, header.indexOf(strings.status) + 1, lastRow, 1).getValues();
  copySheet.getRange(1, 3, lastRow, 1).setValues(rows1Column);
  rows1Column = weekDataSheet.getRange(1, header.indexOf(strings.acc) + 1, lastRow, 1).getValues();
  copySheet.getRange(1, 4, lastRow, 1).setValues(rows1Column);
  rows1Column = weekDataSheet.getRange(1, header.indexOf(strings.portfolio) + 1, lastRow, 1).getValues();
  copySheet.getRange(1, 5, lastRow, 1).setValues(rows1Column);
  rows1Column = weekDataSheet.getRange(1, header.indexOf(strings.url + " " + strings.healthTable) + 1, lastRow, 1).getValues();
  copySheet.getRange(1, 6, lastRow, 1).setValues(rows1Column);


  var data = copySheet.getDataRange().getValues();
  var dataFiltered = data.filter(function(item) {
    Logger.log(item);
    return (item[1].length != 0 && item[1] != strings.date); //date not empty
  });

  var historicalSheet = historicSpreadSheet.getSheetByName("Historical");
  historicalSheet.getRange(historicalSheet.getLastRow() + 1, 1, dataFiltered.length, dataFiltered[0].length).setValues(dataFiltered);
  historicalSheet.getRange("B:B").setNumberFormat('MM/dd/yyyy');
  removeDuplicates(historicalSheet);

  historicalSheet.getRange("A2:F").sort([{
    column: 2,
    ascending: true
  }, {
    column: 1,
    ascending: false
  }]);


  return;
}

/**
 * Removes duplicate rows from the current sheet.
 */
function removeDuplicates(sheet) {
  var data = sheet.getDataRange().getValues();
  var newData = [];
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if (row.join() == newData[j].join()) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}
