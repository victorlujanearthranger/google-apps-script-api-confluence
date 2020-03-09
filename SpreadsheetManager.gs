//some details were removed for privacy

function generateSpreadsheet() {
  pullConfluenceData();
  formatTable(g.healthTable);
  addActiveProjects();
  mergeWithVLookup();
  formatTable(g.mainTable);
  return;

  function pullConfluenceData() {
    g.healthTable.header = [strings.date, strings.url + " " + strings.healthTable, strings.prj, strings.status, strings.comments];
    g.healthTable.sheet.appendRow(g.healthTable.header);
    query = "lines?pageSize=60&pageIndex=0&cql=label+%3D+%22weekly-health%22+and+created+%3E%3D+now+(+%27-" +
      GlobalConfig.daysBack +
      "%27+)&spaceKey=PDS&contentId=768213015&headings=Project%2C+Overall+Status%2C+Comments+for+eStaff%2C+&sortBy=Project";
    var responseContentText = getRestAPIResponse(query);
    var dataJSON = JSON.parse(responseContentText);
    insertRowsWithData(dataJSON, g.healthTable.sheet);

    function getRestAPIResponse(query) {
      var API_URL = 'https://your-domain.atlassian.net/wiki/rest/masterdetail/1.0/detailssummary/' + query;
      var authHeader = 'Basic ' + Utilities.base64Encode(GlobalConfig.confluenceUSERNAME + ':' + GlobalConfig.confluencePASSWORD);
      var options = {
        headers: {
          Authorization: authHeader
        }
      }
      var response = UrlFetchApp.fetch(API_URL, options);
      return response.getContentText()
    }

    function insertRowsWithData(dataJSON, sheet) {
      var rows = [];
      for (var i in dataJSON.detailLines) {
        var line = dataJSON.detailLines[i];
        var values = [];
        var date = "";
        var patt = /([12]\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01]))/;
        var res = patt.exec(line.title);
        if (res) {
          date = res[0];
          values.push("'" + date);
          var url = 'https://wizeline.atlassian.net/wiki' + line.relativeLink;
          values.push(url);
        } else {
          values.push("");
          values.push("");
        }
        for (var x in line.details) {
          var xml = line.details[x].split("<li>").join("<li>\n"); //trick to show the information spaced when the code below doesn't properly navigate lists in order to add a new line \n or a space.
          var document = XmlService.parse("<!DOCTYPE html><html><div>" + xml + "</div></html>");
          var root = document.getRootElement();
          var items = root.getChildren();
          values.push(root.getValue().trim());
        }

        rows.push(values);
      }

      sheet.getRange(2, 1, rows.length, g.healthTable.header.length).setValues(rows);
      g.healthTable.lastRow = sheet.getLastRow();
    }
  }


  function mergeWithVLookup() {
    var vlookup = new VLookupParameters();
    //vlookup will need the two sheet names
    vlookup.sheetA = g.mainTable.sheet.getSheetName();
    vlookup.sheetB = g.healthTable.sheet.getSheetName();

    const additionalHeaders = [strings.status, strings.comments, strings.date, strings.url + " " + strings.healthTable];
    insertColumns(additionalHeaders);

    //we need a range with all of the projects, because this is the key that will help us merge both tables/sheets
    var iniRow = 2;
    var iniCol = g.mainTable.header.indexOf(strings.prj) + 1; //column with projects
    var numRows = g.mainTable.lastRow - 1;
    var numCols = 1;
    vlookup.rangeA = g.mainTable.sheet.getRange(iniRow, iniCol, numRows, numCols).getA1Notation();

    //we also need all the columns to merge. 
    iniCol = 1;
    numRows = g.healthTable.lastRow;
    numCols = g.healthTable.sheet.getLastColumn();
    vlookup.rangeB = g.healthTable.sheet.getRange(iniRow, iniCol, numRows, numCols).getA1Notation();

    //finally build the formula and add it to the relevant cell.
    for (var i in additionalHeaders) {
      addFormula(g.healthTable, additionalHeaders[i], vlookup);
    }

    function addFormula(healthTable, text, vlookup) {
      var mainTable = g.mainTable;
      vlookup.value = healthTable.header.indexOf(text) + 1;
      var emptyValue = "";
      if (text === strings.comments) {
        emptyValue = "No update in the last " + GlobalConfig.daysBack;
      }
      var formula = "=IFERROR(ARRAYFORMULA(VLOOKUP('" + vlookup.sheetA + "'!" + vlookup.rangeA + ",'" + vlookup.sheetB + "'!" + vlookup.rangeB + ", " + vlookup.value + ', FALSE )),"' + emptyValue + '")';
      const column = mainTable.header.indexOf(text) + 1;
      const cell = mainTable.sheet.getRange(2, column);
      cell.setFormula(formula);
    }
  }

  function insertColumns(columns) {
    var table = g.mainTable;
    table.header = table.header.concat(columns);
    table.sheet.getRange(1, 1, 1, table.header.length).setValues([table.header]);
  }

  function formatTable(table) {
    addConditionalFormat(table);
    table.sheet.getRange(1, 1, 1, 100).setFontWeight("bold");
    table.sheet.setFrozenRows(1);
    table.sheet.getRange(1, 1, 1, 100).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    switch (table.name) {
      case strings.healthTable:
        //We change the position of the date and url to the end. Its more important to have the project in the first column in order to use it as a key in the vlookup
        table.sheet.getRange('A:B').moveTo(table.sheet.getRange(1, table.sheet.getLastColumn() + 1));
        table.sheet.deleteColumns(1, 2);
        table.header.push(table.header.shift());
        table.header.push(table.header.shift());
        break;
      case strings.mainTable:
        //We want to change the column order because there is a dashboard that expects the info to be in certain order.
        rearrangeColumns(table.sheet, [table.header.indexOf(strings.prj),
          table.header.indexOf(strings.acc),
          table.header.indexOf(strings.pjm),
          table.header.indexOf(strings.pjPSA),
          table.header.indexOf(strings.portfolio),
          table.header.indexOf(strings.status),
          table.header.indexOf(strings.comments),
          table.header.indexOf(strings.date),
          table.header.indexOf(strings.url + " " + strings.healthTable),
          table.header.indexOf(strings.startDate),
          table.header.indexOf(strings.endDate),
          table.header.indexOf(strings.psaLink)
        ]);
        table.header = table.sheet.getRange(1, 1, 1, table.sheet.getLastColumn()).getValues()[0];
        break;
    }

    function addConditionalFormat(table) {
      var rules = table.sheet.getConditionalFormatRules();
      var range = table.sheet.getRange(2, 1, table.lastRow - 1, table.header.length);
      var rule = SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground("#cbc9c9").setRanges([range]).build();
      rules.push(rule);
      table.sheet.setConditionalFormatRules(rules);
      addConditionalFormattingHealthStatus(table);
    }

    function addConditionalFormattingHealthStatus(table) {
      var range = table.sheet.getRange(2, 1, table.lastRow - 1, table.header.length);
      var rules = table.sheet.getConditionalFormatRules();
      var rule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("GREEN").setBackground("#28B672").setFontColor("#F4F5F7").setRanges([range]).build();
      rules.push(rule);
      rule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("RED").setBackground("#C73B0F").setFontColor("#F4F5F7").setRanges([range]).build();
      rules.push(rule);
      rule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("YELLOW").setBackground("#ECA628").setFontColor("#F4F5F7").setRanges([range]).build();
      rules.push(rule);
      table.sheet.setConditionalFormatRules(rules);
    }
  }
}

//https://tanaikech.github.io/2020/02/10/rearranging-columns-on-google-spreadsheet-using-google-apps-script/
//When [4, 3, 1, 0, 2] is used, the columns A, B, C, D, E are rearranged to E, D, B, A, C.
function rearrangeColumns(sheet, ar) {
  var obj = ar.reduce(function(ar, e, i) {
    return ar.concat({
      from: e + 1,
      to: i + 1
    });
  }, []);
  obj.sort(function(a, b) {
    return a.to < b.to ? -1 : 1;
  });
  obj.forEach(function(o) {
    if (o.from != o.to) sheet.moveColumns(sheet.getRange(1, o.from), o.to);
    obj.forEach(function(e, i) {
      if (e.from < o.from) obj[i].from += 1;
    });
  });
}
