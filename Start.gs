//some details were removed for privacy

function triggerProd() { //Warning: this function has a trigger, if you rename it ensure that the trigger is still there.
  start(ProdConfig, "false");
}

function triggerDev() {
  start(DevConfig, "false");
}

function triggerDevWithEmail() {
  start(DevConfig, "true");
}

function triggerProdWithEmail() { //Warning: this function has a trigger, if you rename it ensure that the trigger is still there.
  start(ProdConfig, "true");
}

function start(config, sendEmail) {
  setup(config);
  try {
    generateSpreadsheet();
    email(config, sendEmail);
  } catch (err) {
    var error = catchToString(err)
    Logger.log(error);
    throw (error);
  }
  return;

  function email(config, sendEmail) {
    var originalData = g.mainTable.sheet.getRange(2, 1, g.mainTable.lastRow - 1, g.mainTable.header.length).getValues();
    g.numberRedProjects = originalData.filter(function(item) {
      return (item[g.mainTable.header.indexOf(strings.status)] === "RED");
    }).length;
    g.numberYellowProjects = originalData.filter(function(item) {
      return (item[g.mainTable.header.indexOf(strings.status)] === "YELLOW");
    }).length;
    g.numberGreenProjects = originalData.filter(function(item) {
      return (item[g.mainTable.header.indexOf(strings.status)] === "GREEN");
    }).length;
    if (g.numberRedProjects > 0 && sendEmail === "true") {
      MailApp.sendEmail(config.emails,
        "Weekly Report: Project Health Status ( " + Utilities.formatDate(new Date(), "GMT", "'Week 'w") + " )",
        "There are ongoing projects that require attention. \n " +
        "\n* Red satus: " + g.numberRedProjects + " projects." +
        "\n* Amber status: " + g.numberYellowProjects + " projects." +
        "\n* Green status: " + g.numberGreenProjects + " projects.\n\n" +
        "Please visit the Project Health Dashboard for more information: https://datastudio.google.com");
    }
  }

  function setup(config) {
    g.config = config;
    g.spreadsheet = SpreadsheetApp.openById(config.sheetId);
    g.mainTable = new Table(strings.mainTable, g.spreadsheet.getSheetByName(strings.mainTable));
    g.healthTable = new Table(strings.healthTable, g.spreadsheet.getSheetByName(strings.healthTable));
    g.mainTable.sheet.clear();
    g.healthTable.sheet.clear();
  }
}

function catchToString(err) {
  var errInfo = "Error Info:\n";
  for (var prop in err) {
    errInfo += "  property: " + prop + "\n    value: [" + err[prop] + "]\n";
  }
  errInfo += "  toString(): " + " value: [" + err.toString() + "]";
  return errInfo;
}
