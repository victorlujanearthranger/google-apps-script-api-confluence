//some details were removed for privacy

function getActiveProjects() {
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify([]),
    'headers': {
      'x-api-key': GlobalConfig.APIKey
    }
  };
  var projects = UrlFetchApp.fetch(GlobalConfig.activeProjectsAPI, options);
  return projects;
}

function projectToRow(project) {
  return [project.Name, project.Account, project.PjM, project.ID, project.Portfolio, project.StartDate, project.EndDate, "https://your-domain.force.com/r/pse__Proj__c/" + project.ID + "/view"];
}

function addActiveProjects() {
  g.mainTable.sheet = g.spreadsheet.getSheetByName(strings.mainTable);

  g.mainTable.header = [strings.prj, strings.acc, strings.pjm, strings.pjPSA, strings.portfolio, strings.startDate, strings.endDate, strings.psaLink];
  g.mainTable.sheet.appendRow(g.mainTable.header);

  var rows = [];
  var activeProjects = JSON.parse(getActiveProjects());
  for (var i in activeProjects) {
    var values = projectToRow(activeProjects[i]);
    rows.push(values);
  }
  g.mainTable.sheet.getRange(2, 1, rows.length, g.mainTable.header.length).setValues(rows);
  g.mainTable.lastRow = g.mainTable.sheet.getLastRow();
}
