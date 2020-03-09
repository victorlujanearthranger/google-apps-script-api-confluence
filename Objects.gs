//some details were removed or changed for privacy and security

var g = {
  spreadsheet: null,
  documentApp: null,
  config: null,
  mainTable: null,
  healthTable: null,
  numberRedProjects: 0,
  numberYellowProjects: 0,
  numberGreenProjects: 0
}

var strings = {
  prj: "Project",
  acc: "Account",
  portfolio: "Portfolio",
  comments: "Comments for eStaff",
  status: "Overall Status",
  pjm: "Project Manager",
  healthTable: "Health Report",
  mainTable: "Full Data",
  date: "Date",
  url: "URL",
  pjPSA: "Project PSA ID",
  startDate: "Start Date",
  endDate: "End Date",
  psaLink: "PSA Link"
};

var Table = function(name, sheet) {
  this.sheet = sheet;
  this.name = name;
  this.header = [];
  this.lastRow = 0;
}

var VLookupParameters = function() {
  this.sheetA = "";
  this.sheetB = "";
  this.rangeA = null;
  this.rangeB = null;
  this.value = "";
}
