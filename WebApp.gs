//Some details were removed for privacy.


function doGet(e) {
  name = e.parameter["name"];
  runHistorical = e.parameter["runHistorical"];
  //Logger.log("invoked. name: " + name + " runHistorical: " + runHistorical);

  var template = HtmlService.createTemplate(returnHTML());
  return template.evaluate();


  function returnHTML() {
    if (name == null || name.length === 0) {
      return htmlError();
    }
    if (runHistorical != null && runHistorical === "true") {
      return htmlHistoricalProcessing(name);
    }

    return htmlDataProcessing(name);
  }

  function htmlDataProcessing(name) {
    switch (name) {
      case ProdConfig.name:
        func = "triggerProd()";
        break;
      case DevConfig.name:
        func = "triggerDev()";
        break;
      default:
        return htmlError();
        break;
    }

    return '<head><title>Script Complete</title></head><script>' +
      '<?= ' + func + ' ?></script><body>  <font face="verdana"><p> <header><h1>Script ' +
      'executed successfully</h1></header></p>You can close this window now. ' +
      'Find the updated information in the <a href="https://datastudio.google.com">' +
      'Dashboard</a></font></body>';
  }


  function htmlHistoricalProcessing(name) {
    switch (name) {
      case ProdConfig.name:
        func = "runProdHistorical()";
        break;
      case DevConfig.name:
        func = "runDevHistorical()";
        break;
      default:
        return htmlError();
        break;
    }

    return '<head><title>Historical Project Health</title></head><script>' +
      '<?= ' + func + ' ?></script><body>  <font face="verdana"><p> <header><h1>Historical script ' +
      'executed successfully</h1></header></p>You can close this window now. Find the updated information in the <a href="https://datastudio.google.com">' +
      'Dashboard</a></font></body>';
  }

  function htmlError() {
    return '<head><title>Delivery Weekly Report</title></head><body>' +
      '<font face="verdana"><p><header><h1>Unknown Command</h1></header></p>' +
      'Consult the team to prevent this error. Good Luck! </font></body>';
  }
}
