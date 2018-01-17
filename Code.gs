function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('Wallace')
      .addItem('Trend Charts', 'viewCharts')
      .addToUi();
}

function viewCharts() {

  var html = HtmlService
  .createTemplateFromFile("Index")
  .evaluate()
  .setTitle("Test Chart")
  .setHeight(450)
  .setWidth(750)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  SpreadsheetApp.getUi().showModelessDialog(html, "Trend Charts")
}

function getDataTable(type) {
  //Extracts and formats the data for use in google.visualization.arrayToDataTable()
  
  
  var months = {1: "Jan",
                2: "Feb",
                3: "Mar",
                4: "Apr",
                5: "May",
                6: "Jun",
                7: "Jul",
                8: "Aug",
                9: "Sep",
                10: "Oct",
                11: "Nov",
                12: "Dec"}
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pivot_sheet = ss.getSheetByName("Quality History");
  
  var user = pivot_sheet.getRange(1, 9).getValue().split(" ")[1].slice(0, -1);
  
  //Data for QR column chart 
  if (type == "QR") {
    //Extract QR data
    var QR_months = pivot_sheet.getRange("C1").getValue();
    var QR_month_data = pivot_sheet.getRange(4, 19, QR_months, 1).getValues();
    var QR_data = pivot_sheet.getRange(4, 22, QR_months, 1).getValues();
    
    var PR_months = pivot_sheet.getRange("D1").getValue();
    if (PR_months > 0) {
      //Extract proofreading QR data (if applicable)
      var PR_month_data = pivot_sheet.getRange(4, 26, PR_months, 1).getValues();
      var PR_data = pivot_sheet.getRange(4, 30, PR_months, 1).getValues();
      
      //Calculate the average proofreaing QR scores for each month
      var PR_avg = {};
      for (var i = 0; i < PR_months; i++) {
        var month_year = months[PR_month_data[i][0].getMonth() + 1] + " " + PR_month_data[i][0].getFullYear().toString().slice(2,4);
        if (Object.keys(PR_avg).indexOf(month_year) == -1) {
          PR_avg[month_year] = [PR_data[i][0], 1];
        }
        else {
          PR_avg[month_year][0] += PR_data[i][0];
          PR_avg[month_year][1] += 1;
        }
      }
      for (var i = 0; i < Object.keys(PR_avg).length; i++) {
        PR_avg[Object.keys(PR_avg)[i]] = PR_avg[Object.keys(PR_avg)[i]][0] / PR_avg[Object.keys(PR_avg)[i]][1];
      }
    }
    
    //Add data series labels
    var labels = [{label: 'Month', id: 'Month'}, 
                  {label: 'QR Score', id: "QR", type: 'number'}];
    if (PR_months > 0) {
      labels.push({label: 'Average QR Score (Proofreading)', id: "PR", type: 'number'});
    }
    
    //Format null values as per google.visualization.arrayToDataTable() requirements
    for (var i = 1; i < QR_months; i++) {
      if (QR_data[i][0] == "") {
        QR_data[i][0] = null;
      }
    }
    
    var data_table = [labels];
    var start_month = 1
    if (QR_months > 12) {
      start_month = QR_months - 12;
    }
    
    //Assemble the array
    for (var i = start_month; i < QR_months; i++) {
      var month_year = months[QR_month_data[i][0].getMonth() + 1] + " " + QR_month_data[i][0].getFullYear().toString().slice(2,4);
      var entry = [month_year, QR_data[i][0]]
      if (PR_months > 0) {
        if (PR_avg[month_year]) {
          entry.push(PR_avg[month_year]);
        }
        else {
          entry.push(null);
        }
      }
      data_table.push(entry)
    }
  }
  
  //Data for time score and IF line charts
  else {
    //Extract time and IF data
    var num_months = pivot_sheet.getRange("B1").getValue();
    var month_data = pivot_sheet.getRange(3, 1, num_months, 1).getValues();
    var time_data = pivot_sheet.getRange(3, 2, num_months, 1).getValues();
    var team_time_data = pivot_sheet.getRange(3, 11, num_months, 1).getValues();
    var IF_data = pivot_sheet.getRange(3, 3, num_months, 1).getValues();
    var team_IF_data = pivot_sheet.getRange(3, 17, num_months, 1).getValues();
    var CPI_data = pivot_sheet.getRange(3, 6, num_months, 1).getValues();
    
    //Add data series labels
    var labels = [{label: 'Month', id: 'Month'}, 
                {label: user, id: user, type: 'number'}];
    if (type == "IF" || type == "Time") {
      labels.push({label: 'Team Average', id: "TAvg", type: 'number'});
    }
    
    //Format null values as per google.visualization.arrayToDataTable() requirements
    for (var i = 1; i < num_months; i++) {
      if (time_data[i][0] == "") {
        time_data[i][0] = null;
      }
      if (team_time_data[i][0] == "") {
        team_time_data[i][0] = null;
      }
      if (IF_data[i][0] == "") {
        IF_data[i][0] = null;
      }
      if (team_IF_data[i][0] == "") {
        team_IF_data[i][0] = null;
      }
      if (CPI_data[i][0] == "") {
        CPI_data[i][0] = null;
      }
    }
    
    var data_table = [labels];
    var start_month = 1
    if (num_months > 12) {
      start_month = num_months - 12;
    }
    //Assemble the arrays
    for (var i = start_month; i < num_months; i++) {
      if (type == "Time") {
        data_table.push([months[month_data[i][0].getMonth() + 1] + " " + month_data[i][0].getFullYear().toString().slice(2,4), 
                        time_data[i][0], 
                        team_time_data[i][0]]);
      }
      else if (type == "IF") {
        data_table.push([months[month_data[i][0].getMonth() + 1] + " " + month_data[i][0].getFullYear().toString().slice(2,4), 
                        IF_data[i][0], 
                        team_IF_data[i][0]]);
      }
      else if (type == "CPI") {
        data_table.push([months[month_data[i][0].getMonth() + 1] + " " + month_data[i][0].getFullYear().toString().slice(2,4), 
                        CPI_data[i][0]]);
      }
    }
  }
  
  return data_table
}


function include(filename) {
  //Adds stylesheet and javascript to Index.html
  
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
