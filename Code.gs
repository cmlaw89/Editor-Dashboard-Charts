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
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pivot_sheet = ss.getSheetByName("Quality History");
  
  var user = pivot_sheet.getRange(1, 9).getValue().split(" ")[1].slice(0, -1);
  
  //Data for QR column chart 
  if (type == "QR") {
    //Extract QR data
    var QR_months = pivot_sheet.getRange("C1").getValue();
    var QR_month_data = pivot_sheet.getRange(4, 19, QR_months, 1).getValues();
    var QR_data = pivot_sheet.getRange(4, 22, QR_months, 1).getValues();
    
    if (QR_months > 0) {
      var QR_dic = {}
      for (var i = 0; i < QR_months; i++) {
        var month_year = dateToMonthYear(QR_month_data[i][0]);
        if (Object.keys(QR_dic).indexOf(month_year) == -1) {
          if (QR_data[i][0] != "") {
            QR_dic[month_year] = [QR_data[i][0], 1];
          }
        }
        else {
          if (QR_data[i][0] != "") {
            QR_dic[month_year][0] += QR_data[i][0];
            QR_dic[month_year][1] += 1;
          }
        }
      }
      for (var i = 0; i < Object.keys(QR_dic).length; i++) {
        QR_dic[Object.keys(QR_dic)[i]] = QR_dic[Object.keys(QR_dic)[i]][0] / QR_dic[Object.keys(QR_dic)[i]][1];
      }
    }
    
    var PR_months = pivot_sheet.getRange("D1").getValue();
    if (PR_months > 0) {
      //Extract proofreading QR data (if applicable)
      var PR_month_data = pivot_sheet.getRange(4, 26, PR_months, 1).getValues();
      var PR_data = pivot_sheet.getRange(4, 30, PR_months, 1).getValues();
      
      //Calculate the average proofreaing QR scores for each month
      var PR_avg = {};
      for (var i = 0; i < PR_months; i++) {
        var month_year = dateToMonthYear(PR_month_data[i][0]);
        if (Object.keys(PR_avg).indexOf(month_year) == -1) {
          if (PR_data[i][0] != "") {
            PR_avg[month_year] = [PR_data[i][0], 1];
          }
        }
        else {
          if (PR_data[i][0] != "") {
            PR_avg[month_year][0] += PR_data[i][0];
            PR_avg[month_year][1] += 1;
          }
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
    
    var data_table = [labels];
    
    var date = new Date()
    var current_month_year = dateToMonthYear(date);
    var entry_counter = 0;
    
    while (entry_counter < Math.min(12, Math.max(Object.keys(QR_dic).length, Object.keys(PR_avg).length))) {
      if ((Object.keys(QR_dic).indexOf(current_month_year) != -1) || (Object.keys(PR_avg).indexOf(current_month_year) != -1)) {
        var entry = [current_month_year];
        if (Object.keys(QR_dic).indexOf(current_month_year) != -1) {
          entry.push(QR_dic[current_month_year])
        }
        else {
          entry.push(null);
        }
        if (Object.keys(PR_avg).indexOf(current_month_year) != -1) {
          entry.push(PR_avg[current_month_year])
        }
        else {
          entry.push(null);
        }
        data_table.push(entry)
        entry_counter += 1;
      }
      date.setMonth(date.getMonth() - 1)
      current_month_year = dateToMonthYear(date);
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
        data_table.push([dateToMonthYear(month_data[i][0]), 
                        time_data[i][0], 
                        team_time_data[i][0]]);
      }
      else if (type == "IF") {
        data_table.push([dateToMonthYear(month_data[i][0]), 
                        IF_data[i][0], 
                        team_IF_data[i][0]]);
      }
      else if (type == "CPI") {
        data_table.push([dateToMonthYear(month_data[i][0]), 
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

function dateToMonthYear(date) {
  
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
  
  return months[date.getMonth() + 1] + " " + date.getFullYear().toString().slice(2,4);
}

