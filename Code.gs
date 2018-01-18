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
    
    var QR_dic = {}
    if (QR_months > 0) {
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
    var PR_avg = {};
    if (PR_months > 0) {
      //Extract proofreading QR data (if applicable)
      var PR_month_data = pivot_sheet.getRange(4, 26, PR_months, 1).getValues();
      var PR_data = pivot_sheet.getRange(4, 30, PR_months, 1).getValues();
      
      //Calculate the average proofreaing QR scores for each month
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
                  {label: 'QR Score', id: "QR", type: 'number'},
                 {label: 'Average QR Score (Proofreading)', id: "PR", type: 'number'}];
    
    var data_table = [labels];
    
    var date = new Date()
    var current_month_year = dateToMonthYear(date);
    var entry_counter = 0;
    
    while (entry_counter < Math.min(12, Math.max(Object.keys(QR_dic).length, Object.keys(PR_avg).length))) {
      if ((Object.keys(QR_dic).indexOf(current_month_year) != -1) || (Object.keys(PR_avg).indexOf(current_month_year) != -1)) {
        var entry = [current_month_year];
        if (Object.keys(QR_dic).indexOf(current_month_year) != -1) {
          entry.push({v:QR_dic[current_month_year], f:getGrade(QR_dic[current_month_year], "QR")})
        }
        else {
          entry.push({v:null, f:null});
        }
        if (Object.keys(PR_avg).indexOf(current_month_year) != -1) {
          entry.push({v:PR_avg[current_month_year], f:getGrade(PR_avg[current_month_year], "PR")})
        }
        else {
          entry.push({v:null, f:null});
        }
        data_table.splice(1, 0, entry);
        entry_counter += 1;
      }
      date.setMonth(date.getMonth() - 1)
      current_month_year = dateToMonthYear(date);
    }
    
    if (PR_months == 0) {
      for (var i = 0; i < data_table.length; i++) {
        data_table[i].pop();
      }
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

function getGrade(score, stage) {
  if (score < 1) {
    return "F"
  }
  else if ((score < 2) && (score >= 1)) {
    if (stage == "QR") {
      return "E"
    }
    else if (stage == "PR") {
      return "F"
    }
  }
  else if ((score >= 2) && (score < 3)) {
    if (stage == "QR") {
      return "D"
    }
    else if (stage == "PR") {
      return "E"
    }
  }
  else if ((score >= 3) && (score < 4)) {
    if (stage == "QR") {
      return "C"
    }
    else if (stage == "PR") {
      return "D"
    }
  }
  else if ((score >= 4) && (score < 5)) {
    if (stage == "QR") {
      return "B2"
    }
    else if (stage == "PR") {
      return "C"
    }
  }
  else if ((score >= 5) && (score < 6)) {
    if (stage == "QR") {
      return "B1"
    }
    else if (stage == "PR") {
      return "B2"
    }
  }
  else if ((score >= 6) && (score < 7)) {
    if (stage == "QR") {
      return "A2"
    }
    else if (stage == "PR") {
      return "B1"
    }
  }
  else if ((score >= 7) && (score < 8)) {
    if (stage == "QR") {
      return "A1"
    }
    else if (stage == "PR") {
      return "A2"
    }
  }
  else if ((score >= 8) && (score < 9)) {
    if (stage == "QR") {
      return "A0"
    }
    else if (stage == "PR") {
      return "A1"
    }
  }
  else if (score >= 9) {
    return "A0"
  }
}


function include(filename) {
  //Adds stylesheet and javascript to Index.html
  
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}









