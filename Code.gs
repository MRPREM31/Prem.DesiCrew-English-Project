function buildAdvancedDashboard() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName("Summary");

  if (!summary) return;

  const lastRow = summary.getLastRow();
  if (lastRow < 2) return;

  let dashboard = ss.getSheetByName("Dashboard");

  if (!dashboard) {
    dashboard = ss.insertSheet("Dashboard");
  } else {

    dashboard.clearContents();

    const charts = dashboard.getCharts();
    charts.forEach(chart => dashboard.removeChart(chart));
  }


  // ---------- TITLE ----------
  dashboard.getRange("A1")
    .setValue("TEAM ANALYTICS DASHBOARD")
    .setFontSize(22)
    .setFontWeight("bold");


  // ---------- CHART 1 : TOTAL TIME ----------
  const chart1 = dashboard.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(summary.getRange("A2:A" + lastRow))
    .addRange(summary.getRange("D2:D" + lastRow))
    .setPosition(3,1,0,0)
    .setOption("title","Total Work Time by Member")
    .build();

  dashboard.insertChart(chart1);


  // ---------- CHART 2 : COMPLETED FILES ----------
  const chart2 = dashboard.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(summary.getRange("A2:A" + lastRow))
    .addRange(summary.getRange("C2:C" + lastRow))
    .setPosition(3,9,0,0)
    .setOption("title","Completed Files by Member")
    .build();

  dashboard.insertChart(chart2);


  // ---------- CHART 3 : TEAM CONTRIBUTION ----------
  const chart3 = dashboard.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(summary.getRange("A2:A" + lastRow))
    .addRange(summary.getRange("C2:C" + lastRow))
    .setPosition(24,5,0,0)   // moved down
    .setOption("title","Team Contribution")
    .build();

  dashboard.insertChart(chart3);



  // ---------- TEAM TOTALS ----------
  const files = summary.getRange("C2:C"+lastRow).getValues().flat();
  const time = summary.getRange("D2:D"+lastRow).getValues().flat();

  const totalFiles = files.reduce((a,b)=>a+b,0);
  const totalTime = time.reduce((a,b)=>a+b,0);

  dashboard.getRange("A44")   // moved down
    .setValue("TEAM TOTALS")
    .setFontWeight("bold");

  dashboard.getRange("A46").setValue("Total Files Completed");
  dashboard.getRange("B46").setValue(totalFiles);

  dashboard.getRange("A47").setValue("Total Work Time (Min)");
  dashboard.getRange("B47").setValue(totalTime);

  // ---------- CONVERT MINUTES TO H:M:S ----------
  const hours = Math.floor(totalTime / 60);
  const minutes = Math.floor(totalTime % 60);
  const seconds = Math.floor((totalTime - Math.floor(totalTime)) * 60);

  // Format as HH:MM:SS
  const formattedTime = 
    String(hours).padStart(2, '0') + ":" +
    String(minutes).padStart(2, '0') + ":" +
    String(seconds).padStart(2, '0');

  // Display in dashboard
  dashboard.getRange("A48").setValue("Total Work Time (H:M:S)");
  dashboard.getRange("B48").setValue(formattedTime);



  // ---------- LEADERBOARD ----------
  const data = summary.getRange(2,1,lastRow-1,4).getValues();

  data.sort((a,b)=>b[3]-a[3]);

  dashboard.getRange("F44")   // moved down
    .setValue("TOP PERFORMERS")
    .setFontSize(16)
    .setFontWeight("bold");

  dashboard.getRange("F46:H46")
    .setValues([["Rank","Name","Total Time"]]);

  const medals=["🥇","🥈","🥉"];

  const leaderboard=[];

  for(let i=0;i<data.length;i++){
    leaderboard.push([
      medals[i] || (i+1),
      data[i][0],
      data[i][3]
    ]);
  }

  dashboard.getRange(47,6,leaderboard.length,3)
    .setValues(leaderboard);



  // ---------- COMPLETION % ----------
  const progressData=[];

  data.forEach(row=>{
    let percent = row[1] === 0 ? 0 : row[2]/row[1];
    progressData.push([row[0],percent]);
  });

  dashboard.getRange("A52:B52")  // moved down
    .setValues([["Name","Completion %"]]);

  dashboard.getRange(53,1,progressData.length,2)
    .setValues(progressData);

  dashboard.getRange("B53:B"+(progressData.length+52))
    .setNumberFormat("0%");
}


// ---------- BACKUP AUTO REFRESH ----------
function refreshDashboard(){
  buildAdvancedDashboard();
}


function buildDailyProduction(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  if(!sheet) return;

  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return;

  const data = sheet.getRange("B2:H"+lastRow).getValues();

  const daily = {};

  data.forEach(row => {

    const name = row[0];
    const date = row[3];
    const time = parseFloat(row[6]) || 0;

    if(!date) return;

    const key = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");

    if(!daily[key]){
      daily[key] = {files:0,time:0,people:new Set()};
    }

    daily[key].files += 1;
    daily[key].time += time;
    if(name) daily[key].people.add(name);

  });


  let prodSheet = ss.getSheetByName("Daily Production");

  if(!prodSheet){
    prodSheet = ss.insertSheet("Daily Production");
  }else{
    prodSheet.clearContents();
    const charts = prodSheet.getCharts();
    charts.forEach(c => prodSheet.removeChart(c));
  }


  // TITLE
  prodSheet.getRange("A1")
  .setValue("DAILY PRODUCTION REPORT")
  .setFontSize(20)
  .setFontWeight("bold");


  // HEADERS
  prodSheet.getRange("A3:D3")
  .setValues([["Date","Files Completed","Total Time (Min)","Head Count"]])
  .setFontWeight("bold")
  .setBackground("#f4b400");


  // DATA
  const rows = [];

  Object.keys(daily)
  .sort((a,b)=> new Date(b) - new Date(a))
  .forEach(date=>{

    rows.push([
      Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "dd MMM yyyy"),
      daily[date].files,
      daily[date].time,
      daily[date].people.size
    ]);

  });

  if(rows.length){
    prodSheet.getRange(4,1,rows.length,4).setValues(rows);
  }

  prodSheet.getRange("C4:C"+(rows.length+3)).setNumberFormat("0.00");

  prodSheet.getRange("A4:D"+(rows.length+3))
  .setBorder(true,true,true,true,true,true);

  prodSheet.setFrozenRows(3);


  // Calculate the last row for the charts (max 20 days + 3 for the header row offset)
  const chartEndRow = Math.min(rows.length, 20) + 3;

  // CHART 1 - PRODUCTION
  const chart1 = prodSheet.newChart()
  .setChartType(Charts.ChartType.COLUMN)
  .addRange(prodSheet.getRange("A3:C" + chartEndRow)) // Use chartEndRow here
  .setPosition(5,6,0,0)
  .setOption("title", "Daily Production Output (Last 20 Days)") // Updated title for clarity
  .build();

  prodSheet.insertChart(chart1);


  // CHART 2 - HEADCOUNT
  const chart2 = prodSheet.newChart()
  .setChartType(Charts.ChartType.BAR)
  .addRange(prodSheet.getRange("A3:A" + chartEndRow)) // Use chartEndRow here
  .addRange(prodSheet.getRange("D3:D" + chartEndRow)) // Use chartEndRow here
  .setPosition(25,6,0,0)
  .setOption("title", "Daily Head Count (Last 20 Days)") // Updated title for clarity
  .build();

  prodSheet.insertChart(chart2);

}


function buildAnnotatorDaily(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  if(!sheet) return;

  const lastRow = sheet.getLastRow();
  const data = sheet.getRange("B2:H"+lastRow).getValues();

  const members = {};
  const dates = new Set();

  data.forEach(r=>{

    const name = r[0];
    const date = r[3];
    const time = parseFloat(r[6]) || 0;

    if(!name || !date) return;

    const d = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(),"dd-MMM");

    dates.add(d);

    if(!members[name]){
      members[name] = {};
    }

    if(!members[name][d]){
      members[name][d] = 0;
    }

    members[name][d] += time;

  });


  const dateList = Array.from(dates).sort((a,b)=>{
    return new Date(b + "-2026") - new Date(a + "-2026");
  });


  let sheet3 = ss.getSheetByName("Member Daily");

  if(!sheet3){
    sheet3 = ss.insertSheet("Member Daily");
  }else{
    sheet3.clear();
  }


  const headers = ["Annotators","Total"].concat(dateList);

  // ---------- HEADER STYLE ----------
  sheet3.getRange(1,1,1,headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#1F4E78")
    .setFontColor("white");


  const rows = [];

  Object.keys(members).forEach(name=>{

    let total = 0;
    const row = [name];

    dateList.forEach(d=>{
      const t = members[name][d] || 0;
      row.push(t);
      total += t;
    });

    row.splice(1,0,total);

    rows.push(row);

  });


  if(rows.length){

    sheet3.getRange(2,1,rows.length,headers.length).setValues(rows);

    // ---------- NAME COLUMN STYLE ----------
    sheet3.getRange(2,1,rows.length,1)
      .setBackground("#E8F0FE")
      .setFontWeight("bold");

    // ---------- TABLE BORDER ----------
    sheet3.getRange(1,1,rows.length+1,headers.length)
      .setBorder(true,true,true,true,true,true);

  }

}

function onEdit(e){
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();

  // Auto clean name when editing File Completed column B
  if(sheetName === "File Complited" && e.range.getColumn() === 2){
    let value = e.range.getValue();

    if(value){
      let clean = value.toString().trim();
      clean = clean.replace(/\s+/g, " ");
      clean = clean.toLowerCase().replace(/\b\w/g, l => l.toUpperCase());

      if(clean !== value){
        e.range.setValue(clean);
      }
    }
  }

  // Refresh dashboards
  if(sheetName === "File Complited"){
    buildDailyProduction();
    buildAnnotatorDaily();
    buildQCDashboard();
  }

  if(sheetName === "Summary"){
    buildAdvancedDashboard();
  }

  if(sheetName === "QC(Prem)"){
    buildQCDashboard();
  }
}

function refreshAll(){
  buildSummary();
  buildDailyProduction();
  buildAnnotatorDaily();
  buildAdvancedDashboard();
  buildQCDashboard();
}

function protectDailyAutoEmail() {
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();

  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const lastSent = props.getProperty("AUTO_EMAIL_DATE");

  if (lastSent === today) {
    throw new Error("Auto email already sent today");
  }

  // Check quota
  const quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) {
    throw new Error("Email quota finished");
  }

  // Save today
  props.setProperty("AUTO_EMAIL_DATE", today);
}


function buildQCDashboard(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qcSheet = ss.getSheetByName("QC(Prem)");
  const fileSheet = ss.getSheetByName("File Complited");

  if(!qcSheet || !fileSheet) return;

  const qcLast = qcSheet.getLastRow();
  const fileLast = fileSheet.getLastRow();

  let dash = ss.getSheetByName("QC Dashboard");

  if(!dash){
    dash = ss.insertSheet("QC Dashboard");
  }else{
    dash.clear();
    dash.getCharts().forEach(c => dash.removeChart(c));
  }

  // ================= COLUMN WIDTH =================
  dash.setColumnWidth(1, 260);
  dash.setColumnWidth(2, 160);
  dash.setColumnWidth(3, 160);
  dash.setColumnWidth(5, 200);
  dash.setColumnWidth(6, 550); // big space for charts

  // ================= TITLE =================
  dash.getRange("A1:C1").merge()
    .setValue("QC ANALYTICS DASHBOARD")
    .setFontSize(24)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1F4E78");

  // ================= QC DATA =================
    const qcData = qcSheet.getRange("A2:I"+qcLast).getValues();

    let totalQCFiles = 0;
    let totalQCTime = 0;
    const dailyQC = {};

    qcData.forEach(r=>{
      const id = r[0];
      const time = parseFloat(r[3]) || 0;
      const date = r[4];
      const approved = r[8]; // Column I

      if(!id || !date) return;

      // Count only approved QC
      if(
        approved !== "Accepted With Minor Changes" &&
        approved !== "Accepted With Major Changes" &&
        approved !== "Accepted"
      ) return;

      totalQCFiles++;
      totalQCTime += time;

      const key = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(),"yyyy-MM-dd");

      if(!dailyQC[key]){
        dailyQC[key] = {files:0,time:0};
      }

      dailyQC[key].files++;
      dailyQC[key].time += time;
    });

  // ================= TRANSCRIPTION =================
  const fileData = fileSheet.getRange("F2:F"+fileLast).getValues().flat();
  const transIDs = fileData.filter(id => id && id !== "");
  const uniqueTrans = [...new Set(transIDs)];

  const totalTransFiles = uniqueTrans.length;
  const pending = totalTransFiles - totalQCFiles;

  // ================= TIME =================
  function convertToHMS(min){
    const h = Math.floor(min / 60);
    const m = Math.floor(min % 60);
    const s = Math.floor((min - Math.floor(min)) * 60);
    return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
  }

  const totalTimeHMS = convertToHMS(totalQCTime);
  const avgTime = totalQCFiles === 0 ? 0 : totalQCTime / totalQCFiles;
  const avgTimeHMS = convertToHMS(avgTime);

  // ================= SUMMARY =================
  dash.getRange("A3")
    .setValue("QC SUMMARY")
    .setFontWeight("bold")
    .setBackground("#FFD966");

  const summaryData = [
    ["Total QC Files", totalQCFiles],
    ["Total QC Time (Min)", totalQCTime.toFixed(2)],
    ["Total QC Time (H:M:S)", totalTimeHMS],
    ["Avg QC Time per File (Min)", avgTime.toFixed(2)],
    ["Avg QC Time per File (H:M:S)", avgTimeHMS],
    ["Total Transcription Files", totalTransFiles],
    ["Pending QC Work", pending]
  ];

  dash.getRange(5,1,summaryData.length,2).setValues(summaryData);

  dash.getRange("A5:A11").setFontWeight("bold").setBackground("#E8F0FE");
  dash.getRange("B5:B11").setBackground("#FCE5CD");
  dash.getRange("A5:B11").setBorder(true,true,true,true,true,true);

  // ================= TYPE TABLE =================
  dash.getRange("A13:B15").setValues([
    ["Type","Count"],
    ["Completed (QC)", totalQCFiles],
    ["Pending", pending]
  ]);

  dash.getRange("A13:B13")
    .setBackground("#1F4E78")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  dash.getRange("A14:B15")
    .setBackground("#F4CCCC");

  dash.getRange("A13:B15")
    .setBorder(true,true,true,true,true,true);

  // ================= DAILY TABLE =================
  dash.getRange("A18:C18")
    .setValues([["Date","QC Files","QC Time (Min)"]])
    .setBackground("#00ACC1")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  const rows = [];

  Object.keys(dailyQC)
    .sort((a,b)=> new Date(b)-new Date(a))
    .forEach(d=>{
      rows.push([
        Utilities.formatDate(new Date(d), Session.getScriptTimeZone(),"dd MMM yyyy"),
        dailyQC[d].files,
        dailyQC[d].time
      ]);
    });

  if(rows.length){
    dash.getRange(19,1,rows.length,3).setValues(rows);
  }

  dash.getRange("A19:C"+(rows.length+18))
    .setBorder(true,true,true,true,true,true);

  // ================= CHART 1 =================
  const chart1 = dash.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dash.getRange("A18:C"+(rows.length+18)))
    .setPosition(2,6,0,0)   // top right
    .setOption("title","QC Daily Output")
    .build();

  dash.insertChart(chart1);

  // ================= CHART 2 =================
  const chart2 = dash.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(dash.getRange("A13:B15"))
    .setPosition(20,6,0,0)  // ALWAYS BELOW chart 1
    .setOption("title","QC vs Pending Work")
    .build();

  dash.insertChart(chart2);

}


function sendDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileSheet = ss.getSheetByName("File Complited");
  const qcSheet = ss.getSheetByName("QC(Prem)");

  const tz = Session.getScriptTimeZone();

  // ===== YESTERDAY DATE =====
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);

  const targetDate = Utilities.formatDate(yesterday, tz, "yyyy-MM-dd");
  const displayDate = Utilities.formatDate(yesterday, tz, "dd MMM yyyy");

  // ================= TRANSCRIPTION =================
  const fileData = fileSheet.getRange("B2:H" + fileSheet.getLastRow()).getValues();

  let totalFiles = 0;
  let totalTime = 0;
  const members = {};

  fileData.forEach(r => {
    const name = r[0];
    const date = r[3];
    const time = parseFloat(r[6]) || 0;

    if (!date) return;

    const rowDate = Utilities.formatDate(new Date(date), tz, "yyyy-MM-dd");

    if (rowDate === targetDate) {
      totalFiles++;
      totalTime += time;

      if (!members[name]) members[name] = 0;
      members[name] += time;
    }
  });

  // ================= QC DATA =================
  const qcData = qcSheet.getRange("A2:I" + qcSheet.getLastRow()).getValues();

  let qcFiles = 0;
  let qcTime = 0;

  qcData.forEach(r => {
    const date = r[4];
    const time = parseFloat(r[3]) || 0;
    const approved = r[8];

    if(!date) return;

    const rowDate = Utilities.formatDate(new Date(date), tz, "yyyy-MM-dd");

    if(
      rowDate === targetDate &&
      (
        approved === "Accepted With Minor Changes" ||
        approved === "Accepted With Major Changes" ||
        approved === "Accepted"
      )
    ){
      qcFiles++;
      qcTime += time;
    }
  });

  // ================= TOP PERFORMER =================
  let topName = "N/A";
  let topTime = 0;

  Object.keys(members).forEach(name => {
    if (members[name] > topTime) {
      topTime = members[name];
      topName = name;
    }
  });

  // Sort members by time (highest to lowest)
  const sortedMembers = Object.keys(members).sort((a, b) => members[b] - members[a]);

  // ================= MEMBER TABLE ROWS - WITH PROPER SPACING AND COLON =================
  const memberRows = sortedMembers.map(name => `
    <div style="display: flex; justify-content: space-between; align-items: center; padding: 14px 0; border-bottom: 1px solid #E9EDF2;">
      <div style="font-size: 15px; font-weight: 500; color: #2C3E50;">${name}</div>
      <div style="font-size: 16px; font-weight: 700; color: #1F4E78;">
        ${members[name].toFixed(2)} <span style="font-size: 12px; font-weight: normal;">minutes</span>
      </div>
    </div>
  `).join('');

  // If no data for the day, show a message
  const noDataMessage = Object.keys(members).length === 0 ? `
    <div style="text-align: center; padding: 40px 20px; color: #8DA3BB; font-size: 14px;">
      No transcription data available for this date
    </div>
  ` : memberRows;

  // ================= PROFESSIONAL SINGLE-COLUMN HTML =================
  const htmlBody = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=yes">
    <title>Daily Production & QC Report</title>
    <style>
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }
      body {
        background-color: #F0F2F5;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', Helvetica, Arial, sans-serif;
        line-height: 1.5;
        margin: 0;
        padding: 16px;
        -webkit-text-size-adjust: 100%;
      }
      .email-container {
        max-width: 550px;
        width: 100%;
        margin: 0 auto;
        background: #ffffff;
        border-radius: 28px;
        overflow: hidden;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08);
      }
      /* Mobile Optimizations */
      @media only screen and (max-width: 480px) {
        body {
          padding: 8px !important;
        }
        .email-container {
          border-radius: 20px !important;
        }
        h1 {
          font-size: 20px !important;
        }
        .stat-card {
          padding: 18px 16px !important;
        }
        .stat-value {
          font-size: 32px !important;
        }
        .top-performer-name {
          font-size: 22px !important;
        }
        .member-item {
          padding: 12px 0 !important;
        }
        .member-name {
          font-size: 14px !important;
        }
        .member-time {
          font-size: 15px !important;
        }
        .footer-section {
          padding: 20px 16px !important;
        }
      }
      a {
        text-decoration: none;
        color: #1F4E78;
      }
      .stat-card {
        background: linear-gradient(135deg, #F8FAFE 0%, #F2F6FC 100%);
        border-radius: 20px;
        padding: 20px;
        border: 1px solid #E9EDF2;
      }
    </style>
  </head>
  <body style="background:#F0F2F5; margin:0; padding:16px;">
    <div class="email-container" style="max-width:550px; width:100%; margin:0 auto; background:#fff; border-radius:28px; overflow:hidden;">
      
      <!-- HEADER -->
      <div style="background: linear-gradient(135deg, #0B2B40 0%, #1F4E78 100%); padding: 32px 24px; text-align: center;">
        <div style="margin-bottom: 12px;">
          <span style="font-size: 44px;">📊</span>
        </div>
        <h1 style="font-size: 24px; font-weight: 700; margin: 0 0 8px; color: #ffffff; line-height: 1.3;">DAILY PRODUCTION & QC REPORT</h1>
        <div style="opacity: 0.95; font-size: 15px; font-weight: 500; background: rgba(255,255,255,0.2); display: inline-block; padding: 6px 18px; border-radius: 40px; margin-top: 8px;">
          📅 ${displayDate}
        </div>
      </div>

      <!-- TOP PERFORMER -->
      <div style="background: linear-gradient(135deg, #FFF9E8 0%, #FFF3E0 100%); padding: 28px 20px; text-align: center;">
        <div style="display: inline-block; background: #F6AE1C; border-radius: 60px; padding: 6px 18px; font-weight: 700; font-size: 13px; color: #2C2C2C; margin-bottom: 12px;">
          🏆 CHAMPION OF THE DAY
        </div>
        <h2 style="font-size: 26px; font-weight: 800; margin: 12px 0 8px; color: #1E3A4D; word-break: break-word;">${topName}</h2>
        <div style="display: inline-block; background: #ffffff; border-radius: 40px; padding: 8px 20px; margin-top: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.05);">
          <span style="font-weight: 700; font-size: 17px; color:#1F4E78;">⏱️ ${topTime.toFixed(2)} minutes</span>
        </div>
      </div>

      <!-- STATISTICS SECTION - SINGLE COLUMN -->
      <div style="padding: 24px 20px;">
        
        <!-- Transcription Card -->
        <div class="stat-card" style="margin-bottom: 16px;">
          <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 16px;">
            <span style="font-size: 32px;">📝</span>
            <h3 style="font-size: 18px; font-weight: 700; color: #1F4E78; margin: 0;">Transcription</h3>
          </div>
          <div style="display: flex; justify-content: space-between; align-items: baseline; flex-wrap: wrap; gap: 12px;">
            <div>
              <div class="stat-value" style="font-size: 36px; font-weight: 800; color: #1E4663;">${totalFiles}</div>
              <div style="color: #5A6E7C; font-size: 13px; margin-top: 4px;">Files Completed</div>
            </div>
            <div style="text-align: right;">
              <div style="font-size: 28px; font-weight: 700; color: #2A6B4E;">${totalTime.toFixed(2)} <span style="font-size: 14px;">min</span></div>
              <div style="color: #5A6E7C; font-size: 13px;">Total Time</div>
            </div>
          </div>
        </div>
        
        <!-- Quality Check Card -->
        <div class="stat-card" style="margin-bottom: 0;">
          <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 16px;">
            <span style="font-size: 32px;">✅</span>
            <h3 style="font-size: 18px; font-weight: 700; color: #1F4E78; margin: 0;">Quality Check</h3>
          </div>
          <div style="display: flex; justify-content: space-between; align-items: baseline; flex-wrap: wrap; gap: 12px;">
            <div>
              <div class="stat-value" style="font-size: 36px; font-weight: 800; color: #1E4663;">${qcFiles}</div>
              <div style="color: #5A6E7C; font-size: 13px; margin-top: 4px;">Files Completed</div>
            </div>
            <div style="text-align: right;">
              <div style="font-size: 28px; font-weight: 700; color: #2A6B4E;">${qcTime.toFixed(2)} <span style="font-size: 14px;">min</span></div>
              <div style="color: #5A6E7C; font-size: 13px;">QC Time</div>
            </div>
          </div>
        </div>
      </div>

      <!-- ANNOTATOR PERFORMANCE - SIMPLE VERTICAL LIST -->
      <div style="padding: 8px 20px 28px 20px;">
        <div style="border-bottom: 2px solid #EFF3F8; margin-bottom: 20px; padding-bottom: 12px;">
          <div style="display: flex; align-items: baseline; justify-content: space-between;">
            <h3 style="font-size: 19px; font-weight: 700; color: #1F4E78; margin: 0;">👨‍💻 Annotator Performance</h3>
            <span style="font-size: 12px; color: #6F8FAC;">Time (minutes)</span>
          </div>
        </div>
        
        <div style="background: #ffffff; border-radius: 20px;">
          ${noDataMessage}
        </div>
        
        ${Object.keys(members).length > 0 ? `
        <div style="margin-top: 16px; padding: 12px; background: #F8FAFE; border-radius: 16px; text-align: center;">
          <div style="font-size: 12px; color: #6F8FAC;">
            📊 Total Team Members: ${Object.keys(members).length} | Total Minutes: ${totalTime.toFixed(2)}
          </div>
        </div>
        ` : ''}
      </div>

      <!-- FOOTER - PROFESSIONAL -->
      <div style="background: #FBFDFF; border-top: 1px solid #E6EDF4; padding: 28px 24px 24px 24px;">
        
        <!-- Vendor Section -->
        <div style="background: linear-gradient(135deg, #FFF5E8 0%, #FFEFE0 100%); border-radius: 20px; padding: 18px; margin-bottom: 20px; text-align: center;">
          <div style="font-weight: 700; font-size: 14px; color: #D45A1F; margin-bottom: 8px; letter-spacing: 0.5px;">VENDOR PARTNER</div>
          <div style="font-weight: 700; color: #2C3E50; font-size: 16px;">Prem Prasad Pradhan</div>
          <div style="font-size: 13px; color: #4F6F8F; margin-top: 8px;">
            📞 +91 98277 75230
          </div>
          <div style="margin-top: 6px;">
            <a href="https://mrprem.in/" target="_blank" style="color: #1F6392; text-decoration: none; font-weight: 600; font-size: 13px;">🌐 mrprem.in</a>
          </div>
        </div>
        
        <!-- Company Section -->
        <div style="text-align: center; margin-bottom: 24px;">
          <div style="font-weight: 700; font-size: 15px; color: #0A3146;">DesiCrew Solutions Pvt. Ltd.</div>
          <div style="font-size: 12px; color: #346A8C; margin: 6px 0;">
            <a href="https://desicrew.in/" target="_blank" style="color: #1F6392; text-decoration: none; font-weight: 500;">desicrew.in</a>
          </div>
          <div style="font-size: 11px; color: #5E7C9A; margin-top: 4px;">
            Empowering rural and urban India together.
          </div>
        </div>

        <!-- Thank You Note -->
        <div style="text-align: center; margin-bottom: 20px; padding: 12px 0; border-top: 1px solid #E9EDF2; border-bottom: 1px solid #E9EDF2;">
          <span style="font-size: 28px;">🙌</span>
          <p style="font-size: 13px; color: #2C4C6C; font-weight: 500; margin: 8px 0 4px;">Thank you for your dedication & precision</p>
          <p style="font-size: 11px; color: #5E7C9A;">Every minute counts toward project excellence</p>
        </div>

        <!-- System Info -->
        <div style="text-align: center; font-size: 10px; color: #8DA3BB; line-height: 1.5;">
          <div>📊 Automated Daily Report System v2.0</div>
          <div style="margin-top: 8px;">
            <a href="https://docs.google.com/forms/d/e/1FAIpQLScIzlm1YLwsnGq9yryrz6_ZYoqSSKJPkm9aVwP3YjoK8c_Tvg/viewform?usp=publish-editor" style="color: #1F4E78; text-decoration: underline;">🔔 Unsubscribe from daily report</a>
          </div>
          <div style="margin-top: 8px; font-size: 9px; color: #B0C4DE;">
            © 2026 MR.PREM — All Rights Reserved
          </div>
        </div>
      </div>
    </div>
  </body>
  </html>
  `;

  // ================= SEND EMAIL =================
  // ✅ NEW — Sends ONE email, BCC all, costs only 1 quota
  const recipients = getEmailsFromSheet();

  MailApp.sendEmail({
    to: "qcds-team-reports@googlegroups.com",
    subject: `📊 Daily Report - ${displayDate}`,
    htmlBody: htmlBody,
    name: "Prem Production Report System"
  });
}

function autoSendExact() {
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();

  const now = new Date();
  const today = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  const hour = now.getHours();
  const minute = now.getMinutes();

  const lastSent = props.getProperty("AUTO_EMAIL_DATE");

  // ✅ Stop if already sent today
  if (lastSent === today) {
    Logger.log("⏭ Already sent today, skipping...");
    return;
  }

  // ✅ Safe time window (12:00 AM → 12:30 AM)
  if (hour === 0 && minute < 30) {
    try {

      Logger.log("🚀 Attempting to send daily report...");

      sendDailyReport();  // 🔥 your main function

      // ✅ Mark as sent ONLY after success
      props.setProperty("AUTO_EMAIL_DATE", today);

      Logger.log("✅ Daily email sent successfully for " + today);

    } catch (err) {

      Logger.log("❌ Email sending failed: " + err);

      // ❗ Do NOT mark as sent → so it retries in next trigger
    }

  } else {
    Logger.log("⏱ Not in sending window");
  }
}



function cleanName(name){
  name = name.trim();
  name = name.replace(/\s+/g, " ");
  name = name.toLowerCase().replace(/\b\w/g, l => l.toUpperCase());
  return name;
}

function submitFileData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  // Clean name properly
  const name = cleanName(data.name);

  // Find first empty row
  const names = sheet.getRange("B2:B").getValues();
  let nextRow = 2;

  for (let i = 0; i < names.length; i++) {
    if (!names[i][0]) {
      nextRow = i + 2;
      break;
    }
  }

  // Check duplicate ID
  const lastRow = sheet.getLastRow();
  const idColumn = sheet.getRange("F2:F" + lastRow).getValues().flat().map(String);

  if (idColumn.includes(String(data.id))) {
    return "ERROR: ID already exists!";
  }

  // Sl No
  const slno = nextRow - 1;

  // Date
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "M/d/yyyy");

  // Convert seconds to minutes
  const minutes = (Number(data.time) / 60).toFixed(2);

  // Insert data
  sheet.getRange(nextRow, 1, 1, 8).setValues([[
    slno,
    name,
    "QC",
    "Completed",
    formattedDate,
    Number(data.id),
    Number(data.time),
    Number(minutes)
  ]]);

  return "SUCCESS";
}

function doGet(e) {
  if (e.parameter.page == "admin") {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle("Admin Panel");
  } 
  else if (e.parameter.page == "dashboard") {
    return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle("Analytics Dashboard");
  }
  else if (e.parameter.page == "email") {   // ✅ ADD THIS
    return HtmlService.createHtmlOutputFromFile('emailsystem')
      .setTitle("Email System");
  }
  else if (e.parameter.page === "report") {
    return HtmlService.createHtmlOutputFromFile('memberreport')
      .setTitle("My Work Report · QCDS");
  }
  else {
    return HtmlService.createHtmlOutputFromFile('form')
      .setTitle("QC File Submission");
  }
}

function fixAllNamesInSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange("B2:B" + lastRow);
  const names = range.getValues();

  const cleaned = names.map(row => {
    if (!row[0]) return [""];
    
    let name = row[0].toString().trim();
    name = name.replace(/\s+/g, " ");
    name = name.toLowerCase().replace(/\b\w/g, l => l.toUpperCase());

    return [name];
  });

  range.setValues(cleaned);
}

function getEmailsFromSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Team Details");

  const emails = sheet.getRange("C2:C" + sheet.getLastRow())
    .getValues()
    .flat()
    .filter(String);

  return emails;
}

function checkQuota(){
  const quota = MailApp.getRemainingDailyQuota();
  const now = new Date();

  Logger.log("⏰ Time: " + now);
  Logger.log("📧 Remaining quota: " + quota);
}

// ✅ NEW — Single email, BCC all
function sendCustomEmail(subject, body){

  MailApp.sendEmail({
    to: "qcds-team-reports@googlegroups.com",
    subject: subject,
    htmlBody: body,
    name: "Prem Production Report System"
  });

}

function getSheetUrl(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getUrl();
}

function getFullDashboardData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const summarySheet = ss.getSheetByName("Summary");
  const prodSheet = ss.getSheetByName("Daily Production");
  const memberSheet = ss.getSheetByName("Member Daily");

  if (!summarySheet) {
    return {
      totalFiles: 0,
      totalTime: 0,
      totalTimeHMS: "00:00:00",
      topPerformers: [],
      dailyProduction: [],
      memberDaily: []
    };
  }

  const summary = summarySheet.getRange("A2:D" + summarySheet.getLastRow()).getValues();

  let totalFiles = 0;
  let totalTime = 0;

  summary.forEach(r => {
    totalFiles += Number(r[2] || 0);
    totalTime += Number(r[3] || 0);
  });

  const h = Math.floor(totalTime / 60);
  const m = Math.floor(totalTime % 60);
  const s = Math.floor((totalTime - Math.floor(totalTime)) * 60);

  const totalTimeHMS =
    String(h).padStart(2, '0') + ":" +
    String(m).padStart(2, '0') + ":" +
    String(s).padStart(2, '0');

  summary.sort((a,b)=>b[3]-a[3]);
  const topPerformers = summary.slice(0,15);

  let prodData = [];
  if(prodSheet && prodSheet.getLastRow() > 3){
    prodData = prodSheet.getRange("A4:D" + prodSheet.getLastRow()).getValues();
  }

  let memberData = [];
  if(memberSheet){
    memberData = memberSheet.getDataRange().getValues();
  }

  return {
    totalFiles,
    totalTime: totalTime.toFixed(2),
    totalTimeHMS,
    topPerformers,
    dailyProduction: prodData,
    memberDaily: memberData
  };
}

function buildSummary(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileSheet = ss.getSheetByName("File Complited");

  const data = fileSheet.getRange("B2:H" + fileSheet.getLastRow()).getValues();

  const members = {};

  data.forEach(r=>{
    const name = r[0];
    const time = parseFloat(r[6]) || 0;

    if(!name) return;

    if(!members[name]){
      members[name] = {files:0,time:0};
    }

    members[name].files++;
    members[name].time += time;
  });

  let summary = ss.getSheetByName("Summary");

  if(!summary){
    summary = ss.insertSheet("Summary");
  }else{
    summary.clear();
  }

  summary.getRange("A1:D1")
    .setValues([["Name","Assigned","Completed Files","Total Time"]])
    .setFontWeight("bold");

  const rows = [];

  Object.keys(members).forEach(name=>{
    rows.push([
      name,
      members[name].files,
      members[name].files,
      members[name].time
    ]);
  });

  if(rows.length){
    summary.getRange(2,1,rows.length,4).setValues(rows);
  }
}

function buildEverything(){
  buildSummary();
  buildDailyProduction();
  buildAnnotatorDaily();
  buildAdvancedDashboard();
}

function getTodayProduction() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2,1,lastRow-1,8).getValues();

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");

  let files = 0;
  let time = 0;
  let members = new Set();

  let dayMap = {};

  data.forEach(row => {
    const name = row[1];
    const date = Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "M/d/yyyy");
    const minutes = Number(row[7]);

    // Today
    if(date === today){
      files++;
      time += minutes;
      members.add(name);
    }

    // Last days
    if(!dayMap[date]){
      dayMap[date] = {files:0, time:0, members: new Set()};
    }

    dayMap[date].files++;
    dayMap[date].time += minutes;
    if(name) dayMap[date].members.add(name);
  });

  const last3 = Object.keys(dayMap)
    .sort((a,b)=> new Date(b) - new Date(a))
    .slice(0,3)
    .map(d => [d, dayMap[d].files, dayMap[d].time, dayMap[d].members.size]);

  return {
    today: [
      today,
      files,
      time,
      members.size
    ],
    last3: last3
  };
}

function getTodayMemberOutput() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2,1,lastRow-1,8).getValues();

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");

  const members = {};

  data.forEach(row => {
    const name = row[1];
    const date = Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "M/d/yyyy");
    const minutes = Number(row[7]) || 0;

    if(date === today){
      if(!members[name]) members[name] = 0;
      members[name] += minutes;
    }
  });

  const result = [];

  Object.keys(members).forEach(name=>{
    result.push([name, members[name]]);
  });

  // Sort highest time first
  result.sort((a,b)=> b[1] - a[1]);

  return result;
}

function getTeamContribution(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  const lastRow = sheet.getLastRow();
  if(lastRow < 2) return [];

  const data = sheet.getRange(2,1,lastRow-1,8).getValues();

  const members = {};

  data.forEach(row=>{
    const name = row[1];
    const minutes = Number(row[7]) || 0;

    if(!name) return;

    if(!members[name]) members[name] = 0;
    members[name] += minutes;
  });

  const result = [];

  Object.keys(members).forEach(name=>{
    result.push([name, members[name]]);
  });

  return result;
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");
  const data = sheet.getDataRange().getValues();

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  let files = 0;
  let time = 0;
  let members = {};

  for (let i = 1; i < data.length; i++) {
    let name = data[i][1];
    let date = data[i][4];
    let min = parseFloat(data[i][7]) || 0;

    if (!date) continue;

    let rowDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");

    if (rowDate === today) {
      files++;
      time += min;

      if (!members[name]) members[name] = 0;
      members[name] += min;
    }
  }

  let memberArr = [];
  for (let m in members) {
    memberArr.push({
      name: m,
      initial: m.split(" ").map(x=>x[0]).join(""),
      time: members[m].toFixed(2)
    });
  }

  return {
    date: today,
    files: files,
    time: time.toFixed(2),
    head: memberArr.length,
    last3: [],
    members: memberArr
  };
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");
  const data = sheet.getDataRange().getValues();

  const rows = data.slice(1);
  const tz = Session.getScriptTimeZone();

  const todayStr = Utilities.formatDate(new Date(), tz, "dd MMM yyyy");
  const yesterdayStr = Utilities.formatDate(new Date(Date.now()-86400000), tz, "dd MMM yyyy");

  let production = {};
  let memberDaily = {};
  let memberTotal = {};
  let todayFiles = 0;
  let todayTime = 0;

  // Live Activity Feed (last 10 entries)
  let liveFeed = [];

  rows.forEach(r => {
    let name = r[1];
    let date = Utilities.formatDate(new Date(r[4]), tz, "dd MMM yyyy");
    let time = parseFloat(r[7]) || 0;

    // Production
    if (!production[date]) production[date] = {files:0, time:0, members:new Set()};
    production[date].files++;
    production[date].time += time;
    production[date].members.add(name);

    // Member daily
    if (!memberDaily[name]) memberDaily[name] = {};
    if (!memberDaily[name][date]) memberDaily[name][date] = 0;
    memberDaily[name][date] += time;

    // Member total
    if (!memberTotal[name]) memberTotal[name] = 0;
    memberTotal[name] += time;

    // Today
    if (date === todayStr) {
      todayFiles++;
      todayTime += time;
    }

    // Live activity
    liveFeed.push({
      name: name,
      time: time.toFixed(2),
      date: date
    });
  });

  // Sort dates properly
  let lastDays = Object.keys(production)
    .sort((a,b)=> new Date(b) - new Date(a))
    .slice(0,4);

  let report = lastDays.map(d => ({
    date:d,
    files:production[d].files,
    time:production[d].time.toFixed(2),
    head:production[d].members.size
  }));

  // Member Today vs Yesterday
  let members = [];
  Object.keys(memberTotal).forEach(name=>{
    members.push({
      name:name,
      today:(memberDaily[name][todayStr] || 0).toFixed(2),
      yesterday:(memberDaily[name][yesterdayStr] || 0).toFixed(2)
    });
  });

  // Top performers
  let top = [];
  Object.keys(memberTotal).forEach(name=>{
    top.push({
      name:name,
      total:memberTotal[name].toFixed(2)
    });
  });

  top.sort((a,b)=>b.total-a.total);

  // Live feed last 10
  liveFeed = liveFeed.reverse().slice(0,10);

  return {
    todayFiles,
    todayTime:todayTime.toFixed(2),
    report,
    members,
    top,
    liveFeed,
    teamFiles: rows.length,
    teamTime: Object.values(memberTotal).reduce((a,b)=>a+b,0).toFixed(2)
  };
}

function submitSkipData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Skip Files");
  const name = cleanName(data.name);
  const lastRow = sheet.getLastRow();
  const nextRow = lastRow + 1;
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "M/d/yyyy");
  const minutes = (Number(data.time) / 60).toFixed(2);
  sheet.getRange(nextRow, 1, 1, 9).setValues([[   // ← 8 changed to 9
    nextRow - 1,
    name,
    "QC",
    "Skipped",
    formattedDate,
    Number(data.id),
    Number(data.time),
    Number(minutes),
    data.reason || ""                              // ← this line added
  ]]);
  return "SUCCESS";
}

function verifyPassword(input) {
  const correct = PropertiesService.getScriptProperties().getProperty("ADMIN_PASS");
  return input === correct;
}

function checkEmailStatus() {
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  const lastSent = props.getProperty("AUTO_EMAIL_DATE");
  const sentToday = (lastSent === today);

  const quota = MailApp.getRemainingDailyQuota();

  return {
    quota: quota,
    sentToday: sentToday
  };
}

// ▶ Run this ONE TIME manually from Apps Script editor to set your password
function setPassword() {
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASS", "Prem123");
  Logger.log("✅ Password set successfully!");
}

function resetPassword(newPass) {
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASS", newPass);
  return "✅ Password updated to: " + newPass;
}

function getMembersForEmailSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Team Details");

  if (!sheet) return [];

  const data = sheet.getRange("B2:C" + sheet.getLastRow()).getValues();

  return data
    .filter(row => row[0] && row[1])
    .map(row => ({
      name: row[0],
      email: row[1]
    }));
}

function sendEmailSystemBatch(subject, htmlBody, emails) {

  try {
    const BATCH_SIZE = 45;
    let batches = [];

    for (let i = 0; i < emails.length; i += BATCH_SIZE) {
      batches.push(emails.slice(i, i + BATCH_SIZE));
    }

    let results = [];

    for (let i = 0; i < batches.length; i++) {
      let batch = batches[i];

      try {
        MailApp.sendEmail({
          to: Session.getActiveUser().getEmail(), // required dummy
          bcc: batch.join(","),
          subject: subject,
          htmlBody: htmlBody,
          name: "Prem Email System"
        });

        results.push({
          batch: i + 1,
          count: batch.length,
          success: true
        });

        Utilities.sleep(2000); // delay to avoid quota error

      } catch (e) {
        results.push({
          batch: i + 1,
          count: batch.length,
          success: false
        });
      }
    }

    return {
      success: true,
      totalBatches: batches.length,
      batches: results
    };

  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

function clearFlagAndSend() {

  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  // ❌ Prevent multiple send
  const lastSent = props.getProperty("AUTO_EMAIL_DATE");

  if (lastSent === today) {
    Logger.log("⛔ Already sent today!");
    return;
  }

  // Step 1: Check quota FIRST
  const quota = MailApp.getRemainingDailyQuota();
  Logger.log("📧 Quota available: " + quota);

  if (quota <= 0) {
    Logger.log("❌ No quota left — try tomorrow");
    return;
  }

  try {
    sendDailyReport();

    // ✅ mark after success
    props.setProperty("AUTO_EMAIL_DATE", today);

    Logger.log("✅ Daily report sent!");
  } catch (err) {
    Logger.log("❌ Error: " + err);
  }
}

// ═══════════════════════════════════════════════════════════════════════
// QCDS EMAIL SYSTEM — COMPLETE BACKEND (Code.gs additions)
// QuantumCoders Data Solutions — Professional Email System v2.0
// ═══════════════════════════════════════════════════════════════════════
//
// SETUP INSTRUCTIONS:
// 1. Add this code to your existing Code.gs (or paste below existing code)
// 2. Create a sheet named "Email History" (auto-created on first send)
// 3. Add 'emailsystem' page in doGet() — already handled below
// 4. Ensure "Team Details" sheet has: Col B = Name, Col C = Email
// 5. Run setPassword() once to set admin password
// ═══════════════════════════════════════════════════════════════════════


// ────────────────────────────────────────────────────────────────────────
// CONSTANTS
// ────────────────────────────────────────────────────────────────────────
const GOOGLE_GROUP_EMAIL = "qcds-team-reports@googlegroups.com";
const EMAIL_FROM_NAME_PREM = "Prem Prasad Pradhan · QuantumCoders Data Solutions";
const EMAIL_FROM_NAME_OFFICIAL = "QuantumCoders Data Solutions";
const EMAIL_HISTORY_SHEET = "Email History";
const BATCH_SIZE = 45;


// ────────────────────────────────────────────────────────────────────────
// doGet — ADD THE EMAIL PAGE ROUTE
// Replace your existing doGet with this version (or merge the email case)
// ────────────────────────────────────────────────────────────────────────
/*
function doGet(e) {
  const page = e.parameter.page;

  if (page === "admin") {
    return HtmlService.createHtmlOutputFromFile('admin').setTitle("Admin Panel");
  }
  if (page === "dashboard") {
    return HtmlService.createHtmlOutputFromFile('dashboard').setTitle("Analytics Dashboard");
  }
  if (page === "email") {
    return HtmlService.createHtmlOutputFromFile('emailsystem').setTitle("Email System · QCDS");
  }
  return HtmlService.createHtmlOutputFromFile('form').setTitle("QC File Submission");
}
*/


// ────────────────────────────────────────────────────────────────────────
// 1. SEND VIA GOOGLE GROUP — Single email, 1 quota, group distributes
// ────────────────────────────────────────────────────────────────────────
function sendEmailViaGroup(subject, htmlBody, senderMode) {

  try {

    const fromName = (senderMode === 'prem')
      ? EMAIL_FROM_NAME_PREM
      : EMAIL_FROM_NAME_OFFICIAL;

    // One single email → Google Group → distributes to all members
    MailApp.sendEmail({
      to: GOOGLE_GROUP_EMAIL,
      subject: subject,
      htmlBody: htmlBody,
      name: fromName
    });

    // Log to history
    logEmailHistory(subject, 'Sent', 'group', senderMode);

    return {
      success: true,
      totalBatches: 1,
      batches: [{ batch: 1, count: 1, success: true }]
    };

  } catch (err) {

    logEmailHistory(subject, 'Failed', 'group', senderMode);

    return {
      success: false,
      error: err.toString()
    };
  }
}


// ────────────────────────────────────────────────────────────────────────
// 2. SEND INDIVIDUAL BCC BATCHES — multiple recipients, BCC protected
// ────────────────────────────────────────────────────────────────────────
function sendEmailSystemBatch(subject, htmlBody, emails, senderMode) {

  try {

    const fromName = (senderMode === 'prem')
      ? EMAIL_FROM_NAME_PREM
      : EMAIL_FROM_NAME_OFFICIAL;

    const batches = [];
    for (let i = 0; i < emails.length; i += BATCH_SIZE) {
      batches.push(emails.slice(i, i + BATCH_SIZE));
    }

    const results = [];

    for (let i = 0; i < batches.length; i++) {
      const batch = batches[i];

      try {
        MailApp.sendEmail({
          to: Session.getActiveUser().getEmail(), // required non-empty "to"
          bcc: batch.join(","),
          subject: subject,
          htmlBody: htmlBody,
          name: fromName
        });

        results.push({ batch: i + 1, count: batch.length, success: true });

        // Rate-limit protection: 2 second delay between batches
        if (i < batches.length - 1) {
          Utilities.sleep(2000);
        }

      } catch (batchErr) {
        results.push({ batch: i + 1, count: batch.length, success: false });
      }
    }

    // Log to history
    const allOk = results.every(r => r.success);
    logEmailHistory(subject, allOk ? 'Sent' : 'Failed', 'individual', senderMode);

    return {
      success: true,
      totalBatches: batches.length,
      batches: results
    };

  } catch (err) {

    logEmailHistory(subject, 'Failed', 'individual', senderMode);

    return {
      success: false,
      error: err.toString()
    };
  }
}


// ────────────────────────────────────────────────────────────────────────
// 3. EMAIL HISTORY — Log & retrieve
// ────────────────────────────────────────────────────────────────────────
function logEmailHistory(subject, status, mode, senderMode) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(EMAIL_HISTORY_SHEET);

  // Auto-create sheet if missing
  if (!sheet) {
    sheet = ss.insertSheet(EMAIL_HISTORY_SHEET);
    sheet.getRange("A1:5").setValues([["Date", "Subject", "Status", "Mode", "Sender"]]);
    sheet.getRange("A1:E1")
      .setFontWeight("bold")
      .setBackground("#1456A8")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 220);
  }

  const tz = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, "dd MMM yyyy, hh:mm a");

  sheet.appendRow([
    now,
    subject,
    status,
    mode || 'group',
    senderMode === 'prem' ? 'Prem Prasad Pradhan' : 'Official QCDS'
  ]);
}


function getEmailHistory() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EMAIL_HISTORY_SHEET);

  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  // Newest first, max 50 rows
  return data.reverse().slice(0, 50);
}


// ────────────────────────────────────────────────────────────────────────
// 4. MEMBERS FOR EMAIL SYSTEM
// ────────────────────────────────────────────────────────────────────────
function getMembersForEmailSystem() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Team Details");

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange("B2:C" + lastRow).getValues();

  return data
    .filter(row => row[0] && row[1])
    .map(row => ({
      name: row[0].toString().trim(),
      email: row[1].toString().trim().toLowerCase()
    }));
}


// ────────────────────────────────────────────────────────────────────────
// 5. QUOTA & STATUS CHECK
// ────────────────────────────────────────────────────────────────────────
function checkEmailStatus() {

  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const lastSent = props.getProperty("AUTO_EMAIL_DATE");

  const quota = MailApp.getRemainingDailyQuota();

  return {
    quota: quota,
    sentToday: (lastSent === today),
    date: today
  };
}


// ────────────────────────────────────────────────────────────────────────
// 6. PASSWORD VERIFICATION (unchanged — already in your code)
// ────────────────────────────────────────────────────────────────────────
// function verifyPassword(input) {
//   const correct = PropertiesService.getScriptProperties().getProperty("ADMIN_PASS");
//   return input === correct;
// }

// Run once to set password:
// function setPassword() { PropertiesService.getScriptProperties().setProperty("ADMIN_PASS", "YourPassword"); }


// ────────────────────────────────────────────────────────────────────────
// 7. ENHANCED DAILY REPORT — Sends via Google Group (1 quota)
// Replaces the old sendDailyReport() function
// ────────────────────────────────────────────────────────────────────────
function sendDailyReportViaGroup() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileSheet = ss.getSheetByName("File Complited");
  const qcSheet = ss.getSheetByName("QC(Prem)");

  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);

  const targetDate = Utilities.formatDate(yesterday, tz, "yyyy-MM-dd");
  const displayDate = Utilities.formatDate(yesterday, tz, "dd MMM yyyy");

  // ── TRANSCRIPTION DATA ──
  const fileData = fileSheet.getRange("B2:H" + fileSheet.getLastRow()).getValues();
  let totalFiles = 0, totalTime = 0;
  const members = {};

  fileData.forEach(r => {
    const name = r[0], date = r[3], time = parseFloat(r[6]) || 0;
    if (!date) return;
    const rowDate = Utilities.formatDate(new Date(date), tz, "yyyy-MM-dd");
    if (rowDate === targetDate) {
      totalFiles++;
      totalTime += time;
      if (!members[name]) members[name] = 0;
      members[name] += time;
    }
  });

  // ── QC DATA ──
  const qcData = qcSheet.getRange("A2:I" + qcSheet.getLastRow()).getValues();
  let qcFiles = 0, qcTime = 0;

  qcData.forEach(r => {
    const date = r[4], time = parseFloat(r[3]) || 0, approved = r[8];
    if (!date) return;
    const rowDate = Utilities.formatDate(new Date(date), tz, "yyyy-MM-dd");
    if (rowDate === targetDate && ["Accepted With Minor Changes","Accepted With Major Changes","Accepted"].includes(approved)) {
      qcFiles++;
      qcTime += time;
    }
  });

  // ── TOP PERFORMER ──
  let topName = "N/A", topTime = 0;
  Object.keys(members).forEach(n => { if (members[n] > topTime) { topTime = members[n]; topName = n; } });

  const sortedMembers = Object.keys(members).sort((a, b) => members[b] - members[a]);
  const memberRows = sortedMembers.map(name =>
    `<div style="display:flex;justify-content:space-between;align-items:center;padding:12px 0;border-bottom:1px solid #E9EDF2">
      <div style="font-size:15px;font-weight:500;color:#2C3E50">${name}</div>
      <div style="font-size:16px;font-weight:700;color:#1456A8">${members[name].toFixed(2)} <span style="font-size:12px;font-weight:normal">min</span></div>
    </div>`
  ).join('');

  // ── BUILD HTML ──
  const htmlBody = buildDailyReportHtml(displayDate, totalFiles, totalTime, qcFiles, qcTime, topName, topTime, memberRows, sortedMembers, members);

  // ── SEND VIA GOOGLE GROUP ──
  MailApp.sendEmail({
    to: GOOGLE_GROUP_EMAIL,
    subject: `📊 Daily Report - ${displayDate}`,
    htmlBody: htmlBody,
    name: EMAIL_FROM_NAME_PREM
  });

  logEmailHistory(`📊 Daily Report - ${displayDate}`, 'Sent', 'group', 'prem');
}


// ────────────────────────────────────────────────────────────────────────
// 8. DAILY REPORT HTML BUILDER
// ────────────────────────────────────────────────────────────────────────
function buildDailyReportHtml(displayDate, totalFiles, totalTime, qcFiles, qcTime, topName, topTime, memberRows, sortedMembers, members) {

  const noData = sortedMembers.length === 0
    ? `<div style="text-align:center;padding:36px 20px;color:#8DA3BB;font-size:14px">No transcription data for this date</div>`
    : memberRows;

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1.0">
  <title>Daily Production & QC Report</title>
</head>
<body style="margin:0;padding:16px;background:#F0F2F5;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif">
<div style="max-width:560px;width:100%;margin:0 auto;background:#ffffff;border-radius:24px;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,0.09)">

  <!-- HEADER with Logo -->
  <div style="background:linear-gradient(135deg,#071629 0%,#0b2040 35%,#1456A8 75%,#1a6bbf 100%);padding:30px 28px;text-align:center">
    <img src="https://www.quantumcoderstechlab.codes/assets/Data%20Solutions/Logo/QCDL.png"
      alt="QCDS Logo" style="height:46px;width:auto;margin-bottom:12px;display:block;margin-left:auto;margin-right:auto">
    <div style="font-size:13px;font-weight:800;color:#ffffff;letter-spacing:1.5px;text-transform:uppercase">QUANTUMCODERS DATA SOLUTIONS</div>
    <div style="font-size:9px;color:rgba(255,255,255,0.5);letter-spacing:2px;text-transform:uppercase;margin-top:3px">Official Communication</div>
    <h1 style="font-size:22px;font-weight:800;margin:14px 0 0;color:#ffffff">DAILY PRODUCTION & QC REPORT</h1>
    <div style="display:inline-block;background:rgba(255,255,255,0.15);padding:6px 18px;border-radius:40px;margin-top:10px;font-size:14px;font-weight:600;color:#fff">📅 ${displayDate}</div>
  </div>

  <!-- TOP PERFORMER -->
  <div style="background:linear-gradient(135deg,#FFF9E8,#FFF3E0);padding:26px 20px;text-align:center">
    <div style="display:inline-block;background:#F6AE1C;border-radius:50px;padding:6px 18px;font-weight:800;font-size:12px;color:#2C2C2C;margin-bottom:10px;letter-spacing:0.5px">🏆 CHAMPION OF THE DAY</div>
    <h2 style="font-size:24px;font-weight:800;margin:10px 0 8px;color:#1E3A4D">${topName}</h2>
    <div style="display:inline-block;background:#fff;border-radius:40px;padding:8px 20px;box-shadow:0 2px 8px rgba(0,0,0,0.05)">
      <span style="font-weight:700;font-size:16px;color:#1456A8">⏱️ ${topTime.toFixed(2)} minutes</span>
    </div>
  </div>

  <!-- STATS -->
  <div style="padding:22px 20px">
    <!-- Transcription -->
    <div style="background:linear-gradient(135deg,#F8FAFE,#F2F6FC);border-radius:18px;padding:18px;border:1px solid #E9EDF2;margin-bottom:14px">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:14px">
        <span style="font-size:28px">📝</span>
        <h3 style="font-size:17px;font-weight:700;color:#1456A8;margin:0">Transcription</h3>
      </div>
      <div style="display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;gap:10px">
        <div>
          <div style="font-size:34px;font-weight:800;color:#1E4663">${totalFiles}</div>
          <div style="color:#5A6E7C;font-size:12px;margin-top:3px">Files Completed</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:26px;font-weight:700;color:#2A6B4E">${totalTime.toFixed(2)} <span style="font-size:13px">min</span></div>
          <div style="color:#5A6E7C;font-size:12px">Total Time</div>
        </div>
      </div>
    </div>
    <!-- QC -->
    <div style="background:linear-gradient(135deg,#F8FAFE,#F2F6FC);border-radius:18px;padding:18px;border:1px solid #E9EDF2">
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:14px">
        <span style="font-size:28px">✅</span>
        <h3 style="font-size:17px;font-weight:700;color:#1456A8;margin:0">Quality Check</h3>
      </div>
      <div style="display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;gap:10px">
        <div>
          <div style="font-size:34px;font-weight:800;color:#1E4663">${qcFiles}</div>
          <div style="color:#5A6E7C;font-size:12px;margin-top:3px">Files QC'd</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:26px;font-weight:700;color:#2A6B4E">${qcTime.toFixed(2)} <span style="font-size:13px">min</span></div>
          <div style="color:#5A6E7C;font-size:12px">QC Time</div>
        </div>
      </div>
    </div>
  </div>

  <!-- ANNOTATOR PERFORMANCE -->
  <div style="padding:4px 20px 26px">
    <div style="border-bottom:2px solid #EFF3F8;margin-bottom:18px;padding-bottom:10px;display:flex;justify-content:space-between;align-items:baseline">
      <h3 style="font-size:18px;font-weight:700;color:#1456A8;margin:0">👨‍💻 Annotator Performance</h3>
      <span style="font-size:11px;color:#6F8FAC">Time (minutes)</span>
    </div>
    ${noData}
    ${sortedMembers.length > 0 ? `
    <div style="margin-top:14px;padding:10px;background:#F8FAFE;border-radius:14px;text-align:center;font-size:11px;color:#6F8FAC">
      📊 Team: ${sortedMembers.length} members · ${totalTime.toFixed(2)} total minutes
    </div>` : ''}
  </div>

  <!-- FOOTER -->
  <div style="background:#FBFDFF;border-top:1px solid #E6EDF4;padding:26px 22px 22px">
    <!-- Vendor -->
    <div style="background:linear-gradient(135deg,#FFF5E8,#FFEFE0);border-radius:18px;padding:16px;margin-bottom:18px;text-align:center">
      <div style="font-weight:800;font-size:13px;color:#D45A1F;margin-bottom:6px;letter-spacing:0.5px">VENDOR PARTNER</div>
      <div style="font-weight:700;color:#2C3E50;font-size:15px">Prem Prasad Pradhan</div>
      <div style="font-size:12px;color:#4F6F8F;margin-top:7px">📞 +91 98277 75230</div>
      <div style="margin-top:5px"><a href="https://mrprem.in/" style="color:#1456A8;text-decoration:none;font-weight:600;font-size:12px">🌐 mrprem.in</a></div>
    </div>
    <!-- Company -->
    <div style="text-align:center;margin-bottom:18px">
      <img src="https://www.quantumcoderstechlab.codes/assets/Data%20Solutions/Logo/QCDL.png"
        alt="QCDS" style="height:32px;width:auto;opacity:0.5;display:block;margin:0 auto 10px">
      <div style="font-weight:700;font-size:14px;color:#0A3146">QuantumCoders Data Solutions</div>
      <div style="font-size:11px;color:#64748b;margin-top:5px">Berhampur, Odisha, India 760001</div>
      <div style="margin-top:4px"><a href="https://www.quantumcoderstechlab.codes" style="color:#1456A8;text-decoration:none;font-weight:600;font-size:11px">www.quantumcoderstechlab.codes</a></div>
    </div>
    <!-- Thank you -->
    <div style="text-align:center;padding:12px 0;border-top:1px solid #E9EDF2;border-bottom:1px solid #E9EDF2;margin-bottom:14px">
      <span style="font-size:26px">🙌</span>
      <p style="font-size:12px;color:#2C4C6C;font-weight:500;margin:7px 0 3px">Thank you for your dedication & precision</p>
      <p style="font-size:10px;color:#5E7C9A">Every minute counts toward project excellence</p>
    </div>
    <!-- Legal -->
    <div style="text-align:center;font-size:10px;color:#8DA3BB;line-height:1.6">
      This is an official internal communication. Please do not reply to this email.<br>
      <a href="https://docs.google.com/forms/d/e/1FAIpQLScIzlm1YLwsnGq9yryrz6_ZYoqSSKJPkm9aVwP3YjoK8c_Tvg/viewform" style="color:#1456A8;text-decoration:underline">🔕 Unsubscribe from daily reports</a><br>
      <div style="margin-top:7px;font-size:9px;color:#CBD5E1">© 2026 QuantumCoders Data Solutions · All Rights Reserved</div>
    </div>
  </div>

</div>
</body>
</html>`;
}


// ────────────────────────────────────────────────────────────────────────
// 9. AUTO SEND — Updated to use Group mode (1 quota only)
// Replace autoSendExact() with this version
// ────────────────────────────────────────────────────────────────────────
function autoSendExactV2() {

  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const hour = now.getHours();
  const minute = now.getMinutes();
  const lastSent = props.getProperty("AUTO_EMAIL_DATE");

  if (lastSent === today) {
    Logger.log("⏭ Already sent today, skipping...");
    return;
  }

  // Send window: midnight to 12:30 AM
  if (hour === 0 && minute < 30) {
    try {
      Logger.log("🚀 Sending daily report via Google Group...");

      sendDailyReportViaGroup(); // ← Uses 1 quota via Google Group

      props.setProperty("AUTO_EMAIL_DATE", today);
      Logger.log("✅ Daily report sent for " + today);

    } catch (err) {
      Logger.log("❌ Send failed: " + err);
    }
  } else {
    Logger.log("⏱ Outside sending window (" + hour + ":" + minute + ")");
  }
}


// ────────────────────────────────────────────────────────────────────────
// 10. HELPER — Get emails from sheet (unchanged, kept for compatibility)
// ────────────────────────────────────────────────────────────────────────
// function getEmailsFromSheet() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName("Team Details");
//   return sheet.getRange("C2:C" + sheet.getLastRow()).getValues().flat().filter(String);
// }


// ────────────────────────────────────────────────────────────────────────
// SETUP GUIDE
// ────────────────────────────────────────────────────────────────────────
/*
╔══════════════════════════════════════════════════════════════════════╗
║   QCDS EMAIL SYSTEM v2.0 — SETUP GUIDE                             ║
╚══════════════════════════════════════════════════════════════════════╝

1. FILE: Save emailsystem.html as 'emailsystem' in your Apps Script project
   (File > New > HTML File > name it 'emailsystem')

2. CODE: Paste this file's contents into Code.gs (add below existing code)
   Make sure not to duplicate existing functions.

3. ROUTE: In your doGet(), add:
   if (page === "email") {
     return HtmlService.createHtmlOutputFromFile('emailsystem').setTitle("Email System");
   }

4. SHEETS: Ensure "Team Details" sheet exists with:
   - Column B: Member Name
   - Column C: Member Email
   The "Email History" sheet will auto-create on first send.

5. PASSWORD: Run setPassword() once from the editor to set your admin password.

6. TRIGGER: For auto daily report, create a time-based trigger:
   - Function: autoSendExactV2
   - Runs every: Hour (midnight trigger)

7. ACCESS: Visit your web app URL with ?page=email
   e.g. https://script.google.com/macros/s/YOUR_ID/exec?page=email

GOOGLE GROUP MODE:
  - Email goes to qcds-team-reports@googlegroups.com
  - Google Group distributes to all members
  - Uses ONLY 1 email quota per send
  - All members must be subscribed to the Google Group

INDIVIDUAL BCC MODE:
  - Sends in batches of 45 via BCC
  - Recipients are hidden from each other
  - Uses 1 quota per batch (multiple batches for large teams)

QUOTA LIMITS:
  - Google Workspace: 1,500 emails/day
  - Gmail: 500 emails/day
  - Group mode always uses just 1 quota regardless of team size

╔══════════════════════════════════════════════════════════════════════╗
║   NEW FUNCTIONS ADDED                                                ║
╠══════════════════════════════════════════════════════════════════════╣
║  sendEmailViaGroup(subject, htmlBody, senderMode)                    ║
║  sendEmailSystemBatch(subject, htmlBody, emails, senderMode)         ║
║  logEmailHistory(subject, status, mode, senderMode)                  ║
║  getEmailHistory()                                                    ║
║  getMembersForEmailSystem()  [unchanged]                              ║
║  checkEmailStatus()          [unchanged]                              ║
║  sendDailyReportViaGroup()   [new - replaces sendDailyReport()]       ║
║  buildDailyReportHtml(...)   [new - reusable HTML builder]            ║
║  autoSendExactV2()           [new - replaces autoSendExact()]         ║
╚══════════════════════════════════════════════════════════════════════╝
*/

// ═══════════════════════════════════════════════════════════════════════
// MEMBER PERSONAL REPORT — Backend Functions
// Add these to your Code.gs
// ═══════════════════════════════════════════════════════════════════════


// ════════════════════════════════════════════════════════════════════════
// MEMBER PERSONAL REPORT — Complete Backend (memberreport.html)
// QuantumCoders Data Solutions — v2.0
// ════════════════════════════════════════════════════════════════════════


// ────────────────────────────────────────────────────────────────────────
// WRAPPER: getFilteredReport — HTML calls this function name
// ────────────────────────────────────────────────────────────────────────
function getFilteredReport(name, fromDate, toDate) {
  return getFilteredUserData(name, fromDate, toDate);
}


// ────────────────────────────────────────────────────────────────────────
// WRAPPER: sendReportEmail — HTML calls this function name
// ────────────────────────────────────────────────────────────────────────
function sendReportEmail(recipientEmail, reportData, memberName) {
  return sendReportViaEmail(recipientEmail, reportData, memberName);
}


// ────────────────────────────────────────────────────────────────────────
// 1. getUniqueNames — Autocomplete dropdown in HTML
//    Returns sorted unique member names from "File Complited" sheet
// ────────────────────────────────────────────────────────────────────────
function getUniqueNames() {

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Column B (index 2) = Name
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

  const names = data
    .map(row => (row[0] || "").toString().trim())
    .filter(name => name !== "");

  // Deduplicate and sort alphabetically
  const unique = [...new Set(names)].sort((a, b) => a.localeCompare(b));

  return unique;
}


// ────────────────────────────────────────────────────────────────────────
// 2. getFilteredUserData — Core report fetch logic
//    Reads "File Complited", filters by name + date range
//    Returns: records, summary stats, chart data, rank
// ────────────────────────────────────────────────────────────────────────
function getFilteredUserData(name, fromDate, toDate) {

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  // Columns A–H: SlNo | Name | WorkType | Status | Date | ID | TimeSec | TimeMin
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();

  const tz   = Session.getScriptTimeZone();

  // Parse from/to from "yyyy-MM-dd" (HTML date input format)
  const fromParts = fromDate.split("-");
  const toParts   = toDate.split("-");

  const from = new Date(
    parseInt(fromParts[0]),
    parseInt(fromParts[1]) - 1,
    parseInt(fromParts[2]),
    0, 0, 0
  );
  const to = new Date(
    parseInt(toParts[0]),
    parseInt(toParts[1]) - 1,
    parseInt(toParts[2]),
    23, 59, 59
  );

  // Normalise search name
  const cleanedInput = name.trim().toLowerCase();

  const records    = [];
  const dailyMap   = {};
  let totalFiles   = 0;
  let totalSeconds = 0;
  let totalMinutes = 0;

  data.forEach(function(row) {

    const rowName = (row[1] || "").toString().trim().toLowerCase();
    const rawDate = row[4];
    const timeSec = parseFloat(row[6]) || 0;
    const timeMin = parseFloat(row[7]) || 0;

    if (!rawDate || !rowName) return;

    // ── Name match ──
    if (!rowName.includes(cleanedInput) && cleanedInput !== rowName) return;

    // ── Robust date parsing ──
    // Sheet stores dates as Date objects OR "M/d/yyyy" strings
    let rowDate;

    if (rawDate instanceof Date) {
      rowDate = new Date(rawDate.getTime());
    } else {
      const str = rawDate.toString().trim();

      // Try "M/d/yyyy"  e.g. 5/4/2026
      const slashMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (slashMatch) {
        rowDate = new Date(
          parseInt(slashMatch[3]),
          parseInt(slashMatch[1]) - 1,
          parseInt(slashMatch[2]),
          12, 0, 0
        );
      } else {
        // Fallback: native parse
        rowDate = new Date(str);
      }
    }

    if (!rowDate || isNaN(rowDate.getTime())) return;

    // ── Date range filter ──
    if (rowDate < from || rowDate > to) return;

    // ── Format display date & chart key ──
    const displayDate = Utilities.formatDate(rowDate, tz, "dd MMM yyyy");
    const chartKey    = Utilities.formatDate(rowDate, tz, "dd MMM");

    totalFiles++;
    totalSeconds += timeSec;
    totalMinutes += timeMin;

    if (!dailyMap[chartKey]) dailyMap[chartKey] = { files: 0, time: 0, sortDate: rowDate };
    dailyMap[chartKey].files++;
    dailyMap[chartKey].time += timeMin;

    records.push({
      slNo:        row[0] !== "" ? row[0] : "",
      name:        row[1] || "",
      workType:    row[2] || "",
      status:      row[3] || "",
      date:        displayDate,
      id:          row[5] !== "" ? row[5] : "",
      timeSec:     timeSec,
      timeMin:     timeMin,
      finalStatus: row[3] || ""
    });
  });

  if (records.length === 0) return { records: [] };

  // ── H:M:S ──
  const totalHMS   = convertMinutesToHMS(totalMinutes);
  const avgMinutes = totalFiles > 0 ? totalMinutes / totalFiles : 0;

  // ── Chart data — sorted by actual date, last 20 days ──
  const chartData = Object.keys(dailyMap)
    .sort(function(a, b) {
      return dailyMap[a].sortDate.getTime() - dailyMap[b].sortDate.getTime();
    })
    .slice(-20)
    .map(function(key) {
      return {
        date:  key,
        files: dailyMap[key].files,
        time:  parseFloat(dailyMap[key].time.toFixed(2))
      };
    });

  // ── Rank ──
  const rank = getRankForMember(name);

  return {
    records:      records,
    totalFiles:   totalFiles,
    totalSeconds: parseFloat(totalSeconds.toFixed(3)),
    totalMinutes: parseFloat(totalMinutes.toFixed(4)),
    totalHMS:     totalHMS,
    avgMinutes:   parseFloat(avgMinutes.toFixed(4)),
    chartData:    chartData,
    rank:         rank
  };
}


// ────────────────────────────────────────────────────────────────────────
// 3. getRankForMember — Member's all-time rank by total work minutes
// ────────────────────────────────────────────────────────────────────────
function getRankForMember(name) {

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("File Complited");

  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 2, lastRow - 1, 7).getValues();
  // Columns read: B(Name) through H(TimeMin) — index [0]=Name, [6]=TimeMin

  const totals = {};

  data.forEach(function(row) {
    const n = (row[0] || "").toString().trim();
    const t = parseFloat(row[6]) || 0;
    if (!n) return;
    if (!totals[n]) totals[n] = 0;
    totals[n] += t;
  });

  const sorted     = Object.keys(totals).sort(function(a, b) { return totals[b] - totals[a]; });
  const cleanInput = name.trim().toLowerCase();

  let position = null;

  for (let i = 0; i < sorted.length; i++) {
    const memberLower = sorted[i].toLowerCase();
    if (memberLower === cleanInput ||
        memberLower.includes(cleanInput) ||
        cleanInput.includes(memberLower)) {
      position = i + 1;
      break;
    }
  }

  return {
    position:     position,
    totalMembers: sorted.length
  };
}


// ────────────────────────────────────────────────────────────────────────
// 4. sendReportViaEmail — Sends personal report email with rate limiting
// ────────────────────────────────────────────────────────────────────────
function sendReportViaEmail(recipientEmail, reportData, memberName) {

  try {

    if (!recipientEmail || !reportData) {
      return { success: false, message: "Invalid email or report data." };
    }

    // ── Daily rate limit: 2 emails per recipient per day ──
    const props    = PropertiesService.getScriptProperties();
    const tz       = Session.getScriptTimeZone();
    const today    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const limitKey = "RPT_EMAIL_" + recipientEmail.replace(/[^a-zA-Z0-9]/g, "_") + "_" + today;
    const sentCount = parseInt(props.getProperty(limitKey) || "0");

    if (sentCount >= 2) {
      return {
        success: false,
        message: "Daily email limit (2 per day) reached for this address. Try again tomorrow."
      };
    }

    const subject  = "📊 Work Report — " + memberName + " · QuantumCoders Data Solutions";
    const htmlBody = buildMemberReportEmail(memberName, reportData);

    MailApp.sendEmail({
      to:       recipientEmail,
      subject:  subject,
      htmlBody: htmlBody,
      name:     "QuantumCoders Data Solutions"
    });

    // Increment counter only after successful send
    props.setProperty(limitKey, String(sentCount + 1));

    return {
      success: true,
      message: "Report sent successfully to " + recipientEmail +
               ". (" + (sentCount + 1) + "/2 emails used today)"
    };

  } catch (err) {
    return {
      success: false,
      message: "Send failed: " + err.toString()
    };
  }
}


// ────────────────────────────────────────────────────────────────────────
// 5. buildMemberReportEmail — Styled HTML email for personal report
// ────────────────────────────────────────────────────────────────────────
function buildMemberReportEmail(memberName, data) {

  const rankText = (data.rank && data.rank.position)
    ? "#" + data.rank.position + " of " + data.rank.totalMembers + " members"
    : "N/A";

  const tableRows = (data.records || []).slice(0, 50).map(function(r, i) {
    const bg = i % 2 === 0 ? "#ffffff" : "#f8fafc";
    return "<tr style='background:" + bg + "'>" +
      "<td style='padding:8px 10px;border:1px solid #e2e8f0;font-size:12px'>" + (r.slNo || "—") + "</td>" +
      "<td style='padding:8px 10px;border:1px solid #e2e8f0;font-size:12px'>" + (r.date || "—") + "</td>" +
      "<td style='padding:8px 10px;border:1px solid #e2e8f0;font-size:12px'>" + (r.workType || "—") + "</td>" +
      "<td style='padding:8px 10px;border:1px solid #e2e8f0;font-size:12px'>" + (r.id || "—") + "</td>" +
      "<td style='padding:8px 10px;border:1px solid #e2e8f0;font-size:12px;text-align:right'>" +
        (r.timeMin ? r.timeMin.toFixed(2) : "—") + "</td>" +
      "<td style='padding:8px 10px;border:1px solid #e2e8f0;font-size:12px'>" + (r.finalStatus || "—") + "</td>" +
    "</tr>";
  }).join("");

  const generatedOn = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy, hh:mm a");

  return "<!DOCTYPE html>" +
  "<html><head><meta charset='UTF-8'>" +
  "<meta name='viewport' content='width=device-width,initial-scale=1.0'></head>" +
  "<body style='margin:0;padding:16px;background:#F0F2F5;" +
    "font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Arial,sans-serif'>" +

  "<div style='max-width:600px;margin:0 auto;background:#fff;" +
    "border-radius:24px;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,0.09)'>" +

  // HEADER
  "<div style='background:linear-gradient(135deg,#071629 0%,#0b2040 35%,#1456A8 75%,#1a6bbf 100%);" +
    "padding:30px 28px;text-align:center'>" +
  "<img src='https://www.quantumcoderstechlab.codes/assets/Data%20Solutions/Logo/QCDL.png' " +
    "alt='QCDS' style='height:44px;width:auto;display:block;margin:0 auto 12px'>" +
  "<div style='font-size:12px;font-weight:800;color:#fff;letter-spacing:1.5px;text-transform:uppercase'>" +
    "QUANTUMCODERS DATA SOLUTIONS</div>" +
  "<h1 style='font-size:20px;font-weight:800;margin:14px 0 0;color:#fff'>PERSONAL WORK REPORT</h1>" +
  "<div style='display:inline-block;background:rgba(255,255,255,0.15);padding:6px 16px;" +
    "border-radius:40px;margin-top:10px;font-size:13px;font-weight:600;color:#fff'>👤 " +
    memberName + "</div></div>" +

  // SUMMARY
  "<div style='padding:24px 24px 8px'>" +
  "<div style='font-size:11px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;" +
    "color:#94a3b8;margin-bottom:14px'>Summary</div>" +
  "<table style='width:100%;border-collapse:collapse'><tr>" +
  "<td style='padding:10px;background:#f8fafc;border-radius:12px;text-align:center;width:25%'>" +
    "<div style='font-size:26px;font-weight:800;color:#1456A8'>" + data.totalFiles + "</div>" +
    "<div style='font-size:11px;color:#64748b;margin-top:3px'>Total Files</div></td>" +
  "<td style='width:3%'></td>" +
  "<td style='padding:10px;background:#f8fafc;border-radius:12px;text-align:center;width:25%'>" +
    "<div style='font-size:22px;font-weight:800;color:#059669'>" +
      parseFloat(data.totalMinutes).toFixed(2) + "</div>" +
    "<div style='font-size:11px;color:#64748b;margin-top:3px'>Total Min</div></td>" +
  "<td style='width:3%'></td>" +
  "<td style='padding:10px;background:#f8fafc;border-radius:12px;text-align:center;width:25%'>" +
    "<div style='font-size:16px;font-weight:800;color:#1456A8'>" + data.totalHMS + "</div>" +
    "<div style='font-size:11px;color:#64748b;margin-top:3px'>H:M:S</div></td>" +
  "<td style='width:3%'></td>" +
  "<td style='padding:10px;background:#fff9e8;border-radius:12px;text-align:center;" +
    "width:19%;border:1px solid #fde68a'>" +
    "<div style='font-size:18px;font-weight:800;color:#d97706'>" + rankText + "</div>" +
    "<div style='font-size:11px;color:#64748b;margin-top:3px'>🏆 Rank</div></td>" +
  "</tr></table></div>" +

  // TABLE
  "<div style='padding:16px 24px 24px'>" +
  "<div style='font-size:11px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;" +
    "color:#94a3b8;margin-bottom:12px'>Work Details (first 50 records)</div>" +
  "<div style='overflow-x:auto'>" +
  "<table style='width:100%;border-collapse:collapse;font-size:12px'>" +
  "<thead><tr style='background:linear-gradient(135deg,#071629,#1456A8)'>" +
  "<th style='padding:9px 10px;border:1px solid #1a6bbf;color:#fff;text-align:left'>Sl</th>" +
  "<th style='padding:9px 10px;border:1px solid #1a6bbf;color:#fff;text-align:left'>Date</th>" +
  "<th style='padding:9px 10px;border:1px solid #1a6bbf;color:#fff;text-align:left'>Work Type</th>" +
  "<th style='padding:9px 10px;border:1px solid #1a6bbf;color:#fff;text-align:left'>ID</th>" +
  "<th style='padding:9px 10px;border:1px solid #1a6bbf;color:#fff;text-align:right'>Min</th>" +
  "<th style='padding:9px 10px;border:1px solid #1a6bbf;color:#fff;text-align:left'>Status</th>" +
  "</tr></thead><tbody>" +
  (tableRows || "<tr><td colspan='6' style='padding:20px;text-align:center;color:#94a3b8'>No records</td></tr>") +
  "</tbody></table></div>" +
  (data.records.length > 50
    ? "<div style='font-size:11px;color:#94a3b8;text-align:center;margin-top:8px'>Showing first 50 of " +
      data.records.length + " records.</div>"
    : "") +
  "</div>" +

  // FOOTER
  "<div style='background:#f8fafc;border-top:1px solid #e2e8f0;padding:20px 24px;text-align:center'>" +
  "<img src='https://www.quantumcoderstechlab.codes/assets/Data%20Solutions/Logo/QCDL.png' " +
    "alt='QCDS' style='height:28px;width:auto;opacity:0.5;display:block;margin:0 auto 10px'>" +
  "<div style='font-size:12px;font-weight:700;color:#0d1b2e'>QuantumCoders Data Solutions</div>" +
  "<div style='font-size:10px;color:#64748b;margin-top:4px'>Berhampur, Odisha, India 760001</div>" +
  "<div style='margin-top:5px'><a href='https://www.quantumcoderstechlab.codes' " +
    "style='color:#1456A8;font-size:10px;text-decoration:none;font-weight:600'>" +
    "www.quantumcoderstechlab.codes</a></div>" +
  "<div style='font-size:10px;color:#94a3b8;margin-top:12px;padding-top:10px;" +
    "border-top:1px solid #e2e8f0'>Automated personal report · QCDS Member Portal<br>" +
    "Generated on " + generatedOn + "</div>" +
  "</div>" +

  "</div></body></html>";
}


// ────────────────────────────────────────────────────────────────────────
// 6. convertMinutesToHMS — Shared helper
// ────────────────────────────────────────────────────────────────────────
function convertMinutesToHMS(minutes) {
  const totalSeconds = Math.floor(minutes * 60);
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = totalSeconds % 60;
  return String(h).padStart(2, "0") + ":" +
         String(m).padStart(2, "0") + ":" +
         String(s).padStart(2, "0");
}