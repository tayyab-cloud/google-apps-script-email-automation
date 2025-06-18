function previewAndSendEmails() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  let rowsToSend = [];

  for (let i = 1; i < data.length; i++) {
    const toggle = data[i][8]?.toString().trim().toUpperCase(); // Column I
    const status = data[i][4];
    const subject = data[i][2];
    const message = data[i][3];
    const lastSubject = data[i][6];
    const lastMessage = data[i][7];

    const subjectChanged = subject !== lastSubject;
    const messageChanged = message !== lastMessage;

    if (toggle === "ON" && (status !== "Sent" || subjectChanged || messageChanged)) {
      rowsToSend.push(i + 1); // Store row number
    }
  }

  if (rowsToSend.length === 0) {
    ui.alert("✅ No new emails to send.");
    return;
  }

  const confirmation = ui.alert(
    "Send Emails Confirmation",
    `You are about to send ${rowsToSend.length} email(s).\n\nDo you want to continue?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmation === ui.Button.YES) {
    sendEmailsFromSheet(); // Call the function you already have
  } else {
    ui.alert("❌ Sending cancelled.");
  }
}


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📧 Email Actions")
    .addItem("Preview & Send Emails", "previewAndSendEmails")
    .addToUi();
}


function createEmailChartInDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Email Log");
  const dashboardSheet = ss.getSheetByName("Dashboard");

  // Clear old charts
  dashboardSheet.getCharts().forEach(chart => dashboardSheet.removeChart(chart));

  const dataRange = logSheet.getRange("A1:F" + logSheet.getLastRow());
  const logData = dataRange.getValues();

  // Prepare or create temp sheet
  const tempSheet = ss.getSheetByName("TempChartData") || ss.insertSheet("TempChartData");
  tempSheet.clear();
  tempSheet.getRange("A1").setValue("Date");
  tempSheet.getRange("B1").setValue("Count");

  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const countsByDate = {};
  const statusCounts = {};

  for (let i = 1; i < logData.length; i++) {
    const rawDate = logData[i][0];
    const status = logData[i][5];

    if (rawDate instanceof Date && status) {
      const plainDate = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
      const formattedDate = Utilities.formatDate(plainDate, tz, "yyyy-MM-dd");

      // Skip future dates
      if (formattedDate > today) continue;

      // Column/Line Chart Data
      if (status === "Sent") {
        countsByDate[formattedDate] = (countsByDate[formattedDate] || 0) + 1;
      }

      // Pie Chart Data
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    }
  }

  // Populate date-wise data (for column and line chart)
  let row = 2;
  Object.keys(countsByDate).sort().forEach(date => {
    tempSheet.getRange(row, 1).setValue(date);
    tempSheet.getRange(row, 2).setValue(countsByDate[date]);
    row++;
  });

  // Add status-wise data (for pie chart)
  tempSheet.getRange("D1").setValue("Status");
  tempSheet.getRange("E1").setValue("Count");

  let pieRow = 2;
  for (let status in statusCounts) {
    tempSheet.getRange(pieRow, 4).setValue(status);
    tempSheet.getRange(pieRow, 5).setValue(statusCounts[status]);
    pieRow++;
  }

  // --- 📊 Column Chart (Email per day)
  const chart1 = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(tempSheet.getRange("A1:B" + (row - 1)))
    .setPosition(2, 2, 0, 0)
    .setOption("title", "📊 Emails Sent Per Day")
    .build();

  dashboardSheet.insertChart(chart1);

  // --- 📈 Line Chart (Same data)
  const chart2 = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(tempSheet.getRange("A1:B" + (row - 1)))
    .setPosition(20, 2, 0, 0)
    .setOption("title", "📈 Email Trend Over Time")
    .build();

  dashboardSheet.insertChart(chart2);

  // --- 🥧 Pie Chart (Status distribution)
  const chart3 = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(tempSheet.getRange("D1:E" + (pieRow - 1)))
    .setPosition(38, 2, 0, 0)
    .setOption("title", "🥧 Email Status Breakdown")
    .build();

  dashboardSheet.insertChart(chart3);
}














function sendEmailsFromSheet() {
  const TEST_MODE = true; // ✅ Test Mode ON — set false to actually send emails

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const logSheet = ss.getSheetByName("Email Log");

  if (!logSheet) {
    Logger.log("❌ 'Email Log' sheet not found.");
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const email = data[i][1];
    const subject = data[i][2];
    const message = data[i][3];
    const status = data[i][4];
    const fileId = data[i][5];
    const lastSubject = data[i][6];
    const lastMessage = data[i][7];
    const toggle = data[i][8]?.toString().trim().toUpperCase(); // Column I
    const lastSentDateCell = sheet.getRange(i + 1, 10); // Column J
    const templateName = data[i][10]?.toString().trim(); // Column K

    if (toggle !== "ON") continue;

    const subjectChanged = subject !== lastSubject;
    const messageChanged = message !== lastMessage;

    if (status !== "Sent" || subjectChanged || messageChanged) {
      let htmlBody = "";
      let hasAttachment = "No";

      // 🧾 Load template
      try {
        const htmlTemplate = HtmlService.createTemplateFromFile(templateName);
        htmlTemplate.name = name;
        htmlTemplate.message = message;
        htmlBody = htmlTemplate.evaluate().getContent();
      } catch (e) {
        Logger.log(`❌ Template "${templateName}" not found for ${name}`);
        sheet.getRange(i + 1, 5).setValue("Failed");
        continue;
      }

      // 📎 Attachment (optional)
      let options = { htmlBody };
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          options.attachments = [file.getBlob()];
          hasAttachment = "Yes";
        } catch (err) {
          Logger.log("📎 Attachment not found for: " + name);
        }
      }

      // 🚦 TEST MODE or REAL EMAIL
      if (TEST_MODE) {
        Logger.log("🧪 TEST MODE — Email Preview:");
        Logger.log("To: " + email);
        Logger.log("Subject: " + subject);
        Logger.log("Body Preview:\n" + htmlBody);
        Logger.log("Attachment: " + hasAttachment);

        sheet.getRange(i + 1, 5).setValue("Tested");         // Status
        sheet.getRange(i + 1, 9).setValue("OFF");            // Toggle OFF
        lastSentDateCell.setValue(new Date());               // Optional: log preview date
      } else {
        try {
          GmailApp.sendEmail(email, subject, "", options);

          sheet.getRange(i + 1, 5).setValue("Sent");         // Status
          sheet.getRange(i + 1, 7).setValue(subject);        // Last Subject
          sheet.getRange(i + 1, 8).setValue(message);        // Last Message
          lastSentDateCell.setValue(new Date());            // Date Sent
          sheet.getRange(i + 1, 9).setValue("OFF");          // Toggle OFF

          logSheet.appendRow([
            new Date(), name, email, subject, message, "Sent", hasAttachment
          ]);

          Logger.log(`✅ Email sent to ${email} using template "${templateName}"`);
        } catch (error) {
          sheet.getRange(i + 1, 5).setValue("Failed");

          logSheet.appendRow([
            new Date(), name, email, subject, message, "Failed", hasAttachment
          ]);

          Logger.log("❌ Error sending to " + email + ": " + error.toString());
        }
      }
    }
  }
}
