// Constants for spreadsheet and sheet names
// Retrieve the spreadsheet ID from script properties
const spreadsheetID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const inputSheetName = 'SplunkOnCall';
const outputSheetName = 'Top 5 Paging Alerts';

function topPagingAlerts() {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  const inputSheet = spreadsheet.getSheetByName(inputSheetName);
  const alertsData = extractAlertsData(inputSheet);
  const totalAlertsCount = calculateTotalAlertsCount(alertsData);
  const top5Alerts = getTop5Alerts(alertsData);

  let outputSheet = prepareOutputSheet(spreadsheet, outputSheetName);
  setHeaders(outputSheet);
  appendTop5Alerts(outputSheet, top5Alerts, totalAlertsCount);
  createFooter(outputSheet, top5Alerts, totalAlertsCount);
}

function extractAlertsData(sheet) {
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  const serviceIndex = headers.indexOf('Service');
  const countIndex = headers.indexOf('Number Of Alerts');

  if (serviceIndex === -1 || countIndex === -1) {
    throw new Error('One or both headers not found');
  }

  return values.slice(1).map(row => ({
    name: row[serviceIndex],
    count: parseInt(row[countIndex], 10) // Ensure count is a number
  }));
}

function calculateTotalAlertsCount(alertsData) {
  return alertsData.reduce((acc, alert) => acc + alert.count, 0);
}

function getTop5Alerts(alertsData) {
  return alertsData.sort((a, b) => b.count - a.count).slice(0, 5);
}

function prepareOutputSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

function setHeaders(sheet) {
  sheet.appendRow(['Count', 'Alert Name']);
  const headerRange = sheet.getRange(1, 1, 1, 2);
  headerRange.setFontWeight('bold').setHorizontalAlignment('center');
}

function appendTop5Alerts(sheet, top5Alerts, totalAlertsCount) {
  top5Alerts.forEach(alert => {
    const percentage = ((alert.count / totalAlertsCount) * 100).toFixed(2);
    sheet.appendRow([`${alert.count} (${percentage}%)`, alert.name]);
  });
}

function createFooter(sheet, top5Alerts, totalAlertsCount) {
  const top5AlertsSum = top5Alerts.reduce((acc, alert) => acc + alert.count, 0);
  const top5Percentage = ((top5AlertsSum / totalAlertsCount) * 100).toFixed(2);

    // Current last row before appending new rows
  let currentLastRow = sheet.getLastRow();

  // Append rows
  sheet.appendRow([' ']); // Blank row as a separator
  sheet.appendRow(['', `Total Alerts: ${totalAlertsCount}`]);
  sheet.appendRow(['', `Top 5 Alerts account for ${top5Percentage}% of Total.`]);

  // Calculate the range that needs formatting
  // +1 for the separator row, and +2 for the two rows of content
  let startRowForFormatting = currentLastRow + 1 + 1; // Skip the separator row
  let numberOfRowsForFormatting = 2; // Number of rows with content to format

  // Get the range for the newly added rows (excluding the blank separator row)
  let rangeForFormatting = sheet.getRange(startRowForFormatting, 2, numberOfRowsForFormatting, 1);

  // Apply bold formatting to the specified range
  rangeForFormatting.setFontWeight('bold');
}
