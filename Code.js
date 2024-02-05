// Constants for spreadsheet ID and sheet names
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const SHEET_NAMES = {
  SplunkOnCall: 'Top 5 Paging Alerts',
  EmailAlerts: 'Top root@ Mail Alerts'
};

/**
 * Initiates the analysis and reporting process for alerts from various sources.
 */
function runAlertsAnalysis() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(SHEET_NAMES).forEach(source => {
    processAlerts(spreadsheet, source);
  });
}

/**
 * General handler for processing alerts of any specified type.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The target spreadsheet.
 * @param {string} source - The source of alerts to process ('SplunkOnCall', 'EmailAlerts', 'Logstash', etc.).
 */
function processAlerts(spreadsheet, source) {
  const alertsData = fetchAlertsData(spreadsheet, source);
  outputData(spreadsheet, SHEET_NAMES[source], alertsData);
}

/**
 * Fetches alert data based on the specified source.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet to fetch data from.
 * @param {string} source - The source of alerts ('SplunkOnCall', 'EmailAlerts', 'Logstash', etc.).
 * @return {Object[]} - An array of alert objects.
 */
function fetchAlertsData(spreadsheet, source) {
  switch (source) {
    case 'SplunkOnCall':
      return fetchSplunkOnCallData(spreadsheet);
    case 'EmailAlerts':
      return fetchEmailAlerts();
    default:
      console.log(`Unknown alert source: ${source}`);
      return [];
  }
}

/**
 * Fetches alert data from the SplunkOnCall sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet The spreadsheet to fetch data from.
 * @return {Object[]} An array of alert objects.
 */
function fetchSplunkOnCallData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('SplunkOnCall');
  if (!sheet) {
    Logger.log(`Sheet not found: ${sheetName}`);
    return [];
  }

  const [columnHeaders, ...dataRows] = sheet.getDataRange().getValues();
  const serviceIndex = columnHeaders.indexOf('Service');
  const countIndex = columnHeaders.indexOf('Number Of Alerts');

  if (serviceIndex === -1 || countIndex === -1) {
    throw new Error('Required columns not found in the SplunkOnCall sheet');
  }

  return dataRows.map(row => ({
    name: row[serviceIndex],
    count: parseInt(row[countIndex], 10) || 0
  })).sort((a, b) => b.count - a.count);
}

/**
 * Outputs the provided alert data to the specified sheet.
 * This function is adapted to handle different types of alerts with generic handling for future data sources.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The target spreadsheet.
 * @param {string} sheetName - The name of the output sheet.
 * @param {Object[]} data - The alert data to output.
 */
function outputData(spreadsheet, sheetName, data) {
  const sheet = prepareSheet(spreadsheet, sheetName);
  const totalAlertsCount = data.reduce((total, {count}) => total + count, 0);

  data.forEach(({count, subject, name}, index) => {
    const percentage = ((count / totalAlertsCount) * 100).toFixed(2);
    const row = [`${count} (${percentage}%)`, subject || name];
    sheet.appendRow(row);

    if (index < 5) {
      sheet.getRange(`D${index + 2}:E${index + 2}`).setValues([row]);
    }
  });

  createFooter(sheet, totalAlertsCount);
}

function calculateTotalAlertsCount(alertsData) {
  return alertsData.reduce((acc, alert) => acc + alert.count, 0);
}

function getTop5Alerts(alertsData) {
  return alertsData.sort((a, b) => b.count - a.count).slice(0, 5);
}

/**
 * Prepares a sheet for output by clearing it or creating a new one if necessary.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet instance.
 * @param {string} sheetName - The name of the sheet to prepare.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The prepared sheet.
 */
function prepareSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  sheet.clear();
  setHeaders(sheet, sheetName);
  return sheet;
}

/**
 * Sets headers for the specified sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to set headers on.
 * @param {string} sheetName - The name of the sheet, used to determine header content.
 */
function setHeaders(sheet, sheetName) {
  // Ensures headers are only set if not already present
  if(sheet.getLastRow() === 0) {
    const titleForD = 'Top 5 Count';
    const titleForE = sheetName.includes('Mail') ? 'Subject' : 'Alert Name';
    sheet.appendRow(['Count', 'Alert Name', '', titleForD, titleForE]);
    sheet.getRange('A1:E1').setFontWeight('bold').setHorizontalAlignment('center');
  }
}

/**
 * Creates a footer on the sheet summarizing the total alerts and percentage of top 5 alerts.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to append the footer to.
 * @param {number} totalAlertsCount - The total count of alerts.
 */
function createFooter(sheet, totalAlertsCount) {
  // Footer explicitly starts on D8 regardless of the data rows
  sheet.getRange("D8").setValue('Total Alerts:');
  sheet.getRange("E8").setValue(totalAlertsCount);
  const top5Percentage = calculateTop5Percentage(sheet, totalAlertsCount);
  sheet.getRange("D9").setValue('Top 5 Alerts % of Total:');
  sheet.getRange("E9").setValue(`${top5Percentage}%`);
  sheet.getRange("D8:E9").setFontWeight('bold');
}

/**
 * Fetches email alerts based on a predefined query. Adapted for generic handling.
 * @return {Object[]} - An array of objects representing deduplicated email alerts and their counts.
 */
function fetchEmailAlerts() {
    const query = '(to:root OR from:root) AND newer_than:100d';
    const threads = GmailApp.search(query);
    let alerts = [];
    for (let thread of threads) {
        const messages = thread.getMessages();
        for (let message of messages) {
            const subject = message.getSubject();
            alerts.push(subject); // Assuming further deduplication and processing happens here
        }
    }

    return deduplicateAlerts(alerts); // Return actual data in the format [{ subject: '', count: 0 }]
}

/**
 * Deduplicates a list of email subjects and counts occurrences.
 *
 * @param {string[]} subjects - The email subjects to deduplicate.
 * @return {Object[]} An array of objects with subject and count properties.
 */
function deduplicateAlerts(subjects) {
    let counts = {};
    subjects.forEach(subject => {
        // Normalize the subject to a lower case for case-insensitive comparison
        let normalizedSubject = subject.toLowerCase();
        counts[normalizedSubject] = (counts[normalizedSubject] || 0) + 1;
    });

    // Convert the counts object into an array of {subject, count} objects
    let alerts = Object.keys(counts).map(subject => ({
        subject: subject,
        count: counts[subject]
    }));

    // Sort alerts by count in descending order for reporting the top alerts
    alerts.sort((a, b) => b.count - a.count);

    return alerts;
}

/**
 * Calculate the percentage of the top 5 alerts out of the total count
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet from which to read the top 5 alerts' counts.
 * @param {number} totalAlertsCount - The total count of all alerts, used as the denominator in the calculation.
 * @return {number} - The percentage of the total alerts that the top 5 alerts constitute, rounded to two decimal places.
 */
function calculateTop5Percentage(sheet, totalAlertsCount) {
  const top5Values = sheet.getRange('A2:A6').getValues();
  let top5Count = 0;
  top5Values.forEach((value) => {
    const count = parseInt(value[0].split(' ')[0]);
    top5Count += count;
  });
  return (top5Count / totalAlertsCount) * 100;
}
