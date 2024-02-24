// Global constants and configurations
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const DOCUMENT_ID = SCRIPT_PROPERTIES.getProperty('DOCUMENT_ID');
const VO_API_KEY = SCRIPT_PROPERTIES.getProperty('VO_API_KEY');
const VO_API_ID = SCRIPT_PROPERTIES.getProperty('VO_API_ID');
const COLOR_PALETTE = {
  White: '#FFFFFF', Black: '#000000', Black75: '#404040',
  Black50: '#7F7F7F', Black25: '#BFBFBF', BlueAAA: '#0C57A8'
};
const cache = CacheService.getScriptCache();

/**
 * Main function to run the alert analysis and reporting process.
 */
async function runAlertsAnalysis() {
  try {
    const doc = DocumentApp.openById(DOCUMENT_ID);
    const body = doc.getBody();
    prepareDocument(body);

    await Promise.all([
      writeAlertsSection(body, 'Top Root@ mail alerts', fetchEmailAlerts),
      writeAlertsSection(body, 'Top Paging alerts', fetchVictorOpsIncidents)
    ]);
  } catch (error) {
    Logger.log(`Error running alerts analysis: ${error}`);
  }
}

/**
 * Prepares the document by setting up a title, subtitle, and a horizontal rule.
 * @param {GoogleAppsScript.Document.Body} body - The body of the document to prepare.
 */
function prepareDocument(body) {
  body.clear();
  appendStyledText(body, 'ALERT REVIEW', 'title');
  appendStyledText(body, 'Q3', 'subtitle');
  body.appendHorizontalRule();
}

/**
 * Writes a specific section of alerts into the document.
 * @param {GoogleAppsScript.Document.Body} body - The body of the Google Doc.
 * @param {string} sectionTitle - The section's title.
 * @param {Array} data - The alerts data to be written.
 */
async function writeAlertsSection(body, sectionTitle, fetchDataFunction) {
  const data = await fetchDataFunction(); // Make sure this is awaited if fetchDataFunction is async
  const totalAlertsCount = calculateTotalAlertsCount(data);

  appendStyledText(body, sectionTitle, 'heading');
  appendStyledText(body, `Last 100d days`, 'subheading');
  appendStyledTable(body, ['Top 5 count', 'Subject'], data.slice(0, 5), totalAlertsCount);
  appendStyledText(body, `The top 5 alerts represent ${calculateTopAlertsPercent(data.slice(0, 5), totalAlertsCount)}% of the total number of alerts (${totalAlertsCount}).`, 'text');

  body.appendHorizontalRule();
}

/**
 * Appends styled text to a document.
 * @param {GoogleAppsScript.Document.Body} body - The body of the Google Doc.
 * @param {string} text - The text to be added.
 * @param {string} styleType - The style type (e.g., title, subtitle).
 */
function appendStyledText(body, text, styleType) {
  // Define styles based on the type
  const styles = {
    title: {FONT_FAMILY: 'Montserrat', FONT_SIZE: 36, BOLD: true, FOREGROUND_COLOR: COLOR_PALETTE.Black, HORIZONTAL_ALIGNMENT: DocumentApp.HorizontalAlignment.CENTER},
    subtitle: {FONT_SIZE: 18, FOREGROUND_COLOR: COLOR_PALETTE.Black75, HORIZONTAL_ALIGNMENT: DocumentApp.HorizontalAlignment.CENTER},
    heading: {FONT_FAMILY: 'Montserrat', FONT_SIZE: 28, BOLD: true, FOREGROUND_COLOR: COLOR_PALETTE.Black},
    subheading: {FONT_FAMILY: 'Montserrat', FONT_SIZE: 18, BOLD: false},
    text: {FONT_FAMILY: 'Montserrat', FONT_SIZE: 14, BOLD: false, FOREGROUND_COLOR: COLOR_PALETTE.Black, HORIZONTAL_ALIGNMENT: DocumentApp.HorizontalAlignment.LEFT}
  };

  const paragraph = body.appendParagraph(text);
  paragraph.setAttributes(styles[styleType]);
}

/**
 * Appends a styled table to the document body.
 * @param {GoogleAppsScript.Document.Body} body - The body of the document.
 * @param {Array<string>} titles - Titles for the table headers.
 * @param {Array<Object>} data - Data to populate the table.
 * @param {number} totalAlertsCount - Total count of alerts for percentage calculation.
 */
function appendStyledTable(body, titles, data, totalAlertsCount) {
  const styles = {
    table: {BACKGROUND_COLOR: null, BOLD: false, BORDER_COLOR: COLOR_PALETTE.BlueAAA, FONT_FAMILY: 'Montserrat', FONT_SIZE: 10, FOREGROUND_COLOR: COLOR_PALETTE.Black},
    title: {BACKGROUND_COLOR: COLOR_PALETTE.BlueAAA, BOLD: true, FONT_SIZE: 12, FOREGROUND_COLOR: COLOR_PALETTE.White},
    content: {BACKGROUND_COLOR: null, BOLD: false, FONT_SIZE: 12, FOREGROUND_COLOR: COLOR_PALETTE.Black}
  };

  const table = body.appendTable();
  table.setAttributes(styles.table);

  const headerRow = table.appendTableRow();
  titles.forEach(title => {
    const cell = headerRow.appendTableCell(title).setBackgroundColor(styles.title.BACKGROUND_COLOR).setBold(styles.title.BOLD).setFontSize(styles.title.FONT_SIZE).setForegroundColor(styles.title.FOREGROUND_COLOR);
  });

  data.forEach(({ count, subject }) => {
    const row = table.appendTableRow();
    const percentage = ((count / totalAlertsCount) * 100).toFixed(2);
    row.appendTableCell(`${count} (${percentage}%)`).setAttributes(styles.content);
    row.appendTableCell(subject).setAttributes(styles.content);
  });
}

/**
 * Calculates the total count of alerts.
 * @param {Array<Object>} data - Data array containing alerts.
 * @returns {number} Total count of alerts.
 */
function calculateTotalAlertsCount(data) {
  return data.reduce((total, {count}) => total + count, 0);
}

/**
 * Calculates the count of top alerts.
 * @param {Array} data - The data array containing the alerts.
 * @param {number} topCount - The number of top items to consider.
 * @return {number} The total count of alerts.
 */
function calculateTopAlertsCount(data, topCount) {
  const topAlerts = [...data].sort((a, b) => b.count - a.count).slice(0, topCount);
  return topAlerts.reduce((total, {count}) => total + count, 0);
}

/**
 * Calculates the percentage represented by the top alerts out of the total number of alerts.
 * @param {Array<Object>} data - The data array containing the alerts.
 * @param {number} totalAlertsCount - The total count of all alerts.
 * @returns {string} The percentage of total alerts represented by the top alerts, formatted as a string.
 */
function calculateTopAlertsPercent(data, totalAlertsCount) {
  const topAlertsSum = data.slice(0, 5).reduce((sum, item) => sum + item.count, 0);
  return ((topAlertsSum / totalAlertsCount) * 100).toFixed(2);
}

/**
 * Fetches email alerts based on a given query. Handles Gmail search internally to abstract away details.
 * @return {Object[]} An array of email alert objects with subject and count.
 */
async function fetchEmailAlerts() {
  const cacheKey = 'emailAlerts';
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    return JSON.parse(cachedData);
  }

  const query = `(to:root OR from:root) AND newer_than:100d`;
  try {
    const threads = await GmailApp.search(query);
    let subjects = threads.flatMap(thread => thread.getMessages().map(message => message.getSubject()));
    const uniqueSubjects = deduplicateAlerts(subjects);
    cache.put(cacheKey, JSON.stringify(uniqueSubjects), 21600); // Cache for 6 hours
    return uniqueSubjects;
  } catch (error) {
    Logger.log('Error fetching email alerts: ' + error.toString());
    return []; // Return empty array on error to maintain function's contract
  }
}

/**
 * Deduplicates a list of email subjects or incident alerts and counts their occurrences.
 * @param {string[]} subjects - The email subjects or incident names to deduplicate.
 * @return {Object[]} An array of objects with subject and count properties.
 */
function deduplicateAlerts(subjects) {
  const counts = subjects.reduce((acc, subject) => {
    let normalizedSubject = subject.toLowerCase().trim();

    // Extract the multiplier if present, e.g., "[3x]" or "[FIRING:1]"
    const multiplierMatch = normalizedSubject.match(/\[(\d+)x\]|\[FIRING:(\d+)\]/);
    let multiplier = 1;
    if (multiplierMatch) {
      multiplier = parseInt(multiplierMatch[1] || multiplierMatch[2], 10);
      normalizedSubject = normalizedSubject.replace(/\[\d+x\]|\[FIRING:\d+\]/, '').trim();
    }

    acc[normalizedSubject] = (acc[normalizedSubject] || 0) + multiplier;
    return acc;
  }, {});

  return Object.entries(counts).map(([subject, count]) => ({
    subject: subject,
    count: count
  })).sort((a, b) => b.count - a.count);
}

/**
 * Fetches VictorOps incidents and returns deduplicated incidents data.
 * @return {Object[]} Array of incident objects with count and subject.
 */
async function fetchVictorOpsIncidents() {
  const cacheKey = 'pagingIncidents';
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    return JSON.parse(cachedData);
  }

  const daysBefore = 60;
  const dateISO8601 = getDateDaysBefore(daysBefore);
  const apiUrl = 'https://api.victorops.com/api-reporting/v2/incidents';
  const queryParams = `?limit=100&startedAfter=${dateISO8601}`;
  const fullUrl = apiUrl + queryParams;

  const headers = {
    'Accept': 'application/json',
    'X-VO-Api-Id': VO_API_ID,
    'X-VO-Api-Key': VO_API_KEY
  };

  const options = {
    'method': 'get',
    'headers': headers,
  };

  try {
    const response = await UrlFetchApp.fetch(fullUrl, options);
    const responseData = JSON.parse(response.getContentText());

    if (responseData && responseData.incidents) {
      let incidents = responseData.incidents.map(incident => incident.service);
      const uniqueIncidents = deduplicateAlerts(incidents);
      cache.put(cacheKey, JSON.stringify(uniqueIncidents), 21600); // Cache for 6 hours
      return deduplicateAlerts(responseData.incidents.map(incident => incident.service));
    } else {
      console.error("No incidents data found in response");
      return [];
    }
  } catch (error) {
    console.error("Failed to fetch VictorOps incidents:", error);
    return []; // Return an empty array in case of an error
  }
}

/**
 * Gets a date that is a certain number of days before the current date.
 * @param {number} daysBefore - The number of days before today.
 * @return {string} The date in ISO 8601 format.
 */
function getDateDaysBefore(daysBefore) {
  const currentDate = new Date();
  currentDate.setDate(currentDate.getDate() - daysBefore);
  return currentDate.toISOString();
}
