// Constants for document ID and API keys
const DOCUMENT_ID = PropertiesService.getScriptProperties().getProperty('DOCUMENT_ID');

// Wikimedia's brand colors https://meta.wikimedia.org/wiki/Brand/colours
const COLOR_PALETTE = {
  White: '#FFFFFF',
  Black: '#000000',
  Black75: '#404040',
  Black50: '#7F7F7F',
  Black25: '#BFBFBF',
  BlueAAA: '#0C57A8'
};

/**
 * Initializes the alert analysis and reporting process.
 */
function runAlertsAnalysis() {
  const document = DocumentApp.openById(DOCUMENT_ID);
  const body = document.getBody();

  prepareDocument(body);
  writeRootEmailAlerts(body);
}

/**
 * Sets up the document with a title, subtitle, and a horizontal rule.
 *
 * @param {GoogleAppsScript.Document.Body} body - The body of the document.
 */
function prepareDocument(body) {
  body.clear();
  appendFormattedTitle(body, 'ALERT REVIEW');
  appendFormattedSubtitle(body, 'Q3');
  body.appendHorizontalRule();
}

/**
 * Writes the "Top Root@ mail alerts" section to the document.
 *
 * @param {GoogleAppsScript.Document.Body} body - The body of the document.
 */
function writeRootEmailAlerts(body) {
  const timePeriod = '100d';
  const topCount = 5;
  const query = `(to:root OR from:root) AND newer_than:${timePeriod}`;
  const data = fetchEmailAlerts(query);

  appendFormattedHeading(body, 'Top Root@ mail alerts');
  appendFormattedSubheading(body, `Last ${timePeriod} days`);
  appendFormattedTable(body, [`Top ${topCount} count`, 'Subject'], data, topCount);

  const totalAlertsCount = calculateTotalAlertsCount(data);
  const topAlertsPercent = calculateTopAlertsPercent(data, topCount);
  body.appendParagraph(`The top ${topCount} alerts represent ${topAlertsPercent}% of the total number of alerts (${totalAlertsCount}).`);
}

/**
 * Appends a formatted title to the document body.
 *
 * @param {GoogleAppsScript.Document.Body} body - The body of the document.
 * @param {string} title - The title text.
 */
function appendFormattedTitle(body, title) {
  const titleStyle = {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Montserrat',
    [DocumentApp.Attribute.FONT_SIZE]: 36,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: COLOR_PALETTE.Black,
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
  };
  body.appendParagraph(title).setAttributes(titleStyle);
}

/**
 * Appends a formatted subtitle to the document body.
 *
 * @param {GoogleAppsScript.Document.Body} body - The body of the document.
 * @param {string} subtitle - The subtitle text.
 */
function appendFormattedSubtitle(body, subtitle) {
  const subtitleStyle = {
    [DocumentApp.Attribute.FONT_SIZE]: 18,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: COLOR_PALETTE.Black75,
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
  };
  body.appendParagraph(subtitle).setAttributes(subtitleStyle);
}

/**
 * Appends a formatted heading to the document body.
 *
 * @param {GoogleAppsScript.Document.Body} body - The body of the document.
 * @param {string} heading - The heading text.
 */
function appendFormattedHeading(body, heading) {
  const headingStyle = {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Montserrat',
    [DocumentApp.Attribute.FONT_SIZE]: 28,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: COLOR_PALETTE.Black,
  };
  body.appendParagraph(heading).setAttributes(headingStyle);
}

/**
 * Appends a formatted subheading to the document body.
 *
 * @param {GoogleAppsScript.Document.Body} body - The body of the document.
 * @param {string} subheading - The subheading text.
 */
function appendFormattedSubheading(body, subheading) {
  const subheadingStyle = {
    [DocumentApp.Attribute.FONT_FAMILY]: 'Montserrat',
    [DocumentApp.Attribute.FONT_SIZE]: 18,
    [DocumentApp.Attribute.BOLD]: false,
  };
  body.appendParagraph(subheading).setAttributes(subheadingStyle);
}

/**
 * Appends a formatted table of top email alerts to the document.
 *
 * @param {GoogleAppsScript.Document.Body} body - The body of the document.
 * @param {Array} titles - Table header titles.
 * @param {Array} data - The data to fill the table with.
 * @param {number} topCount - The number of top items to include.
 */
function appendFormattedTable(body, titles, data, topCount) {
  const tableTitleStyle = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: COLOR_PALETTE.BlueAAA,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Montserrat',
    [DocumentApp.Attribute.FONT_SIZE]: 10,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: COLOR_PALETTE.White,
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.CENTER
  }

  const tableContentStyle = {
    [DocumentApp.Attribute.BACKGROUND_COLOR]: COLOR_PALETTE.White,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Montserrat',
    [DocumentApp.Attribute.FONT_SIZE]: 10,
    [DocumentApp.Attribute.BOLD]: false,
    [DocumentApp.Attribute.FOREGROUND_COLOR]: COLOR_PALETTE.Black,
    [DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]: DocumentApp.HorizontalAlignment.LEFT
  };

  const table = body.appendTable();
  const headerRow = table.appendTableRow();
  titles.forEach(title => headerRow.appendTableCell(title).setAttributes(tableTitleStyle));

  const totalAlertsCount = calculateTotalAlertsCount(data, topCount);

  data.slice(0, topCount).forEach(({subject, count}) => {
    const row = table.appendTableRow();
    const percentage = ((count / totalAlertsCount) * 100).toFixed(2);
    const countCell = row.appendTableCell(`${count} (${percentage}%)`);

    countCell.setAttributes(tableContentStyle);
    const subjectCell = row.appendTableCell(subject);
    subjectCell.setAttributes(tableContentStyle);
  });
}

/**
 * Calculates and returns the total count of alerts.
 *
 * @param {Array} data - The data array containing the alerts.
 * @return {number} The total count of alerts.
 */
function calculateTotalAlertsCount(data) {
  return data.reduce((total, {count}) => total + count, 0);
}

/**
 * Calculates and returns the count of top alerts.
 *
 * @param {Array} data - The data array containing the alerts.
 * @param {number} topCount - The number of top items to consider.
 * @return {number} The total count of alerts.
 */
function calculateTopAlertsCount(data, topCount) {
  const topAlerts = [...data].sort((a, b) => b.count - a.count).slice(0, topCount);
  return topAlerts.reduce((total, {count}) => total + count, 0);
}

/**
 * Calculates and returns the total count of alerts.
 *
 * @param {Array} data - The data array containing the alerts.
 * @param {number} topCount - The number of top items to consider.
 * @return {number} The total count of alerts.
 */
function calculateTopAlertsPercent(data, topCount) {
  const totalAlertsCount = calculateTotalAlertsCount(data);
  const topAlertsCount = calculateTopAlertsCount(data, topCount);
  return ((topAlertsCount / totalAlertsCount) * 100).toFixed(2);
}

/**
 * Fetches and deduplicates email alerts based on a given query.
 * @param {string} query - The search query for Gmail.
 * @return {Object[]} An array of email alert objects with subject and count.
 */
function fetchEmailAlerts(query) {
  const threads = GmailApp.search(query);
  let subjects = [];

  // Extract subjects from each message
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      subjects.push(message.getSubject()); // Collect all subjects
    });
  });

  return deduplicateAlerts(subjects); // Deduplicate and return
}

/**
 * Deduplicates a list of email subjects and counts their occurrences.
 * @param {string[]} subjects - The email subjects to deduplicate.
 * @return {Object[]} An array of objects with subject and count properties.
 */
function deduplicateAlerts(subjects) {
  const counts = subjects.reduce((acc, subject) => {
    // Normalize subject for case-insensitive comparison
    const normalizedSubject = subject.toLowerCase();
    acc[normalizedSubject] = (acc[normalizedSubject] || 0) + 1;
    return acc;
  }, {});

  // Convert counts to an array of { subject, count } objects and sort by count
  return Object.entries(counts).map(([subject, count]) => ({
    subject,
    count
  })).sort((a, b) => b.count - a.count);
}
