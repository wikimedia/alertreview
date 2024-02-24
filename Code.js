// Global constants and configuration
const scriptProperties = PropertiesService.getScriptProperties()
const documentId = scriptProperties.getProperty('DOCUMENT_ID')
const voApiKey = scriptProperties.getProperty('VO_API_KEY')
const voApiId = scriptProperties.getProperty('VO_API_ID')
const colorPalette = {
  white: '#FFFFFF',
  black: '#000000',
  grey75: '#404040',
  grey50: '#7F7F7F',
  grey25: '#BFBFBF',
  blueaaa: '#0C57A8'
}

// --- Caching ---
const cacheService = CacheService.getScriptCache()
const cacheIsEnabled = true

/**
 * Gets a value from the cache only if the cache is enabled.
 * @param {string} key - The key to use for the cache lookup.
 * @returns The cached value, or null if the cache is disabled or the key is not found.
 */
function getFromCache (key) {
  return cacheIsEnabled ? cacheService.get(key) : null
}

/**
 * Puts a value into the cache only if the cache is enabled.
 * @param {string} key - The key to use for the cache store.
 * @param {string} value - The data to store in the cache.
 * @param {number} cacheExpirationSeconds - The time to live (TTL) for the cache entry.
 */
function putInCache (key, data, cacheExpirationSeconds) {
  if (cacheIsEnabled) {
    cacheService.put(key, data, cacheExpirationSeconds)
  }
}

// --- Helper Functions ---

/**
 * Calculates the total count of alerts.
 * @param {Object[]} alertData - Array of alert objects with 'count' properties.
 * @returns {number} Total count of alerts.
 */
function calculateTotalAlertCount (alertData) {
  return alertData.reduce((total, alert) => total + alert.count, 0)
}

/**
 * Calculates the percentage of top alerts out of the total number of alerts.
 * @param {Object[]} topAlerts - The top alerts data.
 * @param {number} totalAlertCount - The total count of all alerts.
 * @returns {string} The percentage of total alerts represented by the top alerts, formatted as a string.
 */
function calculateTopAlertsPercentage (topAlerts, totalAlertCount) {
  const topAlertsSum = topAlerts.reduce((sum, alert) => sum + alert.count, 0)
  return totalAlertCount === 0
    ? '0.00'
    : ((topAlertsSum / totalAlertCount) * 100).toFixed(2)
}

/**
 * Gets a date string that is a certain number of days before the current date.
 * @param {number} daysBefore - The number of days before today.
 * @return {string} The date in ISO 8601 format.
 */
function formatDateISO (daysBefore) {
  const date = new Date()
  date.setDate(date.getDate() - daysBefore)
  return date.toISOString()
}

// Alert Processing Functions

/**
 * Normalizes subject strings and extracts any multiplier.
 * @param {string} subject - The subject string to normalize.
 * @returns {Object} An object containing the normalized subject and multiplier.
 */
function extractNormalizedSubject (subject) {
  let normalized = subject.toLowerCase().trim()
  const multiplierMatch = normalized.match(/\[(\d+)x\]|\[FIRING:(\d+)\]/)
  let multiplier = 1
  if (multiplierMatch) {
    multiplier = parseInt(multiplierMatch[1] || multiplierMatch[2], 10)
    normalized = normalized.replace(/\[\d+x\]|\[FIRING:\d+\]/, '').trim()
  }
  return { normalizedSubject: normalized, multiplier }
}

/**
 * Aggregates and normalizes alerts from a list of subjects.
 * @param {string[]} subjects - The subjects to process.
 * @returns {Object[]} An array of objects each containing a subject and its count.
 */
function aggregateAlerts (subjects) {
  const counts = subjects.reduce((acc, subject) => {
    const { normalizedSubject, multiplier } = extractNormalizedSubject(subject)
    acc[normalizedSubject] = (acc[normalizedSubject] || 0) + multiplier
    return acc
  }, {})

  return Object.entries(counts).map(([subject, count]) => ({
    subject, count
  })).sort((a, b) => b.count - a.count)
}

/**
 * Appends styled text to a Google Document.
 * @param {GoogleAppsScript.Document.Body} body - The document body.
 * @param {string} text - Text to append.
 * @param {string} styleType - The style type (title, subtitle, etc.).
 */
function appendStyledText (body, text, styleType) {
  const styleAttributes = {
    title: { FONT_FAMILY: 'Montserrat', FONT_SIZE: 36, BOLD: true, FOREGROUND_COLOR: colorPalette.black, HORIZONTAL_ALIGNMENT: DocumentApp.HorizontalAlignment.CENTER },
    subtitle: { FONT_SIZE: 18, FOREGROUND_COLOR: colorPalette.black75, HORIZONTAL_ALIGNMENT: DocumentApp.HorizontalAlignment.CENTER },
    heading: { FONT_FAMILY: 'Montserrat', FONT_SIZE: 28, BOLD: true, FOREGROUND_COLOR: colorPalette.black },
    subheading: { FONT_FAMILY: 'Montserrat', FONT_SIZE: 18, BOLD: false },
    text: { FONT_FAMILY: 'Montserrat', FONT_SIZE: 14, BOLD: false, FOREGROUND_COLOR: colorPalette.black, HORIZONTAL_ALIGNMENT: DocumentApp.HorizontalAlignment.LEFT }
  }

  const paragraph = body.appendParagraph(text)
  paragraph.setAttributes(styleAttributes[styleType])
}

/**
 * Appends a styled table to a Google Document.
 * @param {GoogleAppsScript.Document.Body} body - The document body.
 * @param {Array<string>} headers - Table headers.
 * @param {Array<Object>} data - Table data.
 * @param {number} totalAlertCount - Total alert count for percentage calculations.
 */
function appendStyledTable (body, headers, data, totalAlertCount) {
  const styleAttributes = {
    table: { BACKGROUND_COLOR: null, BOLD: false, BORDER_COLOR: colorPalette.blueaaa, FONT_FAMILY: 'Montserrat', FONT_SIZE: 10, FOREGROUND_COLOR: colorPalette.black },
    title: { BACKGROUND_COLOR: colorPalette.blueaaa, BOLD: true, FONT_SIZE: 12, FOREGROUND_COLOR: colorPalette.white },
    content: { BACKGROUND_COLOR: null, BOLD: false, FONT_SIZE: 12, FOREGROUND_COLOR: colorPalette.black }
  }

  const table = body.appendTable()
  table.setAttributes(styleAttributes.table)

  const headerRow = table.appendTableRow()
  headers.forEach(header => {
    const cell = headerRow.appendTableCell(header)
    cell.setBackgroundColor(styleAttributes.title.BACKGROUND_COLOR).setBold(styleAttributes.title.BOLD).setFontSize(styleAttributes.title.FONT_SIZE).setForegroundColor(styleAttributes.title.FOREGROUND_COLOR)
  })

  data.forEach(({ count, subject }) => {
    const row = table.appendTableRow()
    const percentage = ((count / totalAlertCount) * 100).toFixed(2)
    row.appendTableCell(`${count} (${percentage}%)`).setAttributes(styleAttributes.content)
    row.appendTableCell(subject).setAttributes(styleAttributes.content)
  })
}

// --- Fetching Functions ---

/**
 * Fetches email alerts and processes them for reporting.
 * @returns {Promise<Object[]>} Processed email alerts.
 */
async function fetchAndProcessEmailAlerts () {
  const cacheKey = 'emailAlerts'
  const cachedData = getFromCache(cacheKey)

  if (cachedData) {
    return JSON.parse(cachedData)
  }

  const query = '(to:root OR from:root) AND newer_than:100d'

  try {
    const threads = await GmailApp.search(query)
    const subjects = threads.flatMap(thread => thread.getMessages().map(message => message.getSubject()))
    const uniqueSubjects = aggregateAlerts(subjects)
    putInCache(cacheKey, JSON.stringify(uniqueSubjects), 21600) // Cache for 6 hours
    return uniqueSubjects
  } catch (error) {
    const errorMessage = `Error fetching email alerts: ${error.toString()}`
    Logger.log(errorMessage)
    throw new Error(errorMessage)
  }
}

/**
 * Fetches VictorOps incidents and processes them for reporting.
 * @returns {Promise<Object[]>} Processed incidents.
 */
async function fetchAndProcessVictorOpsIncidents () {
  const cacheKey = 'pagingIncidents'
  const cachedData = getFromCache(cacheKey)

  if (cachedData) {
    return JSON.parse(cachedData)
  }

  const daysBefore = 60
  const dateISO8601 = formatDateISO(daysBefore)
  const apiUrl = 'https://api.victorops.com/api-reporting/v2/incidents'
  const queryParams = `?limit=100&startedAfter=${dateISO8601}`
  const fullUrl = apiUrl + queryParams

  const headers = {
    Accept: 'application/json',
    'X-VO-Api-Id': voApiId,
    'X-VO-Api-Key': voApiKey
  }

  const options = {
    method: 'get',
    headers
  }

  try {
    const response = await UrlFetchApp.fetch(fullUrl, options)
    const responseData = JSON.parse(response.getContentText())

    if (responseData && responseData.incidents) {
      const incidents = responseData.incidents.map(incident => incident.service)
      const uniqueIncidents = aggregateAlerts(incidents)
      putInCache(cacheKey, JSON.stringify(uniqueIncidents), 21600) // Cache for 6 hours
      return uniqueIncidents
    } else if (responseData && responseData.error) {
      const errorMessage = `VictorOps API Error: ${responseData.error}`
      console.error(errorMessage)
      throw new Error(errorMessage)
    } else {
      const errorMessage = 'No incidents data found in VictorOps response'
      console.error(errorMessage)
      throw new Error(errorMessage)
    }
  } catch (error) {
    const errorMessage = `Failed to fetch VictorOps incidents: ${error}`
    console.error(errorMessage)
    throw new Error(errorMessage)
  }
}

// --- Document Generation Functions ---

/**
 * Prepares the document by setting up a title, subtitle, and a horizontal rule.
 * @param {GoogleAppsScript.Document.Body} body - The body of the document to prepare.
 */
function prepareDocument (body) {
  body.clear()
  appendStyledText(body, 'ALERT REVIEW', 'title')
  appendStyledText(body, 'Q3', 'subtitle')
  body.appendHorizontalRule()
}

/**
 * Writes a specific section of alerts into the document.
 * @param {GoogleAppsScript.Document.Body} body - The body of the Google Doc.
 * @param {string} sectionTitle - The section's title.
 * @param {Array} data - The alerts data to be written.
 */
async function writeAlertsSection (body, sectionTitle, fetchDataFunction) {
  try {
    const data = await fetchDataFunction()
    const totalAlertsCount = calculateTotalAlertCount(data)

    appendStyledText(body, sectionTitle, 'heading')
    appendStyledText(body, 'Last 100d days', 'subheading')
    appendStyledTable(body, ['Top 5 count', 'Subject'], data.slice(0, 5), totalAlertsCount)
    appendStyledText(body, `The top 5 alerts represent ${calculateTopAlertsPercentage(data.slice(0, 5), totalAlertsCount)}% of the total number of alerts (${totalAlertsCount}).`, 'text')

    body.appendHorizontalRule()
  } catch (error) {
    Logger.log(`Error writing alerts analysis: ${error}`)
  }
}

/**
 * Main function to generate an alert analysis report.
 */
async function generateAlertAnalysisReport () {
  try {
    const doc = DocumentApp.openById(documentId)
    const body = doc.getBody()
    prepareDocument(body)

    await Promise.all([
      writeAlertsSection(body, 'Top Root@ mail alerts', fetchAndProcessEmailAlerts),
      writeAlertsSection(body, 'Top Paging alerts', fetchAndProcessVictorOpsIncidents)
    ])
  } catch (error) {
    Logger.log(`Error running alerts analysis: ${error}`)
  }
}

generateAlertAnalysisReport()
