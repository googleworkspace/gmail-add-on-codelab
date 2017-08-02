/**
 * Finds largest dollar amount from email body.
 * Returns null if no dollar amount is found.
 *
 * @param {String} messageId The ID of the message currently open.
 * @returns {String}
 */
function getLargestAmount(message) {
  return 'TODO';
}

/**
 * Determines date the email was received.
 *
 * @param {String} messageId The ID of the message currently open.
 * @returns {String}
 */
function getReceivedDate(message) {
  return 'TODO';
}

/**
 * Determines expense description by joining sender name and message subject.
 *
 * @param {String} messageId The ID of the message currently open.
 * @returns {String}
 */
function getExpenseDescription(message) {
  return 'TODO';
}

/**
 * Determines most recent spreadsheet URL.
 * Returns null if no URL was previously submitted.
 *
 * @returns {String}
 */
function getSheetUrl() {
  return PropertiesService.getUserProperties().getProperty('SPREADSHEET_URL');
}
