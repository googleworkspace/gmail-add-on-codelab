/**
 * Finds largest dollar amount from email body.
 * Returns null if no dollar amount is found.
 *
 * @param {String} message The message that is currently open.
 * @returns {String}
 */
function getLargestAmount(message) {
  var amount = 0;
  var messageBody = message.getPlainBody();
  var regex = /\$[\d,]+\.\d\d/g;
  var match = regex.exec(messageBody);
  while (match) {
    amount = Math.max(amount, parseFloat(match[0].substring(1).replace(/,/g,'')));
    match = regex.exec(messageBody);
  }
  return amount ? '$' + amount.toFixed(2).toString() : null;
}

/**
 * Determines date the email was received.
 *
 * @param {String} message The message that is currently open.
 * @returns {String}
 */
function getReceivedDate(message) {
  return message.getDate().toLocaleDateString();
}

/**
 * Determines expense description by joining sender name and message subject.
 *
 * @param {String} message The message that is currently open.
 * @returns {String}
 */
function getExpenseDescription(message) {
  var sender = message.getFrom();
  var subject = message.getSubject();
  return sender + " | " + subject;
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
