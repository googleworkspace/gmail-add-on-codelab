/**
 * Copyright 2017 Google Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *   https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Finds largest dollar amount from email body.
 * Returns null if no dollar amount is found.
 *
 * @param {Message} message An email message.
 * @returns {String}
 */
function getLargestAmount(message) {
  return 'TODO';
}

/**
 * Determines date the email was received.
 *
 * @param {Message} message An email message.
 * @returns {String}
 */
function getReceivedDate(message) {
  return 'TODO';
}

/**
 * Determines expense description by joining sender name and message subject.
 *
 * @param {Message} message An email message.
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
