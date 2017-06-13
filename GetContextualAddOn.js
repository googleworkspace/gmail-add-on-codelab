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
 * Returns the contextual add-on data that should be rendered for
 * the current e-mail thread. This function satisfies the requirements of
 * an 'onTriggerFunction' and is specified in the add-on's manifest.
 *
 * @param {String} messageId The ID of the message currently open.
 * @returns {ContextualAddOn}
 */
function getContextualAddOn(messageId) {
  var message = GmailApp.getMessageById(messageId);
  var prefills = [getReceivedDate(message),
                  getLargestAmount(message),
                  getExpenseDescription(message),
                  getSheetUrl()];
  var addOn = AddOnCardBuilder.createContextualAddOn()
      .addCard(createExpensesCard(prefills));
  return addOn.buildResult();
}
