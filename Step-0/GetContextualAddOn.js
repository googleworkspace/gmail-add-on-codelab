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
 * @param {Object} event Event containing the message ID and other context.
 * @returns {Card[]}
 */
function getContextualAddOn(event) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Log Your Expense'));

  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextInput()
    .setFieldName('Date')
    .setTitle('Date'));
  section.addWidget(CardService.newTextInput()
    .setFieldName('Amount')
    .setTitle('Amount'));
  section.addWidget(CardService.newTextInput()
    .setFieldName('Description')
    .setTitle('Description'));
  section.addWidget(CardService.newTextInput()
    .setFieldName('Spreadsheet URL')
    .setTitle('Spreadsheet URL'));

  card.addSection(section);

  return [card.build()];
}
