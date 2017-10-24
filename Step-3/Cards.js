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

var FIELDNAMES = ['Date', 'Amount', 'Description', 'Spreadsheet URL'];

/**
 * Creates the main card users see with form inputs to log expense.
 * Form can be prefilled with values.
 *
 * @param {String[]} opt_prefills Default values for each input field.
 * @param {String} opt_status Optional status displayed at top of card.
 * @returns {Card}
 */
function createExpensesCard(opt_prefills, opt_status) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle('Log Your Expense'));

  if (opt_status) {
    if (opt_status.indexOf('Error: ') == 0) {
      opt_status = '<font color=\'#FF0000\'>' + opt_status + '</font>';
    } else {
      opt_status = '<font color=\'#228B22\'>' + opt_status + '</font>';
    }
    var statusSection = CardService.newCardSection();
    statusSection.addWidget(CardService.newTextParagraph()
      .setText('<b>' + opt_status + '</b>'));
    card.addSection(statusSection);
  }

  var formSection = createFormSection(CardService.newCardSection(),
                                      FIELDNAMES, opt_prefills);
  card.addSection(formSection);

  return card;
}

/**
 * Creates form section to be displayed on card.
 *
 * @param {CardSection} section The card section to which form items are added.
 * @param {String[]} inputNames Names of titles for each input field.
 * @param {String[]} opt_prefills Default values for each input field.
 * @returns {CardSection}
 */
function createFormSection(section, inputNames, opt_prefills) {
  for (var i = 0; i < inputNames.length; i++) {
    var widget = CardService.newTextInput()
      .setFieldName(inputNames[i])
      .setTitle(inputNames[i]);
    if (opt_prefills && opt_prefills[i]) {
      widget.setValue(opt_prefills[i]);
    }
    section.addWidget(widget);
  }

  var submitForm = CardService.newAction().setFunctionName('submitForm');
  var submitButton = CardService.newTextButton()
    .setText('Submit')
    .setOnClickAction(submitForm);
  section.addWidget(CardService.newButtonSet().addButton(submitButton));

  return section;
}

/**
 * Logs form inputs into a spreadsheet given by URL from form.
 * Then displays edit card.
 *
 * @param {Event} e An event object containing form inputs and parameters.
 * @returns {Card}
 */
function submitForm(e) {
  var res = e['formInput'];
  try {
    FIELDNAMES.forEach(function(fieldName) {
      if (! res[fieldName]) {
        throw 'incomplete form';
      }
    });
    var sheet = SpreadsheetApp
      .openByUrl((res['Spreadsheet URL']))
      .getActiveSheet();
    sheet.appendRow(objToArray(res, FIELDNAMES.slice(0, FIELDNAMES.length - 1)));
    PropertiesService.getUserProperties().setProperty('SPREADSHEET_URL', res['Spreadsheet URL']);
    return createExpensesCard(null, 'Logged expense successfully!').build();
  }
  catch (err) {
    if (err == 'Exception: Invalid argument: url') {
      err = 'Invalid URL';
      res['Spreadsheet URL'] = null;
    }
    return createExpensesCard(objToArray(res, FIELDNAMES), 'Error: ' + err).build();
  }
}

/**
 * Returns an array corresponding to the given object and desired ordering of keys.
 *
 * @param {Object} obj Object whose values will be returned as an array.
 * @param {String[]} keys An array of key names in the desired order.
 * @returns {Object[]}
 */
function objToArray(obj, keys) {
  return keys.map(function(key) {
    return obj[key];
  });
}
