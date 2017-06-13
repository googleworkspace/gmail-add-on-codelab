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
 * @param {Array.<String>} opt_prefills Default values for each input field.
 * @param {String} opt_status Optional status displayed at top of card.
 * @returns {Card}
 */
function createExpensesCard(opt_prefills, opt_status) {
  var card = AddOnCardBuilder.createCard()
      .setHeader(AddOnCardBuilder
          .createCardHeader()
          .setTitle('Log Your Expense'));
  if (opt_status) {
    if (opt_status.indexOf('Error: ') == 0) {
      opt_status = '<font color=\'#FF0000\'>' + opt_status + '</font>';
    } else {
      opt_status = '<font color=\'#228B22\'>' + opt_status + '</font>';
    }
    card.addSection(AddOnCardBuilder
        .createCardSection()
        .addWidget(AddOnCardBuilder
            .createTextParagraph('<b>' + opt_status + '</b>')));
  }
  var id = PropertiesService.getUserProperties().getProperty('EXPENSE_ID');
  if (! id) {
    id = '0';
    PropertiesService.getUserProperties().setProperty('EXPENSE_ID', '0');
  }
  var formSection = addFields(AddOnCardBuilder
      .createCardSection()
      .addWidget(AddOnCardBuilder.createTextParagraph('Expense ID #' + id)),
                              FIELDNAMES, opt_prefills);
  return card
      .addSection(formSection
          .addWidget(AddOnCardBuilder
              .createButtonSet()
              .addButton(AddOnCardBuilder
                  .createTextButton('Submit')
                  .setOnClickAction('submitForm'))))
      .addSection(AddOnCardBuilder
          .createCardSection()
          .addWidget(AddOnCardBuilder.createTextInput('Sheet Name', 'Sheet Name'))
          .addWidget(AddOnCardBuilder
              .createButtonSet()
              .addButton(AddOnCardBuilder
                  .createTextButton('New Sheet')
                  .setOnClickAction('createExpensesSheet'))))
      .addAction(AddOnCardBuilder
          .createCardAction()
          .setText('Clear form')
          .setOnClickAction('clearForm', {'Status': opt_status ? opt_status : ''}));
}

/**
 * Creates form section to be displayed on card.
 *
 * @param {CardSection} section The card section to which form items are added.
 * @param {Array.<String>} inputNames Names of titles for each input field.
 * @param {Array.<String>} opt_prefills Default values for each input field.
 * @returns {CardSection}
 */
function addFields(section, inputNames, opt_prefills) {
  for (var i = 0; i < inputNames.length; i++) {
    var widget = AddOnCardBuilder.createTextInput(inputNames[i], inputNames[i]);
    if (opt_prefills && opt_prefills[i]) {
      widget.setValue(opt_prefills[i]);
    }
    section.addWidget(widget);
  }
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
    var props = PropertiesService.getUserProperties();
    var id = props.getProperty('EXPENSE_ID');
    sheet.appendRow([id].concat(objToArray(res, FIELDNAMES.slice(0, FIELDNAMES.length - 1))));
    props.setProperty('EXPENSE_ID', (parseInt(id) + 1).toString());
    props.setProperty('SPREADSHEET_URL', res['Spreadsheet URL']);
    return createEditCard(res, 'Logged expense successfully!')
        .buildResult();
  }
  catch (err) {
    if (err == 'Exception: Invalid argument: url') {
      err = 'Invalid URL';
      res['Spreadsheet URL'] = null;
    }
    return createExpensesCard(objToArray(res, FIELDNAMES), 'Error: ' + err)
        .buildResult();
  }
}

/**
 * Creates a new spreadsheet for the user to log expenses.
 *
 * @param {Event} e An event object containing form inputs and parameters.
 * @returns {Card}
 */
function createExpensesSheet(e) {
  var res = e['formInput'];
  var sheetName = res['Sheet Name'] ? res['Sheet Name'] : 'Expenses';
  newSpreadsheet = SpreadsheetApp.create(sheetName);
  newSpreadsheet.setFrozenRows(1);
  newSpreadsheet
      .getActiveSheet()
      .getRange(1, 1, 1, FIELDNAMES.length)
      .setValues([['Expense ID'].concat(FIELDNAMES.slice(0, FIELDNAMES.length - 1))]);
  newSpreadsheet
      .getRange('Sheet1!A:A')
      .protect()
      .setDescription('IDs Protected')
      .setWarningOnly(true);
  var prefills = objToArray(res, FIELDNAMES.slice(0, FIELDNAMES.length - 1));
  prefills.push(newSpreadsheet.getUrl());
  return createExpensesCard(prefills, 'Created and linked the spreadsheet <i>' +
                            sheetName + '</i> for expenses!')
      .buildResult();
}

/**
 * Recreates the main card without prefilled data.
 *
 * @param {Event} e An event object containing form inputs and parameters.
 * @returns {Card}
 */
function clearForm(e) {
  return createExpensesCard(null, e['parameters']['Status'])
      .buildResult();
}

/**
 * Creates a card letting users edit the most recent expense.
 * Form can be prefilled with values.
 *
 * @param {Array.<String>} prevResults Default values for each input field.
 * @param {String} opt_status Optional status displayed at top of card.
 * @returns {Card}
 */
function createEditCard(prevResults, opt_status) {
  if (prevResults) {
    var prefills = objToArray(prevResults, FIELDNAMES.slice(0, FIELDNAMES.length - 1));
  }
  var card = AddOnCardBuilder.createCard()
      .setHeader(AddOnCardBuilder
          .createCardHeader()
          .setTitle('Edit Your Expense'));
  if (opt_status) {
    if (opt_status.indexOf('Error: ') == 0) {
      opt_status = '<font color=\'#FF0000\'>' + opt_status + '</font>';
    } else {
      opt_status = '<font color=\'#228B22\'>' + opt_status + '</font>';
    }
    card.addSection(AddOnCardBuilder
        .createCardSection()
        .addWidget(AddOnCardBuilder
            .createTextParagraph('<b>' + opt_status + '</b>')))
  }
  var id = (parseInt(PropertiesService.getUserProperties().getProperty('EXPENSE_ID')) - 1)
      .toString();
  var formSection = addFields(AddOnCardBuilder
    .createCardSection()
    .addWidget(AddOnCardBuilder.createTextParagraph('Expense ID #' + id)),
                              FIELDNAMES.slice(0, FIELDNAMES.length - 1), prefills);
  return card
      .addSection(formSection
          .addWidget(AddOnCardBuilder
              .createButtonSet()
              .addButton(AddOnCardBuilder
                  .createTextButton('Edit')
                  .setOnClickAction('editForm'))))
      .addSection(AddOnCardBuilder
          .createCardSection()
          .addWidget(spreadsheetButton('Open Spreadsheet')))
      .addAction(AddOnCardBuilder
          .createCardAction()
          .setText('Clear form')
          .setOnClickAction('clearEditForm', {'Status': opt_status ? opt_status : ''}));
}

/**
 * Edits most recent spreadsheet by removing bottom row and appending new information to sheet.
 * Then displays a edit card again.
 *
 * @param {Event} e An event object containing form inputs and parameters.
 * @returns {Card}
 */
function editForm(e) {
  var res = e['formInput'];
  try {
    FIELDNAMES.slice(0, FIELDNAMES.length - 1).forEach(function(fieldName) {
      if (! res[fieldName]) {
        throw 'incomplete form';
      }
    });
    var props = PropertiesService.getUserProperties();
    var url = props.getProperty('SPREADSHEET_URL');
    var sheet = SpreadsheetApp.openByUrl(url).getActiveSheet();
    var id = parseInt(props.getProperty('EXPENSE_ID')) - 1;
    for (var i = sheet.getLastRow(); i > 0; i--) {
      if (sheet.getRange(i, 1).getValue() == id) {
        var newValues = objToArray(res, FIELDNAMES.slice(0, FIELDNAMES.length - 1));
        sheet.getRange(i, 2, 1, FIELDNAMES.length - 1).setValues([newValues]);
        return createEditCard(res, 'Edited expense successfully!')
            .buildResult();
      }
    }
    throw 'expense ID not found in sheet';
  }
  catch (err) {
    return createEditCard(res, 'Error: ' + err).buildResult();
  }
}

/**
 * Recreates the edit card without prefilled data.
 *
 * @param {Event} e Callback event object, which contains status.
 * @returns {Card}
 */
function clearEditForm(e) {
  return createEditCard(null, e['parameters']['Status'])
      .buildResult();
}

/**
 * Returns a widget linking to spreadsheet.
 *
 * @param {String} text Text displayed on button.
 * @returns {ButtonSet}
 */
function spreadsheetButton(text) {
  return AddOnCardBuilder
              .createButtonSet()
              .addButton(AddOnCardBuilder
                  .createTextButton(text)
                  .setOnClickUrl(PropertiesService
                      .getUserProperties()
                      .getProperty('SPREADSHEET_URL')));
}

/**
 * Returns an array corresponding to the given object and desired ordering of keys.
 *
 * @param {Object} obj Objected whose values will be returned as an array.
 * @param {Array.<string>} keys An array of key names in the desired order.
 * @returns {Array}
 */
function objToArray(obj, keys) {
  return keys.map(function(key) {
    return obj[key];
  });
}
