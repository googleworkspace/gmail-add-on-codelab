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
 * @param {Array.<String>} inputNames Names of titles for each input field.
 * @param {Array.<String>} opt_prefills Default values for each input field.
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
  return section;
}
