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
  var card = createExpensesCard(prefills);

  return [card.build()];
}
