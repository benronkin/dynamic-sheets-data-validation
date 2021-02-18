/**
 *
 * @param {*} e
 */
// eslint-disable-next-line no-unused-vars
function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allCompanies = ss
    .getSheetByName('Companies')
    .getDataRange()
    .getValues()
    .map((row) => row[0].toLowerCase());

  const inDroppableCell =
    ss.getActiveSheet().getName() === 'Edit' &&
    e.range.getColumn() === 1 &&
    e.range.getRow() !== 1;

  if (!inDroppableCell) {
    return;
  }

  const userInput = e.range.getValue().toString().toLowerCase();

  if (userInput.toString().trim().length === 0) {
    return;
  }

  const matchedCompanies = allCompanies.filter((row) =>
    row.includes(userInput)
  );
  matchedCompanies.splice(500, matchedCompanies.length);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(matchedCompanies)
    .build();
  e.range.setDataValidation(rule);
}
