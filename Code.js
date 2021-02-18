const ss = SpreadsheetApp.getActiveSpreadsheet();
let allPartners = ss
  .getSheetByName('Partner Names')
  .getDataRange()
  .getValues()
  .map((row) => row[0]);

/**
 *
 * @param {*} e
 */
// eslint-disable-next-line no-unused-vars
function onEdit(e) {
  const sheet = ss.getActiveSheet();

  const inDroppableCell =
    sheet.getName() === 'View' &&
    e.range.getColumn() === 1 &&
    e.range.getRow() !== 1;

  if (!inDroppableCell) {
    return;
  }

  const userInput = e.range.getValue();

  if (userInput.toString().trim().length === 0) {
    return;
  }

  const matchedPartners = allPartners.filter((row) => row.includes(userInput));
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(matchedPartners)
    .build();
  e.range.setDataValidation(rule);
}


