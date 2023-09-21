var months = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];

function identifyInstallment() {
  let currentRange = SpreadsheetApp.getActive().getActiveRange();

  let date = currentRange.getCell(1, 1).getValue();
  let amount = currentRange.getCell(1, 2).getValue();
  let description = currentRange.getCell(1, 3).getValue();
  let category = currentRange.getCell(1, 4).getValue();
  let paymentType = currentRange.getCell(1, 5).getValue();
  //let installments = description.split("-")[0].toString().split("/")[1].trim();
  let installments = description.split("-")[0].trim();

  return {
    date,
    amount,
    description,
    category,
    paymentType,
    installments
  };
}

function getParcelSheets(currentSheetName, parcels) {
  let [currentMonth, currentYear] = currentSheetName.split("-");
  let currentMonthIndex = months.findIndex(s => s == currentMonth);
  let originalDate = new Date("20" + currentYear, currentMonthIndex);
  let parcelSheetsNames = [];

  for (let parcel = 1; parcel < parcels; parcel++) {
    let parcelDate = new Date(
      originalDate.getFullYear(),
      originalDate.getMonth() + parcel,
      originalDate.getDay() + 1 // handles months with 31 days --'
    ); // Using Date to get next month because it handles December -> January
    // transitioning transparently
    parcelSheetsNames.push(getSheetNameFromDate(parcelDate));
  }

  return parcelSheetsNames;
}

function getSheetNameFromDate(date) {
  let month = months[date.getMonth()];
  let year = date.getFullYear().toString().slice(2, 4);
  return month + '-' + year;
}

function getSheetsToCreate(parcelSheets) {
  let existingSheets = SpreadsheetApp.getActive().getSheets().map(s => s.getSheetName());
  let diff = parcelSheets.filter(x => !existingSheets.includes(x));
  return diff;
}

function createNewMonthSheets(sheetsToCreate) {
  sheetsToCreate.forEach(sheetName => {
    let spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
    // Creates the new month sheet
    let sheet = spreadSheet.getSheetByName("Template").copyTo(spreadSheet)
    sheet.setName(sheetName);

    // Sets the month name in the new sheet
    sheet.setActiveSelection(sheet.getRange('J1'));
    let month = sheetName.split("-")[0];
    spreadSheet.getActiveCell().setValue(month);

    // Adds recurrent expenses
    //let anualExpenses = if (month == "January") //TODO: Add anual expenses automatically?
    // Adds monthly recurrent expenses - REMOVED: It's better to add them manually for better control
    //let monthlyExpenses = getMonthlyRecurrentExpenses(sheetName);

    //sheet.setActiveSelection(sheet.getRange(row=5, column=2, numRows=monthlyExpenses.length, numColumns=5));
    //spreadSheet.getActiveRange().setValues(monthlyExpenses);

    // Moves the new sheet to the last position
    spreadSheet.setActiveSheet(sheet);
    spreadSheet.moveActiveSheet(spreadSheet.getSheets().length)
  });
}

function addParcelToSheets(parcelSheets, expenseInfo) {
  let [originalInstallment, totalInstallments] = expenseInfo.installments.split("/")

  if (totalInstallments - 1 != parcelSheets.length) {
    console.log("ERROR: Number of installments is different from the number of months!")
    return -1;
  }

  for (let i = 0; i < parcelSheets.length; i++) {
    // Defining where to write the installment
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(parcelSheets[i]);
    let dateValues = sheet.getRange("B4:B").getValues();
    let lastRow = 4 + dateValues.filter(String).length; // getLastRow wouldn't work for
    // sheets that has very few records (less than templated categories)
    sheet.setActiveSelection(sheet.getRange("B" + lastRow + ":F" + lastRow));

    // Defining the updated description
    let installment = parseInt(originalInstallment) + i + 1;
    let descr = expenseInfo.description.replace(
      `${originalInstallment}/${totalInstallments}`,
      `${installment}/${totalInstallments}`
    );

    sheet.getActiveRange().setValues([
      [expenseInfo.date, expenseInfo.amount, descr, expenseInfo.category, expenseInfo.paymentType]
    ]);
  }
}

function addInstallments() {
  let expenseInfo = identifyInstallment()
  let currentSheetName = SpreadsheetApp.getActive().getActiveSheet().getSheetName();

  let parcelSheets = getParcelSheets(currentSheetName, expenseInfo.installments.split("/")[1]);
  let sheetsToCreate = getSheetsToCreate(parcelSheets);

  Logger.log(sheetsToCreate);

  createNewMonthSheets(sheetsToCreate);

  addParcelToSheets(parcelSheets, expenseInfo);
}

function getRecurrentExpenses() {
  let plannedExpensesSheet = SpreadsheetApp.getActive().getSheetByName("Planned Expenses");
  // Gets the range from the 4th row and 1st column (A4) until the 40th row 7th column (G40)
  // Then filters out the empty rows
  let plannedExpenses = plannedExpensesSheet.getRange(4, 1, 40, 8).getValues().filter(
    row => !row.every(cell => cell == "")
  )
  return plannedExpenses;
}

/**
 * Gets the monthly expenses
 * @param {string} currentDate - The month-year like the sheet name (May-22)
 */
function getMonthlyRecurrentExpenses(currentDate) {
  let recurrentExpenses = getRecurrentExpenses();
  let monthlyExpenses = recurrentExpenses.filter(row => row[4] == "Monthly").filter(row => row[7] == true);

  [m, y] = currentDate.split("-");
  let month = months.indexOf(m);
  let year = `20${y}`;
  let expenseDate = `${year}-${month + 1}-06`

  // Filter out recurrences that already reached the end date
  let firstDayOfMonth = new Date(year, month, 1);

  console.log(firstDayOfMonth);
  let monthlyExpensesFiltered = monthlyExpenses.filter(
    row => new Date(row[6]).getTime() >= firstDayOfMonth.getTime()
  );

  console.log(monthlyExpensesFiltered);

  let formatedExpenses = monthlyExpensesFiltered.map(row => {
    return [expenseDate, row[1], row[0], row[2], row[3]];
  });
  return formatedExpenses;
}

function addMonthlyRecurrentExpenses() {
  let spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  let currentSheet = spreadSheet.getActiveSheet();

  let monthlyExpenses = getMonthlyRecurrentExpenses(currentSheet.getSheetName());

  let dateValues = currentSheet.getRange("B4:B").getValues();
  let firstEmptyRow = 4 + dateValues.filter(String).length;

  currentSheet.setActiveSelection(currentSheet.getRange(
    row = firstEmptyRow,
    column = 2,
    numRows = monthlyExpenses.length,
    numColumns = 5)
  );

  spreadSheet.getActiveRange().setValues(monthlyExpenses);
}

function addRecurrentExpenses() {

  let anualExpenses = recurrentExpenses.filter(row => row[4] == "Anual");


  console.log(plannedExpenses);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Budget Funcions')
    .addItem('Add installment purchase', 'addInstallments')
    .addItem('Add monthly recurrent expenses', 'addMonthlyRecurrentExpenses')
    .addToUi();
}
