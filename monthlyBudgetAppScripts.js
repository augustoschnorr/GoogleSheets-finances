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
  let installments = description.split("-")[0].toString().split("/")[1].trim();

  return {
    date,
    amount,
    description,
    category,
    paymentType,
    installments
  };
}

function getParcelDates(currentSheetName, parcels) {
  //TODO: Consider the original purchase date here
  //It'll fix the problem of duplicating data in the first month when the
  //purchase was on the previous month.
  let [currentMonth, currentYear] = currentSheetName.split("-");
  let currentMonthIndex = months.findIndex(s => s == currentMonth);
  let originalDate = new Date("20" + currentYear, currentMonthIndex, 6);
  let parcelDates = [];

  for (let parcel = 0; parcel < parcels; parcel++) {
    let parcelDate = new Date(
      originalDate.getFullYear(),
      originalDate.getMonth() + parcel,
      originalDate.getDay()
    );
    parcelDates.push(parcelDate);
  }

  return parcelDates;
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

function addParcelToSheets(parcelDates, parcelSheets, expenseInfo) {
  if (expenseInfo.installments != parcelSheets.length) {
    console.log("ERROR: Number of installments is different from the number of months!")
    return -1;
  }
  for (let i = 1; i < expenseInfo.installments; i++) {
    // Defining where to write the installment
    let sheetName = parcelSheets[i];
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    let dateValues = sheet.getRange("B4:B").getValues();
    let last = 4 + dateValues.filter(String).length;
    sheet.setActiveSelection(sheet.getRange("B" + last + ":F" + last));

    // Defining the updated description
    let installment = i + 1;
    let descr = expenseInfo.description.replace(
      `1/${expenseInfo.installments}`,
      `${installment}/${expenseInfo.installments}`
    );

    // Defining updated date
    let newDate = parcelDates[i];
    sheet.getActiveRange().setValues([
      [newDate, expenseInfo.amount, descr, expenseInfo.category, expenseInfo.paymentType]
    ]);
  }
}

function addInstallments() {
  let expenseInfo = identifyInstallment()
  let currentSheetName = SpreadsheetApp.getActive().getActiveSheet().getSheetName();

  let parcelDates = getParcelDates(currentSheetName, expenseInfo.installments);

  let parcelSheets = parcelDates.map(getSheetNameFromDate)
  let sheetsToCreate = getSheetsToCreate(parcelSheets);

  Logger.log(sheetsToCreate);

  createNewMonthSheets(sheetsToCreate);

  addParcelToSheets(parcelDates, parcelSheets, expenseInfo);
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
