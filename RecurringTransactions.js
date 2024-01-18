const ss = SpreadsheetApp.getActiveSpreadsheet();
const currencies = trimArray(ss.getRangeByName("Currencies").getValues()).reduce((ary, tfx) => ( { ...ary, [tfx[0]]: tfx[1] }), {});

function onOpen() {
  var entries = [
    {
      name : "Update Budget",
      functionName : "calculateRecurringTransactionsForDate"
    },{
      name : "Sort Recurring Transactions",
      functionName : "sortRecurringTransactions"
    }];
  ss.addMenu("Budget", entries);
};

/** Calculates all transactions for a date range.
*
* @customfunction
**/

function calculateRecurringTransactionsForDate() {
  ss.toast("Updating...");
  
  const budgetDates = ss.getRange("BudgetDates").getValues();
  const budgetSubtotals = ss.getRange("BudgetRecurringSubtotals");
  const rx = trimArray(ss.getRange("RecurringTransactions").getValues());
  
  for (let j = 0; j < budgetDates.length; j++) {
    let result = 0;
    let comment = "";

    for (let i = 0; i < rx.length; i++) {
      if ((budgetDates[j][0] <= rx[i][6] || rx[i][6] == "") && matchFreq(rx[i][4], budgetDates[j][0], rx[i][0])) {
        let curr = rx[i][2], amt = rx[i][3], acct = rx[i][5], asterisks = "";
        result += amt * currencies[curr];
        if (acct == "RBC" || acct == "CIBC") {
          asterisks = "**";
        }
        comment += `${rx[i][1]}${asterisks}, ${currencyFormat(amt)} ${curr}\n`;
      }
    }
    
    budgetSubtotals.getCell(j + 1, 1).setValue(result).setNote(comment);
  }
  
  sortRecurringTransactions();
  
  ss.toast("Update complete.");
};

function sortRecurringTransactions() { 
  const recurringTransactions = ss.getRange("RecurringTransactions");
  const matrix = getValuesOrFormulas(recurringTransactions);
  const transactionFrequencies = trimArray(ss.getRange("TransactionFrequencies").getValues());
  
  // Credit: https://dev.to/afewminutesofcode/how-to-create-a-custom-sort-order-in-javascript-3j1p
  const customSort = ({data, sortBy, sortField}) => {
    const sortByObject = sortBy.reduce((obj, item, index) => ({ ...obj, [item]: index }), {});
    return data.sort((a, b) => sortByObject[a[sortField]] - sortByObject[b[sortField]] || a[0] - b[0]);
  };

  const customSorted = customSort({data:matrix, sortBy:transactionFrequencies, sortField:4});

  recurringTransactions.setValues(customSorted);
};

/**
* Returns TRUE if dates meet frequency requirements, else FALSE
*
* @param {string} freq Frequency of match.
* @param {Date} dt Current date.
* @param {Date} recDt Date of first occurrence of transaction.
* @return TRUE if dates meet frequency requirements, else FALSE
**/

function matchFreq (freq, dt, recDt) {
  if (dt < recDt) return false; // The recurrence hasn't started yet

  switch(freq) {
    case "Once":
      return dt.getTime() === recDt.getTime(); // Days match exactly
      break;
    case "Weekly":
      return dt.getDay() == recDt.getDay(); // Days of the week match
      break;
    case "Biweekly":
      return subtractDays(dt, recDt) % 14 == 0; // The number of days between the two dates being compared is evenly divided by 14
      break;
    case "Monthly":
      return (isLastDayOfMonth(dt) && isAfterLastDay(dt, recDt)) || dt.getDate() == recDt.getDate(); // If the recurring transaction date is after the last day of the current month, we consider the dates to be matching OR The days of the month match
      break;
    case "Biweekly after 15":
      return dt.getDate() > 13 && dt.getDate() < 28 && subtractDays(dt, recDt) % 14 == 0; // The date is greater than the 13th and less than the 28th (this accounts for the 15th falling on a Sat) and the number of days between the two dates being compared is evenly divided by 14
      break;
    case "Bimonthly":
      return dt.getDate() == recDt.getDate() && (dt.getMonth() - recDt.getMonth()) % 2 == 0; // The days of the month match and the months are of the same parity
      break;
    case "Quarterly":
      return dt.getDate() == recDt.getDate() && (dt.getMonth() - recDt.getMonth()) % 3 == 0; // The days of the month match and the months are quarters apart
      break;
    case "Triannual":
      return dt.getDate() == recDt.getDate() && (dt.getMonth() - recDt.getMonth()) % 4 == 0; // The days of the month match and the months are four months apart
      break;
    case "Semiannual":
      return dt.getDate() == recDt.getDate() && (dt.getMonth() - recDt.getMonth()) % 6 == 0; // The days of the month match and the months are six months apart
      break;
    case "Annual":
      return dt.getDate() == recDt.getDate() && dt.getMonth() == recDt.getMonth(); // The days of the month match, and the months match
      break;
    default:
      return false;
  }
};

/** UTILS **/


function currencyFormat(num) {
  return num >= 0 ? '$' + num.toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : '($' + (num * -1).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') + ')';
};

function getValuesOrFormulas(range) {
  const rows = range.getNumRows();
  const cols = range.getNumColumns();
  const matrix = new Array(rows).fill().map(() => new Array(cols).fill());
  
  for (let i = 1; i < rows + 1; i++) {
    if (range.getCell(i, 1).getValue() == "") break;
    for (let j = 1; j < cols + 1; j++) {
      const thisCell = range.getCell(i, j);
      matrix[i-1][j-1] = thisCell.getFormula() != "" ? thisCell.getFormula() : thisCell.getValue();
    }
  }
  
  return matrix;
};

/**
* I can't believe this isn't a standard Javascript library
*
* @param {Date} dt1 The first date.
* @param {Date} dt2 The second date.
* @return The number of days between the two given dates.
* @customfunction
**/

function subtractDays(dt1, dt2) {
  return Math.round((dt1.getTime()-dt2.getTime())/(24*3600*1000)); 
};

function getLastDayOfMonth(dt) {
  return new Date(dt.getFullYear(), dt.getMonth() + 1, 0).getDate();
};

function isLastDayOfMonth(dt) {
  let d = new Date(dt.getFullYear(), dt.getMonth() + 1, 0);
  return d.getDate() == dt.getDate();
};

function isAfterLastDay(dt, recDt) {
  if (!isLeapYear(dt.getFullYear()) && dt.getMonth() == 1 && recDt.getDate() > 28) {
    return true;
  }
  if (isLeapYear(dt.getFullYear()) && dt.getMonth() == 1 && recDt.getDate() > 29) {
    return true;
  }
  if ((dt.getMonth() == 3 || dt.getMonth() == 5 || dt.getMonth() == 8 || dt.getMonth() == 10) && recDt.getDate() > 30) {
    return true;
  }
  return false;
};

function isLeapYear(year) {
  return ((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0);
};

function trimArray(ary) {
  return ary.filter(a => a[0] != "");
};
