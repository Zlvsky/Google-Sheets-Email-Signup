/*
Google app-script to collect data via simple POST request.

https://github.com/Zlvsky
https://x.com/czaleskii

 * In order to enable this script, follow these steps:
 * 
 * From your Google Sheet, from the "Extensions" menu select "Apps Script"
 * Paste this whole file into the script code editor and hit Save icon.
 * From the "Deploy" menu, select Deploy as web app
 * Choose to execute the app as yourself, and allow "Anyone", even anonymous to execute the script.
 * Now click Deploy. You may be asked to review permissions now.
 * The URL that you get will be the webhook that you can use in your POST request anywhere.
 * You can test this webhook in your browser first by pasting it. It will say "Use POST method to send data to this URL.".
 * Last all you have to do is set up your request function in your app or website.
*/

let isNewSheet = false;
let postedData = [];
const EXCLUDE_PROPERTY = "e_gs_exclude";
const ORDER_PROPERTY = "e_gs_order";
const SHEET_NAME_PROPERTY = "e_gs_SheetName";

function doGet(e) {
  return HtmlService.createHtmlOutput(
    "Use POST method to send data to this URL."
  );
}

function doPost(e) {
  let params = JSON.stringify(e.parameter);
  params = JSON.parse(params);

  params["date"] = new Date().toISOString();
  postedData = params;
  insertToSheet(params);

  return HtmlService.createHtmlOutput();
}

const flattenObject = (ob) => {
  let toReturn = {};
  for (let i in ob) {
    if (!ob.hasOwnProperty(i)) {
      continue;
    }
    if (typeof ob[i] !== "object" || ob[i] === null) {
      // Handle null as non-object
      toReturn[i] = ob[i];
      continue;
    }
    let flatObject = flattenObject(ob[i]);
    for (let x in flatObject) {
      if (!flatObject.hasOwnProperty(x)) {
        continue;
      }
      toReturn[i + "." + x] = flatObject[x];
    }
  }
  return toReturn;
};

const getHeaders = (formSheet, keys) => {
  let headers = [];
  if (!isNewSheet) {
    const lastCol = formSheet.getLastColumn();
    if (lastCol > 0) {
      headers = formSheet.getRange(1, 1, lastCol).getValues()[0];
    }
  }
  const newHeaders = keys.filter((h) => !headers.includes(h));
  headers = [...headers, ...newHeaders];
  headers = getColumnsOrder(headers);
  headers = excludeColumns(headers);
  headers = headers.filter(
    (header) =>
      ![EXCLUDE_PROPERTY, ORDER_PROPERTY, SHEET_NAME_PROPERTY].includes(header)
  );
  return headers;
};

const getValues = (headers, flat) => {
  const values = [];
  headers.forEach((h) => values.push(flat[h] === undefined ? "" : flat[h])); // Ensure undefined values are written as empty strings
  return values;
};

const insertRowData = (sheet, row, values, bold = false) => {
  if (values.length === 0) return; // Do not attempt to write if there are no values/headers
  const currentRow = sheet.getRange(row, 1, 1, values.length);
  currentRow
    .setValues([values])
    .setFontWeight(bold ? "bold" : "normal")
    .setHorizontalAlignment("center");
};

const setHeaders = (sheet, values) => {
  if (values.length > 0) {
    // Only set headers if there are any
    insertRowData(sheet, 1, values, true);
  }
};

const setValues = (sheet, values) => {
  const lastRow = Math.max(sheet.getLastRow(), sheet.getFrozenRows()); // Consider frozen rows for new data row
  // If sheet is new or headers were just written, lastRow might be 1.
  // If headers are present, data starts at row 2.
  // If no headers (e.g. all columns excluded), data starts at row 1.
  let targetRow;
  if (sheet.getLastRow() === 0 && sheet.getFrozenRows() === 0) {
    // Completely empty sheet
    targetRow = 1;
  } else if (
    sheet.getLastRow() === 1 &&
    sheet
      .getRange(1, 1, Math.max(1, sheet.getLastColumn()))
      .getFontWeights()[0][0] === "bold"
  ) {
    // Headers likely exist at row 1
    targetRow = 2;
    if (sheet.getLastRow() < targetRow - 1) {
      // Ensure row exists before inserting after
      sheet.insertRowAfter(targetRow - 1);
    } else {
      sheet.insertRowAfter(sheet.getLastRow());
    }
  } else {
    targetRow = lastRow + 1;
    sheet.insertRowAfter(lastRow);
  }
  insertRowData(sheet, targetRow, values);
};

const getFormSheet = (sheetName) => {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let formSheet = activeSpreadsheet.getSheetByName(sheetName);
  if (formSheet == null) {
    formSheet = activeSpreadsheet.insertSheet(sheetName);
    isNewSheet = true;
  } else {
    isNewSheet = false; // Ensure isNewSheet is correctly set if sheet exists
  }
  return formSheet;
};

const insertToSheet = (data) => {
  const flat = flattenObject(data);
  const keys = Object.keys(flat);
  const targetSheetName = getSheetName(data);
  const formSheet = getFormSheet(targetSheetName);

  const headers = getHeaders(formSheet, keys);
  const values = getValues(headers, flat);

  // Only set headers if it's a new sheet or if new headers were added
  // or if existing headers don't match (e.g. order changed)
  let existingHeaders = [];
  if (
    !isNewSheet &&
    formSheet.getLastColumn() > 0 &&
    formSheet.getLastRow() > 0
  ) {
    existingHeaders = formSheet
      .getRange(1, 1, 1, formSheet.getLastColumn())
      .getValues()[0];
  }

  if (
    isNewSheet ||
    headers.length > existingHeaders.length ||
    JSON.stringify(headers) !==
      JSON.stringify(existingHeaders.slice(0, headers.length))
  ) {
    if (headers.length > 0) {
      // Check if there are any headers to set
      if (!isNewSheet && formSheet.getLastRow() > 0)
        formSheet.getRange(1, 1, formSheet.getLastColumn()).clearContent(); // Clear old headers if re-writing
      setHeaders(formSheet, headers);
    }
  }

  if (values.some((val) => val !== "" && val !== undefined)) {
    // Only insert if there's actual data
    setValues(formSheet, values);
  }
};

const getSheetName = (data) =>
  data[SHEET_NAME_PROPERTY] || data["form_name"] || "Sheet1"; // Added fallback "Sheet1"

const stringToArray = (str) => {
  if (typeof str !== "string") return [];
  return str.split(",").map((el) => el.trim());
};

const getColumnsOrder = (headers) => {
  if (!postedData[ORDER_PROPERTY]) {
    return headers;
  }
  let sortingArr = stringToArray(postedData[ORDER_PROPERTY]);
  sortingArr = sortingArr.filter((h) => headers.includes(h));
  let remainingHeaders = headers.filter((h) => !sortingArr.includes(h));
  return [...sortingArr, ...remainingHeaders];
};

const excludeColumns = (headers) => {
  if (!postedData[EXCLUDE_PROPERTY]) {
    return headers;
  }
  const columnsToExclude = stringToArray(postedData[EXCLUDE_PROPERTY]);
  return headers.filter((header) => !columnsToExclude.includes(header));
};
