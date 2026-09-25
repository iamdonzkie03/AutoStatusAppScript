/*******************************************************
 * PROCUREMENT DATA VALIDATOR
 * Based on the conditions in Book1.xlsx
 *
 * IMPORTANT:
 * 1. Row 1 must contain your column headers.
 * 2. Header names should correspond to the names below.
 * 3. Run setupValidationTrigger() ONCE manually.
 * 4. Do NOT use a simple trigger named onEdit() for the popup.
 *    This script uses an INSTALLABLE onEdit trigger.
 *******************************************************/

const CONFIG = {

  // Sheet to validate.
  // Leave blank ("") to validate the active sheet.
  SHEET_NAME: "",

  HEADER_ROW: 1,

  STATUS_HEADER: "STATUS",

  // Values accepted as "not applicable"
  NOT_APPLICABLE: [
    "N/A",
    "NA",
    "NOT APPLICABLE"
  ],

  // Date fields from the Excel file.
  DATE_FIELDS: [
    "POSTING DATE",
    "PRE-BID CONFERENCE",
    "ELIGIBILITY SCREENING",
    "SUBMISSION OF BIDS",
    "DETAILED BID EVALUATION",
    "POST-QUALIFICATION",
    "RECOMMENDATION OF AWARD",
    "NOA DATE",
    "NTP DATE",
    "DATE PREPARED (PO)",
    "PO DATE"
  ],

  // Numeric fields
  NUMERIC_FIELDS: [
    "TOTAL ABC",
    "PR TOTAL ABC",
    "PO TOTAL COST"
  ],

  // Alphanumeric fields
  TEXT_FIELDS: [
    "PR NO.",
    "PHILGEPS REFERENCE NO.",
    "PROJECT ID",
    "BAC RESOLUTION NO.",
    "PO NO.",
    "SUPPLIER",
    "PROCUREMENT METHOD",
    "PROCUREMENT METHOD PO",
    "APP CLASSIFICATION",
    "PROJECT TITLE",
    "REMARKS"
  ]
};


/*******************************************************
 * INSTALLABLE TRIGGER SETUP
 *
 * Run this function manually ONE TIME.
 *******************************************************/
function setupValidationTrigger() {

  const ss = SpreadsheetApp.getActive();

  // Remove old copies of this trigger
  ScriptApp.getProjectTriggers().forEach(function(trigger) {

    if (trigger.getHandlerFunction() === "procurementOnEdit") {
      ScriptApp.deleteTrigger(trigger);
    }

  });

  // Create new installable onEdit trigger
  ScriptApp.newTrigger("procurementOnEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    "Validation trigger installed successfully.\n\n" +
    "The procurement validation will now run automatically whenever data is edited."
  );
}


/*******************************************************
 * MAIN EDIT FUNCTION
 *******************************************************/
function procurementOnEdit(e) {

  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();

  // If a specific sheet was configured, ignore other sheets.
  if (
    CONFIG.SHEET_NAME &&
    sheet.getName() !== CONFIG.SHEET_NAME
  ) {
    return;
  }

  // Ignore header row
  if (range.getRow() <= CONFIG.HEADER_ROW) {
    return;
  }

  const firstRow = range.getRow();
  const numberOfRows = range.getNumRows();

  const headers = getHeaders_(sheet);

  if (!headers.length) {
    return;
  }

  let allErrors = [];

  // Validate every affected row.
  for (let i = 0; i < numberOfRows; i++) {

    const rowNumber = firstRow + i;

    const rowValues = sheet
      .getRange(rowNumber, 1, 1, headers.length)
      .getValues()[0];

    const errors = validateRow_(
      sheet,
      rowNumber,
      headers,
      rowValues
    );

    if (errors.length) {

      allErrors.push({
        row: rowNumber,
        errors: errors
      });

    }
  }

  if (!allErrors.length) {
    return;
  }

  // Display ALL errors in one modal.
  showValidationModal_(allErrors);
}


/*******************************************************
 * GET HEADERS
 *******************************************************/
function getHeaders_(sheet) {

  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    return [];
  }

  const values = sheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn)
    .getDisplayValues()[0];

  return values.map(function(header) {

    return normalizeHeader_(header);

  });
}


/*******************************************************
 * NORMALIZE HEADER
 *******************************************************/
function normalizeHeader_(value) {

  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}


/*******************************************************
 * CREATE HEADER MAP
 *******************************************************/
function createHeaderMap_(headers) {

  const map = {};

  headers.forEach(function(header, index) {

    if (header) {
      map[header] = index;
    }

  });

  return map;
}


/*******************************************************
 * VALIDATE ONE ROW
 *******************************************************/
function validateRow_(
  sheet,
  rowNumber,
  headers,
  rowValues
) {

  const errors = [];

  const headerMap = createHeaderMap_(headers);

  const status = getField_(
    headerMap,
    rowValues,
    "STATUS"
  );

  const normalizedStatus =
    String(status || "")
      .trim()
      .toUpperCase();

  if (!normalizedStatus) {
    return errors;
  }


  /*****************************************************
   * GENERAL FORMAT VALIDATION
   *
   * Validate ALL applicable fields so the modal
   * reports every invalid column.
   *****************************************************/

  CONFIG.DATE_FIELDS.forEach(function(field) {

    if (!hasColumn_(headerMap, field)) {
      return;
    }

    const value = getField_(
      headerMap,
      rowValues,
      field
    );

    if (isBlank_(value)) {
      return;
    }

    if (isNotApplicable_(value)) {
      return;
    }

    if (!isValidDate_(value)) {

      errors.push({
        field: field,
        message: "Please use correct date format."
      });

    }

  });


  /*****************************************************
   * NUMERIC VALIDATION
   *****************************************************/

  CONFIG.NUMERIC_FIELDS.forEach(function(field) {

    if (!hasColumn_(headerMap, field)) {
      return;
    }

    const value = getField_(
      headerMap,
      rowValues,
      field
    );

    if (isBlank_(value)) {
      return;
    }

    if (!isValidNumeric_(value)) {

      errors.push({
        field: field,
        message: "Please enter a valid numeric value."
      });

    }

  });


  /*****************************************************
   * STATUS-SPECIFIC RULES
   *****************************************************/

  switch (normalizedStatus) {

    case "ACTIVE":

      validateActive_(
        headerMap,
        rowValues,
        errors
      );

      break;


    case "CLOSED":

      validateClosed_(
        headerMap,
        rowValues,
        errors
      );

      break;


    case "AWARDED":

      validateAwarded_(
        headerMap,
        rowValues,
        errors
      );

      break;


    case "FAILED":

      validateFailed_(
        headerMap,
        rowValues,
        errors
      );

      break;


    case "PURCHASE ORDER":

      validatePurchaseOrder_(
        headerMap,
        rowValues,
        errors
      );

      break;

  }


  return removeDuplicateErrors_(errors);
}


/*******************************************************
 * ACTIVE STATUS
 *******************************************************/
function validateActive_(
  headerMap,
  rowValues,
  errors
) {

  /*
   * Excel:
   *
   * If Posting Date is blank
   * If Eligibility Screening is blank
   * If Submission of Bids is blank
   * If Project ID is blank
   */

  validateRequired_(
    headerMap,
    rowValues,
    "POSTING DATE",
    "There is no data inputted.",
    errors
  );

  validateRequired_(
    headerMap,
    rowValues,
    "ELIGIBILITY SCREENING",
    "There is no data inputted.",
    errors
  );

  validateRequired_(
    headerMap,
    rowValues,
    "SUBMISSION OF BIDS",
    "There is no data inputted.",
    errors
  );

  validateRequired_(
    headerMap,
    rowValues,
    "PROJECT ID",
    "Please input the Project ID",
    errors
  );


  /*
   * Eligibility Screening and Submission of Bids
   * should not be on or before today.
   *
   * If they are on/before today:
   * "Not applicable"
   *
   * This is interpreted as the Excel condition.
   */

  validateActiveDateLogic_(
    headerMap,
    rowValues,
    errors
  );
}


/*******************************************************
 * CLOSED STATUS
 *******************************************************/
function validateClosed_(
  headerMap,
  rowValues,
  errors
) {

  validateRequired_(
    headerMap,
    rowValues,
    "POSTING DATE",
    "There is no data inputted.",
    errors
  );

  validateRequired_(
    headerMap,
    rowValues,
    "ELIGIBILITY SCREENING",
    "There is no data inputted.",
    errors
  );

  validateRequired_(
    headerMap,
    rowValues,
    "SUBMISSION OF BIDS",
    "There is no data inputted.",
    errors
  );

  validateRequired_(
    headerMap,
    rowValues,
    "PROJECT ID",
    "Please input the Project ID",
    errors
  );


  /*
   * Closed status generally means the bidding period
   * has already progressed/passed.
   */
}


/*******************************************************
 * AWARDED STATUS
 *******************************************************/
function validateAwarded_(
  headerMap,
  rowValues,
  errors
) {

  const requiredFields = [

    "PROJECT ID",
    "POSTING DATE",
    "ELIGIBILITY SCREENING",
    "SUBMISSION OF BIDS",
    "DETAILED BID EVALUATION",
    "POST-QUALIFICATION",
    "NOA DATE",
    "SUPPLIER"

  ];


  requiredFields.forEach(function(field) {

    validateRequired_(
      headerMap,
      rowValues,
      field,
      "There is no inputted data.",
      errors
    );

  });


  /*
   * Date fields must have valid formats.
   *
   * General date validation already runs above.
   */
}


/*******************************************************
 * FAILED STATUS
 *******************************************************/
function validateFailed_(
  headerMap,
  rowValues,
  errors
) {

  const requiredFields = [

    "PROJECT ID",
    "POSTING DATE",
    "ELIGIBILITY SCREENING",
    "SUBMISSION OF BIDS",
    "DETAILED BID EVALUATION",
    "POST-QUALIFICATION",
    "NOA DATE",
    "SUPPLIER"

  ];


  requiredFields.forEach(function(field) {

    validateRequired_(
      headerMap,
      rowValues,
      field,
      "There is no inputted data.",
      errors
    );

  });

}


/*******************************************************
 * PURCHASE ORDER STATUS
 *******************************************************/
function validatePurchaseOrder_(
  headerMap,
  rowValues,
  errors
) {

  const requiredFields = [

    "PROJECT ID",
    "POSTING DATE",
    "ELIGIBILITY SCREENING",
    "SUBMISSION OF BIDS",
    "DETAILED BID EVALUATION",
    "POST-QUALIFICATION",
    "NOA DATE",
    "NTP DATE",
    "PO DATE",
    "PO NO.",
    "PO TOTAL COST",
    "SUPPLIER"

  ];


  requiredFields.forEach(function(field) {

    validateRequired_(
      headerMap,
      rowValues,
      field,
      "There is no inputted data.",
      errors
    );

  });

}


/*******************************************************
 * REQUIRED FIELD VALIDATION
 *******************************************************/
function validateRequired_(
  headerMap,
  rowValues,
  field,
  message,
  errors
) {

  if (!hasColumn_(headerMap, field)) {
    return;
  }

  const value = getField_(
    headerMap,
    rowValues,
    field
  );

  if (isBlank_(value)) {

    errors.push({
      field: field,
      message: message
    });

  }
}


/*******************************************************
 * ACTIVE DATE LOGIC
 *******************************************************/
function validateActiveDateLogic_(
  headerMap,
  rowValues,
  errors
) {

  const eligibility =
    getField_(
      headerMap,
      rowValues,
      "ELIGIBILITY SCREENING"
    );

  const submission =
    getField_(
      headerMap,
      rowValues,
      "SUBMISSION OF BIDS"
    );


  if (
    isBlank_(eligibility) ||
    isBlank_(submission)
  ) {
    return;
  }


  if (
    isNotApplicable_(eligibility) ||
    isNotApplicable_(submission)
  ) {
    return;
  }


  const eligibilityDate =
    convertToDate_(eligibility);

  const submissionDate =
    convertToDate_(submission);


  if (!eligibilityDate || !submissionDate) {
    return;
  }


  const today = startOfDay_(
    new Date()
  );


  /*
   * Excel condition:
   * If Eligibility Screening and Submission
   * of Bids are on or before today.
   */

  if (
    eligibilityDate <= today &&
    submissionDate <= today
  ) {

    errors.push({
      field:
        "ELIGIBILITY SCREENING & SUBMISSION OF BIDS",
      message:
        "Not applicable"
    });

  }


  /*
   * Excel also contains:
   * If Eligibility Screening and Submission
   * of Bids are after today.
   *
   * We retain this logic as an informational
   * validation condition.
   */

}


/*******************************************************
 * DATE VALIDATION
 *******************************************************/
function isValidDate_(value) {

  // Real Google Sheets Date value
  if (Object.prototype.toString.call(value) === "[object Date]") {

    return !isNaN(value.getTime());

  }


  const text =
    String(value || "").trim();


  if (!text) {
    return false;
  }


  if (isNotApplicable_(text)) {
    return true;
  }


  /*
   * Expected:
   *
   * January 5, 2026
   * September 25, 2026
   *
   * Format:
   * MMMM d, YYYY
   */

  const regex =
    /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$/i;


  if (!regex.test(text)) {
    return false;
  }


  const parsed =
    new Date(text);


  if (isNaN(parsed.getTime())) {
    return false;
  }


  // Prevent JavaScript from accepting
  // invalid dates such as February 31.
  const parts =
    text.match(
      /^([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})$/
    );


  if (!parts) {
    return false;
  }


  const monthName = parts[1];
  const day = Number(parts[2]);
  const year = Number(parts[3]);


  const monthIndex =
    new Date(
      monthName + " 1, " + year
    ).getMonth();


  return (
    parsed.getFullYear() === year &&
    parsed.getMonth() === monthIndex &&
    parsed.getDate() === day
  );
}


/*******************************************************
 * CONVERT VALUE TO DATE
 *******************************************************/
function convertToDate_(value) {

  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {

    if (isNaN(value.getTime())) {
      return null;
    }

    return startOfDay_(value);
  }


  const text =
    String(value || "").trim();


  if (!text || isNotApplicable_(text)) {
    return null;
  }


  const date =
    new Date(text);


  if (isNaN(date.getTime())) {
    return null;
  }


  return startOfDay_(date);
}


/*******************************************************
 * NUMERIC VALIDATION
 *******************************************************/
function isValidNumeric_(value) {

  if (
    typeof value === "number" &&
    !isNaN(value)
  ) {
    return true;
  }


  const text =
    String(value || "")
      .trim()
      .replace(/,/g, "");


  if (!text) {
    return false;
  }


  /*
   * Accept:
   *
   * 100
   * 100.00
   * 1,000.50
   * 1000.50
   */

  return /^-?\d+(\.\d+)?$/.test(text);
}


/*******************************************************
 * NOT APPLICABLE
 *******************************************************/
function isNotApplicable_(value) {

  const text =
    String(value || "")
      .trim()
      .toUpperCase();


  return CONFIG.NOT_APPLICABLE.indexOf(text) !== -1;
}


/*******************************************************
 * BLANK CHECK
 *******************************************************/
function isBlank_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return true;
  }


  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {

    return isNaN(value.getTime());

  }


  return String(value).trim() === "";
}


/*******************************************************
 * GET FIELD
 *******************************************************/
function getField_(
  headerMap,
  rowValues,
  field
) {

  const normalized =
    normalizeHeader_(field);


  if (
    headerMap[normalized] === undefined
  ) {
    return "";
  }


  return rowValues[
    headerMap[normalized]
  ];
}


/*******************************************************
 * COLUMN EXISTS
 *******************************************************/
function hasColumn_(
  headerMap,
  field
) {

  return (
    headerMap[
      normalizeHeader_(field)
    ] !== undefined
  );
}


/*******************************************************
 * START OF DAY
 *******************************************************/
function startOfDay_(date) {

  const d =
    new Date(date);

  d.setHours(
    0,
    0,
    0,
    0
  );

  return d;
}


/*******************************************************
 * REMOVE DUPLICATE ERRORS
 *******************************************************/
function removeDuplicateErrors_(errors) {

  const seen = {};

  return errors.filter(function(error) {

    const key =
      error.field +
      "|" +
      error.message;


    if (seen[key]) {
      return false;
    }


    seen[key] = true;

    return true;

  });
}


/*******************************************************
 * MODAL POPUP
 *******************************************************/
function showValidationModal_(allErrors) {

  let html = `
    <html>
      <head>
        <style>

          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            color: #222;
          }

          h2 {
            margin-top: 0;
            color: #b00020;
          }

          .intro {
            margin-bottom: 18px;
          }

          .row {
            border: 1px solid #ddd;
            border-radius: 6px;
            padding: 12px;
            margin-bottom: 12px;
          }

          .row-title {
            font-weight: bold;
            margin-bottom: 8px;
          }

          .error {
            padding: 6px 0;
            border-bottom: 1px solid #eee;
          }

          .error:last-child {
            border-bottom: none;
          }

          .field {
            font-weight: bold;
          }

          .message {
            margin-left: 8px;
          }

          button {
            margin-top: 15px;
            padding: 8px 20px;
            border: none;
            border-radius: 4px;
            background: #1a73e8;
            color: white;
            cursor: pointer;
          }

        </style>
      </head>

      <body>

        <h2>Procurement Data Validation</h2>

        <div class="intro">
          Please review the following data issues:
        </div>
  `;


  allErrors.forEach(function(rowData) {

    html += `
      <div class="row">

        <div class="row-title">
          Row ${rowData.row}
        </div>
    `;


    rowData.errors.forEach(function(error) {

      html += `
        <div class="error">

          <span class="field">
            ${escapeHtml_(error.field)}
          </span>

          <span class="message">
            ${escapeHtml_(error.message)}
          </span>

        </div>
      `;

    });


    html += `
      </div>
    `;

  });


  html += `

        <button onclick="google.script.host.close()">
          Close
        </button>

      </body>
    </html>
  `;


  const output =
    HtmlService
      .createHtmlOutput(html)
      .setWidth(600)
      .setHeight(500);


  SpreadsheetApp
    .getUi()
    .showModalDialog(
      output,
      "Validation Warning"
    );
}


/*******************************************************
 * ESCAPE HTML
 *******************************************************/
function escapeHtml_(text) {

  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
