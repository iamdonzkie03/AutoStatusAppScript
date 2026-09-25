/*******************************************************
 * PROCUREMENT MONITORING - AUTOMATIC ERROR MODAL
 *
 * PURPOSE:
 * Automatically validates a row whenever data is edited.
 * All detected errors are displayed in ONE centered modal.
 *
 * IMPORTANT:
 * Create an INSTALLABLE "On edit" trigger for:
 *     handleEdit
 *
 * Do NOT create a trigger for "onEdit".
 *******************************************************/


/**
 * Main edit handler.
 *
 * This function should be connected to an INSTALLABLE
 * "On edit" trigger.
 */
function handleEdit(e) {

  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  // Ignore header row
  if (row <= 1) return;

  // Ignore edits outside the actual data area
  if (e.range.getNumRows() > 1) {
    // Validate each affected row when pasting multiple rows
    const firstRow = e.range.getRow();
    const lastRow = firstRow + e.range.getNumRows() - 1;

    for (let r = firstRow; r <= lastRow; r++) {
      validateRowAndShowModal_(sheet, r);
    }

    return;
  }

  validateRowAndShowModal_(sheet, row);
}


/**
 * Validate one complete row.
 */
function validateRowAndShowModal_(sheet, rowNumber) {

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) return;

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const values = sheet
    .getRange(rowNumber, 1, 1, lastColumn)
    .getValues()[0];

  const displayValues = sheet
    .getRange(rowNumber, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const columns = buildColumnMap_(headers);

  const errors = [];

  const status = getValue_(values, displayValues, columns, "STATUS");

  if (!status) {
    return;
  }

  const normalizedStatus = normalize_(status);


  /********************************************************
   * ACTIVE
   ********************************************************/
  if (normalizedStatus === "active") {

    validateDateFormats_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        "POSTING DATE",
        "PRE-PROCUREMENT CONFERENCE",
        "PRE-BID CONFERENCE",
        "ELIGIBILITY SCREENING",
        "SUBMISSION OF BIDS"
      ]
    );

    validateProjectId_(errors, rowNumber, columns, values, displayValues);

    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        "POSTING DATE",
        "PRE-PROCUREMENT CONFERENCE",
        "PROJECT ID",
        "PRE-BID CONFERENCE",
        "ELIGIBILITY SCREENING",
        "SUBMISSION OF BIDS"
      ]
    );

    validateBidDatesAgainstToday_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues
    );
  }


  /********************************************************
   * CLOSED
   ********************************************************/
  if (normalizedStatus === "closed") {

    validateDateFormats_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        "POSTING DATE",
        "PRE-PROCUREMENT CONFERENCE",
        "PRE-BID CONFERENCE",
        "ELIGIBILITY SCREENING",
        "SUBMISSION OF BIDS"
      ]
    );

    validateProjectId_(errors, rowNumber, columns, values, displayValues);

    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        "POSTING DATE",
        "PRE-PROCUREMENT CONFERENCE",
        "PROJECT ID",
        "PRE-BID CONFERENCE",
        "ELIGIBILITY SCREENING",
        "SUBMISSION OF BIDS"
      ]
    );
  }


  /********************************************************
   * AWARDED
   ********************************************************/
  if (normalizedStatus === "awarded") {

    const requiredAwardedFields = [
      "PROJECT ID",
      "PRE-PROCUREMENT CONFERENCE",
      "PRE-BID CONFERENCE",
      "POSTING DATE",
      "ELIGIBILITY SCREENING",
      "SUBMISSION OF BIDS",
      "POST-QUALIFICATION",
      "DETAILED BID EVALUATION",
      "NOA DATE",
      "SUPPLIER"
    ];

    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      requiredAwardedFields
    );

    validateDateFormats_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        "POSTING DATE",
        "PRE-PROCUREMENT CONFERENCE",
        "PRE-BID CONFERENCE",
        "ELIGIBILITY SCREENING",
        "SUBMISSION OF BIDS",
        "POST-QUALIFICATION",
        "DETAILED BID EVALUATION",
        "NOA DATE"
      ]
    );

    validateProjectId_(errors, rowNumber, columns, values, displayValues);

    validateSupplier_(errors, rowNumber, columns, values, displayValues);

    validateBidDatesAgainstToday_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues
    );
  }


  /********************************************************
   * FAILED
   ********************************************************/
  if (normalizedStatus === "failed") {

    const requiredFailedFields = [
      "PROJECT ID",
      "PRE-PROCUREMENT CONFERENCE",
      "PRE-BID CONFERENCE",
      "POSTING DATE",
      "ELIGIBILITY SCREENING",
      "SUBMISSION OF BIDS",
      "POST-QUALIFICATION",
      "DETAILED BID EVALUATION",
      "NOA DATE",
      "SUPPLIER"
    ];

    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      requiredFailedFields
    );

    validateDateFormats_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        "POSTING DATE",
        "PRE-PROCUREMENT CONFERENCE",
        "PRE-BID CONFERENCE",
        "ELIGIBILITY SCREENING",
        "SUBMISSION OF BIDS",
        "POST-QUALIFICATION",
        "DETAILED BID EVALUATION",
        "NOA DATE"
      ]
    );

    validateProjectId_(errors, rowNumber, columns, values, displayValues);

    validateSupplier_(errors, rowNumber, columns, values, displayValues);
  }


  /********************************************************
   * PURCHASE ORDER
   ********************************************************/
  if (
    normalizedStatus === "purchase order" ||
    normalizedStatus === "purchaseorder"
  ) {

    const requiredPOFields = [
      "PROJECT ID",
      "POSTING DATE",
      "ELIGIBILITY SCREENING",
      "SUBMISSION OF BIDS",
      "POST-QUALIFICATION",
      "DETAILED BID EVALUATION",
      "NOA DATE",
      "SUPPLIER",
      "PO DATE",
      "PO NO.",
      "PO TOTAL COST"
    ];

    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      requiredPOFields
    );

    validateDateFormats_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        "POSTING DATE",
        "ELIGIBILITY SCREENING",
        "SUBMISSION OF BIDS",
        "POST-QUALIFICATION",
        "DETAILED BID EVALUATION",
        "NOA DATE",
        "NTP DATE",
        "PO DATE"
      ]
    );

    validateProjectId_(errors, rowNumber, columns, values, displayValues);

    validateSupplier_(errors, rowNumber, columns, values, displayValues);

    validateNumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      "PO TOTAL COST"
    );

    validatePOStatusFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues
    );
  }


  /********************************************************
   * TOTAL ABC
   ********************************************************/
  validateNumericField_(
    errors,
    rowNumber,
    columns,
    values,
    displayValues,
    "TOTAL ABC"
  );


  /********************************************************
   * SHOW ALL ERRORS IN ONE MODAL
   ********************************************************/
  if (errors.length > 0) {

    const uniqueErrors = removeDuplicateErrors_(errors);

    showErrorModal_(
      sheet.getName(),
      rowNumber,
      uniqueErrors
    );
  }
}


/**
 * Creates a header -> column number map.
 */
function buildColumnMap_(headers) {

  const map = {};

  headers.forEach(function(header, index) {

    if (!header) return;

    const normalized = normalizeHeader_(header);

    if (normalized) {
      map[normalized] = index;
    }
  });

  return map;
}


/**
 * Normalize normal values.
 */
function normalize_(value) {

  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}


/**
 * Normalize column headers.
 */
function normalizeHeader_(header) {

  return String(header || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ")
    .replace(/\s*\.\s*$/, ".");
}


/**
 * Get a value based on the column name.
 */
function getValue_(values, displayValues, columns, header) {

  const key = normalizeHeader_(header);

  if (columns[key] === undefined) {
    return "";
  }

  const index = columns[key];

  return displayValues[index] !== ""
    ? displayValues[index]
    : values[index];
}


/**
 * Check whether a field is blank.
 */
function isBlank_(values, displayValues, columns, header) {

  return String(
    getValue_(values, displayValues, columns, header)
  ).trim() === "";
}


/**
 * Required field validation.
 */
function validateRequiredFields_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues,
  fields
) {

  fields.forEach(function(field) {

    if (columns[normalizeHeader_(field)] === undefined) {
      return;
    }

    if (
      isBlank_(
        values,
        displayValues,
        columns,
        field
      )
    ) {

      addError_(
        errors,
        rowNumber,
        field,
        "There is no data inputted."
      );
    }
  });
}


/**
 * Date validation.
 *
 * Accepted:
 * - Actual Google Sheets date
 * - MMMM d, YYYY
 * - N/A
 * - NA
 * - Not Applicable
 */
function validateDateFormats_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues,
  fields
) {

  fields.forEach(function(field) {

    const key = normalizeHeader_(field);

    if (columns[key] === undefined) {
      return;
    }

    const index = columns[key];

    const rawValue = values[index];
    const displayValue = String(displayValues[index] || "").trim();

    // Blank is handled separately by required-field rules
    if (displayValue === "") {
      return;
    }

    if (isValidDateValue_(rawValue, displayValue)) {
      return;
    }

    addError_(
      errors,
      rowNumber,
      field,
      "Please use correct date format."
    );
  });
}


/**
 * Determines if a date value is valid.
 */
function isValidDateValue_(rawValue, displayValue) {

  // Google Sheets date object
  if (
    Object.prototype.toString.call(rawValue) === "[object Date]" &&
    !isNaN(rawValue.getTime())
  ) {
    return true;
  }

  const normalized = displayValue
    .trim()
    .toUpperCase();

  // Allowed non-date values
  if (
    normalized === "N/A" ||
    normalized === "NA" ||
    normalized === "NOT APPLICABLE"
  ) {
    return true;
  }

  /*
   * Required date format:
   * January 1, 2026
   * September 25, 2026
   */
  const datePattern =
    /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$/i;

  if (!datePattern.test(displayValue)) {
    return false;
  }

  const parsed = new Date(displayValue);

  return !isNaN(parsed.getTime());
}


/**
 * Project ID validation.
 */
function validateProjectId_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues
) {

  const key = "PROJECT ID";

  if (columns[key] === undefined) {
    return;
  }

  const value = String(
    getValue_(
      values,
      displayValues,
      columns,
      key
    )
  ).trim();

  if (value === "") {
    return;
  }

  // Alphanumeric validation
  if (!/^[A-Za-z0-9\-_\/. ]+$/.test(value)) {

    addError_(
      errors,
      rowNumber,
      key,
      "Please use a valid alphanumeric Project ID."
    );
  }
}


/**
 * Supplier validation.
 */
function validateSupplier_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues
) {

  const key = "SUPPLIER";

  if (columns[key] === undefined) {
    return;
  }

  const value = String(
    getValue_(
      values,
      displayValues,
      columns,
      key
    )
  ).trim();

  if (value === "") {
    return;
  }

  if (!/^[A-Za-z0-9\-&,.'()\/ ]+$/.test(value)) {

    addError_(
      errors,
      rowNumber,
      key,
      "Please use a valid alphanumeric supplier name."
    );
  }
}


/**
 * Numeric field validation.
 *
 * Accepts:
 * 1
 * 1.00
 * 1000000.50
 * 1,000,000.50
 */
function validateNumericField_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues,
  field
) {

  const key = normalizeHeader_(field);

  if (columns[key] === undefined) {
    return;
  }

  const index = columns[key];

  const rawValue = values[index];
  const displayValue = String(displayValues[index] || "").trim();

  if (displayValue === "") {
    return;
  }

  // Actual numeric spreadsheet value
  if (
    typeof rawValue === "number" &&
    !isNaN(rawValue)
  ) {
    return;
  }

  const cleaned = displayValue
    .replace(/,/g, "")
    .replace(/\s/g, "");

  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) {

    addError_(
      errors,
      rowNumber,
      field,
      "Please enter a numeric value with decimals."
    );
  }
}


/**
 * Check Eligibility Screening and Submission of Bids
 * against today's date.
 */
function validateBidDatesAgainstToday_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues
) {

  const fields = [
    "ELIGIBILITY SCREENING",
    "SUBMISSION OF BIDS"
  ];

  const today = new Date();

  today.setHours(0, 0, 0, 0);

  fields.forEach(function(field) {

    const key = normalizeHeader_(field);

    if (columns[key] === undefined) {
      return;
    }

    const index = columns[key];

    const rawValue = values[index];
    const displayValue = String(displayValues[index] || "").trim();

    if (displayValue === "") {
      return;
    }

    // N/A is not compared with today's date
    const normalized = displayValue.toUpperCase();

    if (
      normalized === "N/A" ||
      normalized === "NA" ||
      normalized === "NOT APPLICABLE"
    ) {
      return;
    }

    let dateValue = null;

    if (
      Object.prototype.toString.call(rawValue) === "[object Date]" &&
      !isNaN(rawValue.getTime())
    ) {
      dateValue = new Date(rawValue);
    } else if (
      isValidDateValue_(rawValue, displayValue)
    ) {
      dateValue = new Date(displayValue);
    }

    if (!dateValue) {
      return;
    }

    dateValue.setHours(0, 0, 0, 0);

    /*
     * According to the Excel conditions:
     *
     * Active:
     * Eligibility Screening and Submission of Bids
     * should not already be on/before today.
     *
     * If after today, it is applicable.
     */
    if (dateValue <= today) {

      addError_(
        errors,
        rowNumber,
        field,
        "This date is on or before today and is not applicable for the current status."
      );
    }
  });
}


/**
 * Purchase Order-specific validations.
 */
function validatePOStatusFields_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues
) {

  const fields = [
    "PO DATE"
  ];

  fields.forEach(function(field) {

    const key = normalizeHeader_(field);

    if (columns[key] === undefined) {
      return;
    }

    if (
      isBlank_(
        values,
        displayValues,
        columns,
        field
      )
    ) {
      return;
    }

    const index = columns[key];

    if (
      !isValidDateValue_(
        values[index],
        String(displayValues[index] || "")
      )
    ) {

      addError_(
        errors,
        rowNumber,
        field,
        "Please use correct date format."
      );
    }
  });

  // PO Number
  const poNoKey = "PO NO.";

  if (columns[poNoKey] !== undefined) {

    const poNo = String(
      getValue_(
        values,
        displayValues,
        columns,
        poNoKey
      )
    ).trim();

    if (
      poNo !== "" &&
      !/^[A-Za-z0-9\-\/. ]+$/.test(poNo)
    ) {

      addError_(
        errors,
        rowNumber,
        poNoKey,
        "Please use a valid alphanumeric PO Number."
      );
    }
  }
}


/**
 * Add an error to the error collection.
 */
function addError_(
  errors,
  rowNumber,
  column,
  message
) {

  errors.push({
    row: rowNumber,
    column: column,
    message: message
  });
}


/**
 * Remove duplicate errors.
 */
function removeDuplicateErrors_(errors) {

  const seen = {};
  const result = [];

  errors.forEach(function(error) {

    const key =
      error.row +
      "|" +
      error.column +
      "|" +
      error.message;

    if (!seen[key]) {

      seen[key] = true;
      result.push(error);
    }
  });

  return result;
}


/**
 * Display the error modal.
 */
function showErrorModal_(
  sheetName,
  rowNumber,
  errors
) {

  const template =
    HtmlService.createTemplateFromFile("ErrorModal");

  template.sheetName = sheetName;
  template.rowNumber = rowNumber;
  template.errors = errors;

  const html = template
    .evaluate()
    .setWidth(650)
    .setHeight(500);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      "⚠ Procurement Data Validation"
    );
}
