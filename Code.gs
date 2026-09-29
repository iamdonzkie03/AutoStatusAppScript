/***************************************************************
 * PROCUREMENT STATUS VALIDATOR
 * Google Sheets Apps Script
 *
 * TARGET SHEET:
 *   Data List
 *
 * IMPORTANT:
 *   1. Paste this entire code into ONE Apps Script file.
 *   2. Run setupValidator() ONCE manually.
 *   3. Authorize the script.
 *   4. After that, validation runs automatically.
 *
 * No custom Google Sheets menu is created.
 ***************************************************************/


/* ============================================================
 * CONFIGURATION
 * ============================================================ */

const CONFIG = {
  SHEET_NAME: 'Data List',

  HEADER_ROW: 1,

  STATUS_HEADER: 'STATUS',

  VALID_STATUSES: [
    'Active',
    'Closed',
    'Failed',
    'Awarded',
    'Purchase Order',
    'Realigned Item',
    'Cancelled PR',
    'Cancelled PO'
  ],

  DATE_FORMAT: 'MMMM d, yyyy',

  // Maximum number of errors displayed in one dialog.
  // Increase if necessary.
  MAX_ERRORS_DISPLAYED: 100
};


/* ============================================================
 * RULES FROM EXCEL FILE
 *
 * true/required:
 *   field must contain valid data
 *
 * false/blank:
 *   field must be blank
 *
 * zero:
 *   field must contain zero / 0.00
 *
 * before:
 *   date must be today or earlier
 *
 * after:
 *   date must be today or later
 *
 * exact:
 *   exact text is required
 * ============================================================ */

const STATUS_RULES = {

  'Active': {
    'TOTAL ABC': 'required',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'required',
    'PRE-PROCUREMENT CONFERENCE': 'required',
    'POSTING DATE': 'before',
    'PHILGEPS REFERENCE NO.': 'required',
    'PROJECT ID': 'required',
    'PRE-BID CONFERENCE': 'required',
    'ELIGIBILITY SCREENING': 'required',
    'SUBMISSION OF BIDS': 'required',
    'DETAILED BID EVALUATION': 'blank',
    'POST-QUALIFICATION': 'blank',
    'NOA DATE': 'blank',
    'NTP DATE': 'blank',
    'BAC RESOLUTION NO.': 'blank',
    'SUPPLIER': 'blank',
    'DATE PREPARED (PO)': 'blank',
    'PO NO.': 'blank',
    'PO TOTAL COST': 'zero',
    'PROJECT TITLE': 'required',
    'PROCUREMENT METHOD': 'required',
    'REMARKS': 'blank'
  },

  'Closed': {
    'TOTAL ABC': 'required',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'required',
    'PRE-PROCUREMENT CONFERENCE': 'required',
    'POSTING DATE': 'required',
    'PHILGEPS REFERENCE NO.': 'required',
    'PROJECT ID': 'required',
    'PRE-BID CONFERENCE': 'required',
    'ELIGIBILITY SCREENING': 'after',
    'SUBMISSION OF BIDS': 'after',
    'DETAILED BID EVALUATION': 'blank',
    'POST-QUALIFICATION': 'blank',
    'NOA DATE': 'blank',
    'NTP DATE': 'blank',
    'BAC RESOLUTION NO.': 'blank',
    'SUPPLIER': 'blank',
    'DATE PREPARED (PO)': 'blank',
    'PO NO.': 'blank',
    'PO TOTAL COST': 'zero',
    'PROJECT TITLE': 'required',
    'PROCUREMENT METHOD': 'required',
    'REMARKS': 'blank'
  },

  'Failed': {
    'TOTAL ABC': 'required',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'required',
    'PRE-PROCUREMENT CONFERENCE': 'required',
    'POSTING DATE': 'required',
    'PHILGEPS REFERENCE NO.': 'required',
    'PROJECT ID': 'required',
    'PRE-BID CONFERENCE': 'required',
    'ELIGIBILITY SCREENING': 'after',
    'SUBMISSION OF BIDS': 'after',
    'DETAILED BID EVALUATION': 'blank',
    'POST-QUALIFICATION': 'blank',
    'NOA DATE': 'blank',
    'NTP DATE': 'blank',
    'BAC RESOLUTION NO.': 'required',
    'SUPPLIER': 'blank',
    'DATE PREPARED (PO)': 'blank',
    'PO NO.': 'blank',
    'PO TOTAL COST': 'zero',
    'PROJECT TITLE': 'required',
    'PROCUREMENT METHOD': 'required',
    'REMARKS': 'blank'
  },

  'Awarded': {
    'TOTAL ABC': 'required',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'required',
    'PRE-PROCUREMENT CONFERENCE': 'required',
    'POSTING DATE': 'required',
    'PHILGEPS REFERENCE NO.': 'required',
    'PROJECT ID': 'required',
    'PRE-BID CONFERENCE': 'required',
    'ELIGIBILITY SCREENING': 'after',
    'SUBMISSION OF BIDS': 'after',
    'DETAILED BID EVALUATION': 'required',
    'POST-QUALIFICATION': 'required',
    'NOA DATE': 'required',
    'NTP DATE': 'required',
    'BAC RESOLUTION NO.': 'required',
    'SUPPLIER': 'required',
    'DATE PREPARED (PO)': 'blank',
    'PO NO.': 'blank',
    'PO TOTAL COST': 'zero',
    'PROJECT TITLE': 'required',
    'PROCUREMENT METHOD': 'required',
    'REMARKS': 'blank'
  },

  'Purchase Order': {
    'TOTAL ABC': 'required',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'required',
    'PRE-PROCUREMENT CONFERENCE': 'required',
    'POSTING DATE': 'required',
    'PHILGEPS REFERENCE NO.': 'required',
    'PROJECT ID': 'required',
    'PRE-BID CONFERENCE': 'required',
    'ELIGIBILITY SCREENING': 'after',
    'SUBMISSION OF BIDS': 'after',
    'DETAILED BID EVALUATION': 'required',
    'POST-QUALIFICATION': 'required',
    'NOA DATE': 'required',
    'NTP DATE': 'required',
    'BAC RESOLUTION NO.': 'required',
    'SUPPLIER': 'required',
    'DATE PREPARED (PO)': 'required',
    'PO NO.': 'required',
    'PO TOTAL COST': 'required',
    'PROJECT TITLE': 'required',
    'PROCUREMENT METHOD': 'required',
    'REMARKS': 'blank'
  },

  'Realigned Item': {
    'TOTAL ABC': 'zero',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'zero',
    'PRE-PROCUREMENT CONFERENCE': 'blank',
    'POSTING DATE': 'blank',
    'PHILGEPS REFERENCE NO.': 'blank',
    'PROJECT ID': 'blank',
    'PRE-BID CONFERENCE': 'blank',
    'ELIGIBILITY SCREENING': 'blank',
    'SUBMISSION OF BIDS': 'blank',
    'DETAILED BID EVALUATION': 'blank',
    'POST-QUALIFICATION': 'blank',
    'NOA DATE': 'blank',
    'NTP DATE': 'blank',
    'BAC RESOLUTION NO.': 'blank',
    'SUPPLIER': 'blank',
    'DATE PREPARED (PO)': 'blank',
    'PO NO.': 'blank',
    'PO TOTAL COST': 'blank',
    'PROJECT TITLE': 'blank',
    'PROCUREMENT METHOD': 'blank',
    'REMARKS': 'blank'
  },

  'Cancelled PR': {
    'TOTAL ABC': 'required',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'zero',
    'PRE-PROCUREMENT CONFERENCE': 'blank',
    'POSTING DATE': 'blank',
    'PHILGEPS REFERENCE NO.': 'blank',
    'PROJECT ID': 'blank',
    'PRE-BID CONFERENCE': 'blank',
    'ELIGIBILITY SCREENING': 'blank',
    'SUBMISSION OF BIDS': 'blank',
    'DETAILED BID EVALUATION': 'blank',
    'POST-QUALIFICATION': 'blank',
    'NOA DATE': 'blank',
    'NTP DATE': 'blank',
    'BAC RESOLUTION NO.': 'blank',
    'SUPPLIER': 'blank',
    'DATE PREPARED (PO)': 'blank',
    'PO NO.': 'blank',
    'PO TOTAL COST': 'blank',
    'PROJECT TITLE': 'blank',
    'PROCUREMENT METHOD': 'blank',
    'REMARKS': 'exact:Cancelled PR'
  },

  'Cancelled PO': {
    'TOTAL ABC': 'required',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'required',
    'PRE-PROCUREMENT CONFERENCE': 'required',
    'POSTING DATE': 'required',
    'PHILGEPS REFERENCE NO.': 'required',
    'PROJECT ID': 'required',
    'PRE-BID CONFERENCE': 'required',
    'ELIGIBILITY SCREENING': 'required',
    'SUBMISSION OF BIDS': 'required',
    'DETAILED BID EVALUATION': 'required',
    'POST-QUALIFICATION': 'required',
    'NOA DATE': 'required',
    'NTP DATE': 'required',
    'BAC RESOLUTION NO.': 'required',
    'SUPPLIER': 'required',
    'DATE PREPARED (PO)': 'required',
    'PO NO.': 'required',
    'PO TOTAL COST': 'blank',
    'PROJECT TITLE': 'required',
    'PROCUREMENT METHOD': 'required',
    'REMARKS': 'exact:Cancelled PO'
  }
};


/* ============================================================
 * FORMAT DEFINITIONS FROM EXCEL
 * ============================================================ */

const FIELD_FORMATS = {

  'TOTAL ABC': 'numeric',

  'PR NO.': 'text',

  'PR TOTAL ABC': 'numeric',

  'PRE-PROCUREMENT CONFERENCE': 'date',

  'POSTING DATE': 'date',

  'PHILGEPS REFERENCE NO.': 'numeric',

  'PROJECT ID': 'text',

  'PRE-BID CONFERENCE': 'date',

  'ELIGIBILITY SCREENING': 'date',

  'SUBMISSION OF BIDS': 'date',

  'DETAILED BID EVALUATION': 'date',

  'POST-QUALIFICATION': 'date',

  'NOA DATE': 'date',

  'NTP DATE': 'date',

  'BAC RESOLUTION NO.': 'text',

  'SUPPLIER': 'text',

  'DATE PREPARED (PO)': 'date',

  'PO NO.': 'text',

  'PO TOTAL COST': 'numeric',

  'PROJECT TITLE': 'text',

  'PROCUREMENT METHOD': 'text',

  'REMARKS': 'text'
};


/* ============================================================
 * INSTALLATION
 *
 * RUN THIS FUNCTION ONE TIME MANUALLY.
 *
 * It creates an installable onEdit trigger.
 *
 * No menu is created.
 * ============================================================ */

function setupValidator() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove duplicate validator triggers.
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {

    const handler = trigger.getHandlerFunction();

    if (handler === 'validatorOnEdit') {
      ScriptApp.deleteTrigger(trigger);
    }

  });

  // Create the installable edit trigger.
  ScriptApp.newTrigger('validatorOnEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Make sure the target sheet exists.
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet "' + CONFIG.SHEET_NAME + '" was not found.'
    );
  }

  SpreadsheetApp.getActive().toast(
    'Procurement validator installed successfully.',
    'Validator',
    5
  );
}


/* ============================================================
 * AUTOMATIC EDIT TRIGGER
 * ============================================================ */

function validatorOnEdit(e) {

  try {

    if (!e || !e.range) return;

    const range = e.range;
    const sheet = range.getSheet();

    // ONLY Data List
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;

    // Ignore header
    if (range.getRow() <= CONFIG.HEADER_ROW) return;

    const headers = getHeaders_(sheet);

    const statusColumn =
      headers.indexOf(CONFIG.STATUS_HEADER) + 1;

    if (statusColumn <= 0) return;

    const firstRow = range.getRow();
    const lastRow =
      firstRow + range.getNumRows() - 1;

    const allErrors = [];

    /*
     * Process every affected row.
     */
    for (
      let row = firstRow;
      row <= lastRow;
      row++
    ) {

      /*
       * ------------------------------------------------------
       * 1. AUTOMATICALLY DETERMINE STATUS
       * ------------------------------------------------------
       */
      const automaticStatus =
        determineAutomaticStatus_(
          sheet,
          row,
          headers
        );

      /*
       * Write the automatically determined status.
       *
       * We only write when a status can actually be determined.
       */
      if (automaticStatus) {

        const statusCell =
          sheet.getRange(
            row,
            statusColumn
          );

        const currentStatus =
          normalizeText_(
            statusCell.getDisplayValue()
          );

        if (currentStatus !== automaticStatus) {

          statusCell.setValue(
            automaticStatus
          );

        }

      }

      /*
       * ------------------------------------------------------
       * 2. GET THE RESULTING STATUS
       * ------------------------------------------------------
       */
      const status =
        normalizeText_(
          sheet
            .getRange(
              row,
              statusColumn
            )
            .getDisplayValue()
        );

      /*
       * Nothing to validate until a status has been determined.
       */
      if (!status) continue;

      /*
       * ------------------------------------------------------
       * 3. VALIDATE THE ROW
       * ------------------------------------------------------
       */
      if (!STATUS_RULES[status]) {

        allErrors.push({
          row: row,
          status: status,
          field: 'STATUS',
          message:
            'Invalid automatically determined status.'
        });

        continue;
      }

      const rowErrors =
        validateRow_(
          sheet,
          row,
          headers,
          status
        );

      rowErrors.forEach(
        error => allErrors.push(error)
      );

    }

    /*
     * --------------------------------------------------------
     * 4. SHOW ONE MODAL CONTAINING ALL ERRORS
     * --------------------------------------------------------
     */
    if (allErrors.length > 0) {

      showValidationModal_(
        allErrors
      );

    }

  } catch (error) {

    console.error(
      'validatorOnEdit error:',
      error
    );

  }
}

/* ============================================================
 * AUTOMATIC STATUS DETERMINATION
 * ============================================================ */

function determineAutomaticStatus_(
  sheet,
  row,
  headers
) {

  const data = {};

  headers.forEach(
    (header, index) => {

      if (!header) return;

      data[header] =
        sheet
          .getRange(
            row,
            index + 1
          )
          .getValue();

    }
  );


  /*
   * ----------------------------------------------------------
   * DATE-BASED STATUS TRANSITION
   * ----------------------------------------------------------
   *
   * Active:
   *   POSTING DATE is today or earlier.
   *
   * Closed / Failed / Awarded / Purchase Order:
   *   ELIGIBILITY SCREENING AND SUBMISSION OF BIDS
   *   are both after today.
   *
   * The more specific statuses are checked first so that
   * completed procurement stages are not overwritten by Closed.
   */

  const today = normalizeDate_(new Date());

  const postingDate =
    normalizeDate_(
      data['POSTING DATE']
    );

  const eligibilityDate =
    normalizeDate_(
      data['ELIGIBILITY SCREENING']
    );

  const submissionDate =
    normalizeDate_(
      data['SUBMISSION OF BIDS']
    );

  const postingIsBeforeOrToday =
    postingDate &&
    postingDate.getTime() <= today.getTime();

  const eligibilityIsAfterToday =
    eligibilityDate &&
    eligibilityDate.getTime() > today.getTime();

  const submissionIsAfterToday =
    submissionDate &&
    submissionDate.getTime() > today.getTime();

  const futureEligibilityAndBids =
    eligibilityIsAfterToday &&
    submissionIsAfterToday;


  /*
   * ----------------------------------------------------------
   * CANCELLED PR
   * ----------------------------------------------------------
   */

  if (
    textEquals_(
      data['REMARKS'],
      'Cancelled PR'
    ) &&
    isZero_(
      data['PR TOTAL ABC'],
      displayValue_(
        sheet,
        row,
        headers,
        'PR TOTAL ABC'
      )
    )
  ) {

    return 'Cancelled PR';

  }


  /*
   * ----------------------------------------------------------
   * CANCELLED PO
   * ----------------------------------------------------------
   */

  if (
    textEquals_(
      data['REMARKS'],
      'Cancelled PO'
    ) &&
    hasValue_(data['PO NO.']) &&
    !hasValue_(data['PO TOTAL COST'])
  ) {

    return 'Cancelled PO';

  }


  /*
   * ----------------------------------------------------------
   * REALIGNED ITEM
   * ----------------------------------------------------------
   */

  if (
    isZero_(
      data['TOTAL ABC'],
      displayValue_(
        sheet,
        row,
        headers,
        'TOTAL ABC'
      )
    ) &&
    isZero_(
      data['PR TOTAL ABC'],
      displayValue_(
        sheet,
        row,
        headers,
        'PR TOTAL ABC'
      )
    ) &&
    hasValue_(data['PR NO.'])
  ) {

    const realignmentFields = [
      'PRE-PROCUREMENT CONFERENCE',
      'POSTING DATE',
      'PHILGEPS REFERENCE NO.',
      'PROJECT ID',
      'PRE-BID CONFERENCE',
      'ELIGIBILITY SCREENING',
      'SUBMISSION OF BIDS',
      'DETAILED BID EVALUATION',
      'POST-QUALIFICATION',
      'NOA DATE',
      'NTP DATE',
      'BAC RESOLUTION NO.',
      'SUPPLIER',
      'DATE PREPARED (PO)',
      'PO NO.',
      'PO TOTAL COST',
      'PROJECT TITLE',
      'PROCUREMENT METHOD'
    ];

    const allBlank =
      realignmentFields.every(
        field =>
          !hasValue_(data[field])
      );

    if (allBlank) {
      return 'Realigned Item';
    }
  }


  /*
   * ----------------------------------------------------------
   * PURCHASE ORDER
   * ----------------------------------------------------------
   *
   * Requires the future-date transition AND all PO fields.
   */

  if (
    futureEligibilityAndBids &&
    hasValue_(
      data['DATE PREPARED (PO)']
    ) &&
    hasValue_(
      data['PO NO.']
    ) &&
    hasValue_(
      data['PO TOTAL COST']
    ) &&
    hasValue_(
      data['SUPPLIER']
    ) &&
    hasValue_(
      data['NOA DATE']
    ) &&
    hasValue_(
      data['NTP DATE']
    )
  ) {

    return 'Purchase Order';

  }


  /*
   * ----------------------------------------------------------
   * AWARDED
   * ----------------------------------------------------------
   *
   * Requires the future-date transition, award fields, and
   * PO TOTAL COST = 0.
   */

  if (
    futureEligibilityAndBids &&
    hasValue_(
      data['DETAILED BID EVALUATION']
    ) &&
    hasValue_(
      data['POST-QUALIFICATION']
    ) &&
    hasValue_(
      data['NOA DATE']
    ) &&
    hasValue_(
      data['NTP DATE']
    ) &&
    hasValue_(
      data['BAC RESOLUTION NO.']
    ) &&
    hasValue_(
      data['SUPPLIER']
    ) &&
    !hasValue_(
      data['DATE PREPARED (PO)']
    ) &&
    !hasValue_(
      data['PO NO.']
    ) &&
    isZero_(
      data['PO TOTAL COST'],
      displayValue_(
        sheet,
        row,
        headers,
        'PO TOTAL COST'
      )
    )
  ) {

    return 'Awarded';

  }


  /*
   * ----------------------------------------------------------
   * FAILED
   * ----------------------------------------------------------
   *
   * Requires the future-date transition, BAC resolution,
   * no supplier/award dates, and PO TOTAL COST = 0.
   */

  if (
    futureEligibilityAndBids &&
    hasValue_(
      data['BAC RESOLUTION NO.']
    ) &&
    !hasValue_(
      data['SUPPLIER']
    ) &&
    !hasValue_(
      data['DETAILED BID EVALUATION']
    ) &&
    !hasValue_(
      data['POST-QUALIFICATION']
    ) &&
    !hasValue_(
      data['NOA DATE']
    ) &&
    !hasValue_(
      data['NTP DATE']
    ) &&
    isZero_(
      data['PO TOTAL COST'],
      displayValue_(
        sheet,
        row,
        headers,
        'PO TOTAL COST'
      )
    )
  ) {

    return 'Failed';

  }


  /*
   * ----------------------------------------------------------
   * CLOSED
   * ----------------------------------------------------------
   *
   * Closed is selected when:
   *   - eligibility screening is after today
   *   - submission of bids is after today
   *   - the row has the basic procurement information
   *
   * More specific statuses above take priority.
   */

  const basicFields = [
    'TOTAL ABC',
    'PR NO.',
    'PR TOTAL ABC',
    'PRE-PROCUREMENT CONFERENCE',
    'POSTING DATE',
    'PHILGEPS REFERENCE NO.',
    'PROJECT ID',
    'PRE-BID CONFERENCE',
    'ELIGIBILITY SCREENING',
    'SUBMISSION OF BIDS',
    'PROJECT TITLE',
    'PROCUREMENT METHOD'
  ];

  const basicConditionsMet =
    basicFields.every(
      field =>
        hasValue_(data[field])
    );

  if (
    basicConditionsMet &&
    futureEligibilityAndBids
  ) {

    return 'Closed';

  }


  /*
   * ----------------------------------------------------------
   * ACTIVE
   * ----------------------------------------------------------
   *
   * Active is selected when:
   *   - POSTING DATE is today or earlier
   *   - the basic procurement information exists
   *   - PO TOTAL COST is 0 / 0.00
   */

  if (
    basicConditionsMet &&
    postingIsBeforeOrToday &&
    isZero_(
      data['PO TOTAL COST'],
      displayValue_(
        sheet,
        row,
        headers,
        'PO TOTAL COST'
      )
    )
  ) {

    return 'Active';

  }


  /*
   * ----------------------------------------------------------
   * No complete status yet.
   * ----------------------------------------------------------
   */

  return null;
}

/* ============================================================
 * AUTOMATIC STATUS HELPERS
 * ============================================================ */

function hasValue_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return false;
  }

  if (
    typeof value === 'string' &&
    value.trim() === ''
  ) {
    return false;
  }

  return true;
}


function textEquals_(
  value,
  expected
) {

  if (!hasValue_(value)) {
    return false;
  }

  return String(value)
    .trim()
    .toLowerCase() ===
    String(expected)
      .trim()
      .toLowerCase();
}


function displayValue_(
  sheet,
  row,
  headers,
  field
) {

  const index =
    headers.indexOf(field);

  if (index < 0) {
    return '';
  }

  return sheet
    .getRange(
      row,
      index + 1
    )
    .getDisplayValue();

}


/* ============================================================
 * GET HEADERS
 * ============================================================ */

function getHeaders_(sheet) {

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    return [];
  }

  return sheet
    .getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(header => normalizeText_(header));

}


/* ============================================================
 * VALIDATE ONE ROW
 * ============================================================ */

function validateRow_(
  sheet,
  row,
  headers,
  status
) {

  const errors = [];

  const rules = STATUS_RULES[status];

  headers.forEach((header, index) => {

    if (!header) {
      return;
    }

    // STATUS itself is handled separately.
    if (header === CONFIG.STATUS_HEADER) {
      return;
    }

    const rule = rules[header];

    /*
     * If the workbook doesn't define a condition for a column,
     * don't validate it.
     */
    if (!rule) {
      return;
    }

    const column = index + 1;

    const range = sheet.getRange(row, column);

    const value = range.getValue();

    const displayValue = normalizeText_(
      range.getDisplayValue()
    );

    const fieldFormat = FIELD_FORMATS[header];

    /*
     * ----------------------------------------------------------
     * REQUIRED
     * ----------------------------------------------------------
     */

    if (rule === 'required') {

      if (isBlankValue_(value, displayValue)) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message:
            'This field is required for status "' +
            status +
            '".'
        });

        return;
      }

      /*
       * A required field must ALSO have the correct format.
       */

      const formatError = validateFormat_(
        header,
        value,
        displayValue,
        fieldFormat
      );

      if (formatError) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message: formatError
        });

      }

      return;
    }


    /*
     * ----------------------------------------------------------
     * BLANK
     * ----------------------------------------------------------
     */

    if (rule === 'blank') {

      if (!isBlankValue_(value, displayValue)) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message:
            'This field must be blank for status "' +
            status +
            '".'
        });

      }

      return;
    }


    /*
     * ----------------------------------------------------------
     * ZERO
     * ----------------------------------------------------------
     */

    if (rule === 'zero') {

      if (isBlankValue_(value, displayValue)) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message:
            'This field must contain 0 or 0.00 for status "' +
            status +
            '".'
        });

        return;
      }

      if (!isZero_(value, displayValue)) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message:
            'This field must be 0 or 0.00 for status "' +
            status +
            '".'
        });

      }

      return;
    }


    /*
     * ----------------------------------------------------------
     * BEFORE TODAY
     * ----------------------------------------------------------
     */

    if (rule === 'before') {

      if (isBlankValue_(value, displayValue)) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message:
            'A date is required and it must be today or earlier.'
        });

        return;
      }

      const dateError = validateDate_(
        value,
        displayValue,
        'before'
      );

      if (dateError) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message: dateError
        });

      }

      return;
    }


    /*
     * ----------------------------------------------------------
     * AFTER TODAY
     * ----------------------------------------------------------
     */

    if (rule === 'after') {

      if (isBlankValue_(value, displayValue)) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message:
            'A date is required and it must be today or later.'
        });

        return;
      }

      const dateError = validateDate_(
        value,
        displayValue,
        'after'
      );

      if (dateError) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message: dateError
        });

      }

      return;
    }


    /*
     * ----------------------------------------------------------
     * EXACT TEXT
     * ----------------------------------------------------------
     */

    if (rule.indexOf('exact:') === 0) {

      const expected = rule.substring(6);

      if (displayValue !== expected) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message:
            'This field must contain exactly "' +
            expected +
            '".'
        });

      }

      return;
    }

  });

  return errors;
}


/* ============================================================
 * FORMAT VALIDATION
 * ============================================================ */

function validateFormat_(
  field,
  value,
  displayValue,
  format
) {

  if (!format) {
    return null;
  }


  /*
   * NUMERIC
   */

  if (format === 'numeric') {

    if (!isNumeric_(value, displayValue)) {

      return (
        'Invalid number format. This field must contain a numeric value.'
      );

    }

    return null;
  }


if (format === 'text') {

  if (!isTextValue_(displayValue)) {

    return (
      'This field cannot be blank.'
    );

  }

  return null;
}
  /*
   * DATE
   */

  if (format === 'date') {

    if (!isValidDateValue_(value, displayValue)) {

      return (
        'Invalid date format. Expected format: ' +
        CONFIG.DATE_FORMAT +
        '.'
      );

    }

    return null;
  }


  return null;
}


/* ============================================================
 * DATE VALIDATION
 * ============================================================ */

function validateDate_(
  value,
  displayValue,
  direction
) {

  if (!isValidDateValue_(value, displayValue)) {

    return (
      'Invalid date format. Expected format: ' +
      CONFIG.DATE_FORMAT +
      '.'
    );

  }

  const date = normalizeDate_(
    value
  );

  if (!date) {

    return (
      'Invalid date. Expected format: ' +
      CONFIG.DATE_FORMAT +
      '.'
    );

  }

  const today = normalizeDate_(
    new Date()
  );

  const dateNumber = date.getTime();
  const todayNumber = today.getTime();


  if (
    direction === 'before' &&
    dateNumber > todayNumber
  ) {

    return (
      'Invalid date. The date must be today or earlier.'
    );

  }


  if (
    direction === 'after' &&
    dateNumber < todayNumber
  ) {

    return (
      'Invalid date. The date must be today or later.'
    );

  }


  return null;
}


/* ============================================================
 * DATE TYPE CHECK
 * ============================================================ */

function isValidDateValue_(
  value,
  displayValue
) {

  /*
   * Best case:
   * Google Sheets stored the value as an actual Date object.
   */

  if (
    Object.prototype.toString.call(value) ===
    '[object Date]' &&
    !isNaN(value.getTime())
  ) {

    return true;

  }


  /*
   * If it is text, don't accept arbitrary JavaScript dates.
   * Require the expected displayed pattern:
   *
   * January 5, 2026
   * September 29, 2026
   */

  if (!displayValue) {
    return false;
  }

  const datePattern =
    /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$/;

  if (!datePattern.test(displayValue)) {
    return false;
  }

  const parsed = new Date(displayValue);

  return (
    !isNaN(parsed.getTime())
  );
}


/* ============================================================
 * NORMALIZE DATE
 * ============================================================ */

function normalizeDate_(dateValue) {

  const date = new Date(dateValue);

  if (isNaN(date.getTime())) {
    return null;
  }

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate()
  );
}


/* ============================================================
 * NUMERIC CHECK
 * ============================================================ */

function isNumeric_(
  value,
  displayValue
) {

  /*
   * Actual Google Sheets number.
   */

  if (
    typeof value === 'number' &&
    !isNaN(value)
  ) {

    return true;

  }


  /*
   * Numeric text.
   *
   * Supports:
   * 100
   * 100.00
   * 1,000
   * 1,000.00
   */

  if (!displayValue) {
    return false;
  }

  const cleaned = displayValue
    .replace(/,/g, '')
    .replace(/\s/g, '');

  return /^-?\d+(\.\d+)?$/.test(cleaned);
}


/* ============================================================
 * ZERO CHECK
 * ============================================================ */

function isZero_(
  value,
  displayValue
) {

  if (
    typeof value === 'number' &&
    !isNaN(value)
  ) {

    return value === 0;

  }

  if (!displayValue) {
    return false;
  }

  const cleaned = displayValue
    .replace(/,/g, '')
    .replace(/\s/g, '');

  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) {
    return false;
  }

  return Number(cleaned) === 0;
}


/* ============================================================
 * FREE-FORM TEXT CHECK
 * ============================================================ */

function isTextValue_(text) {

  /*
   * Accept any non-blank text, numbers, punctuation,
   * symbols, spaces and special characters.
   */

  if (
    text === null ||
    text === undefined
  ) {
    return false;
  }

  return String(text).trim() !== '';
}


/* ============================================================
 * BLANK CHECK
 * ============================================================ */

function isBlankValue_(
  value,
  displayValue
) {

  if (
    value === null ||
    value === undefined
  ) {

    return true;

  }

  if (
    typeof value === 'string' &&
    value.trim() === ''
  ) {

    return true;

  }

  if (
    typeof displayValue === 'string' &&
    displayValue.trim() === ''
  ) {

    return true;

  }

  return false;
}


/* ============================================================
 * NORMALIZE TEXT
 * ============================================================ */

function normalizeText_(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return '';

  }

  return String(value).trim();

}


/* ============================================================
 * HTML MODAL
 * ============================================================ */

function showValidationModal_(
  errors
) {

  /*
   * Prevent an enormous dialog.
   */

  const displayErrors =
    errors.slice(
      0,
      CONFIG.MAX_ERRORS_DISPLAYED
    );

  let html =
    '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<base target="_top">' +
    '<style>' +

    '* {' +
    '  box-sizing: border-box;' +
    '}' +

    'body {' +
    '  margin: 0;' +
    '  padding: 0;' +
    '  font-family: Arial, sans-serif;' +
    '  background: #f5f7fa;' +
    '  color: #202124;' +
    '}' +

    '.container {' +
    '  padding: 22px;' +
    '}' +

    '.header {' +
    '  background: #b3261e;' +
    '  color: white;' +
    '  margin: -22px -22px 20px -22px;' +
    '  padding: 18px 22px;' +
    '}' +

    '.header h2 {' +
    '  margin: 0 0 6px 0;' +
    '  font-size: 21px;' +
    '}' +

    '.header p {' +
    '  margin: 0;' +
    '  font-size: 13px;' +
    '  opacity: .95;' +
    '}' +

    '.summary {' +
    '  background: white;' +
    '  border-left: 5px solid #b3261e;' +
    '  padding: 13px 15px;' +
    '  margin-bottom: 16px;' +
    '  border-radius: 5px;' +
    '  box-shadow: 0 1px 3px rgba(0,0,0,.10);' +
    '}' +

    '.summary strong {' +
    '  font-size: 16px;' +
    '}' +

    '.error {' +
    '  background: white;' +
    '  margin-bottom: 10px;' +
    '  padding: 13px 15px;' +
    '  border-radius: 5px;' +
    '  border: 1px solid #dadce0;' +
    '}' +

    '.error-title {' +
    '  font-weight: bold;' +
    '  margin-bottom: 5px;' +
    '  color: #b3261e;' +
    '}' +

    '.error-message {' +
    '  font-size: 13px;' +
    '  line-height: 1.45;' +
    '}' +

    '.details {' +
    '  color: #5f6368;' +
    '  font-size: 12px;' +
    '  margin-bottom: 5px;' +
    '}' +

    '.footer {' +
    '  margin-top: 18px;' +
    '  text-align: right;' +
    '}' +

    'button {' +
    '  border: none;' +
    '  border-radius: 5px;' +
    '  padding: 10px 22px;' +
    '  background: #1a73e8;' +
    '  color: white;' +
    '  font-size: 14px;' +
    '  cursor: pointer;' +
    '}' +

    'button:hover {' +
    '  background: #1557b0;' +
    '}' +

    '.scroll {' +
    '  max-height: 390px;' +
    '  overflow-y: auto;' +
    '  padding-right: 4px;' +
    '}' +

    '</style>' +
    '</head>' +

    '<body>' +

    '<div class="container">' +

    '<div class="header">' +
    '<h2>⚠ Procurement Data Validation Error</h2>' +
    '<p>Please correct the following field(s).</p>' +
    '</div>' +

    '<div class="summary">' +
    '<strong>' +
    escapeHtml_(
      errors.length +
      ' validation error' +
      (errors.length === 1 ? '' : 's') +
      ' found'
    ) +
    '</strong>' +
    '</div>' +

    '<div class="scroll">';


  displayErrors.forEach(
    function(error) {

      html +=
        '<div class="error">' +

        '<div class="details">' +
        'Row ' +
        escapeHtml_(
          error.row
        ) +
        ' &nbsp; | &nbsp; Status: ' +
        escapeHtml_(
          error.status
        ) +
        '</div>' +

        '<div class="error-title">' +
        escapeHtml_(
          error.field
        ) +
        '</div>' +

        '<div class="error-message">' +
        escapeHtml_(
          error.message
        ) +
        '</div>' +

        '</div>';

    }
  );


  if (
    errors.length >
    CONFIG.MAX_ERRORS_DISPLAYED
  ) {

    html +=
      '<div class="error">' +
      '<div class="error-message">' +
      'Only the first ' +
      CONFIG.MAX_ERRORS_DISPLAYED +
      ' errors are displayed.' +
      '</div>' +
      '</div>';

  }


  html +=
    '</div>' +

    '<div class="footer">' +
    '<button onclick="google.script.host.close()">' +
    'Close' +
    '</button>' +
    '</div>' +

    '</div>' +

    '</body>' +
    '</html>';


  const output =
    HtmlService
      .createHtmlOutput(html)
      .setWidth(650)
      .setHeight(560);


  SpreadsheetApp
    .getUi()
    .showModalDialog(
      output,
      'Data Validation'
    );
}


/* ============================================================
 * HTML ESCAPE
 * ============================================================ */

function escapeHtml_(value) {

  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');

}


/* ============================================================
 * OPTIONAL MANUAL VALIDATION
 *
 * This function is useful for testing.
 *
 * It validates the currently active row on Data List.
 * ============================================================ */

function validateActiveRow() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet();

  if (
    sheet.getName() !==
    CONFIG.SHEET_NAME
  ) {

    SpreadsheetApp
      .getUi()
      .alert(
        'This validator only runs on the "Data List" sheet.'
      );

    return;
  }

  const headers =
    getHeaders_(sheet);

  const statusColumn =
    headers.indexOf(
      CONFIG.STATUS_HEADER
    ) + 1;

  if (statusColumn <= 0) {
    return;
  }

  const row =
    sheet
      .getActiveRange()
      .getRow();

  if (row <= CONFIG.HEADER_ROW) {
    return;
  }

  const status =
    normalizeText_(
      sheet
        .getRange(row, statusColumn)
        .getDisplayValue()
    );

  if (!status) {
    return;
  }

  const errors =
    validateRow_(
      sheet,
      row,
      headers,
      status
    );

  if (errors.length > 0) {
    showValidationModal_(errors);
  } else {
    SpreadsheetApp
      .getActive()
      .toast(
        'No validation errors found.',
        'Validation',
        3
      );
  }
}
