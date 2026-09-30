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
 *   date must be strictly before today
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

  'Closed': {
    'TOTAL ABC': 'required',
    'PR NO.': 'required',
    'PR TOTAL ABC': 'required',
    'PRE-PROCUREMENT CONFERENCE': 'required',
    'POSTING DATE': 'required',
    'PHILGEPS REFERENCE NO.': 'required',
    'PROJECT ID': 'required',
    'PRE-BID CONFERENCE': 'required',
    'ELIGIBILITY SCREENING': 'before',
    'SUBMISSION OF BIDS': 'before',
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

const NA_ALLOWED_FIELDS = new Set([
  'PRE-PROCUREMENT CONFERENCE',
  'POSTING DATE',
  'PRE-BID CONFERENCE',
  'ELIGIBILITY SCREENING',
  'SUBMISSION OF BIDS',
  'DETAILED BID EVALUATION',
  'POST-QUALIFICATION'
]);

function isNAValue_(value, displayValue) {
  const text = normalizeText_(
    displayValue !== undefined ? displayValue : value
  ).toUpperCase();

  return text === 'NA' || text === 'N/A';
}

function isAcceptedValue_(field, value, displayValue) {
  return NA_ALLOWED_FIELDS.has(field) && isNAValue_(value, displayValue);
}


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
  if (!ss) throw new Error('No active spreadsheet found.');

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + CONFIG.SHEET_NAME + '" was not found.');
  }

  // Remove legacy/competing validator triggers.
  // Automatic STATUS calculation is intentionally handled only by the
  // simple onEdit(e) trigger so every editor gets the same behavior.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (
      handler === 'validatorOnEdit' ||
      handler === 'validatorOnOpen' ||
      handler === 'validatorOnChange'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Only the refresh-on-open handler remains installable.
  // It is not part of the edit-time STATUS engine.
  ScriptApp.newTrigger('validatorOnOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  // Refresh existing rows immediately.
  refreshAllStatuses_(sheet);

  SpreadsheetApp.flush();

  ss.toast(
    'Automatic STATUS is enabled for all editors. Validation is separated from the edit-time status engine.',
    'Validator',
    5
  );
}

/* ============================================================
 * SIMPLE ON EDIT - STATUS ENGINE
 *
 * This function deliberately contains NO modal/UI logic.
 * It is responsible only for changing STATUS.
 * ============================================================ */

function onEdit(e) {
  // Simple triggers run in the context of the user who edited the sheet.
  // Keep this path lightweight and use it ONLY for automatic STATUS
  // calculation so shared editors do not depend on the trigger owner's
  // authorization.
  try {
    if (!e || !e.range) return;

    const range = e.range;
    const sheet = range.getSheet();

    if (sheet.getName() !== CONFIG.SHEET_NAME) return;
    if (range.getRow() <= CONFIG.HEADER_ROW) return;

    const headers = getHeaders_(sheet);
    const statusColumn = headers.indexOf(CONFIG.STATUS_HEADER) + 1;

    if (statusColumn <= 0) return;

    const firstRow = range.getRow();
    const lastRow = firstRow + range.getNumRows() - 1;

    for (let row = firstRow; row <= lastRow; row++) {
      updateAutomaticStatusFast_(sheet, row, headers, statusColumn);
    }
  } catch (error) {
    console.error('onEdit status engine error:', error);
  }
}


/* ============================================================
 * SIMPLE ON OPEN - STATUS ENGINE
 * ============================================================ */

function onOpen(e) {
  // The installable validatorOnOpen trigger handles status refresh.
}


/* ============================================================
 * STATUS UPDATE FOR ONE ROW
 * ============================================================ */

function updateRowStatus_(sheet, row, headers, statusColumn) {
  updateAutomaticStatusFast_(sheet, row, headers, statusColumn);
}

/**
 * Fast automatic status engine used by the simple onEdit trigger.
 *
 * This intentionally does not validate the row and does not open any UI.
 * It only determines and writes the automatic procurement status.
 */
function updateAutomaticStatusFast_(sheet, row, headers, statusColumn) {
  if (row <= CONFIG.HEADER_ROW) return;

  const lastColumn = sheet.getLastColumn();

  // One read for the entire row.
  const rowValues = sheet
    .getRange(row, 1, 1, lastColumn)
    .getValues()[0];

  const rowDisplayValues = sheet
    .getRange(row, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const data = {};
  const displayData = {};

  headers.forEach((header, index) => {
    if (!header) return;
    data[header] = rowValues[index];
    displayData[header] = rowDisplayValues[index];
  });

  const automaticStatus = determineAutomaticStatusFromData_(
    data,
    displayData
  );

  const statusCell = sheet.getRange(row, statusColumn);
  const currentStatus = normalizeText_(statusCell.getDisplayValue());

  if (!automaticStatus) {
    if (currentStatus !== '') {
      statusCell.clearContent();
    }
    return;
  }

  if (currentStatus !== automaticStatus) {
    statusCell.setValue(automaticStatus);
  }
}

/**
 * Shared data-only status calculation.
 *
 * The existing determineAutomaticStatus_() remains available for the
 * validation/legacy path. The simple-trigger path uses this function
 * so it never performs one getRange() call per field.
 */
function determineAutomaticStatusFromData_(data, displayData) {
  const rowHasAnyData = Object.keys(data).some(function(header) {
    return header !== CONFIG.STATUS_HEADER && hasValue_(data[header]);
  });

  if (!rowHasAnyData) return '';

  const today = normalizeDate_(new Date());

  const postingDate = isNAValue_('', displayData['POSTING DATE'])
    ? null
    : normalizeDate_(data['POSTING DATE']);

  const eligibilityDate = isNAValue_('', displayData['ELIGIBILITY SCREENING'])
    ? null
    : normalizeDate_(data['ELIGIBILITY SCREENING']);

  const submissionDate = isNAValue_('', displayData['SUBMISSION OF BIDS'])
    ? null
    : normalizeDate_(data['SUBMISSION OF BIDS']);

  const postingIsBeforeOrToday =
    postingDate && postingDate.getTime() <= today.getTime();

  const eligibilityIsBeforeToday =
    eligibilityDate && eligibilityDate.getTime() < today.getTime();

  const submissionIsBeforeToday =
    submissionDate && submissionDate.getTime() < today.getTime();

  const eligibilityIsTodayOrAfter =
    eligibilityDate && eligibilityDate.getTime() >= today.getTime();

  const submissionIsTodayOrAfter =
    submissionDate && submissionDate.getTime() >= today.getTime();

  const activeDateCondition =
    eligibilityIsTodayOrAfter === true ||
    submissionIsTodayOrAfter === true;

  const closedDateCondition =
    eligibilityIsBeforeToday === true &&
    submissionIsBeforeToday === true;

  const futureEligibilityAndBids =
    eligibilityDate &&
    submissionDate &&
    eligibilityDate.getTime() > today.getTime() &&
    submissionDate.getTime() > today.getTime();

  // Cancelled/realigned statuses are checked before the normal
  // procurement progression.
  if (
    textEquals_(data['REMARKS'], 'Cancelled PR') &&
    isZero_(data['PR TOTAL ABC'], displayData['PR TOTAL ABC'])
  ) {
    return 'Cancelled PR';
  }

  if (
    textEquals_(data['REMARKS'], 'Cancelled PO') &&
    hasValue_(data['PO NO.']) &&
    !hasValue_(data['PO TOTAL COST'])
  ) {
    return 'Cancelled PO';
  }

  if (
    isZero_(data['TOTAL ABC'], displayData['TOTAL ABC']) &&
    isZero_(data['PR TOTAL ABC'], displayData['PR TOTAL ABC']) &&
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
      'PROCUREMENT METHOD',
      'REMARKS'
    ];

    const realigned = realignmentFields.every(function(field) {
      return isBlankValue_(data[field], displayData[field]);
    });

    if (realigned) return 'Realigned Item';
  }

  const basicConditionsMet =
    hasValue_(data['TOTAL ABC']) &&
    hasValue_(data['PR NO.']) &&
    hasValue_(data['PR TOTAL ABC']) &&
    hasValue_(data['PRE-PROCUREMENT CONFERENCE']) &&
    hasValue_(data['POSTING DATE']) &&
    hasValue_(data['PHILGEPS REFERENCE NO.']) &&
    hasValue_(data['PROJECT ID']) &&
    hasValue_(data['PRE-BID CONFERENCE']) &&
    hasValue_(data['ELIGIBILITY SCREENING']) &&
    hasValue_(data['SUBMISSION OF BIDS']) &&
    hasValue_(data['PROJECT TITLE']) &&
    hasValue_(data['PROCUREMENT METHOD']);

  if (
    basicConditionsMet &&
    postingIsBeforeOrToday === true &&
    activeDateCondition === true &&
    isZero_(data['PO TOTAL COST'], displayData['PO TOTAL COST'])
  ) {
    return 'Active';
  }

  if (
    basicConditionsMet &&
    closedDateCondition &&
    isZero_(data['PO TOTAL COST'], displayData['PO TOTAL COST'])
  ) {
    return 'Closed';
  }

  if (
    basicConditionsMet &&
    futureEligibilityAndBids &&
    hasValue_(data['BAC RESOLUTION NO.']) &&
    isBlankValue_(data['SUPPLIER'], displayData['SUPPLIER']) &&
    isZero_(data['PO TOTAL COST'], displayData['PO TOTAL COST'])
  ) {
    return 'Failed';
  }

  if (
    basicConditionsMet &&
    futureEligibilityAndBids &&
    hasValue_(data['BAC RESOLUTION NO.']) &&
    hasValue_(data['SUPPLIER']) &&
    hasValue_(data['DETAILED BID EVALUATION']) &&
    hasValue_(data['POST-QUALIFICATION']) &&
    hasValue_(data['NOA DATE']) &&
    hasValue_(data['NTP DATE']) &&
    isBlankValue_(data['DATE PREPARED (PO)'], displayData['DATE PREPARED (PO)']) &&
    isBlankValue_(data['PO NO.'], displayData['PO NO.']) &&
    isZero_(data['PO TOTAL COST'], displayData['PO TOTAL COST'])
  ) {
    return 'Awarded';
  }

  if (
    basicConditionsMet &&
    futureEligibilityAndBids &&
    hasValue_(data['BAC RESOLUTION NO.']) &&
    hasValue_(data['SUPPLIER']) &&
    hasValue_(data['DETAILED BID EVALUATION']) &&
    hasValue_(data['POST-QUALIFICATION']) &&
    hasValue_(data['NOA DATE']) &&
    hasValue_(data['NTP DATE']) &&
    hasValue_(data['DATE PREPARED (PO)']) &&
    hasValue_(data['PO NO.']) &&
    hasValue_(data['PO TOTAL COST'])
  ) {
    return 'Purchase Order';
  }

  return '';
}


/* ============================================================
 * AUTOMATIC OPEN TRIGGER
 * ============================================================ */

function validatorOnOpen(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return;
    refreshAllStatuses_(sheet);
  } catch (error) {
    console.error('validatorOnOpen error:', error);
  }
}


/*
 * Recalculate every data row.
 *
 * This is important because a status can change simply because
 * today's date has changed, even when nobody edited the row.
 */
function refreshAllStatuses_(sheet) {
  if (!sheet || sheet.getName() !== CONFIG.SHEET_NAME) return;

  const headers = getHeaders_(sheet);
  const statusColumn = headers.indexOf(CONFIG.STATUS_HEADER) + 1;
  if (statusColumn <= 0) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;

  for (let row = CONFIG.HEADER_ROW + 1; row <= lastRow; row++) {
    updateRowStatus_(sheet, row, headers, statusColumn);
  }
}


/* ============================================================
 * AUTOMATIC EDIT TRIGGER
 * ============================================================ */

function validatorOnEdit(e) {
  // Deprecated. Automatic STATUS calculation is handled exclusively by
  // the simple onEdit(e) trigger. This function intentionally does nothing
  // so an old installable trigger cannot compete with the status engine.
  return;
}
/* ============================================================
 * AUTOMATIC STATUS DETERMINATION
 * ============================================================ */

function determineAutomaticStatus_(
  sheet,
  row,
  headers
) {

  // Read the entire edited row in ONE spreadsheet call for values and
  // ONE call for display values. The previous implementation called
  // getRange().getValue() separately for every column, which made each
  // edit unnecessarily slow.
  const lastColumn = sheet.getLastColumn();
  const rowValues = sheet
    .getRange(row, 1, 1, lastColumn)
    .getValues()[0];
  const rowDisplayValues = sheet
    .getRange(row, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const data = {};
  const displayData = {};

  headers.forEach((header, index) => {
    if (!header) return;
    data[header] = rowValues[index];
    displayData[header] = rowDisplayValues[index];
  });


  // Keep STATUS blank when the entire procurement row has no data.
  const rowHasAnyData = headers
    .filter(header => header && header !== CONFIG.STATUS_HEADER)
    .some(header => hasValue_(data[header]));

  if (!rowHasAnyData) {
    return '';
  }


  /*
   * ----------------------------------------------------------
   * DATE-BASED STATUS TRANSITION
   * ----------------------------------------------------------
   *
   * Active:
   *   ELIGIBILITY SCREENING OR SUBMISSION OF BIDS is today or later.
   *
   * Closed:
   *   ELIGIBILITY SCREENING AND SUBMISSION OF BIDS are both
   *   strictly before today.
   *
   * Failed / Awarded / Purchase Order:
   *   retain their existing downstream-stage rules.
   *
   * The more specific statuses are checked first so that
   * completed procurement stages are not overwritten by Closed.
   */

  const today = normalizeDate_(new Date());

  const postingDate =
    isNAValue_('', data['POSTING DATE'])
      ? null
      : normalizeDate_(data['POSTING DATE']);

  const eligibilityDate =
    isNAValue_('', data['ELIGIBILITY SCREENING'])
      ? null
      : normalizeDate_(data['ELIGIBILITY SCREENING']);

  const submissionDate =
    isNAValue_('', data['SUBMISSION OF BIDS'])
      ? null
      : normalizeDate_(data['SUBMISSION OF BIDS']);

  const postingIsBeforeOrToday =
    postingDate &&
    postingDate.getTime() <= today.getTime();

  const eligibilityIsBeforeToday =
    eligibilityDate &&
    eligibilityDate.getTime() < today.getTime();

  const submissionIsBeforeToday =
    submissionDate &&
    submissionDate.getTime() < today.getTime();

  const eligibilityIsTodayOrAfter =
    eligibilityDate &&
    eligibilityDate.getTime() >= today.getTime();

  const submissionIsTodayOrAfter =
    submissionDate &&
    submissionDate.getTime() >= today.getTime();

  // Active when either eligibility screening OR submission of bids
  // is today or later.
  const activeDateCondition =
    eligibilityIsTodayOrAfter === true ||
    submissionIsTodayOrAfter === true;

  // Closed only when BOTH dates are strictly before today.
  const closedDateCondition =
    eligibilityIsBeforeToday === true &&
    submissionIsBeforeToday === true;

  // Existing downstream statuses (Failed, Awarded, Purchase Order)
  // still require BOTH eligibility screening and submission of bids
  // to be strictly after today.
  const futureEligibilityAndBids =
    eligibilityDate &&
    submissionDate &&
    eligibilityDate.getTime() > today.getTime() &&
    submissionDate.getTime() > today.getTime();

  // IMPORTANT:
  // Active = either eligibility OR submission is TODAY OR LATER.
  // Closed = BOTH eligibility AND submission are STRICTLY BEFORE TODAY.


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
   * BASIC PROCUREMENT FIELDS
   * ----------------------------------------------------------
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
    basicFields.every(field => hasValue_(data[field]));

  /*
   * ----------------------------------------------------------
   * ACTIVE
   * ----------------------------------------------------------
   * Active when EITHER eligibility screening OR submission of
   * bids is today or later.
   */
  if (
    basicConditionsMet &&
    postingIsBeforeOrToday === true &&
    activeDateCondition === true &&
    isZero_(
      data['PO TOTAL COST'],
      displayValue_(sheet, row, headers, 'PO TOTAL COST')
    )
  ) {
    return 'Active';
  }

  /*
   * ----------------------------------------------------------
   * CLOSED
   * ----------------------------------------------------------
   * Closed only when BOTH eligibility screening and submission
   * of bids are strictly before today.
   */
  if (
    basicConditionsMet &&
    closedDateCondition
  ) {
    return 'Closed';
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

      if (isAcceptedValue_(header, value, displayValue)) {
        return;
      }

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

      if (isAcceptedValue_(header, value, displayValue)) {
        return;
      }

      if (isBlankValue_(value, displayValue)) {

        errors.push({
          row: row,
          status: status,
          field: header,
          message:
            'A date is required and it must be before today.'
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

      if (isAcceptedValue_(header, value, displayValue)) {
        return;
      }

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
    dateNumber >= todayNumber
  ) {

    return (
      'Invalid date. The date must be before today.'
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

  if (
    dateValue === null ||
    dateValue === undefined ||
    dateValue === ''
  ) {
    return null;
  }

  /*
   * Google Sheets normally returns real Date objects.
   */
  if (
    Object.prototype.toString.call(dateValue) ===
    '[object Date]'
  ) {

    if (isNaN(dateValue.getTime())) {
      return null;
    }

    return new Date(
      dateValue.getFullYear(),
      dateValue.getMonth(),
      dateValue.getDate()
    );
  }

  /*
   * Support displayed date text such as:
   * September 29, 2026
   * 09/29/2026
   * 9/29/2026
   *
   * The first pattern is preferred because it is unambiguous.
   */
  const text = String(dateValue).trim();

  if (!text) {
    return null;
  }

  let match =
    text.match(
      /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})$/i
    );

  if (match) {

    const monthNames = [
      'january',
      'february',
      'march',
      'april',
      'may',
      'june',
      'july',
      'august',
      'september',
      'october',
      'november',
      'december'
    ];

    const month =
      monthNames.indexOf(
        match[1].toLowerCase()
      );

    const day = Number(match[2]);
    const year = Number(match[3]);

    const parsed =
      new Date(year, month, day);

    if (
      parsed.getFullYear() === year &&
      parsed.getMonth() === month &&
      parsed.getDate() === day
    ) {
      return parsed;
    }

    return null;
  }

  /*
   * Support common numeric date displays.
   * Google Sheets locale usually determines the display,
   * so use the JavaScript parser only as a fallback.
   */
  const parsed = new Date(text);

  if (isNaN(parsed.getTime())) {
    return null;
  }

  return new Date(
    parsed.getFullYear(),
    parsed.getMonth(),
    parsed.getDate()
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
