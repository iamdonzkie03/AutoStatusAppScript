/**
 * PROCUREMENT AUTO STATUS / VALIDATOR
 * Target sheet: Data List
 *
 * Run setupValidator() ONCE after installing this code.
 *
 * Status rules:
 * - ACTIVE: core procurement fields complete; eligibility/submission are
 *   today or later.
 * - CLOSED: core procurement fields complete; eligibility/submission are
 *   both before today.
 * - AWARDED: downstream award fields complete; Supplier required.
 * - FAILED: downstream award fields complete; Supplier NOT required.
 * - PURCHASE ORDER: all PO fields complete and PO TOTAL COST > 0.
 * - CANCELLED PO: all PO fields complete and PO TOTAL COST = 0/0.00.
 * - CANCELLED PR: PR TOTAL ABC = 0 and Remarks = Cancelled PR.
 * - REALIGNED ITEM: TOTAL ABC = 0/0.00 AND Remarks contains a word beginning with "realign".
 *
 * Blank fields are not errors unless the selected status requires them.
 * No custom spreadsheet menu is created.
 */

const CONFIG = {
  SHEET_NAME: 'Data List',
  HEADER_ROW: 1,
  STATUS_HEADER: 'STATUS',
  DATE_FORMAT: 'MMMM d, yyyy',
  MAX_ERRORS_DISPLAYED: 100
};

const PHILGEPS_REQUIRED_FIELD = 'PHILGEPS REFERENCE NO.';

const CORE_FIELDS = [
  'PRE-PROCUREMENT CONFERENCE',
  'POSTING DATE',
  'PROJECT ID',
  'PRE-BID CONFERENCE',
  'ELIGIBILITY SCREENING',
  'SUBMISSION OF BIDS',
  'PROCUREMENT METHOD',
  'PROJECT TITLE',
  PHILGEPS_REQUIRED_FIELD
];

const AWARDED_FIELDS = CORE_FIELDS.concat([
  'DETAILED BID EVALUATION',
  'POST-QUALIFICATION',
  'BAC RESOLUTION NO.',
  'NOA DATE',
  'SUPPLIER'
]);

const FAILED_FIELDS = [
  PHILGEPS_REQUIRED_FIELD,
  'PRE-PROCUREMENT CONFERENCE',
  'POSTING DATE',
  'PRE-BID CONFERENCE',
  'ELIGIBILITY SCREENING',
  'SUBMISSION OF BIDS',
  'DETAILED BID EVALUATION',
  'POST-QUALIFICATION',
  'NOA DATE',
  'BAC RESOLUTION NO.'
];

const PO_FIELDS = AWARDED_FIELDS.concat([
  'NTP DATE',
  'PO NO.',
  'DATE PREPARED (PO)',
  'PO TOTAL COST'
]);

const DATE_FIELDS = [
  'PRE-PROCUREMENT CONFERENCE',
  'POSTING DATE',
  'PRE-BID CONFERENCE',
  'ELIGIBILITY SCREENING',
  'SUBMISSION OF BIDS',
  'DETAILED BID EVALUATION',
  'POST-QUALIFICATION',
  'NOA DATE',
  'NTP DATE',
  'DATE PREPARED (PO)'
];

const NUMERIC_FIELDS = [
  'TOTAL ABC',
  'PR TOTAL ABC',
  'PHILGEPS REFERENCE NO.',
  'PO TOTAL COST'
];

/**
 * Install/refresh the validator triggers.
 * Run manually once.
 */
function setupValidator() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet found.');

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + CONFIG.SHEET_NAME + '" was not found.');
  }

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const fn = trigger.getHandlerFunction();
    if (fn === 'validatorOnEdit' || fn === 'validatorOnOpen') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('validatorOnEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger('validatorOnOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  refreshAllStatuses_(sheet);
  SpreadsheetApp.flush();

  ss.toast('Procurement validator installed.', 'Validator', 4);
}

/**
 * Simple trigger fallback for automatic status changes.
 * No UI/modal code is used here.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;
    if (e.range.getRow() <= CONFIG.HEADER_ROW) return;

    processEditedRows_(sheet, e.range);
  } catch (err) {
    console.error('onEdit error: ' + err);
  }
}

/**
 * Simple onOpen intentionally does not create a menu.
 * The installable validatorOnOpen trigger performs refresh.
 */
function onOpen(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);
    if (sheet) refreshAllStatuses_(sheet);
  } catch (err) {
    console.error('onOpen error: ' + err);
  }
}

/**
 * Installable onOpen trigger.
 */
function validatorOnOpen(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);
    if (sheet) refreshAllStatuses_(sheet);
  } catch (err) {
    console.error('validatorOnOpen error: ' + err);
  }
}

/**
 * Installable onEdit trigger.
 * Status is calculated first; validation is then performed only
 * against the status-specific requirements.
 */
function validatorOnEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.SHEET_NAME) return;
    if (e.range.getRow() <= CONFIG.HEADER_ROW) return;

    const errors = processEditedRows_(sheet, e.range, true);
    if (errors.length) showValidationModal_(errors);
  } catch (err) {
    console.error('validatorOnEdit error: ' + err);
  }
}

/**
 * Process all rows touched by an edit.
 */
function processEditedRows_(sheet, range, collectErrors) {
  const headers = getHeaders_(sheet);
  const statusColumn = headers.indexOf(CONFIG.STATUS_HEADER) + 1;
  if (statusColumn <= 0) return [];

  const errors = [];
  const firstRow = Math.max(range.getRow(), CONFIG.HEADER_ROW + 1);
  const lastRow = firstRow + range.getNumRows() - 1;

  for (let row = firstRow; row <= lastRow; row++) {
    const status = determineStatus_(sheet, row, headers);
    const cell = sheet.getRange(row, statusColumn);
    const current = normalize_(cell.getDisplayValue());

    if (current !== status) {
      if (status) cell.setValue(status);
      else cell.clearContent();
    }

    if (collectErrors && status) {
      const rowErrors = validateRow_(sheet, row, headers, status);
      Array.prototype.push.apply(errors, rowErrors);
    }
  }

  SpreadsheetApp.flush();
  return errors;
}

/**
 * Recalculate all procurement rows.
 */
function refreshAllStatuses_(sheet) {
  if (!sheet || sheet.getName() !== CONFIG.SHEET_NAME) return;

  const headers = getHeaders_(sheet);
  const statusColumn = headers.indexOf(CONFIG.STATUS_HEADER) + 1;
  if (statusColumn <= 0) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;

  for (let row = CONFIG.HEADER_ROW + 1; row <= lastRow; row++) {
    const status = determineStatus_(sheet, row, headers);
    const cell = sheet.getRange(row, statusColumn);
    const current = normalize_(cell.getDisplayValue());

    if (current !== status) {
      if (status) cell.setValue(status);
      else cell.clearContent();
    }
  }
}

/**
 * Determine exactly one automatic status.
 *
 * Priority:
 * 1. REALIGNED ITEM
 * 2. CANCELLED PR
 * 3. PURCHASE ORDER / CANCELLED PO
 * 4. AWARDED
 * 5. FAILED
 * 6. CLOSED
 * 7. ACTIVE
 * 8. blank
 */
function determineStatus_(sheet, row, headers) {
  const data = getRowData_(sheet, row, headers);

  if (isCompletelyBlank_(data)) return '';

  // REALIGNED ITEM is determined first from its own explicit conditions.
  // It must not be blocked by the general PhilGEPS gate.
  // Conditions:
  //   1) TOTAL ABC = 0 or 0.00
  //   2) REMARKS contains Realign / Realigned / Realignment, etc.
  if (isRealignedItem_(data, sheet, row, headers)) {
    return 'Realigned Item';
  }

  // PhilGEPS Reference No. is mandatory and numeric for all other statuses.
  // No other automatic status is assigned until it is present and numeric.
  if (!requiredPhilGEPSValid_(data)) return '';

  // Cancelled PR.
  if (
    isZero_(data['PR TOTAL ABC'], display_(sheet, row, headers, 'PR TOTAL ABC')) &&
    /cancel+ed\s+pr/i.test(normalize_(data['REMARKS']))
  ) {
    return 'Cancelled PR';
  }

  // Purchase-order statuses require ALL PO fields first.
  if (fieldsComplete_(data, PO_FIELDS)) {
    const poCost = getNumber_(data['PO TOTAL COST']);
    if (poCost !== null) {
      if (poCost > 0) return 'Purchase Order';
      if (poCost === 0) return 'Cancelled PO';
    }
  }

  // Awarded requires Supplier.
  if (fieldsComplete_(data, AWARDED_FIELDS)) {
    if (downstreamDatesBeforeToday_(data)) return 'Awarded';
  }

  // Failed deliberately does NOT require Supplier or Project ID/Method/Title.
  if (fieldsComplete_(data, FAILED_FIELDS)) {
    if (downstreamDatesBeforeToday_(data)) return 'Failed';
  }

  // Closed: core fields complete and eligibility/submission both before today.
  if (fieldsComplete_(data, CORE_FIELDS)) {
    const eligibility = getDate_(data['ELIGIBILITY SCREENING']);
    const submission = getDate_(data['SUBMISSION OF BIDS']);

    if (
      eligibility &&
      submission &&
      isBeforeToday_(eligibility) &&
      isBeforeToday_(submission)
    ) {
      return 'Closed';
    }

    // Active: core fields complete and either date is today or later.
    if (
      (eligibility && !isBeforeToday_(eligibility)) ||
      (submission && !isBeforeToday_(submission))
    ) {
      return 'Active';
    }
  }

  return '';
}

function isRealignedItem_(data, sheet, row, headers) {
  const totalAbcDisplay = display_(sheet, row, headers, 'TOTAL ABC');
  const remarksDisplay = display_(sheet, row, headers, 'REMARKS');

  const totalAbc = getNumber_(data['TOTAL ABC']);
  const totalAbcFromDisplay = getNumber_(totalAbcDisplay);
  const zeroTotalAbc =
    (totalAbc !== null && totalAbc === 0) ||
    (totalAbcFromDisplay !== null && totalAbcFromDisplay === 0);

  const remarks = normalize_(data['REMARKS'] || remarksDisplay);

  // Match a standalone word beginning with "realign":
  // Realign, Realigned, Realignment, Realignments, etc.
  const hasRealignmentRemark = /\\brealign\\w*\\b/i.test(remarks);

  return zeroTotalAbc && hasRealignmentRemark;
}

function requiredPhilGEPSValid_(data) {
  return hasValue_(data[PHILGEPS_REQUIRED_FIELD]) &&
    isNumeric_(data[PHILGEPS_REQUIRED_FIELD], normalize_(data[PHILGEPS_REQUIRED_FIELD]));
}

function downstreamDatesBeforeToday_(data) {
  const eligibility = getDate_(data['ELIGIBILITY SCREENING']);
  const submission = getDate_(data['SUBMISSION OF BIDS']);

  return !!(
    eligibility &&
    submission &&
    isBeforeToday_(eligibility) &&
    isBeforeToday_(submission)
  );
}

/**
 * Status-specific validation.
 * Blank fields that are not required by the selected status are ignored.
 */
function validateRow_(sheet, row, headers, status) {
  const errors = [];
  const data = getRowData_(sheet, row, headers);

  const required = getRequiredFields_(status);

  required.forEach(function(field) {
    if (!Object.prototype.hasOwnProperty.call(data, field)) return;

    const value = data[field];
    const display = display_(sheet, row, headers, field);

    if (isBlank_(value, display)) {
      errors.push({
        row: row,
        status: status,
        field: field,
        message: 'This field is required for status "' + status + '".'
      });
      return;
    }

    const formatError = validateFieldFormat_(field, value, display);
    if (formatError) {
      errors.push({
        row: row,
        status: status,
        field: field,
        message: formatError
      });
    }
  });

  // Validate any entered non-required fields too, but never treat blanks as errors.
  headers.forEach(function(field) {
    if (!field || field === CONFIG.STATUS_HEADER) return;
    const value = data[field];
    const display = display_(sheet, row, headers, field);
    if (isBlank_(value, display)) return;

    const formatError = validateFieldFormat_(field, value, display);
    if (formatError) {
      const already = errors.some(function(e) {
        return e.field === field;
      });
      if (!already) {
        errors.push({
          row: row,
          status: status,
          field: field,
          message: formatError
        });
      }
    }
  });

  return errors;
}

function getRequiredFields_(status) {
  switch (status) {
    case 'Active':
    case 'Closed':
      return CORE_FIELDS;

    case 'Awarded':
      return AWARDED_FIELDS;

    case 'Failed':
      return FAILED_FIELDS;

    case 'Purchase Order':
    case 'Cancelled PO':
      return PO_FIELDS;

    case 'Cancelled PR':
      return [PHILGEPS_REQUIRED_FIELD, 'PR NO.', 'PR TOTAL ABC', 'REMARKS'];

    case 'Realigned Item':
      return [PHILGEPS_REQUIRED_FIELD, 'TOTAL ABC', 'PR NO.', 'PR TOTAL ABC'];

    default:
      return [];
  }
}

function validateFieldFormat_(field, value, display) {
  if (isBlank_(value, display)) return null;

  if (DATE_FIELDS.indexOf(field) >= 0) {
    if (!isValidDate_(value, display)) {
      return 'Invalid date format. Expected format: ' +
        CONFIG.DATE_FORMAT + '.';
    }
    return null;
  }

  if (NUMERIC_FIELDS.indexOf(field) >= 0) {
    if (!isNumeric_(value, display)) {
      return 'Invalid number format. This field must contain a numeric value.';
    }
    return null;
  }

  if (field === 'PO NO.') {
    if (!/^[A-Za-z0-9\-\/_.]+$/.test(display)) {
      return 'Invalid PO number. Use letters, numbers, hyphens, slashes or underscores only.';
    }
  }

  return null;
}

function fieldsComplete_(data, fields) {
  return fields.every(function(field) {
    return hasValue_(data[field]);
  });
}

function getHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return [];

  return sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(normalizeHeader_);
}

function normalizeHeader_(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/\\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function getRowData_(sheet, row, headers) {
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const data = {};

  headers.forEach(function(header, i) {
    if (header) data[header] = values[i];
  });

  return data;
}

function display_(sheet, row, headers, field) {
  const index = headers.indexOf(field);
  if (index < 0) return '';

  return sheet.getRange(row, index + 1).getDisplayValue();
}

function hasValue_(value) {
  if (value === null || value === undefined) return false;
  if (typeof value === 'string') return value.trim() !== '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return !isNaN(value.getTime());
  }
  return true;
}

function isBlank_(value, display) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string' && value.trim() === '') return true;
  if (display !== undefined && String(display).trim() === '') return true;
  return false;
}

function isCompletelyBlank_(data) {
  return Object.keys(data).every(function(field) {
    return field === CONFIG.STATUS_HEADER || !hasValue_(data[field]);
  });
}

function normalize_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function getDate_(value) {
  if (
    Object.prototype.toString.call(value) === '[object Date]' &&
    !isNaN(value.getTime())
  ) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const text = normalize_(value);
  if (!text || /^n\/?a$/i.test(text)) return null;

  const m = text.match(
    /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})$/i
  );

  if (m) {
    const months = [
      'january','february','march','april','may','june',
      'july','august','september','october','november','december'
    ];
    const month = months.indexOf(m[1].toLowerCase());
    const day = Number(m[2]);
    const year = Number(m[3]);
    const date = new Date(year, month, day);

    if (
      date.getFullYear() === year &&
      date.getMonth() === month &&
      date.getDate() === day
    ) return date;

    return null;
  }

  return null;
}

function isValidDate_(value, display) {
  if (
    Object.prototype.toString.call(value) === '[object Date]' &&
    !isNaN(value.getTime())
  ) {
    return true;
  }

  if (/^n\/?a$/i.test(normalize_(display))) return true;

  return !!getDate_(display);
}

function normalizeDate_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function isBeforeToday_(date) {
  const d = normalizeDate_(date);
  const today = normalizeDate_(new Date());
  return d.getTime() < today.getTime();
}

function getNumber_(value) {
  if (typeof value === 'number' && !isNaN(value)) return value;

  const text = normalize_(value).replace(/,/g, '').replace(/\s/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(text)) return null;

  return Number(text);
}

function isNumeric_(value, display) {
  return getNumber_(value) !== null ||
    getNumber_(display) !== null;
}

function isZero_(value, display) {
  const n = getNumber_(value);
  if (n !== null) return n === 0;

  const d = getNumber_(display);
  return d !== null && d === 0;
}

function showValidationModal_(errors) {
  const shown = errors.slice(0, CONFIG.MAX_ERRORS_DISPLAYED);

  let html =
    '<!doctype html><html><head><base target="_top">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;margin:0;background:#f5f7fa;color:#202124}' +
    '.wrap{padding:22px}.head{background:#b3261e;color:#fff;margin:-22px -22px 18px;padding:18px 22px}' +
    '.head h2{margin:0 0 6px;font-size:20px}.head p{margin:0;font-size:13px}' +
    '.summary{background:#fff;border-left:5px solid #b3261e;padding:12px 15px;margin-bottom:14px}' +
    '.scroll{max-height:390px;overflow:auto}.error{background:#fff;border:1px solid #dadce0;border-radius:5px;padding:12px 14px;margin-bottom:9px}' +
    '.meta{font-size:12px;color:#5f6368;margin-bottom:5px}.field{font-weight:bold;color:#b3261e;margin-bottom:4px}' +
    '.msg{font-size:13px;line-height:1.4}.footer{text-align:right;margin-top:16px}' +
    'button{border:0;border-radius:5px;padding:9px 20px;background:#1a73e8;color:#fff;cursor:pointer}' +
    '</style></head><body><div class="wrap">' +
    '<div class="head"><h2>⚠ Procurement Data Validation Error</h2>' +
    '<p>Please correct the following field(s).</p></div>' +
    '<div class="summary"><strong>' + errors.length +
    ' validation error' + (errors.length === 1 ? '' : 's') + ' found.</strong></div>' +
    '<div class="scroll">';

  shown.forEach(function(error) {
    html += '<div class="error">' +
      '<div class="meta">Row ' + escapeHtml_(error.row) +
      ' &nbsp;|&nbsp; Status: ' + escapeHtml_(error.status) + '</div>' +
      '<div class="field">' + escapeHtml_(error.field) + '</div>' +
      '<div class="msg">' + escapeHtml_(error.message) + '</div>' +
      '</div>';
  });

  if (errors.length > shown.length) {
    html += '<div class="error"><div class="msg">Only the first ' +
      shown.length + ' errors are displayed.</div></div>';
  }

  html += '</div><div class="footer">' +
    '<button onclick="google.script.host.close()">Close</button>' +
    '</div></div></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html)
      .setWidth(650)
      .setHeight(560),
    'Data Validation'
  );
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Optional manual test: validates the currently selected row.
 */
function validateActiveRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet || sheet.getName() !== CONFIG.SHEET_NAME) {
    SpreadsheetApp.getUi().alert(
      'This validator only runs on the "Data List" sheet.'
    );
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= CONFIG.HEADER_ROW) return;

  const headers = getHeaders_(sheet);
  const statusColumn = headers.indexOf(CONFIG.STATUS_HEADER) + 1;
  if (statusColumn <= 0) return;

  const status = normalize_(sheet.getRange(row, statusColumn).getDisplayValue());
  if (!status) return;

  const errors = validateRow_(sheet, row, headers, status);
  if (errors.length) {
    showValidationModal_(errors);
  } else {
    ss.toast('No validation errors found.', 'Validation', 3);
  }
}
