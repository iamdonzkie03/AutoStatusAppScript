/**
 * PROCUREMENT MONITORING / AUTO STATUS
 * Single-file Apps Script. HTML is embedded below.
 *
 * Automatic behavior:
 *  - onOpen: opens the validation monitor automatically.
 *  - onEdit: validates edited rows and automatically updates STATUS.
 *  - The monitor polls the validation result and displays all errors.
 *
 * Status rules:
 *  Active          = bidding is ongoing / not yet completed
 *  Closed          = Eligibility Screening AND Submission of Bids
 *                    are on or before today, before award/failure
 *  Failed          = post-bidding fields complete, Supplier is blank
 *  Awarded         = post-bidding fields complete, Supplier is present
 *  Purchase Order  = Awarded + all PO fields complete
 *
 * NOTE:
 * A background onEdit trigger should not attempt to create a new
 * browser modal. The monitor is opened by onOpen and remains connected
 * to onEdit through UserProperties.
 */

const CFG = {
  TARGET_SHEET: 'Data List',
  HEADER_ROW: 1,
  STATUS_HEADER: 'STATUS',
  ERROR_PROPERTY: 'PROCUREMENT_VALIDATION_RESULT',
  POLL_INTERVAL_MS: 600
};


/* =========================
 * AUTOMATIC EVENTS
 * =========================
 *
 * IMPORTANT:
 * The actual automation uses INSTALLABLE triggers.
 * This is required because the validation monitor uses
 * authorized Apps Script services and an HTML dialog.
 *
 * Run setupAutomaticTriggers() ONCE from the Apps Script editor.
 * No custom Sheet menu is created. The setup creates only the
 * installable onEdit trigger; onOpen is the native simple trigger.
 */

function setupAutomaticTriggers() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(CFG.TARGET_SHEET);

  if (!targetSheet) {
    throw new Error('Sheet "' + CFG.TARGET_SHEET + '" was not found.');
  }

  /*
   * Remove previous copies of our installable triggers so
   * repeated setup does not create duplicate executions.
   */
  ScriptApp.getProjectTriggers().forEach(function(trigger) {

    const handler = trigger.getHandlerFunction();

    if (
      handler === 'handleOpen_' ||
      handler === 'handleEdit_'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp
    .newTrigger('handleEdit_')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  /*
   * Run the same logic immediately so the current sheet
   * does not need to be closed/reopened after setup.
   */
  handleOpen_();
}


function refreshAllStatuses_() {

  const sheet = getTargetSheet_();

  if (!sheet) return;

  const lastRow = sheet.getLastRow();

  if (lastRow <= CFG.HEADER_ROW) {
    saveValidationResult_(sheet, []);
    return;
  }

  let allErrors = [];

  for (
    let row = CFG.HEADER_ROW + 1;
    row <= lastRow;
    row++
  ) {

    const result =
      processRow_(sheet, row);

    allErrors =
      allErrors.concat(
        result.errors
      );
  }

  allErrors =
    removeDuplicateErrors_(allErrors);

  /*
   * Save the current validation state so the modeless monitor
   * can display errors immediately on opening.
   */
  saveValidationResult_(
    sheet,
    allErrors
  );

  SpreadsheetApp.flush();
}


/*
 * Installable OPEN trigger.
 */
function handleOpen_(e) {

  try {
    refreshAllStatuses_();
  } catch (err) {
    console.error('handleOpen_/refreshAllStatuses_:', err);
  }

}


/*
 * Installable EDIT trigger.
 */
function handleEdit_(e) {

  if (!e || !e.range) return;

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(1500)) return;

  try {

    const range = e.range;
    const sheet = range.getSheet();

    // This automation runs ONLY on the Data List sheet.
    if (sheet.getName() !== CFG.TARGET_SHEET) {
      return;
    }

    /*
     * Ignore the header row.
     */
    if (
      range.getLastRow() <= CFG.HEADER_ROW
    ) {
      return;
    }

    const firstRow = Math.max(
      range.getRow(),
      CFG.HEADER_ROW + 1
    );

    const lastRow =
      range.getRow() +
      range.getNumRows() -
      1;

    let allErrors = [];

    for (
      let row = firstRow;
      row <= lastRow;
      row++
    ) {

      const result =
        processRow_(sheet, row);

      allErrors =
        allErrors.concat(
          result.errors
        );
    }

    allErrors =
      removeDuplicateErrors_(allErrors);

    saveValidationResult_(
      sheet,
      allErrors
    );

  } catch (err) {

    console.error(
      'handleEdit_:',
      err
    );

  } finally {

    lock.releaseLock();
  }
}


/*
 * Kept as lightweight fallbacks for compatibility.
 *
 * The installable onEdit trigger is authoritative for data processing.
 * The native simple onOpen trigger owns the UI because opening an
 * HtmlService window is a spreadsheet UI operation.
 */
function onOpen(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Do not run or open the monitor for any sheet other than Data List.
  if (!sheet || sheet.getName() !== CFG.TARGET_SHEET) {
    return;
  }

  // Native simple onOpen owns the UI action. This is deliberately
  // separate from the authorized installable trigger used for data work.
  try {
    openValidationMonitor_();
  } catch (err) {
    console.error('onOpen/openValidationMonitor_:', err);
  }
}


function onEdit(e) {
  return;
}


function getTargetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CFG.TARGET_SHEET);
}


/* =========================
 * ROW PROCESSING
 * ========================= */

function processRow_(sheet, row) {
  const data = getRowData_(sheet, row);

  // STATUS itself must never count as user input.
  // Otherwise an existing status (for example "Active") makes an
  // otherwise blank row look non-empty and causes the script to assign
  // a status again.
  if (!data || isProcurementInputEmpty_(data)) {
    setStatus_(sheet, row, '');
    return { status: '', errors: [] };
  }

  const errors = [];
  const status = determineStatus_(data, errors);

  // A blank status is always written back as blank.
  setStatus_(sheet, row, status);

  /*
   * Validate the values that were actually entered BEFORE
   * applying status-specific required-field rules.
   *
   * This is important when a row is incomplete and therefore
   * has no status yet. Invalid dates/numbers must still be
   * reported in the modal.
   */
  validateEnteredFields_(data, errors);

  /*
   * If a procurement row has started but does not yet qualify
   * for a status, use Active's required-field rules only for
   * validation. This does NOT write Active into STATUS.
   */
  const validationStatus =
    status || 'Active';

  validateForStatus_(
    data,
    validationStatus,
    errors
  );

  return {
    status: status,
    errors: removeDuplicateErrors_(errors)
  };
}


/* =========================
 * AUTOMATIC STATUS
 * ========================= */

function determineStatus_(data, errors) {

  // Never assign a status to a row that has no procurement input.
  // This check is intentionally independent of the STATUS cell itself.
  if (isProcurementInputEmpty_(data)) {
    return '';
  }

  /*
   * Do not assign Active/Closed/Awarded/Failed/Purchase Order
   * until POSTING DATE has been entered and is valid.
   * Validation still runs separately, so an invalid Posting Date
   * is reported instead of silently producing a status.
   */
  const postingKey = normalizeHeader_('POSTING DATE');

  if (data.columns[postingKey] === undefined) {
    return '';
  }

  const postingIndex = data.columns[postingKey];
  const postingRaw = data.values[postingIndex];
  const postingDisplay = String(
    data.displayValues[postingIndex] || ''
  ).trim();

  if (
    !postingDisplay ||
    !isValidDate_(postingRaw, postingDisplay)
  ) {
    return '';
  }

  const awardFields = [
    'PRE-PROCUREMENT CONFERENCE',
    'PRE-BID CONFERENCE',
    'POSTING DATE',
    'PROJECT ID',
    'ELIGIBILITY SCREENING',
    'SUBMISSION OF BIDS',
    'DETAILED BID EVALUATION',
    'POST-QUALIFICATION',
    'NOA DATE'
  ];

  const poFields = [
    'PRE-PROCUREMENT CONFERENCE',
    'PRE-BID CONFERENCE',
    'POSTING DATE',
    'PROJECT ID',
    'ELIGIBILITY SCREENING',
    'SUBMISSION OF BIDS',
    'DETAILED BID EVALUATION',
    'POST-QUALIFICATION',
    'NOA DATE',
    'NTP DATE',
    'PO DATE',
    'PO NO.',
    'PO TOTAL COST'
  ];

  const basicFields = [
    'POSTING DATE',
    'PRE-PROCUREMENT CONFERENCE',
    'PROJECT ID',
    'PRE-BID CONFERENCE',
    'ELIGIBILITY SCREENING',
    'SUBMISSION OF BIDS'
  ];

  const awardComplete = allFieldsPresent_(data, awardFields);
  const awardValid = allFieldsCorrectlyFormatted_(data, awardFields);

  const supplier = String(getValue_(data, 'SUPPLIER') || '').trim();
  const hasSupplier = supplier !== '';

  const poComplete = allFieldsPresent_(data, poFields);
  const poValid = allFieldsCorrectlyFormatted_(data, poFields);

  /*
   * Purchase Order is the final stage.
   * Supplier must exist because PO is an awarded procurement.
   */
  if (
    awardComplete &&
    awardValid &&
    hasSupplier &&
    poComplete &&
    poValid
  ) {
    return 'Purchase Order';
  }

  /*
   * Awarded = completed post-bidding information + Supplier.
   */
  if (
    awardComplete &&
    awardValid &&
    hasSupplier
  ) {
    return 'Awarded';
  }

  /*
   * Failed = completed post-bidding information + NO Supplier.
   *
   * Supplier is deliberately NOT required here.
   */
  if (
    awardComplete &&
    awardValid &&
    !hasSupplier
  ) {
    return 'Failed';
  }

  /*
   * Before Awarded/Failed, determine Active vs Closed.
   */
  const eligibility = getDate_(data, 'ELIGIBILITY SCREENING');
  const submission = getDate_(data, 'SUBMISSION OF BIDS');

  const today = startOfDay_(new Date());

  /*
   * Once both bidding dates have passed, status is Closed.
   * Awarded/Failed has already been checked above.
   */
  if (
    eligibility &&
    submission &&
    eligibility <= today &&
    submission <= today
  ) {
    return 'Closed';
  }

  return 'Active';
}


/* =========================
 * ENTERED-VALUE VALIDATION
 * ========================= */

/*
 * Validate every date/number that the user has actually entered,
 * independently of STATUS.
 *
 * This guarantees that an invalid date or number is reported even
 * when the row is incomplete and STATUS is still blank.
 */
function validateEnteredFields_(data, errors) {

  const dateFields = [
    'POSTING DATE',
    'PRE-PROCUREMENT CONFERENCE',
    'PRE-BID CONFERENCE',
    'ELIGIBILITY SCREENING',
    'SUBMISSION OF BIDS',
    'POST-QUALIFICATION',
    'DETAILED BID EVALUATION',
    'NOA DATE',
    'NTP DATE',
    'PO DATE'
  ];

  dateFields.forEach(function(field) {

    const key = normalizeHeader_(field);

    if (data.columns[key] === undefined) return;

    const index = data.columns[key];

    const raw = data.values[index];

    const display =
      String(data.displayValues[index] || '').trim();

    if (!display) return;

    if (!isValidDate_(raw, display)) {
      addError_(
        errors,
        data.row,
        field,
        'Please use correct date format.'
      );
    }
  });

  /*
   * Numeric fields are checked whenever they contain a value.
   */
  [
    'TOTAL ABC',
    'PO TOTAL COST'
  ].forEach(function(field) {

    const key = normalizeHeader_(field);

    if (data.columns[key] === undefined) return;

    const index = data.columns[key];

    const raw = data.values[index];

    const display =
      String(data.displayValues[index] || '').trim();

    if (!display) return;

    if (!isValidNumeric_(raw, display)) {
      addError_(
        errors,
        data.row,
        field,
        'Please enter a numeric value with decimals.'
      );
    }
  });
}


/* =========================
 * STATUS-SPECIFIC VALIDATION
 * ========================= */

function validateForStatus_(data, status, errors) {

  if (status === 'Active') {

    validateDateFields_(
      data,
      errors,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS'
      ]
    );

    validateRequiredFields_(
      data,
      errors,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PROJECT ID',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS'
      ]
    );

    validateProjectId_(data, errors);

    validateDateRelativeToToday_(
      data,
      errors,
      true
    );
  }


  else if (status === 'Closed') {

    validateDateFields_(
      data,
      errors,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS'
      ]
    );

    validateRequiredFields_(
      data,
      errors,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PROJECT ID',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS'
      ]
    );

    validateProjectId_(data, errors);
  }


  else if (
    status === 'Awarded' ||
    status === 'Failed'
  ) {

    validateRequiredFields_(
      data,
      errors,
      [
        'PROJECT ID',
        'PRE-PROCUREMENT CONFERENCE',
        'PRE-BID CONFERENCE',
        'POSTING DATE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS',
        'POST-QUALIFICATION',
        'DETAILED BID EVALUATION',
        'NOA DATE'
      ]
    );

    validateDateFields_(
      data,
      errors,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS',
        'POST-QUALIFICATION',
        'DETAILED BID EVALUATION',
        'NOA DATE'
      ]
    );

    validateProjectId_(data, errors);

    /*
     * Supplier:
     *   Awarded -> required
     *   Failed  -> must be blank / not required
     */
    const supplier = String(
      getValue_(data, 'SUPPLIER') || ''
    ).trim();

    if (status === 'Awarded') {

      if (!supplier) {
        addError_(
          errors,
          data.row,
          'SUPPLIER',
          'There is no data inputted.'
        );
      } else {
        validateTextField_(
          data,
          errors,
          'SUPPLIER'
        );
      }

    } else if (status === 'Failed') {

      /*
       * No Supplier error for Failed.
       * Supplier is the field that distinguishes Failed
       * from Awarded.
       */
      if (supplier) {
        addError_(
          errors,
          data.row,
          'SUPPLIER',
          'Supplier should be blank for Failed status.'
        );
      }
    }
  }


  else if (status === 'Purchase Order') {

    validateRequiredFields_(
      data,
      errors,
      [
        'PROJECT ID',
        'PRE-PROCUREMENT CONFERENCE',
        'PRE-BID CONFERENCE',
        'POSTING DATE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS',
        'POST-QUALIFICATION',
        'DETAILED BID EVALUATION',
        'NOA DATE',
        'NTP DATE',
        'PO DATE',
        'PO NO.',
        'PO TOTAL COST',
        'SUPPLIER'
      ]
    );

    validateDateFields_(
      data,
      errors,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS',
        'POST-QUALIFICATION',
        'DETAILED BID EVALUATION',
        'NOA DATE',
        'NTP DATE',
        'PO DATE'
      ]
    );

    validateProjectId_(data, errors);

    validateTextField_(
      data,
      errors,
      'SUPPLIER'
    );

    validateTextField_(
      data,
      errors,
      'PO NO.'
    );

    validateNumericField_(
      data,
      errors,
      'PO TOTAL COST'
    );
  }

  /*
   * TOTAL ABC is checked for every status when present.
   */
  validateNumericField_(
    data,
    errors,
    'TOTAL ABC'
  );
}


/* =========================
 * REQUIRED FIELDS
 * ========================= */

function validateRequiredFields_(data, errors, fields) {

  fields.forEach(function(field) {

    const key = normalizeHeader_(field);

    if (data.columns[key] === undefined) return;

    const value = getValue_(data, field);

    if (String(value || '').trim() === '') {

      addError_(
        errors,
        data.row,
        field,
        'There is no data inputted.'
      );
    }
  });
}


function allFieldsPresent_(data, fields) {

  return fields.every(function(field) {

    const key = normalizeHeader_(field);

    if (data.columns[key] === undefined) {
      return false;
    }

    return String(
      getValue_(data, field) || ''
    ).trim() !== '';
  });
}


/* =========================
 * DATE VALIDATION
 * ========================= */

function validateDateFields_(data, errors, fields) {

  fields.forEach(function(field) {

    const key = normalizeHeader_(field);

    if (data.columns[key] === undefined) return;

    const index = data.columns[key];

    const raw = data.values[index];

    const display = String(
      data.displayValues[index] || ''
    ).trim();

    if (display === '') return;

    if (!isValidDate_(raw, display)) {

      addError_(
        errors,
        data.row,
        field,
        'Please use correct date format.'
      );
    }
  });
}


function allFieldsCorrectlyFormatted_(data, fields) {

  return fields.every(function(field) {

    const key = normalizeHeader_(field);

    if (data.columns[key] === undefined) {
      return false;
    }

    const index = data.columns[key];

    const raw = data.values[index];

    const display = String(
      data.displayValues[index] || ''
    ).trim();

    if (display === '') return false;

    if (isDateField_(field)) {
      return isValidDate_(raw, display);
    }

    if (
      field === 'TOTAL ABC' ||
      field === 'PO TOTAL COST'
    ) {
      return isValidNumeric_(raw, display);
    }

    return true;
  });
}


function isDateField_(field) {

  const fields = [
    'POSTING DATE',
    'PRE-PROCUREMENT CONFERENCE',
    'PRE-BID CONFERENCE',
    'ELIGIBILITY SCREENING',
    'SUBMISSION OF BIDS',
    'POST-QUALIFICATION',
    'DETAILED BID EVALUATION',
    'NOA DATE',
    'NTP DATE',
    'PO DATE'
  ];

  return fields.indexOf(
    normalizeHeader_(field)
  ) !== -1;
}


function isValidDate_(raw, display) {

  if (
    raw instanceof Date &&
    !isNaN(raw.getTime())
  ) {
    return true;
  }

  const normalized = display
    .trim()
    .toLowerCase();

  if (
    normalized === 'n/a' ||
    normalized === 'na' ||
    normalized === 'not applicable'
  ) {
    return true;
  }

  /*
   * Required text date format:
   * January 5, 2026
   */
  const pattern =
    /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$/i;

  if (!pattern.test(display)) {
    return false;
  }

  const parsed = new Date(display);

  return !isNaN(parsed.getTime());
}


/* =========================
 * ACTIVE DATE RULES
 * ========================= */

function validateDateRelativeToToday_(
  data,
  errors,
  reportPastDates
) {

  if (!reportPastDates) return;

  const today = startOfDay_(new Date());

  [
    'ELIGIBILITY SCREENING',
    'SUBMISSION OF BIDS'
  ].forEach(function(field) {

    const date = getDate_(data, field);

    if (!date) return;

    if (date <= today) {

      addError_(
        errors,
        data.row,
        field,
        'Not applicable'
      );
    }
  });
}


function getDate_(data, field) {

  const key = normalizeHeader_(field);

  if (data.columns[key] === undefined) {
    return null;
  }

  const index = data.columns[key];

  const raw = data.values[index];

  const display = String(
    data.displayValues[index] || ''
  ).trim();

  if (!display) return null;

  if (
    raw instanceof Date &&
    !isNaN(raw.getTime())
  ) {

    return startOfDay_(
      new Date(raw)
    );
  }

  if (!isValidDate_(raw, display)) {
    return null;
  }

  return startOfDay_(
    new Date(display)
  );
}


function startOfDay_(date) {

  date.setHours(
    0,
    0,
    0,
    0
  );

  return date;
}


/* =========================
 * PROJECT ID
 * ========================= */

function validateProjectId_(data, errors) {

  const key = 'PROJECT ID';

  if (data.columns[key] === undefined) return;

  const value = String(
    getValue_(data, key) || ''
  ).trim();

  /*
   * PROJECT ID is free-form text.
   * Letters, numbers, spaces, punctuation and symbols are all allowed.
   * The only invalid value for this format check is blank.
   */
  if (!value) {
    return;
  }
}


/* =========================
 * FREE-FORM TEXT
 * =========================
 *
 * Former "alphanumeric" fields now accept:
 *   - letters
 *   - numbers
 *   - spaces
 *   - punctuation
 *   - symbols
 *   - special characters
 *   - any combination of the above
 *
 * The format validator only rejects a blank value when the field
 * is required. It does NOT restrict characters.
 * ========================= */

function validateTextField_(
  data,
  errors,
  field
) {

  const key = normalizeHeader_(field);

  if (data.columns[key] === undefined) return;

  const value = String(
    getValue_(data, field) || ''
  ).trim();

  /*
   * Free-form text:
   * No character whitelist is applied.
   */
  if (!value) {
    return;
  }
}


/* =========================
 * NUMERIC
 * ========================= */

function validateNumericField_(
  data,
  errors,
  field
) {

  const key = normalizeHeader_(field);

  if (data.columns[key] === undefined) return;

  const index = data.columns[key];

  const raw = data.values[index];

  const display = String(
    data.displayValues[index] || ''
  ).trim();

  if (!display) return;

  if (
    typeof raw === 'number' &&
    !isNaN(raw)
  ) {
    return;
  }

  const cleaned = display
    .replace(/,/g, '')
    .replace(/\s/g, '');

  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) {

    addError_(
      errors,
      data.row,
      field,
      'Please enter a numeric value with decimals.'
    );
  }
}


function isValidNumeric_(raw, display) {

  if (
    typeof raw === 'number' &&
    !isNaN(raw)
  ) {
    return true;
  }

  return /^-?\d+(\.\d+)?$/.test(
    display
      .replace(/,/g, '')
      .replace(/\s/g, '')
  );
}


/* =========================
 * SHEET DATA
 * ========================= */

function getRowData_(sheet, row) {

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) return null;

  const headers =
    sheet
      .getRange(
        CFG.HEADER_ROW,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const values =
    sheet
      .getRange(
        row,
        1,
        1,
        lastColumn
      )
      .getValues()[0];

  const displayValues =
    sheet
      .getRange(
        row,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  return {
    sheet: sheet,
    row: row,
    headers: headers,
    values: values,
    displayValues: displayValues,
    columns: buildColumnMap_(headers)
  };
}


function buildColumnMap_(headers) {

  const map = {};

  headers.forEach(function(header, index) {

    const key = normalizeHeader_(header);

    if (key) {
      map[key] = index;
    }
  });

  return map;
}


function getValue_(data, field) {

  const key = normalizeHeader_(field);

  if (data.columns[key] === undefined) {
    return '';
  }

  const index = data.columns[key];

  if (
    data.displayValues[index] !== ''
  ) {
    return data.displayValues[index];
  }

  return data.values[index];
}


function normalizeHeader_(value) {

  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}


const PROCUREMENT_INPUT_FIELDS = [
  'POSTING DATE',
  'PRE-PROCUREMENT CONFERENCE',
  'PROJECT ID',
  'PRE-BID CONFERENCE',
  'ELIGIBILITY SCREENING',
  'SUBMISSION OF BIDS',
  'DETAILED BID EVALUATION',
  'POST-QUALIFICATION',
  'NOA DATE',
  'NTP DATE',
  'PO DATE',
  'PO NO.',
  'PO TOTAL COST',
  'SUPPLIER',
  'TOTAL ABC'
];

function isProcurementInputEmpty_(data) {
  /*
   * Only explicit procurement input fields count.
   * STATUS, formulas, helper columns, and formatting artifacts
   * do not make an otherwise blank row a procurement record.
   */
  return PROCUREMENT_INPUT_FIELDS.every(function(field) {
    const key = normalizeHeader_(field);

    if (data.columns[key] === undefined) {
      return true;
    }

    const index = data.columns[key];

    return String(
      data.displayValues[index] || ''
    ).trim() === '';
  });
}


function isDataRowEmpty_(data) {

  // Ignore the STATUS column completely when deciding whether
  // the row contains actual procurement data.
  return data.headers.every(function(header, index) {

    if (
      normalizeHeader_(header) ===
      CFG.STATUS_HEADER
    ) {
      return true;
    }

    return String(
      data.displayValues[index] || ''
    ).trim() === '';
  });
}


/* =========================
 * STATUS WRITING
 * ========================= */

function setStatus_(sheet, row, status) {

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) return;

  const headers =
    sheet
      .getRange(
        CFG.HEADER_ROW,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  let statusColumn = -1;

  for (let i = 0; i < headers.length; i++) {

    if (
      normalizeHeader_(headers[i]) ===
      CFG.STATUS_HEADER
    ) {

      statusColumn = i + 1;
      break;
    }
  }

  if (statusColumn === -1) {
    return;
  }

  const cell =
    sheet.getRange(row, statusColumn);

  const current =
    String(
      cell.getDisplayValue() || ''
    ).trim();

  if (current !== status) {
    cell.setValue(status);
  }
}


/* =========================
 * ERROR OBJECTS
 * ========================= */

function addError_(errors, row, column, message) {
  errors.push({
    row: row,
    column: column,
    message: message
  });
}


/* =========================
 * ERROR STORAGE
 * ========================= */

function saveValidationResult_(
  sheet,
  errors
) {

  const props =
    PropertiesService
      .getUserProperties();

  if (errors.length === 0) {

    props.deleteProperty(
      CFG.ERROR_PROPERTY
    );

    return;
  }

  props.setProperty(
    CFG.ERROR_PROPERTY,
    JSON.stringify({
      timestamp: Date.now(),
      sheet: sheet.getName(),
      errors: errors
    })
  );
}


function runValidationNow() {
  /*
   * The HTML monitor actively refreshes validation instead of relying
   * exclusively on the onEdit trigger. This makes the modeless monitor resilient
   * to delayed/missed trigger executions and catches changes immediately.
   */
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(1500)) {
    return getValidationResult();
  }

  try {
    refreshAllStatuses_();
    return getValidationResult();
  } catch (err) {
    console.error('runValidationNow:', err);
    return getValidationResult();
  } finally {
    lock.releaseLock();
  }
}


function getValidationResult() {

  const value =
    PropertiesService
      .getUserProperties()
      .getProperty(
        CFG.ERROR_PROPERTY
      );

  if (!value) return null;

  try {
    return JSON.parse(value);
  } catch (err) {
    return null;
  }
}


function clearValidationResult() {

  PropertiesService
    .getUserProperties()
    .deleteProperty(
      CFG.ERROR_PROPERTY
    );

  return true;
}


function removeDuplicateErrors_(errors) {

  const seen = {};
  const result = [];

  errors.forEach(function(error) {

    const key =
      error.row +
      '|' +
      error.column +
      '|' +
      error.message;

    if (!seen[key]) {

      seen[key] = true;
      result.push(error);
    }
  });

  return result;
}


/* =========================
 * HTML MONITOR
 * ========================= */

function openValidationMonitor_() {

  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (!activeSheet || activeSheet.getName() !== CFG.TARGET_SHEET) {
    return;
  }

  const html =
    HtmlService
      .createHtmlOutput(
        getValidationMonitorHtml_()
      )
      .setWidth(720)
      .setHeight(540);

  SpreadsheetApp
    .getUi()
    .showModelessDialog(
      html,
      'Procurement Data Validation'
    );
}


function getValidationMonitorHtml_() {

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">

<style>

* {
  box-sizing: border-box;
}

html,
body {
  margin: 0;
  padding: 0;
  width: 100%;
  height: 100%;
  font-family: Arial, sans-serif;
  background: #f5f6f8;
  color: #202124;
}

.container {
  padding: 22px;
}

.header {
  display: flex;
  align-items: center;
  gap: 13px;
  margin-bottom: 16px;
}

.icon {
  width: 45px;
  height: 45px;
  border-radius: 50%;
  background: #f4b400;
  color: white;
  font-size: 26px;
  font-weight: bold;
  display: flex;
  align-items: center;
  justify-content: center;
}

.title {
  font-size: 20px;
  font-weight: bold;
}

.subtitle {
  font-size: 12px;
  color: #666;
  margin-top: 3px;
}

.status {
  background: white;
  border: 1px solid #ddd;
  border-radius: 7px;
  padding: 11px 14px;
  margin-bottom: 12px;
}

.status.good {
  color: #188038;
}

.status.bad {
  color: #b31412;
}

.info {
  background: white;
  border: 1px solid #ddd;
  border-radius: 8px;
  padding: 11px 14px;
  margin-bottom: 12px;
}

.error-list {
  background: white;
  border: 1px solid #ddd;
  border-radius: 8px;
  max-height: 315px;
  overflow-y: auto;
}

.error {
  padding: 13px 15px;
  border-bottom: 1px solid #eee;
}

.error:last-child {
  border-bottom: none;
}

.location {
  font-weight: bold;
  margin-bottom: 5px;
}

.column {
  color: #b31412;
}

.message {
  color: #555;
  font-size: 13px;
  line-height: 1.4;
}

.empty {
  background: white;
  border: 1px solid #ddd;
  border-radius: 8px;
  text-align: center;
  padding: 50px 20px;
  color: #777;
}

.footer {
  display: flex;
  justify-content: flex-end;
  margin-top: 15px;
}

button {
  border: none;
  border-radius: 6px;
  padding: 10px 20px;
  cursor: pointer;
  font-weight: bold;
  font-size: 14px;
  background: #1a73e8;
  color: white;
}

</style>
</head>

<body>

<div class="container">

  <div class="header">

    <div class="icon">!</div>

    <div>
      <div class="title">
        Procurement Data Validation
      </div>

      <div class="subtitle">
        Automatic validation monitor — keep editing the sheet — keep editing the sheet
      </div>
    </div>

  </div>

  <div id="status"
       class="status good">
    Monitoring spreadsheet changes...
  </div>

  <div id="content">
    <div class="empty">
      No validation errors.
    </div>
  </div>

  <div class="footer">
    <button onclick="google.script.host.close()">
      Close
    </button>
  </div>

</div>

<script>

let lastTimestamp = 0;
let validationBusy = false;

function checkValidation() {

  if (validationBusy) return;
  validationBusy = true;

  google.script.run

    .withSuccessHandler(function(result) {

      validationBusy = false;

      if (!result) {
        showNoErrors();
        return;
      }

      lastTimestamp =
        result.timestamp || Date.now();

      showErrors(result);

    })

    .withFailureHandler(function(error) {

      validationBusy = false;

      const status =
        document.getElementById('status');

      status.className =
        'status bad';

      status.textContent =
        'Validation monitor error.';

      console.error(error);

    })

    .runValidationNow();
}


function showErrors(result) {

  const status =
    document.getElementById('status');

  status.className =
    'status bad';

  status.innerHTML =
    '<strong>' +
    result.errors.length +
    '</strong> validation error(s) found.';

  let html =

    '<div class="info">' +
      '<strong>Sheet:</strong> ' +
      escapeHtml(result.sheet) +
    '</div>' +

    '<div class="error-list">';

  result.errors.forEach(function(error) {

    html +=

      '<div class="error">' +

        '<div class="location">' +
          'Row ' +
          escapeHtml(error.row) +
          ' — ' +
          '<span class="column">' +
            escapeHtml(error.column) +
          '</span>' +
        '</div>' +

        '<div class="message">' +
          escapeHtml(error.message) +
        '</div>' +

      '</div>';

  });

  html += '</div>';

  document.getElementById('content')
    .innerHTML = html;
}


function showNoErrors() {

  const status =
    document.getElementById('status');

  status.className =
    'status good';

  status.innerHTML =
    '<strong>✓</strong> No validation errors.';

  document.getElementById('content')
    .innerHTML =
      '<div class="empty">' +
        'All current entries meet the ' +
        'validation rules.' +
      '</div>';
}


function escapeHtml(value) {

  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}


checkValidation();

setInterval(
  checkValidation,
  600
);

</script>

</body>
</html>
`;
}
