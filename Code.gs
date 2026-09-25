/************************************************************
 * PROCUREMENT MONITORING VALIDATION SYSTEM
 *
 * SINGLE FILE VERSION
 * HTML IS EMBEDDED INSIDE THIS CODE.
 *
 * IMPORTANT:
 * The edit trigger validates the row and stores the errors.
 * A persistent validation window checks for new errors and
 * displays the centered modal.
 ************************************************************/


/* ==========================================================
   CONFIGURATION
   ========================================================== */

const CONFIG = {
  HEADER_ROW: 1,

  // How frequently the HTML client checks for new errors.
  POLL_MS: 700,

  // Properties key used to communicate between Apps Script
  // and the HTML interface.
  ERROR_PROPERTY: 'PROCUREMENT_VALIDATION_ERRORS'
};


/* ==========================================================
   SPREADSHEET OPEN
   ========================================================== */

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('Procurement Validation')
    .addItem('Open Validation Monitor', 'openValidationMonitor')
    .addItem('Test Validation Modal', 'testValidationModal')
    .addToUi();
}


/* ==========================================================
   ON EDIT
   ========================================================== */

/**
 * IMPORTANT:
 *
 * This is the actual onEdit function.
 *
 * Do not rename this to handleEdit.
 *
 * This function only validates the edited row and stores
 * the result. It does NOT attempt to directly open a modal.
 */
function onEdit(e) {

  if (!e || !e.range) {
    return;
  }

  try {

    const range = e.range;
    const sheet = range.getSheet();

    const firstRow = range.getRow();
    const lastRow =
      firstRow + range.getNumRows() - 1;

    // Ignore header.
    if (lastRow <= CONFIG.HEADER_ROW) {
      return;
    }

    let allErrors = [];

    for (
      let row = Math.max(
        firstRow,
        CONFIG.HEADER_ROW + 1
      );
      row <= lastRow;
      row++
    ) {

      const errors =
        validateRow_(sheet, row);

      if (errors.length > 0) {
        allErrors = allErrors.concat(errors);
      }
    }


    /*
     * Remove duplicate errors.
     */
    allErrors =
      removeDuplicateErrors_(allErrors);


    /*
     * Store errors for the HTML monitor.
     */
    const props =
      PropertiesService.getUserProperties();


    if (allErrors.length > 0) {

      props.setProperty(
        CONFIG.ERROR_PROPERTY,
        JSON.stringify({
          timestamp: Date.now(),
          sheet: sheet.getName(),
          errors: allErrors
        })
      );

    } else {

      /*
       * Clear previous errors when the edited data
       * is corrected.
       */
      props.deleteProperty(
        CONFIG.ERROR_PROPERTY
      );
    }

  } catch (error) {

    console.error(
      'onEdit validation error:',
      error
    );
  }
}


/* ==========================================================
   VALIDATE ROW
   ========================================================== */

function validateRow_(sheet, rowNumber) {

  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {
    return [];
  }


  const headers =
    sheet
      .getRange(
        CONFIG.HEADER_ROW,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  const values =
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        lastColumn
      )
      .getValues()[0];


  const displayValues =
    sheet
      .getRange(
        rowNumber,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  const columns =
    buildColumnMap_(headers);


  const errors = [];


  const status =
    getValue_(
      values,
      displayValues,
      columns,
      'STATUS'
    );


  if (!status) {
    return errors;
  }


  const normalizedStatus =
    normalize_(status);


  /* ========================================================
     COMMON NUMERIC VALIDATION
     ======================================================== */

  validateNumericField_(
    errors,
    rowNumber,
    columns,
    values,
    displayValues,
    'TOTAL ABC'
  );


  /* ========================================================
     ACTIVE
     ======================================================== */

  if (normalizedStatus === 'active') {

    validateDateFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS'
      ]
    );


    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PROJECT ID',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS'
      ]
    );


    validateDateRelativeToToday_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues
    );
  }


  /* ========================================================
     CLOSED
     ======================================================== */

  else if (normalizedStatus === 'closed') {

    validateDateFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS'
      ]
    );


    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      [
        'POSTING DATE',
        'PRE-PROCUREMENT CONFERENCE',
        'PROJECT ID',
        'PRE-BID CONFERENCE',
        'ELIGIBILITY SCREENING',
        'SUBMISSION OF BIDS'
      ]
    );
  }


  /* ========================================================
     AWARDED
     ======================================================== */

  else if (normalizedStatus === 'awarded') {

    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
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
        'SUPPLIER'
      ]
    );


    validateDateFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
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


    validateAlphanumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      'PROJECT ID'
    );


    validateAlphanumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      'SUPPLIER'
    );


    validateDateRelativeToToday_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues
    );
  }


  /* ========================================================
     FAILED
     ======================================================== */

  else if (normalizedStatus === 'failed') {

    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
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
        'SUPPLIER'
      ]
    );


    validateDateFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
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


    validateAlphanumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      'PROJECT ID'
    );


    validateAlphanumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      'SUPPLIER'
    );
  }


  /* ========================================================
     PURCHASE ORDER
     ======================================================== */

  else if (
    normalizedStatus === 'purchase order' ||
    normalizedStatus === 'purchaseorder'
  ) {

    validateRequiredFields_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
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
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
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


    validateAlphanumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      'PROJECT ID'
    );


    validateAlphanumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      'SUPPLIER'
    );


    validateAlphanumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      'PO NO.'
    );


    validateNumericField_(
      errors,
      rowNumber,
      columns,
      values,
      displayValues,
      'PO TOTAL COST'
    );
  }


  return errors;
}


/* ==========================================================
   REQUIRED FIELD VALIDATION
   ========================================================== */

function validateRequiredFields_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues,
  fields
) {

  fields.forEach(function(field) {

    const key =
      normalizeHeader_(field);

    if (columns[key] === undefined) {
      return;
    }


    const value =
      getValue_(
        values,
        displayValues,
        columns,
        field
      );


    if (
      String(value || '').trim() === ''
    ) {

      addError_(
        errors,
        rowNumber,
        field,
        'There is no data inputted.'
      );
    }
  });
}


/* ==========================================================
   DATE VALIDATION
   ========================================================== */

function validateDateFields_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues,
  fields
) {

  fields.forEach(function(field) {

    const key =
      normalizeHeader_(field);


    if (columns[key] === undefined) {
      return;
    }


    const index =
      columns[key];


    const raw =
      values[index];


    const display =
      String(
        displayValues[index] || ''
      ).trim();


    /*
     * Blank values are handled by required validation.
     */
    if (display === '') {
      return;
    }


    if (
      !isValidDate_(raw, display)
    ) {

      addError_(
        errors,
        rowNumber,
        field,
        'Please use correct date format.'
      );
    }
  });
}


/**
 * Date formats allowed by the Excel file:
 *
 * MMMM d, YYYY
 * N/A
 * NA
 * Not Applicable
 */
function isValidDate_(raw, display) {

  if (
    Object.prototype.toString.call(raw)
      === '[object Date]' &&
    !isNaN(raw.getTime())
  ) {
    return true;
  }


  const normalized =
    display
      .trim()
      .toLowerCase();


  if (
    normalized === 'n/a' ||
    normalized === 'na' ||
    normalized === 'not applicable'
  ) {
    return true;
  }


  const pattern =
    /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$/i;


  if (!pattern.test(display)) {
    return false;
  }


  const date =
    new Date(display);


  return !isNaN(date.getTime());
}


/* ==========================================================
   DATE RELATIVE TO TODAY
   ========================================================== */

function validateDateRelativeToToday_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues
) {

  const fields = [
    'ELIGIBILITY SCREENING',
    'SUBMISSION OF BIDS'
  ];


  const today =
    new Date();


  today.setHours(
    0,
    0,
    0,
    0
  );


  fields.forEach(function(field) {

    const key =
      normalizeHeader_(field);


    if (columns[key] === undefined) {
      return;
    }


    const index =
      columns[key];


    const raw =
      values[index];


    const display =
      String(
        displayValues[index] || ''
      ).trim();


    if (!display) {
      return;
    }


    const normalized =
      display.toLowerCase();


    if (
      normalized === 'n/a' ||
      normalized === 'na' ||
      normalized === 'not applicable'
    ) {
      return;
    }


    if (!isValidDate_(raw, display)) {
      return;
    }


    const date =
      raw instanceof Date
        ? new Date(raw)
        : new Date(display);


    date.setHours(
      0,
      0,
      0,
      0
    );


    /*
     * According to the Excel conditions:
     *
     * On or before today:
     * "Not applicable"
     *
     * After today:
     * also considered not applicable depending on
     * the applicable status/rule.
     *
     * We therefore report the condition rather than
     * silently ignoring it.
     */

    if (date <= today) {

      addError_(
        errors,
        rowNumber,
        field,
        'Not applicable'
      );
    }
  });
}


/* ==========================================================
   ALPHANUMERIC
   ========================================================== */

function validateAlphanumericField_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues,
  field
) {

  const key =
    normalizeHeader_(field);


  if (columns[key] === undefined) {
    return;
  }


  const value =
    String(
      getValue_(
        values,
        displayValues,
        columns,
        field
      )
    ).trim();


  if (!value) {
    return;
  }


  if (
    !/^[A-Za-z0-9\s\-\/.&(),'#]+$/.test(value)
  ) {

    addError_(
      errors,
      rowNumber,
      field,
      'Please use a valid alphanumeric value.'
    );
  }
}


/* ==========================================================
   NUMERIC
   ========================================================== */

function validateNumericField_(
  errors,
  rowNumber,
  columns,
  values,
  displayValues,
  field
) {

  const key =
    normalizeHeader_(field);


  if (columns[key] === undefined) {
    return;
  }


  const index =
    columns[key];


  const raw =
    values[index];


  const display =
    String(
      displayValues[index] || ''
    ).trim();


  if (!display) {
    return;
  }


  if (
    typeof raw === 'number' &&
    !isNaN(raw)
  ) {
    return;
  }


  const cleaned =
    display
      .replace(/,/g, '')
      .replace(/\s/g, '');


  if (
    !/^-?\d+(\.\d+)?$/.test(cleaned)
  ) {

    addError_(
      errors,
      rowNumber,
      field,
      'Please enter a numeric value with decimals.'
    );
  }
}


/* ==========================================================
   COLUMN MAP
   ========================================================== */

function buildColumnMap_(headers) {

  const map = {};


  headers.forEach(
    function(header, index) {

      const key =
        normalizeHeader_(header);


      if (key) {
        map[key] = index;
      }
    }
  );


  return map;
}


/* ==========================================================
   GET VALUE
   ========================================================== */

function getValue_(
  values,
  displayValues,
  columns,
  header
) {

  const key =
    normalizeHeader_(header);


  if (
    columns[key] === undefined
  ) {
    return '';
  }


  const index =
    columns[key];


  if (
    displayValues[index] !== ''
  ) {
    return displayValues[index];
  }


  return values[index];
}


/* ==========================================================
   NORMALIZATION
   ========================================================== */

function normalize_(value) {

  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}


function normalizeHeader_(header) {

  return String(header || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}


/* ==========================================================
   ADD ERROR
   ========================================================== */

function addError_(
  errors,
  row,
  column,
  message
) {

  errors.push({
    row: row,
    column: column,
    message: message
  });
}


/* ==========================================================
   REMOVE DUPLICATES
   ========================================================== */

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


/* ==========================================================
   OPEN VALIDATION MONITOR
   ========================================================== */

/**
 * Open the persistent HTML validation monitor.
 *
 * This is the ONLY function that directly opens HTML UI.
 */
function openValidationMonitor() {

  const html =
    HtmlService
      .createHtmlOutput(
        getValidationMonitorHtml_()
      )
      .setWidth(720)
      .setHeight(520);


  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      'Procurement Validation Monitor'
    );
}


/* ==========================================================
   GET CURRENT ERRORS
   ========================================================== */

function getCurrentValidationErrors() {

  const props =
    PropertiesService.getUserProperties();


  const json =
    props.getProperty(
      CONFIG.ERROR_PROPERTY
    );


  if (!json) {
    return null;
  }


  try {

    return JSON.parse(json);

  } catch (error) {

    props.deleteProperty(
      CONFIG.ERROR_PROPERTY
    );

    return null;
  }
}


/* ==========================================================
   CLEAR ERRORS
   ========================================================== */

function clearValidationErrors() {

  PropertiesService
    .getUserProperties()
    .deleteProperty(
      CONFIG.ERROR_PROPERTY
    );

  return true;
}


/* ==========================================================
   TEST MODAL
   ========================================================== */

function testValidationModal() {

  const testData = {

    timestamp: Date.now(),

    sheet: 'TEST SHEET',

    errors: [

      {
        row: 5,
        column: 'PROJECT ID',
        message: 'There is no data inputted.'
      },

      {
        row: 5,
        column: 'POSTING DATE',
        message: 'Please use correct date format.'
      },

      {
        row: 5,
        column: 'ELIGIBILITY SCREENING',
        message: 'Not applicable'
      }

    ]
  };


  PropertiesService
    .getUserProperties()
    .setProperty(
      CONFIG.ERROR_PROPERTY,
      JSON.stringify(testData)
    );


  openValidationMonitor();
}


/* ==========================================================
   EMBEDDED HTML
   ========================================================== */

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

#monitor {
  padding: 24px;
}

.header {
  display: flex;
  align-items: center;
  gap: 14px;
  margin-bottom: 18px;
}

.icon {
  width: 46px;
  height: 46px;
  border-radius: 50%;
  background: #f4b400;
  color: white;
  font-size: 27px;
  font-weight: bold;
  display: flex;
  align-items: center;
  justify-content: center;
}

.title {
  font-size: 21px;
  font-weight: bold;
}

.subtitle {
  font-size: 13px;
  color: #666;
  margin-top: 4px;
}

.info {
  background: white;
  border: 1px solid #ddd;
  border-radius: 8px;
  padding: 12px 15px;
  margin-bottom: 14px;
}

.error-list {
  background: white;
  border: 1px solid #ddd;
  border-radius: 8px;
  max-height: 300px;
  overflow-y: auto;
}

.error {
  padding: 13px 15px;
  border-bottom: 1px solid #eee;
}

.error:last-child {
  border-bottom: none;
}

.row {
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

.buttons {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  margin-top: 17px;
}

button {
  border: none;
  border-radius: 6px;
  padding: 10px 20px;
  cursor: pointer;
  font-weight: bold;
  font-size: 14px;
}

.ok {
  background: #1a73e8;
  color: white;
}

.clear {
  background: #e8eaed;
  color: #202124;
}

.empty {
  text-align: center;
  padding: 45px 20px;
  color: #777;
}

.hidden {
  display: none;
}

</style>

</head>


<body>

<div id="monitor">

  <div class="header">

    <div class="icon">
      !
    </div>

    <div>

      <div class="title">
        Procurement Data Validation
      </div>

      <div class="subtitle">
        Monitoring the spreadsheet for validation errors
      </div>

    </div>

  </div>


  <div id="content">

    <div class="empty">
      Monitoring is active.
      <br><br>
      Edit a procurement record to check for errors.
    </div>

  </div>


  <div class="buttons">

    <button
      class="clear"
      onclick="clearErrors()">

      Clear

    </button>


    <button
      class="ok"
      onclick="google.script.host.close()">

      Close

    </button>

  </div>

</div>


<script>

let lastTimestamp = 0;


/*
 * Poll Apps Script for new validation errors.
 */
function checkErrors() {

  google.script.run

    .withSuccessHandler(function(data) {

      if (!data) {
        return;
      }


      if (
        data.timestamp !== lastTimestamp
      ) {

        lastTimestamp =
          data.timestamp;

        displayErrors(data);
      }

    })

    .withFailureHandler(function(error) {

      console.error(error);

    })

    .getCurrentValidationErrors();
}


/*
 * Render errors.
 */
function displayErrors(data) {

  const content =
    document.getElementById('content');


  let html = '';


  html +=
    '<div class="info">' +

      '<strong>Sheet:</strong> ' +
      escapeHtml(data.sheet) +

      '&nbsp;&nbsp; | &nbsp;&nbsp;' +

      '<strong>Errors:</strong> ' +
      data.errors.length +

    '</div>';


  html +=
    '<div class="error-list">';


  data.errors.forEach(function(error) {

    html +=
      '<div class="error">' +

        '<div class="row">' +

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


  content.innerHTML = html;
}


/*
 * Clear errors.
 */
function clearErrors() {

  google.script.run

    .withSuccessHandler(function() {

      lastTimestamp = 0;

      document.getElementById(
        'content'
      ).innerHTML =

        '<div class="empty">' +
        'Monitoring is active.' +
        '<br><br>' +
        'No validation errors.' +
        '</div>';

    })

    .clearValidationErrors();
}


/*
 * HTML escaping.
 */
function escapeHtml(value) {

  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}


/*
 * Check approximately every 700 ms.
 */
setInterval(
  checkErrors,
  ${CONFIG.POLL_MS}
);


/*
 * Check immediately.
 */
checkErrors();

</script>

</body>

</html>
`;
}
