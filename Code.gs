/**
 * PROCUREMENT STATUS AUTOMATION
 *
 * Runs:
 * 1. Automatically when the spreadsheet is opened
 * 2. Automatically when data is edited
 *
 * Date format:
 * MMMM d, yyyy
 * Example: September 24, 2026
 *
 * IMPORTANT:
 * - Column positions do NOT matter.
 * - The script finds columns using their header names.
 * - Header names are matched case-insensitively.
 */


const SHEET_NAME = "Data List";


/**
 * ============================================================
 * SETUP INSTALLABLE TRIGGERS
 * ============================================================
 *
 * Run setupTriggers() ONCE manually from Apps Script.
 *
 * This creates installable triggers for:
 * - Spreadsheet open
 * - Spreadsheet edit
 *
 * Installable triggers are required because modal dialogs use
 * SpreadsheetApp.getUi(), which cannot be reliably called from
 * a simple onEdit trigger.
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === "onOpen" || handler === "onEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // onOpen stays as the simple trigger so it can interact with
  // the user's open spreadsheet UI.
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    "Setup Complete",
    "The edit trigger has been installed. Close and reopen the spreadsheet to test validation.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * ============================================================
 * ON OPEN
 * ============================================================
 *
 * Updates statuses and displays the validation modal when the
 * spreadsheet is opened.
 */
function onOpen(e) {
  // Runs automatically when the spreadsheet is opened.
  // No custom menu is created.
  if (!e) return;

  try {
    const invalidDateFields = updateAllStatuses();
    SpreadsheetApp.flush();

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);

    if (sheet) {
      showValidationDialog(
        invalidDateFields,
        getActiveMissingData(sheet)
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "Validation Error",
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/**
 * ============================================================
 * ON EDIT
 * ============================================================
 *
 * This is the main edit handler.
 *
 * IMPORTANT:
 * To display a modal after an edit, this function must be
 * installed as a Spreadsheet -> On edit trigger.
 *
 * Do not create any custom menu for this feature.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== SHEET_NAME) return;
  if (e.range.getRow() === 1) return;

  const headerMap = getHeaderMap(sheet);

  if (!headerMap.STATUS) return;

  const firstRow = e.range.getRow();
  const numberOfRows = e.range.getNumRows();
  const invalidDateFields = [];

  for (let i = 0; i < numberOfRows; i++) {
    const row = firstRow + i;

    const invalidFields = normalizeInputDates(
      sheet,
      row,
      headerMap
    );

    invalidFields.forEach(function(field) {
      invalidDateFields.push("Row " + row + ": " + field);
    });

    updateStatusForRow(sheet, row, headerMap);
  }

  SpreadsheetApp.flush();
}


/**
 * ============================================================
 * MANUAL VALIDATION MODAL TEST
 * ============================================================
 */
function openValidationModal() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      "Validation Error",
      'Sheet "' + SHEET_NAME + '" was not found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const headerMap = getHeaderMap(sheet);
  const invalidDateFields = collectInvalidDateFields(sheet, headerMap);
  const missingRows = getActiveMissingData(sheet);

  if (invalidDateFields.length === 0 && missingRows.length === 0) {
    SpreadsheetApp.getUi().alert(
      "Validation Complete",
      "No invalid date formats or missing ACTIVE-row data were found.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  showValidationDialog(invalidDateFields, missingRows);
}


/**
 * ============================================================
 * COMBINED VALIDATION DIALOG
 * ============================================================
 *
 * Displays one dialog for both:
 * 1. Invalid date formats
 * 2. Missing required data in ACTIVE rows
 *
 * If both problems exist, they are displayed together.
 */
function showValidationDialog(invalidDateFields, missingRows) {
  const hasInvalidDates = invalidDateFields && invalidDateFields.length > 0;
  const hasMissingData = missingRows && missingRows.length > 0;

  if (!hasInvalidDates && !hasMissingData) return;

  let sectionsHtml = "";

  if (hasInvalidDates) {
    const dateHtml = invalidDateFields.map(function(item) {
      return '<div class="item error-item">' +
        escapeHtml(item) +
      '</div>';
    }).join("");

    sectionsHtml +=
      '<div class="section">' +
        '<h3>Invalid Date Format</h3>' +
        '<p class="section-note">All currently invalid date fields are listed below. Use <b>MMMM d, yyyy</b>, for example: September 24, 2026. You may also enter <b>N/A</b>, <b>NA</b>, or <b>Not Applicable</b>.</p>' +
        '<div class="list">' + dateHtml + '</div>' +
      '</div>';
  }

  if (hasMissingData) {
    const missingHtml = missingRows.map(function(item) {
      const fields = item.fields.map(function(field) {
        return escapeHtml(field);
      }).join(", ");

      return '<div class="item">' +
        '<div class="row-number">Row ' + item.row + '</div>' +
        '<div class="fields">' + fields + '</div>' +
      '</div>';
    }).join("");

    sectionsHtml +=
      '<div class="section">' +
        '<h3>Required Data Missing</h3>' +
        '<p class="section-note">The following ACTIVE row(s) are missing required data:</p>' +
        '<div class="list">' + missingHtml + '</div>' +
      '</div>';
  }

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html><head><base target="_top">' +
    '<style>' +
      '*{box-sizing:border-box}' +
      'html,body{margin:0;padding:0;width:100%;height:100%;font-family:Arial,sans-serif;background:#fff;color:#202124}' +
      'body{display:flex;align-items:center;justify-content:center;padding:24px}' +
      '.dialog{width:100%;max-width:520px;margin:auto}' +
      '.header{text-align:center;margin-bottom:18px}' +
      '.icon{width:48px;height:48px;margin:0 auto 10px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#fff3cd;color:#8a6d1d;font-size:25px;font-weight:bold}' +
      'h2{margin:0;font-size:20px;font-weight:600}' +
      '.intro{margin:8px 0 18px;text-align:center;color:#5f6368;font-size:13px;line-height:1.45}' +
      '.section{margin-bottom:16px}' +
      '.section:last-child{margin-bottom:0}' +
      'h3{margin:0 0 6px;font-size:14px;font-weight:600}' +
      '.section-note{margin:0 0 8px;color:#5f6368;font-size:12px;line-height:1.4}' +
      '.list{max-height:150px;overflow-y:auto;border:1px solid #dadce0;border-radius:8px;background:#f8f9fa;padding:8px}' +
      '.item{background:#fff;border:1px solid #e0e3e7;border-radius:6px;padding:9px 11px;margin-bottom:7px;font-size:13px;line-height:1.4}' +
      '.item:last-child{margin-bottom:0}' +
      '.error-item{color:#b3261e}' +
      '.row-number{font-weight:600}' +
      '.fields{color:#5f6368;margin-top:3px}' +
      '.footer{text-align:center;margin-top:18px}' +
      '.note{margin:0 0 13px;color:#5f6368;font-size:12px}' +
      'button{min-width:100px;border:0;border-radius:6px;padding:9px 22px;background:#1a73e8;color:#fff;font-size:13px;font-weight:600;cursor:pointer}' +
      'button:hover{background:#1765cc}' +
    '</style></head>' +
    '<body><div class="dialog">' +
      '<div class="header">' +
        '<div class="icon">!</div>' +
        '<h2>Procurement Data Validation</h2>' +
        '<p class="intro">Please review the following item(s) before continuing.</p>' +
      '</div>' +
      sectionsHtml +
      '<div class="footer">' +
        '<p class="note">Correct the items above and continue entering the procurement data.</p>' +
        '<button onclick="google.script.host.close()">OK</button>' +
      '</div>' +
    '</div></body></html>'
  ).setWidth(580).setHeight(hasInvalidDates && hasMissingData ? 560 : 440);

  SpreadsheetApp.getUi().showModalDialog(html, "Procurement Data Validation");
}

/**
 * ============================================================
 * UPDATE ALL STATUSES
 * ============================================================
 */
function updateAllStatuses() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet "' + SHEET_NAME + '" was not found.'
    );
  }

  const headerMap = getHeaderMap(sheet);

  if (!headerMap.STATUS) {
    throw new Error(
      'STATUS column was not found in "' + SHEET_NAME + '".'
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const invalidDateFields = [];

  for (let row = 2; row <= lastRow; row++) {

    const invalidFields = normalizeInputDates(
      sheet,
      row,
      headerMap
    );

    invalidFields.forEach(function(field) {
      invalidDateFields.push(
        "Row " + row + ": " + field
      );
    });

    updateStatusForRow(sheet, row, headerMap);
  }

  return invalidDateFields;
}


/**
 * ============================================================
 * COLLECT ALL INVALID DATE FIELDS
 * ============================================================
 *
 * Scans every data row and every supported date column.
 * This ensures the validation dialog lists all invalid date
 * fields currently present in the sheet.
 */
function collectInvalidDateFields(sheet, map) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const invalidDateFields = [];

  for (let row = 2; row <= lastRow; row++) {
    const invalidFields = normalizeInputDates(
      sheet,
      row,
      map
    );

    invalidFields.forEach(function(field) {
      invalidDateFields.push(
        "Row " + row + ": " + field
      );
    });
  }

  return invalidDateFields;
}


/**
 * ============================================================
 * FIND COLUMN HEADERS
 * ============================================================
 *
 * Column positions can be anywhere in the sheet.
 */
function getHeaderMap(sheet) {

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    return {};
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const map = {};

  headers.forEach(function(header, index) {

    if (!header) return;

    const normalized = normalizeHeader(header);

    const column = index + 1;

    switch (normalized) {

    case "PRE PROCUREMENT CONFERENCE":
    map.PRE_PROCUREMENT = column;
    break;

    case "PRE BID CONFERENCE":
    map.PRE_BID_CONFERENCE = column;
    break;

    case "PROJECT ID":
    map.PROJECT_ID = column;
    break;

      case "TOTAL ABC":
        map.PPMP_TOTAL_COST = column;
        break;

      case "PR NO":
        map.PR_NO = column;
        break;

      case "PR TOTAL ABC":
        map.PR_TOTAL_COST = column;
        break;

      case "POSTING DATE":
        map.POSTING_DATE = column;
        break;

      case "SUBMISSION OF BIDS":
      case "SUBMISSION OF BID":
        map.SUBMISSION_OF_BIDS = column;
        break;

      case "ELIGIBILITY SCREENING":
      case "ELIGIBILITY SCREEN":
        map.ELIGIBILITY_SCREENING = column;
        break;

      case "BAC RESOLUTION NO":
        map.BAC_RESOLUTION = column;
        break;

      case "NOA DATE":
        map.NOA_DATE = column;
        break;

      case "SUPPLIER":
        map.SUPPLIER = column;
        break;

      case "NTP DATE":
        map.NTP_DATE = column;
        break;

      case "PO NO":
        map.PO_NO = column;
        break;

      case "PO TOTAL COST":
        map.PO_TOTAL_COST = column;
        break;

      case "REMARKS":
        map.REMARKS = column;
        break;

      case "STATUS":
        map.STATUS = column;
        break;
    }
  });

  return map;
}
/**
 * ============================================================
 * CHECK REQUIRED DATA FOR ACTIVE STATUS
 * ============================================================
 *
 * Every row whose STATUS is ACTIVE must contain:
 * - Pre-Procurement Conference
 * - Pre-Bid Conference
 * - Project ID
 *
 * Blank values are detected using the actual cell value,
 * so empty cells, empty strings, and whitespace are treated
 * as missing.
 */
function getActiveMissingData(sheet) {
  const map = getHeaderMap(sheet);

  // Validation cannot run unless all required headers exist.
  if (
    !map.STATUS ||
    !map.PRE_PROCUREMENT ||
    !map.PRE_BID_CONFERENCE ||
    !map.PROJECT_ID
  ) {
    return [];
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return [];
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getValues();

  const missingRows = [];

  values.forEach(function(row, index) {
    const actualRow = index + 2;

    // Normalize the status so "Active", "ACTIVE", etc. are equivalent.
    const status = String(
      row[map.STATUS - 1] === null ||
      row[map.STATUS - 1] === undefined
        ? ""
        : row[map.STATUS - 1]
    ).trim().toUpperCase();

    if (status !== "ACTIVE") {
      return;
    }

    const missingFields = [];

    if (!hasValue(row[map.PRE_PROCUREMENT - 1])) {
      missingFields.push("PRE-PROCUREMENT CONFERENCE");
    }

    if (!hasValue(row[map.PRE_BID_CONFERENCE - 1])) {
      missingFields.push("PRE-BID CONFERENCE");
    }

    if (!hasValue(row[map.PROJECT_ID - 1])) {
      missingFields.push("PROJECT ID");
    }

    if (missingFields.length > 0) {
      missingRows.push({
        row: actualRow,
        fields: missingFields
      });
    }
  });

  return missingRows;
}

/**
 * ============================================================
 * CHECK REQUIRED DATA FOR ACTIVE STATUS
 * ============================================================
 */
function showActiveMissingDataModal() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  if (!sheet) return;

  const missingRows = getActiveMissingData(sheet);

  if (missingRows.length === 0) return;

  showCenteredMissingDataDialog(missingRows);
}


/**
 * ============================================================
 * CENTERED ACTIVE MISSING-DATA MODAL
 * ============================================================
 *
 * The dialog HTML is generated directly in Code.gs.
 * No separate HTML file is required.
 */
function showCenteredMissingDataDialog(missingRows) {

  const rowsHtml = missingRows.map(function(item) {
    const safeFields = item.fields
      .map(function(field) {
        return escapeHtml(field);
      })
      .join(", ");

    return (
      '<div class="row-item">' +
        '<div class="row-number">Row ' + item.row + '</div>' +
        '<div class="fields">' + safeFields + '</div>' +
      '</div>'
    );
  }).join("");

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html><head><base target="_top">' +
    '<style>' +
      '*{box-sizing:border-box}' +
      'html,body{margin:0;padding:0;width:100%;height:100%;font-family:Arial,sans-serif;background:#fff;color:#202124}' +
      'body{display:flex;align-items:center;justify-content:center;padding:24px}' +
      '.dialog{width:100%;max-width:510px;margin:auto}' +
      '.header{text-align:center;margin-bottom:18px}' +
      '.icon{width:48px;height:48px;margin:0 auto 10px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#fff3cd;color:#8a6d1d;font-size:25px;font-weight:bold}' +
      'h2{margin:0;font-size:20px;font-weight:600}' +
      '.intro{margin:8px 0 0;text-align:center;color:#5f6368;font-size:13px;line-height:1.45}' +
      '.list{max-height:235px;overflow-y:auto;border:1px solid #dadce0;border-radius:8px;background:#f8f9fa;padding:8px}' +
      '.row-item{background:#fff;border:1px solid #e0e3e7;border-radius:6px;padding:9px 11px;margin-bottom:7px;font-size:13px;line-height:1.4}' +
      '.row-item:last-child{margin-bottom:0}' +
      '.row-number{font-weight:600}' +
      '.fields{color:#5f6368;margin-top:3px}' +
      '.footer{text-align:center;margin-top:16px}' +
      '.note{margin:0 0 13px;color:#5f6368;font-size:12px}' +
      'button{min-width:100px;border:0;border-radius:6px;padding:9px 22px;background:#1a73e8;color:#fff;font-size:13px;font-weight:600;cursor:pointer}' +
      'button:hover{background:#1765cc}' +
    '</style></head>' +
    '<body><div class="dialog">' +
      '<div class="header">' +
        '<div class="icon">!</div>' +
        '<h2>Required Data Missing</h2>' +
        '<p class="intro">The following ACTIVE row(s) have missing required data:</p>' +
      '</div>' +
      '<div class="list">' + rowsHtml + '</div>' +
      '<div class="footer">' +
        '<p class="note">Please enter the required data for the Active Status row(s).</p>' +
        '<button onclick="google.script.host.close()">OK</button>' +
      '</div>' +
    '</div></body></html>'
  ).setWidth(560).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, "Required Data Missing");
}


/**
 * Escapes spreadsheet values before placing them in HTML.
 */
function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}


/**
 * ============================================================
 * NORMALIZE HEADER
 * ============================================================
 */
function normalizeHeader(value) {

  return String(value)
    .trim()
    .toUpperCase()
    .replace(/[.\-_/]+/g, " ")
    .replace(/\s+/g, " ");
}


/**
 * ============================================================
 * NORMALIZE INPUT DATES
 * ============================================================
 *
 * The following columns accept either:
 *
 * 1. MMMM d, yyyy
 *    Example: September 24, 2026
 * 2. N/A
 * 3. NA
 * 4. Not Applicable
 *
 * Accepted N/A values are treated as intentional non-date entries
 * and are not reported as invalid date formats.
 */
function normalizeInputDates(sheet, row, map) {

  const dateColumns = [
    { column: map.PRE_PROCUREMENT, name: "PRE-PROCUREMENT CONFERENCE" },
    { column: map.PRE_BID_CONFERENCE, name: "PRE-BID CONFERENCE" },
    { column: map.POSTING_DATE, name: "POSTING DATE" },
    { column: map.ELIGIBILITY_SCREENING, name: "ELIGIBILITY SCREENING" },
    { column: map.SUBMISSION_OF_BIDS, name: "SUBMISSION OF BIDS" }
  ];

  const invalidFields = [];

  dateColumns.forEach(function(item) {

    const column = item.column;

    if (!column) return;

    const cell = sheet.getRange(row, column);
    const value = cell.getValue();

    if (!hasValue(value)) return;

    /* Already a real date value. */
    if (Object.prototype.toString.call(value) === "[object Date]") {
      if (!isNaN(value.getTime())) {
        cell.setNumberFormat("mmmm d, yyyy");
      }
      return;
    }

    /* Accept N/A, NA, and Not Applicable as valid non-date values. */
    if (typeof value === "string" && isAcceptedNonDateValue(value)) {
      return;
    }

    /* Convert text entered as MMMM d, yyyy into a real Date. */
    if (typeof value === "string") {
      const parsed = parseInputDate(value);

      if (parsed) {
        cell.setValue(parsed);
        cell.setNumberFormat("mmmm d, yyyy");
      } else {
        invalidFields.push(item.name);
      }
    }
  });

  return invalidFields;
}

/**
 * Parses the requested input format:
 * MMMM d, yyyy
 *
 * Example: September 24, 2026
 */
function parseInputDate(value) {

  const text = String(value).trim();

  if (!text || isAcceptedNonDateValue(text)) return null;

  const match = text.match(
    /^(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})$/i
  );

  if (!match) return null;

  const monthNames = [
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december"
  ];

  const month = monthNames.indexOf(match[1].toLowerCase());
  const day = Number(match[2]);
  const year = Number(match[3]);

  const date = new Date(year, month, day);

  /* Reject impossible dates such as February 30. */
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month ||
    date.getDate() !== day
  ) {
    return null;
  }

  return startOfDay(date);
}


/**
 * ============================================================
 * ACCEPTED NON-DATE VALUES
 * ============================================================
 *
 * These values are valid alternatives to a date in the date
 * validation columns:
 * - N/A
 * - NA
 * - Not Applicable
 */
function isAcceptedNonDateValue(value) {
  const text = String(value).trim().toUpperCase();

  return (
    text === "N/A" ||
    text === "NA" ||
    text === "NOT APPLICABLE"
  );
}


/**
 * ============================================================
 * UPDATE STATUS FOR ONE ROW
 * ============================================================
 */
function updateStatusForRow(sheet, row, map) {

  const lastColumn = sheet.getLastColumn();

  const rowValues = sheet
    .getRange(row, 1, 1, lastColumn)
    .getValues()[0];

  const rowDisplayValues = sheet
    .getRange(row, 1, 1, lastColumn)
    .getDisplayValues()[0];


  /**
   * ----------------------------------------------------------
   * GET VALUES
   * ----------------------------------------------------------
   */

  const ppmpTotalCost = getValue(
    rowValues,
    rowDisplayValues,
    map.PPMP_TOTAL_COST
  );

  const postingDate = getValue(
    rowValues,
    rowDisplayValues,
    map.POSTING_DATE
  );

  const submissionOfBids = getValue(
    rowValues,
    rowDisplayValues,
    map.SUBMISSION_OF_BIDS
  );

  const eligibilityScreening = getValue(  
    rowValues,
    rowDisplayValues,
    map.ELIGIBILITY_SCREENING
  );

  const bacResolution = getValue(
    rowValues,
    rowDisplayValues,
    map.BAC_RESOLUTION
  );

  const noaDate = getValue(
    rowValues,
    rowDisplayValues,
    map.NOA_DATE
  );

  const supplier = getValue(
    rowValues,
    rowDisplayValues,
    map.SUPPLIER
  );

  const ntpDate = getValue(
    rowValues,
    rowDisplayValues,
    map.NTP_DATE
  );

  const poNo = getValue(
    rowValues,
    rowDisplayValues,
    map.PO_NO
  );

  const poTotalCost = getValue(
    rowValues,
    rowDisplayValues,
    map.PO_TOTAL_COST
  );

  const remarks = getValue(
    rowValues,
    rowDisplayValues,
    map.REMARKS
  );


  /**
   * ----------------------------------------------------------
   * DATE VALUES
   * ----------------------------------------------------------
   */

  const posting = parseSheetDate(postingDate);

  const submission = parseSheetDate(submissionOfBids);

  const eligibility = parseSheetDate(eligibilityScreening);

  const today = startOfDay(new Date());


  /**
   * ----------------------------------------------------------
   * STATUS DETERMINATION
   * ----------------------------------------------------------
   */

  let status = "";


  /*
   * RULE 1
   *
   * PPMP Total Cost = Equal to Zero
   *
   * STATUS = Realigned Item
   */
  if (
    String(ppmpTotalCost).trim().toUpperCase() === "EQUAL TO ZERO" ||
    String(ppmpTotalCost).trim() === "0" ||
    Number(ppmpTotalCost) === 0
  ) {

    status = "Realigned Item";
  }


  /*
   * RULE 2
   *
   * Cancelled PR
   */
  else if (
    String(remarks)
      .trim()
      .toUpperCase() === "CANCELLED PR"
  ) {

    status = "Cancelled PR";
  }


  /*
   * RULE 3
   *
   * Cancelled PO
   *
   * BAC Resolution
   * NOA Date
   * Supplier
   * NTP Date
   */
  else if (
    String(remarks)
      .trim()
      .toUpperCase() === "CANCELLED PO" &&
    hasValue(bacResolution) &&
    hasValue(noaDate) &&
    hasValue(supplier) &&
    hasValue(ntpDate)
  ) {

    status = "Cancelled PO";
  }


  /*
   * RULE 4
   *
   * PURCHASE ORDER
   *
   * BAC Resolution
   * NOA Date
   * Supplier
   * NTP Date
   * PO No.
   * PO Total Cost
   */
  else if (
    hasValue(bacResolution) &&
    hasValue(noaDate) &&
    hasValue(supplier) &&
    hasValue(ntpDate) &&
    hasValue(poNo) &&
    hasValue(poTotalCost)
  ) {

    status = "Purchase Order";
  }


  /*
   * RULE 5
   *
   * AWARDED
   *
   * BAC Resolution
   * NOA Date
   * Supplier
   *
   * NTP / PO information is not yet complete.
   */
  else if (
    hasValue(bacResolution) &&
    hasValue(noaDate) &&
    hasValue(supplier)
  ) {

    status = "Awarded";
  }


  /*
   * RULE 6
   *
   * FAILED
   *
   * BAC Resolution is present
   * NOA Date is blank
   */
  else if (
    hasValue(bacResolution) &&
    !hasValue(noaDate)
  ) {

    status = "Failed";
  }


  /*
   * RULE 7
   *
   * ACTIVE
   *
   * Posting Date exists
   * Submission of Bids exists
   * Submission of Bids is ON or BEFORE TODAY
   * BAC Resolution is blank
   */
  else if (
    posting &&
    submission &&
    submission <= today && eligibility && eligibility <= today &&
    !hasValue(bacResolution)
  ) {

    status = "Active";
  }


  /*
   * RULE 8
   *
   * CLOSED
   *
   * Posting Date exists
   * Submission of Bids exists
   * Submission of Bids is AFTER TODAY
   * BAC Resolution is blank
   */
  else if (
    posting &&
    submission &&
    submission > today && eligibility && eligibility > today &&
    !hasValue(bacResolution)
  ) {

    status = "Closed";
  }


  /*
   * ----------------------------------------------------------
   * WRITE STATUS
   * ----------------------------------------------------------
   */

  if (map.STATUS) {

    const statusCell = sheet.getRange(
      row,
      map.STATUS
    );

    const currentStatus =
      String(statusCell.getDisplayValue()).trim();

    if (currentStatus !== status) {

      statusCell.setValue(status);
    }
  }
}


/**
 * ============================================================
 * GET VALUE
 * ============================================================
 */
function getValue(rowValues, rowDisplayValues, column) {

  if (!column) return "";

  const index = column - 1;

  const actualValue = rowValues[index];

  const displayValue = rowDisplayValues[index];

  /*
   * Prefer actual date/value when available.
   */
  if (
    actualValue !== null &&
    actualValue !== undefined &&
    actualValue !== ""
  ) {

    return actualValue;
  }

  return displayValue || "";
}


/**
 * ============================================================
 * CHECK IF VALUE EXISTS
 * ============================================================
 */
function hasValue(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return false;
  }

  if (
    typeof value === "string" &&
    value.trim() === ""
  ) {

    return false;
  }

  return true;
}


/**
 * ============================================================
 * PARSE DATE
 * ============================================================
 *
 * Supports:
 *
 * September 24, 2026
 * September 24 2026
 * Sep 24, 2026
 * N/A / NA / Not Applicable (returns null)
 * Google Sheets Date objects
 */
function parseSheetDate(value) {

  if (!value) return null;

  /*
   * Accepted non-date values.
   * These are intentionally treated as having no date.
   */
  if (isAcceptedNonDateValue(value)) {
    return null;
  }


  /*
   * Google Sheets date object
   */
  if (Object.prototype.toString.call(value) === "[object Date]") {

    if (isNaN(value.getTime())) {
      return null;
    }

    return startOfDay(value);
  }


  /*
   * Text date
   */
  if (typeof value === "string") {

    const text = value.trim();

    if (!text) return null;


    /*
     * Convert:
     *
     * September 24, 2026
     *
     * to a JavaScript date.
     */
    const parsed = new Date(text);

    if (!isNaN(parsed.getTime())) {

      return startOfDay(parsed);
    }
  }


  return null;
}


/**
 * ============================================================
 * START OF DAY
 * ============================================================
 */
function startOfDay(date) {

  const result = new Date(date);

  result.setHours(
    0,
    0,
    0,
    0
  );

  return result;
}
