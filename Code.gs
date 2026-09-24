/***************************************************************
 * PROCUREMENT STATUS AUTOMATION
 *
 * Works on the ACTIVE SHEET.
 *
 * The column positions do not matter.
 * Columns are identified by their header names.
 *
 * DATE FORMAT:
 * MMMM d, yyyy
 * Example: September 24, 2026
 *
 * STATUS IS UPDATED:
 * 1. When the spreadsheet is opened
 * 2. When a user edits the spreadsheet
 *
 * ACTIVE STATUS REQUIRED FIELDS:
 * - Pre-procurement
 * - Pre-bid Conference
 * - Project ID
 ***************************************************************/


/***************************************************************
 * ON OPEN
 ***************************************************************/
function onOpen(e) {

  updateAllStatuses();

  SpreadsheetApp.flush();

  // Check Active rows for missing required information.
  checkActiveRequiredFieldsOnOpen();
}


/***************************************************************
 * ON EDIT
 ***************************************************************/
function onEdit(e) {

  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  /*
   * Work on whichever sheet was edited.
   */
  const headerMap = getHeaderMap(sheet);

  if (!headerMap.STATUS) return;

  const firstRow = e.range.getRow();
  const numberOfRows = e.range.getNumRows();

  /*
   * Do not process the header row.
   */
  if (firstRow <= 1) {

    if (firstRow === 1) {
      return;
    }
  }


  /*
   * Update the edited rows.
   */
  for (
    let i = 0;
    i < numberOfRows;
    i++
  ) {

    const row = firstRow + i;

    if (row > 1) {

      updateStatusForRow(
        sheet,
        row,
        headerMap
      );
    }
  }


  SpreadsheetApp.flush();


  /*
   * Check for missing required data.
   *
   * We use a toast here because popup dialogs
   * are unreliable when called directly from
   * an onEdit trigger.
   */
  showActiveMissingDataToast(sheet);
}


/***************************************************************
 * UPDATE ALL STATUSES
 ***************************************************************/
function updateAllStatuses() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getActiveSheet();

  if (!sheet) return;


  const headerMap =
    getHeaderMap(sheet);


  if (!headerMap.STATUS) {

    throw new Error(
      'STATUS column was not found on the active sheet.'
    );
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) return;


  /*
   * Update every data row.
   */
  for (
    let row = 2;
    row <= lastRow;
    row++
  ) {

    updateStatusForRow(
      sheet,
      row,
      headerMap
    );
  }
}


/***************************************************************
 * FIND ALL COLUMN HEADERS
 *
 * Column positions may be anywhere.
 ***************************************************************/
function getHeaderMap(sheet) {

  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {
    return {};
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  const map = {};


  headers.forEach(function(header, index) {

    if (
      header === null ||
      header === undefined
    ) {
      return;
    }


    const normalized =
      normalizeHeader(header);


    const column =
      index + 1;


    switch (normalized) {


      /*
       * PPMP TOTAL COST
       */
      case "PPMP TOTAL COST":

        map.PPMP_TOTAL_COST =
          column;

        break;


      /*
       * PR NO.
       */
      case "PR NO.":
      case "PR NO":

        map.PR_NO =
          column;

        break;


      /*
       * PR TOTAL COST
       */
      case "PR TOTAL COST":

        map.PR_TOTAL_COST =
          column;

        break;


      /*
       * POSTING DATE
       */
      case "POSTING DATE":

        map.POSTING_DATE =
          column;

        break;


      /*
       * SUBMISSION OF BIDS
       */
      case "SUBMISSION OF BIDS":
      case "SUBMISSION OF BID":

        map.SUBMISSION_OF_BIDS =
          column;

        break;


      /*
       * BAC RESOLUTION
       */
      case "BAC RESOLUTION":

        map.BAC_RESOLUTION =
          column;

        break;


      /*
       * NOA DATE
       */
      case "NOA DATE":

        map.NOA_DATE =
          column;

        break;


      /*
       * SUPPLIER
       */
      case "SUPPLIER":

        map.SUPPLIER =
          column;

        break;


      /*
       * NTP DATE
       */
      case "NTP DATE":

        map.NTP_DATE =
          column;

        break;


      /*
       * PO NO.
       */
      case "PO NO":
      case "PO NO.":

        map.PO_NO =
          column;

        break;


      /*
       * PO TOTAL COST
       */
      case "PO TOTAL COST":

        map.PO_TOTAL_COST =
          column;

        break;


      /*
       * REMARKS
       */
      case "REMARKS":

        map.REMARKS =
          column;

        break;


      /*
       * PRE-PROCUREMENT
       */
      case "PRE-PROCUREMENT":
      case "PRE PROCUREMENT":
      case "PREPROCUREMENT":
      case "PRE-PROCUREMENT CONFERENCE":
      case "PRE PROCUREMENT CONFERENCE":
      case "PREPROCUREMENT CONFERENCE":

        map.PRE_PROCUREMENT =
          column;

        break;


      /*
       * PRE-BID CONFERENCE
       */
      case "PRE-BID CONFERENCE":
      case "PRE BID CONFERENCE":
      case "PREBID CONFERENCE":

        map.PRE_BID_CONFERENCE =
          column;

        break;


      /*
       * PROJECT ID
       */
      case "PROJECT ID":

        map.PROJECT_ID =
          column;

        break;


      /*
       * STATUS
       */
      case "STATUS":

        map.STATUS =
          column;

        break;
    }

  });


  return map;
}


/***************************************************************
 * NORMALIZE HEADER
 ***************************************************************/
function normalizeHeader(value) {

  return String(value)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}


/***************************************************************
 * UPDATE STATUS FOR ONE ROW
 ***************************************************************/
function updateStatusForRow(
  sheet,
  row,
  map
) {

  const lastColumn =
    sheet.getLastColumn();


  /*
   * Get actual values.
   */
  const rowValues =
    sheet
      .getRange(
        row,
        1,
        1,
        lastColumn
      )
      .getValues()[0];


  /*
   * Get displayed values.
   *
   * This is useful for values such as
   * "Equal to Zero".
   */
  const rowDisplayValues =
    sheet
      .getRange(
        row,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  /*
   * Get values.
   */
  const ppmpTotalCost =
    getValue(
      rowValues,
      rowDisplayValues,
      map.PPMP_TOTAL_COST
    );


  const postingDate =
    getValue(
      rowValues,
      rowDisplayValues,
      map.POSTING_DATE
    );


  const submissionOfBids =
    getValue(
      rowValues,
      rowDisplayValues,
      map.SUBMISSION_OF_BIDS
    );


  const bacResolution =
    getValue(
      rowValues,
      rowDisplayValues,
      map.BAC_RESOLUTION
    );


  const noaDate =
    getValue(
      rowValues,
      rowDisplayValues,
      map.NOA_DATE
    );


  const supplier =
    getValue(
      rowValues,
      rowDisplayValues,
      map.SUPPLIER
    );


  const ntpDate =
    getValue(
      rowValues,
      rowDisplayValues,
      map.NTP_DATE
    );


  const poNo =
    getValue(
      rowValues,
      rowDisplayValues,
      map.PO_NO
    );


  const poTotalCost =
    getValue(
      rowValues,
      rowDisplayValues,
      map.PO_TOTAL_COST
    );


  const remarks =
    getValue(
      rowValues,
      rowDisplayValues,
      map.REMARKS
    );


  /*
   * Convert dates.
   */
  const posting =
    parseSheetDate(
      postingDate
    );


  const submission =
    parseSheetDate(
      submissionOfBids
    );


  const today =
    startOfDay(
      new Date()
    );


  let status = "";


  /*************************************************************
   * RULE 1
   *
   * REALIGNED ITEM
   *
   * PPMP TOTAL COST:
   * 0
   * 0.00
   * zero
   * Equal to Zero
   *************************************************************/
  if (
    isRealignedItem(
      ppmpTotalCost
    )
  ) {

    status =
      "Realigned Item";
  }


  /*************************************************************
   * RULE 2
   *
   * CANCELLED PR
   *************************************************************/
  else if (
    normalizeText(remarks) ===
    "CANCELLED PR"
  ) {

    status =
      "Cancelled PR";
  }


  /*************************************************************
   * RULE 3
   *
   * CANCELLED PO
   *************************************************************/
  else if (
    normalizeText(remarks) ===
      "CANCELLED PO" &&

    hasValue(bacResolution) &&
    hasValue(noaDate) &&
    hasValue(supplier) &&
    hasValue(ntpDate)
  ) {

    status =
      "Cancelled PO";
  }


  /*************************************************************
   * RULE 4
   *
   * PURCHASE ORDER
   *************************************************************/
  else if (

    hasValue(bacResolution) &&
    hasValue(noaDate) &&
    hasValue(supplier) &&
    hasValue(ntpDate) &&
    hasValue(poNo) &&
    hasValue(poTotalCost)

  ) {

    status =
      "Purchase Order";
  }


  /*************************************************************
   * RULE 5
   *
   * AWARDED
   *************************************************************/
  else if (

    hasValue(bacResolution) &&
    hasValue(noaDate) &&
    hasValue(supplier)

  ) {

    status =
      "Awarded";
  }


  /*************************************************************
   * RULE 6
   *
   * FAILED
   *
   * BAC Resolution exists
   * NOA Date is blank
   *************************************************************/
  else if (

    hasValue(bacResolution) &&
    !hasValue(noaDate)

  ) {

    status =
      "Failed";
  }


  /*************************************************************
   * RULE 7
   *
   * ACTIVE
   *
   * Posting Date exists
   * Submission of Bids exists
   * Submission of Bids is on or before today
   * BAC Resolution is blank
   *************************************************************/
  else if (

    posting &&
    submission &&
    submission <= today &&
    !hasValue(bacResolution)

  ) {

    status =
      "Active";
  }


  /*************************************************************
   * RULE 8
   *
   * CLOSED
   *
   * Posting Date exists
   * Submission of Bids exists
   * Submission of Bids is after today
   * BAC Resolution is blank
   *************************************************************/
  else if (

    posting &&
    submission &&
    submission > today &&
    !hasValue(bacResolution)

  ) {

    status =
      "Closed";
  }


  /*************************************************************
   * WRITE STATUS
   *************************************************************/
  if (map.STATUS) {

    const statusCell =
      sheet.getRange(
        row,
        map.STATUS
      );


    const currentStatus =
      String(
        statusCell.getDisplayValue()
      ).trim();


    if (
      currentStatus !== status
    ) {

      statusCell.setValue(
        status
      );
    }
  }
}


/***************************************************************
 * REALIGNED ITEM CHECK
 ***************************************************************/
function isRealignedItem(value) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return false;
  }


  /*
   * Numeric zero.
   *
   * This catches:
   * 0
   * 0.00
   * 0.000
   */
  if (
    typeof value === "number"
  ) {

    return value === 0;
  }


  const text =
    String(value)
      .trim()
      .toUpperCase();


  /*
   * Text zero.
   */
  if (
    text === "ZERO"
  ) {

    return true;
  }


  /*
   * Equal to Zero.
   */
  if (
    text === "EQUAL TO ZERO"
  ) {

    return true;
  }


  /*
   * Text numeric zero.
   */
  if (
    text === "0" ||
    text === "0.00"
  ) {

    return true;
  }


  return false;
}


/***************************************************************
 * CHECK WHETHER A VALUE EXISTS
 ***************************************************************/
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


/***************************************************************
 * GET VALUE FROM ROW
 ***************************************************************/
function getValue(
  rowValues,
  rowDisplayValues,
  column
) {

  if (!column) return "";


  const index =
    column - 1;


  const actualValue =
    rowValues[index];


  const displayValue =
    rowDisplayValues[index];


  if (
    actualValue !== null &&
    actualValue !== undefined &&
    actualValue !== ""
  ) {

    return actualValue;
  }


  return displayValue || "";
}


/***************************************************************
 * NORMALIZE TEXT
 ***************************************************************/
function normalizeText(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";
  }


  return String(value)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}


/***************************************************************
 * PARSE SHEET DATE
 *
 * Handles:
 *
 * September 24, 2026
 * September 24 2026
 * Sep 24, 2026
 * Google Sheets Date objects
 ***************************************************************/
function parseSheetDate(value) {

  if (!value) {
    return null;
  }


  /*
   * Google Sheets Date object.
   */
  if (
    Object.prototype.toString.call(value) ===
    "[object Date]"
  ) {

    if (
      isNaN(
        value.getTime()
      )
    ) {

      return null;
    }


    return startOfDay(
      value
    );
  }


  /*
   * Text date.
   */
  if (
    typeof value === "string"
  ) {

    const text =
      value.trim();


    if (!text) {
      return null;
    }


    const parsed =
      new Date(text);


    if (
      !isNaN(
        parsed.getTime()
      )
    ) {

      return startOfDay(
        parsed
      );
    }
  }


  return null;
}


/***************************************************************
 * START OF DAY
 ***************************************************************/
function startOfDay(date) {

  const result =
    new Date(date);


  result.setHours(