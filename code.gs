// ============================================
// GOOGLE APPS SCRIPT CODE
// Deploy this as a Web App
// ============================================

// Configuration - Update this with your Google Sheet ID
const SHEET_ID = '1ijXeQPdci9xI1u622vikjA2xGlCRhIVNAPCf93Slg1A';
const SHEET_NAME = 'Release Order Tracker';

// ==================== MAIN HANDLER ====================
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getNextRONumber') {
    return getNextRONumber();
  }

  // NEW: Action to fetch details of an existing RO
  if (action === 'fetchRO') {
    return fetchRODetails(e.parameter.roNumber);
  }

  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    message: 'Invalid action'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;

    if (action === 'saveReleaseOrder') {
      return saveReleaseOrder(requestData.data);
    }

    // NEW: Action to update an existing RO
    if (action === 'updateReleaseOrder') {
      return updateReleaseOrder(requestData.data);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Invalid action'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== GET NEXT R.O. NUMBER ====================
// ==================== GET DYNAMIC R.O. NUMBER ====================
function getNextRONumber() {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();

    // 1. Calculate Current Financial Year (Starts April 1st) [cite: 54]
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth(); // 0 = Jan, 3 = April

    let startYear, endYear;
    if (month < 3) { // If Jan, Feb, or Mar, we are still in the previous FY
      startYear = year - 1;
      endYear = year;
    } else { // From April 1st onwards, we start the new FY
      startYear = year;
      endYear = year + 1;
    }

    // Creates format "26-27" [cite: 69]
    const currentFY = String(startYear).slice(-2) + '-' + String(endYear).slice(-2);

    // 2. If the sheet is empty, start fresh with the current FY [cite: 65]
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        nextRONumber: currentFY + '/0001'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Get and parse the last recorded R.O. number [cite: 66, 67]
    const lastRONumber = sheet.getRange(lastRow, 1).getValue();
    if (!lastRONumber) {
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        nextRONumber: currentFY + '/0001'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const parts = lastRONumber.split('/');
    const lastFY = parts[0]; // e.g., "25-26" [cite: 68]
    const lastNum = parseInt(parts[1]); // e.g., 0002 [cite: 70]

    // 4. Reset count if the Financial Year has changed
    let nextRONumber;
    if (lastFY === currentFY) {
      // Same year: Increment by 1 
      const nextCount = (lastNum + 1).toString().padStart(4, '0');
      nextRONumber = currentFY + '/' + nextCount;
    } else {
      // New year started: Reset to 0001
      nextRONumber = currentFY + '/0001';
    }

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      nextRONumber: nextRONumber
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== SAVE RELEASE ORDER ====================
function saveReleaseOrder(data) {
  try {
    const sheet = getOrCreateSheet();

    // Append new row with data
    sheet.appendRow([
      data.roNumber,
      data.date,
      data.publicationName,
      data.edition,
      data.client,
      data.scheduledDateSS,
      data.clientExtra,
      data.scheduledDate,
      data.dateExtra,
      data.size,
      data.templateType,
      data.position,
      data.material,
      data.rate,
      data.sizeSS
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Release order saved successfully'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== NEW FUNCTIONS ====================

// Helper to safely format Google Sheets Date objects back to DD/MM/YYYY
function formatSheetDate(dateVal) {
  if (dateVal instanceof Date) {
    const dd = String(dateVal.getDate()).padStart(2, '0');
    const mm = String(dateVal.getMonth() + 1).padStart(2, '0');
    const yyyy = dateVal.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }
  return dateVal; // Return as-is if it's already a text string
}

function fetchRODetails(roNumber) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  // OPTIMIZATION: Loop backwards (from last row up to row 1)
  // Most likely to find recent ROs quickly
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] == roNumber) { // Column 0 is RO Number
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: {
          roNumber: data[i][0],
          date: formatSheetDate(data[i][1]), // Safely format the Generation Date
          publicationName: data[i][2],
          edition: data[i][3],
          client: data[i][4],
          clientExtra: data[i][6],          // NEW
          scheduledDate: formatSheetDate(data[i][7]), // Safely format the Scheduled Date
          dateExtra: data[i][8],            // NEW
          size: data[i][9],
          templateType: data[i][10],
          position: data[i][11],
          material: data[i][12],
          rate: data[i][13]
        }
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    message: 'RO Number not found'
  })).setMimeType(ContentService.MimeType.JSON);
}

function updateReleaseOrder(data) {
  const sheet = getOrCreateSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();

  // OPTIMIZATION: Loop backwards
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][0] == data.roNumber) {
      // Found it! Update the row
      const rowData = [
        data.roNumber,
        data.date,
        data.publicationName,
        data.edition,
        data.client,
        data.clientExtra,
        data.scheduledDateSS,
        data.scheduledDate,
        data.dateExtra,
        data.size,
        data.templateType,
        data.position,
        data.material,
        data.rate,
        data.sizeSS
      ];

      // i + 1 because sheet rows are 1-based
      sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: 'RO Updated Successfully'
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    message: 'RO Number to update was not found'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==================== GET OR CREATE SHEET ====================
function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  // If sheet doesn't exist, create it
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);

    // Add headers
    const headers = [
      'R.O. No.',
      'Date',
      'Publication Name',
      'Edition',
      'Client',
      'Schedule Date',
      'Client Extra',
      'Scheduled Date In RO',
      'Date Extra',
      'Size In RO',
      'RO Type',
      'Position',
      'Material',
      'Rate In RO',
      'Size',
      'Rate'
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4a86e8');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('#ffffff');

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }

  return sheet;
}

// ==================== TEST FUNCTIONS ====================
// Use these to test your setup

function testGetNextRONumber() {
  const result = getNextRONumber();
  Logger.log(result.getContent());
}

function testSaveReleaseOrder() {
  const testData = {
    roNumber: '25-26/0001',
    date: '30/12/2025',
    publicationName: 'Test Publication',
    edition: 'Test Edition',
    client: 'Test Client',
    scheduledDate: '31/12/2025 MUST',
    size: '26CM (W) X 33CM (H) = 858 SQCM',
    position: 'BEST',
    material: 'Ad. Enclosed',
    rate: '2000 - less 20% + GST as applicable'
  };

  const result = saveReleaseOrder(testData);
  Logger.log(result.getContent());
}