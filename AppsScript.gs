/**
 * Hardware Information Sync to Google Spreadsheet
 * This script receives hardware info from PowerShell and updates the spreadsheet
 *
 * Setup:
 * 1. Open your Google Spreadsheet
 * 2. Go to Extensions > Apps Script
 * 3. Replace all code with this file
 * 4. Deploy as Web App (Deploy > New deployment)
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL and use it in sync_to_spreadsheet.ps1
 */

// Configuration
const SHEET_NAME = 'Sheet1'; // Change this to your sheet name if needed
const COMPUTER_NAME_COLUMN = 'ComputerName'; // Column to use for identifying unique records

/**
 * Handle POST requests from PowerShell script
 */
function doPost(e) {
  try {
    // Log incoming request
    Logger.log('Received POST request');
    Logger.log('Content type: ' + e.postData.type);
    Logger.log('Content length: ' + e.postData.length);

    // Parse incoming JSON data
    const requestData = JSON.parse(e.postData.contents);
    Logger.log('Parsed JSON successfully');

    const records = requestData.data;
    Logger.log('Records count: ' + (records ? records.length : 0));

    if (!records || records.length === 0) {
      Logger.log('No records received');
      return createResponse(false, 'No data received');
    }

    // Get the active spreadsheet and sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Spreadsheet ID: ' + spreadsheet.getId());

    let sheet = spreadsheet.getSheetByName(SHEET_NAME);

    // Create sheet if it doesn't exist
    if (!sheet) {
      Logger.log('Creating new sheet: ' + SHEET_NAME);
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    // Process the data
    Logger.log('Starting to update spreadsheet');
    const result = updateSpreadsheet(sheet, records);
    Logger.log('Update completed successfully');

    return createResponse(true, 'Data synchronized successfully', {
      updated: result.updated,
      added: result.added,
      total: result.total,
      spreadsheetUrl: spreadsheet.getUrl()
    });

  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return createResponse(false, 'Error: ' + error.toString());
  }
}

/**
 * Handle GET requests (for testing)
 */
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({
      status: 'success',
      message: 'Hardware Info Sync Web App is running',
      timestamp: new Date().toISOString()
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Update spreadsheet with hardware info records
 */
function updateSpreadsheet(sheet, records) {
  // Get all headers from the first record
  const headers = Object.keys(records[0]);

  // Get existing data
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  let existingHeaders = [];
  let existingData = [];
  let computerNameColIndex = -1;

  if (lastRow > 0) {
    // Read existing headers
    existingHeaders = sheet.getRange(1, 1, 1, lastCol > 0 ? lastCol : headers.length).getValues()[0];
    computerNameColIndex = existingHeaders.indexOf(COMPUTER_NAME_COLUMN);

    // Read existing data (if any)
    if (lastRow > 1) {
      existingData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    }
  } else {
    // Write headers if sheet is empty
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    existingHeaders = headers;
    computerNameColIndex = headers.indexOf(COMPUTER_NAME_COLUMN);

    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
  }

  // Create a map of existing records by ComputerName
  const existingRecordsMap = new Map();
  if (computerNameColIndex >= 0) {
    for (let i = 0; i < existingData.length; i++) {
      const computerName = existingData[i][computerNameColIndex];
      if (computerName) {
        existingRecordsMap.set(computerName, i + 2); // +2 for 1-based index and header row
      }
    }
  }

  let updatedCount = 0;
  let addedCount = 0;

  // Process each record
  for (const record of records) {
    const computerName = record[COMPUTER_NAME_COLUMN];
    const rowData = headers.map(header => record[header] || '');

    if (existingRecordsMap.has(computerName)) {
      // Update existing row
      const rowIndex = existingRecordsMap.get(computerName);
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
      updatedCount++;
    } else {
      // Append new row
      sheet.appendRow(rowData);
      addedCount++;
    }
  }

  // Sort by ComputerName column
  if (computerNameColIndex >= 0 && sheet.getLastRow() > 1) {
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    dataRange.sort(computerNameColIndex + 1); // +1 for 1-based column index
  }

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  return {
    updated: updatedCount,
    added: addedCount,
    total: sheet.getLastRow() - 1 // Exclude header
  };
}

/**
 * Create JSON response
 */
function createResponse(success, message, data = {}) {
  const response = {
    status: success ? 'success' : 'error',
    message: message,
    timestamp: new Date().toISOString(),
    ...data
  };

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Test function (run this from Apps Script editor to test)
 */
function testUpdateSpreadsheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const testData = [
    {
      LastUpdated: '2025-10-07 12:00:00',
      ComputerName: 'TEST-PC-001',
      AllRegisteredUsers: 'Administrator\nUser1',
      PrimaryUser: 'User1',
      CurrentUser: 'User1',
      IPAddress: '192.168.1.100',
      MACAddress: 'AA-BB-CC-DD-EE-FF',
      Manufacturer: 'Test Manufacturer',
      Model: 'Test Model',
      SerialNumber: 'TEST123',
      OSName: 'Windows 11 Pro',
      OSVersion: '10.0.22000',
      OSArchitecture: '64-bit',
      CPU: 'Test CPU',
      CPUCores: '8',
      CPULogicalProcessors: '16',
      TotalMemoryGB: '32',
      RAMSlots: 'Slot1:16GB\nSlot2:16GB',
      MemoryDetails: 'DIMM1: 16GB\nDIMM2: 16GB',
      GPU: 'Test GPU',
      Disk: 'Test Disk 1TB',
      MotherboardManufacturer: 'Test Board Mfg',
      MotherboardProduct: 'Test Board',
      NetworkAdapter: 'Test Adapter'
    }
  ];

  const result = updateSpreadsheet(sheet, testData);
  Logger.log('Test result: ' + JSON.stringify(result));
}
