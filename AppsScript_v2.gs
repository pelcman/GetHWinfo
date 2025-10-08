/**
 * Hardware Information Sync to Google Spreadsheet V2
 * Updated to match actual CSV structure from systeminfo.ps1
 */

// Configuration
const COMPUTER_NAME_COLUMN = 'ComputerName';

/**
 * Handle POST requests from PowerShell script
 */
function doPost(e) {
  try {
    Logger.log('=== POST Request Received ===');
    Logger.log('Content type: ' + e.postData.type);
    Logger.log('Content length: ' + e.postData.length);

    // Parse incoming JSON data
    const requestData = JSON.parse(e.postData.contents);
    Logger.log('Parsed JSON successfully');

    const records = requestData.data;
    Logger.log('Records count: ' + (records ? records.length : 0));

    if (!records || records.length === 0) {
      Logger.log('ERROR: No records received');
      return createResponse(false, 'No data received');
    }

    // Log first record structure
    Logger.log('First record keys: ' + Object.keys(records[0]).join(', '));

    // Get the active spreadsheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Spreadsheet ID: ' + spreadsheet.getId());
    Logger.log('Spreadsheet name: ' + spreadsheet.getName());

    // Get or create the first sheet
    let sheet = spreadsheet.getSheets()[0];
    if (!sheet) {
      Logger.log('No sheets found, creating new sheet');
      sheet = spreadsheet.insertSheet('Hardware Info');
    }

    Logger.log('Using sheet: ' + sheet.getName());

    // Process the data
    Logger.log('=== Starting Update ===');
    const result = updateSpreadsheet(sheet, records);
    Logger.log('=== Update Completed ===');
    Logger.log('Result: ' + JSON.stringify(result));

    return createResponse(true, 'Data synchronized successfully', {
      updated: result.updated,
      added: result.added,
      total: result.total,
      spreadsheetUrl: spreadsheet.getUrl()
    });

  } catch (error) {
    Logger.log('ERROR in doPost: ' + error.toString());
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
      message: 'Hardware Info Sync Web App V2 is running',
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
  Logger.log('Headers from data: ' + headers.join(', '));

  // Get existing data
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  Logger.log('Sheet status - Last row: ' + lastRow + ', Last col: ' + lastCol);

  let existingHeaders = [];
  let existingData = [];
  let computerNameColIndex = -1;

  if (lastRow > 0 && lastCol > 0) {
    // Read existing headers
    existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log('Existing headers: ' + existingHeaders.join(', '));

    // Find ComputerName column
    computerNameColIndex = existingHeaders.indexOf(COMPUTER_NAME_COLUMN);
    Logger.log('ComputerName column index: ' + computerNameColIndex);

    // Read existing data (if any)
    if (lastRow > 1) {
      existingData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      Logger.log('Existing data rows: ' + existingData.length);
    }
  } else {
    Logger.log('Sheet is empty - creating new headers');
    // Write headers if sheet is empty
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    existingHeaders = headers;
    computerNameColIndex = headers.indexOf(COMPUTER_NAME_COLUMN);
    Logger.log('Created headers - ComputerName index: ' + computerNameColIndex);

    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    Logger.log('Formatted header row');
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
    Logger.log('Found ' + existingRecordsMap.size + ' existing computer records');
  }

  let updatedCount = 0;
  let addedCount = 0;

  // Process each record
  for (const record of records) {
    const computerName = record[COMPUTER_NAME_COLUMN];
    Logger.log('Processing computer: ' + computerName);

    const rowData = headers.map(header => record[header] || '');

    if (existingRecordsMap.has(computerName)) {
      // Update existing row
      const rowIndex = existingRecordsMap.get(computerName);
      Logger.log('Updating existing row ' + rowIndex + ' for ' + computerName);
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
      updatedCount++;
    } else {
      // Append new row
      Logger.log('Adding new row for ' + computerName);
      sheet.appendRow(rowData);
      addedCount++;
    }
  }

  Logger.log('Updated: ' + updatedCount + ', Added: ' + addedCount);

  // Sort by ComputerName column
  if (computerNameColIndex >= 0 && sheet.getLastRow() > 1) {
    Logger.log('Sorting by ComputerName');
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    dataRange.sort(computerNameColIndex + 1); // +1 for 1-based column index
  }

  // Auto-resize columns
  Logger.log('Auto-resizing columns');
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  const finalTotal = sheet.getLastRow() - 1; // Exclude header
  Logger.log('Final total rows: ' + finalTotal);

  return {
    updated: updatedCount,
    added: addedCount,
    total: finalTotal
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
 * Test function with actual CSV structure
 */
function testUpdateSpreadsheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  // Match actual CSV structure
  const testData = [
    {
      "ComputerName": "TEST-PC-001",
      "Users": "Administrator\nTestUser",
      "Manufacturer": "Test Manufacturer",
      "Motherboard": "Test Board",
      "CPU": "Test CPU",
      "Cores": "8",
      "Processors": "16",
      "GPU": "Test GPU",
      "RAM": "64",
      "RAM Details": "DIMM 0: 16GB\nDIMM 1: 16GB",
      "Disk": "Test SSD 1TB\nTest HDD 2TB",
      "IP": "192.168.1.100",
      "MAC": "AA-BB-CC-DD-EE-FF",
      "NetworkAdapter": "Test Adapter",
      "OSName": "Windows 11 Pro",
      "OSVersion": "10.0.26100",
      "LastUpdated": "2025-10-07 19:00:00"
    }
  ];

  const result = updateSpreadsheet(sheet, testData);
  Logger.log('Test result: ' + JSON.stringify(result));
}
