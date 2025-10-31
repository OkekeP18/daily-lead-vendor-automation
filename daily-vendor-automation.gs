/**
 * DAILY VENDOR REPORT AUTOMATION
 * 
 * Automates the ingestion of CSV vendor reports via Gmail,
 * processes them into Google Sheets, and maintains weekly summaries.
 * 
 * Saves ~28 hours/month of manual data processing
 * Processes 9-17 reports daily across 20+ vendor integrations
 * 
 * Setup Instructions:
 * 1. Replace YOUR_SHEET_ID_HERE with actual destination sheet IDs
 * 2. Replace YOUR_REPORTS_DRIVE_FOLDER_ID with your Google Drive folder ID
 * 3. Run setupDailyTrigger() once to create the daily 8 AM automation
 * 4. Run setupWeeklyTrigger() once to create the Saturday 10 AM summary
 */

///////////////////////////////////////////////////////////////////
// CONFIGURATION - Replace these values with your own
///////////////////////////////////////////////////////////////////

const VENDOR_CONFIG = {
  'Vendor 1 to the LV': {
    sheetId: 'YOUR_SHEET_ID_HERE',
    tabName: 'Vendor 1',
    processFunction: 'processVendor1'
  },
  'Vendor 2 to the LV': {
    sheetId: 'YOUR_SHEET_ID_HERE',
    tabName: 'Vendor 2',
    processFunction: 'processVendor2'
  },
  // Add remaining 18+ vendors following same pattern
};

const REPORTS_DRIVE_FOLDER_ID = 'YOUR_REPORTS_DRIVE_FOLDER_ID';

///////////////////////////////////////////////////////////////////
// MAIN AUTOMATION - Runs daily at 8 AM (Tuesday-Saturday only)
///////////////////////////////////////////////////////////////////

function automateEmailReports() {
  const today = new Date().getDay(); // 0=Sunday, 1=Monday, 2=Tuesday, etc.
  
  // Skip if Sunday (0) or Monday (1)
  if (today === 0 || today === 1) {
    console.log('Skipping automation - Sunday or Monday');
    return;
  }
  
  console.log('Starting email report automation...');
  
  // Get today's date range for Gmail search
  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);
  
  // Search for scheduled report emails with CSV attachments
  const searchQuery = `in:updates subject:"Scheduled Report:" has:attachment -subject:"weekly" after:${formatDateForGmail(todayStart)}`;
  const threads = GmailApp.search(searchQuery);
  
  console.log(`Found ${threads.length} email threads`);
  
  let processedReports = [];
  let skippedCount = 0;
  
  // Process each email thread
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const attachments = message.getAttachments();
      
      attachments.forEach(attachment => {
        if (attachment.getName().endsWith('.csv')) {
          const result = processReportAttachment(attachment, today);
          if (result) {
            processedReports.push(result);
            console.log('Waiting 10 seconds before next vendor...');
            Utilities.sleep(10000);
          } else {
            skippedCount++;
          }
        }
      });
    });
  });
  
  // Run cleanup functions once at the end if any reports were processed
  if (processedReports.length > 0) {
    console.log('Running final cleanup functions...');
    clearUnwantedDispositions();
    formatAllDates();
    console.log('Cleanup functions completed');
  }
  
  console.log(`Automation complete. Processed: ${processedReports.length}, Skipped: ${skippedCount}`);
}

///////////////////////////////////////////////////////////////////
// REPORT PROCESSING
///////////////////////////////////////////////////////////////////

function processReportAttachment(attachment, dayOfWeek) {
  const fileName = attachment.getName();
  const vendorName = fileName.replace('.csv', '');
  
  console.log(`Processing: ${fileName}`);
  
  try {
    // Parse CSV content
    const csvContent = attachment.getDataAsString();
    const csvData = Utilities.parseCsv(csvContent);
    
    // Skip if CSV only has headers or is empty
    if (csvData.length <= 1) {
      console.log(`Skipping ${fileName} - empty or headers only`);
      return false;
    }
    
    // Upload to Drive (only if CSV has actual data)
    console.log(`${fileName} has data - uploading to Drive...`);
    uploadCSVToDrive(attachment, extractDateFromCSV(csvData[1][0]));
    
    // Get the macro function for this vendor
    const macroFunction = getVendorMacroFunction(vendorName, dayOfWeek);
    if (!macroFunction) {
      console.log(`Warning: No macro function found for vendor: ${vendorName}`);
      return false;
    }
    
    // Trigger the macro to prepare destination sheet
    console.log(`Triggering macro for: ${vendorName}`);
    macroFunction();
    
    // Paste CSV data into Source sheet
    console.log(`Pasting data into Source sheet: ${VENDOR_CONFIG[vendorName]?.tabName}`);
    pasteDataToSourceSheet(vendorName, csvData);
    
    console.log(`Successfully processed: ${fileName}`);
    return vendorName;
    
  } catch (error) {
    console.log(`Error processing ${fileName}: ${error.toString()}`);
    return false;
  }
}

///////////////////////////////////////////////////////////////////
// SHEET OPERATIONS
///////////////////////////////////////////////////////////////////

/**
 * Processes vendor data for regular days (Tuesday-Saturday)
 * Creates 2 new rows for daily report data
 */
function processVendorData(sheetId) {
  var destinationSpreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = destinationSpreadsheet.getActiveSheet();
  
  var lastRow = sheet.getLastRow();
  var formulaRange = sheet.getRange(lastRow, 1, 1, 19); // Columns A to S
  
  var destinationRange = formulaRange.offset(0, 0, 2);
  formulaRange.autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  // Convert formulas to values
  formulaRange.copyTo(formulaRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  // Force sync to prevent duplicate rows
  SpreadsheetApp.flush();
  Utilities.sleep(3000);
}

/**
 * Processes vendor data for Monday reports only
 * Skips a row to separate weekly data blocks
 */
function processVendorData2(sheetId) {
  var destinationSpreadsheet = SpreadsheetApp.openById(sheetId);
  var sheet = destinationSpreadsheet.getActiveSheet();
  
  var lastRow = sheet.getLastRow();
  
  // Use second-to-last row (actual data, not sum row)
  var formulaRange = sheet.getRange(lastRow - 1, 1, 1, 19);
  
  // Copy to skip one row from current last row
  var destinationRange = sheet.getRange(lastRow + 1, 1, 1, 19);
  formulaRange.copyTo(destinationRange);
  
  // Convert formulas to values
  formulaRange.copyTo(formulaRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  // Force sync
  SpreadsheetApp.flush();
  Utilities.sleep(3000);
}

/**
 * Clears unwanted disposition talk times from specific vendor tabs
 * Handles: "No Agents - Disconnect", "Duplicate", "Caller Disconnected", "Abandon", "No Disposition"
 */
function clearUnwantedDispositions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetTabs = ['Vendor 1', 'Vendor 2', 'Vendor 3']; // Add your vendor tabs
  
  var clearTriggers = [
    'No Agents - Disconnect',
    'Duplicate', 
    'Caller Disconnected',
    'Caller Disconnected - At Closed Message',
    'Abandon',
    'No Disposition'
  ];
  
  for (var t = 0; t < targetTabs.length; t++) {
    var sheet = ss.getSheetByName(targetTabs[t]);
    
    if (!sheet) {
      console.log("Sheet '" + targetTabs[t] + "' not found");
      continue;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    
    var columnG = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
    
    for (var i = 0; i < columnG.length; i++) {
      var cellValue = columnG[i][0].toString().trim();
      if (clearTriggers.includes(cellValue)) {
        sheet.getRange(i + 2, 9).clearContent();
      }
    }
  }
}

/**
 * Formats all dates in Column A across all sheets to MM/DD/YYYY
 */
function formatAllDates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 1) continue;
    
    var columnA = sheet.getRange(1, 1, lastRow, 1);
    columnA.setNumberFormat('MM/DD/YYYY');
  }
}

/**
 * Pastes CSV data into the Source sheet
 */
function pasteDataToSourceSheet(csvAttachmentName, csvData) {
  const tabName = VENDOR_CONFIG[csvAttachmentName]?.tabName || csvAttachmentName;
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet();
  const vendorTab = sourceSheet.getSheetByName(tabName);
  
  if (!vendorTab) {
    console.log(`Tab not found: ${tabName}`);
    return;
  }
  
  // Clear existing data (keep headers)
  const lastRow = vendorTab.getLastRow();
  if (lastRow > 1) {
    vendorTab.getRange(2, 1, lastRow - 1, vendorTab.getLastColumn()).clear();
  }
  
  // Paste new data (skip CSV headers)
  if (csvData.length > 1) {
    const dataOnly = csvData.slice(1);
    const range = vendorTab.getRange(2, 1, dataOnly.length, dataOnly[0].length);
    range.setValues(dataOnly);
  }
}

///////////////////////////////////////////////////////////////////
// SMART VENDOR MAPPING
///////////////////////////////////////////////////////////////////

/**
 * Intelligently routes to skip-row or regular function based on first data of week
 */
function getVendorMacroFunction(vendorName, dayOfWeek) {
  const shouldSkipRow = isFirstDataOfWeek(vendorName);
  
  const vendorSheetId = VENDOR_CONFIG[vendorName]?.sheetId;
  if (!vendorSheetId) {
    console.log(`No sheet ID configured for: ${vendorName}`);
    return null;
  }
  
  console.log(`${vendorName}: ${shouldSkipRow ? 'First data - SKIP ROW' : 'Regular data'}`);
  
  if (shouldSkipRow) {
    return () => processVendorData2(vendorSheetId);
  } else {
    return () => processVendorData(vendorSheetId);
  }
}

/**
 * Checks if this is the first data of the week for a vendor
 */
function isFirstDataOfWeek(csvAttachmentName) {
  try {
    const vendorSheetId = VENDOR_CONFIG[csvAttachmentName]?.sheetId;
    if (!vendorSheetId) {
      console.log(`Warning: No sheet ID for ${csvAttachmentName}, defaulting to skip row`);
      return true;
    }
    
    const destinationSpreadsheet = SpreadsheetApp.openById(vendorSheetId);
    const sheet = destinationSpreadsheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    // Check columns A-Q for actual data
    const dataRange = sheet.getRange(lastRow, 1, 1, 17);
    const dataValues = dataRange.getValues()[0];
    
    const hasActualData = dataValues.some(value => 
      value !== '' && value !== null && value !== undefined
    );
    
    const isFirstData = !hasActualData;
    console.log(`${csvAttachmentName}: First data of week = ${isFirstData}`);
    
    return isFirstData;
    
  } catch (error) {
    console.log(`Error checking first data: ${error.toString()}`);
    return true;
  }
}

///////////////////////////////////////////////////////////////////
// DRIVE OPERATIONS
///////////////////////////////////////////////////////////////////

/**
 * Uploads CSV files to organized Google Drive folder structure
 * Structure: Reports > [Month Year] > Daily > Daily [DD/MM/YY]
 */
function uploadCSVToDrive(attachment, csvDate) {
  try {
    console.log(`Uploading ${attachment.getName()} to Google Drive...`);
    
    const targetFolder = getOrCreateDailyFolder(csvDate);
    if (!targetFolder) {
      console.log(`Failed to create folder structure for ${attachment.getName()}`);
      return false;
    }
    
    targetFolder.createFile(attachment);
    console.log(`Uploaded: ${attachment.getName()}`);
    return true;
    
  } catch (error) {
    console.log(`Error uploading to Drive: ${error.toString()}`);
    return false;
  }
}

/**
 * Creates or retrieves folder structure
 * Format: Reports > [Month Year] > Daily > Daily [DD/MM/YY]
 */
function getOrCreateDailyFolder(csvDate) {
  try {
    const rootFolder = DriveApp.getFolderById(REPORTS_DRIVE_FOLDER_ID);
    
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                       'July', 'August', 'September', 'October', 'November', 'December'];
    const monthName = monthNames[csvDate.getMonth()];
    const year = csvDate.getFullYear();
    const monthFolderName = `${monthName} ${year}`;
    
    const monthFolder = getOrCreateFolder(rootFolder, monthFolderName);
    const dailyFolder = getOrCreateFolder(monthFolder, 'Daily');
    
    const day = csvDate.getDate().toString().padStart(2, '0');
    const month = (csvDate.getMonth() + 1).toString().padStart(2, '0');
    const shortYear = csvDate.getFullYear().toString().slice(-2);
    const dailyDateFolderName = `Daily ${day}/${month}/${shortYear}`;
    
    return getOrCreateFolder(dailyFolder, dailyDateFolderName);
    
  } catch (error) {
    console.log(`Error creating folder structure: ${error.toString()}`);
    return null;
  }
}

function getOrCreateFolder(parentFolder, folderName) {
  const existingFolders = parentFolder.getFoldersByName(folderName);
  return existingFolders.hasNext() ? existingFolders.next() : parentFolder.createFolder(folderName);
}

///////////////////////////////////////////////////////////////////
// WEEKLY SUMMARY
///////////////////////////////////////////////////////////////////

/**
 * Runs Saturday at 10 AM on ALL destination sheets
 * Creates SUM formulas for the week's data
 */
function weeklySum() {
  console.log('Starting weekly sum for all vendor sheets...');
  
  let successCount = 0;
  let errorCount = 0;
  
  for (const [vendorName, config] of Object.entries(VENDOR_CONFIG)) {
    try {
      console.log(`Processing ${config.tabName}...`);
      processWeeklySumForSheet(config.sheetId, config.tabName);
      successCount++;
    } catch (error) {
      console.log(`Error processing ${config.tabName}: ${error.toString()}`);
      errorCount++;
    }
  }
  
  console.log(`Weekly sum completed. Success: ${successCount}, Errors: ${errorCount}`);
}

function processWeeklySumForSheet(sheetId, sheetName) {
  const destinationSpreadsheet = SpreadsheetApp.openById(sheetId);
  const sheet = destinationSpreadsheet.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  console.log(`${sheetName}: Last row is ${lastRow}`);
  
  let emptyRow = lastRow + 1;
  
  // Check if sums already exist
  const existingSumR = sheet.getRange(emptyRow, 18).getValue();
  const existingSumS = sheet.getRange(emptyRow, 19).getValue();
  
  if (existingSumR !== "" || existingSumS !== "") {
    console.log(`${sheetName}: Sums already exist. Skipping.`);
    return;
  }
  
  // Find rows with dates (up to 5)
  const rowsToSum = findRowsWithDates(sheet, lastRow, 5);
  
  if (rowsToSum.length === 0) {
    console.log(`${sheetName}: No rows with dates found.`);
    return;
  }
  
  const firstRow = rowsToSum[0];
  const lastRowToSum = rowsToSum[rowsToSum.length - 1];
  
  const formulaR = `=SUM(R${firstRow}:R${lastRowToSum})`;
  const formulaS = `=SUM(S${firstRow}:S${lastRowToSum})`;
  
  sheet.getRange(emptyRow, 18).setFormula(formulaR);
  sheet.getRange(emptyRow, 19).setFormula(formulaS);
  sheet.getRange(emptyRow, 18, 1, 2).setBackground('#00ff00');
  
  console.log(`${sheetName}: SUM formulas added to row ${emptyRow}`);
}

function findRowsWithDates(sheet, startRow, maxRows) {
  const rowsWithDates = [];
  let currentRow = startRow;
  
  while (rowsWithDates.length < maxRows && currentRow > 1) {
    const cellValue = sheet.getRange(currentRow, 1).getValue();
    
    if (cellValue instanceof Date || (typeof cellValue === 'string' && cellValue.trim() !== '')) {
      if (cellValue instanceof Date || isValidDateString(cellValue)) {
        rowsWithDates.unshift(currentRow);
      } else {
        break;
      }
    } else {
      break;
    }
    currentRow--;
  }
  
  return rowsWithDates;
}

function isValidDateString(str) {
  if (typeof str !== 'string') return false;
  const datePatterns = [
    /^\d{1,2}\/\d{1,2}\/\d{4}$/,
    /^\d{4}-\d{1,2}-\d{1,2}$/,
    /^\d{1,2}-\d{1,2}-\d{4}$/
  ];
  return datePatterns.some(pattern => pattern.test(str.trim()));
}

///////////////////////////////////////////////////////////////////
// TRIGGER SETUP
///////////////////////////////////////////////////////////////////

function setupDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'automateEmailReports') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('automateEmailReports')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  console.log('Daily trigger set for 8 AM (Tue-Sat only)');
}

function setupWeeklyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'weeklySum') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('weeklySum')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(10)
    .create();
    
  console.log('Weekly trigger set for Saturday 10 AM');
}

///////////////////////////////////////////////////////////////////
// HELPER FUNCTIONS
///////////////////////////////////////////////////////////////////

function formatDateForGmail(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}/${month}/${day}`;
}

function extractDateFromCSV(dateValue) {
  try {
    let date;
    
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue);
    } else if (typeof dateValue === 'number') {
      date = new Date((dateValue - 25569) * 86400 * 1000);
    } else {
      date = new Date(dateValue);
    }
    
    if (isNaN(date.getTime())) {
      console.log(`Invalid date: ${dateValue}`);
      return null;
    }
    
    return date;
  } catch (error) {
    console.log(`Error extracting date: ${error.toString()}`);
    return null;
  }
}

///////////////////////////////////////////////////////////////////
// TEST FUNCTIONS
///////////////////////////////////////////////////////////////////

function testAutomation() {
  console.log('Running test automation...');
  automateEmailReports();
}

function testWeeklySum() {
  console.log('Running test weekly sum...');
  weeklySum();
}
