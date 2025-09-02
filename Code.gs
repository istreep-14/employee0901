/**
 * Bar Employee CRM - Google Apps Script Backend (Fixed Date Handling)
 * 
 * Setup Instructions:
 * 1. Create a new Google Apps Script project (script.google.com)
 * 2. Replace Code.gs content with this script
 * 3. Add the HTML file as 'index.html' 
 * 4. Deploy as Web App with execute as "Me" and access "Anyone"
 * 5. Copy the web app URL and test the connection
 */

// Configuration
const CRM_SHEET_NAME = 'Bar_Employee_CRM_Data';
const POSITIONS_SHEET_NAME = 'Position_Config';
const PHOTOS_FOLDER_NAME = 'Bar_Employee_Photos';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Headers in order
const HEADERS = [
  'Emp Id', 'First Name', 'Last Name', 'Phone', 'Email', 
  'Position', 'Status', 'Note', 'Photo URL', 'Created Date', 'Last Modified',
  'Is Manager', 'Is Assistant Manager', 'Is Me'
];

// Default positions with icons (matching HTML)
const DEFAULT_POSITIONS = [
  { name: 'Bartender', icon: '🍸' },
  { name: 'Server', icon: '🍽️' },
  { name: 'Manager', icon: '👔' },
  { name: 'Host', icon: '🎯' },
  { name: 'Kitchen Staff', icon: '👨‍🍳' },
  { name: 'Security', icon: '🛡️' },
  { name: 'Assistant Manager', icon: '🎖️' }
];

/**
 * Utility function to safely convert string dates to Date objects
 */
function safeParseDate(dateInput) {
  if (!dateInput) return null;
  
  try {
    // If it's already a Date object, return it
    if (dateInput instanceof Date) {
      return isNaN(dateInput.getTime()) ? null : dateInput;
    }
    
    // If it's a string, try to parse it
    if (typeof dateInput === 'string') {
      const parsed = new Date(dateInput);
      return isNaN(parsed.getTime()) ? null : parsed;
    }
    
    // If it's a number (timestamp), convert it
    if (typeof dateInput === 'number') {
      const parsed = new Date(dateInput);
      return isNaN(parsed.getTime()) ? null : parsed;
    }
    
    return null;
  } catch (error) {
    Logger.log('Date parsing error: ' + error.toString() + ' for input: ' + dateInput);
    return null;
  }
}

/**
 * Utility function to format date for display
 */
function formatDateForDisplay(dateInput) {
  const date = safeParseDate(dateInput);
  return date ? date.toISOString() : '';
}

/**
 * Serves the main web app HTML
 */
function doGet(e) {
  try {
    const html = HtmlService.createTemplateFromFile('index');
    const htmlOutput = html.evaluate()
      .setTitle('Bar Employee CRM')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    
    return htmlOutput;
  } catch (error) {
    Logger.log('Error serving HTML: ' + error.toString());
    
    const errorHtml = HtmlService.createHtmlOutput(`
      <html>
        <body style="font-family: Arial, sans-serif; padding: 20px; text-align: center;">
          <h2>🚫 CRM System Error</h2>
          <p>There was an error loading the Bar Employee CRM.</p>
          <p><strong>Error:</strong> ${error.toString()}</p>
          <button onclick="location.reload()">🔄 Reload Page</button>
        </body>
      </html>
    `).setTitle('CRM Error');
    
    return errorHtml;
  }
}

/**
 * Include CSS and JS files in HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get or create the CRM data sheet
 */
function getCRMSheet() {
  try {
    const props = PropertiesService.getScriptProperties();
    let sheetId = props.getProperty('CRM_SHEET_ID');
    let sheet = null;
    
    if (sheetId) {
      try {
        const spreadsheet = SpreadsheetApp.openById(sheetId);
        sheet = spreadsheet.getSheetByName(CRM_SHEET_NAME);
        
        if (sheet && !validateSheetHeaders(sheet)) {
          Logger.log('Sheet headers invalid, attempting migration...');
          if (!tryMigrateSheet(sheet)) {
            Logger.log('Migration not possible, reinitializing sheet...');
            initializeSheet(sheet);
          }
        }
      } catch (e) {
        Logger.log('Existing sheet not accessible: ' + e.toString());
        sheetId = null;
      }
    }
    
    if (!sheet) {
      const spreadsheet = SpreadsheetApp.create('Bar Employee CRM Data - ' + new Date().getFullYear());
      sheet = spreadsheet.getActiveSheet();
      sheet.setName(CRM_SHEET_NAME);
      
      sheetId = spreadsheet.getId();
      props.setProperty('CRM_SHEET_ID', sheetId);
      
      initializeSheet(sheet);
      initializePositionsSheet(spreadsheet);
      
      Logger.log('Created new CRM sheet: ' + sheetId);
    }
    
    return sheet;
  } catch (error) {
    throw new Error('Failed to access or create CRM data sheet: ' + error.toString());
  }
}

/**
 * Get or create positions configuration sheet
 */
function getPositionsSheet() {
  try {
    const props = PropertiesService.getScriptProperties();
    const sheetId = props.getProperty('CRM_SHEET_ID');
    
    if (!sheetId) {
      throw new Error('No CRM sheet ID found');
    }
    
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    let posSheet = spreadsheet.getSheetByName(POSITIONS_SHEET_NAME);
    
    if (!posSheet) {
      posSheet = initializePositionsSheet(spreadsheet);
    }
    
    return posSheet;
  } catch (error) {
    Logger.log('Error getting positions sheet: ' + error.toString());
    return null;
  }
}

/**
 * Initialize positions configuration sheet
 */
function initializePositionsSheet(spreadsheet) {
  try {
    const posSheet = spreadsheet.insertSheet(POSITIONS_SHEET_NAME);
    
    posSheet.getRange(1, 1, 1, 2).setValues([['Position Name', 'Icon']]);
    
    const positionData = DEFAULT_POSITIONS.map(pos => [pos.name, pos.icon]);
    if (positionData.length > 0) {
      posSheet.getRange(2, 1, positionData.length, 2).setValues(positionData);
    }
    
    const headerRange = posSheet.getRange(1, 1, 1, 2);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    
    posSheet.setColumnWidth(1, 200);
    posSheet.setColumnWidth(2, 80);
    posSheet.setFrozenRows(1);
    
    Logger.log('Initialized positions sheet');
    return posSheet;
  } catch (error) {
    Logger.log('Error initializing positions sheet: ' + error.toString());
    return null;
  }
}

/**
 * Validate sheet headers
 */
function validateSheetHeaders(sheet) {
  try {
    const existingHeaders = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).getValues()[0];
    return HEADERS.every((header, index) => existingHeaders[index] === header);
  } catch (error) {
    return false;
  }
}

/**
 * Initialize the main data sheet with headers
 */
function initializeSheet(sheet) {
  try {
    sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS]);
    
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    
    sheet.setColumnWidth(1, 100);  // Emp ID
    sheet.setColumnWidth(2, 120);  // First Name
    sheet.setColumnWidth(3, 120);  // Last Name
    sheet.setColumnWidth(4, 140);  // Phone
    sheet.setColumnWidth(5, 180);  // Email
    sheet.setColumnWidth(6, 200);  // Position
    sheet.setColumnWidth(7, 80);   // Status
    sheet.setColumnWidth(8, 250);  // Note
    sheet.setColumnWidth(9, 200);  // Photo URL
    sheet.setColumnWidth(10, 120); // Created Date
    sheet.setColumnWidth(11, 120); // Last Modified
    sheet.setColumnWidth(12, 100); // Is Manager
    sheet.setColumnWidth(13, 150); // Is Assistant Manager
    sheet.setColumnWidth(14, 80);  // Is Me
    
    sheet.setFrozenRows(1);
    
    const statusRange = sheet.getRange(DATA_START_ROW, 7, 1000, 1);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Inactive'])
      .setAllowInvalid(false)
      .setHelpText('Select Active or Inactive')
      .build();
    statusRange.setDataValidation(statusRule);

    // Add checkbox validations for boolean columns
    const boolRange = sheet.getRange(DATA_START_ROW, 12, 1000, 3);
    try { boolRange.insertCheckboxes(); } catch (e) { /* older Apps Script? ignore */ }
    
    Logger.log('Sheet initialized successfully');
    return { success: true };
  } catch (error) {
    Logger.log('Error initializing sheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Attempt to migrate an older sheet schema to the latest headers
 */
function tryMigrateSheet(sheet) {
  try {
    const existingHeadersRow = sheet.getRange(HEADER_ROW, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)).getValues()[0];
    const oldHeaders = [
      'Emp Id', 'First Name', 'Last Name', 'Phone', 'Email',
      'Position', 'Status', 'Note', 'Photo URL', 'Created Date', 'Last Modified'
    ];
    const existingFirst11 = existingHeadersRow.slice(0, oldHeaders.length);
    const isOldSchema = oldHeaders.every((h, i) => existingFirst11[i] === h);
    if (!isOldSchema) {
      return false;
    }
    // Append new headers in place
    const newHeaderValues = [HEADERS];
    sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues(newHeaderValues);
    // Set widths for new columns
    sheet.setColumnWidth(12, 100);
    sheet.setColumnWidth(13, 150);
    sheet.setColumnWidth(14, 80);
    // Add checkbox validations for boolean columns
    const lastRow = sheet.getLastRow();
    const numRows = Math.max(1000, lastRow - HEADER_ROW);
    const boolRange = sheet.getRange(DATA_START_ROW, 12, numRows, 3);
    try { boolRange.insertCheckboxes(); } catch (e) { /* ignore */ }
    Logger.log('Migrated sheet headers to latest schema');
    return true;
  } catch (e) {
    Logger.log('Migration failed: ' + e.toString());
    return false;
  }
}

/**
 * Get system information and verify connection
 */
function getSystemInfo() {
  try {
    const props = PropertiesService.getScriptProperties();
    let sheetId = props.getProperty('CRM_SHEET_ID');
    
    let sheetAccessible = false;
    try {
      if (sheetId) {
        const sheet = SpreadsheetApp.openById(sheetId);
        sheetAccessible = !!sheet;
      }
    } catch (e) {
      Logger.log('Sheet access test failed: ' + e.toString());
    }
    
    if (!sheetAccessible) {
      try {
        getCRMSheet();
        sheetId = props.getProperty('CRM_SHEET_ID');
        sheetAccessible = true;
      } catch (e) {
        Logger.log('Sheet creation failed: ' + e.toString());
      }
    }
    
    const userEmail = Session.getActiveUser().getEmail();
    const scriptTimeZone = Session.getScriptTimeZone();
    
    return {
      success: true,
      info: {
        scriptId: ScriptApp.getScriptId(),
        sheetId: sheetId,
        sheetAccessible: sheetAccessible,
        timezone: scriptTimeZone,
        userEmail: userEmail,
        lastSync: new Date().toISOString(),
        version: '2.0',
        positionsCount: DEFAULT_POSITIONS.length
      }
    };
  } catch (error) {
    Logger.log('Error in getSystemInfo: ' + error.toString());
    return { 
      success: false, 
      error: error.toString(),
      info: {
        scriptId: ScriptApp.getScriptId() || 'unknown',
        sheetAccessible: false,
        version: '2.0'
      }
    };
  }
}

/**
 * Get the Google Sheet URL
 */
function getSheetUrl() {
  try {
    const props = PropertiesService.getScriptProperties();
    const sheetId = props.getProperty('CRM_SHEET_ID');
    
    if (sheetId) {
      return { 
        success: true, 
        url: `https://docs.google.com/spreadsheets/d/${sheetId}/edit`
      };
    } else {
      const sheet = getCRMSheet();
      const newSheetId = props.getProperty('CRM_SHEET_ID');
      return { 
        success: true, 
        url: `https://docs.google.com/spreadsheets/d/${newSheetId}/edit`
      };
    }
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Get all employees from the sheet with enhanced error handling
 */
function getAllEmployees() {
  try {
    Logger.log('getAllEmployees: Starting to load employees...');
    
    const sheet = getCRMSheet();
    const lastRow = sheet.getLastRow();
    
    Logger.log(`getAllEmployees: Sheet has ${lastRow} rows total`);
    
    if (lastRow < DATA_START_ROW) {
      Logger.log('getAllEmployees: No data rows found, returning empty array');
      return { success: true, employees: [], message: 'No employees found - ready to add data!' };
    }
    
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, HEADERS.length);
    const values = dataRange.getValues();
    
    Logger.log(`getAllEmployees: Retrieved ${values.length} data rows`);
    
    const employees = [];
    let processedCount = 0;
    let skippedCount = 0;
    
    values.forEach((row, index) => {
      try {
        if (!row[0] || row[0].toString().trim() === '') {
          skippedCount++;
          return;
        }
        
        const employee = {
          empId: row[0] ? row[0].toString().trim() : '',
          firstName: row[1] ? row[1].toString().trim() : '',
          lastName: row[2] ? row[2].toString().trim() : '',
          phone: row[3] ? row[3].toString().trim() : '',
          email: row[4] ? row[4].toString().trim() : '',
          position: row[5] ? row[5].toString().trim() : '',
          status: row[6] ? row[6].toString().trim() : 'Active',
          note: row[7] ? row[7].toString().trim() : '',
          photoUrl: row[8] ? row[8].toString().trim() : '',
          createdDate: formatDateForDisplay(row[9]),
          lastModified: formatDateForDisplay(row[10]),
          isManager: !!row[11],
          isAssistantManager: !!row[12],
          isMe: !!row[13]
        };
        
        employees.push(employee);
        processedCount++;
      } catch (rowError) {
        Logger.log(`Error processing row ${index + DATA_START_ROW}: ${rowError.toString()}`);
        skippedCount++;
      }
    });
    
    Logger.log(`getAllEmployees: Successfully processed ${processedCount} employees, skipped ${skippedCount} rows`);
    
    return { 
      success: true, 
      employees: employees,
      message: `Loaded ${employees.length} employees from Google Sheets`
    };
    
  } catch (error) {
    Logger.log('Critical error in getAllEmployees: ' + error.toString());
    return { 
      success: false, 
      error: `Failed to load employees: ${error.toString()}`,
      employees: []
    };
  }
}

/**
 * Save all employees to the sheet with enhanced validation and fixed date handling
 */
function saveAllEmployees(employees) {
  try {
    Logger.log(`saveAllEmployees: Starting to save ${employees ? employees.length : 0} employees`);
    
    if (!Array.isArray(employees)) {
      throw new Error('Invalid data format: employees must be an array');
    }
    
    const sheet = getCRMSheet();
    const now = new Date();
    
    // Validate and clean each employee record
    const validEmployees = [];
    const errors = [];
    
    employees.forEach((emp, index) => {
      try {
        if (!emp || typeof emp !== 'object') {
          errors.push(`Employee ${index + 1}: Invalid employee object`);
          return;
        }
        
        if (!emp.empId || emp.empId.toString().trim() === '') {
          errors.push(`Employee ${index + 1}: Employee ID is required`);
          return;
        }
        
        // Parse dates safely
        let createdDate = safeParseDate(emp.createdDate);
        let lastModified = safeParseDate(emp.lastModified);
        
        // If dates are invalid, use defaults
        if (!createdDate) {
          createdDate = now;
        }
        if (!lastModified) {
          lastModified = now;
        }
        
        const cleanEmployee = {
          empId: emp.empId.toString().trim(),
          firstName: (emp.firstName || '').toString().trim(),
          lastName: (emp.lastName || '').toString().trim(),
          phone: (emp.phone || '').toString().trim(),
          email: (emp.email || '').toString().trim(),
          position: (emp.position || '').toString().trim(),
          status: (emp.status || 'Active').toString().trim(),
          note: (emp.note || '').toString().trim(),
          photoUrl: (emp.photoUrl || '').toString().trim(),
          createdDate: createdDate,
          lastModified: lastModified,
          isManager: !!emp.isManager,
          isAssistantManager: !!emp.isAssistantManager,
          isMe: !!emp.isMe
        };
        
        validEmployees.push(cleanEmployee);
      } catch (empError) {
        errors.push(`Employee ${index + 1}: ${empError.toString()}`);
      }
    });
    
    if (errors.length > 0) {
      Logger.log('Validation errors found: ' + errors.join('; '));
      if (validEmployees.length === 0) {
        return { 
          success: false, 
          error: 'No valid employees to save. Errors: ' + errors.join('; ')
        };
      }
    }
    
    // Check for duplicate Employee IDs
    const empIds = new Set();
    const duplicates = [];
    validEmployees.forEach(emp => {
      if (empIds.has(emp.empId)) {
        duplicates.push(emp.empId);
      } else {
        empIds.add(emp.empId);
      }
    });
    
    if (duplicates.length > 0) {
      return {
        success: false,
        error: `Duplicate Employee IDs found: ${duplicates.join(', ')}`
      };
    }
    
    // Clear existing data (keep headers)
    try {
      const lastRow = sheet.getLastRow();
      if (lastRow >= DATA_START_ROW) {
        const clearRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW + 1, HEADERS.length);
        clearRange.clear();
        Logger.log(`Cleared ${lastRow - HEADER_ROW + 1} existing rows`);
      }
    } catch (clearError) {
      Logger.log('Error clearing existing data: ' + clearError.toString());
    }
    
    if (validEmployees.length > 0) {
      // Prepare data rows with proper Date objects
      const dataRows = validEmployees.map(emp => [
        emp.empId,
        emp.firstName,
        emp.lastName,
        emp.phone,
        emp.email,
        emp.position,
        emp.status,
        emp.note,
        emp.photoUrl,
        emp.createdDate,  // This is already a Date object
        emp.lastModified, // This is already a Date object
        emp.isManager,
        emp.isAssistantManager,
        emp.isMe
      ]);
      
      try {
        const range = sheet.getRange(DATA_START_ROW, 1, dataRows.length, HEADERS.length);
        range.setValues(dataRows);
        
        // Format date columns
        const createdDateRange = sheet.getRange(DATA_START_ROW, 10, dataRows.length, 1);
        const lastModifiedRange = sheet.getRange(DATA_START_ROW, 11, dataRows.length, 1);
        createdDateRange.setNumberFormat('MM/dd/yyyy hh:mm:ss');
        lastModifiedRange.setNumberFormat('MM/dd/yyyy hh:mm:ss');
        
        Logger.log(`Successfully wrote ${dataRows.length} rows to sheet`);
      } catch (writeError) {
        throw new Error(`Failed to write data to sheet: ${writeError.toString()}`);
      }
    }
    
    let message = `Successfully saved ${validEmployees.length} employees to Google Sheets`;
    if (errors.length > 0) {
      message += `. Note: ${errors.length} records had validation issues.`;
    }
    
    Logger.log(message);
    return { 
      success: true, 
      message: message,
      saved: validEmployees.length,
      errors: errors.length
    };
    
  } catch (error) {
    const errorMessage = `Failed to save employees: ${error.toString()}`;
    Logger.log('Critical error in saveAllEmployees: ' + errorMessage);
    return { 
      success: false, 
      error: errorMessage
    };
  }
}

/**
 * Sync data from sheet (refresh from source)
 */
function syncFromSheet() {
  try {
    const result = getAllEmployees();
    if (result.success) {
      return { 
        success: true, 
        employees: result.employees,
        message: `Synced ${result.employees.length} employees from sheet`
      };
    } else {
      return result;
    }
  } catch (error) {
    Logger.log('Error in syncFromSheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get positions list with icons
 */
function getPositionsList() {
  try {
    Logger.log('getPositionsList: Loading positions...');
    
    const posSheet = getPositionsSheet();
    
    if (!posSheet) {
      Logger.log('getPositionsList: No positions sheet found, returning defaults');
      return { 
        success: true, 
        positions: DEFAULT_POSITIONS,
        message: 'Using default positions'
      };
    }
    
    const lastRow = posSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('getPositionsList: Positions sheet empty, returning defaults');
      return { 
        success: true, 
        positions: DEFAULT_POSITIONS,
        message: 'No custom positions found, using defaults'
      };
    }
    
    const dataRange = posSheet.getRange(2, 1, lastRow - 1, 2);
    const values = dataRange.getValues();
    
    const positions = [];
    values.forEach((row, index) => {
      if (row[0] && row[0].toString().trim() !== '') {
        positions.push({
          name: row[0].toString().trim(),
          icon: row[1] ? row[1].toString().trim() : '📋'
        });
      }
    });
    
    if (positions.length === 0) {
      Logger.log('getPositionsList: No valid custom positions, returning defaults');
      return { 
        success: true, 
        positions: DEFAULT_POSITIONS,
        message: 'No valid custom positions found, using defaults'
      };
    }
    
    Logger.log(`getPositionsList: Loaded ${positions.length} custom positions`);
    return { 
      success: true, 
      positions: positions,
      message: `Loaded ${positions.length} custom positions`
    };
    
  } catch (error) {
    Logger.log('Error in getPositionsList: ' + error.toString());
    return { 
      success: true, 
      positions: DEFAULT_POSITIONS,
      message: 'Error loading custom positions, using defaults',
      error: error.toString()
    };
  }
}

/**
 * Save positions list
 */
function savePositionsList(positions) {
  try {
    Logger.log(`savePositionsList: Saving ${positions ? positions.length : 0} positions`);
    
    if (!Array.isArray(positions)) {
      throw new Error('Positions must be an array');
    }
    
    const posSheet = getPositionsSheet();
    
    if (!posSheet) {
      throw new Error('Could not access positions configuration sheet');
    }
    
    const lastRow = posSheet.getLastRow();
    if (lastRow > 1) {
      const clearRange = posSheet.getRange(2, 1, lastRow - 1, 2);
      clearRange.clear();
    }
    
    if (positions && positions.length > 0) {
      const positionData = positions.map(pos => {
        if (typeof pos === 'string') {
          return [pos, '📋'];
        } else if (pos && typeof pos === 'object') {
          return [
            (pos.name || pos).toString().trim(),
            (pos.icon || '📋').toString().trim()
          ];
        } else {
          return ['Unknown Position', '📋'];
        }
      });
      
      if (positionData.length > 0) {
        const range = posSheet.getRange(2, 1, positionData.length, 2);
        range.setValues(positionData);
        Logger.log(`Saved ${positionData.length} positions to sheet`);
      }
    }
    
    return { 
      success: true, 
      message: `Successfully saved ${positions ? positions.length : 0} positions`,
      count: positions ? positions.length : 0
    };
    
  } catch (error) {
    Logger.log('Error in savePositionsList: ' + error.toString());
    return { 
      success: false, 
      error: `Failed to save positions: ${error.toString()}`
    };
  }
}

/**
 * Get/Set "me" employee ID (for highlighting current user)
 */
function getMeEmployeeId() {
  try {
    const props = PropertiesService.getScriptProperties();
    const empId = props.getProperty('ME_EMPLOYEE_ID') || '';
    return { success: true, empId: empId };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function setMeEmployeeId(empId) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('ME_EMPLOYEE_ID', empId || '');
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function clearMeEmployeeId() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('ME_EMPLOYEE_ID');
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Enhanced test function to verify complete setup
 */
function testSetup() {
  try {
    Logger.log('=== TESTING BAR EMPLOYEE CRM SETUP ===');
    
    const results = {
      success: true,
      tests: {},
      details: {},
      errors: []
    };
    
    // Test 1: Sheet creation/access
    try {
      Logger.log('Test 1: Testing sheet access...');
      const sheet = getCRMSheet();
      const sheetId = PropertiesService.getScriptProperties().getProperty('CRM_SHEET_ID');
      results.tests.sheetAccess = true;
      results.details.sheetId = sheetId;
      Logger.log('✓ Sheet access successful - ID: ' + sheetId);
    } catch (e) {
      results.tests.sheetAccess = false;
      results.errors.push('Sheet access failed: ' + e.toString());
      results.success = false;
      Logger.log('✗ Sheet access failed: ' + e.toString());
    }
    
    // Test 2: System info
    try {
      Logger.log('Test 2: Testing system info...');
      const sysInfo = getSystemInfo();
      results.tests.systemInfo = sysInfo.success;
      results.details.systemInfo = sysInfo.info;
      Logger.log('✓ System info retrieved successfully');
    } catch (e) {
      results.tests.systemInfo = false;
      results.errors.push('System info failed: ' + e.toString());
      Logger.log('✗ System info failed: ' + e.toString());
    }
    
    // Test 3: Positions loading
    try {
      Logger.log('Test 3: Testing positions...');
      const positions = getPositionsList();
      results.tests.positions = positions.success;
      results.details.positionsCount = positions.positions ? positions.positions.length : 0;
      Logger.log('✓ Positions loaded: ' + results.details.positionsCount);
    } catch (e) {
      results.tests.positions = false;
      results.errors.push('Positions loading failed: ' + e.toString());
      Logger.log('✗ Positions loading failed: ' + e.toString());
    }
    
    // Test 4: Employee data operations
    try {
      Logger.log('Test 4: Testing employee data operations...');
      const employees = getAllEmployees();
      results.tests.employeeData = employees.success;
      results.details.employeeCount = employees.employees ? employees.employees.length : 0;
      Logger.log('✓ Employee data operations successful - Count: ' + results.details.employeeCount);
    } catch (e) {
      results.tests.employeeData = false;
      results.errors.push('Employee data operations failed: ' + e.toString());
      Logger.log('✗ Employee data operations failed: ' + e.toString());
    }
    
    // Test 5: User preferences
    try {
      Logger.log('Test 5: Testing user preferences...');
      const meResult = getMeEmployeeId();
      results.tests.userPrefs = meResult.success;
      results.details.meEmployeeId = meResult.empId || 'Not set';
      Logger.log('✓ User preferences working - Me ID: ' + results.details.meEmployeeId);
    } catch (e) {
      results.tests.userPrefs = false;
      results.errors.push('User preferences failed: ' + e.toString());
      Logger.log('✗ User preferences failed: ' + e.toString());
    }
    
    // Test 6: Sheet URL access
    try {
      Logger.log('Test 6: Testing sheet URL access...');
      const urlResult = getSheetUrl();
      results.tests.sheetUrl = urlResult.success;
      results.details.sheetUrl = urlResult.url;
      Logger.log('✓ Sheet URL access successful');
    } catch (e) {
      results.tests.sheetUrl = false;
      results.errors.push('Sheet URL access failed: ' + e.toString());
      Logger.log('✗ Sheet URL access failed: ' + e.toString());
    }
    
    // Test 7: Date handling
    try {
      Logger.log('Test 7: Testing date handling...');
      const testDate = new Date();
      const parsedDate = safeParseDate(testDate.toISOString());
      const formattedDate = formatDateForDisplay(testDate);
      results.tests.dateHandling = !!parsedDate && !!formattedDate;
      results.details.dateHandling = {
        originalDate: testDate.toISOString(),
        parsedDate: parsedDate ? parsedDate.toISOString() : null,
        formattedDate: formattedDate
      };
      Logger.log('✓ Date handling working correctly');
    } catch (e) {
      results.tests.dateHandling = false;
      results.errors.push('Date handling failed: ' + e.toString());
      Logger.log('✗ Date handling failed: ' + e.toString());
    }
    
    // Final summary
    const passedTests = Object.values(results.tests).filter(test => test === true).length;
    const totalTests = Object.keys(results.tests).length;
    
    Logger.log(`=== TEST SUMMARY: ${passedTests}/${totalTests} PASSED ===`);
    
    if (results.success) {
      Logger.log('🎉 ALL TESTS PASSED - CRM is ready to use!');
      results.message = `Setup test completed successfully! All ${totalTests} tests passed.`;
    } else {
      Logger.log('⚠️  SOME TESTS FAILED - Check errors above');
      results.message = `Setup test completed with issues. ${passedTests}/${totalTests} tests passed.`;
      results.success = false;
    }
    
    results.details.testSummary = `${passedTests}/${totalTests} tests passed`;
    results.details.timestamp = new Date().toISOString();
    
    return results;
    
  } catch (error) {
    Logger.log('CRITICAL ERROR in testSetup: ' + error.toString());
    return { 
      success: false, 
      error: 'Critical test failure: ' + error.toString(),
      message: 'Setup test failed to complete',
      details: { timestamp: new Date().toISOString() }
    };
  }
}

/**
 * Quick connectivity test (lighter version)
 */
function testConnection() {
  try {
    const sheet = getCRMSheet();
    const sysInfo = getSystemInfo();
    
    return {
      success: true,
      message: 'Connection test passed',
      details: {
        sheetAccessible: !!sheet,
        systemInfoAvailable: sysInfo.success,
        timestamp: new Date().toISOString()
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      message: 'Connection test failed'
    };
  }
}

/**
 * Initialize sample data (for new installations)
 */
function initializeSampleData() {
  try {
    Logger.log('Initializing sample data...');
    
    // Check if data already exists
    const existing = getAllEmployees();
    if (existing.success && existing.employees && existing.employees.length > 0) {
      return {
        success: false,
        message: 'Sample data not added - existing employees found',
        existingCount: existing.employees.length
      };
    }
    
    // Sample employees with proper date handling
    const now = new Date().toISOString();
    const sampleEmployees = [
      {
        empId: 'MGR001',
        firstName: 'Alex',
        lastName: 'Johnson',
        phone: '(555) 123-4567',
        email: 'alex.johnson@yourbar.com',
        position: 'Manager',
        status: 'Active',
        note: 'Bar manager with 8+ years experience',
        photoUrl: '',
        createdDate: now,
        lastModified: now
      },
      {
        empId: 'BTD001',
        firstName: 'Sarah',
        lastName: 'Martinez',
        phone: '(555) 987-6543',
        email: 'sarah.martinez@yourbar.com',
        position: 'Bartender',
        status: 'Active',
        note: 'Expert mixologist, specializes in craft cocktails',
        photoUrl: '',
        createdDate: now,
        lastModified: now
      },
      {
        empId: 'SRV001',
        firstName: 'Mike',
        lastName: 'Chen',
        phone: '(555) 456-7890',
        email: 'mike.chen@yourbar.com',
        position: 'Server',
        status: 'Active',
        note: 'Excellent customer service, knows wine pairings',
        photoUrl: '',
        createdDate: now,
        lastModified: now
      }
    ];
    
    const saveResult = saveAllEmployees(sampleEmployees);
    
    if (saveResult.success) {
      Logger.log('Sample data initialized successfully');
      return {
        success: true,
        message: 'Sample data added successfully',
        sampleCount: sampleEmployees.length
      };
    } else {
      return {
        success: false,
        message: 'Failed to add sample data',
        error: saveResult.error
      };
    }
    
  } catch (error) {
    Logger.log('Error initializing sample data: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Failed to initialize sample data'
    };
  }
}

/**
 * Backup and restore functions
 */
function createBackup() {
  try {
    Logger.log('Creating backup...');
    
    const employees = getAllEmployees();
    const positions = getPositionsList();
    const sysInfo = getSystemInfo();
    
    const backup = {
      timestamp: new Date().toISOString(),
      version: '2.0',
      employees: employees.success ? employees.employees : [],
      positions: positions.success ? positions.positions : DEFAULT_POSITIONS,
      systemInfo: sysInfo.success ? sysInfo.info : {},
      employeeCount: employees.success ? employees.employees.length : 0,
      positionsCount: positions.success ? positions.positions.length : 0
    };
    
    // Save backup to script properties (for recent backup)
    try {
      const props = PropertiesService.getScriptProperties();
      props.setProperty('LAST_BACKUP', JSON.stringify({
        timestamp: backup.timestamp,
        employeeCount: backup.employeeCount,
        success: true
      }));
    } catch (propsError) {
      Logger.log('Warning: Could not save backup info to properties: ' + propsError.toString());
    }
    
    Logger.log(`Backup created successfully - ${backup.employeeCount} employees, ${backup.positionsCount} positions`);
    
    return {
      success: true,
      backup: backup,
      message: `Backup created with ${backup.employeeCount} employees and ${backup.positionsCount} positions`
    };
    
  } catch (error) {
    Logger.log('Error creating backup: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Failed to create backup'
    };
  }
}

function getLastBackupInfo() {
  try {
    const props = PropertiesService.getScriptProperties();
    const backupInfo = props.getProperty('LAST_BACKUP');
    
    if (backupInfo) {
      const info = JSON.parse(backupInfo);
      return { success: true, backupInfo: info };
    } else {
      return { success: false, message: 'No backup information found' };
    }
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Advanced sheet management
 */
function repairSheet() {
  try {
    Logger.log('Starting sheet repair...');
    
    const sheet = getCRMSheet();
    
    // Verify and fix headers
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
    const currentHeaders = headerRange.getValues()[0];
    
    let headersFixed = false;
    HEADERS.forEach((expectedHeader, index) => {
      if (currentHeaders[index] !== expectedHeader) {
        Logger.log(`Fixing header ${index + 1}: "${currentHeaders[index]}" -> "${expectedHeader}"`);
        sheet.getRange(HEADER_ROW, index + 1).setValue(expectedHeader);
        headersFixed = true;
      }
    });
    
    // Re-apply header formatting
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    
    // Check and fix column widths
    const columnWidths = [100, 120, 120, 140, 180, 200, 80, 250, 200, 120, 120];
    columnWidths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });
    
    // Re-apply data validation
    const statusRange = sheet.getRange(DATA_START_ROW, 7, 1000, 1);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Inactive'])
      .setAllowInvalid(false)
      .setHelpText('Select Active or Inactive')
      .build();
    statusRange.setDataValidation(statusRule);
    
    sheet.setFrozenRows(1);
    
    Logger.log('Sheet repair completed');
    
    return {
      success: true,
      message: 'Sheet repair completed successfully',
      headersFixed: headersFixed
    };
    
  } catch (error) {
    Logger.log('Error repairing sheet: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      message: 'Sheet repair failed'
    };
  }
}

/**
 * Upload employee photo
 */
function uploadEmployeePhoto(dataUrl, fileName, empId) {
  try {
    // Get or create photos folder
    let photosFolder;
    const folders = DriveApp.getFoldersByName(PHOTOS_FOLDER_NAME);
    
    if (folders.hasNext()) {
      photosFolder = folders.next();
    } else {
      photosFolder = DriveApp.createFolder(PHOTOS_FOLDER_NAME);
    }
    
    // Extract base64 data
    const base64Data = dataUrl.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg');
    
    // Create unique filename
    const timestamp = new Date().getTime();
    const cleanEmpId = empId.replace(/[^a-zA-Z0-9]/g, '_');
    const uniqueFileName = `${cleanEmpId}_${timestamp}_${fileName}`;
    
    // Delete existing photos for this employee
    const existingFiles = photosFolder.getFilesByName(cleanEmpId);
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      if (file.getName().startsWith(cleanEmpId + '_')) {
        photosFolder.removeFile(file);
      }
    }
    
    // Upload new photo
    const file = photosFolder.createFile(blob.setName(uniqueFileName));
    
    // Make file publicly viewable
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Get the public URL
    const photoUrl = `https://drive.google.com/uc?id=${file.getId()}`;
    
    Logger.log(`Uploaded photo for employee ${empId}: ${photoUrl}`);
    return { success: true, url: photoUrl };
  } catch (error) {
    Logger.log('Error uploading photo: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Backup data (export to JSON)
 */
function exportBackup() {
  try {
    const employeesResult = getAllEmployees();
    const positionsResult = getPositionsList();
    
    if (!employeesResult.success) {
      throw new Error('Failed to get employees data');
    }
    
    const backup = {
      timestamp: new Date().toISOString(),
      employees: employeesResult.employees || [],
      positions: positionsResult.success ? positionsResult.positions : DEFAULT_POSITIONS,
      version: '2.0'
    };
    
    return { success: true, backup: backup };
  } catch (error) {
    Logger.log('Error creating backup: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Clean up old photo files (run periodically)
 */
function cleanupOldPhotos() {
  try {
    const folders = DriveApp.getFoldersByName(PHOTOS_FOLDER_NAME);
    if (!folders.hasNext()) {
      return { success: true, message: 'No photos folder found' };
    }
    
    const photosFolder = folders.next();
    const files = photosFolder.getFiles();
    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - 6); // 6 months ago
    
    let deletedCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getDateCreated() < cutoffDate) {
        photosFolder.removeFile(file);
        deletedCount++;
      }
    }
    
    Logger.log(`Cleaned up ${deletedCount} old photo files`);
    return { success: true, message: `Deleted ${deletedCount} old files` };
  } catch (error) {
    Logger.log('Error cleaning up photos: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
