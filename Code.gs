/**
 * Simplified Bar Employee CRM - Google Apps Script
 * 
 * Setup Instructions:
 * 1. Create a new Google Apps Script project (script.google.com)
 * 2. Replace Code.gs content with this script
 * 3. Add the HTML file as 'index.html'
 * 4. Deploy as Web App with execute as "Me" and access "Anyone"
 */

// Configuration
const CRM_SHEET_NAME = 'Bar_Employee_CRM_Data';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Headers in order
const HEADERS = ['Emp Id', 'First Name', 'Last Name', 'Phone', 'Email', 'Position', 'Status', 'Note', 'Created Date', 'Last Modified'];

/**
 * Serves the main web app HTML
 */
function doGet() {
  const html = HtmlService.createTemplateFromFile('index');
  return html.evaluate()
    .setTitle('Bar Employee CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get or create the CRM data sheet
 */
function getCRMSheet() {
  const scriptId = ScriptApp.getScriptId();
  let sheet = null;
  
  try {
    const props = PropertiesService.getScriptProperties();
    const sheetId = props.getProperty('CRM_SHEET_ID');
    
    if (sheetId) {
      try {
        const spreadsheet = SpreadsheetApp.openById(sheetId);
        sheet = spreadsheet.getSheetByName(CRM_SHEET_NAME);
      } catch (e) {
        Logger.log('Existing sheet not accessible, will create new one');
      }
    }
    
    // Create new sheet if none exists
    if (!sheet) {
      const spreadsheet = SpreadsheetApp.create('Bar Employee CRM Data - ' + new Date().getFullYear());
      sheet = spreadsheet.getActiveSheet();
      sheet.setName(CRM_SHEET_NAME);
      
      props.setProperty('CRM_SHEET_ID', spreadsheet.getId());
      initializeSheet(sheet);
    }
    
    return sheet;
  } catch (error) {
    throw new Error('Failed to access or create CRM data sheet: ' + error.toString());
  }
}

/**
 * Initialize the sheet with headers
 */
function initializeSheet(sheet) {
  try {
    // Set headers
    sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS]);
    
    // Format headers
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 100); // Emp ID
    sheet.setColumnWidth(2, 120); // First Name
    sheet.setColumnWidth(3, 120); // Last Name
    sheet.setColumnWidth(4, 120); // Phone
    sheet.setColumnWidth(5, 150); // Email
    sheet.setColumnWidth(6, 150); // Position
    sheet.setColumnWidth(7, 80);  // Status
    sheet.setColumnWidth(8, 200); // Note
    sheet.setColumnWidth(9, 120); // Created Date
    sheet.setColumnWidth(10, 120); // Last Modified
    
    sheet.setFrozenRows(1);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
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
 * Get all employees from the sheet
 */
function getAllEmployees() {
  try {
    const sheet = getCRMSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: true, employees: [] };
    }
    
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, HEADERS.length);
    const values = dataRange.getValues();
    
    const employees = values
      .filter(row => row[0] !== '') // Filter out empty rows
      .map(row => ({
        empId: row[0] || '',
        firstName: row[1] || '',
        lastName: row[2] || '',
        phone: row[3] || '',
        email: row[4] || '',
        position: row[5] || '',
        status: row[6] || 'Active',
        note: row[7] || '',
        createdDate: row[8] ? row[8].toISOString() : '', // Convert Date to ISO string
        lastModified: row[9] ? row[9].toISOString() : '' // Convert Date to ISO string
      }));
    
    return { success: true, employees: employees };
  } catch (error) {
    Logger.log('Error in getAllEmployees: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Save all employees to the sheet
 */
function saveAllEmployees(employees) {
  try {
    const sheet = getCRMSheet();
    const now = new Date();
    
    // Ensure headers exist
    const existingHeaders = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).getValues()[0];
    const hasHeaders = existingHeaders.some(header => header !== '');
    
    if (!hasHeaders) {
      initializeSheet(sheet);
    }
    
    // Clear existing data (keep headers)
    const lastRow = sheet.getLastRow();
    if (lastRow >= DATA_START_ROW) {
      sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW + 1, HEADERS.length).clear();
    }
    
    if (employees && employees.length > 0) {
      const dataRows = employees.map(emp => [
        emp.empId || '',
        emp.firstName || '',
        emp.lastName || '',
        emp.phone || '',
        emp.email || '',
        emp.position || '',
        emp.status || 'Active',
        emp.note || '',
        emp.createdDate ? new Date(emp.createdDate) : now, // Convert string back to Date
        now // Always update last modified with current date
      ]);
      
      const range = sheet.getRange(DATA_START_ROW, 1, dataRows.length, HEADERS.length);
      range.setValues(dataRows);
    }
    
    return { success: true, message: `Saved ${employees ? employees.length : 0} employees` };
  } catch (error) {
    Logger.log('Error in saveAllEmployees: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Sync data from external sheet
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
    return { success: false, error: error.toString() };
  }
}
