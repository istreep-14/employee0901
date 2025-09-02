/**
 * Bar Employee CRM - Standalone Web App
 * Creates its own Google Sheet for data storage and syncing
 * 
 * Setup Instructions:
 * 1. Create a new Google Apps Script project (script.google.com)
 * 2. Replace Code.gs content with this script
 * 3. Add the HTML file as 'index.html'
 * 4. Deploy as Web App with execute as "Me" and access "Anyone"
 * 5. Use the web app URL to access your CRM
 */

// Configuration
const CRM_SHEET_NAME = 'Bar_Employee_CRM_Data';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Headers in order
const HEADERS = ['Emp Id', 'First Name', 'Last Name', 'Phone', 'Email', 'Position', 'Status', 'Note', 'Photo URL', 'Created Date', 'Last Modified'];

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
 * Include external files (for CSS/JS in HTML template)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get or create the CRM data sheet
 * This creates a dedicated Google Sheet for the CRM data
 */
function getCRMSheet() {
  const scriptId = ScriptApp.getScriptId();
  let sheet = null;
  
  try {
    // Try to get existing sheet from script properties
    const props = PropertiesService.getScriptProperties();
    const sheetId = props.getProperty('CRM_SHEET_ID');
    
    if (sheetId) {
      try {
        const spreadsheet = SpreadsheetApp.openById(sheetId);
        sheet = spreadsheet.getSheetByName(CRM_SHEET_NAME);
      } catch (e) {
        Logger.log('Existing sheet not accessible, will create new one: ' + e.toString());
      }
    }
    
    // Create new sheet if none exists or not accessible
    if (!sheet) {
      const spreadsheet = SpreadsheetApp.create('Bar Employee CRM Data - ' + new Date().getFullYear());
      sheet = spreadsheet.getActiveSheet();
      sheet.setName(CRM_SHEET_NAME);
      
      // Store the sheet ID for future use
      props.setProperty('CRM_SHEET_ID', spreadsheet.getId());
      
      // Initialize with headers
      initializeSheet(sheet);
      
      Logger.log('Created new CRM sheet: ' + spreadsheet.getId());
    }
    
    return sheet;
  } catch (error) {
    Logger.log('Error in getCRMSheet: ' + error.toString());
    throw new Error('Failed to access or create CRM data sheet: ' + error.toString());
  }
}

/**
 * Initialize the sheet with headers
 */
function initializeSheet(sheet) {
  try {
    if (!sheet) {
      sheet = getCRMSheet();
    }
    
    // Set headers
    sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS]);
    
    // Format headers
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setBorder(true, true, true, true, true, true);
    
    // Set column widths for better display
    sheet.setColumnWidth(1, 100); // Emp ID
    sheet.setColumnWidth(2, 120); // First Name
    sheet.setColumnWidth(3, 120); // Last Name
    sheet.setColumnWidth(4, 120); // Phone
    sheet.setColumnWidth(5, 150); // Email
    sheet.setColumnWidth(6, 150); // Position
    sheet.setColumnWidth(7, 80);  // Status
    sheet.setColumnWidth(8, 200); // Note
    sheet.setColumnWidth(9, 200); // Photo URL
    sheet.setColumnWidth(10, 120); // Created Date
    sheet.setColumnWidth(11, 120); // Last Modified
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    Logger.log('Sheet initialized with headers');
    return { success: true, message: 'Sheet initialized successfully' };
    
  } catch (error) {
    Logger.log('Error in initializeSheet: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get the Google Sheet URL for users to access the data directly
 */
function getSheetUrl() {
  try {
    const props = PropertiesService.getScriptProperties();
    const sheetId = props.getProperty('CRM_SHEET_ID');
    
    if (sheetId) {
      return { 
        success: true, 
        url: `https://docs.google.com/spreadsheets/d/${sheetId}/edit`,
        sheetId: sheetId
      };
    } else {
      // Try to create sheet if it doesn't exist
      const sheet = getCRMSheet();
      const newSheetId = props.getProperty('CRM_SHEET_ID');
      return { 
        success: true, 
        url: `https://docs.google.com/spreadsheets/d/${newSheetId}/edit`,
        sheetId: newSheetId
      };
    }
  } catch (error) {
    Logger.log('Error in getSheetUrl: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Persist positions list in script properties
 */
function getPositionsList() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('CRM_POSITIONS_LIST') || '';
    const positions = raw ? JSON.parse(raw) : [];
    return { success: true, positions };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function savePositionsList(list) {
  try {
    const normalized = (list || []).map(function(s){ return String(s || '').trim(); }).filter(function(s){ return s.length > 0; });
    PropertiesService.getScriptProperties().setProperty('CRM_POSITIONS_LIST', JSON.stringify(normalized));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Persist and retrieve the "Me" employee id
 */
function getMeEmployeeId() {
  try {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('CRM_ME_EMP_ID') || '';
    return { success: true, empId: id };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function setMeEmployeeId(empId) {
  try {
    PropertiesService.getScriptProperties().setProperty('CRM_ME_EMP_ID', String(empId || ''));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function clearMeEmployeeId() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('CRM_ME_EMP_ID');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
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
      .filter(row => row[0] !== '') // Filter out empty rows (Emp Id is required)
      .map(row => ({
        empId: row[0] || '',
        firstName: row[1] || '',
        lastName: row[2] || '',
        phone: row[3] || '',
        email: row[4] || '',
        position: row[5] || '',
        status: row[6] || 'Active',
        note: row[7] || '',
        photoUrl: row[8] || '',
        createdDate: row[9] || '',
        lastModified: row[10] || ''
      }));
    
    Logger.log(`Retrieved ${employees.length} employees`);
    return { success: true, employees: employees };
    
  } catch (error) {
    Logger.log('Error in getAllEmployees: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Save all employees to the sheet (replaces existing data)
 */
function saveAllEmployees(employees) {
  try {
    const sheet = getCRMSheet();
    const now = new Date();
    
    // Ensure headers are present and up-to-date
    const existingHeaders = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).getValues()[0];
    const hasHeaders = existingHeaders.some(header => header !== '');
    
    if (!hasHeaders) {
      initializeSheet(sheet);
    }
    
    // Clear existing data (keep headers)
    const lastRow = sheet.getLastRow();
    if (lastRow >= DATA_START_ROW) {
      const maxCols = Math.max(HEADERS.length, sheet.getLastColumn());
      sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW + 1, maxCols).clear();
    }
    
    if (employees && employees.length > 0) {
      // Prepare data rows
      const dataRows = employees.map(emp => [
        emp.empId || '',
        emp.firstName || '',
        emp.lastName || '',
        emp.phone || '',
        emp.email || '',
        emp.position || '',
        emp.status || 'Active',
        emp.note || '',
        emp.photoUrl || '',
        emp.createdDate || now,
        now // Always update last modified
      ]);
      
      // Write data to sheet
      const range = sheet.getRange(DATA_START_ROW, 1, dataRows.length, HEADERS.length);
      range.setValues(dataRows);
      
      Logger.log(`Saved ${employees.length} employees`);
    }
    
    return { success: true, message: `Saved ${employees ? employees.length : 0} employees` };
    
  } catch (error) {
    Logger.log('Error in saveAllEmployees: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Add a single employee
 */
function addEmployee(employee) {
  try {
    const sheet = getCRMSheet();
    const now = new Date();
    
    // Check for duplicate Emp ID
    const existingData = getAllEmployees();
    if (existingData.success) {
      const duplicate = existingData.employees.find(emp => emp.empId === employee.empId);
      if (duplicate) {
        return { success: false, error: 'Employee ID already exists' };
      }
    }
    
    // Ensure headers exist
    const existingHeaders = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).getValues()[0];
    const hasHeaders = existingHeaders.some(header => header !== '');
    
    if (!hasHeaders) {
      initializeSheet(sheet);
    }
    
    // Add to the end of the sheet
    const newRow = [
      employee.empId || '',
      employee.firstName || '',
      employee.lastName || '',
      employee.phone || '',
      employee.email || '',
      employee.position || '',
      employee.status || 'Active',
      employee.note || '',
      employee.photoUrl || '',
      now, // Created date
      now  // Last modified
    ];
    
    sheet.appendRow(newRow);
    
    Logger.log(`Added employee: ${employee.empId}`);
    return { success: true, message: 'Employee added successfully' };
    
  } catch (error) {
    Logger.log('Error in addEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Update an existing employee
 */
function updateEmployee(employee, originalEmpId) {
  try {
    const sheet = getCRMSheet();
    const lastRow = sheet.getLastRow();
    const now = new Date();
    
    if (lastRow < DATA_START_ROW) {
      return { success: false, error: 'No employees found' };
    }
    
    // Find the employee row
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, HEADERS.length);
    const values = dataRange.getValues();
    const rowIndex = values.findIndex(row => row[0] === originalEmpId);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Check for duplicate Emp ID if it's being changed
    if (employee.empId !== originalEmpId) {
      const duplicate = values.findIndex(row => row[0] === employee.empId);
      if (duplicate !== -1) {
        return { success: false, error: 'New Employee ID already exists' };
      }
    }
    
    // Keep original created date
    const originalCreatedDate = values[rowIndex][9] || now;
    
    // Update the row
    const targetRow = DATA_START_ROW + rowIndex;
    const updatedRow = [
      employee.empId || '',
      employee.firstName || '',
      employee.lastName || '',
      employee.phone || '',
      employee.email || '',
      employee.position || '',
      employee.status || 'Active',
      employee.note || '',
      employee.photoUrl || '',
      originalCreatedDate, // Keep original created date
      now // Update last modified
    ];
    
    sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([updatedRow]);
    
    Logger.log(`Updated employee: ${originalEmpId} -> ${employee.empId}`);
    return { success: true, message: 'Employee updated successfully' };
    
  } catch (error) {
    Logger.log('Error in updateEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Delete an employee
 */
function deleteEmployee(empId) {
  try {
    const sheet = getCRMSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: false, error: 'No employees found' };
    }
    
    // Find the employee row
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1);
    const empIds = dataRange.getValues().flat();
    const rowIndex = empIds.findIndex(id => id === empId);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Delete the row
    const targetRow = DATA_START_ROW + rowIndex;
    sheet.deleteRow(targetRow);
    
    Logger.log(`Deleted employee: ${empId}`);
    return { success: true, message: 'Employee deleted successfully' };
    
  } catch (error) {
    Logger.log('Error in deleteEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Upload a photo to Google Drive and return an embeddable URL
 */
function uploadEmployeePhoto(dataUrl, fileName, empId) {
  try {
    if (!dataUrl) {
      return { success: false, error: 'Missing dataUrl' };
    }
    const matches = dataUrl.match(/^data:(.+);base64,(.*)$/);
    if (!matches) {
      return { success: false, error: 'Invalid data URL' };
    }
    const contentType = matches[1];
    const base64Data = matches[2];
    const bytes = Utilities.base64Decode(base64Data);
    const tz = Session.getScriptTimeZone() || 'Etc/GMT';
    const timestamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
    const safeEmpId = (empId || 'employee').toString().replace(/[^a-zA-Z0-9_-]+/g, '_');
    const baseName = `${safeEmpId}_${timestamp}`;
    const finalName = fileName ? `${baseName}_${fileName}` : `${baseName}.png`;
    const blob = Utilities.newBlob(bytes, contentType, finalName);

    const folderName = 'Employee CRM Photos';
    let folderIter = DriveApp.getFoldersByName(folderName);
    let folder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder(folderName);

    const file = folder.createFile(blob).setName(finalName);
    // Make viewable via link for embedding in <img>
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // In some domains, ANYONE_WITH_LINK may be restricted; ignore if fails
    }
    const id = file.getId();
    const viewUrl = `https://drive.google.com/uc?export=view&id=${id}`;
    return { success: true, url: viewUrl, id: id };
  } catch (error) {
    Logger.log('Error in uploadEmployeePhoto: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Sync data from external sheet (if users edit the Google Sheet directly)
 */
function syncFromSheet() {
  try {
    // This function allows the web app to pull in changes made directly to the Google Sheet
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
 * Get system information for the web app
 */
function getSystemInfo() {
  try {
    const props = PropertiesService.getScriptProperties();
    const sheetId = props.getProperty('CRM_SHEET_ID');
    
    return {
      success: true,
      info: {
        scriptId: ScriptApp.getScriptId(),
        sheetId: sheetId,
        sheetUrl: sheetId ? `https://docs.google.com/spreadsheets/d/${sheetId}/edit` : null,
        timezone: Session.getScriptTimeZone(),
        userEmail: Session.getActiveUser().getEmail()
      }
    };
  } catch (error) {
    Logger.log('Error in getSystemInfo: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
