/**
 * Bar Employee CRM - Google Apps Script Backend
 * This script handles the Google Sheets integration for the Bar Employee CRM
 * 
 * Setup Instructions:
 * 1. Open Google Apps Script (script.google.com)
 * 2. Create a new project
 * 3. Replace Code.gs content with this script
 * 4. Create a Google Sheet with "Sheet2" tab
 * 5. Deploy as web app with execute permissions for "Anyone"
 * 6. Copy the web app URL to use in your HTML file
 */

// Configuration
const SHEET_NAME = 'Sheet2';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Headers in order (must match your CRM structure)
const HEADERS = ['Emp Id', 'First Name', 'Last Name', 'Phone', 'Email', 'Position', 'Note'];

/**
 * Initialize the sheet with headers if they don't exist
 */
function initializeSheet() {
  const sheet = getSheet();
  
  // Check if headers exist
  const existingHeaders = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = existingHeaders.some(header => header !== '');
  
  if (!hasHeaders) {
    // Set headers
    sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS]);
    
    // Format headers
    const headerRange = sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
    headerRange.setBackground('#f8f9fa');
    headerRange.setFontWeight('bold');
    headerRange.setBorder(true, true, true, true, true, true);
    
    Logger.log('Sheet initialized with headers');
  }
  
  return sheet;
}

/**
 * Get or create the target sheet
 */
function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    Logger.log(`Created new sheet: ${SHEET_NAME}`);
  }
  
  return sheet;
}

/**
 * Handle HTTP requests (GET and POST)
 */
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    
    Logger.log(`Received ${action} request`);
    
    switch (action) {
      case 'getAllEmployees':
        return createResponse(getAllEmployees());
      
      case 'saveAllEmployees':
        return createResponse(saveAllEmployees(requestData.employees));
      
      case 'addEmployee':
        return createResponse(addEmployee(requestData.employee));
      
      case 'updateEmployee':
        return createResponse(updateEmployee(requestData.employee, requestData.originalEmpId));
      
      case 'deleteEmployee':
        return createResponse(deleteEmployee(requestData.empId));
      
      case 'initializeSheet':
        initializeSheet();
        return createResponse({ success: true, message: 'Sheet initialized' });
      
      default:
        return createResponse({ success: false, error: 'Unknown action' });
    }
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    return createResponse({ success: false, error: error.toString() });
  }
}

/**
 * Handle GET requests (for testing)
 */
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'test') {
    return createResponse({ 
      success: true, 
      message: 'Bar Employee CRM API is working!',
      timestamp: new Date().toISOString()
    });
  }
  
  if (action === 'getAllEmployees') {
    return createResponse(getAllEmployees());
  }
  
  return createResponse({ success: false, error: 'Invalid GET request' });
}

/**
 * Get all employees from the sheet
 */
function getAllEmployees() {
  try {
    const sheet = initializeSheet();
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
        note: row[6] || ''
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
    const sheet = initializeSheet();
    
    // Clear existing data (keep headers)
    const lastRow = sheet.getLastRow();
    if (lastRow >= DATA_START_ROW) {
      sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW + 1, HEADERS.length).clear();
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
        emp.note || ''
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
    const sheet = initializeSheet();
    
    // Check for duplicate Emp ID
    const existingData = getAllEmployees();
    if (existingData.success) {
      const duplicate = existingData.employees.find(emp => emp.empId === employee.empId);
      if (duplicate) {
        return { success: false, error: 'Employee ID already exists' };
      }
    }
    
    // Add to the end of the sheet
    const newRow = [
      employee.empId || '',
      employee.firstName || '',
      employee.lastName || '',
      employee.phone || '',
      employee.email || '',
      employee.position || '',
      employee.note || ''
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
    const sheet = initializeSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: false, error: 'No employees found' };
    }
    
    // Find the employee row
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1); // Just Emp ID column
    const empIds = dataRange.getValues().flat();
    const rowIndex = empIds.findIndex(id => id === originalEmpId);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Check for duplicate Emp ID if it's being changed
    if (employee.empId !== originalEmpId) {
      const duplicate = empIds.findIndex(id => id === employee.empId);
      if (duplicate !== -1) {
        return { success: false, error: 'New Employee ID already exists' };
      }
    }
    
    // Update the row
    const targetRow = DATA_START_ROW + rowIndex;
    const updatedRow = [
      employee.empId || '',
      employee.firstName || '',
      employee.lastName || '',
      employee.phone || '',
      employee.email || '',
      employee.position || '',
      employee.note || ''
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
    const sheet = initializeSheet();
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
 * Create a standardized response
 */
function createResponse(data) {
  const response = ContentService.createTextOutput(JSON.stringify(data));
  response.setMimeType(ContentService.MimeType.JSON);
  
  // Add CORS headers to allow cross-origin requests
  return response;
}

/**
 * Test function - run this to test your setup
 */
function testSetup() {
  console.log('Testing Bar Employee CRM setup...');
  
  // Initialize sheet
  initializeSheet();
  console.log('✓ Sheet initialized');
  
  // Test adding an employee
  const testEmployee = {
    empId: 'TEST001',
    firstName: 'John',
    lastName: 'Doe',
    phone: '555-0123',
    email: 'john.doe@example.com',
    position: 'Bartender, Server',
    note: 'Test employee'
  };
  
  const addResult = addEmployee(testEmployee);
  console.log('Add result:', addResult);
  
  // Test getting all employees
  const getResult = getAllEmployees();
  console.log('Get result:', getResult);
  
  // Clean up test data
  if (getResult.success && getResult.employees.length > 0) {
    const deleteResult = deleteEmployee('TEST001');
    console.log('Delete result:', deleteResult);
  }
  
  console.log('✓ Test completed successfully!');
}

/**
 * Create a simple web app interface for testing
 */
function doGet(e) {
  if (e.parameter.action === 'testInterface') {
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>Bar Employee CRM API Test</title>
          <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            button { padding: 10px 20px; margin: 5px; }
            .result { margin: 20px 0; padding: 10px; background: #f5f5f5; border-radius: 4px; }
          </style>
        </head>
        <body>
          <h1>Bar Employee CRM API Test Interface</h1>
          <p>This interface helps you test the Google Apps Script API.</p>
          
          <button onclick="testAPI('getAllEmployees')">Get All Employees</button>
          <button onclick="testAPI('initializeSheet')">Initialize Sheet</button>
          
          <div id="result" class="result">
            Click a button to test the API...
          </div>
          
          <script>
            function testAPI(action) {
              const resultDiv = document.getElementById('result');
              resultDiv.innerHTML = 'Loading...';
              
              fetch(window.location.href.split('?')[0], {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ action: action })
              })
              .then(response => response.json())
              .then(data => {
                resultDiv.innerHTML = '<pre>' + JSON.stringify(data, null, 2) + '</pre>';
              })
              .catch(error => {
                resultDiv.innerHTML = 'Error: ' + error.message;
              });
            }
          </script>
        </body>
      </html>
    `);
  }
  
  // Default response for other GET requests
  return createResponse(getAllEmployees());
}
