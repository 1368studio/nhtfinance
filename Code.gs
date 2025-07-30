// =================== NHT_Code.gs - FIXED SHEET NAMES ===================

/**
 * CẤU HÌNH TÊN SHEET - THAY ĐỔI CHỖ NÀY THEO TÊN SHEET THỰC TẾ CỦA BẠN
 */
const SHEET_NAMES = {
  TRANSACTIONS: 'Transactions',    // Thay bằng tên sheet giao dịch thực tế
  ACCOUNTS: 'Accounts',           // Thay bằng tên sheet tài khoản thực tế  
  CATEGORIES: 'Categories',       // Thay bằng tên sheet danh mục thực tế
  CUSTOMERS: 'Customers',         // Thay bằng tên sheet khách hàng thực tế
  SUPPLIERS: 'Suppliers',         // Thay bằng tên sheet nhà cung cấp thực tế
  USERS: 'Users'                  // Thay bằng tên sheet users thực tế
};

/**
 * Main entry point for the web application
 */
function doGet() {
  console.log('🚀 Starting NHT Finance System...');
  
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Hệ thống quản lý tài chính NHT')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    console.error('❌ Error in doGet:', error);
    
    // Return emergency HTML page
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>Error - NHT Finance</title>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <style>
            body { 
              font-family: Arial, sans-serif; 
              text-align: center; 
              padding: 50px; 
              background: #f8f9fa;
            }
            .error-container {
              background: white;
              padding: 40px;
              border-radius: 12px;
              box-shadow: 0 4px 12px rgba(0,0,0,0.1);
              max-width: 500px;
              margin: 0 auto;
            }
            .error { color: #dc3545; }
            .btn {
              padding: 12px 24px;
              background: #007bff;
              color: white;
              border: none;
              border-radius: 6px;
              cursor: pointer;
              font-size: 16px;
              margin: 10px;
            }
            .btn:hover { background: #0056b3; }
          </style>
        </head>
        <body>
          <div class="error-container">
            <h1 class="error">⚠️ Lỗi hệ thống</h1>
            <p>Không thể tải ứng dụng. Vui lòng thử lại sau.</p>
            <button class="btn" onclick="window.location.reload()">🔄 Làm mới trang</button>
            <details style="margin-top: 20px; text-align: left;">
              <summary style="cursor: pointer;">Chi tiết lỗi</summary>
              <pre style="background: #f8f9fa; padding: 10px; border-radius: 4px; font-size: 12px;">${error.message}</pre>
            </details>
          </div>
        </body>
      </html>
    `);
  }
}

/**
 * Include HTML files for templating
 */
function include(filename) {
  try {
    console.log(`📄 Including file: ${filename}`);
    
    const content = HtmlService.createHtmlOutputFromFile(filename).getContent();
    console.log(`✅ Successfully included file: ${filename}`);
    return content;
    
  } catch (error) {
    console.error(`❌ Error including file ${filename}:`, error);
    
    // Return appropriate fallback content
    if (filename.includes('Stylesheet') || filename === 'Stylesheet') {
      return `
        /* Error loading ${filename}: ${error.message} */
        body { 
          font-family: Arial, sans-serif; 
          margin: 20px; 
          background: #f8f9fa;
        }
        .error { 
          color: #dc3545; 
          background: #f8d7da; 
          padding: 15px; 
          border-radius: 6px; 
          border: 1px solid #f5c6cb;
          margin: 20px 0;
        }
      `;
    } else if (filename.includes('JavaScript') || filename === 'JavaScript') {
      return `
        <script>
        console.error('Failed to load JavaScript file: ${filename}');
        document.addEventListener('DOMContentLoaded', function() {
          const errorDiv = document.createElement('div');
          errorDiv.innerHTML = '<div style="background: #f8d7da; color: #721c24; padding: 15px; border-radius: 6px; margin: 20px;">⚠️ JavaScript Loading Error: Could not load main JavaScript file.</div>';
          document.body.insertBefore(errorDiv, document.body.firstChild);
        });
        </script>
      `;
    }
    
    return `/* Error loading ${filename}: ${error.message} */`;
  }
}

// =================== SIMPLE AUTHENTICATION ===================

/**
 * Check login - Main authentication function
 */
function checkLogin(username, password) {
  try {
    console.log(`🔐 Checking login for: ${username}`);
    
    // Get Users sheet
    const sheet = getSheetByName(SHEET_NAMES.USERS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log('⚠️ No user data found');
      return { 
        success: false, 
        message: 'Chưa có dữ liệu người dùng trong hệ thống.' 
      };
    }
    
    // Find user in sheet
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // CỘT ĐÃ SỬA: Username=1, Password=5, Status=6
      
      const dbUsername = String(row[1] || '').trim();  // Cột B
      const dbStatus = String(row[6] || '').trim();    // Cột G - Status
      
      if (dbUsername === username && dbStatus === 'Active') {
        const dbPassword = String(row[5] || '').trim(); // Cột F - Password
        
        // Compare password (supports both plain text and hashed)
        if (isPasswordMatch(password, dbPassword)) {
          console.log(`✅ Login success for: ${username}`);
          
          return {
            success: true,
            user: {
              id: row[0],           // Cột A
              username: row[1],     // Cột B
              email: row[2] || '',  // Cột C
              fullName: row[3] || username, // Cột D
              roleId: row[4] || 'USER',     // Cột E
              status: row[6]        // Cột G
            }
          };
        }
      }
    }
    
    console.log(`❌ Login failed for: ${username}`);
    return { 
      success: false, 
      message: 'Tên đăng nhập hoặc mật khẩu không đúng' 
    };
    
  } catch (error) {
    console.error('❌ Error in checkLogin:', error);
    return { 
      success: false, 
      message: 'Lỗi hệ thống: ' + error.toString() 
    };
  }
}

/**
 * Compare password (supports both hash and plain text)
 */
function isPasswordMatch(inputPassword, dbPassword) {
  // If password in DB starts with hash prefix
  if (dbPassword.startsWith('HASH_') || dbPassword.length > 50) {
    return verifyPassword(inputPassword, dbPassword);
  }
  
  // If plain text, compare directly
  return inputPassword === dbPassword;
}

/**
 * Get all data needed for the app
 */
function getAllData() {
  try {
    console.log('📊 Loading all app data...');
    
    return {
      transactions: getTransactions(),
      accounts: getAccounts(), 
      categories: getCategories(),
      customers: getCustomers(),
      suppliers: getSuppliers()
    };
    
  } catch (error) {
    console.error('❌ Error loading app data:', error);
    return {
      transactions: [],
      accounts: [],
      categories: [],
      customers: [],
      suppliers: []
    };
  }
}

// =================== HELPER FUNCTIONS ===================

/**
 * Get active spreadsheet
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    console.error('❌ Error getting spreadsheet:', error);
    throw new Error('Không thể truy cập spreadsheet');
  }
}

/**
 * Get sheet by name, create if doesn't exist
 */
function getSheetByName(sheetName) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log(`📋 Creating new sheet: ${sheetName}`);
      sheet = ss.insertSheet(sheetName);
      initializeSheet(sheet, sheetName);
    }
    
    return sheet;
  } catch (error) {
    console.error(`❌ Error getting sheet ${sheetName}:`, error);
    throw new Error(`Không thể truy cập sheet ${sheetName}`);
  }
}

/**
 * Initialize sheet with headers based on sheet type
 */
function initializeSheet(sheet, sheetName) {
  const headers = getHeadersForSheet(sheetName);
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    console.log(`✅ Initialized sheet ${sheetName} with headers`);
  }
}

/**
 * Get appropriate headers for each sheet type
 */
function getHeadersForSheet(sheetName) {
  const headerMap = {
    [SHEET_NAMES.TRANSACTIONS]: ['ID', 'Ngày', 'Loại', 'Danh mục', 'Số tiền', 'Tài khoản', 'Ghi chú', 'Trạng thái'],
    [SHEET_NAMES.ACCOUNTS]: ['ID', 'Tên tài khoản', 'Loại', 'Số dư', 'Ngày tạo'],
    [SHEET_NAMES.CATEGORIES]: ['ID', 'Tên danh mục', 'Loại', 'Mô tả'],
    [SHEET_NAMES.CUSTOMERS]: ['ID', 'Tên', 'Điện thoại', 'Email', 'Địa chỉ', 'Ngày tạo'],
    [SHEET_NAMES.SUPPLIERS]: ['ID', 'Tên', 'Điện thoại', 'Email', 'Địa chỉ', 'Ngày tạo'],
    [SHEET_NAMES.USERS]: ['ID', 'Username', 'Email', 'FullName', 'RoleID', 'Password', 'Status', 'CreatedDate']
  };
  
  return headerMap[sheetName] || [];
}

// =================== ID GENERATION FUNCTIONS ===================

/**
 * Generate structured ID with prefix and auto-increment number
 */
function generateStructuredId(prefix, sheetName, idColumn = 0) {
  try {
    const sheet = getSheetByName(sheetName);
    if (!sheet) return prefix + '000001';
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return prefix + '000001';
    
    // Filter IDs with correct prefix
    const ids = data.slice(1)
                 .map(row => String(row[idColumn] || ''))
                 .filter(id => id.startsWith(prefix));
    
    if (ids.length === 0) return prefix + '000001';
    
    // Find max number
    let maxNumber = 0;
    ids.forEach(id => {
      const numberPart = id.substring(prefix.length);
      const number = parseInt(numberPart, 10);
      if (!isNaN(number) && number > maxNumber) {
        maxNumber = number;
      }
    });
    
    // Generate next ID
    const nextNumber = maxNumber + 1;
    const paddedNumber = String(nextNumber).padStart(6, '0');
    
    return prefix + paddedNumber;
  } catch (e) {
    console.error("Error generating structured ID: " + e.toString());
    return prefix + '000001';
  }
}

// ID generators for different entities
function generateTransactionId() {
  return generateStructuredId('GD', SHEET_NAMES.TRANSACTIONS);
}

function generateAccountId() {
  return generateStructuredId('TK', SHEET_NAMES.ACCOUNTS);
}

function generateCategoryId() {
  return generateStructuredId('DM', SHEET_NAMES.CATEGORIES);
}

function generateCustomerId() {
  return generateStructuredId('KH', SHEET_NAMES.CUSTOMERS);
}

function generateSupplierId() {
  return generateStructuredId('NCC', SHEET_NAMES.SUPPLIERS);
}

function generateUserId() {
  return generateStructuredId('USER', SHEET_NAMES.USERS);
}

// =================== UTILITY FUNCTIONS ===================

/**
 * Get current timestamp
 */
function getCurrentTimestamp() {
  return new Date().toISOString();
}

/**
 * Get current date in Vietnamese format
 */
function getCurrentDate() {
  return new Date().toLocaleDateString('vi-VN');
}

/**
 * Hash password (simple Base64 encoding for basic security)
 */
function hashPassword(password) {
  return Utilities.base64Encode(password + 'NHT_SALT_2024');
}

/**
 * Verify password
 */
function verifyPassword(password, hashedPassword) {
  return hashPassword(password) === hashedPassword;
}

/**
 * Handle API errors
 */
function handleApiError(error) {
  console.error('API Error:', error);
  return {
    success: false,
    message: 'Có lỗi xảy ra',
    error: error.toString()
  };
}

// =================== TRANSACTION FUNCTIONS ===================

/**
 * Get all transactions
 */
function getTransactions() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],
      date: row[1],
      type: row[2],
      category: row[3],
      amount: row[4],
      account: row[5],
      note: row[6],
      status: row[7]
    }));
  } catch (error) {
    console.error('Error getting transactions:', error);
    return [];
  }
}

/**
 * Add new transaction
 */
function addTransaction(transaction) {
  try {
    const sheet = getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const id = generateTransactionId();
    const timestamp = getCurrentDate();
    
    sheet.appendRow([
      id,
      transaction.date || timestamp,
      transaction.type,
      transaction.category,
      transaction.amount,
      transaction.account,
      transaction.note || '',
      transaction.status || 'Hoàn thành'
    ]);
    
    console.log(`✅ Added transaction: ${id}`);
    return { success: true, id: id, message: 'Thêm giao dịch thành công' };
  } catch (error) {
    console.error('Error adding transaction:', error);
    return handleApiError(error);
  }
}

/**
 * Update transaction
 */
function updateTransaction(transaction) {
  try {
    const sheet = getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === transaction.id) {
        sheet.getRange(i + 1, 1, 1, 8).setValues([[
          transaction.id,
          transaction.date,
          transaction.type,
          transaction.category,
          transaction.amount,
          transaction.account,
          transaction.note || '',
          transaction.status || 'Hoàn thành'
        ]]);
        console.log(`✅ Updated transaction: ${transaction.id}`);
        return { success: true, message: 'Cập nhật giao dịch thành công' };
      }
    }
    
    return { success: false, message: 'Không tìm thấy giao dịch' };
  } catch (error) {
    console.error('Error updating transaction:', error);
    return handleApiError(error);
  }
}

/**
 * Delete transaction
 */
function deleteTransaction(transactionId) {
  try {
    const sheet = getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === transactionId) {
        sheet.deleteRow(i + 1);
        console.log(`✅ Deleted transaction: ${transactionId}`);
        return { success: true, message: 'Xóa giao dịch thành công' };
      }
    }
    
    return { success: false, message: 'Không tìm thấy giao dịch' };
  } catch (error) {
    console.error('Error deleting transaction:', error);
    return handleApiError(error);
  }
}

// =================== ACCOUNT FUNCTIONS ===================

/**
 * Get all accounts
 */
function getAccounts() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.ACCOUNTS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],
      name: row[1],
      type: row[2],
      balance: row[3],
      createdDate: row[4]
    }));
  } catch (error) {
    console.error('Error getting accounts:', error);
    return [];
  }
}

/**
 * Add new account
 */
function addAccount(account) {
  try {
    const sheet = getSheetByName(SHEET_NAMES.ACCOUNTS);
    const id = generateAccountId();
    const timestamp = getCurrentDate();
    
    sheet.appendRow([
      id,
      account.name,
      account.type,
      account.balance || 0,
      timestamp
    ]);
    
    console.log(`✅ Added account: ${id}`);
    return { success: true, id: id, message: 'Thêm tài khoản thành công' };
  } catch (error) {
    console.error('Error adding account:', error);
    return handleApiError(error);
  }
}

// =================== CATEGORY FUNCTIONS ===================

/**
 * Get all categories
 */
function getCategories() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.CATEGORIES);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],
      name: row[1],
      type: row[2],
      description: row[3]
    }));
  } catch (error) {
    console.error('Error getting categories:', error);
    return [];
  }
}

/**
 * Add new category
 */
function addCategory(category) {
  try {
    const sheet = getSheetByName(SHEET_NAMES.CATEGORIES);
    const id = generateCategoryId();
    
    sheet.appendRow([
      id,
      category.name,
      category.type,
      category.description || ''
    ]);
    
    console.log(`✅ Added category: ${id}`);
    return { success: true, id: id, message: 'Thêm danh mục thành công' };
  } catch (error) {
    console.error('Error adding category:', error);
    return handleApiError(error);
  }
}

// =================== CUSTOMER FUNCTIONS ===================

/**
 * Get all customers
 */
function getCustomers() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.CUSTOMERS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],
      name: row[1],
      phone: row[2],
      email: row[3],
      address: row[4],
      createdDate: row[5]
    }));
  } catch (error) {
    console.error('Error getting customers:', error);
    return [];
  }
}

/**
 * Add new customer
 */
function addCustomer(customer) {
  try {
    const sheet = getSheetByName(SHEET_NAMES.CUSTOMERS);
    const id = generateCustomerId();
    const timestamp = getCurrentDate();
    
    sheet.appendRow([
      id,
      customer.name,
      customer.phone || '',
      customer.email || '',
      customer.address || '',
      timestamp
    ]);
    
    console.log(`✅ Added customer: ${id}`);
    return { success: true, id: id, message: 'Thêm khách hàng thành công' };
  } catch (error) {
    console.error('Error adding customer:', error);
    return handleApiError(error);
  }
}

// =================== SUPPLIER FUNCTIONS ===================

/**
 * Get all suppliers
 */
function getSuppliers() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.SUPPLIERS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],
      name: row[1],
      phone: row[2],
      email: row[3],
      address: row[4],
      createdDate: row[5]
    }));
  } catch (error) {
    console.error('Error getting suppliers:', error);
    return [];
  }
}

/**
 * Add new supplier
 */
function addSupplier(supplier) {
  try {
    const sheet = getSheetByName(SHEET_NAMES.SUPPLIERS);
    const id = generateSupplierId();
    const timestamp = getCurrentDate();
    
    sheet.appendRow([
      id,
      supplier.name,
      supplier.phone || '',
      supplier.email || '',
      supplier.address || '',
      timestamp
    ]);
    
    console.log(`✅ Added supplier: ${id}`);
    return { success: true, id: id, message: 'Thêm nhà cung cấp thành công' };
  } catch (error) {
    console.error('Error adding supplier:', error);
    return handleApiError(error);
  }
}

// =================== DASHBOARD & STATISTICS FUNCTIONS ===================

/**
 * Get dashboard statistics
 */
function getDashboardStats() {
  try {
    const transactions = getTransactions();
    const accounts = getAccounts();
    
    // Calculate totals
    let totalIncome = 0;
    let totalExpense = 0;
    let monthlyIncome = 0;
    let monthlyExpense = 0;
    
    const currentMonth = new Date().getMonth();
    const currentYear = new Date().getFullYear();
    
    transactions.forEach(transaction => {
      const amount = parseFloat(transaction.amount) || 0;
      const transactionDate = new Date(transaction.date);
      
      if (transaction.type === 'Thu') {
        totalIncome += amount;
        if (transactionDate.getMonth() === currentMonth && 
            transactionDate.getFullYear() === currentYear) {
          monthlyIncome += amount;
        }
      } else if (transaction.type === 'Chi') {
        totalExpense += amount;
        if (transactionDate.getMonth() === currentMonth && 
            transactionDate.getFullYear() === currentYear) {
          monthlyExpense += amount;
        }
      }
    });
    
    // Calculate total balance
    let totalBalance = 0;
    accounts.forEach(account => {
      totalBalance += parseFloat(account.balance) || 0;
    });
    
    return {
      totalIncome,
      totalExpense,
      totalBalance,
      monthlyIncome,
      monthlyExpense,
      monthlyBalance: monthlyIncome - monthlyExpense,
      transactionCount: transactions.length,
      accountCount: accounts.length
    };
  } catch (error) {
    console.error('Error getting dashboard stats:', error);
    return {
      totalIncome: 0,
      totalExpense: 0,
      totalBalance: 0,
      monthlyIncome: 0,
      monthlyExpense: 0,
      monthlyBalance: 0,
      transactionCount: 0,
      accountCount: 0
    };
  }
}

// =================== TESTING & DEBUG FUNCTIONS ===================

/**
 * Test function to verify system is working
 */
function testSystem() {
  try {
    console.log('🧪 Testing NHT Finance System...');
    
    const tests = {
      spreadsheet: false,
      sheets: false,
      data: false
    };
    
    // Test spreadsheet access
    try {
      const ss = getSpreadsheet();
      tests.spreadsheet = true;
      console.log('✅ Spreadsheet access: OK');
    } catch (error) {
      console.error('❌ Spreadsheet access: FAILED');
    }
    
    // Test sheet access
    try {
      const usersSheet = getSheetByName(SHEET_NAMES.USERS);
      tests.sheets = true;
      console.log('✅ Sheet access: OK');
    } catch (error) {
      console.error('❌ Sheet access: FAILED');
    }
    
    // Test data access
    try {
      const data = getAllData();
      tests.data = true;
      console.log('✅ Data access: OK');
    } catch (error) {
      console.error('❌ Data access: FAILED');
    }
    
    const overallSuccess = Object.values(tests).every(test => test === true);
    
    console.log('🎯 System test completed:', overallSuccess ? 'PASSED' : 'FAILED');
    
    return {
      success: overallSuccess,
      tests: tests,
      message: overallSuccess ? 'All tests passed' : 'Some tests failed'
    };
    
  } catch (error) {
    console.error('❌ System test error:', error);
    return handleApiError(error);
  }
}

/**
 * Get system info
 */
function getSystemInfo() {
  try {
    const ss = getSpreadsheet();
    
    return {
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      owner: ss.getOwner().getEmail(),
      lastModified: ss.getLastUpdated(),
      timezone: ss.getSpreadsheetTimeZone(),
      locale: ss.getSpreadsheetLocale(),
      sheetCount: ss.getSheets().length,
      version: '2.0.0 - Fixed Sheet Names'
    };
  } catch (error) {
    console.error('Error getting system info:', error);
    return { error: error.toString() };
  }
}

/**
 * DEBUG: Get all sheet names in the spreadsheet
 */
function getAllSheetNames() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const sheetNames = sheets.map(sheet => sheet.getName());
    
    console.log('📋 All sheet names in spreadsheet:', sheetNames);
    return {
      success: true,
      sheetNames: sheetNames,
      count: sheetNames.length
    };
  } catch (error) {
    console.error('Error getting sheet names:', error);
    return handleApiError(error);
  }
}

// =================== END OF FILE ===================

console.log('✅ NHT_Code.gs loaded successfully with FIXED SHEET NAMES!');
console.log('🔧 Remember to update SHEET_NAMES constant with your actual sheet names');
console.log('📋 Use getAllSheetNames() to see all sheets in your spreadsheet');
