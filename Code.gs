// =================== NHT_Code.gs - FIXED FOR YOUR EXCEL STRUCTURE ===================

/**
 * CẤU HÌNH TÊN SHEET - KHỚP VỚI FILE EXCEL CỦA BẠN
 */
const SHEET_NAMES = {
  TRANSACTIONS: 'Transactions',
  ACCOUNTS: 'Accounts',
  CATEGORIES: 'Categories',
  CUSTOMERS: 'Customers',
  SUPPLIERS: 'Suppliers',
  USERS: 'Users'
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
    
    // Find user in sheet - KHỚP VỚI CẤU TRÚC EXCEL CỦA BẠN
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Cấu trúc: ID(0), Username(1), Email(2), FullName(3), RoleID(4), Password(5), Status(6)
      
      const dbUsername = String(row[1] || '').trim();  // Cột B - Username
      const dbStatus = String(row[6] || '').trim();    // Cột G - Status
      
      if (dbUsername === username && dbStatus === 'Active') {
        const dbPassword = String(row[5] || '').trim(); // Cột F - Password
        
        // Compare password (supports both plain text and hashed)
        if (isPasswordMatch(password, dbPassword)) {
          console.log(`✅ Login success for: ${username}`);
          
          return {
            success: true,
            user: {
              id: row[0],           // Cột A - ID
              username: row[1],     // Cột B - Username
              email: row[2] || '',  // Cột C - Email
              fullName: row[3] || username, // Cột D - FullName
              roleId: row[4] || 'USER',     // Cột E - RoleID
              status: row[6]        // Cột G - Status
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
 * Get appropriate headers for each sheet type - KHỚP VỚI FILE EXCEL
 */
function getHeadersForSheet(sheetName) {
  const headerMap = {
    [SHEET_NAMES.TRANSACTIONS]: ['ID', 'Ngày', 'Loại', 'Danh mục', 'Số tiền', 'Tài khoản nguồn', 'Tài khoản đích', 'Ghi chú', 'Số hóa đơn', 'Ngày hóa đơn', 'Tên đối tượng', 'Loại đối tượng', 'Nhân viên/Bộ phận', 'Trạng thái thanh toán', 'Ngày đến hạn', 'Ngày thanh toán'],
    [SHEET_NAMES.ACCOUNTS]: ['ID', 'Tên', 'Loại', 'Số dư đầu kỳ', 'Số dư hiện tại', 'Icon', 'Thông tin ngân hàng', 'Số tài khoản'],
    [SHEET_NAMES.CATEGORIES]: ['ID', 'Tên', 'Loại', 'Icon'],
    [SHEET_NAMES.CUSTOMERS]: ['ID', 'Tên', 'Số điện thoại', 'Email', 'Địa chỉ', 'Mã số thuế', 'Người liên hệ', 'Ghi chú', 'Số dư công nợ'],
    [SHEET_NAMES.SUPPLIERS]: ['ID', 'Tên', 'Số điện thoại', 'Email', 'Địa chỉ', 'Mã số thuế', 'Người liên hệ', 'Ghi chú', 'Số dư công nợ'],
    [SHEET_NAMES.USERS]: ['ID', 'Username', 'Email', 'FullName', 'RoleID', 'Password', 'Status']
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
  return generateStructuredId('NV', SHEET_NAMES.USERS);
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

// =================== TRANSACTION FUNCTIONS - FIXED FOR YOUR STRUCTURE ===================

/**
 * Get all transactions - KHỚP VỚI CẤU TRÚC 16 CỘT CỦA BẠN
 */
function getTransactions() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],                    // ID
      date: row[1],                  // Ngày
      type: row[2],                  // Loại
      category: row[3],              // Danh mục
      amount: row[4],                // Số tiền
      account: row[5],               // Tài khoản nguồn
      targetAccount: row[6],         // Tài khoản đích
      note: row[7],                  // Ghi chú
      invoiceNumber: row[8],         // Số hóa đơn
      invoiceDate: row[9],           // Ngày hóa đơn
      objectName: row[10],           // Tên đối tượng
      objectType: row[11],           // Loại đối tượng
      employee: row[12],             // Nhân viên/Bộ phận
      status: row[13],               // Trạng thái thanh toán
      dueDate: row[14],              // Ngày đến hạn
      paymentDate: row[15]           // Ngày thanh toán
    }));
  } catch (error) {
    console.error('Error getting transactions:', error);
    return [];
  }
}

/**
 * Add new transaction - CẬP NHẬT THEO CẤU TRÚC MỚI
 */
function addTransaction(transaction) {
  try {
    const sheet = getSheetByName(SHEET_NAMES.TRANSACTIONS);
    const id = generateTransactionId();
    const timestamp = getCurrentDate();
    
    sheet.appendRow([
      id,                                    // ID
      transaction.date || timestamp,         // Ngày
      transaction.type,                      // Loại
      transaction.category,                  // Danh mục
      transaction.amount,                    // Số tiền
      transaction.account,                   // Tài khoản nguồn
      transaction.targetAccount || '',       // Tài khoản đích
      transaction.note || '',                // Ghi chú
      transaction.invoiceNumber || '',       // Số hóa đơn
      transaction.invoiceDate || '',         // Ngày hóa đơn
      transaction.objectName || '',          // Tên đối tượng
      transaction.objectType || '',          // Loại đối tượng
      transaction.employee || '',            // Nhân viên/Bộ phận
      transaction.status || 'Chưa thanh toán', // Trạng thái thanh toán
      transaction.dueDate || '',             // Ngày đến hạn
      transaction.paymentDate || ''          // Ngày thanh toán
    ]);
    
    console.log(`✅ Added transaction: ${id}`);
    return { success: true, id: id, message: 'Thêm giao dịch thành công' };
  } catch (error) {
    console.error('Error adding transaction:', error);
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

// =================== ACCOUNT FUNCTIONS - FIXED FOR YOUR STRUCTURE ===================

/**
 * Get all accounts - KHỚP VỚI CẤU TRÚC 8 CỘT CỦA BẠN
 */
function getAccounts() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.ACCOUNTS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],                    // ID
      name: row[1],                  // Tên
      type: row[2],                  // Loại
      initialBalance: row[3],        // Số dư đầu kỳ
      balance: row[4],               // Số dư hiện tại
      icon: row[5],                  // Icon
      bankInfo: row[6],              // Thông tin ngân hàng
      accountNumber: row[7]          // Số tài khoản
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
    
    sheet.appendRow([
      id,
      account.name,
      account.type,
      account.initialBalance || 0,
      account.balance || account.initialBalance || 0,
      account.icon || '💰',
      account.bankInfo || '',
      account.accountNumber || ''
    ]);
    
    console.log(`✅ Added account: ${id}`);
    return { success: true, id: id, message: 'Thêm tài khoản thành công' };
  } catch (error) {
    console.error('Error adding account:', error);
    return handleApiError(error);
  }
}

// =================== CATEGORY FUNCTIONS - FIXED FOR YOUR STRUCTURE ===================

/**
 * Get all categories - KHỚP VỚI CẤU TRÚC 4 CỘT CỦA BẠN
 */
function getCategories() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.CATEGORIES);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],        // ID
      name: row[1],      // Tên
      type: row[2],      // Loại
      icon: row[3]       // Icon
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
      category.icon || '📁'
    ]);
    
    console.log(`✅ Added category: ${id}`);
    return { success: true, id: id, message: 'Thêm danh mục thành công' };
  } catch (error) {
    console.error('Error adding category:', error);
    return handleApiError(error);
  }
}

// =================== CUSTOMER FUNCTIONS - FIXED FOR YOUR STRUCTURE ===================

/**
 * Get all customers - KHỚP VỚI CẤU TRÚC 9 CỘT CỦA BẠN
 */
function getCustomers() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.CUSTOMERS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],          // ID
      name: row[1],        // Tên
      phone: row[2],       // Số điện thoại
      email: row[3],       // Email
      address: row[4],     // Địa chỉ
      taxCode: row[5],     // Mã số thuế
      contact: row[6],     // Người liên hệ
      note: row[7],        // Ghi chú
      balance: row[8]      // Số dư công nợ
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
    
    sheet.appendRow([
      id,
      customer.name,
      customer.phone || '',
      customer.email || '',
      customer.address || '',
      customer.taxCode || '',
      customer.contact || '',
      customer.note || '',
      customer.balance || 0
    ]);
    
    console.log(`✅ Added customer: ${id}`);
    return { success: true, id: id, message: 'Thêm khách hàng thành công' };
  } catch (error) {
    console.error('Error adding customer:', error);
    return handleApiError(error);
  }
}

// =================== SUPPLIER FUNCTIONS - FIXED FOR YOUR STRUCTURE ===================

/**
 * Get all suppliers - KHỚP VỚI CẤU TRÚC 9 CỘT CỦA BẠN
 */
function getSuppliers() {
  try {
    const sheet = getSheetByName(SHEET_NAMES.SUPPLIERS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],          // ID
      name: row[1],        // Tên
      phone: row[2],       // Số điện thoại
      email: row[3],       // Email
      address: row[4],     // Địa chỉ
      taxCode: row[5],     // Mã số thuế
      contact: row[6],     // Người liên hệ
      note: row[7],        // Ghi chú
      balance: row[8]      // Số dư công nợ
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
    
    sheet.appendRow([
      id,
      supplier.name,
      supplier.phone || '',
      supplier.email || '',
      supplier.address || '',
      supplier.taxCode || '',
      supplier.contact || '',
      supplier.note || '',
      supplier.balance || 0
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
      
      // Phân loại theo loại giao dịch trong Excel của bạn
      if (transaction.type === 'Doanh thu' || transaction.type === 'Thu') {
        totalIncome += amount;
        if (transactionDate.getMonth() === currentMonth && 
            transactionDate.getFullYear() === currentYear) {
          monthlyIncome += amount;
        }
      } else if (transaction.type === 'Chi phí' || transaction.type === 'Chi') {
        totalExpense += amount;
        if (transactionDate.getMonth() === currentMonth && 
            transactionDate.getFullYear() === currentYear) {
          monthlyExpense += amount;
        }
      }
    });
    
    // Calculate total balance using current balance from accounts
    let totalBalance = 0;
    accounts.forEach(account => {
      totalBalance += parseFloat(account.balance) || 0; // Sử dụng "Số dư hiện tại"
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
      console.log('📊 Data summary:');
      console.log(`- Transactions: ${data.transactions.length}`);
      console.log(`- Accounts: ${data.accounts.length}`);
      console.log(`- Categories: ${data.categories.length}`);
      console.log(`- Customers: ${data.customers.length}`);
      console.log(`- Suppliers: ${data.suppliers.length}`);
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
      version: '2.1.0 - Fixed for Excel Structure'
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

/**
 * DEBUG: Test data loading with detailed output
 */
function testDataLoad() {
  try {
    console.log('🔍 Testing data load...');
    
    const results = {};
    
    // Test Transactions
    console.log('\n--- Testing Transactions ---');
    try {
      const transactions = getTransactions();
      results.transactions = {
        success: true,
        count: transactions.length,
        sample: transactions.length > 0 ? transactions[0] : null
      };
      console.log(`✅ Transactions: ${transactions.length} records`);
      if (transactions.length > 0) {
        console.log('📝 Sample transaction:', transactions[0]);
      }
    } catch (error) {
      results.transactions = { success: false, error: error.toString() };
      console.log('❌ Transactions: FAILED');
    }
    
    // Test Accounts
    console.log('\n--- Testing Accounts ---');
    try {
      const accounts = getAccounts();
      results.accounts = {
        success: true,
        count: accounts.length,
        sample: accounts.length > 0 ? accounts[0] : null
      };
      console.log(`✅ Accounts: ${accounts.length} records`);
      if (accounts.length > 0) {
        console.log('📝 Sample account:', accounts[0]);
      }
    } catch (error) {
      results.accounts = { success: false, error: error.toString() };
      console.log('❌ Accounts: FAILED');
    }
    
    // Test Categories
    console.log('\n--- Testing Categories ---');
    try {
      const categories = getCategories();
      results.categories = {
        success: true,
        count: categories.length,
        sample: categories.length > 0 ? categories[0] : null
      };
      console.log(`✅ Categories: ${categories.length} records`);
      if (categories.length > 0) {
        console.log('📝 Sample category:', categories[0]);
      }
    } catch (error) {
      results.categories = { success: false, error: error.toString() };
      console.log('❌ Categories: FAILED');
    }
    
    // Test Customers
    console.log('\n--- Testing Customers ---');
    try {
      const customers = getCustomers();
      results.customers = {
        success: true,
        count: customers.length,
        sample: customers.length > 0 ? customers[0] : null
      };
      console.log(`✅ Customers: ${customers.length} records`);
    } catch (error) {
      results.customers = { success: false, error: error.toString() };
      console.log('❌ Customers: FAILED');
    }
    
    // Test Suppliers
    console.log('\n--- Testing Suppliers ---');
    try {
      const suppliers = getSuppliers();
      results.suppliers = {
        success: true,
        count: suppliers.length,
        sample: suppliers.length > 0 ? suppliers[0] : null
      };
      console.log(`✅ Suppliers: ${suppliers.length} records`);
    } catch (error) {
      results.suppliers = { success: false, error: error.toString() };
      console.log('❌ Suppliers: FAILED');
    }
    
    return {
      success: true,
      results: results,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ Error in testDataLoad:', error);
    return handleApiError(error);
  }
}

// =================== END OF FILE ===================

console.log('✅ NHT_Code.gs loaded successfully - FIXED FOR YOUR EXCEL STRUCTURE!');
console.log('📊 Main changes:');
console.log('- Fixed sheet names: Transactions, Accounts, Categories, etc.');
console.log('- Fixed column mapping for all sheets');
console.log('- Fixed transaction types: "Doanh thu" and "Chi phí"');
console.log('- Updated account balance field');
console.log('🧪 Use testDataLoad() to verify data loading');
console.log('🔐 Use testSystem() to test full system');
