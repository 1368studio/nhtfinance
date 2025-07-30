// =================== NHT_Code.gs - PHIÊN BẢN HOÀN CHỈNH ===================

/**
 * Main entry point for the web application
 * Serves the main HTML file with proper configuration
 */
function doGet() {
  console.log('🚀 Starting NHT Finance System...');
  
  try {
    // Initialize system if needed (chỉ chạy 1 lần)
    const usersSheet = getSheetByName('Users');
    if (usersSheet) {
      const userData = usersSheet.getDataRange().getValues();
      
      if (userData.length <= 1) {
        console.log('🆕 First run detected, initializing sample data...');
        initializeSampleData();
      }
    }
    
    // Return HTML with correct file names
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
 * FIXED: This function handles the template inclusion properly
 */
function include(filename) {
  try {
    console.log(`📄 Including file: ${filename}`);
    
    // Map old names to new names for backward compatibility
    const fileMap = {
      'Stylesheet': 'Stylesheet',
      'JavaScript': 'JavaScript'
    };
    
    const actualFilename = fileMap[filename] || filename;
    const content = HtmlService.createHtmlOutputFromFile(actualFilename).getContent();
    
    console.log(`✅ Successfully included file: ${actualFilename}`);
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
        .sidebar {
          width: 250px;
          background: #343a40;
          color: white;
          position: fixed;
          height: 100vh;
          padding: 20px;
        }
        .main-content {
          margin-left: 250px;
          padding: 20px;
        }
        .modal {
          display: none;
          position: fixed;
          z-index: 1000;
          left: 0;
          top: 0;
          width: 100%;
          height: 100%;
          background-color: rgba(0,0,0,0.5);
        }
        .modal-content {
          background: white;
          margin: 10% auto;
          padding: 20px;
          border-radius: 8px;
          width: 90%;
          max-width: 500px;
        }
      `;
    } else if (filename.includes('JavaScript') || filename === 'JavaScript') {
      return `
        <script>
        // Error loading ${filename}: ${error.message}
        console.error('Failed to load JavaScript file: ${filename}');
        
        // Emergency fallback functionality
        document.addEventListener('DOMContentLoaded', function() {
          console.log('🚨 Emergency mode: Basic functionality only');
          
          // Show error message
          const errorDiv = document.createElement('div');
          errorDiv.innerHTML = \`
            <div style="background: #f8d7da; color: #721c24; padding: 15px; border-radius: 6px; margin: 20px; border: 1px solid #f5c6cb;">
              <h4>⚠️ JavaScript Loading Error</h4>
              <p>Could not load main JavaScript file. The application is running in limited mode.</p>
              <button onclick="window.location.reload()" style="padding: 8px 16px; background: #dc3545; color: white; border: none; border-radius: 4px; cursor: pointer;">
                🔄 Reload Page
              </button>
            </div>
          \`;
          
          const content = document.querySelector('.content');
          if (content) {
            content.insertBefore(errorDiv, content.firstChild);
          } else {
            document.body.insertBefore(errorDiv, document.body.firstChild);
          }
        });
        </script>
      `;
    }
    
    return `/* Error loading ${filename}: ${error.message} */`;
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
    'Giao dịch': ['ID', 'Ngày', 'Loại', 'Danh mục', 'Số tiền', 'Tài khoản', 'Ghi chú', 'Trạng thái'],
    'Tài khoản': ['ID', 'Tên tài khoản', 'Loại', 'Số dư', 'Ngày tạo'],
    'Danh mục': ['ID', 'Tên danh mục', 'Loại', 'Mô tả'],
    'Khách hàng': ['ID', 'Tên', 'Điện thoại', 'Email', 'Địa chỉ', 'Ngày tạo'],
    'Nhà cung cấp': ['ID', 'Tên', 'Điện thoại', 'Email', 'Địa chỉ', 'Ngày tạo'],
    'Sản phẩm': ['ID', 'Tên sản phẩm', 'Mã SP', 'Giá', 'Tồn kho', 'Danh mục'],
    'Users': ['ID', 'Username', 'Email', 'FullName', 'RoleID', 'CreatedDate', 'Password', 'Status'],
    'Roles': ['RoleID', 'RoleName', 'Description'],
    'Permissions': ['RoleID', 'MenuID', 'CanView', 'CanAdd', 'CanEdit', 'CanDelete'],
    'Menus': ['MenuID', 'MenuName', 'Description', 'URL', 'Icon']
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
  return generateStructuredId('GD', 'Giao dịch');
}

function generateAccountId() {
  return generateStructuredId('TK', 'Tài khoản');
}

function generateCategoryId() {
  return generateStructuredId('DM', 'Danh mục');
}

function generateCustomerId() {
  return generateStructuredId('KH', 'Khách hàng');
}

function generateSupplierId() {
  return generateStructuredId('NCC', 'Nhà cung cấp');
}

function generateProductId() {
  return generateStructuredId('SP', 'Sản phẩm');
}

function generateUserId() {
  return generateStructuredId('USER', 'Users');
}

// =================== UTILITY FUNCTIONS ===================

/**
 * Format currency for display
 */
function formatCurrency(amount) {
  try {
    return new Intl.NumberFormat('vi-VN', {
      style: 'currency',
      currency: 'VND'
    }).format(amount || 0);
  } catch (error) {
    return (amount || 0).toLocaleString('vi-VN') + ' đ';
  }
}

/**
 * Generate unique ID
 */
function generateId() {
  return 'ID_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

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
 * Validate email format
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email);
}

/**
 * Validate password strength
 */
function isValidPassword(password) {
  if (!password || password.length < 6) return false;
  return true;
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
    const sheet = getSheetByName('Giao dịch');
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
    const sheet = getSheetByName('Giao dịch');
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
    return { success: true, id: id };
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
    const sheet = getSheetByName('Giao dịch');
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
        return { success: true };
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
    const sheet = getSheetByName('Giao dịch');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === transactionId) {
        sheet.deleteRow(i + 1);
        console.log(`✅ Deleted transaction: ${transactionId}`);
        return { success: true };
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
    const sheet = getSheetByName('Tài khoản');
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
    const sheet = getSheetByName('Tài khoản');
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
    return { success: true, id: id };
  } catch (error) {
    console.error('Error adding account:', error);
    return handleApiError(error);
  }
}

/**
 * Update account
 */
function updateAccount(account) {
  try {
    const sheet = getSheetByName('Tài khoản');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === account.id) {
        sheet.getRange(i + 1, 1, 1, 5).setValues([[
          account.id,
          account.name,
          account.type,
          account.balance || 0,
          data[i][4] // Keep original creation date
        ]]);
        console.log(`✅ Updated account: ${account.id}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy tài khoản' };
  } catch (error) {
    console.error('Error updating account:', error);
    return handleApiError(error);
  }
}

/**
 * Delete account
 */
function deleteAccount(accountId) {
  try {
    const sheet = getSheetByName('Tài khoản');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === accountId) {
        sheet.deleteRow(i + 1);
        console.log(`✅ Deleted account: ${accountId}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy tài khoản' };
  } catch (error) {
    console.error('Error deleting account:', error);
    return handleApiError(error);
  }
}

// =================== CATEGORY FUNCTIONS ===================

/**
 * Get all categories
 */
function getCategories() {
  try {
    const sheet = getSheetByName('Danh mục');
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
    const sheet = getSheetByName('Danh mục');
    const id = generateCategoryId();
    
    sheet.appendRow([
      id,
      category.name,
      category.type,
      category.description || ''
    ]);
    
    console.log(`✅ Added category: ${id}`);
    return { success: true, id: id };
  } catch (error) {
    console.error('Error adding category:', error);
    return handleApiError(error);
  }
}

/**
 * Update category
 */
function updateCategory(category) {
  try {
    const sheet = getSheetByName('Danh mục');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === category.id) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[
          category.id,
          category.name,
          category.type,
          category.description || ''
        ]]);
        console.log(`✅ Updated category: ${category.id}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy danh mục' };
  } catch (error) {
    console.error('Error updating category:', error);
    return handleApiError(error);
  }
}

/**
 * Delete category
 */
function deleteCategory(categoryId) {
  try {
    const sheet = getSheetByName('Danh mục');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === categoryId) {
        sheet.deleteRow(i + 1);
        console.log(`✅ Deleted category: ${categoryId}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy danh mục' };
  } catch (error) {
    console.error('Error deleting category:', error);
    return handleApiError(error);
  }
}

// =================== CUSTOMER FUNCTIONS ===================

/**
 * Get all customers
 */
function getCustomers() {
  try {
    const sheet = getSheetByName('Khách hàng');
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
    const sheet = getSheetByName('Khách hàng');
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
    return { success: true, id: id };
  } catch (error) {
    console.error('Error adding customer:', error);
    return handleApiError(error);
  }
}

/**
 * Update customer
 */
function updateCustomer(customer) {
  try {
    const sheet = getSheetByName('Khách hàng');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === customer.id) {
        sheet.getRange(i + 1, 1, 1, 6).setValues([[
          customer.id,
          customer.name,
          customer.phone || '',
          customer.email || '',
          customer.address || '',
          data[i][5] // Keep original creation date
        ]]);
        console.log(`✅ Updated customer: ${customer.id}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy khách hàng' };
  } catch (error) {
    console.error('Error updating customer:', error);
    return handleApiError(error);
  }
}

/**
 * Delete customer
 */
function deleteCustomer(customerId) {
  try {
    const sheet = getSheetByName('Khách hàng');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === customerId) {
        sheet.deleteRow(i + 1);
        console.log(`✅ Deleted customer: ${customerId}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy khách hàng' };
  } catch (error) {
    console.error('Error deleting customer:', error);
    return handleApiError(error);
  }
}

// =================== SUPPLIER FUNCTIONS ===================

/**
 * Get all suppliers
 */
function getSuppliers() {
  try {
    const sheet = getSheetByName('Nhà cung cấp');
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
    const sheet = getSheetByName('Nhà cung cấp');
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
    return { success: true, id: id };
  } catch (error) {
    console.error('Error adding supplier:', error);
    return handleApiError(error);
  }
}

/**
 * Update supplier
 */
function updateSupplier(supplier) {
  try {
    const sheet = getSheetByName('Nhà cung cấp');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === supplier.id) {
        sheet.getRange(i + 1, 1, 1, 6).setValues([[
          supplier.id,
          supplier.name,
          supplier.phone || '',
          supplier.email || '',
          supplier.address || '',
          data[i][5] // Keep original creation date
        ]]);
        console.log(`✅ Updated supplier: ${supplier.id}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy nhà cung cấp' };
  } catch (error) {
    console.error('Error updating supplier:', error);
    return handleApiError(error);
  }
}

/**
 * Delete supplier
 */
function deleteSupplier(supplierId) {
  try {
    const sheet = getSheetByName('Nhà cung cấp');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === supplierId) {
        sheet.deleteRow(i + 1);
        console.log(`✅ Deleted supplier: ${supplierId}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy nhà cung cấp' };
  } catch (error) {
    console.error('Error deleting supplier:', error);
    return handleApiError(error);
  }
}

// =================== USER MANAGEMENT FUNCTIONS ===================

/**
 * Get all users
 */
function getUsers() {
  try {
    const sheet = getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      id: row[0],
      username: row[1],
      email: row[2],
      fullName: row[3],
      roleId: row[4],
      createdDate: row[5],
      status: row[7] || 'Active'
    }));
  } catch (error) {
    console.error('Error getting users:', error);
    return [];
  }
}

/**
 * Authenticate user
 */
function authenticateUser(username, password) {
  try {
    const sheet = getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username && data[i][7] === 'Active') {
        const hashedPassword = data[i][6];
        if (verifyPassword(password, hashedPassword)) {
          return {
            success: true,
            user: {
              id: data[i][0],
              username: data[i][1],
              email: data[i][2],
              fullName: data[i][3],
              roleId: data[i][4]
            }
          };
        }
      }
    }
    
    return { success: false, message: 'Tên đăng nhập hoặc mật khẩu không đúng' };
  } catch (error) {
    console.error('Error authenticating user:', error);
    return handleApiError(error);
  }
}

/**
 * Add new user
 */
function addUser(user) {
  try {
    const sheet = getSheetByName('Users');
    const id = generateUserId();
    const timestamp = getCurrentDate();
    
    // Validate email
    if (user.email && !isValidEmail(user.email)) {
      return { success: false, message: 'Email không hợp lệ' };
    }
    
    // Validate password
    if (!isValidPassword(user.password)) {
      return { success: false, message: 'Mật khẩu phải có ít nhất 6 ký tự' };
    }
    
    // Check if username already exists
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][1] === user.username) {
        return { success: false, message: 'Tên đăng nhập đã tồn tại' };
      }
    }
    
    sheet.appendRow([
      id,
      user.username,
      user.email || '',
      user.fullName || '',
      user.roleId || 'USER',
      timestamp,
      hashPassword(user.password),
      user.status || 'Active'
    ]);
    
    console.log(`✅ Added user: ${id}`);
    return { success: true, id: id };
  } catch (error) {
    console.error('Error adding user:', error);
    return handleApiError(error);
  }
}

/**
 * Update user
 */
function updateUser(user) {
  try {
    const sheet = getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === user.id) {
        // Validate email if provided
        if (user.email && !isValidEmail(user.email)) {
          return { success: false, message: 'Email không hợp lệ' };
        }
        
        const updatedRow = [
          user.id,
          user.username || data[i][1],
          user.email || data[i][2],
          user.fullName || data[i][3],
          user.roleId || data[i][4],
          data[i][5], // Keep original creation date
          user.password ? hashPassword(user.password) : data[i][6], // Only update password if provided
          user.status || data[i][7]
        ];
        
        sheet.getRange(i + 1, 1, 1, 8).setValues([updatedRow]);
        console.log(`✅ Updated user: ${user.id}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy người dùng' };
  } catch (error) {
    console.error('Error updating user:', error);
    return handleApiError(error);
  }
}

/**
 * Delete user
 */
function deleteUser(userId) {
  try {
    const sheet = getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userId) {
        sheet.deleteRow(i + 1);
        console.log(`✅ Deleted user: ${userId}`);
        return { success: true };
      }
    }
    
    return { success: false, message: 'Không tìm thấy người dùng' };
  } catch (error) {
    console.error('Error deleting user:', error);
    return handleApiError(error);
  }
}

/**
 * Get roles
 */
function getRoles() {
  try {
    const sheet = getSheetByName('Roles');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      roleId: row[0],
      roleName: row[1],
      description: row[2]
    }));
  } catch (error) {
    console.error('Error getting roles:', error);
    return [];
  }
}

/**
 * Get permissions by role
 */
function getPermissionsByRole(roleId) {
  try {
    const permissionsSheet = getSheetByName('Permissions');
    const menusSheet = getSheetByName('Menus');
    
    const permissionsData = permissionsSheet.getDataRange().getValues();
    const menusData = menusSheet.getDataRange().getValues();
    
    let permissionsMap = {};
    const lowerRoleId = roleId.toLowerCase();
    
    // Map existing permissions for the role
    for (let i = 1; i < permissionsData.length; i++) {
      if (permissionsData[i][0].toLowerCase() === lowerRoleId) {
        permissionsMap[permissionsData[i][1]] = {
          roleId: permissionsData[i][0],
          menuId: permissionsData[i][1],
          canView: String(permissionsData[i][2]).toLowerCase() === 'true',
          canAdd: String(permissionsData[i][3]).toLowerCase() === 'true',
          canEdit: String(permissionsData[i][4]).toLowerCase() === 'true',
          canDelete: String(permissionsData[i][5]).toLowerCase() === 'true'
        };
      }
    }
    
    let permissions = [];
    
    // Ensure all menus are included
    for (let i = 1; i < menusData.length; i++) {
      let menuId = menusData[i][0];
      
      if (permissionsMap[menuId]) {
        permissions.push(permissionsMap[menuId]);
      } else {
        permissions.push({
          roleId: roleId,
          menuId: menuId,
          canView: false,
          canAdd: false,
          canEdit: false,
          canDelete: false
        });
      }
    }
    
    return permissions;
  } catch (error) {
    console.error('Error getting permissions by role:', error);
    return [];
  }
}

/**
 * Update permissions
 */
function updatePermissions(permissions) {
  try {
    const sheet = getSheetByName('Permissions');
    const data = sheet.getDataRange().getValues();
    
    let lastRow = data.length;
    
    for (let i = 0; i < permissions.length; i++) {
      let permissionRoleId = permissions[i].roleId.toLowerCase();
      let permissionMenuId = permissions[i].menuId.toLowerCase();
      let rowIndex = data.findIndex(row => 
        row[0].toLowerCase() === permissionRoleId && 
        row[1].toLowerCase() === permissionMenuId
      );
      
      if (rowIndex !== -1) {
        // Update existing permission
        sheet.getRange(rowIndex + 1, 3, 1, 4).setValues([[
          permissions[i].canView ? 'TRUE' : 'FALSE',
          permissions[i].canAdd ? 'TRUE' : 'FALSE',
          permissions[i].canEdit ? 'TRUE' : 'FALSE',
          permissions[i].canDelete ? 'TRUE' : 'FALSE'
        ]]);
      } else {
        // Add new permission
        lastRow++;
        sheet.getRange(lastRow, 1, 1, 6).setValues([[
          permissions[i].roleId,
          permissions[i].menuId,
          permissions[i].canView ? 'TRUE' : 'FALSE',
          permissions[i].canAdd ? 'TRUE' : 'FALSE',
          permissions[i].canEdit ? 'TRUE' : 'FALSE',
          permissions[i].canDelete ? 'TRUE' : 'FALSE'
        ]]);
      }
    }
    
    console.log('✅ Updated permissions successfully');
    return { success: true };
  } catch (error) {
    console.error('Error updating permissions:', error);
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

/**
 * Get transaction statistics by category
 */
function getTransactionStatsByCategory() {
  try {
    const transactions = getTransactions();
    const categories = getCategories();
    
    const stats = {};
    
    // Initialize categories
    categories.forEach(category => {
      stats[category.name] = {
        income: 0,
        expense: 0,
        count: 0
      };
    });
    
    // Calculate statistics
    transactions.forEach(transaction => {
      const amount = parseFloat(transaction.amount) || 0;
      const category = transaction.category;
      
      if (!stats[category]) {
        stats[category] = { income: 0, expense: 0, count: 0 };
      }
      
      stats[category].count++;
      
      if (transaction.type === 'Thu') {
        stats[category].income += amount;
      } else if (transaction.type === 'Chi') {
        stats[category].expense += amount;
      }
    });
    
    return stats;
  } catch (error) {
    console.error('Error getting transaction stats by category:', error);
    return {};
  }
}

/**
 * Get monthly transaction trend
 */
function getMonthlyTransactionTrend(year) {
  try {
    const transactions = getTransactions();
    const targetYear = year || new Date().getFullYear();
    
    const monthlyData = {};
    
    // Initialize months
    for (let month = 0; month < 12; month++) {
      const monthName = new Date(targetYear, month, 1).toLocaleDateString('vi-VN', { month: 'long' });
      monthlyData[monthName] = {
        income: 0,
        expense: 0,
        count: 0
      };
    }
    
    // Calculate monthly data
    transactions.forEach(transaction => {
      const transactionDate = new Date(transaction.date);
      
      if (transactionDate.getFullYear() === targetYear) {
        const monthName = transactionDate.toLocaleDateString('vi-VN', { month: 'long' });
        const amount = parseFloat(transaction.amount) || 0;
        
        if (monthlyData[monthName]) {
          monthlyData[monthName].count++;
          
          if (transaction.type === 'Thu') {
            monthlyData[monthName].income += amount;
          } else if (transaction.type === 'Chi') {
            monthlyData[monthName].expense += amount;
          }
        }
      }
    });
    
    return monthlyData;
  } catch (error) {
    console.error('Error getting monthly transaction trend:', error);
    return {};
  }
}

// =================== DATA INITIALIZATION FUNCTIONS ===================

/**
 * Initialize sample data for testing
 */
function initializeSampleData() {
  try {
    console.log('🔄 Initializing sample data...');
    
    // Initialize roles first
    initializeSampleRoles();
    
    // Initialize menus
    initializeSampleMenus();
    
    // Initialize users
    initializeSampleUsers();
    
    // Initialize accounts
    initializeSampleAccounts();
    
    // Initialize categories
    initializeSampleCategories();
    
    // Initialize customers
    initializeSampleCustomers();
    
    // Initialize suppliers
    initializeSampleSuppliers();
    
    // Initialize transactions
    initializeSampleTransactions();
    
    console.log('✅ Sample data initialized successfully');
    return { success: true, message: 'Dữ liệu mẫu đã được tạo thành công' };
  } catch (error) {
    console.error('Error initializing sample data:', error);
    return handleApiError(error);
  }
}

/**
 * Initialize sample roles
 */
function initializeSampleRoles() {
  try {
    const sheet = getSheetByName('Roles');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      sheet.appendRow(['ADMIN', 'Administrator', 'Quản trị viên hệ thống']);
      sheet.appendRow(['USER', 'User', 'Người dùng thông thường']);
      sheet.appendRow(['MANAGER', 'Manager', 'Quản lý']);
      console.log('✅ Sample roles created');
    }
  } catch (error) {
    console.error('Error initializing sample roles:', error);
  }
}

/**
 * Initialize sample menus
 */
function initializeSampleMenus() {
  try {
    const sheet = getSheetByName('Menus');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      sheet.appendRow(['dashboard', 'Tổng quan', 'Trang tổng quan hệ thống', '/dashboard', 'fas fa-chart-line']);
      sheet.appendRow(['transactions', 'Giao dịch', 'Quản lý giao dịch thu chi', '/transactions', 'fas fa-exchange-alt']);
      sheet.appendRow(['accounts', 'Tài khoản', 'Quản lý tài khoản ngân hàng', '/accounts', 'fas fa-university']);
      sheet.appendRow(['categories', 'Danh mục', 'Quản lý danh mục thu chi', '/categories', 'fas fa-tags']);
      sheet.appendRow(['customers', 'Khách hàng', 'Quản lý khách hàng', '/customers', 'fas fa-users']);
      sheet.appendRow(['suppliers', 'Nhà cung cấp', 'Quản lý nhà cung cấp', '/suppliers', 'fas fa-truck']);
      sheet.appendRow(['reports', 'Báo cáo', 'Báo cáo thống kê', '/reports', 'fas fa-chart-bar']);
      sheet.appendRow(['settings', 'Cài đặt', 'Cài đặt hệ thống', '/settings', 'fas fa-cog']);
      console.log('✅ Sample menus created');
    }
  } catch (error) {
    console.error('Error initializing sample menus:', error);
  }
}

/**
 * Initialize sample users
 */
function initializeSampleUsers() {
  try {
    const sheet = getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      sheet.appendRow([
        generateUserId(),
        'admin',
        'admin@nht.com',
        'Administrator',
        'ADMIN',
        getCurrentDate(),
        hashPassword('admin123'),
        'Active'
      ]);
      
      sheet.appendRow([
        generateUserId(),
        'user',
        'user@nht.com',
        'Demo User',
        'USER',
        getCurrentDate(),
        hashPassword('user123'),
        'Active'
      ]);
      
      console.log('✅ Sample users created');
    }
  } catch (error) {
    console.error('Error initializing sample users:', error);
  }
}

/**
 * Initialize sample accounts
 */
function initializeSampleAccounts() {
  try {
    const sheet = getSheetByName('Tài khoản');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      sheet.appendRow([generateAccountId(), 'Tiền mặt', 'Tiền mặt', 2000000, getCurrentDate()]);
      sheet.appendRow([generateAccountId(), 'Vietcombank', 'Ngân hàng', 15000000, getCurrentDate()]);
      sheet.appendRow([generateAccountId(), 'Techcombank', 'Ngân hàng', 8000000, getCurrentDate()]);
      sheet.appendRow([generateAccountId(), 'Ví MoMo', 'Ví điện tử', 800000, getCurrentDate()]);
      
      console.log('✅ Sample accounts created');
    }
  } catch (error) {
    console.error('Error initializing sample accounts:', error);
  }
}

/**
 * Initialize sample categories
 */
function initializeSampleCategories() {
  try {
    const sheet = getSheetByName('Danh mục');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      // Income categories
      sheet.appendRow([generateCategoryId(), 'Lương', 'Thu', 'Lương tháng']);
      sheet.appendRow([generateCategoryId(), 'Thưởng', 'Thu', 'Tiền thưởng']);
      sheet.appendRow([generateCategoryId(), 'Đầu tư', 'Thu', 'Lợi nhuận đầu tư']);
      
      // Expense categories
      sheet.appendRow([generateCategoryId(), 'Ăn uống', 'Chi', 'Chi phí ăn uống']);
      sheet.appendRow([generateCategoryId(), 'Đi lại', 'Chi', 'Chi phí di chuyển']);
      sheet.appendRow([generateCategoryId(), 'Nhà ở', 'Chi', 'Tiền nhà, điện nước']);
      sheet.appendRow([generateCategoryId(), 'Giải trí', 'Chi', 'Chi phí giải trí']);
      sheet.appendRow([generateCategoryId(), 'Mua sắm', 'Chi', 'Mua sắm cá nhân']);
      
      console.log('✅ Sample categories created');
    }
  } catch (error) {
    console.error('Error initializing sample categories:', error);
  }
}

/**
 * Initialize sample customers
 */
function initializeSampleCustomers() {
  try {
    const sheet = getSheetByName('Khách hàng');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      sheet.appendRow([generateCustomerId(), 'Nguyễn Văn A', '0901234567', 'nguyenvana@email.com', 'Hà Nội', getCurrentDate()]);
      sheet.appendRow([generateCustomerId(), 'Trần Thị B', '0902345678', 'tranthib@email.com', 'TP.HCM', getCurrentDate()]);
      sheet.appendRow([generateCustomerId(), 'Lê Hoàng C', '0903456789', 'lehoangc@email.com', 'Đà Nẵng', getCurrentDate()]);
      
      console.log('✅ Sample customers created');
    }
  } catch (error) {
    console.error('Error initializing sample customers:', error);
  }
}

/**
 * Initialize sample suppliers
 */
function initializeSampleSuppliers() {
  try {
    const sheet = getSheetByName('Nhà cung cấp');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      sheet.appendRow([generateSupplierId(), 'Công ty ABC', '0281234567', 'abc@company.com', 'Hà Nội', getCurrentDate()]);
      sheet.appendRow([generateSupplierId(), 'Công ty XYZ', '0282345678', 'xyz@company.com', 'TP.HCM', getCurrentDate()]);
      sheet.appendRow([generateSupplierId(), 'Công ty DEF', '0283456789', 'def@company.com', 'Đà Nẵng', getCurrentDate()]);
      
      console.log('✅ Sample suppliers created');
    }
  } catch (error) {
    console.error('Error initializing sample suppliers:', error);
  }
}

/**
 * Initialize sample transactions
 */
function initializeSampleTransactions() {
  try {
    const sheet = getSheetByName('Giao dịch');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      const today = new Date();
      const currentMonth = today.getMonth();
      const currentYear = today.getFullYear();
      
      // Sample transactions for current month
      const sampleTransactions = [
        ['Thu', 'Lương', 15000000, 'Vietcombank', 'Lương tháng 12'],
        ['Chi', 'Ăn uống', 500000, 'Tiền mặt', 'Ăn trưa với đồng nghiệp'],
        ['Chi', 'Đi lại', 200000, 'Ví MoMo', 'Xe ôm đi làm'],
        ['Thu', 'Thưởng', 3000000, 'Vietcombank', 'Thưởng cuối năm'],
        ['Chi', 'Nhà ở', 8000000, 'Techcombank', 'Tiền nhà tháng 12'],
        ['Chi', 'Giải trí', 1200000, 'Tiền mặt', 'Xem phim và shopping'],
        ['Chi', 'Ăn uống', 800000, 'Ví MoMo', 'Đi ăn gia đình'],
        ['Thu', 'Đầu tư', 2500000, 'Vietcombank', 'Lợi nhuận cổ phiếu'],
        ['Chi', 'Mua sắm', 1500000, 'Techcombank', 'Mua quần áo'],
        ['Chi', 'Đi lại', 300000, 'Tiền mặt', 'Taxi đi sân bay']
      ];
      
      sampleTransactions.forEach((trans, index) => {
        const randomDay = Math.floor(Math.random() * 28) + 1;
        const transactionDate = new Date(currentYear, currentMonth, randomDay);
        
        sheet.appendRow([
          generateTransactionId(),
          transactionDate.toLocaleDateString('vi-VN'),
          trans[0], // type
          trans[1], // category
          trans[2], // amount
          trans[3], // account
          trans[4], // note
          'Hoàn thành'
        ]);
      });
      
      console.log('✅ Sample transactions created');
    }
  } catch (error) {
    console.error('Error initializing sample transactions:', error);
  }
}

// =================== BACKUP & RESTORE FUNCTIONS ===================

/**
 * Export all data
 */
function exportAllData() {
  try {
    const data = {
      transactions: getTransactions(),
      accounts: getAccounts(),
      categories: getCategories(),
      customers: getCustomers(),
      suppliers: getSuppliers(),
      users: getUsers(),
      roles: getRoles(),
      exportDate: getCurrentTimestamp(),
      version: '1.0.0'
    };
    
    console.log('✅ Data exported successfully');
    return { success: true, data: data };
  } catch (error) {
    console.error('Error exporting data:', error);
    return handleApiError(error);
  }
}

/**
 * Clear all data (for testing purposes)
 */
function clearAllData() {
  try {
    const sheetNames = ['Giao dịch', 'Tài khoản', 'Danh mục', 'Khách hàng', 'Nhà cung cấp'];
    
    sheetNames.forEach(sheetName => {
      try {
        const sheet = getSheetByName(sheetName);
        const lastRow = sheet.getLastRow();
        
        if (lastRow > 1) {
          sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
        }
      } catch (error) {
        console.warn(`Warning: Could not clear sheet ${sheetName}:`, error);
      }
    });
    
    console.log('✅ All data cleared successfully');
    return { success: true, message: 'Đã xóa tất cả dữ liệu' };
  } catch (error) {
    console.error('Error clearing data:', error);
    return handleApiError(error);
  }
}

// =================== SEARCH FUNCTIONS ===================

/**
 * Search transactions
 */
function searchTransactions(query, filters = {}) {
  try {
    let transactions = getTransactions();
    
    // Apply text search
    if (query && query.trim()) {
      const searchTerm = query.toLowerCase();
      transactions = transactions.filter(transaction => 
        (transaction.note || '').toLowerCase().includes(searchTerm) ||
        (transaction.category || '').toLowerCase().includes(searchTerm) ||
        (transaction.account || '').toLowerCase().includes(searchTerm)
      );
    }
    
    // Apply filters
    if (filters.type) {
      transactions = transactions.filter(t => t.type === filters.type);
    }
    
    if (filters.category) {
      transactions = transactions.filter(t => t.category === filters.category);
    }
    
    if (filters.account) {
      transactions = transactions.filter(t => t.account === filters.account);
    }
    
    if (filters.dateFrom) {
      transactions = transactions.filter(t => new Date(t.date) >= new Date(filters.dateFrom));
    }
    
    if (filters.dateTo) {
      transactions = transactions.filter(t => new Date(t.date) <= new Date(filters.dateTo));
    }
    
    if (filters.amountMin) {
      transactions = transactions.filter(t => parseFloat(t.amount) >= parseFloat(filters.amountMin));
    }
    
    if (filters.amountMax) {
      transactions = transactions.filter(t => parseFloat(t.amount) <= parseFloat(filters.amountMax));
    }
    
    return transactions;
  } catch (error) {
    console.error('Error searching transactions:', error);
    return [];
  }
}

// =================== VALIDATION FUNCTIONS ===================

/**
 * Validate transaction data
 */
function validateTransaction(transaction) {
  const errors = [];
  
  if (!transaction.date) {
    errors.push('Ngày giao dịch không được để trống');
  }
  
  if (!transaction.type || !['Thu', 'Chi'].includes(transaction.type)) {
    errors.push('Loại giao dịch không hợp lệ');
  }
  
  if (!transaction.category || transaction.category.trim() === '') {
    errors.push('Danh mục không được để trống');
  }
  
  if (!transaction.amount || isNaN(parseFloat(transaction.amount)) || parseFloat(transaction.amount) <= 0) {
    errors.push('Số tiền phải là số dương');
  }
  
  if (!transaction.account || transaction.account.trim() === '') {
    errors.push('Tài khoản không được để trống');
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Validate account data
 */
function validateAccount(account) {
  const errors = [];
  
  if (!account.name || account.name.trim() === '') {
    errors.push('Tên tài khoản không được để trống');
  }
  
  if (!account.type || account.type.trim() === '') {
    errors.push('Loại tài khoản không được để trống');
  }
  
  if (account.balance !== undefined && isNaN(parseFloat(account.balance))) {
    errors.push('Số dư phải là số');
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Validate category data
 */
function validateCategory(category) {
  const errors = [];
  
  if (!category.name || category.name.trim() === '') {
    errors.push('Tên danh mục không được để trống');
  }
  
  if (!category.type || !['Thu', 'Chi'].includes(category.type)) {
    errors.push('Loại danh mục không hợp lệ');
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Validate customer data
 */
function validateCustomer(customer) {
  const errors = [];
  
  if (!customer.name || customer.name.trim() === '') {
    errors.push('Tên khách hàng không được để trống');
  }
  
  if (customer.email && !isValidEmail(customer.email)) {
    errors.push('Email không hợp lệ');
  }
  
  if (customer.phone && !/^[0-9+\-\s()]+$/.test(customer.phone)) {
    errors.push('Số điện thoại không hợp lệ');
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * Validate supplier data
 */
function validateSupplier(supplier) {
  const errors = [];
  
  if (!supplier.name || supplier.name.trim() === '') {
    errors.push('Tên nhà cung cấp không được để trống');
  }
  
  if (supplier.email && !isValidEmail(supplier.email)) {
    errors.push('Email không hợp lệ');
  }
  
  if (supplier.phone && !/^[0-9+\-\s()]+$/.test(supplier.phone)) {
    errors.push('Số điện thoại không hợp lệ');
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

// =================== LOGGING & DEBUG FUNCTIONS ===================

/**
 * Log activity
 */
function logActivity(action, details = '') {
  try {
    const sheet = getSheetByName('Activity_Log');
    sheet.appendRow([
      getCurrentTimestamp(),
      action,
      details,
      Session.getActiveUser().getEmail()
    ]);
  } catch (error) {
    console.log('Could not log activity:', error);
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
      version: '1.0.0'
    };
  } catch (error) {
    console.error('Error getting system info:', error);
    return { error: error.toString() };
  }
}

/**
 * Test function to verify system is working
 */
function testSystem() {
  try {
    console.log('🧪 Testing NHT Finance System...');
    
    const tests = {
      spreadsheet: false,
      users: false,
      transactions: false,
      accounts: false,
      categories: false
    };
    
    // Test spreadsheet access
    try {
      const ss = getSpreadsheet();
      tests.spreadsheet = true;
      console.log('✅ Spreadsheet access: OK');
    } catch (error) {
      console.error('❌ Spreadsheet access: FAILED');
    }
    
    // Test user authentication
    try {
      const auth = authenticateUser('admin', 'admin123');
      tests.users = auth.success;
      console.log('✅ User authentication: OK');
    } catch (error) {
      console.error('❌ User authentication: FAILED');
    }
    
    // Test data access
    try {
      const transactions = getTransactions();
      tests.transactions = Array.isArray(transactions);
      console.log('✅ Transactions access: OK');
    } catch (error) {
      console.error('❌ Transactions access: FAILED');
    }
    
    try {
      const accounts = getAccounts();
      tests.accounts = Array.isArray(accounts);
      console.log('✅ Accounts access: OK');
    } catch (error) {
      console.error('❌ Accounts access: FAILED');
    }
    
    try {
      const categories = getCategories();
      tests.categories = Array.isArray(categories);
      console.log('✅ Categories access: OK');
    } catch (error) {
      console.error('❌ Categories access: FAILED');
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

// =================== REPORT FUNCTIONS ===================

/**
 * Generate monthly report
 */
function generateMonthlyReport(year, month) {
  try {
    const transactions = getTransactions();
    const targetDate = new Date(year, month - 1, 1); // month is 0-based
    
    const monthlyTransactions = transactions.filter(transaction => {
      const transactionDate = new Date(transaction.date);
      return transactionDate.getFullYear() === year && 
             transactionDate.getMonth() === month - 1;
    });
    
    let totalIncome = 0;
    let totalExpense = 0;
    const categoryStats = {};
    
    monthlyTransactions.forEach(transaction => {
      const amount = parseFloat(transaction.amount) || 0;
      
      if (transaction.type === 'Thu') {
        totalIncome += amount;
      } else if (transaction.type === 'Chi') {
        totalExpense += amount;
      }
      
      // Category statistics
      if (!categoryStats[transaction.category]) {
        categoryStats[transaction.category] = {
          income: 0,
          expense: 0,
          count: 0
        };
      }
      
      categoryStats[transaction.category].count++;
      if (transaction.type === 'Thu') {
        categoryStats[transaction.category].income += amount;
      } else {
        categoryStats[transaction.category].expense += amount;
      }
    });
    
    return {
      success: true,
      data: {
        period: `${month}/${year}`,
        totalIncome,
        totalExpense,
        netProfit: totalIncome - totalExpense,
        transactionCount: monthlyTransactions.length,
        categoryStats,
        transactions: monthlyTransactions
      }
    };
    
  } catch (error) {
    console.error('Error generating monthly report:', error);
    return handleApiError(error);
  }
}

/**
 * Generate yearly report
 */
function generateYearlyReport(year) {
  try {
    const transactions = getTransactions();
    
    const yearlyTransactions = transactions.filter(transaction => {
      const transactionDate = new Date(transaction.date);
      return transactionDate.getFullYear() === year;
    });
    
    let totalIncome = 0;
    let totalExpense = 0;
    const monthlyStats = {};
    const categoryStats = {};
    
    // Initialize monthly stats
    for (let month = 0; month < 12; month++) {
      const monthName = new Date(year, month, 1).toLocaleDateString('vi-VN', { month: 'long' });
      monthlyStats[monthName] = {
        income: 0,
        expense: 0,
        count: 0
      };
    }
    
    yearlyTransactions.forEach(transaction => {
      const amount = parseFloat(transaction.amount) || 0;
      const transactionDate = new Date(transaction.date);
      const monthName = transactionDate.toLocaleDateString('vi-VN', { month: 'long' });
      
      if (transaction.type === 'Thu') {
        totalIncome += amount;
        if (monthlyStats[monthName]) {
          monthlyStats[monthName].income += amount;
        }
      } else if (transaction.type === 'Chi') {
        totalExpense += amount;
        if (monthlyStats[monthName]) {
          monthlyStats[monthName].expense += amount;
        }
      }
      
      if (monthlyStats[monthName]) {
        monthlyStats[monthName].count++;
      }
      
      // Category statistics
      if (!categoryStats[transaction.category]) {
        categoryStats[transaction.category] = {
          income: 0,
          expense: 0,
          count: 0
        };
      }
      
      categoryStats[transaction.category].count++;
      if (transaction.type === 'Thu') {
        categoryStats[transaction.category].income += amount;
      } else {
        categoryStats[transaction.category].expense += amount;
      }
    });
    
    return {
      success: true,
      data: {
        year,
        totalIncome,
        totalExpense,
        netProfit: totalIncome - totalExpense,
        transactionCount: yearlyTransactions.length,
        monthlyStats,
        categoryStats
      }
    };
    
  } catch (error) {
    console.error('Error generating yearly report:', error);
    return handleApiError(error);
  }
}

// =================== IMPORT/EXPORT FUNCTIONS ===================

/**
 * Import transactions from CSV data
 */
function importTransactionsFromCSV(csvData) {
  try {
    const lines = csvData.split('\n');
    const headers = lines[0].split(',');
    
    // Validate headers
    const requiredHeaders = ['date', 'type', 'category', 'amount', 'account'];
    const missingHeaders = requiredHeaders.filter(header => 
      !headers.some(h => h.toLowerCase().trim() === header)
    );
    
    if (missingHeaders.length > 0) {
      return {
        success: false,
        message: `Thiếu các cột bắt buộc: ${missingHeaders.join(', ')}`
      };
    }
    
    const sheet = getSheetByName('Giao dịch');
    let importedCount = 0;
    let errorCount = 0;
    const errors = [];
    
    for (let i = 1; i < lines.length; i++) {
      if (!lines[i].trim()) continue; // Skip empty lines
      
      const values = lines[i].split(',');
      
      try {
        const transaction = {
          date: values[0]?.trim(),
          type: values[1]?.trim(),
          category: values[2]?.trim(),
          amount: parseFloat(values[3]?.trim()) || 0,
          account: values[4]?.trim(),
          note: values[5]?.trim() || ''
        };
        
        // Validate transaction
        const validation = validateTransaction(transaction);
        if (!validation.isValid) {
          errors.push(`Dòng ${i + 1}: ${validation.errors.join(', ')}`);
          errorCount++;
          continue;
        }
        
        // Add transaction
        const id = generateTransactionId();
        sheet.appendRow([
          id,
          transaction.date,
          transaction.type,
          transaction.category,
          transaction.amount,
          transaction.account,
          transaction.note,
          'Hoàn thành'
        ]);
        
        importedCount++;
        
      } catch (error) {
        errors.push(`Dòng ${i + 1}: Lỗi xử lý dữ liệu`);
        errorCount++;
      }
    }
    
    return {
      success: true,
      message: `Import thành công ${importedCount} giao dịch${errorCount > 0 ? `, ${errorCount} lỗi` : ''}`,
      importedCount,
      errorCount,
      errors: errors.slice(0, 10) // Limit error messages
    };
    
  } catch (error) {
    console.error('Error importing transactions from CSV:', error);
    return handleApiError(error);
  }
}

/**
 * Export transactions to CSV format
 */
function exportTransactionsToCSV() {
  try {
    const transactions = getTransactions();
    
    let csvContent = 'ID,Ngày,Loại,Danh mục,Số tiền,Tài khoản,Ghi chú,Trạng thái\n';
    
    transactions.forEach(transaction => {
      csvContent += [
        transaction.id,
        transaction.date,
        transaction.type,
        transaction.category,
        transaction.amount,
        transaction.account,
        `"${transaction.note || ''}"`, // Wrap in quotes to handle commas
        transaction.status
      ].join(',') + '\n';
    });
    
    return {
      success: true,
      data: csvContent,
      filename: `giao_dich_${new Date().toISOString().split('T')[0]}.csv`
    };
    
  } catch (error) {
    console.error('Error exporting transactions to CSV:', error);
    return handleApiError(error);
  }
}

// =================== UTILITY & HELPER FUNCTIONS ===================

/**
 * Format number with Vietnamese locale
 */
function formatNumber(number) {
  try {
    return new Intl.NumberFormat('vi-VN').format(number || 0);
  } catch (error) {
    return (number || 0).toString();
  }
}

/**
 * Parse Vietnamese formatted number
 */
function parseVietnameseNumber(numberString) {
  try {
    if (typeof numberString === 'number') return numberString;
    return parseFloat(String(numberString).replace(/[.,]/g, match => match === ',' ? '.' : ''));
  } catch (error) {
    return 0;
  }
}

/**
 * Get date range for period
 */
function getDateRange(period) {
  const today = new Date();
  const ranges = {
    'today': {
      from: new Date(today.getFullYear(), today.getMonth(), today.getDate()),
      to: new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59)
    },
    'this_week': {
      from: new Date(today.getFullYear(), today.getMonth(), today.getDate() - today.getDay()),
      to: new Date(today.getFullYear(), today.getMonth(), today.getDate() - today.getDay() + 6, 23, 59, 59)
    },
    'this_month': {
      from: new Date(today.getFullYear(), today.getMonth(), 1),
      to: new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59)
    },
    'this_year': {
      from: new Date(today.getFullYear(), 0, 1),
      to: new Date(today.getFullYear(), 11, 31, 23, 59, 59)
    }
  };
  
  return ranges[period] || ranges['this_month'];
}

/**
 * Generate random sample data for testing
 */
function generateRandomTransactions(count = 50) {
  try {
    const sheet = getSheetByName('Giao dịch');
    const categories = getCategories();
    const accounts = getAccounts();
    
    if (categories.length === 0 || accounts.length === 0) {
      return { success: false, message: 'Cần có danh mục và tài khoản trước khi tạo giao dịch mẫu' };
    }
    
    const incomeCategories = categories.filter(c => c.type === 'Thu');
    const expenseCategories = categories.filter(c => c.type === 'Chi');
    
    let addedCount = 0;
    
    for (let i = 0; i < count; i++) {
      // Random date within last 6 months
      const randomDate = new Date();
      randomDate.setMonth(randomDate.getMonth() - Math.floor(Math.random() * 6));
      randomDate.setDate(Math.floor(Math.random() * 28) + 1);
      
      // Random type (70% expense, 30% income)
      const isIncome = Math.random() < 0.3;
      const type = isIncome ? 'Thu' : 'Chi';
      
      // Random category based on type
      const availableCategories = isIncome ? incomeCategories : expenseCategories;
      const category = availableCategories[Math.floor(Math.random() * availableCategories.length)];
      
      // Random amount
      const amount = isIncome 
        ? Math.floor(Math.random() * 20000000) + 5000000 // 5M - 25M for income
        : Math.floor(Math.random() * 2000000) + 50000;   // 50K - 2M for expense
      
      // Random account
      const account = accounts[Math.floor(Math.random() * accounts.length)];
      
      // Random note
      const notes = isIncome 
        ? ['Lương tháng', 'Thưởng', 'Freelance', 'Bán hàng', 'Lợi nhuận đầu tư']
        : ['Mua sắm', 'Ăn uống', 'Đi lại', 'Giải trí', 'Sinh hoạt phí'];
      const note = notes[Math.floor(Math.random() * notes.length)];
      
      sheet.appendRow([
        generateTransactionId(),
        randomDate.toLocaleDateString('vi-VN'),
        type,
        category.name,
        amount,
        account.name,
        note,
        'Hoàn thành'
      ]);
      
      addedCount++;
    }
    
    console.log(`✅ Generated ${addedCount} random transactions`);
    return { 
      success: true, 
      message: `Đã tạo ${addedCount} giao dịch mẫu thành công`,
      count: addedCount 
    };
    
  } catch (error) {
    console.error('Error generating random transactions:', error);
    return handleApiError(error);
  }
}

// =================== SYSTEM MAINTENANCE FUNCTIONS ===================

/**
 * Clean up old logs and temporary data
 */
function cleanupSystem() {
  try {
    let cleanedItems = 0;
    
    // Clean up activity logs older than 3 months
    try {
      const logSheet = getSheetByName('Activity_Log');
      const logData = logSheet.getDataRange().getValues();
      const threeMonthsAgo = new Date();
      threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
      
      for (let i = logData.length - 1; i >= 1; i--) {
        const logDate = new Date(logData[i][0]);
        if (logDate < threeMonthsAgo) {
          logSheet.deleteRow(i + 1);
          cleanedItems++;
        }
      }
    } catch (error) {
      console.warn('Could not clean activity logs:', error);
    }
    
    console.log(`✅ System cleanup completed, removed ${cleanedItems} old items`);
    return { 
      success: true, 
      message: `Đã dọn dẹp ${cleanedItems} mục cũ`,
      cleanedItems 
    };
    
  } catch (error) {
    console.error('Error during system cleanup:', error);
    return handleApiError(error);
  }
}

/**
 * Backup all data to a new spreadsheet
 */
function backupData() {
  try {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const backupName = `NHT_Finance_Backup_${timestamp}`;
    
    // Create new spreadsheet for backup
    const backupSS = SpreadsheetApp.create(backupName);
    const originalSS = getSpreadsheet();
    
    // Copy all sheets
    const sheets = originalSS.getSheets();
    let copiedSheets = 0;
    
    sheets.forEach(sheet => {
      try {
        sheet.copyTo(backupSS);
        copiedSheets++;
      } catch (error) {
        console.warn(`Could not copy sheet ${sheet.getName()}:`, error);
      }
    });
    
    // Remove default sheet if backup has content
    try {
      if (copiedSheets > 0) {
        const defaultSheet = backupSS.getSheetByName('Sheet1');
        if (defaultSheet && backupSS.getSheets().length > 1) {
          backupSS.deleteSheet(defaultSheet);
        }
      }
    } catch (error) {
      console.warn('Could not remove default sheet:', error);
    }
    
    console.log(`✅ Backup created: ${backupName}`);
    return { 
      success: true, 
      message: `Backup được tạo thành công: ${backupName}`,
      backupId: backupSS.getId(),
      backupUrl: backupSS.getUrl(),
      sheetsCopied: copiedSheets
    };
    
  } catch (error) {
    console.error('Error creating backup:', error);
    return handleApiError(error);
  }
}

// =================== END OF FILE ===================

console.log('✅ NHT_Code.gs loaded successfully - All functions defined');
console.log('📊 Total functions available:', Object.getOwnPropertyNames(this).filter(name => typeof this[name] === 'function').length);
console.log('🎯 System ready for deployment!');
