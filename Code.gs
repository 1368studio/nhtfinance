// =================== NHT_Code.gs - HOÀN THIỆN ===================

/**
 * Main entry point for the web application
 * Serves the main HTML file with proper configuration
 */
function doGet() {
  console.log('🚀 Starting NHT Finance System...');
  
  try {
    // Initialize system if needed (chỉ chạy 1 lần)
    const usersSheet = getSheetByName('Users');
    const userData = usersSheet.getDataRange().getValues();
    
    if (userData.length <= 1) {
      console.log('🆕 First run detected, initializing sample data...');
      initializeSampleData();
    }
    
    // QUAN TRỌNG: Phải dùng 'Index' chứ không phải 'NHT_Index'
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Hệ thống quản lý tài chính NHT')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    console.error('❌ Error in doGet:', error);
    
    // Return simple HTML page
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>Error - NHT Finance</title>
          <style>
            body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
            .error { color: #dc3545; }
          </style>
        </head>
        <body>
          <h1 class="error">⚠️ Lỗi hệ thống</h1>
          <p>Không thể tải ứng dụng. Vui lòng thử lại sau.</p>
          <button onclick="window.location.reload()">Làm mới trang</button>
          <p><small>Chi tiết lỗi: ${error.message}</small></p>
        </body>
      </html>
    `);
  }
}

/**
 * Include HTML files for templating
 * This function allows us to include CSS and JS files in the main HTML
 * QUAN TRỌNG: Tên file phải khớp với tên file thực tế trong Google Apps Script
 */
function include(filename) {
  try {
    console.log(`📄 Including file: ${filename}`);
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error(`❌ Error including file ${filename}:`, error);
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
    
    // Initialize roles
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
  const sheet = getSheetByName('Roles');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    sheet.appendRow(['ADMIN', 'Administrator', 'Quản trị viên hệ thống']);
    sheet.appendRow(['USER', 'User', 'Người dùng thông thường']);
    sheet.appendRow(['MANAGER', 'Manager', 'Quản lý']);
  }
}

/**
 * Initialize sample menus
 */
function initializeSampleMenus() {
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
  }
}

/**
 * Initialize sample users
 */
function initializeSampleUsers() {
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
      'User Demo',
      'USER',
      getCurrentDate(),
      hashPassword('user123'),
      'Active'
    ]);
  }
}

/**
 * Initialize sample accounts
 */
function initializeSampleAccounts() {
  const sheet = getSheetByName('Tài khoản');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    sheet.appendRow([generateAccountId(), 'Tiền mặt', 'Tiền mặt', 1000000, getCurrentDate()]);
    sheet.appendRow([generateAccountId(), 'Vietcombank', 'Ngân hàng', 5000000, getCurrentDate()]);
    sheet.appendRow([generateAccountId(), 'Techcombank', 'Ngân hàng', 3000000, getCurrentDate()]);
    sheet.appendRow([generateAccountId(), 'Ví MoMo', 'Ví điện tử', 500000, getCurrentDate()]);
  }
}

/**
 * Initialize sample categories
 */
function initializeSampleCategories() {
  const sheet = getSheetByName('Danh mục');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    // Income categories
    sheet.appendRow([generateCategoryId(), 'Lương', 'Thu', 'Lương tháng']);
    sheet.appendRow([generateCategoryId(), 'Thưởng', 'Thu', 'Tiền thưởng']);
    sheet.appendRow([generateCategoryId(), 'Đầu tư', 'Thu', 'Lợi nhuận đầu tư']);
    
    // Expense categories
    sheet.appendRow([generateCategoryId(), 'Ăn uống', 'Chi', 'Chi phí ăn uống']);
    sheet.appendRow([generateCategoryId(), 'Đi lại', 'Chi', 'Chi phí đi lại']);
    sheet.appendRow([generateCategoryId(), 'Nhà ở', 'Chi', 'Tiền nhà, điện nước']);
    sheet.appendRow([generateCategoryId(), 'Giải trí', 'Chi', 'Chi phí giải trí']);
    sheet.appendRow([generateCategoryId(), 'Mua sắm', 'Chi', 'Mua sắm cá nhân']);
  }
}

/**
 * Initialize sample customers
 */
function initializeSampleCustomers() {
  const sheet = getSheetByName('Khách hàng');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    sheet.appendRow([generateCustomerId(), 'Nguyễn Văn A', '0901234567', 'nguyenvana@email.com', 'Hà Nội', getCurrentDate()]);
    sheet.appendRow([generateCustomerId(), 'Trần Thị B', '0902345678', 'tranthib@email.com', 'TP.HCM', getCurrentDate()]);
    sheet.appendRow([generateCustomerId(), 'Lê Hoàng C', '0903456789', 'lehoangc@email.com', 'Đà Nẵng', getCurrentDate()]);
  }
}

/**
 * Initialize sample suppliers
 */
function initializeSampleSuppliers() {
  const sheet = getSheetByName('Nhà cung cấp');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    sheet.appendRow([generateSupplierId(), 'Công ty ABC', '0281234567', 'abc@company.com', 'Hà Nội', getCurrentDate()]);
    sheet.appendRow([generateSupplierId(), 'Công ty XYZ', '0282345678', 'xyz@company.com', 'TP.HCM', getCurrentDate()]);
    sheet.appendRow([generateSupplierId(), 'Công ty DEF', '0283456789', 'def@company.com', 'Đà Nẵng', getCurrentDate()]);
  }
}

/**
 * Initialize sample transactions
 */
function initializeSampleTransactions() {
  const sheet = getSheetByName('Giao dịch');
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    
    // Sample transactions for current month
    for (let i = 1; i <= 10; i++) {
      const randomDay = Math.floor(Math.random() * 28) + 1;
      const transactionDate = new Date(currentYear, currentMonth, randomDay);
      
      if (i % 3 === 0) {
        // Income transaction
        sheet.appendRow([
          generateTransactionId(),
          transactionDate.toLocaleDateString('vi-VN'),
          'Thu',
          'Lương',
          Math.floor(Math.random() * 5000000) + 1000000,
          'Vietcombank',
          `Giao dịch thu nhập ${i}`,
          'Hoàn thành'
        ]);
      } else {
        // Expense transaction
        const categories = ['Ăn uống', 'Đi lại', 'Nhà ở', 'Giải trí', 'Mua sắm'];
        const randomCategory = categories[Math.floor(Math.random() * categories.length)];
        
        sheet.appendRow([
          generateTransactionId(),
          transactionDate.toLocaleDateString('vi-VN'),
          'Chi',
          randomCategory,
          Math.floor(Math.random() * 1000000) + 50000,
          'Tiền mặt',
          `Giao dịch chi tiêu ${i}`,
          'Hoàn thành'
        ]);
      }
    }
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
      const sheet = getSheetByName(sheetName);
      const lastRow = sheet.getLastRow();
      
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
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
        transaction.note.toLowerCase().includes(searchTerm) ||
        transaction.category.toLowerCase().includes(searchTerm) ||
        transaction.account.toLowerCase().includes(searchTerm)
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

// =================== END OF FILE ===================

console.log('✅ NHT_Code.gs loaded successfully - All functions defined');
