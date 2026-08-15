/**
 * CashFlow v2 - Google Apps Script API
 * Hỗ trợ: Transactions, Loans, Members
 * 
 * Hướng dẫn:
 * 1. Mở Google Sheets → Tiện ích mở rộng → Apps Script
 * 2. Xóa code cũ, dán code này vào → Lưu
 * 3. Chọn hàm "setupSheet" → ▶️ Run → cấp quyền
 * 4. Deploy → New deployment → Web app → Anyone → Deploy → copy URL
 * 5. Dán URL vào CashFlow → Cài đặt → Kiểm tra → Xong!
 */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet Transactions
  let sheet = ss.getSheetByName('Transactions');
  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
    sheet.appendRow(['ID', 'Loại', 'Số tiền', 'Danh mục', 'Ghi chú', 'Ngày', 'Thành viên', 'Ngày tạo']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:H1').setFontWeight('bold');
  }
  
  // Sheet Loans
  let loanSheet = ss.getSheetByName('Loans');
  if (!loanSheet) {
    loanSheet = ss.insertSheet('Loans');
    loanSheet.appendRow(['ID', 'Tên', 'Emoji', 'Loại', 'Gốc vay', 'Lãi suất', 'Kỳ hạn', 'Trả/tháng', 'Ngày BĐ', 'Ghi chú', 'Ngày tạo']);
    loanSheet.setFrozenRows(1);
    loanSheet.getRange('A1:K1').setFontWeight('bold');
  }
  
  // Sheet Members
  let memberSheet = ss.getSheetByName('Members');
  if (!memberSheet) {
    memberSheet = ss.insertSheet('Members');
    memberSheet.appendRow(['ID', 'Tên', 'Avatar', 'Màu']);
    memberSheet.setFrozenRows(1);
    memberSheet.getRange('A1:D1').setFontWeight('bold');
  }
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return -1;
}

function responseJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action || 'getAll';
  
  if (action === 'ping') {
    return responseJson({ status: 'ok', message: 'Kết nối thành công!' });
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    // ==================== TRANSACTIONS ====================
    if (action === 'getAll') {
      const sheet = getSheet('Transactions');
      if (!sheet) return responseJson({ success: true, data: [] });
      const data = sheet.getDataRange().getValues();
      const transactions = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          transactions.push({
            id: data[i][0], type: data[i][1], amount: Number(data[i][2]),
            category: data[i][3], note: data[i][4], date: data[i][5],
            memberId: data[i][6], createdAt: data[i][7]
          });
        }
      }
      return responseJson({ success: true, data: transactions });
    }
    
    if (action === 'add') {
      const obj = JSON.parse(e.parameter.data);
      getSheet('Transactions').appendRow([
        obj.id || '', obj.type || '', obj.amount || 0,
        obj.category || '', obj.note || '', obj.date || '',
        obj.memberId || '', obj.createdAt || new Date().toISOString()
      ]);
      return responseJson({ success: true, message: 'Đã thêm giao dịch' });
    }
    
    if (action === 'update') {
      const id = e.parameter.id;
      const obj = JSON.parse(e.parameter.data);
      const sheet = getSheet('Transactions');
      const rowNum = findRowById(sheet, id);
      if (rowNum > -1) {
        sheet.getRange(rowNum, 1, 1, 8).setValues([[
          obj.id || id, obj.type || '', obj.amount || 0,
          obj.category || '', obj.note || '', obj.date || '',
          obj.memberId || '', obj.createdAt || ''
        ]]);
        return responseJson({ success: true });
      }
      throw new Error('Không tìm thấy giao dịch');
    }
    
    if (action === 'delete') {
      const sheet = getSheet('Transactions');
      const rowNum = findRowById(sheet, e.parameter.id);
      if (rowNum > -1) { sheet.deleteRow(rowNum); return responseJson({ success: true }); }
      throw new Error('Không tìm thấy');
    }
    
    if (action === 'sync') {
      const transactions = JSON.parse(e.parameter.data);
      const sheet = getSheet('Transactions');
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      if (transactions && transactions.length > 0) {
        const rows = transactions.map(t => [
          t.id || '', t.type || '', t.amount || 0,
          t.category || '', t.note || '', t.date || '',
          t.memberId || '', t.createdAt || ''
        ]);
        sheet.getRange(2, 1, rows.length, 8).setValues(rows);
      }
      return responseJson({ success: true, message: 'Đã đồng bộ giao dịch' });
    }
    
    // ==================== LOANS ====================
    if (action === 'getLoans') {
      const sheet = getSheet('Loans');
      if (!sheet) return responseJson({ success: true, data: [] });
      const data = sheet.getDataRange().getValues();
      const loans = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          loans.push({
            id: data[i][0], name: data[i][1], emoji: data[i][2],
            loanType: data[i][3], principal: Number(data[i][4]),
            interestRate: Number(data[i][5]), termMonths: Number(data[i][6]),
            monthlyPayment: Number(data[i][7]), startDate: data[i][8],
            note: data[i][9], createdAt: data[i][10]
          });
        }
      }
      return responseJson({ success: true, data: loans });
    }
    
    if (action === 'syncLoans') {
      const loans = JSON.parse(e.parameter.data);
      const sheet = getSheet('Loans');
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      if (loans && loans.length > 0) {
        const rows = loans.map(l => [
          l.id || '', l.name || '', l.emoji || '', l.loanType || '',
          l.principal || 0, l.interestRate || 0, l.termMonths || 0,
          l.monthlyPayment || 0, l.startDate || '', l.note || '', l.createdAt || ''
        ]);
        sheet.getRange(2, 1, rows.length, 11).setValues(rows);
      }
      return responseJson({ success: true, message: 'Đã đồng bộ khoản vay' });
    }
    
    // ==================== MEMBERS ====================
    if (action === 'getMembers') {
      const sheet = getSheet('Members');
      if (!sheet) return responseJson({ success: true, data: [] });
      const data = sheet.getDataRange().getValues();
      const members = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          members.push({
            id: data[i][0], name: data[i][1], avatar: data[i][2], color: data[i][3]
          });
        }
      }
      return responseJson({ success: true, data: members });
    }
    
    if (action === 'syncMembers') {
      const members = JSON.parse(e.parameter.data);
      const sheet = getSheet('Members');
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      if (members && members.length > 0) {
        const rows = members.map(m => [m.id || '', m.name || '', m.avatar || '', m.color || '']);
        sheet.getRange(2, 1, rows.length, 4).setValues(rows);
      }
      return responseJson({ success: true, message: 'Đã đồng bộ thành viên' });
    }
    
    return responseJson({ success: false, error: 'Hành động không hợp lệ' });
    
  } catch (error) {
    return responseJson({ success: false, error: error.toString() });
  } finally {
    lock.releaseLock();
  }
}
