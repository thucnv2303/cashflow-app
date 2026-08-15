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
    sheet.appendRow(['ID', 'Loại', 'Số tiền', 'Danh mục', 'Ghi chú', 'Ngày', 'Thành viên', 'Ngày tạo', 'Chi cho ai']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:I1').setFontWeight('bold');
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
    memberSheet.appendRow(['ID', 'Tên', 'Avatar', 'Màu', 'AvatarImg', 'AvatarId']);
    memberSheet.setFrozenRows(1);
    memberSheet.getRange('A1:F1').setFontWeight('bold');
  }

  // Sheet Budgets
  let budgetSheet = ss.getSheetByName('Budgets');
  if (!budgetSheet) {
    budgetSheet = ss.insertSheet('Budgets');
    budgetSheet.appendRow(['CategoryID', 'Số tiền', 'Cập nhật']);
    budgetSheet.setFrozenRows(1);
    budgetSheet.getRange('A1:C1').setFontWeight('bold');
  }

  // Sheet CustomCategories
  let catSheet = ss.getSheetByName('CustomCategories');
  if (!catSheet) {
    catSheet = ss.insertSheet('CustomCategories');
    catSheet.appendRow(['ID', 'Tên', 'Emoji', 'Loại', 'Ngày tạo']);
    catSheet.setFrozenRows(1);
    catSheet.getRange('A1:E1').setFontWeight('bold');
  }

  // Sheet SavingsGoals
  let savingsSheet = ss.getSheetByName('SavingsGoals');
  if (!savingsSheet) {
    savingsSheet = ss.insertSheet('SavingsGoals');
    savingsSheet.appendRow(['ID', 'Tên', 'Emoji', 'Mục tiêu', 'Hiện có', 'Hạn ngày', 'Thành viên', 'Ngày tạo']);
    savingsSheet.setFrozenRows(1);
    savingsSheet.getRange('A1:H1').setFontWeight('bold');
  }

  // Sheet SavingsLogs
  let logsSheet = ss.getSheetByName('SavingsLogs');
  if (!logsSheet) {
    logsSheet = ss.insertSheet('SavingsLogs');
    logsSheet.appendRow(['ID', 'GoalID', 'GoalName', 'Loại', 'Số tiền', 'Thành viên', 'Ngày', 'Ghi chú', 'Ngày tạo']);
    logsSheet.setFrozenRows(1);
    logsSheet.getRange('A1:I1').setFontWeight('bold');
  }

  // Sheet PendingTransactions (Bank Webhook Inbox)
  let pendingSheet = ss.getSheetByName('PendingTransactions');
  if (!pendingSheet) {
    pendingSheet = ss.insertSheet('PendingTransactions');
    pendingSheet.appendRow(['ID', 'Ngân hàng', 'Loại', 'Số tiền', 'Nội dung thô', 'Danh mục gợi ý', 'Thành viên', 'Ngày giờ', 'Trạng thái', 'Ngày tạo']);
    pendingSheet.setFrozenRows(1);
    pendingSheet.getRange('A1:J1').setFontWeight('bold');
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

function responseJson(data, callback) {
  var json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // Support both JSON body in postData and standard form parameters
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      e.parameter = e.parameter || {};
      for (var k in body) { e.parameter[k] = body[k]; }
    } catch(err) {}
  }
  return doGet(e);
}

function doGet(e) {
  e = e || {};
  e.parameter = e.parameter || {};
  const action = e.parameter.action || 'getAll';
  const cb = e.parameter.callback || null;
  
  if (action === 'ping') {
    return responseJson({ status: 'ok', message: 'Kết nối thành công!' }, cb);
  }
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    // ==================== TRANSACTIONS ====================
    if (action === 'getAll') {
      const sheet = getSheet('Transactions');
      if (!sheet) return responseJson({ success: true, data: [] }, cb);
      const data = sheet.getDataRange().getValues();
      const transactions = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          transactions.push({
            id: String(data[i][0]),
            type: String(data[i][1] || 'expense'),
            amount: Number(data[i][2] || 0),
            category: String(data[i][3] || 'other'),
            note: String(data[i][4] || ''),
            date: String(data[i][5] || ''),
            memberId: String(data[i][6] || ''),
            createdAt: String(data[i][7] || ''),
            beneficiaryId: String(data[i][8] || '')
          });
        }
      }
      return responseJson({ success: true, data: transactions }, cb);
    }
    
    if (action === 'add') {
      const obj = JSON.parse(e.parameter.data);
      getSheet('Transactions').appendRow([
        obj.id || '', obj.type || '', obj.amount || 0,
        obj.category || '', obj.note || '', obj.date || '',
        obj.memberId || '', obj.createdAt || new Date().toISOString(),
        obj.beneficiaryId || ''
      ]);
      return responseJson({ success: true, message: 'Đã thêm giao dịch' }, cb);
    }
    
    if (action === 'update') {
      const id = e.parameter.id;
      const obj = JSON.parse(e.parameter.data);
      const sheet = getSheet('Transactions');
      const rowNum = findRowById(sheet, id);
      if (rowNum > -1) {
        sheet.getRange(rowNum, 1, 1, 9).setValues([[
          obj.id || id, obj.type || '', obj.amount || 0,
          obj.category || '', obj.note || '', obj.date || '',
          obj.memberId || '', obj.createdAt || '', obj.beneficiaryId || ''
        ]]);
        return responseJson({ success: true }, cb);
      }
      throw new Error('Không tìm thấy giao dịch');
    }
    
    if (action === 'delete') {
      const sheet = getSheet('Transactions');
      const rowNum = findRowById(sheet, e.parameter.id);
      if (rowNum > -1) { sheet.deleteRow(rowNum); return responseJson({ success: true }, cb); }
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
          t.memberId || '', t.createdAt || '', t.beneficiaryId || ''
        ]);
        sheet.getRange(2, 1, rows.length, 9).setValues(rows);
      }
      return responseJson({ success: true, message: 'Đã đồng bộ giao dịch' }, cb);
    }
    
    // ==================== LOANS ====================
    if (action === 'getLoans') {
      const sheet = getSheet('Loans');
      if (!sheet) return responseJson({ success: true, data: [] }, cb);
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
      return responseJson({ success: true, data: loans }, cb);
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
      return responseJson({ success: true, message: 'Đã đồng bộ khoản vay' }, cb);
    }
    
    // ==================== MEMBERS ====================
    if (action === 'getMembers') {
      const sheet = getSheet('Members');
      if (!sheet) return responseJson({ success: true, data: [] }, cb);
      const data = sheet.getDataRange().getValues();
      const members = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          members.push({
            id: String(data[i][0]),
            name: String(data[i][1] || ''),
            avatar: String(data[i][2] || '👤'),
            color: String(data[i][3] || '#e77d3e'),
            avatarImg: String(data[i][4] || ''),
            avatarId: String(data[i][5] || '')
          });
        }
      }
      return responseJson({ success: true, data: members }, cb);
    }
    
    if (action === 'syncMembers') {
      const members = JSON.parse(e.parameter.data);
      const sheet = getSheet('Members');
      if (!sheet) return responseJson({ success: false, error: 'Sheet Members không tồn tại' }, cb);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      if (members && members.length > 0) {
        const rows = members.map(m => [
          m.id || '',
          m.name || '',
          m.avatar || '',
          m.color || '',
          m.avatarImg || '',
          m.avatarId || ''
        ]);
        sheet.getRange(2, 1, rows.length, 6).setValues(rows);
      }
      return responseJson({ success: true, message: 'Đã đồng bộ thành viên' }, cb);
    }
    
    // ==================== BUDGETS ====================
    if (action === 'getBudgets') {
      const sheet = getSheet('Budgets');
      if (!sheet) return responseJson({ success: true, data: {} }, cb);
      const data = sheet.getDataRange().getValues();
      const budgets = {};
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          budgets[String(data[i][0])] = Number(data[i][1] || 0);
        }
      }
      return responseJson({ success: true, data: budgets }, cb);
    }

    if (action === 'syncBudgets') {
      const budgets = JSON.parse(e.parameter.data);
      const sheet = getSheet('Budgets');
      if (!sheet) return responseJson({ success: false, error: 'Sheet Budgets không tồn tại' }, cb);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      if (budgets && typeof budgets === 'object') {
        const rows = Object.keys(budgets).map(k => [k, Number(budgets[k] || 0), new Date().toISOString()]);
        if (rows.length > 0) {
          sheet.getRange(2, 1, rows.length, 3).setValues(rows);
        }
      }
      return responseJson({ success: true, message: 'Đã đồng bộ ngân sách' }, cb);
    }

    // ==================== CUSTOM CATEGORIES ====================
    if (action === 'getCustomCats') {
      const sheet = getSheet('CustomCategories');
      if (!sheet) return responseJson({ success: true, data: [] }, cb);
      const data = sheet.getDataRange().getValues();
      const cats = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          cats.push({
            id: String(data[i][0]),
            label: String(data[i][1] || ''),
            emoji: String(data[i][2] || '📦'),
            type: String(data[i][3] || 'expense'),
            createdAt: String(data[i][4] || '')
          });
        }
      }
      return responseJson({ success: true, data: cats }, cb);
    }

    if (action === 'syncCustomCats') {
      const cats = JSON.parse(e.parameter.data);
      const sheet = getSheet('CustomCategories');
      if (!sheet) return responseJson({ success: false, error: 'Sheet CustomCategories không tồn tại' }, cb);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      if (cats && cats.length > 0) {
        const rows = cats.map(c => [c.id || '', c.label || '', c.emoji || '📦', c.type || 'expense', c.createdAt || new Date().toISOString()]);
        sheet.getRange(2, 1, rows.length, 5).setValues(rows);
      }
      return responseJson({ success: true, message: 'Đã đồng bộ danh mục tùy chỉnh' }, cb);
    }

    // ==================== SAVINGS GOALS & LOGS ====================
    if (action === 'getSavings') {
      const gSheet = getSheet('SavingsGoals');
      const lSheet = getSheet('SavingsLogs');
      const goals = [];
      const logs = [];

      if (gSheet) {
        const data = gSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][0]) {
            goals.push({
              id: String(data[i][0]),
              name: String(data[i][1] || ''),
              emoji: String(data[i][2] || '🐷'),
              targetAmount: Number(data[i][3] || 0),
              currentAmount: Number(data[i][4] || 0),
              targetDate: String(data[i][5] || ''),
              memberId: String(data[i][6] || 'family'),
              createdAt: String(data[i][7] || '')
            });
          }
        }
      }

      if (lSheet) {
        const data = lSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][0]) {
            logs.push({
              id: String(data[i][0]),
              goalId: String(data[i][1] || ''),
              goalName: String(data[i][2] || ''),
              type: String(data[i][3] || 'deposit'),
              amount: Number(data[i][4] || 0),
              memberId: String(data[i][5] || 'family'),
              date: String(data[i][6] || ''),
              note: String(data[i][7] || ''),
              createdAt: String(data[i][8] || '')
            });
          }
        }
      }

      return responseJson({ success: true, data: { goals, logs } }, cb);
    }

    if (action === 'syncSavings') {
      const payload = typeof e.parameter.goals === 'string' ? JSON.parse(e.parameter.goals) : e.parameter.goals;
      const goals = Array.isArray(payload) ? payload : (e.parameter.goals ? JSON.parse(e.parameter.goals) : []);
      const logs = e.parameter.logs ? (typeof e.parameter.logs === 'string' ? JSON.parse(e.parameter.logs) : e.parameter.logs) : [];

      const gSheet = getSheet('SavingsGoals');
      if (gSheet) {
        const lr = gSheet.getLastRow();
        if (lr > 1) gSheet.getRange(2, 1, lr - 1, gSheet.getLastColumn()).clearContent();
        if (goals && goals.length > 0) {
          const rows = goals.map(g => [g.id || '', g.name || '', g.emoji || '🐷', Number(g.targetAmount || 0), Number(g.currentAmount || 0), g.targetDate || '', g.memberId || 'family', g.createdAt || '']);
          gSheet.getRange(2, 1, rows.length, 8).setValues(rows);
        }
      }

      const lSheet = getSheet('SavingsLogs');
      if (lSheet) {
        const lr = lSheet.getLastRow();
        if (lr > 1) lSheet.getRange(2, 1, lr - 1, lSheet.getLastColumn()).clearContent();
        if (logs && logs.length > 0) {
          const rows = logs.map(l => [l.id || '', l.goalId || '', l.goalName || '', l.type || 'deposit', Number(l.amount || 0), l.memberId || 'family', l.date || '', l.note || '', l.createdAt || '']);
          lSheet.getRange(2, 1, rows.length, 9).setValues(rows);
        }
      }

      return responseJson({ success: true, message: 'Đã đồng bộ tiết kiệm' }, cb);
    }

    // ==================== BANK NOTIFICATION WEBHOOK & PENDING INBOX ====================
    if (action === 'bankNotification') {
      const rawText = String(e.parameter.text || e.parameter.body || e.parameter.content || e.parameter.message || '');
      const title = String(e.parameter.title || e.parameter.sender || '');
      const bank = String(e.parameter.bank || title || 'Ngân hàng');
      const member = String(e.parameter.member || 'dad');

      if (!rawText && !title) {
        return responseJson({ success: false, error: 'Thiếu nội dung thông báo' }, cb);
      }

      const classified = classifyBankNotification(rawText, title, bank);
      const pendingSheet = getSheet('PendingTransactions');
      if (!pendingSheet) return responseJson({ success: false, error: 'Sheet PendingTransactions không tồn tại' }, cb);

      const id = 'pend_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
      const nowStr = new Date().toISOString();
      const row = [
        id,
        bank,
        classified.type,
        classified.amount,
        classified.note,
        classified.category,
        member,
        nowStr,
        'pending',
        nowStr
      ];

      pendingSheet.appendRow(row);
      return responseJson({
        success: true,
        message: 'Đã nhận thông báo ngân hàng',
        transaction: {
          id: id,
          bank: bank,
          type: classified.type,
          amount: classified.amount,
          note: classified.note,
          category: classified.category,
          memberId: member,
          date: nowStr,
          status: 'pending'
        }
      }, cb);
    }

    if (action === 'getPending') {
      const sheet = getSheet('PendingTransactions');
      if (!sheet) return responseJson({ success: true, data: [] }, cb);
      const data = sheet.getDataRange().getValues();
      const pending = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && String(data[i][8]) === 'pending') {
          pending.push({
            id: String(data[i][0]),
            bank: String(data[i][1] || 'Ngân hàng'),
            type: String(data[i][2] || 'expense'),
            amount: Number(data[i][3] || 0),
            note: String(data[i][4] || ''),
            category: String(data[i][5] || 'other_expense'),
            memberId: String(data[i][6] || 'family'),
            date: String(data[i][7] || ''),
            status: String(data[i][8] || 'pending'),
            createdAt: String(data[i][9] || '')
          });
        }
      }
      return responseJson({ success: true, data: pending }, cb);
    }

    if (action === 'approvePending') {
      const pId = String(e.parameter.pendingId || e.parameter.id);
      const data = typeof e.parameter.data === 'string' ? JSON.parse(e.parameter.data) : e.parameter.data;

      // 1. Add to Transactions sheet
      const tSheet = getSheet('Transactions');
      if (tSheet && data) {
        const transId = data.id || ('trans_' + new Date().getTime());
        tSheet.appendRow([
          transId,
          data.type || 'expense',
          Number(data.amount || 0),
          data.category || 'other_expense',
          data.note || '',
          data.date || new Date().toISOString().slice(0, 10),
          data.memberId || 'dad',
          new Date().toISOString(),
          data.beneficiaryId || data.memberId || 'family'
        ]);
      }

      // 2. Mark pending row as approved
      const pSheet = getSheet('PendingTransactions');
      if (pSheet && pId) {
        const r = findRowById(pSheet, pId);
        if (r > -1) {
          pSheet.getRange(r, 9).setValue('approved');
        }
      }

      return responseJson({ success: true, message: 'Đã duyệt giao dịch vào sổ' }, cb);
    }

    if (action === 'deletePending') {
      const pId = String(e.parameter.id || e.parameter.pendingId);
      const pSheet = getSheet('PendingTransactions');
      if (pSheet && pId) {
        const r = findRowById(pSheet, pId);
        if (r > -1) {
          pSheet.getRange(r, 9).setValue('rejected');
        }
      }
      return responseJson({ success: true, message: 'Đã bỏ qua giao dịch' }, cb);
    }
    
    return responseJson({ success: false, error: 'Hành động không hợp lệ' }, cb);
    
  } catch (error) {
    return responseJson({ success: false, error: error.toString() }, cb);
  } finally {
    lock.releaseLock();
  }
}

// ==================== SMART VIETNAMESE CLASSIFIER ====================
function classifyBankNotification(rawText, title, bank) {
  const fullText = (title + ' ' + rawText + ' ' + (bank || '')).toLowerCase();
  
  // 1. Extract Amount
  let amount = 0;
  const amtMatch = rawText.match(/(?:[\+\-]|gd:?\s*[\+\-]?|ps:?\s*[\+\-]?)?\s*([\d\.\,]{3,15})\s*(?:vnd|vnđ|đ|d\b)/i) ||
                   rawText.match(/([\d\.\,]{4,15})\s*(?:vnd|vnđ|đ|d\b)/i);
  if (amtMatch && amtMatch[1]) {
    const cleanNum = amtMatch[1].replace(/[\.\,]/g, '');
    amount = Number(cleanNum) || 0;
  }

  // 2. Extract Type
  let type = 'expense';
  if (rawText.includes('+') || fullText.includes('nhan tien') || fullText.includes('nhận tiền') || fullText.includes('cong tien') || fullText.includes('cộng tiền') || fullText.includes('credit')) {
    type = 'income';
  } else if (rawText.includes('-') || fullText.includes('tru tien') || fullText.includes('trừ tiền') || fullText.includes('thanh toan') || fullText.includes('thanh toán') || fullText.includes('debit')) {
    type = 'expense';
  }

  // 3. Smart Category Prediction based on keywords
  let category = type === 'income' ? 'salary' : 'other_expense';
  if (type === 'expense') {
    if (/highlands|phở|bún|cơm|quán|cafe|cà phê|trà sữa|starbucks|kfc|lotteria|pizza|food|shopeefood|grabfood|befood|baemin|tokyo deli|haidilao|gogi|kichi|nhà hàng|bánh mì|ăn sáng|ăn trưa|ăn tối|lẩu|nướng/i.test(fullText)) {
      category = 'food';
    } else if (/grab|be |xanh sm|taxi|xăng|petrolimex|pvoil|gửi xe|giữ xe|vé xe|cầu đường|epass|vetc|đỗ xe/i.test(fullText)) {
      category = 'transport';
    } else if (/shopee|lazada|tiki|sendo|winmart|co\.?op|bách hóa|bach hoa|siêu thị|uniqlo|zara|h&m|mua sắm|shop|store|mall|quần áo|giày|mỹ phẩm/i.test(fullText)) {
      category = 'shopping';
    } else if (/evn|điện lực|tiền điện|tiền nước|cấp nước|viettel|vnpt|fpt|internet|wifi|mobifone|vinaphone|chung cư|phí quản lý|vệ sinh|rác/i.test(fullText)) {
      category = 'bills';
    } else if (/thuốc|pharmacity|long châu|an khang|bệnh viện|phòng khám|bác sĩ|nha khoa|răng|khám bệnh|y tế/i.test(fullText)) {
      category = 'health';
    } else if (/học phí|trường|mầm non|tiểu học|tiếng anh|ila|vus|apollo|sách|vở|dụng cụ học tập/i.test(fullText)) {
      category = 'education';
    } else if (/cgv|bhd|lotte cinema|rạp|vé xem phim|netflix|spotify|youtube|steam|game|playstation|du lịch|khách sạn|resort|vé máy bay/i.test(fullText)) {
      category = 'entertainment';
    } else if (/nội thất|điện máy|điện thoại|laptop|sửa nhà|decor|gia dụng/i.test(fullText)) {
      category = 'house';
    } else if (/chứng khoán|cổ phiếu|ssi|vps|tcbs|vndirect|tiết kiệm|gửi tiền|vàng|sjc|doji/i.test(fullText)) {
      category = 'invest';
    }
  } else {
    // Income prediction
    if (/lương|salary|payroll|thu nhập|thưởng|bonus/i.test(fullText)) {
      category = 'salary';
    } else if (/đầu tư|cổ tức|lãi|tiết kiệm|interest/i.test(fullText)) {
      category = 'investment';
    } else {
      category = 'other_income';
    }
  }

  // 4. Clean note
  let cleanNote = rawText.replace(/\r?\n|\r/g, ' ').slice(0, 100);

  return {
    amount: amount,
    type: type,
    category: category,
    note: cleanNote
  };
}
