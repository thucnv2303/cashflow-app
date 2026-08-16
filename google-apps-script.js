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

function safeParseJson(data) {
  if (!data) return null;
  if (typeof data === 'object') return data;
  if (typeof data === 'string') {
    try { return JSON.parse(data); } catch(e) { return null; }
  }
  return null;
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

  if (action === 'portal' || action === 'auth') {
    const memberId = e.parameter.member || 'mom';
    return renderMemberPortalHtml(memberId);
  }

  if (action === 'registerMemberEmail') {
    const email = e.parameter.email || '';
    const member = e.parameter.member || 'mom';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let meSheet = ss.getSheetByName('MemberEmails');
    if (!meSheet) {
      meSheet = ss.insertSheet('MemberEmails');
      meSheet.appendRow(['Email', 'Role', 'RegisteredAt', 'Status']);
      meSheet.setFrozenRows(1);
      meSheet.getRange('A1:D1').setFontWeight('bold');
    }
    meSheet.appendRow([email, member, new Date().toISOString(), 'active']);
    return responseJson({ success: true, message: 'Đã đăng ký email ' + email }, cb);
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
      const obj = safeParseJson(e.parameter.data) || {};
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
      const obj = safeParseJson(e.parameter.data) || {};
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
      const transactions = safeParseJson(e.parameter.data) || [];
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
      const loans = safeParseJson(e.parameter.data) || [];
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
      const members = safeParseJson(e.parameter.data) || [];
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
      const budgets = safeParseJson(e.parameter.data) || {};
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
      const cats = safeParseJson(e.parameter.data) || [];
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
      const isDirect = String(e.parameter.direct || e.parameter.auto || '') === '1' || String(e.parameter.direct || '') === 'true';

      if (!rawText && !title) {
        return responseJson({ success: false, error: 'Thiếu nội dung thông báo' }, cb);
      }

      const classified = classifyBankNotification(rawText, title, bank);
      const nowStr = new Date().toISOString();
      const id = 'pend_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);

      // If direct mode enabled, write straight into Transactions sheet!
      if (isDirect && classified.amount > 0) {
        const transSheet = getSheet('Transactions');
        if (transSheet) {
          const transId = 'trans_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
          const dateOnly = nowStr.slice(0, 10);
          transSheet.appendRow([
            transId,
            classified.type,
            classified.amount,
            classified.category,
            classified.note || 'Nhập nhanh từ Shortcut',
            dateOnly,
            member,
            nowStr,
            'family'
          ]);
        }
      }

      // Also record in PendingTransactions
      const pendingSheet = getSheet('PendingTransactions');
      if (pendingSheet) {
        pendingSheet.appendRow([
          id,
          bank,
          classified.type,
          classified.amount,
          classified.note,
          classified.category,
          member,
          nowStr,
          isDirect ? 'approved' : 'pending',
          nowStr
        ]);
      }

      return responseJson({
        success: true,
        message: isDirect ? 'Đã ghi nhận trực tiếp vào sổ giao dịch! 🎉' : 'Đã nhận thông báo ngân hàng',
        transaction: {
          id: id,
          bank: bank,
          type: classified.type,
          amount: classified.amount,
          note: classified.note,
          category: classified.category,
          memberId: member,
          date: nowStr,
          status: isDirect ? 'approved' : 'pending'
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
      const pendingId = String(e.parameter.pendingId || e.parameter.id);
      const data = safeParseJson(e.parameter.data);
      const transSheet = getSheet('Transactions');
      if (transSheet && data) {
        const transId = data.id || ('trans_' + new Date().getTime());
        transSheet.appendRow([
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
      const pendingSheet = getSheet('PendingTransactions');
      if (pendingSheet && pendingId) {
        const rowNum = findRowById(pendingSheet, pendingId);
        if (rowNum > -1) {
          pendingSheet.getRange(rowNum, 9).setValue('approved');
        }
      }
      return responseJson({ success: true, message: 'Đã duyệt giao dịch' }, cb);
    }

    if (action === 'deletePending') {
      const pendingId = String(e.parameter.id || e.parameter.pendingId);
      const pendingSheet = getSheet('PendingTransactions');
      if (pendingSheet && pendingId) {
        const rowNum = findRowById(pendingSheet, pendingId);
        if (rowNum > -1) {
          pendingSheet.getRange(rowNum, 9).setValue('rejected');
        }
      }
      return responseJson({ success: true, message: 'Đã bỏ qua' }, cb);
    }

    if (action === 'scanGmail') {
      const res = scanBankEmails();
      return responseJson({
        success: true,
        addedCount: res.addedCount,
        message: 'Đã quét Gmail thành công! Đã thêm ' + res.addedCount + ' giao dịch.'
      }, cb);
    }

    return responseJson({ success: false, error: 'Unknown action: ' + action }, cb);
  } catch (err) {
    return responseJson({ success: false, error: err.toString() }, cb);
  } finally {
    lock.releaseLock();
  }
}

// ==================== SMART VIETNAMESE CLASSIFIER ====================
function classifyBankNotification(rawText, title, bank) {
  const fullText = (title + ' ' + rawText + ' ' + (bank || '')).toLowerCase();
  
  // 1. Extract Amount with support for k, tr, triệu, nghìn, cành, lít
  let amount = 0;
  
  const kMatch = rawText.match(/(\d+(?:[\.,]\d+)?)\s*(?:k|nghìn|nghin|ng|cành|canh)\b/i);
  const trMatch = rawText.match(/(\d+)(?:[\.,](\d+))?\s*(?:tr|triệu|trieu|m)\b/i) || rawText.match(/(\d+)\s*tr\s*(\d+)/i);
  const stdMatch = rawText.match(/(?:Số tiền|Số tiền GD|So tien|Giá trị GD|Số tiền giao dịch|Số dư thay đổi|PS|GD)[\:\s]*([\+\-]?\s*[\d\.\,]{3,15})\s*(?:vnd|vnđ|đ|d\b)/i) ||
                   rawText.match(/(?:[\+\-]|gd:?\s*[\+\-]?|ps:?\s*[\+\-]?)?\s*([\d\.\,]{3,15})\s*(?:vnd|vnđ|đ|d\b)/i) ||
                   rawText.match(/([\d\.\,]{4,15})\s*(?:vnd|vnđ|đ|d\b)/i);
  const plainMatch = rawText.match(/\b(\d{4,12})\b/);

  if (kMatch && kMatch[1]) {
    const num = parseFloat(kMatch[1].replace(',', '.'));
    amount = Math.round(num * 1000);
  } else if (trMatch) {
    if (trMatch[2]) {
      const whole = parseInt(trMatch[1], 10);
      const dec = trMatch[2].length === 1 ? parseInt(trMatch[2], 10) * 100000 : parseInt(trMatch[2], 10) * Math.pow(10, 6 - trMatch[2].length);
      amount = whole * 1000000 + dec;
    } else {
      const num = parseFloat(trMatch[1]);
      amount = Math.round(num * 1000000);
    }
  } else if (stdMatch && stdMatch[1]) {
    const cleanNum = stdMatch[1].replace(/[\.\,\s\+\-]/g, '');
    amount = Number(cleanNum) || 0;
  } else if (plainMatch && plainMatch[1]) {
    amount = Number(plainMatch[1]) || 0;
  }

  // 2. Extract Type
  let type = 'expense';
  if (rawText.includes('+') || fullText.includes('nhan tien') || fullText.includes('nhận tiền') || fullText.includes('cong tien') || fullText.includes('cộng tiền') || fullText.includes('thu ') || fullText.includes('credit') || fullText.includes('tang so du')) {
    type = 'income';
  } else if (rawText.includes('-') || fullText.includes('tru tien') || fullText.includes('trừ tiền') || fullText.includes('thanh toan') || fullText.includes('thanh toán') || fullText.includes('chi ') || fullText.includes('debit') || fullText.includes('giam so du')) {
    type = 'expense';
  }

  // 3. Smart Category Prediction based on keywords
  let category = type === 'income' ? 'salary' : 'other_expense';
  if (type === 'expense') {
    if (/ăn|uống|highlands|phở|bún|cơm|quán|cafe|cà phê|trà sữa|starbucks|kfc|lotteria|pizza|food|shopeefood|grabfood|befood|baemin|tokyo deli|haidilao|gogi|kichi|nhà hàng|bánh mì|ăn sáng|ăn trưa|ăn tối|lẩu|nướng|nhậu|bia|chè|bánh/i.test(fullText)) {
      category = 'food';
    } else if (/xăng|xang|grab|be |xanh sm|taxi|petrolimex|pvoil|gửi xe|giữ xe|vé xe|cầu đường|epass|vetc|đỗ xe|xe buýt|bus|rửa xe|sửa xe/i.test(fullText)) {
      category = 'transport';
    } else if (/mua|sắm|shopee|lazada|tiki|sendo|winmart|co\.?op|bách hóa|bach hoa|siêu thị|uniqlo|zara|h&m|mall|quần|áo|giày|dép|túi|mỹ phẩm|son/i.test(fullText)) {
      category = 'shopping';
    } else if (/điện|nước|evn|điện lực|tiền điện|tiền nước|cấp nước|viettel|vnpt|fpt|internet|wifi|mobifone|vinaphone|chung cư|phí quản lý|vệ sinh|rác|gas/i.test(fullText)) {
      category = 'bills';
    } else if (/thuốc|thuoc|pharmacity|long châu|an khang|bệnh viện|benh vien|phòng khám|bác sĩ|nha khoa|răng|khám|y tế|clinic|spa/i.test(fullText)) {
      category = 'health';
    } else if (/học|hoc|học phí|trường|mầm non|tiểu học|tiếng anh|ila|vus|apollo|sách|vở|dụng cụ học tập|khóa học|course/i.test(fullText)) {
      category = 'education';
    } else if (/xem phim|cgv|bhd|lotte cinema|rạp|vé xem|netflix|spotify|youtube|steam|game|playstation|du lịch|khách sạn|resort|vé máy bay|karaoke|bowling|billiard/i.test(fullText)) {
      category = 'entertainment';
    } else if (/nội thất|điện máy|điện thoại|laptop|sửa nhà|decor|gia dụng|thuê nhà|tiền nhà/i.test(fullText)) {
      category = 'house';
    } else if (/chứng khoán|cổ phiếu|ssi|vps|tcbs|vndirect|tiết kiệm|gửi tiền|vàng|sjc|doji/i.test(fullText)) {
      category = 'invest';
    }
  } else {
    if (/lương|salary|payroll|thu nhập|thưởng|bonus/i.test(fullText)) {
      category = 'salary';
    } else if (/đầu tư|cổ tức|lãi|tiết kiệm|interest/i.test(fullText)) {
      category = 'investment';
    } else {
      category = 'other_income';
    }
  }

  // 4. Clean note - Extract specific merchant / ND field if present
  let cleanNote = '';
  const ndMatch = rawText.match(/(?:ND|Noi dung|Nội dung|GD tại|tại|Ref|Desc)[\:\s]+([^.]+?)(?:\s*(?:vao luc|vào lúc|ngay|ngày|luc|lúc|Han muc|Hạn mức|SD|Số dư|\.|$))/i);
  if (ndMatch && ndMatch[1]) {
    cleanNote = ndMatch[1].trim();
  } else {
    cleanNote = rawText
      .replace(/(?:Số tiền|Số tiền GD|So tien|Giá trị GD|Số tiền giao dịch|Số dư thay đổi|PS|GD)[\:\s]*[\+\-]?\s*[\d\.\,]{3,15}\s*(?:vnd|vnđ|đ|d\b)/gi, '')
      .replace(/(\d+(?:[\.,]\d+)?)\s*(?:k|nghìn|nghin|ng|cành|canh)\b/gi, '')
      .replace(/(\d+(?:[\.,]\d+)?)\s*(?:tr|triệu|trieu|m)\b/gi, '')
      .replace(/\b\d{4,12}\b/g, '')
      .replace(/\r?\n|\r/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  if (!cleanNote || cleanNote.length < 2) {
    cleanNote = rawText.slice(0, 100);
  }

  return {
    amount: amount,
    type: type,
    category: category,
    note: cleanNote
  };
}

// ==================== GMAIL CLOUD AUTO-SCAN ====================
function scanBankEmails(targetMember) {
  const member = targetMember || 'dad';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let pSheet = ss.getSheetByName('PendingTransactions');
  if (!pSheet) {
    setupSheet();
    pSheet = ss.getSheetByName('PendingTransactions');
  }

  let label = GmailApp.getUserLabelByName('CashFlow_Processed');
  if (!label) {
    label = GmailApp.createLabel('CashFlow_Processed');
  }

  const queries = [
    'from:(vpbank.com.vn OR techcombank.com.vn OR mbbank.com.vn OR vietcombank.com.vn OR timo.vn OR acb.com.vn OR vib.com.vn OR tpbank.com.vn OR bidv.com.vn OR cake.vn) -label:CashFlow_Processed newer_than:3d',
    'subject:("Biến động số dư" OR "Thông báo giao dịch" OR "Transaction Alert" OR "Thông báo thay đổi số dư") -label:CashFlow_Processed newer_than:3d'
  ];

  let addedCount = 0;
  const processedThreadIds = {};

  for (let q = 0; q < queries.length; q++) {
    const threads = GmailApp.search(queries[q], 0, 15);
    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      if (processedThreadIds[thread.getId()]) continue;
      processedThreadIds[thread.getId()] = true;

      const messages = thread.getMessages();
      for (let j = 0; j < messages.length; j++) {
        const msg = messages[j];
        const subject = msg.getSubject() || '';
        const sender = msg.getFrom() || '';
        const bodyText = msg.getPlainBody() || '';
        const date = msg.getDate();

        let bank = 'Ngân hàng';
        if (/vpbank/i.test(sender) || /vpbank/i.test(subject) || /vpbank/i.test(bodyText)) bank = 'VPBank';
        else if (/techcombank/i.test(sender) || /techcombank/i.test(subject)) bank = 'Techcombank';
        else if (/mbbank|mb bank/i.test(sender) || /mbbank|mb bank/i.test(subject)) bank = 'MB Bank';
        else if (/vietcombank|vcb/i.test(sender) || /vietcombank/i.test(subject)) bank = 'Vietcombank';
        else if (/timo/i.test(sender) || /timo/i.test(subject)) bank = 'Timo';
        else if (/acb/i.test(sender) || /acb/i.test(subject)) bank = 'ACB';
        else if (/vib/i.test(sender) || /vib/i.test(subject)) bank = 'VIB';
        else if (/tpbank/i.test(sender) || /tpbank/i.test(subject)) bank = 'TPBank';
        else if (/bidv/i.test(sender) || /bidv/i.test(subject)) bank = 'BIDV';
        else if (/cake/i.test(sender) || /cake/i.test(subject)) bank = 'CAKE';
        else if (/momo/i.test(sender) || /momo/i.test(subject)) bank = 'MoMo';

        const classified = classifyBankNotification(bodyText, subject, bank);
        if (classified.amount > 0) {
          const id = 'email_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
          const dateStr = date ? date.toISOString() : new Date().toISOString();
          pSheet.appendRow([
            id,
            bank,
            classified.type,
            classified.amount,
            classified.note || subject,
            classified.category,
            member,
            dateStr,
            'pending',
            dateStr
          ]);
          addedCount++;
        }
      }
      thread.addLabel(label);
    }
  }

  return { success: true, addedCount: addedCount };
}

function scanMemberBankEmails(memberId) {
  return scanBankEmails(memberId || 'mom');
}

// Trigger functions for each member role
function triggerScanMom() { return scanBankEmails('mom'); }
function triggerScanDad() { return scanBankEmails('dad'); }
function triggerScanChild() { return scanBankEmails('child'); }
function triggerScanOther() { return scanBankEmails('other'); }

// Setup auto-trigger for a specific member
function setupMemberTrigger(memberId) {
  const funcName = 'triggerScan' + memberId.charAt(0).toUpperCase() + memberId.slice(1);
  // Remove existing trigger for this member
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === funcName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create new trigger every 5 minutes
  ScriptApp.newTrigger(funcName)
    .timeBased()
    .everyMinutes(5)
    .create();
  return 'Đã thiết lập tự động quét mỗi 5 phút cho ' + memberId;
}

// Bật tự động chạy mỗi 1 phút trên Cloud Google (cho chủ tài khoản)
function setupGmailTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scanBankEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('scanBankEmails')
    .timeBased()
    .everyMinutes(1)
    .create();
  return "Đã thiết lập tự động quét Gmail mỗi 1 phút thành công!";
}

// ==================== MEMBER AUTH & SCAN PORTAL HTML ====================
function renderMemberPortalHtml(memberId) {
  const label = memberId === 'mom' ? 'Vợ 🌸' : (memberId === 'dad' ? 'Chồng 👔' : (memberId === 'child' ? 'Con 🧒' : 'Thành viên gia đình'));
  const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>CashFlow - Kích hoạt Quét Email</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #0f172a; color: #f8fafc; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; padding: 16px; box-sizing: border-box; }
    .card { background: #1e293b; border: 1px solid #334155; border-radius: 20px; padding: 28px; max-width: 440px; width: 100%; text-align: center; box-shadow: 0 10px 30px rgba(0,0,0,0.5); }
    .emoji { font-size: 3rem; margin-bottom: 12px; }
    h1 { font-size: 1.3rem; margin: 0 0 8px 0; color: #f8fafc; }
    p { font-size: 0.85rem; color: #94a3b8; line-height: 1.5; margin: 0 0 20px 0; }
    .btn { background: #e77d3e; color: #fff; border: none; border-radius: 12px; padding: 14px 24px; font-size: 1rem; font-weight: 700; cursor: pointer; width: 100%; transition: 0.2s; box-shadow: 0 4px 14px rgba(231, 125, 62, 0.4); }
    .btn:hover { background: #d97706; transform: translateY(-1px); }
    .btn:disabled { opacity: 0.6; cursor: not-allowed; }
    .status { margin-top: 18px; padding: 14px; border-radius: 10px; font-size: 0.85rem; display: none; line-height: 1.5; }
    .status.success { background: rgba(34, 197, 94, 0.15); color: #4ade80; border: 1px solid rgba(34, 197, 94, 0.3); }
    .status.loading { background: rgba(59, 130, 246, 0.15); color: #60a5fa; border: 1px solid rgba(59, 130, 246, 0.3); }
    .status.error { background: rgba(248, 113, 113, 0.15); color: #f87171; border: 1px solid rgba(248, 113, 113, 0.3); }
    .steps { text-align: left; margin-top: 14px; font-size: 0.8rem; color: #94a3b8; line-height: 1.6; }
    .steps .done { color: #4ade80; }
    .steps .active { color: #60a5fa; font-weight: 600; }
  </style>
</head>
<body>
  <div class="card">
    <div class="emoji">🌸</div>
    <h1>Kích hoạt Quét Email Ngân Hàng</h1>
    <p>Xin chào <strong>${label}</strong>!<br>Bấm nút bên dưới để cho phép hệ thống đọc email biến động ngân hàng của bạn và đưa vào sổ chi tiêu gia đình.<br><em style="font-size:0.78rem;">Hệ thống chỉ đọc email từ ngân hàng, không đọc email cá nhân khác.</em></p>
    
    <button class="btn" id="scanBtn" onclick="startActivation()">🚀 Kích hoạt & Quét Email Ngân Hàng</button>
    
    <div id="statusBox" class="status"></div>
    
    <div id="stepsBox" class="steps" style="display:none;">
      <div id="step1">⬜ Bước 1: Cấp quyền đọc Gmail...</div>
      <div id="step2">⬜ Bước 2: Quét email ngân hàng...</div>
      <div id="step3">⬜ Bước 3: Thiết lập quét tự động 24/7...</div>
    </div>
  </div>

  <script>
    function startActivation() {
      var btn = document.getElementById('scanBtn');
      var box = document.getElementById('statusBox');
      var steps = document.getElementById('stepsBox');
      btn.disabled = true;
      btn.textContent = 'Đang kích hoạt... ⏳';
      steps.style.display = 'block';
      
      // Step 1: Permission (implicit when calling GmailApp)
      document.getElementById('step1').innerHTML = '🔄 Bước 1: Đang xin cấp quyền đọc Gmail...';
      document.getElementById('step1').className = 'active';
      box.style.display = 'block';
      box.className = 'status loading';
      box.textContent = 'Đang kết nối với Gmail của bạn...';
      
      // Step 2: Scan emails
      google.script.run
        .withSuccessHandler(function(scanRes) {
          document.getElementById('step1').innerHTML = '✅ Bước 1: Đã cấp quyền đọc Gmail';
          document.getElementById('step1').className = 'done';
          document.getElementById('step2').innerHTML = '✅ Bước 2: Đã quét xong! Tìm thấy ' + (scanRes.addedCount || 0) + ' giao dịch mới';
          document.getElementById('step2').className = 'done';
          document.getElementById('step3').innerHTML = '🔄 Bước 3: Đang thiết lập quét tự động...';
          document.getElementById('step3').className = 'active';
          box.textContent = 'Đang thiết lập chế độ quét tự động 24/7...';
          
          // Step 3: Setup auto-trigger
          google.script.run
            .withSuccessHandler(function(triggerRes) {
              document.getElementById('step3').innerHTML = '✅ Bước 3: Đã thiết lập quét tự động 24/7!';
              document.getElementById('step3').className = 'done';
              btn.disabled = false;
              btn.textContent = '🔄 Quét lại ngay';
              box.className = 'status success';
              box.innerHTML = '🎉 <strong>Hoàn tất!</strong><br>Đã thêm ' + (scanRes.addedCount || 0) + ' giao dịch mới.<br>Hệ thống sẽ tự động quét email ngân hàng của bạn mỗi 5 phút.<br><br><strong>Bạn có thể đóng trang này.</strong> Mọi thứ đã được thiết lập xong!';
            })
            .withFailureHandler(function(err) {
              document.getElementById('step3').innerHTML = '⚠️ Bước 3: Không thể tự động hóa (bạn có thể quét thủ công)';
              btn.disabled = false;
              btn.textContent = '🔄 Quét lại ngay';
              box.className = 'status success';
              box.innerHTML = '🎉 Đã quét xong ' + (scanRes.addedCount || 0) + ' giao dịch!<br>Lưu ý: Quét tự động chưa kích hoạt. Bạn có thể bấm nút để quét thủ công.';
            })
            .setupMemberTrigger('${memberId}');
        })
        .withFailureHandler(function(err) {
          document.getElementById('step1').innerHTML = '✅ Bước 1: Đã cấp quyền';
          document.getElementById('step1').className = 'done';
          document.getElementById('step2').innerHTML = '❌ Bước 2: Lỗi khi quét';
          btn.disabled = false;
          btn.textContent = '🔄 Thử lại';
          box.className = 'status error';
          box.textContent = 'Lỗi: ' + err.toString();
        })
        .scanMemberBankEmails('${memberId}');
    }
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle('CashFlow - Kích hoạt Quét Email')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
