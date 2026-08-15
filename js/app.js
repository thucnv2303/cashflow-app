// Google Apps Script code to embed (for copy button)
const APPS_SCRIPT_CODE = `/**
 * Hướng dẫn cài đặt:
 * 1. Mở Google Sheets, tạo một Spreadsheet mới.
 * 2. Menu: Tiện ích mở rộng > Apps Script.
 * 3. Xóa code có sẵn, dán toàn bộ nội dung này vào.
 * 4. Chọn hàm "setupSheet" từ dropdown > nhấn ▶️ Run > cấp quyền.
 * 5. Deploy > New deployment > Web app > Anyone > Deploy > copy URL.
 * 6. Dán URL vào Cài đặt trong app CashFlow.
 */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Transactions');
  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
    sheet.appendRow(['ID', 'Loại', 'Số tiền', 'Danh mục', 'Ghi chú', 'Ngày', 'Ngày tạo']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:G1').setFontWeight('bold');
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Transactions');
  if (!sheet) { setupSheet(); sheet = ss.getSheetByName('Transactions'); }
  return sheet;
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) return i + 1;
  }
  return -1;
}

function rowToObject(row) {
  return { id: row[0], type: row[1], amount: Number(row[2]), category: row[3], note: row[4], date: row[5], createdAt: row[6] };
}

function objectToRow(obj) {
  return [obj.id||'', obj.type||'', obj.amount||0, obj.category||'', obj.note||'', obj.date||'', obj.createdAt||new Date().toISOString()];
}

function responseJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action || 'getAll';
  if (action === 'ping') return responseJson({ status: 'ok', message: 'Kết nối thành công!' });
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet();
    
    if (action === 'getAll') {
      const data = sheet.getDataRange().getValues();
      const transactions = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) transactions.push(rowToObject(data[i]));
      }
      return responseJson({ success: true, data: transactions });
    }
    else if (action === 'add') {
      const obj = JSON.parse(e.parameter.data);
      sheet.appendRow(objectToRow(obj));
      return responseJson({ success: true, message: 'Đã thêm giao dịch' });
    }
    else if (action === 'update') {
      const id = e.parameter.id;
      const obj = JSON.parse(e.parameter.data);
      const rowNum = findRowById(sheet, id);
      if (rowNum > -1) {
        const newRow = objectToRow(obj);
        sheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);
        return responseJson({ success: true, message: 'Đã cập nhật' });
      }
      throw new Error('Không tìm thấy giao dịch');
    }
    else if (action === 'delete') {
      const rowNum = findRowById(sheet, e.parameter.id);
      if (rowNum > -1) { sheet.deleteRow(rowNum); return responseJson({ success: true }); }
      throw new Error('Không tìm thấy giao dịch');
    }
    else if (action === 'sync') {
      const transactions = JSON.parse(e.parameter.data);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      if (transactions && transactions.length > 0) {
        const rows = transactions.map(t => objectToRow(t));
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      }
      return responseJson({ success: true, message: 'Đã đồng bộ' });
    }
    return responseJson({ success: false, error: 'Hành động không hợp lệ' });
  } catch (error) {
    return responseJson({ success: false, error: error.toString() });
  } finally {
    lock.releaseLock();
  }
}`;

const App = {
  currentPage: 'dashboard',
  currentMonth: getCurrentMonth(),
  analyticsMonth: getCurrentMonth(),
  editingId: null,
  selectedType: 'expense',
  selectedCategory: null,

  // Initialize app
  init() {
    this.initTheme();
    this.bindEvents();
    this.initFilterMonth();
    this.initSyncStatus();
    this.renderCurrentPage();
    // Populate the hidden textarea with Apps Script code
    const codeEl = document.getElementById('appsScriptCode');
    if (codeEl) codeEl.value = APPS_SCRIPT_CODE;
  },
  
  // ==================== THEME ====================
  initTheme() { 
    const theme = Storage.getTheme();
    document.documentElement.setAttribute('data-theme', theme);
    this.updateThemeToggleIcon(theme);
    Charts.setTheme(theme === 'dark');
  },
  
  updateThemeToggleIcon(theme) {
    const icon = theme === 'dark' ? '🌙' : '☀️';
    document.querySelectorAll('.theme-icon').forEach(el => {
      el.textContent = icon;
    });
  },
  
  toggleTheme() {
    const currentTheme = Storage.getTheme();
    const newTheme = currentTheme === 'light' ? 'dark' : 'light';
    Storage.setTheme(newTheme);
    document.documentElement.setAttribute('data-theme', newTheme);
    this.updateThemeToggleIcon(newTheme);
    Charts.setTheme(newTheme === 'dark');
    this.renderCurrentPage();
  },
  
  // ==================== NAVIGATION ====================
  navigate(page) {
    this.currentPage = page;
    document.querySelectorAll('.nav-item').forEach(el => {
      el.classList.toggle('active', el.dataset.page === page);
    });
    document.querySelectorAll('.page').forEach(el => {
      el.classList.remove('active');
    });
    const pageEl = document.getElementById(`page-${page}`);
    if (pageEl) pageEl.classList.add('active');
    this.renderCurrentPage();
  },
  
  renderCurrentPage() {
    if (this.currentPage === 'dashboard') this.renderDashboard();
    else if (this.currentPage === 'transactions') this.renderTransactions();
    else if (this.currentPage === 'analytics') this.renderAnalytics();
    else if (this.currentPage === 'settings') this.renderSettings();
  },
  
  // ==================== SYNC STATUS ====================
  initSyncStatus() {
    this.updateSyncStatusBar();
    // If online mode, do initial sync
    if (Storage.isOnline()) {
      this.syncFromSheets(true);
    }
  },

  updateSyncStatusBar() {
    const bar = document.getElementById('syncStatus');
    if (!bar) return;
    
    if (Storage.isOnline()) {
      bar.style.display = 'flex';
      bar.className = 'sync-status connected';
      bar.querySelector('.sync-text').textContent = '🟢 Đã kết nối Google Sheets';
    } else {
      bar.style.display = 'none';
    }
  },

  async syncFromSheets(silent = false) {
    if (!Storage.isOnline()) return;
    
    const bar = document.getElementById('syncStatus');
    if (bar) {
      bar.style.display = 'flex';
      bar.className = 'sync-status';
      bar.querySelector('.sync-text').textContent = '🔄 Đang đồng bộ...';
    }
    
    try {
      const success = await Storage.syncFromSheets();
      if (bar) {
        bar.className = 'sync-status connected';
        bar.querySelector('.sync-text').textContent = '🟢 Đã đồng bộ';
      }
      if (success) this.renderCurrentPage();
      if (!silent) this.showToast('Đã đồng bộ dữ liệu từ Sheets ✅');
    } catch (e) {
      if (bar) {
        bar.className = 'sync-status offline';
        bar.querySelector('.sync-text').textContent = '🔴 Lỗi đồng bộ';
      }
      if (!silent) this.showToast('Lỗi đồng bộ: ' + e.message, 'error');
    }
  },

  // ==================== INIT FILTERS ====================
  initFilterMonth() {
    const filterMonth = document.getElementById('filterMonth');
    if (filterMonth) {
      const now = new Date();
      const y = now.getFullYear();
      const m = (now.getMonth() + 1).toString().padStart(2, '0');
      filterMonth.value = `${y}-${m}`;
    }
    this.populateCategoryFilter();
  },

  populateCategoryFilter() {
    const filterCategory = document.getElementById('filterCategory');
    if (!filterCategory) return;
    filterCategory.innerHTML = '<option value="all">Tất cả</option>';
    getAllCategories().forEach(cat => {
      filterCategory.innerHTML += `<option value="${cat.id}">${cat.emoji} ${cat.label}</option>`;
    });
  },

  // ==================== DASHBOARD ====================
  renderDashboard() {
    const { year, month } = this.currentMonth;
    const transactions = Storage.getByMonth(year, month);
    
    let income = 0, expense = 0;
    transactions.forEach(t => {
      if (t.type === 'income') income += t.amount;
      else expense += t.amount;
    });
    const balance = income - expense;
    const savingsRate = income > 0 ? ((income - expense) / income * 100) : 0;
    
    // Previous month comparison
    const prev = navigateMonth(year, month, -1);
    const prevTransactions = Storage.getByMonth(prev.year, prev.month);
    let prevIncome = 0, prevExpense = 0;
    prevTransactions.forEach(t => {
      if (t.type === 'income') prevIncome += t.amount;
      else prevExpense += t.amount;
    });
    
    // Update cards
    document.getElementById('totalIncome').textContent = formatCurrency(income);
    document.getElementById('totalExpense').textContent = formatCurrency(expense);
    document.getElementById('balance').textContent = formatCurrency(balance);
    document.getElementById('savingsRate').textContent = savingsRate.toFixed(1) + '%';
    
    this.updateChangeBadge('incomeChange', calcChange(income, prevIncome), false);
    this.updateChangeBadge('expenseChange', calcChange(expense, prevExpense), true);
    
    document.getElementById('currentMonth').textContent = formatMonthLabel(year, month);
    
    // Charts
    Charts.renderCategoryChart('categoryChart', transactions);
    const monthsData = Storage.getLastNMonths(6);
    Charts.renderMonthlyChart('monthlyChart', monthsData);
    Charts.renderTrendChart('trendChart', monthsData);
    
    this.renderRecentTransactions(transactions);
  },

  updateChangeBadge(elementId, changePercent, invertColor) {
    const el = document.getElementById(elementId);
    if (!el) return;
    const sign = changePercent >= 0 ? '+' : '';
    el.textContent = `${sign}${changePercent.toFixed(1)}%`;
    el.className = 'card-change';
    if (invertColor) {
      el.classList.add(changePercent <= 0 ? 'positive' : 'negative');
    } else {
      el.classList.add(changePercent >= 0 ? 'positive' : 'negative');
    }
  },

  renderRecentTransactions(transactions) {
    const container = document.getElementById('recentTransactions');
    if (!container) return;
    const sorted = [...transactions].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 5);
    if (sorted.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <span class="empty-icon">📝</span>
          <p>Chưa có giao dịch nào</p>
          <p class="text-muted">Bấm nút + để thêm giao dịch đầu tiên</p>
        </div>`;
    } else {
      container.innerHTML = sorted.map(t => this.renderRecentItem(t)).join('');
    }
  },
  
  // ==================== TRANSACTIONS PAGE ====================
  renderTransactions() {
    let transactions = Storage.getLocal();
    
    const filterMonth = document.getElementById('filterMonth')?.value;
    const filterType = document.getElementById('filterType')?.value || 'all';
    const filterCategory = document.getElementById('filterCategory')?.value || 'all';
    const filterSort = document.getElementById('filterSort')?.value || 'date-desc';
    
    if (filterMonth) {
      const [fy, fm] = filterMonth.split('-').map(Number);
      transactions = transactions.filter(t => {
        const d = new Date(t.date);
        return d.getFullYear() === fy && (d.getMonth() + 1) === fm;
      });
    }
    if (filterType !== 'all') transactions = transactions.filter(t => t.type === filterType);
    if (filterCategory !== 'all') transactions = transactions.filter(t => t.category === filterCategory);
    
    switch (filterSort) {
      case 'date-desc': transactions.sort((a, b) => new Date(b.date) - new Date(a.date)); break;
      case 'date-asc': transactions.sort((a, b) => new Date(a.date) - new Date(b.date)); break;
      case 'amount-desc': transactions.sort((a, b) => b.amount - a.amount); break;
      case 'amount-asc': transactions.sort((a, b) => a.amount - b.amount); break;
    }
    
    const tbody = document.getElementById('transactionsBody');
    const emptyEl = document.getElementById('emptyTransactions');
    const tableEl = document.getElementById('transactionsTable');
    if (!tbody) return;
    
    if (transactions.length === 0) {
      tbody.innerHTML = '';
      if (tableEl) tableEl.style.display = 'none';
      if (emptyEl) emptyEl.style.display = 'flex';
    } else {
      if (tableEl) tableEl.style.display = '';
      if (emptyEl) emptyEl.style.display = 'none';
      tbody.innerHTML = transactions.map(t => this.renderTransactionRow(t)).join('');
      
      let totalIncome = 0, totalExpense = 0;
      transactions.forEach(t => {
        if (t.type === 'income') totalIncome += t.amount;
        else totalExpense += t.amount;
      });
      const net = totalIncome - totalExpense;
      const totalEl = document.getElementById('tableTotal');
      if (totalEl) {
        totalEl.innerHTML = `<strong class="${net >= 0 ? 'amount-col income' : 'amount-col expense'}">${formatCurrency(net)}</strong>`;
      }
    }
  },
  
  // ==================== ANALYTICS PAGE ====================
  renderAnalytics() {
    const { year, month } = this.analyticsMonth;
    const transactions = Storage.getByMonth(year, month);
    const prev = navigateMonth(year, month, -1);
    const prevTransactions = Storage.getByMonth(prev.year, prev.month);
    
    let income = 0, expense = 0;
    transactions.forEach(t => { if (t.type === 'income') income += t.amount; else expense += t.amount; });
    const savings = income - expense;
    const savingsRate = income > 0 ? ((savings / income) * 100) : 0;
    
    let prevIncome = 0, prevExpense = 0;
    prevTransactions.forEach(t => { if (t.type === 'income') prevIncome += t.amount; else prevExpense += t.amount; });
    const prevSavings = prevIncome - prevExpense;
    const prevSavingsRate = prevIncome > 0 ? ((prevSavings / prevIncome) * 100) : 0;
    
    document.getElementById('analyticsIncome').textContent = formatCurrency(income);
    document.getElementById('analyticsExpense').textContent = formatCurrency(expense);
    document.getElementById('analyticsSavings').textContent = formatCurrency(savings);
    document.getElementById('analyticsSavingsRate').textContent = savingsRate.toFixed(1) + '%';
    document.getElementById('analyticsCurrentMonth').textContent = formatMonthLabel(year, month);
    
    this.updateAnalyticsChange('analyticsIncomeChange', income, prevIncome, false);
    this.updateAnalyticsChange('analyticsExpenseChange', expense, prevExpense, true);
    this.updateAnalyticsChange('analyticsSavingsChange', savings, prevSavings, false);
    this.updateAnalyticsChange('analyticsSavingsRateChange', savingsRate, prevSavingsRate, false);
    
    this.renderTopCategories(transactions);
    const monthsData = Storage.getLastNMonths(6);
    Charts.renderCategoryTrendChart('categoryTrendChart', monthsData);
  },

  updateAnalyticsChange(containerId, current, previous, invertColor) {
    const container = document.getElementById(containerId);
    if (!container) return;
    const change = calcChange(current, previous);
    const sign = change >= 0 ? '+' : '';
    const badge = container.querySelector('.change-badge');
    if (badge) {
      badge.textContent = `${sign}${change.toFixed(1)}%`;
      badge.className = 'change-badge';
      if (invertColor) badge.classList.add(change <= 0 ? 'up' : 'down');
      else badge.classList.add(change >= 0 ? 'up' : 'down');
    }
  },

  renderTopCategories(transactions) {
    const container = document.getElementById('topCategories');
    if (!container) return;
    const expenses = transactions.filter(t => t.type === 'expense');
    if (expenses.length === 0) {
      container.innerHTML = `<div class="empty-state"><span class="empty-icon">📊</span><p>Chưa có dữ liệu</p></div>`;
      return;
    }
    const categoryTotals = {};
    expenses.forEach(t => { categoryTotals[t.category] = (categoryTotals[t.category] || 0) + t.amount; });
    const totalExpense = expenses.reduce((sum, t) => sum + t.amount, 0);
    const sorted = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]);
    container.innerHTML = sorted.map(([catId, amount]) => {
      const cat = getCategoryById(catId);
      const percent = ((amount / totalExpense) * 100).toFixed(1);
      return `
        <div class="top-category-item">
          <div class="top-cat-info"><span class="top-cat-emoji">${cat ? cat.emoji : '📦'}</span><span class="top-cat-label">${cat ? cat.label : 'Khác'}</span></div>
          <div class="top-cat-bar-wrapper"><div class="top-cat-bar" style="width: ${percent}%"></div></div>
          <div class="top-cat-values"><span class="top-cat-amount">${formatCurrency(amount)}</span><span class="top-cat-percent">${percent}%</span></div>
        </div>`;
    }).join('');
  },

  // ==================== SETTINGS PAGE ====================
  renderSettings() {
    const syncToggle = document.getElementById('syncToggle');
    const apiConfig = document.getElementById('apiConfig');
    const apiUrl = document.getElementById('apiUrl');
    
    if (syncToggle) syncToggle.checked = Storage.getSyncMode() === 'sheets';
    if (apiConfig) apiConfig.style.display = Storage.getSyncMode() === 'sheets' ? 'block' : 'none';
    if (apiUrl) apiUrl.value = Storage.getApiUrl();
    
    // Show sync actions if connected
    if (Storage.isOnline()) {
      const syncActions = document.getElementById('syncActions');
      if (syncActions) syncActions.style.display = 'flex';
    }
  },

  async handleTestConnection() {
    const apiUrl = document.getElementById('apiUrl')?.value?.trim();
    if (!apiUrl) {
      this.showToast('Vui lòng nhập URL', 'error');
      return;
    }
    
    Storage.setApiUrl(apiUrl);
    
    const statusEl = document.getElementById('connectionStatus');
    const iconEl = document.getElementById('statusIcon');
    const textEl = document.getElementById('statusText');
    const actionsEl = document.getElementById('syncActions');
    
    if (statusEl) { statusEl.style.display = 'flex'; statusEl.className = 'connection-status loading'; }
    if (iconEl) iconEl.textContent = '⏳';
    if (textEl) textEl.textContent = 'Đang kiểm tra kết nối...';
    
    const result = await Storage.testConnection();
    
    if (result.success) {
      if (statusEl) statusEl.className = 'connection-status success';
      if (iconEl) iconEl.textContent = '✅';
      if (textEl) textEl.textContent = result.message;
      if (actionsEl) actionsEl.style.display = 'flex';
      Storage.setSyncMode('sheets');
      this.updateSyncStatusBar();
      this.showToast('Kết nối thành công! ✅');
    } else {
      if (statusEl) statusEl.className = 'connection-status error';
      if (iconEl) iconEl.textContent = '❌';
      if (textEl) textEl.textContent = result.message;
      if (actionsEl) actionsEl.style.display = 'none';
      this.showToast('Lỗi kết nối: ' + result.message, 'error');
    }
  },

  handleSyncToggle() {
    const checked = document.getElementById('syncToggle')?.checked;
    const apiConfig = document.getElementById('apiConfig');
    
    if (checked) {
      Storage.setSyncMode('sheets');
      if (apiConfig) apiConfig.style.display = 'block';
    } else {
      Storage.setSyncMode('local');
      if (apiConfig) apiConfig.style.display = 'none';
      const syncActions = document.getElementById('syncActions');
      if (syncActions) syncActions.style.display = 'none';
    }
    this.updateSyncStatusBar();
  },

  async handleSyncPull() {
    this.showToast('Đang tải dữ liệu từ Sheet...');
    await this.syncFromSheets(false);
  },

  async handleSyncPush() {
    if (!confirm('Đẩy tất cả dữ liệu local lên Sheet? Dữ liệu cũ trên Sheet sẽ bị ghi đè.')) return;
    this.showToast('Đang đẩy dữ liệu lên Sheet...');
    try {
      await Storage.uploadAllToSheets();
      this.showToast('Đã đẩy dữ liệu lên Sheet thành công! ✅');
    } catch (e) {
      this.showToast('Lỗi: ' + e.message, 'error');
    }
  },

  handleCopyScript() {
    const code = document.getElementById('appsScriptCode')?.value || APPS_SCRIPT_CODE;
    navigator.clipboard.writeText(code).then(() => {
      this.showToast('Đã copy code Apps Script! 📋');
    }).catch(() => {
      // Fallback
      const textarea = document.createElement('textarea');
      textarea.value = code;
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      this.showToast('Đã copy code Apps Script! 📋');
    });
  },

  handleExportJson() {
    const data = Storage.getLocal();
    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `cashflow_backup_${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    this.showToast('Đã xuất dữ liệu JSON 📤');
  },

  handleImportJson(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = JSON.parse(evt.target.result);
        if (!Array.isArray(data)) throw new Error('Dữ liệu không hợp lệ');
        Storage.saveLocal(data);
        this.showToast(`Đã nhập ${data.length} giao dịch! ✅`);
        this.renderCurrentPage();
        // Also sync to sheets if online
        if (Storage.isOnline()) {
          Storage.uploadAllToSheets().catch(err => console.warn('Sheet sync after import failed:', err));
        }
      } catch (err) {
        this.showToast('File không hợp lệ: ' + err.message, 'error');
      }
    };
    reader.readAsText(file);
    e.target.value = ''; // Reset input
  },

  handleClearData() {
    if (!confirm('Bạn có chắc muốn XÓA TOÀN BỘ dữ liệu? Hành động này không thể hoàn tác!')) return;
    if (!confirm('Xác nhận lần nữa: XÓA TẤT CẢ giao dịch?')) return;
    Storage.saveLocal([]);
    this.showToast('Đã xóa toàn bộ dữ liệu 🗑️');
    this.renderCurrentPage();
  },

  // ==================== TRANSACTION CRUD ====================
  openModal(editId = null) {
    this.editingId = editId;
    const modal = document.getElementById('transactionModal');
    if (!modal) return;
    const form = document.getElementById('transactionForm');
    if (form) form.reset();
    
    if (editId) {
      document.getElementById('modalTitle').textContent = 'Sửa giao dịch';
      document.getElementById('editId').value = editId;
      const t = Storage.getLocal().find(x => x.id === editId);
      if (t) {
        this.selectedType = t.type;
        this.selectedCategory = t.category;
        document.getElementById('amount').value = t.amount;
        document.getElementById('date').value = t.date;
        document.getElementById('note').value = t.note || '';
        document.querySelectorAll('.type-btn').forEach(btn => {
          btn.classList.remove('active');
          if (btn.dataset.type === t.type) btn.classList.add('active');
        });
      }
    } else {
      document.getElementById('modalTitle').textContent = 'Thêm giao dịch';
      document.getElementById('editId').value = '';
      this.selectedType = 'expense';
      this.selectedCategory = null;
      document.getElementById('date').value = new Date().toISOString().split('T')[0];
      document.querySelectorAll('.type-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.dataset.type === 'expense') btn.classList.add('active');
      });
    }
    this.renderCategories();
    modal.classList.add('active');
    setTimeout(() => document.getElementById('amount')?.focus(), 100);
  },
  
  closeModal() {
    const modal = document.getElementById('transactionModal');
    if (modal) modal.classList.remove('active');
    this.editingId = null;
  },
  
  handleSubmit(e) {
    e.preventDefault();
    const amount = parseFloat(document.getElementById('amount').value);
    const date = document.getElementById('date').value;
    const note = document.getElementById('note').value;
    
    if (!amount || amount <= 0) { this.showToast('Vui lòng nhập số tiền hợp lệ', 'error'); return; }
    if (!this.selectedCategory) { this.showToast('Vui lòng chọn danh mục', 'error'); return; }
    if (!date) { this.showToast('Vui lòng chọn ngày', 'error'); return; }
    
    const data = { type: this.selectedType, amount, category: this.selectedCategory, date, note };
    
    if (this.editingId) {
      Storage.update(this.editingId, data);
      this.showToast('Đã cập nhật giao dịch ✅');
    } else {
      Storage.add(data);
      this.showToast('Đã thêm giao dịch mới ✅');
    }
    this.closeModal();
    this.renderCurrentPage();
  },
  
  deleteTransaction(id) {
    if (confirm('Bạn có chắc chắn muốn xóa giao dịch này?')) {
      Storage.delete(id);
      this.showToast('Đã xóa giao dịch 🗑️');
      this.renderCurrentPage();
    }
  },
  
  // Category selection in form
  selectType(type) {
    this.selectedType = type;
    this.selectedCategory = null;
    document.querySelectorAll('.type-btn').forEach(btn => {
      btn.classList.remove('active');
      if (btn.dataset.type === type) btn.classList.add('active');
    });
    this.renderCategories();
  },
  
  selectCategory(categoryId) {
    this.selectedCategory = categoryId;
    document.querySelectorAll('.category-btn').forEach(el => {
      el.classList.toggle('active', el.dataset.id === categoryId);
    });
  },
  
  renderCategories() {
    const container = document.getElementById('categoryGrid');
    if (!container) return;
    const cats = CATEGORIES[this.selectedType];
    container.innerHTML = cats.map(cat => `
      <button type="button" class="category-btn ${this.selectedCategory === cat.id ? 'active' : ''}" data-id="${cat.id}">
        <span class="cat-icon">${cat.emoji}</span>
        <span>${cat.label}</span>
      </button>
    `).join('');
  },
  
  // ==================== EXPORT CSV ====================
  exportCSV() {
    const csv = Storage.exportCSV();
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `cashflow_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    this.showToast('Đã xuất dữ liệu CSV 📥');
  },
  
  // ==================== TOAST ====================
  showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `<span>${message}</span>`;
    container.appendChild(toast);
    setTimeout(() => {
      toast.classList.add('hide');
      setTimeout(() => toast.remove(), 300);
    }, 3000);
  },
  
  // ==================== RENDER HELPERS ====================
  renderTransactionRow(t) {
    const cat = getCategoryById(t.category);
    const amountClass = t.type === 'income' ? 'income' : 'expense';
    const amountPrefix = t.type === 'income' ? '+' : '-';
    const typeLabel = t.type === 'income' ? 'Thu' : 'Chi';
    return `
      <tr>
        <td>${formatDate(t.date)}</td>
        <td><span class="badge">${typeLabel}</span></td>
        <td>${cat ? cat.emoji + ' ' + cat.label : '📦 Khác'}</td>
        <td>${t.note || '—'}</td>
        <td class="amount-col ${amountClass}">${amountPrefix}${formatCurrency(t.amount)}</td>
        <td>
          <div class="action-btns">
            <button class="edit-btn" data-id="${t.id}" title="Sửa">✏️</button>
            <button class="delete-btn" data-id="${t.id}" title="Xóa">🗑️</button>
          </div>
        </td>
      </tr>`;
  },
  
  renderRecentItem(t) {
    const cat = getCategoryById(t.category);
    const isIncome = t.type === 'income';
    return `
      <div class="recent-item">
        <div class="recent-icon">${cat ? cat.emoji : '📦'}</div>
        <div class="recent-info">
          <div class="recent-title">${cat ? cat.label : 'Khác'}${t.note ? ' · ' + t.note : ''}</div>
          <div class="recent-date">${formatDate(t.date)}</div>
        </div>
        <div class="recent-amount ${isIncome ? 'income' : 'expense'}">
          ${isIncome ? '+' : '-'}${formatCurrency(t.amount)}
        </div>
      </div>`;
  },
  
  // ==================== EVENT BINDINGS ====================
  bindEvents() {
    // Nav clicks
    document.addEventListener('click', (e) => {
      const navItem = e.target.closest('[data-page]');
      if (navItem) { e.preventDefault(); this.navigate(navItem.dataset.page); }
    });
    
    // Month navigation (Dashboard)
    document.getElementById('prevMonth')?.addEventListener('click', () => {
      this.currentMonth = navigateMonth(this.currentMonth.year, this.currentMonth.month, -1);
      this.renderDashboard();
    });
    document.getElementById('nextMonth')?.addEventListener('click', () => {
      this.currentMonth = navigateMonth(this.currentMonth.year, this.currentMonth.month, 1);
      this.renderDashboard();
    });
    
    // Month navigation (Analytics)
    document.getElementById('analyticsPrevMonth')?.addEventListener('click', () => {
      this.analyticsMonth = navigateMonth(this.analyticsMonth.year, this.analyticsMonth.month, -1);
      this.renderAnalytics();
    });
    document.getElementById('analyticsNextMonth')?.addEventListener('click', () => {
      this.analyticsMonth = navigateMonth(this.analyticsMonth.year, this.analyticsMonth.month, 1);
      this.renderAnalytics();
    });
    
    // FAB and add button
    document.getElementById('addTransactionBtn')?.addEventListener('click', (e) => {
      e.stopPropagation(); this.openModal();
    });
    document.getElementById('fabAdd')?.addEventListener('click', () => this.openModal());
    
    // Modal
    document.getElementById('modalClose')?.addEventListener('click', () => this.closeModal());
    document.getElementById('transactionModal')?.addEventListener('click', (e) => {
      if (e.target.id === 'transactionModal') this.closeModal();
    });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') this.closeModal(); });
    
    // Form
    document.getElementById('transactionForm')?.addEventListener('submit', (e) => this.handleSubmit(e));
    
    // Type toggle
    document.querySelectorAll('.type-btn').forEach(btn => {
      btn.addEventListener('click', (e) => { e.preventDefault(); this.selectType(e.currentTarget.dataset.type); });
    });
    
    // Category grid
    document.getElementById('categoryGrid')?.addEventListener('click', (e) => {
      const btn = e.target.closest('.category-btn');
      if (btn) this.selectCategory(btn.dataset.id);
    });
    
    // Filters
    ['filterMonth', 'filterType', 'filterCategory', 'filterSort'].forEach(id => {
      document.getElementById(id)?.addEventListener('change', () => this.renderTransactions());
    });
    
    // Theme toggle (both)
    document.getElementById('themeToggle')?.addEventListener('click', () => this.toggleTheme());
    document.getElementById('themeToggleMobile')?.addEventListener('click', () => this.toggleTheme());
    
    // Export CSV (both)
    document.getElementById('exportBtn')?.addEventListener('click', () => this.exportCSV());
    document.getElementById('exportBtnMobile')?.addEventListener('click', () => this.exportCSV());
    
    // Edit/Delete in table
    document.getElementById('transactionsBody')?.addEventListener('click', (e) => {
      const editBtn = e.target.closest('.edit-btn');
      const delBtn = e.target.closest('.delete-btn');
      if (editBtn) this.openModal(editBtn.dataset.id);
      if (delBtn) this.deleteTransaction(delBtn.dataset.id);
    });

    // ==================== SETTINGS EVENTS ====================
    document.getElementById('syncToggle')?.addEventListener('change', () => this.handleSyncToggle());
    document.getElementById('testConnectionBtn')?.addEventListener('click', () => this.handleTestConnection());
    document.getElementById('syncPullBtn')?.addEventListener('click', () => this.handleSyncPull());
    document.getElementById('syncPushBtn')?.addEventListener('click', () => this.handleSyncPush());
    document.getElementById('copyScriptBtn')?.addEventListener('click', () => this.handleCopyScript());
    document.getElementById('exportJsonBtn')?.addEventListener('click', () => this.handleExportJson());
    document.getElementById('importJsonInput')?.addEventListener('change', (e) => this.handleImportJson(e));
    document.getElementById('clearDataBtn')?.addEventListener('click', () => this.handleClearData());
  }
};

// Start app
document.addEventListener('DOMContentLoaded', () => App.init());
