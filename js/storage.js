const STORAGE_KEY = 'cashflow_transactions';
const THEME_KEY = 'cashflow_theme';
const API_URL_KEY = 'cashflow_api_url';
const SYNC_MODE_KEY = 'cashflow_sync_mode';

const MEMBERS_KEY = 'cashflow_members';
const FAMILY_NAME_KEY = 'cashflow_family_name';
const SETUP_DONE_KEY = 'cashflow_setup_done';
const LOANS_KEY = 'cashflow_loans';
const REMINDER_KEY = 'cashflow_reminders';
const BUDGETS_KEY = 'cashflow_budgets';
const CUSTOM_CATS_KEY = 'cashflow_custom_categories';
const SAVINGS_GOAL_KEY = 'cashflow_savings_goal';

const Storage = {
  // ==================== CONFIG ====================
  getApiUrl() { return localStorage.getItem(API_URL_KEY) || ''; },
  setApiUrl(url) { localStorage.setItem(API_URL_KEY, url); },
  getSyncMode() { return localStorage.getItem(SYNC_MODE_KEY) || 'local'; },
  setSyncMode(mode) { localStorage.setItem(SYNC_MODE_KEY, mode); },
  isOnline() { return this.getSyncMode() === 'sheets' && this.getApiUrl() !== ''; },

  // ==================== SETUP FLAGS ====================
  isSetupDone() { return localStorage.getItem(SETUP_DONE_KEY) === 'true'; },
  setSetupDone() { localStorage.setItem(SETUP_DONE_KEY, 'true'); },
  getFamilyName() { return localStorage.getItem(FAMILY_NAME_KEY) || 'Gia đình'; },
  setFamilyName(name) { localStorage.setItem(FAMILY_NAME_KEY, name); },

  // ==================== THEME ====================
  getTheme() { return localStorage.getItem(THEME_KEY) || 'light'; },
  setTheme(theme) { localStorage.setItem(THEME_KEY, theme); },

  // ==================== MEMBERS CRUD ====================
  getMembers() { return JSON.parse(localStorage.getItem(MEMBERS_KEY) || '[]'); },
  saveMembers(members) { localStorage.setItem(MEMBERS_KEY, JSON.stringify(members)); },
  addMember(member) {
    const members = this.getMembers();
    const newMember = { ...member, id: generateId() };
    members.push(newMember);
    this.saveMembers(members);
    if (this.isOnline()) {
      this.syncMembersToSheets().catch(e => console.warn(e));
    }
    return newMember;
  },
  updateMember(id, data) {
    const members = this.getMembers();
    const idx = members.findIndex(m => m.id === id);
    if (idx !== -1) {
      members[idx] = { ...members[idx], ...data };
      this.saveMembers(members);
      if (this.isOnline()) {
        this.syncMembersToSheets().catch(e => console.warn(e));
      }
      return members[idx];
    }
    return null;
  },
  deleteMember(id) {
    let members = this.getMembers();
    members = members.filter(m => m.id !== id);
    this.saveMembers(members);
    if (this.isOnline()) {
      this.syncMembersToSheets().catch(e => console.warn(e));
    }
  },
  getMemberById(id) { return this.getMembers().find(m => m.id === id); },

  // ==================== LOANS CRUD ====================
  getLoans() { return JSON.parse(localStorage.getItem(LOANS_KEY) || '[]'); },
  saveLoans(loans) { localStorage.setItem(LOANS_KEY, JSON.stringify(loans)); },
  addLoan(loan) {
    const loans = this.getLoans();
    const newLoan = { ...loan, id: generateId(), createdAt: new Date().toISOString() };
    loans.push(newLoan);
    this.saveLoans(loans);
    if (this.isOnline()) {
      this._sheetApiCall('addLoan', { data: newLoan }).catch(e => console.warn(e));
    }
    return newLoan;
  },
  updateLoan(id, data) {
    const loans = this.getLoans();
    const idx = loans.findIndex(l => l.id === id);
    if (idx !== -1) {
      loans[idx] = { ...loans[idx], ...data };
      this.saveLoans(loans);
      if (this.isOnline()) {
        this._sheetApiCall('updateLoan', { id, data: loans[idx] }).catch(e => console.warn(e));
      }
      return loans[idx];
    }
    return null;
  },
  deleteLoan(id) {
    let loans = this.getLoans();
    loans = loans.filter(l => l.id !== id);
    this.saveLoans(loans);
    if (this.isOnline()) {
      this._sheetApiCall('deleteLoan', { id }).catch(e => console.warn(e));
    }
  },

  // ==================== BUDGETS CRUD ====================
  getBudgets() {
    const raw = localStorage.getItem(BUDGETS_KEY);
    if (!raw) return { ...DEFAULT_BUDGETS };
    try {
      return { ...DEFAULT_BUDGETS, ...JSON.parse(raw) };
    } catch(e) {
      return { ...DEFAULT_BUDGETS };
    }
  },
  saveBudgets(budgets) {
    localStorage.setItem(BUDGETS_KEY, JSON.stringify(budgets));
    if (this.isOnline()) {
      this.syncBudgetsToSheets().catch(e => console.warn(e));
    }
  },
  getCategoryBudget(catId) {
    const budgets = this.getBudgets();
    return budgets[catId] !== undefined ? budgets[catId] : (DEFAULT_BUDGETS[catId] || 0);
  },
  async fetchBudgetsFromSheets() {
    const res = await this._sheetApiCall('getBudgets');
    if (res && res.success) return res.data;
    throw new Error(res.error || 'Không thể lấy ngân sách từ Sheets');
  },
  async syncBudgetsToSheets() {
    if (!this.isOnline()) return false;
    const budgets = this.getBudgets();
    return this._sheetApiCall('syncBudgets', { data: budgets });
  },

  // ==================== SAVINGS GOAL ====================
  getSavingsGoal() {
    const raw = localStorage.getItem(SAVINGS_GOAL_KEY);
    return raw ? Number(raw) : 5000000;
  },
  saveSavingsGoal(amount) {
    localStorage.setItem(SAVINGS_GOAL_KEY, String(amount || 0));
  },

  // ==================== CUSTOM CATEGORIES CRUD ====================
  getCustomCategories() {
    const raw = localStorage.getItem(CUSTOM_CATS_KEY);
    if (!raw) return [];
    try {
      return JSON.parse(raw);
    } catch(e) {
      return [];
    }
  },
  saveCustomCategories(cats) {
    localStorage.setItem(CUSTOM_CATS_KEY, JSON.stringify(cats));
    if (this.isOnline()) {
      this.syncCustomCatsToSheets().catch(e => console.warn(e));
    }
  },
  addCustomCategory(cat) {
    const cats = this.getCustomCategories();
    const id = 'custom_' + generateId();
    const newCat = { id, label: cat.label, emoji: cat.emoji || '📦', type: 'expense', createdAt: new Date().toISOString() };
    cats.push(newCat);
    this.saveCustomCategories(cats);
    return newCat;
  },
  deleteCustomCategory(id) {
    let cats = this.getCustomCategories();
    cats = cats.filter(c => c.id !== id);
    this.saveCustomCategories(cats);
    // Also delete its budget
    const budgets = this.getBudgets();
    delete budgets[id];
    this.saveBudgets(budgets);
  },
  getAllExpenseCategories() {
    return getExpenseCategories();
  },
  async fetchCustomCatsFromSheets() {
    const res = await this._sheetApiCall('getCustomCats');
    if (res && res.success) return res.data;
    throw new Error(res.error || 'Không thể lấy danh mục tùy chỉnh từ Sheets');
  },
  async syncCustomCatsToSheets() {
    if (!this.isOnline()) return false;
    const cats = this.getCustomCategories();
    return this._sheetApiCall('syncCustomCats', { data: cats });
  },

  // ==================== TRANSACTIONS LOCAL STORAGE ====================
  getLocal() {
    const data = localStorage.getItem(STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  },
  saveLocal(transactions) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(transactions));
  },

  // ==================== GOOGLE SHEETS API ====================
  async _sheetApiCall(action, params = {}) {
    const url = this.getApiUrl();
    if (!url) throw new Error("Chưa cấu hình API URL");
    
    // For actions sending data, try POST first with plain text body (avoids CORS preflight)
    const payload = { action, ...params };
    const payloadStr = JSON.stringify(payload);
    
    let fetchUrl = `${url}?action=${action}`;
    for (const key in params) {
      fetchUrl += `&${key}=${encodeURIComponent(typeof params[key] === 'object' ? JSON.stringify(params[key]) : params[key])}`;
    }

    try {
      if (['syncMembers', 'syncLoans', 'syncBudgets', 'syncCustomCats', 'sync', 'add', 'update'].includes(action)) {
        try {
          const controller = new AbortController();
          const to = setTimeout(() => controller.abort(), 12000);
          const res = await fetch(url, {
            method: 'POST',
            body: payloadStr,
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            signal: controller.signal
          });
          clearTimeout(to);
          const text = await res.text();
          try { 
            const parsed = JSON.parse(text);
            if (parsed && (parsed.success || parsed.status === 'ok')) return parsed;
          } catch(e) {}
        } catch(postErr) {
          console.warn('POST failed, fallback to GET/JSONP:', postErr);
        }
      }
      return await this._fetchWithTimeout(fetchUrl, 15000);
    } catch (fetchError) {
      console.warn('Fetch failed, trying JSONP fallback:', fetchError.message);
      try {
        return await this._jsonpCall(fetchUrl, 15000);
      } catch (jsonpError) {
        throw new Error('Không thể kết nối. Kiểm tra: (1) URL đúng format .../exec, (2) Deploy → Execute as: Me, Who has access: Anyone, (3) Đã chạy setupSheet() và cấp quyền');
      }
    }
  },

  async _fetchWithTimeout(fetchUrl, timeoutMs) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const response = await fetch(fetchUrl, {
        method: 'GET',
        redirect: 'follow',
        signal: controller.signal
      });
      clearTimeout(timeout);
      const text = await response.text();
      try {
        return JSON.parse(text);
      } catch {
        if (text.includes('<!DOCTYPE') || text.includes('<html')) {
          throw new Error('Apps Script trả về HTML thay vì JSON');
        }
        throw new Error('Phản hồi không phải JSON');
      }
    } catch (error) {
      clearTimeout(timeout);
      throw error;
    }
  },

  // JSONP fallback - bypass CORS bằng script injection
  _jsonpCall(url, timeoutMs) {
    return new Promise((resolve, reject) => {
      const callbackName = '_cashflow_cb_' + Date.now();
      const timeout = setTimeout(() => {
        delete window[callbackName];
        script.remove();
        reject(new Error('JSONP timeout'));
      }, timeoutMs);

      window[callbackName] = (data) => {
        clearTimeout(timeout);
        delete window[callbackName];
        script.remove();
        resolve(data);
      };

      const script = document.createElement('script');
      const separator = url.includes('?') ? '&' : '?';
      script.src = `${url}${separator}callback=${callbackName}`;
      script.onerror = () => {
        clearTimeout(timeout);
        delete window[callbackName];
        script.remove();
        reject(new Error('JSONP script load failed'));
      };
      document.head.appendChild(script);
    });
  },

  async testConnection() {
    const url = this.getApiUrl();
    if (!url) return { success: false, message: 'Chưa nhập URL' };
    if (!url.includes('/exec')) {
      return { success: false, message: 'URL phải kết thúc bằng /exec. Kiểm tra lại URL deploy.' };
    }
    
    try {
      // Thử fetch trực tiếp trước
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 12000);
      try {
        const response = await fetch(`${url}?action=ping`, {
          method: 'GET',
          redirect: 'follow',
          signal: controller.signal
        });
        clearTimeout(timeout);
        const text = await response.text();
        
        // Phát hiện Google Login page
        if (text.includes('accounts.google.com') || text.includes('ServiceLogin') || text.includes('identifierview')) {
          return { 
            success: false, 
            message: '⚠️ Apps Script yêu cầu đăng nhập! Khi Deploy phải chọn: Execute as → Me, Who has access → Anyone. Sau đó tạo New deployment mới.' 
          };
        }
        
        // Phát hiện HTML error page
        if (text.includes('<!DOCTYPE') || text.includes('<html')) {
          return { 
            success: false, 
            message: '⚠️ Apps Script trả về HTML thay vì JSON. Cần Deploy lại: Execute as → Me, Who has access → Anyone.' 
          };
        }
        
        try {
          const data = JSON.parse(text);
          if (data && data.status === 'ok') {
            return { success: true, message: data.message || 'Kết nối thành công! ✅' };
          }
          if (data && data.error) {
            return { success: false, message: 'Lỗi từ server: ' + data.error };
          }
          return { success: false, message: 'Phản hồi không hợp lệ' };
        } catch {
          return { success: false, message: 'Phản hồi không phải JSON hợp lệ. Kiểm tra code Apps Script.' };
        }
      } catch (fetchErr) {
        clearTimeout(timeout);
        console.warn('Fetch test failed:', fetchErr.message);
        
        // CORS error → thử JSONP
        if (fetchErr.message === 'Failed to fetch' || fetchErr.name === 'TypeError') {
          try {
            const data = await this._jsonpCall(`${url}?action=ping`, 12000);
            if (data && data.status === 'ok') {
              return { success: true, message: 'Kết nối thành công (JSONP)! ✅' };
            }
            return { success: false, message: 'JSONP response không hợp lệ' };
          } catch (jsonpErr) {
            return { 
              success: false, 
              message: '⚠️ Không kết nối được. Lỗi CORS + JSONP. Kiểm tra:\n1. Deploy → Execute as: Me\n2. Who has access: Anyone\n3. Đã chạy setupSheet() và cấp quyền\n4. Copy code MỚI NHẤT từ app vào Apps Script' 
            };
          }
        }
        
        if (fetchErr.name === 'AbortError') {
          return { success: false, message: 'Quá thời gian kết nối (12s). Kiểm tra lại URL.' };
        }
        return { success: false, message: 'Lỗi kết nối: ' + fetchErr.message };
      }
    } catch (error) {
      console.error('Test connection error:', error);
      return { success: false, message: 'Lỗi: ' + error.message };
    }
  },

  async fetchFromSheets() {
    const res = await this._sheetApiCall('getAll');
    if (res && res.success) return res.data;
    throw new Error(res.error || 'Không thể lấy dữ liệu từ Sheets');
  },

  async fetchMembersFromSheets() {
    const res = await this._sheetApiCall('getMembers');
    if (res && res.success) return res.data;
    throw new Error(res.error || 'Không thể lấy thành viên từ Sheets');
  },

  async fetchLoansFromSheets() {
    const res = await this._sheetApiCall('getLoans');
    if (res && res.success) return res.data;
    throw new Error(res.error || 'Không thể lấy khoản vay từ Sheets');
  },

  async addToSheets(transaction) {
    return this._sheetApiCall('add', { data: transaction });
  },
  async updateInSheets(id, data) {
    return this._sheetApiCall('update', { id, data });
  },
  async deleteFromSheets(id) {
    return this._sheetApiCall('delete', { id });
  },
  
  async uploadAllToSheets() {
    const localData = this.getLocal();
    return this._sheetApiCall('sync', { data: localData });
  },
  async syncMembersToSheets() {
    if (!this.isOnline()) return false;
    const members = this.getMembers();
    return this._sheetApiCall('syncMembers', { data: members });
  },
  async syncLoansToSheets() {
    if (!this.isOnline()) return false;
    const loans = this.getLoans();
    return this._sheetApiCall('syncLoans', { data: loans });
  },

  // ==================== UNIFIED TRANSACTIONS API ====================
  async getAll() {
    if (this.isOnline()) {
      try {
        const sheetsData = await this.fetchFromSheets();
        this.saveLocal(sheetsData);
        return sheetsData;
      } catch (e) {
        console.warn('Sheets unavailable, using local cache:', e);
        return this.getLocal();
      }
    }
    return this.getLocal();
  },
  
  async add(transaction) {
    const newT = { 
      ...transaction, 
      id: generateId(), 
      createdAt: new Date().toISOString() 
    };
    
    const all = this.getLocal();
    all.push(newT);
    this.saveLocal(all);
    
    if (this.isOnline()) {
      this.addToSheets(newT).catch(e => console.warn('Sheet sync failed (add):', e));
    }
    return newT;
  },
  
  async update(id, data) {
    const all = this.getLocal();
    const idx = all.findIndex(t => t.id === id);
    if (idx !== -1) {
      all[idx] = { ...all[idx], ...data };
      this.saveLocal(all);
      
      if (this.isOnline()) {
        this.updateInSheets(id, all[idx]).catch(e => console.warn('Sheet sync failed (update):', e));
      }
      return all[idx];
    }
    return null;
  },
  
  async delete(id) {
    let all = this.getLocal();
    all = all.filter(t => t.id !== id);
    this.saveLocal(all);
    
    if (this.isOnline()) {
      this.deleteFromSheets(id).catch(e => console.warn('Sheet sync failed (delete):', e));
    }
  },

  getByMonth(year, month) {
    const transactions = this.getLocal();
    return transactions.filter(t => {
      const date = new Date(t.date);
      return date.getFullYear() === year && (date.getMonth() + 1) === month;
    });
  },
  
  getLastNMonths(n) {
    const { year: currentYear, month: currentMonth } = getCurrentMonth();
    const result = [];
    for (let i = n - 1; i >= 0; i--) {
      let m = currentMonth - i;
      let y = currentYear;
      while (m <= 0) {
        m += 12;
        y -= 1;
      }
      result.push({
        year: y,
        month: m,
        transactions: this.getByMonth(y, m)
      });
    }
    return result;
  },

  // ==================== UTILS ====================
  exportCSV() {
    const transactions = this.getLocal();
    let csv = "Ngày,Loại,Danh mục,Thành viên,Ghi chú,Số tiền\n";
    
    transactions.sort((a, b) => new Date(b.date) - new Date(a.date)).forEach(t => {
      const type = t.type === 'income' ? 'Thu nhập' : 'Chi tiêu';
      const categoryLabel = (typeof getCategoryById === 'function') 
        ? (getCategoryById(t.category)?.label || t.category)
        : t.category;
      
      let memberName = '';
      if (t.memberId) {
        const member = this.getMemberById(t.memberId);
        if (member) memberName = member.label || member.id;
      }
      
      const amount = t.amount;
      const note = (t.note || '').replace(/"/g, '""');
      csv += `${t.date},"${type}","${categoryLabel}","${memberName}","${note}",${amount}\n`;
    });
    
    return "\ufeff" + csv;
  },
  
  async syncFromSheets() {
    if (!this.isOnline()) return false;
    try {
      const data = await this.fetchFromSheets();
      if (Array.isArray(data)) this.saveLocal(data);
      
      try {
        const membersData = await this.fetchMembersFromSheets();
        if (Array.isArray(membersData) && membersData.length > 0) {
          const localMembers = this.getMembers();
          // Smart merge: never wipe out a member's valid avatarImg if the incoming sync is blank
          const merged = membersData.map(inc => {
            const local = localMembers.find(l => l.id === inc.id);
            if (local && local.avatarImg && (!inc.avatarImg || inc.avatarImg === '')) {
              return { ...inc, avatarImg: local.avatarImg, avatarId: local.avatarId };
            }
            return inc;
          });
          this.saveMembers(merged);
        }
      } catch (e) { console.warn('Sync members failed:', e); }

      try {
        const loansData = await this.fetchLoansFromSheets();
        if (Array.isArray(loansData)) this.saveLoans(loansData);
      } catch (e) { console.warn('Sync loans failed:', e); }

      try {
        const budgetsData = await this.fetchBudgetsFromSheets();
        if (budgetsData && typeof budgetsData === 'object') {
          const current = this.getBudgets();
          localStorage.setItem(BUDGETS_KEY, JSON.stringify({ ...current, ...budgetsData }));
        }
      } catch (e) { console.warn('Sync budgets failed:', e); }

      try {
        const catsData = await this.fetchCustomCatsFromSheets();
        if (Array.isArray(catsData)) {
          this.saveCustomCategories(catsData);
        }
      } catch (e) { console.warn('Sync custom cats failed:', e); }

      return true;
    } catch (e) {
      console.warn('Sync failed:', e);
      return false;
    }
  }
};
