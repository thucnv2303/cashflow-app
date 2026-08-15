const STORAGE_KEY = 'cashflow_transactions';
const THEME_KEY = 'cashflow_theme';
const API_URL_KEY = 'cashflow_api_url';
const SYNC_MODE_KEY = 'cashflow_sync_mode';

const MEMBERS_KEY = 'cashflow_members';
const FAMILY_NAME_KEY = 'cashflow_family_name';
const SETUP_DONE_KEY = 'cashflow_setup_done';
const LOANS_KEY = 'cashflow_loans';

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
  getTheme() { return localStorage.getItem(THEME_KEY) || 'dark'; },
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
      this._sheetApiCall('addMember', { data: newMember }).catch(e => console.warn(e));
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
        this._sheetApiCall('updateMember', { id, data: members[idx] }).catch(e => console.warn(e));
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
      this._sheetApiCall('deleteMember', { id }).catch(e => console.warn(e));
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
    let fetchUrl = `${url}?action=${action}`;
    for (const key in params) {
      fetchUrl += `&${key}=${encodeURIComponent(typeof params[key] === 'object' ? JSON.stringify(params[key]) : params[key])}`;
    }
    const response = await fetch(fetchUrl);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const result = await response.json();
    return result;
  },

  async testConnection() {
    try {
      const res = await this._sheetApiCall('ping');
      if (res && res.status === 'ok') {
        return { success: true, message: res.message || 'Kết nối thành công!' };
      }
      return { success: false, message: 'Phản hồi không hợp lệ từ máy chủ' };
    } catch (error) {
      console.error('Test connection error:', error);
      return { success: false, message: 'Lỗi kết nối: ' + error.message };
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
      this.saveLocal(data);
      
      try {
        const membersData = await this.fetchMembersFromSheets();
        this.saveMembers(membersData);
      } catch (e) { console.warn('Sync members failed:', e); }

      try {
        const loansData = await this.fetchLoansFromSheets();
        this.saveLoans(loansData);
      } catch (e) { console.warn('Sync loans failed:', e); }

      return true;
    } catch (e) {
      console.warn('Sync failed:', e);
      return false;
    }
  }
};
