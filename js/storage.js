const STORAGE_KEY = 'cashflow_transactions';
const THEME_KEY = 'cashflow_theme';
const API_URL_KEY = 'cashflow_api_url';
const SYNC_MODE_KEY = 'cashflow_sync_mode'; // 'local' or 'sheets'

const Storage = {
  // ==================== CONFIG ====================
  
  // Get/Set API URL
  getApiUrl() { 
    return localStorage.getItem(API_URL_KEY) || ''; 
  },
  
  setApiUrl(url) { 
    localStorage.setItem(API_URL_KEY, url); 
  },
  
  // Get/Set sync mode ('local' or 'sheets')
  getSyncMode() { 
    return localStorage.getItem(SYNC_MODE_KEY) || 'local'; 
  },
  
  setSyncMode(mode) { 
    localStorage.setItem(SYNC_MODE_KEY, mode); 
  },
  
  // Check if using Google Sheets sync
  isOnline() { 
    return this.getSyncMode() === 'sheets' && this.getApiUrl() !== ''; 
  },
  
  // ==================== LOCAL STORAGE ====================
  
  // Get all transactions from localStorage
  getLocal() {
    const data = localStorage.getItem(STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  },
  
  // Save all transactions to localStorage
  saveLocal(transactions) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(transactions));
  },
  
  // ==================== GOOGLE SHEETS API ====================
  
  // Helper to make GET requests to Google Apps Script
  async _sheetApiCall(action, params = {}) {
    const url = this.getApiUrl();
    if (!url) throw new Error("Chưa cấu hình API URL");
    
    let fetchUrl = `${url}?action=${action}`;
    
    // Thêm các parameters vào URL
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

  // Test connection to Google Sheets API
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
  
  // Fetch all transactions from Google Sheets
  async fetchFromSheets() {
    const res = await this._sheetApiCall('getAll');
    if (res && res.success) {
      return res.data;
    }
    throw new Error(res.error || 'Không thể lấy dữ liệu từ Sheets');
  },
  
  // Add transaction to Google Sheets
  async addToSheets(transaction) {
    return this._sheetApiCall('add', { data: transaction });
  },
  
  // Update transaction in Google Sheets
  async updateInSheets(id, data) {
    return this._sheetApiCall('update', { id, data });
  },
  
  // Delete from Google Sheets
  async deleteFromSheets(id) {
    return this._sheetApiCall('delete', { id });
  },
  
  // Upload all local data to Google Sheets (initial sync)
  async uploadAllToSheets() {
    const localData = this.getLocal();
    return this._sheetApiCall('sync', { data: localData });
  },
  
  // ==================== UNIFIED API ====================
  // These are the main methods used by the app
  
  async getAll() {
    if (this.isOnline()) {
      try {
        const sheetsData = await this.fetchFromSheets();
        // Cập nhật lại cache local
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
    
    // Always save locally
    const all = this.getLocal();
    all.push(newT);
    this.saveLocal(all);
    
    // If online, also sync to sheets (fire and forget)
    if (this.isOnline()) {
      this.addToSheets(newT).catch(e => console.warn('Sheet sync failed (add):', e));
    }
    
    return newT;
  },
  
  async update(id, data) {
    // Update locally first
    const all = this.getLocal();
    const idx = all.findIndex(t => t.id === id);
    if (idx !== -1) {
      all[idx] = { ...all[idx], ...data };
      this.saveLocal(all);
      
      // If online, also sync to sheets
      if (this.isOnline()) {
        this.updateInSheets(id, all[idx]).catch(e => console.warn('Sheet sync failed (update):', e));
      }
      return all[idx];
    }
    return null;
  },
  
  async delete(id) {
    // Delete locally first
    let all = this.getLocal();
    all = all.filter(t => t.id !== id);
    this.saveLocal(all);
    
    // If online, also sync to sheets
    if (this.isOnline()) {
      this.deleteFromSheets(id).catch(e => console.warn('Sheet sync failed (delete):', e));
    }
  },
  
  // Get by month (always from local cache for speed)
  getByMonth(year, month) {
    const transactions = this.getLocal();
    return transactions.filter(t => {
      const date = new Date(t.date);
      return date.getFullYear() === year && (date.getMonth() + 1) === month;
    });
  },
  
  // Get transactions for last N months (for charts)
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
  
  // Theme
  getTheme() { 
    return localStorage.getItem(THEME_KEY) || 'dark'; 
  },
  
  setTheme(theme) { 
    localStorage.setItem(THEME_KEY, theme); 
  },
  
  // Export CSV
  exportCSV() {
    const transactions = this.getLocal();
    let csv = "Ngày,Loại,Danh mục,Ghi chú,Số tiền\n";
    
    transactions.sort((a, b) => new Date(b.date) - new Date(a.date)).forEach(t => {
      const type = t.type === 'income' ? 'Thu nhập' : 'Chi tiêu';
      // Fallback if getCategoryById is not defined yet
      const categoryLabel = (typeof getCategoryById === 'function') 
        ? (getCategoryById(t.category)?.label || t.category)
        : t.category;
      const amount = t.amount;
      const note = (t.note || '').replace(/"/g, '""');
      csv += `${t.date},"${type}","${categoryLabel}","${note}",${amount}\n`;
    });
    
    return "\ufeff" + csv; // Add BOM for Excel UTF-8 support
  },
  
  // Sync: pull from sheets and update local cache
  async syncFromSheets() {
    if (!this.isOnline()) return false;
    try {
      const data = await this.fetchFromSheets();
      this.saveLocal(data);
      return true;
    } catch (e) {
      console.warn('Sync failed:', e);
      return false;
    }
  }
};
