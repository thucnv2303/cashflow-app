const STORAGE_KEY = 'cashflow_transactions';
const THEME_KEY = 'cashflow_theme';

const Storage = {
  // Get all transactions
  getAll() {
    const data = localStorage.getItem(STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  },
  
  // Save all transactions
  saveAll(transactions) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(transactions));
  },
  
  // Add a transaction
  add(transaction) {
    const transactions = this.getAll();
    const newTransaction = {
      ...transaction,
      id: generateId(),
      createdAt: new Date().toISOString()
    };
    transactions.push(newTransaction);
    this.saveAll(transactions);
    return newTransaction;
  },
  
  // Update a transaction by id
  update(id, data) {
    const transactions = this.getAll();
    const index = transactions.findIndex(t => t.id === id);
    if (index !== -1) {
      transactions[index] = { ...transactions[index], ...data };
      this.saveAll(transactions);
      return transactions[index];
    }
    return null;
  },
  
  // Delete a transaction by id
  delete(id) {
    let transactions = this.getAll();
    transactions = transactions.filter(t => t.id !== id);
    this.saveAll(transactions);
  },
  
  // Get transactions for a specific month
  getByMonth(year, month) {
    const transactions = this.getAll();
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
  
  // Get theme preference
  getTheme() {
    return localStorage.getItem(THEME_KEY) || 'dark';
  },
  
  // Save theme preference  
  setTheme(theme) {
    localStorage.setItem(THEME_KEY, theme);
  },
  
  // Export all as CSV string
  exportCSV() {
    const transactions = this.getAll();
    let csv = "Ngày,Loại,Danh mục,Ghi chú,Số tiền\n";
    
    transactions.sort((a, b) => new Date(b.date) - new Date(a.date)).forEach(t => {
      const type = t.type === 'income' ? 'Thu nhập' : 'Chi tiêu';
      const category = getCategoryById(t.category)?.label || t.category;
      const amount = t.amount;
      const note = (t.note || '').replace(/"/g, '""');
      csv += `${formatDate(t.date)},"${type}","${category}","${note}",${amount}\n`;
    });
    
    return "\ufeff" + csv; // Add BOM for Excel UTF-8 support
  }
};
