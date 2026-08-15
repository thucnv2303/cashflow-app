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
    this.renderCurrentPage();
  },
  
  // Theme management
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
  
  // Navigation
  navigate(page) {
    this.currentPage = page;
    
    // Update nav active state
    document.querySelectorAll('.nav-item').forEach(el => {
      el.classList.toggle('active', el.dataset.page === page);
    });
    
    // Hide all pages, show selected
    document.querySelectorAll('.page').forEach(el => {
      el.classList.remove('active');
    });
    const pageEl = document.getElementById(`page-${page}`);
    if (pageEl) {
      pageEl.classList.add('active');
    }
    
    this.renderCurrentPage();
  },
  
  renderCurrentPage() {
    if (this.currentPage === 'dashboard') this.renderDashboard();
    else if (this.currentPage === 'transactions') this.renderTransactions();
    else if (this.currentPage === 'analytics') this.renderAnalytics();
  },
  
  // Initialize filter month to current
  initFilterMonth() {
    const filterMonth = document.getElementById('filterMonth');
    if (filterMonth) {
      const now = new Date();
      const y = now.getFullYear();
      const m = (now.getMonth() + 1).toString().padStart(2, '0');
      filterMonth.value = `${y}-${m}`;
    }
    // Populate category filter
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

  // Render Dashboard
  renderDashboard() {
    const { year, month } = this.currentMonth;
    const transactions = Storage.getByMonth(year, month);
    
    // Calculate totals
    let income = 0, expense = 0;
    transactions.forEach(t => {
      if (t.type === 'income') income += t.amount;
      else expense += t.amount;
    });
    const balance = income - expense;
    const savingsRate = income > 0 ? ((income - expense) / income * 100) : 0;
    
    // Previous month for comparison
    const prev = navigateMonth(year, month, -1);
    const prevTransactions = Storage.getByMonth(prev.year, prev.month);
    let prevIncome = 0, prevExpense = 0;
    prevTransactions.forEach(t => {
      if (t.type === 'income') prevIncome += t.amount;
      else prevExpense += t.amount;
    });
    
    // Update summary cards
    document.getElementById('totalIncome').textContent = formatCurrency(income);
    document.getElementById('totalExpense').textContent = formatCurrency(expense);
    document.getElementById('balance').textContent = formatCurrency(balance);
    document.getElementById('savingsRate').textContent = savingsRate.toFixed(1) + '%';
    
    // Update change badges
    const incomeChange = calcChange(income, prevIncome);
    const expenseChange = calcChange(expense, prevExpense);
    this.updateChangeBadge('incomeChange', incomeChange, false);
    this.updateChangeBadge('expenseChange', expenseChange, true);
    
    // Update month label
    document.getElementById('currentMonth').textContent = formatMonthLabel(year, month);
    
    // Render charts
    Charts.renderCategoryChart('categoryChart', transactions);
    const monthsData = Storage.getLastNMonths(6);
    Charts.renderMonthlyChart('monthlyChart', monthsData);
    Charts.renderTrendChart('trendChart', monthsData);
    
    // Render recent transactions
    this.renderRecentTransactions(transactions);
  },

  updateChangeBadge(elementId, changePercent, invertColor) {
    const el = document.getElementById(elementId);
    if (!el) return;
    const sign = changePercent >= 0 ? '+' : '';
    el.textContent = `${sign}${changePercent.toFixed(1)}%`;
    el.className = 'card-change';
    if (invertColor) {
      // For expenses: increase is bad (negative), decrease is good (positive)
      el.classList.add(changePercent <= 0 ? 'positive' : 'negative');
    } else {
      // For income: increase is good (positive), decrease is bad (negative)
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
        </div>
      `;
    } else {
      container.innerHTML = sorted.map(t => this.renderRecentItem(t)).join('');
    }
  },
  
  // Render Transactions page
  renderTransactions() {
    let transactions = Storage.getAll();
    
    // Apply filters
    const filterMonth = document.getElementById('filterMonth')?.value;
    const filterType = document.getElementById('filterType')?.value || 'all';
    const filterCategory = document.getElementById('filterCategory')?.value || 'all';
    const filterSort = document.getElementById('filterSort')?.value || 'date-desc';
    
    // Filter by month
    if (filterMonth) {
      const [fy, fm] = filterMonth.split('-').map(Number);
      transactions = transactions.filter(t => {
        const d = new Date(t.date);
        return d.getFullYear() === fy && (d.getMonth() + 1) === fm;
      });
    }
    
    // Filter by type
    if (filterType !== 'all') {
      transactions = transactions.filter(t => t.type === filterType);
    }
    
    // Filter by category
    if (filterCategory !== 'all') {
      transactions = transactions.filter(t => t.category === filterCategory);
    }
    
    // Sort
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
      
      // Calculate totals
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
  
  // Render Analytics page
  renderAnalytics() {
    const { year, month } = this.analyticsMonth;
    const transactions = Storage.getByMonth(year, month);
    
    // Previous month
    const prev = navigateMonth(year, month, -1);
    const prevTransactions = Storage.getByMonth(prev.year, prev.month);
    
    // Calculate current month
    let income = 0, expense = 0;
    transactions.forEach(t => {
      if (t.type === 'income') income += t.amount;
      else expense += t.amount;
    });
    const savings = income - expense;
    const savingsRate = income > 0 ? ((savings / income) * 100) : 0;
    
    // Calculate previous month
    let prevIncome = 0, prevExpense = 0;
    prevTransactions.forEach(t => {
      if (t.type === 'income') prevIncome += t.amount;
      else prevExpense += t.amount;
    });
    const prevSavings = prevIncome - prevExpense;
    const prevSavingsRate = prevIncome > 0 ? ((prevSavings / prevIncome) * 100) : 0;
    
    // Update values
    document.getElementById('analyticsIncome').textContent = formatCurrency(income);
    document.getElementById('analyticsExpense').textContent = formatCurrency(expense);
    document.getElementById('analyticsSavings').textContent = formatCurrency(savings);
    document.getElementById('analyticsSavingsRate').textContent = savingsRate.toFixed(1) + '%';
    
    // Update month label
    document.getElementById('analyticsCurrentMonth').textContent = formatMonthLabel(year, month);
    
    // Update change badges
    this.updateAnalyticsChange('analyticsIncomeChange', income, prevIncome, false);
    this.updateAnalyticsChange('analyticsExpenseChange', expense, prevExpense, true);
    this.updateAnalyticsChange('analyticsSavingsChange', savings, prevSavings, false);
    this.updateAnalyticsChange('analyticsSavingsRateChange', savingsRate, prevSavingsRate, false);
    
    // Top categories
    this.renderTopCategories(transactions);
    
    // Charts
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
      if (invertColor) {
        badge.classList.add(change <= 0 ? 'up' : 'down');
      } else {
        badge.classList.add(change >= 0 ? 'up' : 'down');
      }
    }
  },

  renderTopCategories(transactions) {
    const container = document.getElementById('topCategories');
    if (!container) return;
    
    const expenses = transactions.filter(t => t.type === 'expense');
    if (expenses.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <span class="empty-icon">📊</span>
          <p>Chưa có dữ liệu</p>
        </div>
      `;
      return;
    }
    
    const categoryTotals = {};
    expenses.forEach(t => {
      categoryTotals[t.category] = (categoryTotals[t.category] || 0) + t.amount;
    });
    
    const totalExpense = expenses.reduce((sum, t) => sum + t.amount, 0);
    const sorted = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1]);
    
    container.innerHTML = sorted.map(([catId, amount]) => {
      const cat = getCategoryById(catId);
      const percent = ((amount / totalExpense) * 100).toFixed(1);
      return `
        <div class="top-category-item">
          <div class="top-cat-info">
            <span class="top-cat-emoji">${cat ? cat.emoji : '📦'}</span>
            <span class="top-cat-label">${cat ? cat.label : 'Khác'}</span>
          </div>
          <div class="top-cat-bar-wrapper">
            <div class="top-cat-bar" style="width: ${percent}%"></div>
          </div>
          <div class="top-cat-values">
            <span class="top-cat-amount">${formatCurrency(amount)}</span>
            <span class="top-cat-percent">${percent}%</span>
          </div>
        </div>
      `;
    }).join('');
  },
  
  // Transaction CRUD
  openModal(editId = null) {
    this.editingId = editId;
    const modal = document.getElementById('transactionModal');
    if (!modal) return;
    
    const form = document.getElementById('transactionForm');
    if (form) form.reset();
    
    if (editId) {
      document.getElementById('modalTitle').textContent = 'Sửa giao dịch';
      document.getElementById('editId').value = editId;
      const t = Storage.getAll().find(x => x.id === editId);
      if (t) {
        this.selectedType = t.type;
        this.selectedCategory = t.category;
        document.getElementById('amount').value = t.amount;
        document.getElementById('date').value = t.date;
        document.getElementById('note').value = t.note || '';
        
        // Update type toggle UI
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
    
    if (!amount || amount <= 0) {
      this.showToast('Vui lòng nhập số tiền hợp lệ', 'error');
      return;
    }
    if (!this.selectedCategory) {
      this.showToast('Vui lòng chọn danh mục', 'error');
      return;
    }
    if (!date) {
      this.showToast('Vui lòng chọn ngày', 'error');
      return;
    }
    
    const data = {
      type: this.selectedType,
      amount,
      category: this.selectedCategory,
      date,
      note
    };
    
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
      <button type="button" class="category-btn ${this.selectedCategory === cat.id ? 'active' : ''}" 
           data-id="${cat.id}">
        <span class="cat-icon">${cat.emoji}</span>
        <span>${cat.label}</span>
      </button>
    `).join('');
  },
  
  // Export
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
  
  // Toast notifications
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
  
  // Helper: render a transaction row
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
      </tr>
    `;
  },
  
  // Helper: render a recent transaction item  
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
      </div>
    `;
  },
  
  // Bind all events
  bindEvents() {
    // Nav clicks (both sidebar and any data-page buttons)
    document.addEventListener('click', (e) => {
      const navItem = e.target.closest('[data-page]');
      if (navItem) {
        e.preventDefault();
        this.navigate(navItem.dataset.page);
      }
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
      e.stopPropagation();
      this.openModal();
    });
    document.getElementById('fabAdd')?.addEventListener('click', () => this.openModal());
    
    // Modal close
    document.getElementById('modalClose')?.addEventListener('click', () => this.closeModal());
    document.getElementById('transactionModal')?.addEventListener('click', (e) => {
      if (e.target.id === 'transactionModal') this.closeModal();
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') this.closeModal();
    });
    
    // Form submit
    document.getElementById('transactionForm')?.addEventListener('submit', (e) => this.handleSubmit(e));
    
    // Type toggle in modal
    document.querySelectorAll('.type-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        this.selectType(e.currentTarget.dataset.type);
      });
    });
    
    // Category selection in modal (delegated)
    document.getElementById('categoryGrid')?.addEventListener('click', (e) => {
      const btn = e.target.closest('.category-btn');
      if (btn) this.selectCategory(btn.dataset.id);
    });
    
    // Filter changes
    ['filterMonth', 'filterType', 'filterCategory', 'filterSort'].forEach(id => {
      document.getElementById(id)?.addEventListener('change', () => this.renderTransactions());
    });
    
    // Theme toggle (both desktop sidebar and mobile)
    document.getElementById('themeToggle')?.addEventListener('click', () => this.toggleTheme());
    document.getElementById('themeToggleMobile')?.addEventListener('click', () => this.toggleTheme());
    
    // Export button (both desktop sidebar and mobile)
    document.getElementById('exportBtn')?.addEventListener('click', () => this.exportCSV());
    document.getElementById('exportBtnMobile')?.addEventListener('click', () => this.exportCSV());
    
    // Delegated click for edit/delete buttons in table
    document.getElementById('transactionsBody')?.addEventListener('click', (e) => {
      const editBtn = e.target.closest('.edit-btn');
      const delBtn = e.target.closest('.delete-btn');
      if (editBtn) this.openModal(editBtn.dataset.id);
      if (delBtn) this.deleteTransaction(delBtn.dataset.id);
    });
  }
};

// Start app
document.addEventListener('DOMContentLoaded', () => App.init());
