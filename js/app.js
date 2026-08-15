// ==================== APPS SCRIPT CODE (for copy button) ====================
const APPS_SCRIPT_CODE = `/**
 * CashFlow - Google Apps Script API
 * Dán code này vào Apps Script của Google Sheet.
 * Chạy hàm setupSheet() trước khi deploy.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Transactions');
  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
    sheet.appendRow(['ID','Loại','Số tiền','Danh mục','Ghi chú','Ngày','Thành viên','Ngày tạo']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:H1').setFontWeight('bold');
  }
  let loanSheet = ss.getSheetByName('Loans');
  if (!loanSheet) {
    loanSheet = ss.insertSheet('Loans');
    loanSheet.appendRow(['ID','Tên','Emoji','Loại','Gốc vay','Lãi suất','Kỳ hạn','Trả/tháng','Ngày BĐ','Ghi chú','Ngày tạo']);
    loanSheet.setFrozenRows(1);
    loanSheet.getRange('A1:K1').setFontWeight('bold');
  }
  let memberSheet = ss.getSheetByName('Members');
  if (!memberSheet) {
    memberSheet = ss.insertSheet('Members');
    memberSheet.appendRow(['ID','Tên','Avatar','Màu']);
    memberSheet.setFrozenRows(1);
    memberSheet.getRange('A1:D1').setFontWeight('bold');
  }
}
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
function findRow(sheet,id){var d=sheet.getDataRange().getValues();for(var i=1;i<d.length;i++){if(d[i][0]===id)return i+1;}return -1;}
function responseJson(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);}
function doGet(e) {
  var action = e.parameter.action || 'getAll';
  if (action==='ping') return responseJson({status:'ok',message:'Kết nối thành công!'});
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (action==='getAll') {
      var sheet=getSheet('Transactions');if(!sheet)return responseJson({success:true,data:[]});
      var d=sheet.getDataRange().getValues(),r=[];
      for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:d[i][0],type:d[i][1],amount:Number(d[i][2]),category:d[i][3],note:d[i][4],date:d[i][5],memberId:d[i][6],createdAt:d[i][7]});}
      return responseJson({success:true,data:r});
    }
    if (action==='add'){var o=JSON.parse(e.parameter.data);getSheet('Transactions').appendRow([o.id||'',o.type||'',o.amount||0,o.category||'',o.note||'',o.date||'',o.memberId||'',o.createdAt||'']);return responseJson({success:true});}
    if (action==='update'){var id=e.parameter.id,o=JSON.parse(e.parameter.data),s=getSheet('Transactions'),r=findRow(s,id);if(r>-1){s.getRange(r,1,1,8).setValues([[o.id||id,o.type||'',o.amount||0,o.category||'',o.note||'',o.date||'',o.memberId||'',o.createdAt||'']]);return responseJson({success:true});}throw new Error('Not found');}
    if (action==='delete'){var s=getSheet('Transactions'),r=findRow(s,e.parameter.id);if(r>-1){s.deleteRow(r);return responseJson({success:true});}throw new Error('Not found');}
    if (action==='sync'){var ts=JSON.parse(e.parameter.data),s=getSheet('Transactions'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ts&&ts.length>0){var rows=ts.map(function(t){return[t.id||'',t.type||'',t.amount||0,t.category||'',t.note||'',t.date||'',t.memberId||'',t.createdAt||''];});s.getRange(2,1,rows.length,8).setValues(rows);}return responseJson({success:true});}
    if (action==='getLoans'){var s=getSheet('Loans');if(!s)return responseJson({success:true,data:[]});var d=s.getDataRange().getValues(),r=[];for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:d[i][0],name:d[i][1],emoji:d[i][2],loanType:d[i][3],principal:Number(d[i][4]),interestRate:Number(d[i][5]),termMonths:Number(d[i][6]),monthlyPayment:Number(d[i][7]),startDate:d[i][8],note:d[i][9],createdAt:d[i][10]});}return responseJson({success:true,data:r});}
    if (action==='syncLoans'){var ls=JSON.parse(e.parameter.data),s=getSheet('Loans'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ls&&ls.length>0){var rows=ls.map(function(l){return[l.id||'',l.name||'',l.emoji||'',l.loanType||'',l.principal||0,l.interestRate||0,l.termMonths||0,l.monthlyPayment||0,l.startDate||'',l.note||'',l.createdAt||''];});s.getRange(2,1,rows.length,11).setValues(rows);}return responseJson({success:true});}
    if (action==='getMembers'){var s=getSheet('Members');if(!s)return responseJson({success:true,data:[]});var d=s.getDataRange().getValues(),r=[];for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:d[i][0],name:d[i][1],avatar:d[i][2],color:d[i][3]});}return responseJson({success:true,data:r});}
    if (action==='syncMembers'){var ms=JSON.parse(e.parameter.data),s=getSheet('Members'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ms&&ms.length>0){var rows=ms.map(function(m){return[m.id||'',m.name||'',m.avatar||'',m.color||''];});s.getRange(2,1,rows.length,4).setValues(rows);}return responseJson({success:true});}
    return responseJson({success:false,error:'Unknown action'});
  } catch(err) {return responseJson({success:false,error:err.toString()});} finally {lock.releaseLock();}
}`;

// ==================== MAIN APP ====================
const App = {
  currentPage: 'dashboard',
  currentMonth: getCurrentMonth(),
  analyticsMonth: getCurrentMonth(),
  editingId: null,
  editingLoanId: null,
  selectedType: 'expense',
  selectedCategory: null,
  selectedMemberId: null,
  selectedLoanType: null,

  // ==================== INIT ====================
  init() {
    this.initTheme();
    this.bindEvents();
    this.initFilterMonth();
    this.initSyncStatus();
    // Populate hidden script textarea
    const codeEl = document.getElementById('appsScriptCode');
    if (codeEl) codeEl.value = APPS_SCRIPT_CODE;
    // Check if setup is done
    if (!Storage.isSetupDone()) {
      this.showSetupModal();
    } else {
      this.renderCurrentPage();
    }
  },

  // ==================== SETUP WIZARD ====================
  showSetupModal() {
    const modal = document.getElementById('setupModal');
    if (modal) modal.classList.add('active');
    this.setupMemberRows = [
      { avatar: '👨', name: '' },
      { avatar: '👩', name: '' }
    ];
    this.renderSetupMembers();
  },

  renderSetupMembers() {
    const list = document.getElementById('setupMemberList');
    if (!list) return;
    list.innerHTML = this.setupMemberRows.map((m, i) => `
      <div class="setup-member-row" data-index="${i}">
        <button type="button" class="avatar-picker" data-index="${i}">${m.avatar}</button>
        <input type="text" value="${m.name}" placeholder="Tên thành viên" class="setup-member-name" data-index="${i}">
        ${this.setupMemberRows.length > 1 ? `<button type="button" class="remove-member-btn" data-index="${i}">✕</button>` : ''}
      </div>
    `).join('');
  },

  handleSetupSubmit(e) {
    e.preventDefault();
    const familyName = document.getElementById('familyName')?.value?.trim();
    if (!familyName) { this.showToast('Vui lòng nhập tên gia đình', 'error'); return; }

    // Collect member data from inputs
    const nameInputs = document.querySelectorAll('.setup-member-name');
    const members = [];
    this.setupMemberRows.forEach((row, i) => {
      const name = nameInputs[i]?.value?.trim();
      if (name) {
        members.push({
          name,
          avatar: row.avatar,
          color: MEMBER_COLORS[i % MEMBER_COLORS.length]
        });
      }
    });

    if (members.length === 0) { this.showToast('Thêm ít nhất 1 thành viên', 'error'); return; }

    Storage.setFamilyName(familyName);
    members.forEach(m => Storage.addMember(m));
    Storage.setSetupDone();

    const modal = document.getElementById('setupModal');
    if (modal) modal.classList.remove('active');
    this.showToast('Chào mừng ' + familyName + '! 🎉');
    this.renderCurrentPage();
  },

  showAvatarPicker(index, button) {
    // Close any existing popover
    document.querySelectorAll('.avatar-popover').forEach(p => p.remove());
    const popover = document.createElement('div');
    popover.className = 'avatar-popover';
    popover.innerHTML = AVATARS.map(a =>
      `<button type="button" class="avatar-option" data-emoji="${a.emoji}" data-index="${index}">${a.emoji}</button>`
    ).join('');
    button.parentElement.style.position = 'relative';
    button.parentElement.appendChild(popover);

    popover.addEventListener('click', (ev) => {
      const opt = ev.target.closest('.avatar-option');
      if (opt) {
        const idx = parseInt(opt.dataset.index);
        this.setupMemberRows[idx].avatar = opt.dataset.emoji;
        button.textContent = opt.dataset.emoji;
        popover.remove();
      }
    });

    // Close on outside click
    setTimeout(() => {
      document.addEventListener('click', function close(ev) {
        if (!popover.contains(ev.target) && ev.target !== button) {
          popover.remove();
          document.removeEventListener('click', close);
        }
      });
    }, 10);
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
    document.querySelectorAll('.theme-icon').forEach(el => { el.textContent = icon; });
  },
  toggleTheme() {
    const newTheme = Storage.getTheme() === 'light' ? 'dark' : 'light';
    Storage.setTheme(newTheme);
    document.documentElement.setAttribute('data-theme', newTheme);
    this.updateThemeToggleIcon(newTheme);
    Charts.setTheme(newTheme === 'dark');
    this.renderCurrentPage();
  },

  // ==================== NAVIGATION ====================
  navigate(page) {
    this.currentPage = page;
    document.querySelectorAll('.nav-item').forEach(el => el.classList.toggle('active', el.dataset.page === page));
    document.querySelectorAll('.page').forEach(el => el.classList.remove('active'));
    const pageEl = document.getElementById(`page-${page}`);
    if (pageEl) pageEl.classList.add('active');
    this.renderCurrentPage();
  },
  renderCurrentPage() {
    switch(this.currentPage) {
      case 'dashboard': this.renderDashboard(); break;
      case 'transactions': this.renderTransactions(); break;
      case 'loans': this.renderLoans(); break;
      case 'analytics': this.renderAnalytics(); break;
      case 'settings': this.renderSettings(); break;
    }
  },

  // ==================== SYNC ====================
  initSyncStatus() {
    this.updateSyncStatusBar();
    if (Storage.isOnline()) this.syncFromSheets(true);
  },
  updateSyncStatusBar() {
    const bar = document.getElementById('syncStatus');
    if (!bar) return;
    if (Storage.isOnline()) {
      bar.style.display = 'flex'; bar.className = 'sync-status connected';
      bar.querySelector('.sync-text').textContent = '🟢 Đã kết nối Google Sheets';
    } else { bar.style.display = 'none'; }
  },
  async syncFromSheets(silent = false) {
    if (!Storage.isOnline()) return;
    const bar = document.getElementById('syncStatus');
    if (bar) { bar.style.display = 'flex'; bar.className = 'sync-status'; bar.querySelector('.sync-text').textContent = '🔄 Đang đồng bộ...'; }
    try {
      await Storage.syncFromSheets();
      if (bar) { bar.className = 'sync-status connected'; bar.querySelector('.sync-text').textContent = '🟢 Đã đồng bộ'; }
      this.renderCurrentPage();
      if (!silent) this.showToast('Đã đồng bộ ✅');
    } catch(e) {
      if (bar) { bar.className = 'sync-status offline'; bar.querySelector('.sync-text').textContent = '🔴 Lỗi đồng bộ'; }
      if (!silent) this.showToast('Lỗi đồng bộ', 'error');
    }
  },

  // ==================== INIT FILTERS ====================
  initFilterMonth() {
    const f = document.getElementById('filterMonth');
    if (f) { const d = new Date(); f.value = `${d.getFullYear()}-${(d.getMonth()+1).toString().padStart(2,'0')}`; }
    this.populateCategoryFilter();
  },
  populateCategoryFilter() {
    const f = document.getElementById('filterCategory');
    if (!f) return;
    f.innerHTML = '<option value="all">Tất cả</option>';
    getAllCategories().forEach(c => { f.innerHTML += `<option value="${c.id}">${c.emoji} ${c.label}</option>`; });
  },

  // ==================== DASHBOARD ====================
  renderDashboard() {
    const { year, month } = this.currentMonth;
    const transactions = Storage.getByMonth(year, month);
    const prev = navigateMonth(year, month, -1);
    const prevTransactions = Storage.getByMonth(prev.year, prev.month);

    let income = 0, expense = 0;
    transactions.forEach(t => { if (t.type === 'income') income += t.amount; else expense += t.amount; });
    const balance = income - expense;
    const savingsRate = income > 0 ? ((income - expense) / income * 100) : 0;

    let prevIncome = 0, prevExpense = 0;
    prevTransactions.forEach(t => { if (t.type === 'income') prevIncome += t.amount; else prevExpense += t.amount; });

    // Greeting
    const greeting = document.getElementById('greetingText');
    if (greeting) greeting.textContent = `Xin chào, ${Storage.getFamilyName()}! 👋`;

    // Family avatars
    this.renderFamilyAvatars();

    // Mood
    const expenseChange = calcChange(expense, prevExpense);
    const mood = analyzeMood(savingsRate, expenseChange);
    const moodCard = document.getElementById('moodCard');
    if (moodCard) {
      moodCard.style.background = mood.bgColor;
      moodCard.style.borderColor = mood.color;
    }
    const moodEmoji = document.getElementById('moodEmoji');
    if (moodEmoji) moodEmoji.textContent = mood.emoji;
    const moodLabel = document.getElementById('moodLabel');
    if (moodLabel) { moodLabel.textContent = mood.label; moodLabel.style.color = mood.color; }
    const moodMsg = document.getElementById('moodMessage');
    if (moodMsg) moodMsg.textContent = mood.message;

    // Smart tips
    const tips = generateSmartTips(transactions, prevTransactions);
    const tipsCont = document.getElementById('tipsContainer');
    if (tipsCont) {
      if (tips.length > 0) {
        tipsCont.style.display = 'flex';
        tipsCont.innerHTML = tips.map(t => `<div class="tip-item ${t.type}"><span>${t.emoji}</span><span>${t.text}</span></div>`).join('');
      } else { tipsCont.style.display = 'none'; }
    }

    // Summary cards
    document.getElementById('totalIncome').textContent = formatCurrency(income);
    document.getElementById('totalExpense').textContent = formatCurrency(expense);
    document.getElementById('balance').textContent = formatCurrency(balance);
    document.getElementById('savingsRate').textContent = savingsRate.toFixed(1) + '%';
    this.updateChangeBadge('incomeChange', calcChange(income, prevIncome), false);
    this.updateChangeBadge('expenseChange', expenseChange, true);
    document.getElementById('currentMonth').textContent = formatMonthLabel(year, month);

    // Charts
    Charts.renderCategoryChart('categoryChart', transactions);
    const monthsData = Storage.getLastNMonths(6);
    Charts.renderMonthlyChart('monthlyChart', monthsData);
    Charts.renderTrendChart('trendChart', monthsData);
    this.renderRecentTransactions(transactions);
  },

  renderFamilyAvatars() {
    const container = document.getElementById('familyAvatars');
    if (!container) return;
    const members = Storage.getMembers();
    container.innerHTML = members.map(m =>
      `<div class="family-avatar" style="border-color: ${m.color}" title="${m.name}">${m.avatar}</div>`
    ).join('');
  },

  updateChangeBadge(id, change, invert) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = `${change >= 0 ? '+' : ''}${change.toFixed(1)}%`;
    el.className = 'card-change';
    el.classList.add(invert ? (change <= 0 ? 'positive' : 'negative') : (change >= 0 ? 'positive' : 'negative'));
  },

  renderRecentTransactions(transactions) {
    const c = document.getElementById('recentTransactions');
    if (!c) return;
    const sorted = [...transactions].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 5);
    if (sorted.length === 0) {
      c.innerHTML = `<div class="empty-state"><span class="empty-icon">📝</span><p>Chưa có giao dịch nào</p><p class="text-muted">Bấm nút + để thêm giao dịch đầu tiên</p></div>`;
    } else {
      c.innerHTML = sorted.map(t => this.renderRecentItem(t)).join('');
    }
  },

  // ==================== TRANSACTIONS PAGE ====================
  renderTransactions() {
    let txs = Storage.getLocal();
    const fm = document.getElementById('filterMonth')?.value;
    const ft = document.getElementById('filterType')?.value || 'all';
    const fc = document.getElementById('filterCategory')?.value || 'all';
    const fs = document.getElementById('filterSort')?.value || 'date-desc';
    if (fm) { const [y,m] = fm.split('-').map(Number); txs = txs.filter(t => { const d=new Date(t.date); return d.getFullYear()===y&&(d.getMonth()+1)===m; }); }
    if (ft !== 'all') txs = txs.filter(t => t.type === ft);
    if (fc !== 'all') txs = txs.filter(t => t.category === fc);
    switch(fs) {
      case 'date-desc': txs.sort((a,b) => new Date(b.date)-new Date(a.date)); break;
      case 'date-asc': txs.sort((a,b) => new Date(a.date)-new Date(b.date)); break;
      case 'amount-desc': txs.sort((a,b) => b.amount-a.amount); break;
      case 'amount-asc': txs.sort((a,b) => a.amount-b.amount); break;
    }
    const tbody = document.getElementById('transactionsBody');
    const empty = document.getElementById('emptyTransactions');
    const table = document.getElementById('transactionsTable');
    if (!tbody) return;
    if (txs.length === 0) {
      tbody.innerHTML = ''; if (table) table.style.display = 'none'; if (empty) empty.style.display = 'flex';
    } else {
      if (table) table.style.display = ''; if (empty) empty.style.display = 'none';
      tbody.innerHTML = txs.map(t => this.renderTransactionRow(t)).join('');
      let ti=0, te=0; txs.forEach(t => { if(t.type==='income') ti+=t.amount; else te+=t.amount; });
      const net = ti-te;
      const tot = document.getElementById('tableTotal');
      if (tot) tot.innerHTML = `<strong class="${net>=0?'amount-col income':'amount-col expense'}">${formatCurrency(net)}</strong>`;
    }
  },

  // ==================== LOANS PAGE ====================
  renderLoans() {
    const loans = Storage.getLoans();
    let totalDebt = 0, totalMonthly = 0, totalInterest = 0;
    loans.forEach(l => {
      const status = calculateLoanStatus(l);
      totalDebt += status.remainingBalance;
      totalMonthly += l.monthlyPayment;
      totalInterest += status.totalInterestPaid;
    });

    document.getElementById('totalDebt').textContent = formatCurrency(totalDebt);
    document.getElementById('monthlyPaymentTotal').textContent = formatCurrency(totalMonthly);
    document.getElementById('totalInterestPaid').textContent = formatCurrency(totalInterest);

    const list = document.getElementById('loansList');
    if (!list) return;
    if (loans.length === 0) {
      list.innerHTML = `<div class="empty-state"><span class="empty-icon">💳</span><p>Chưa có khoản vay nào</p><p class="text-muted">Bấm "Thêm khoản vay" để bắt đầu theo dõi</p></div>`;
    } else {
      list.innerHTML = loans.map(l => this.renderLoanCard(l)).join('');
    }
  },

  renderLoanCard(loan) {
    const status = calculateLoanStatus(loan);
    const loanType = LOAN_TYPES.find(t => t.id === loan.loanType);
    const emoji = loan.emoji || (loanType ? loanType.emoji : '💳');
    const endDate = status.estimatedEndDate;
    const endDateStr = `${endDate.getMonth()+1}/${endDate.getFullYear()}`;

    return `
      <div class="loan-card">
        <div class="loan-header">
          <div class="loan-title">
            <span class="loan-emoji">${emoji}</span>
            <div>
              <h4>${loan.name}</h4>
              ${loan.note ? `<span class="loan-note">${loan.note}</span>` : ''}
            </div>
          </div>
          <div class="loan-actions">
            <button class="edit-btn edit-loan-btn" data-id="${loan.id}" title="Sửa">✏️</button>
            <button class="delete-btn delete-loan-btn" data-id="${loan.id}" title="Xóa">🗑️</button>
          </div>
        </div>
        <div class="loan-progress">
          <div class="loan-progress-bar"><div class="loan-progress-fill" style="width:${status.progressPercent.toFixed(1)}%"></div></div>
          <div class="loan-progress-text">
            <span>Đã trả ${status.progressPercent.toFixed(1)}%</span>
            <span>Dự kiến hết: ${endDateStr}</span>
          </div>
        </div>
        <div class="loan-stats">
          <div class="loan-stat"><div class="loan-stat-label">Gốc vay</div><div class="loan-stat-value">${formatCurrency(loan.principal)}</div></div>
          <div class="loan-stat"><div class="loan-stat-label">Dư nợ còn lại</div><div class="loan-stat-value">${formatCurrency(status.remainingBalance)}</div></div>
          <div class="loan-stat"><div class="loan-stat-label">Lãi suất</div><div class="loan-stat-value">${loan.interestRate}%/năm</div></div>
          <div class="loan-stat"><div class="loan-stat-label">Trả/tháng</div><div class="loan-stat-value">${formatCurrency(loan.monthlyPayment)}</div></div>
        </div>
      </div>`;
  },

  // ==================== ANALYTICS PAGE ====================
  renderAnalytics() {
    const { year, month } = this.analyticsMonth;
    const txs = Storage.getByMonth(year, month);
    const prev = navigateMonth(year, month, -1);
    const prevTxs = Storage.getByMonth(prev.year, prev.month);
    let inc=0, exp=0; txs.forEach(t => { if(t.type==='income') inc+=t.amount; else exp+=t.amount; });
    let pInc=0, pExp=0; prevTxs.forEach(t => { if(t.type==='income') pInc+=t.amount; else pExp+=t.amount; });
    const sav = inc-exp, sr = inc>0?((sav/inc)*100):0;
    const pSav = pInc-pExp, pSr = pInc>0?((pSav/pInc)*100):0;

    document.getElementById('analyticsIncome').textContent = formatCurrency(inc);
    document.getElementById('analyticsExpense').textContent = formatCurrency(exp);
    document.getElementById('analyticsSavings').textContent = formatCurrency(sav);
    document.getElementById('analyticsSavingsRate').textContent = sr.toFixed(1)+'%';
    document.getElementById('analyticsCurrentMonth').textContent = formatMonthLabel(year, month);

    this.updateAnalyticsChange('analyticsIncomeChange', inc, pInc, false);
    this.updateAnalyticsChange('analyticsExpenseChange', exp, pExp, true);
    this.updateAnalyticsChange('analyticsSavingsChange', sav, pSav, false);
    this.updateAnalyticsChange('analyticsSavingsRateChange', sr, pSr, false);
    this.renderTopCategories(txs);
    Charts.renderCategoryTrendChart('categoryTrendChart', Storage.getLastNMonths(6));
  },

  updateAnalyticsChange(cId, cur, prev, inv) {
    const c = document.getElementById(cId); if(!c) return;
    const ch = calcChange(cur, prev);
    const badge = c.querySelector('.change-badge');
    if(badge) { badge.textContent = `${ch>=0?'+':''}${ch.toFixed(1)}%`; badge.className='change-badge'; badge.classList.add(inv?(ch<=0?'up':'down'):(ch>=0?'up':'down')); }
  },

  renderTopCategories(txs) {
    const c = document.getElementById('topCategories'); if(!c) return;
    const exps = txs.filter(t => t.type==='expense');
    if(exps.length===0) { c.innerHTML = `<div class="empty-state"><span class="empty-icon">📊</span><p>Chưa có dữ liệu</p></div>`; return; }
    const totals = {}; exps.forEach(t => { totals[t.category]=(totals[t.category]||0)+t.amount; });
    const total = exps.reduce((s,t)=>s+t.amount,0);
    c.innerHTML = Object.entries(totals).sort((a,b)=>b[1]-a[1]).map(([id,amt]) => {
      const cat = getCategoryById(id); const pct = ((amt/total)*100).toFixed(1);
      return `<div class="top-category-item"><div class="top-cat-info"><span class="top-cat-emoji">${cat?cat.emoji:'📦'}</span><span class="top-cat-label">${cat?cat.label:'Khác'}</span></div><div class="top-cat-bar-wrapper"><div class="top-cat-bar" style="width:${pct}%"></div></div><div class="top-cat-values"><span class="top-cat-amount">${formatCurrency(amt)}</span><span class="top-cat-percent">${pct}%</span></div></div>`;
    }).join('');
  },

  // ==================== SETTINGS PAGE ====================
  renderSettings() {
    const st = document.getElementById('syncToggle');
    const ac = document.getElementById('apiConfig');
    const au = document.getElementById('apiUrl');
    if (st) st.checked = Storage.getSyncMode() === 'sheets';
    if (ac) ac.style.display = Storage.getSyncMode() === 'sheets' ? 'block' : 'none';
    if (au) au.value = Storage.getApiUrl();
    if (Storage.isOnline()) { const sa = document.getElementById('syncActions'); if(sa) sa.style.display = 'flex'; }
  },

  async handleTestConnection() {
    const url = document.getElementById('apiUrl')?.value?.trim();
    if (!url) { this.showToast('Nhập URL', 'error'); return; }
    Storage.setApiUrl(url);
    const se = document.getElementById('connectionStatus'), si = document.getElementById('statusIcon'), st = document.getElementById('statusText');
    if(se) { se.style.display='flex'; se.className='connection-status loading'; } if(si) si.textContent='⏳'; if(st) st.textContent='Đang kiểm tra...';
    const r = await Storage.testConnection();
    if (r.success) {
      if(se) se.className='connection-status success'; if(si) si.textContent='✅'; if(st) st.textContent=r.message;
      const sa = document.getElementById('syncActions'); if(sa) sa.style.display='flex';
      Storage.setSyncMode('sheets'); this.updateSyncStatusBar(); this.showToast('Kết nối thành công! ✅');
    } else {
      if(se) se.className='connection-status error'; if(si) si.textContent='❌'; if(st) st.textContent=r.message;
      this.showToast('Lỗi: '+r.message, 'error');
    }
  },

  handleSyncToggle() {
    const checked = document.getElementById('syncToggle')?.checked;
    Storage.setSyncMode(checked ? 'sheets' : 'local');
    const ac = document.getElementById('apiConfig'); if(ac) ac.style.display = checked ? 'block' : 'none';
    if (!checked) { const sa = document.getElementById('syncActions'); if(sa) sa.style.display='none'; }
    this.updateSyncStatusBar();
  },
  async handleSyncPull() { this.showToast('Đang tải...'); await this.syncFromSheets(false); },
  async handleSyncPush() {
    if (!confirm('Đẩy dữ liệu lên Sheet? Dữ liệu cũ trên Sheet sẽ bị ghi đè.')) return;
    try { await Storage.uploadAllToSheets(); await Storage.syncMembersToSheets(); await Storage.syncLoansToSheets(); this.showToast('Đã đẩy lên Sheet ✅'); }
    catch(e) { this.showToast('Lỗi: '+e.message,'error'); }
  },
  handleCopyScript() {
    const code = APPS_SCRIPT_CODE;
    navigator.clipboard.writeText(code).then(() => this.showToast('Đã copy! 📋')).catch(() => {
      const ta = document.createElement('textarea'); ta.value = code; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta); this.showToast('Đã copy! 📋');
    });
  },
  handleExportJson() {
    const data = { transactions: Storage.getLocal(), members: Storage.getMembers(), loans: Storage.getLoans(), familyName: Storage.getFamilyName() };
    const blob = new Blob([JSON.stringify(data,null,2)], {type:'application/json'});
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob);
    link.download = `cashflow_backup_${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    this.showToast('Đã xuất JSON 📤');
  },
  handleImportJson(e) {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = JSON.parse(evt.target.result);
        if (data.transactions) Storage.saveLocal(data.transactions);
        if (data.members) Storage.saveMembers(data.members);
        if (data.loans) Storage.saveLoans(data.loans);
        if (data.familyName) Storage.setFamilyName(data.familyName);
        this.showToast('Đã nhập dữ liệu! ✅'); this.renderCurrentPage();
      } catch(err) { this.showToast('File không hợp lệ','error'); }
    };
    reader.readAsText(file); e.target.value = '';
  },
  handleClearData() {
    if (!confirm('XÓA TOÀN BỘ dữ liệu? Không thể hoàn tác!')) return;
    if (!confirm('Xác nhận lần nữa?')) return;
    Storage.saveLocal([]); Storage.saveLoans([]); this.showToast('Đã xóa 🗑️'); this.renderCurrentPage();
  },

  // ==================== TRANSACTION MODAL ====================
  openModal(editId = null) {
    this.editingId = editId;
    const modal = document.getElementById('transactionModal'); if(!modal) return;
    const form = document.getElementById('transactionForm'); if(form) form.reset();
    this.renderMemberSelector();

    if (editId) {
      document.getElementById('modalTitle').textContent = 'Sửa giao dịch';
      document.getElementById('editId').value = editId;
      const t = Storage.getLocal().find(x => x.id === editId);
      if (t) {
        this.selectedType = t.type; this.selectedCategory = t.category; this.selectedMemberId = t.memberId || null;
        document.getElementById('amount').value = t.amount;
        document.getElementById('date').value = t.date;
        document.getElementById('note').value = t.note || '';
        document.querySelectorAll('.type-btn').forEach(b => { b.classList.remove('active'); if(b.dataset.type===t.type) b.classList.add('active'); });
        this.renderMemberSelector();
      }
    } else {
      document.getElementById('modalTitle').textContent = 'Thêm giao dịch';
      document.getElementById('editId').value = '';
      this.selectedType = 'expense'; this.selectedCategory = null;
      // Auto-select first member
      const members = Storage.getMembers();
      this.selectedMemberId = members.length > 0 ? members[0].id : null;
      document.getElementById('date').value = new Date().toISOString().split('T')[0];
      document.querySelectorAll('.type-btn').forEach(b => { b.classList.remove('active'); if(b.dataset.type==='expense') b.classList.add('active'); });
      this.renderMemberSelector();
    }
    this.renderCategories();
    modal.classList.add('active');
    setTimeout(() => document.getElementById('amount')?.focus(), 100);
  },

  renderMemberSelector() {
    const c = document.getElementById('memberSelector'); if(!c) return;
    const members = Storage.getMembers();
    c.innerHTML = members.map(m => `
      <button type="button" class="member-btn ${this.selectedMemberId === m.id ? 'active' : ''}" data-id="${m.id}" style="--member-color:${m.color}">
        <span class="member-avatar">${m.avatar}</span>
        <span>${m.name}</span>
      </button>
    `).join('');
  },

  closeModal() {
    document.getElementById('transactionModal')?.classList.remove('active');
    this.editingId = null;
  },

  handleSubmit(e) {
    e.preventDefault();
    const amount = parseFloat(document.getElementById('amount').value);
    const date = document.getElementById('date').value;
    const note = document.getElementById('note').value;
    if (!amount || amount <= 0) { this.showToast('Nhập số tiền hợp lệ','error'); return; }
    if (!this.selectedCategory) { this.showToast('Chọn danh mục','error'); return; }
    if (!date) { this.showToast('Chọn ngày','error'); return; }
    const data = { type: this.selectedType, amount, category: this.selectedCategory, date, note, memberId: this.selectedMemberId };
    if (this.editingId) { Storage.update(this.editingId, data); this.showToast('Đã cập nhật ✅'); }
    else { Storage.add(data); this.showToast('Đã thêm giao dịch ✅'); }
    this.closeModal(); this.renderCurrentPage();
  },

  deleteTransaction(id) {
    if (confirm('Xóa giao dịch này?')) { Storage.delete(id); this.showToast('Đã xóa 🗑️'); this.renderCurrentPage(); }
  },

  selectType(type) {
    this.selectedType = type; this.selectedCategory = null;
    document.querySelectorAll('.type-btn').forEach(b => { b.classList.remove('active'); if(b.dataset.type===type) b.classList.add('active'); });
    this.renderCategories();
  },
  selectCategory(catId) {
    this.selectedCategory = catId;
    document.querySelectorAll('.category-btn').forEach(el => el.classList.toggle('active', el.dataset.id===catId));
  },
  renderCategories() {
    const c = document.getElementById('categoryGrid'); if(!c) return;
    c.innerHTML = CATEGORIES[this.selectedType].map(cat =>
      `<button type="button" class="category-btn ${this.selectedCategory===cat.id?'active':''}" data-id="${cat.id}"><span class="cat-icon">${cat.emoji}</span><span>${cat.label}</span></button>`
    ).join('');
  },

  // ==================== LOAN MODAL ====================
  openLoanModal(editId = null) {
    this.editingLoanId = editId;
    const modal = document.getElementById('loanModal'); if(!modal) return;
    const form = document.getElementById('loanForm'); if(form) form.reset();
    this.renderLoanTypes();

    if (editId) {
      document.getElementById('loanModalTitle').textContent = 'Sửa khoản vay';
      document.getElementById('editLoanId').value = editId;
      const l = Storage.getLoans().find(x => x.id === editId);
      if (l) {
        this.selectedLoanType = l.loanType;
        document.getElementById('loanName').value = l.name;
        document.getElementById('loanPrincipal').value = l.principal;
        document.getElementById('loanRate').value = l.interestRate;
        document.getElementById('loanTerm').value = l.termMonths;
        document.getElementById('loanMonthly').value = l.monthlyPayment;
        document.getElementById('loanStartDate').value = l.startDate;
        document.getElementById('loanNote').value = l.note || '';
        this.renderLoanTypes();
      }
    } else {
      document.getElementById('loanModalTitle').textContent = 'Thêm khoản vay';
      document.getElementById('editLoanId').value = '';
      this.selectedLoanType = null;
      document.getElementById('loanStartDate').value = new Date().toISOString().split('T')[0];
    }
    modal.classList.add('active');
  },

  closeLoanModal() {
    document.getElementById('loanModal')?.classList.remove('active');
    this.editingLoanId = null;
  },

  renderLoanTypes() {
    const c = document.getElementById('loanTypeGrid'); if(!c) return;
    c.innerHTML = LOAN_TYPES.map(lt =>
      `<button type="button" class="loan-type-btn ${this.selectedLoanType===lt.id?'active':''}" data-id="${lt.id}"><span class="loan-type-emoji">${lt.emoji}</span><span>${lt.label}</span></button>`
    ).join('');
  },

  handleLoanSubmit(e) {
    e.preventDefault();
    const data = {
      name: document.getElementById('loanName').value,
      loanType: this.selectedLoanType,
      emoji: LOAN_TYPES.find(t => t.id === this.selectedLoanType)?.emoji || '💳',
      principal: parseFloat(document.getElementById('loanPrincipal').value),
      interestRate: parseFloat(document.getElementById('loanRate').value),
      termMonths: parseInt(document.getElementById('loanTerm').value),
      monthlyPayment: parseFloat(document.getElementById('loanMonthly').value),
      startDate: document.getElementById('loanStartDate').value,
      note: document.getElementById('loanNote').value
    };
    if (!data.name || !data.principal) { this.showToast('Nhập đầy đủ thông tin','error'); return; }
    if (this.editingLoanId) { Storage.updateLoan(this.editingLoanId, data); this.showToast('Đã cập nhật ✅'); }
    else { Storage.addLoan(data); this.showToast('Đã thêm khoản vay ✅'); }
    this.closeLoanModal(); this.renderLoans();
  },

  deleteLoan(id) {
    if (confirm('Xóa khoản vay này?')) { Storage.deleteLoan(id); this.showToast('Đã xóa 🗑️'); this.renderLoans(); }
  },

  // ==================== EXPORT CSV ====================
  exportCSV() {
    const csv = Storage.exportCSV();
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob);
    link.download = `cashflow_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    this.showToast('Đã xuất CSV 📥');
  },

  // ==================== TOAST ====================
  showToast(msg, type='success') {
    const c = document.getElementById('toastContainer'); if(!c) return;
    const t = document.createElement('div'); t.className = `toast toast-${type}`;
    t.innerHTML = `<span>${msg}</span>`; c.appendChild(t);
    setTimeout(() => { t.classList.add('hide'); setTimeout(() => t.remove(), 300); }, 3000);
  },

  // ==================== RENDER HELPERS ====================
  renderTransactionRow(t) {
    const cat = getCategoryById(t.category);
    const member = t.memberId ? Storage.getMemberById(t.memberId) : null;
    const memberBadge = member ? `<span class="family-avatar" style="border-color:${member.color};width:24px;height:24px;font-size:0.8rem;display:inline-flex">${member.avatar}</span>` : '';
    return `<tr>
      <td>${formatDate(t.date)}</td>
      <td>${memberBadge}</td>
      <td><span class="badge">${t.type==='income'?'Thu':'Chi'}</span></td>
      <td>${cat?cat.emoji+' '+cat.label:'📦 Khác'}</td>
      <td>${t.note||'—'}</td>
      <td class="amount-col ${t.type==='income'?'income':'expense'}">${t.type==='income'?'+':'-'}${formatCurrency(t.amount)}</td>
      <td><div class="action-btns"><button class="edit-btn" data-id="${t.id}">✏️</button><button class="delete-btn" data-id="${t.id}">🗑️</button></div></td>
    </tr>`;
  },

  renderRecentItem(t) {
    const cat = getCategoryById(t.category);
    const member = t.memberId ? Storage.getMemberById(t.memberId) : null;
    const isInc = t.type === 'income';
    return `<div class="recent-item">
      <div class="recent-icon">${cat?cat.emoji:'📦'}</div>
      <div class="recent-info">
        <div class="recent-title">${member?member.avatar+' ':''} ${cat?cat.label:'Khác'}${t.note?' · '+t.note:''}</div>
        <div class="recent-date">${formatDate(t.date)}</div>
      </div>
      <div class="recent-amount ${isInc?'income':'expense'}">${isInc?'+':'-'}${formatCurrency(t.amount)}</div>
    </div>`;
  },

  // ==================== EVENT BINDINGS ====================
  bindEvents() {
    // Nav
    document.addEventListener('click', (e) => { const n=e.target.closest('[data-page]'); if(n){e.preventDefault();this.navigate(n.dataset.page);} });
    // Dashboard month nav
    document.getElementById('prevMonth')?.addEventListener('click', () => { this.currentMonth=navigateMonth(this.currentMonth.year,this.currentMonth.month,-1); this.renderDashboard(); });
    document.getElementById('nextMonth')?.addEventListener('click', () => { this.currentMonth=navigateMonth(this.currentMonth.year,this.currentMonth.month,1); this.renderDashboard(); });
    // Analytics month nav
    document.getElementById('analyticsPrevMonth')?.addEventListener('click', () => { this.analyticsMonth=navigateMonth(this.analyticsMonth.year,this.analyticsMonth.month,-1); this.renderAnalytics(); });
    document.getElementById('analyticsNextMonth')?.addEventListener('click', () => { this.analyticsMonth=navigateMonth(this.analyticsMonth.year,this.analyticsMonth.month,1); this.renderAnalytics(); });
    // Transaction modal
    document.getElementById('addTransactionBtn')?.addEventListener('click', (e) => { e.stopPropagation(); this.openModal(); });
    document.getElementById('fabAdd')?.addEventListener('click', () => this.openModal());
    document.getElementById('modalClose')?.addEventListener('click', () => this.closeModal());
    document.getElementById('transactionModal')?.addEventListener('click', (e) => { if(e.target.id==='transactionModal') this.closeModal(); });
    document.getElementById('transactionForm')?.addEventListener('submit', (e) => this.handleSubmit(e));
    document.querySelectorAll('.type-btn').forEach(b => b.addEventListener('click', (e) => { e.preventDefault(); this.selectType(e.currentTarget.dataset.type); }));
    document.getElementById('categoryGrid')?.addEventListener('click', (e) => { const b=e.target.closest('.category-btn'); if(b) this.selectCategory(b.dataset.id); });
    // Member selector in form
    document.getElementById('memberSelector')?.addEventListener('click', (e) => {
      const b = e.target.closest('.member-btn');
      if (b) { this.selectedMemberId = b.dataset.id; this.renderMemberSelector(); }
    });
    // Loan modal
    document.getElementById('addLoanBtn')?.addEventListener('click', () => this.openLoanModal());
    document.getElementById('loanModalClose')?.addEventListener('click', () => this.closeLoanModal());
    document.getElementById('loanModal')?.addEventListener('click', (e) => { if(e.target.id==='loanModal') this.closeLoanModal(); });
    document.getElementById('loanForm')?.addEventListener('submit', (e) => this.handleLoanSubmit(e));
    document.getElementById('loanTypeGrid')?.addEventListener('click', (e) => { const b=e.target.closest('.loan-type-btn'); if(b){this.selectedLoanType=b.dataset.id;this.renderLoanTypes();} });
    // Loan list edit/delete
    document.getElementById('loansList')?.addEventListener('click', (e) => {
      const eb=e.target.closest('.edit-loan-btn'); const db=e.target.closest('.delete-loan-btn');
      if(eb) this.openLoanModal(eb.dataset.id);
      if(db) this.deleteLoan(db.dataset.id);
    });
    // Filters
    ['filterMonth','filterType','filterCategory','filterSort'].forEach(id => { document.getElementById(id)?.addEventListener('change', () => this.renderTransactions()); });
    // Theme
    document.getElementById('themeToggle')?.addEventListener('click', () => this.toggleTheme());
    document.getElementById('themeToggleMobile')?.addEventListener('click', () => this.toggleTheme());
    // Export CSV
    document.getElementById('exportBtn')?.addEventListener('click', () => this.exportCSV());
    document.getElementById('exportBtnMobile')?.addEventListener('click', () => this.exportCSV());
    // Table edit/delete
    document.getElementById('transactionsBody')?.addEventListener('click', (e) => {
      const eb=e.target.closest('.edit-btn'); const db=e.target.closest('.delete-btn');
      if(eb) this.openModal(eb.dataset.id); if(db) this.deleteTransaction(db.dataset.id);
    });
    // Settings
    document.getElementById('syncToggle')?.addEventListener('change', () => this.handleSyncToggle());
    document.getElementById('testConnectionBtn')?.addEventListener('click', () => this.handleTestConnection());
    document.getElementById('syncPullBtn')?.addEventListener('click', () => this.handleSyncPull());
    document.getElementById('syncPushBtn')?.addEventListener('click', () => this.handleSyncPush());
    document.getElementById('copyScriptBtn')?.addEventListener('click', () => this.handleCopyScript());
    document.getElementById('exportJsonBtn')?.addEventListener('click', () => this.handleExportJson());
    document.getElementById('importJsonInput')?.addEventListener('change', (e) => this.handleImportJson(e));
    document.getElementById('clearDataBtn')?.addEventListener('click', () => this.handleClearData());
    // Setup
    document.getElementById('setupForm')?.addEventListener('submit', (e) => this.handleSetupSubmit(e));
    document.getElementById('addSetupMember')?.addEventListener('click', () => {
      if (this.setupMemberRows.length < 6) {
        this.setupMemberRows.push({ avatar: AVATARS[this.setupMemberRows.length % AVATARS.length].emoji, name: '' });
        this.renderSetupMembers();
      } else { this.showToast('Tối đa 6 thành viên','error'); }
    });
    // Delegated: setup member list
    document.getElementById('setupMemberList')?.addEventListener('click', (e) => {
      const picker = e.target.closest('.avatar-picker');
      if (picker) this.showAvatarPicker(parseInt(picker.dataset.index), picker);
      const rm = e.target.closest('.remove-member-btn');
      if (rm) { this.setupMemberRows.splice(parseInt(rm.dataset.index), 1); this.renderSetupMembers(); }
    });
    document.getElementById('setupMemberList')?.addEventListener('input', (e) => {
      if (e.target.classList.contains('setup-member-name')) {
        this.setupMemberRows[parseInt(e.target.dataset.index)].name = e.target.value;
      }
    });
    // Escape key
    document.addEventListener('keydown', (e) => { if(e.key==='Escape') { this.closeModal(); this.closeLoanModal(); } });
  }
};

document.addEventListener('DOMContentLoaded', () => App.init());
