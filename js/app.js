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
    sheet.appendRow(['ID','Loại','Số tiền','Danh mục','Ghi chú','Ngày','Thành viên','Ngày tạo','Chi cho ai']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:I1').setFontWeight('bold');
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
    memberSheet.appendRow(['ID','Tên','Avatar','Màu','AvatarImg','AvatarId']);
    memberSheet.setFrozenRows(1);
    memberSheet.getRange('A1:F1').setFontWeight('bold');
  }
}
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
function findRow(sheet,id){var d=sheet.getDataRange().getValues();for(var i=1;i<d.length;i++){if(d[i][0]===id)return i+1;}return -1;}
function responseJson(data, callback){
  var json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
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
  var action = e.parameter.action || 'getAll';
  var cb = e.parameter.callback || null;
  if (action==='ping') return responseJson({status:'ok',message:'Kết nối thành công!'}, cb);
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (action==='getAll') {
      var sheet=getSheet('Transactions');if(!sheet)return responseJson({success:true,data:[]}, cb);
      var d=sheet.getDataRange().getValues(),r=[];
      for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:String(d[i][0]),type:String(d[i][1]||'expense'),amount:Number(d[i][2]||0),category:String(d[i][3]||'other'),note:String(d[i][4]||''),date:String(d[i][5]||''),memberId:String(d[i][6]||''),createdAt:String(d[i][7]||''),beneficiaryId:String(d[i][8]||'')});}
      return responseJson({success:true,data:r}, cb);
    }
    if (action==='add'){var o=JSON.parse(e.parameter.data);getSheet('Transactions').appendRow([o.id||'',o.type||'',o.amount||0,o.category||'',o.note||'',o.date||'',o.memberId||'',o.createdAt||'',o.beneficiaryId||'']);return responseJson({success:true}, cb);}
    if (action==='update'){var id=e.parameter.id,o=JSON.parse(e.parameter.data),s=getSheet('Transactions'),r=findRow(s,id);if(r>-1){s.getRange(r,1,1,9).setValues([[o.id||id,o.type||'',o.amount||0,o.category||'',o.note||'',o.date||'',o.memberId||'',o.createdAt||'',o.beneficiaryId||'']]);return responseJson({success:true}, cb);}throw new Error('Not found');}
    if (action==='delete'){var s=getSheet('Transactions'),r=findRow(s,e.parameter.id);if(r>-1){s.deleteRow(r);return responseJson({success:true}, cb);}throw new Error('Not found');}
    if (action==='sync'){var ts=JSON.parse(e.parameter.data),s=getSheet('Transactions'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ts&&ts.length>0){var rows=ts.map(function(t){return[t.id||'',t.type||'',t.amount||0,t.category||'',t.note||'',t.date||'',t.memberId||'',t.createdAt||'',t.beneficiaryId||''];});s.getRange(2,1,rows.length,9).setValues(rows);}return responseJson({success:true}, cb);}
    if (action==='getLoans'){var s=getSheet('Loans');if(!s)return responseJson({success:true,data:[]}, cb);var d=s.getDataRange().getValues(),r=[];for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:d[i][0],name:d[i][1],emoji:d[i][2],loanType:d[i][3],principal:Number(d[i][4]),interestRate:Number(d[i][5]),termMonths:Number(d[i][6]),monthlyPayment:Number(d[i][7]),startDate:d[i][8],note:d[i][9],createdAt:d[i][10]});}return responseJson({success:true,data:r}, cb);}
    if (action==='syncLoans'){var ls=JSON.parse(e.parameter.data),s=getSheet('Loans'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ls&&ls.length>0){var rows=ls.map(function(l){return[l.id||'',l.name||'',l.emoji||'',l.loanType||'',l.principal||0,l.interestRate||0,l.termMonths||0,l.monthlyPayment||0,l.startDate||'',l.note||'',l.createdAt||''];});s.getRange(2,1,rows.length,11).setValues(rows);}return responseJson({success:true}, cb);}
    if (action==='getMembers'){var s=getSheet('Members');if(!s)return responseJson({success:true,data:[]}, cb);var d=s.getDataRange().getValues(),r=[];for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:String(d[i][0]),name:String(d[i][1]||''),avatar:String(d[i][2]||'👤'),color:String(d[i][3]||'#e77d3e'),avatarImg:String(d[i][4]||''),avatarId:String(d[i][5]||'')});}return responseJson({success:true,data:r}, cb);}
    if (action==='syncMembers'){var ms=JSON.parse(e.parameter.data),s=getSheet('Members'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ms&&ms.length>0){var rows=ms.map(function(m){return[m.id||'',m.name||'',m.avatar||'',m.color||'',m.avatarImg||'',m.avatarId||''];});s.getRange(2,1,rows.length,6).setValues(rows);}return responseJson({success:true}, cb);}
    return responseJson({success:false,error:'Unknown action'}, cb);
  } catch(err) {return responseJson({success:false,error:err.toString()}, cb);} finally {lock.releaseLock();}
}`;

// ==================== MAIN APP ====================
const App = {
  currentPage: 'dashboard',
  currentMonth: getCurrentMonth(),
  analyticsMonth: getCurrentMonth(),
  dashboardScope: 'all',
  editingId: null,
  editingLoanId: null,
  selectedType: 'expense',
  selectedCategory: null,
  selectedMemberId: null,
  selectedBeneficiaryId: 'family',
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
    // Init notifications
    if ('Notification' in window && Notification.permission === 'granted') {
      this.registerServiceWorker();
      this.scheduleReminderCheck();
    }
  },

  // ==================== SETUP WIZARD ====================
  showSetupModal() {
    const modal = document.getElementById('setupModal');
    if (modal) modal.classList.add('active');
    this.setupMemberRows = [
      { avatarId: 'avatar_dad', name: '' },
      { avatarId: 'avatar_mom', name: '' }
    ];
    this.renderSetupMembers();
  },

  renderSetupMembers() {
    const list = document.getElementById('setupMemberList');
    if (!list) return;
    list.innerHTML = this.setupMemberRows.map((m, i) => {
      const av = AVATARS.find(a => a.id === m.avatarId) || AVATARS[0];
      return `
      <div class="setup-member-row" data-index="${i}">
        <button type="button" class="avatar-picker" data-index="${i}"><img src="${av.img}" alt="${av.label}" class="avatar-img"></button>
        <input type="text" value="${m.name}" placeholder="Tên thành viên" class="setup-member-name" data-index="${i}">
        ${this.setupMemberRows.length > 1 ? `<button type="button" class="remove-member-btn" data-index="${i}">✕</button>` : ''}
      </div>`;
    }).join('');
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
        const av = AVATARS.find(a => a.id === row.avatarId) || AVATARS[i % AVATARS.length];
        members.push({
          name,
          avatarId: row.avatarId,
          avatar: av.emoji,
          avatarImg: av.img,
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
    const currentRow = this.setupMemberRows[index];
    const currentAv = AVATARS.find(a => a.id === currentRow.avatarId) || AVATARS[0];
    const currentImg = currentRow.avatarImg || currentAv.img;

    this.openAvatarModal({
      currentImg: currentImg,
      onSelect: (selected) => {
        this.setupMemberRows[index].avatarId = selected.id || 'custom';
        this.setupMemberRows[index].avatarImg = selected.img;
        this.setupMemberRows[index].avatar = selected.emoji || '👤';
        button.innerHTML = `<img src="${selected.img}" alt="Avatar" class="avatar-img">`;
      }
    });
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

    // Auto-sync when user switches back to app tab
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible' && Storage.isOnline()) {
        this.syncFromSheets(true);
      }
    });
    window.addEventListener('focus', () => {
      if (Storage.isOnline()) {
        this.syncFromSheets(true);
      }
    });

    // Periodic auto-sync every 30 seconds
    setInterval(() => {
      if (Storage.isOnline() && document.visibilityState === 'visible') {
        this.syncFromSheets(true);
      }
    }, 30000);
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
      this.renderFamilyAvatars();
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
    let transactions = Storage.getByMonth(year, month);
    const prev = navigateMonth(year, month, -1);
    let prevTransactions = Storage.getByMonth(prev.year, prev.month);

    // Apply dashboard scope filter (all | family | personal)
    if (this.dashboardScope === 'family') {
      transactions = transactions.filter(t => !t.beneficiaryId || t.beneficiaryId === 'family');
      prevTransactions = prevTransactions.filter(t => !t.beneficiaryId || t.beneficiaryId === 'family');
    } else if (this.dashboardScope === 'personal') {
      transactions = transactions.filter(t => t.beneficiaryId && t.beneficiaryId !== 'family');
      prevTransactions = prevTransactions.filter(t => t.beneficiaryId && t.beneficiaryId !== 'family');
    }

    // Update scope buttons active status
    document.querySelectorAll('.scope-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.scope === this.dashboardScope);
    });

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
    container.innerHTML = members.map(m => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      return `<div class="family-avatar" style="border-color: ${m.color}" title="${m.name}"><img src="${imgSrc}" alt="${m.name}" class="avatar-img"></div>`;
    }).join('');
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
    this.renderMemberManagement();
    this.renderReminderSettings();
  },

  // ==================== MEMBER MANAGEMENT ====================
  _editingMemberId: null,

  renderMemberManagement() {
    const list = document.getElementById('memberManagementList');
    if (!list) return;
    const members = Storage.getMembers();
    list.innerHTML = members.map((m, i) => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      const isEditing = this._editingMemberId === m.id;
      return `
        <div class="member-manage-row" data-id="${m.id}">
          <div class="member-manage-avatar" data-id="${m.id}" data-action="change-avatar" title="Đổi avatar">
            <img src="${imgSrc}" alt="${m.name}">
          </div>
          <div class="member-manage-info">
            ${isEditing
              ? `<input class="member-manage-name-input" value="${m.name}" data-id="${m.id}" autofocus>`
              : `<span class="member-manage-name">${m.name}</span>`
            }
            <span><span class="member-manage-color" style="background:${m.color}"></span> ${isEditing ? 'Đang sửa...' : ''}</span>
          </div>
          <div class="member-manage-actions">
            ${isEditing
              ? `<button class="save-member-btn" data-id="${m.id}" title="Lưu">✓</button>
                 <button data-id="${m.id}" data-action="cancel-edit" title="Hủy">✕</button>`
              : `<button data-id="${m.id}" data-action="edit-member" title="Sửa tên">✏️</button>
                 <button class="delete-member-btn" data-id="${m.id}" data-action="delete-member" title="Xóa">🗑️</button>`
            }
          </div>
        </div>`;
    }).join('');
  },

  handleMemberManagementClick(e) {
    const btn = e.target.closest('button[data-action], div[data-action]');
    if (!btn) {
      // Check save button
      const saveBtn = e.target.closest('.save-member-btn');
      if (saveBtn) {
        const id = saveBtn.dataset.id;
        const input = document.querySelector(`.member-manage-name-input[data-id="${id}"]`);
        if (input && input.value.trim()) {
          Storage.updateMember(id, { name: input.value.trim() });
          this._editingMemberId = null;
          this.renderMemberManagement();
          this.showToast('Đã cập nhật ✅');
          this.renderFamilyAvatars();
        }
        return;
      }
      return;
    }
    const action = btn.dataset.action;
    const id = btn.dataset.id;

    if (action === 'edit-member') {
      this._editingMemberId = id;
      this.renderMemberManagement();
      setTimeout(() => {
        document.querySelector(`.member-manage-name-input[data-id="${id}"]`)?.focus();
      }, 50);
    } else if (action === 'cancel-edit') {
      this._editingMemberId = null;
      this.renderMemberManagement();
    } else if (action === 'delete-member') {
      const member = Storage.getMemberById(id);
      const members = Storage.getMembers();
      if (members.length <= 1) {
        this.showToast('Cần ít nhất 1 thành viên', 'error');
        return;
      }
      if (confirm(`Xóa thành viên "${member?.name}"? Giao dịch cũ vẫn giữ nguyên.`)) {
        Storage.deleteMember(id);
        this.renderMemberManagement();
        this.renderFamilyAvatars();
        this.showToast('Đã xóa 🗑️');
      }
    } else if (action === 'change-avatar') {
      this.showMemberAvatarPicker(id, btn);
    }
  },

  showMemberAvatarPicker(memberId, button) {
    const member = Storage.getMemberById(memberId);
    if (!member) return;
    const currentAv = AVATARS.find(a => a.id === member.avatarId) || AVATARS[0];
    const currentImg = member.avatarImg || currentAv.img;

    this.openAvatarModal({
      currentImg: currentImg,
      onSelect: (selected) => {
        Storage.updateMember(memberId, {
          avatarId: selected.id || 'custom',
          avatarImg: selected.img,
          avatar: selected.emoji || '👤'
        });
        this.renderMemberManagement();
        this.renderFamilyAvatars();
        this.showToast('Đã đổi avatar thành công! ✅');
      }
    });
  },

  // ==================== AVATAR MODAL CONTROLLER ====================
  _avatarModalCallback: null,
  _currentSelectedAvatar: null,

  openAvatarModal({ currentImg, onSelect }) {
    const modal = document.getElementById('avatarModal');
    if (!modal) return;
    this._avatarModalCallback = onSelect;
    this._currentSelectedAvatar = { img: currentImg || AVATARS[0].img };

    // Update preview
    const previewImg = document.getElementById('uploadPreviewImg');
    if (previewImg) previewImg.src = this._currentSelectedAvatar.img;

    // Render presets
    this.renderAvatarPresets();

    modal.classList.add('active');
  },

  closeAvatarModal() {
    const modal = document.getElementById('avatarModal');
    if (modal) modal.classList.remove('active');
    this._avatarModalCallback = null;
    this._currentSelectedAvatar = null;
  },

  renderAvatarPresets() {
    const grid = document.getElementById('avatarPresetsGrid');
    if (!grid) return;
    grid.innerHTML = AVATARS.map(a => {
      const isSelected = this._currentSelectedAvatar && this._currentSelectedAvatar.img === a.img;
      return `
        <button type="button" class="avatar-preset-card ${isSelected ? 'selected' : ''}" data-avatar-id="${a.id}">
          <img src="${a.img}" alt="${a.label}" class="preset-img">
          <span class="preset-label">${a.label}</span>
        </button>
      `;
    }).join('');
  },

  handleAvatarFileSelect(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    this.showToast('Đang xử lý ảnh...');
    processImageFile(file, 128).then(dataUrl => {
      this._currentSelectedAvatar = {
        id: 'uploaded_' + Date.now(),
        img: dataUrl,
        emoji: '📸',
        label: 'Ảnh tự tải'
      };
      const previewImg = document.getElementById('uploadPreviewImg');
      if (previewImg) previewImg.src = dataUrl;
      // Deselect presets
      document.querySelectorAll('.avatar-preset-card').forEach(c => c.classList.remove('selected'));
      this.showToast('Đã tải ảnh lên! Bấm Áp dụng để lưu ✅');
    }).catch(err => {
      this.showToast(err.message || 'Lỗi đọc ảnh', 'error');
    });
  },

  handlePresetSelect(e) {
    const card = e.target.closest('.avatar-preset-card');
    if (!card) return;
    const avId = card.dataset.avatarId;
    const av = AVATARS.find(a => a.id === avId);
    if (!av) return;

    this._currentSelectedAvatar = {
      id: av.id,
      img: av.img,
      emoji: av.emoji,
      label: av.label
    };

    const previewImg = document.getElementById('uploadPreviewImg');
    if (previewImg) previewImg.src = av.img;

    document.querySelectorAll('.avatar-preset-card').forEach(c => c.classList.remove('selected'));
    card.classList.add('selected');
  },

  confirmAvatarSelection() {
    if (this._currentSelectedAvatar && this._avatarModalCallback) {
      this._avatarModalCallback(this._currentSelectedAvatar);
    }
    this.closeAvatarModal();
  },

  handleAddMember() {
    const members = Storage.getMembers();
    if (members.length >= 6) {
      this.showToast('Tối đa 6 thành viên', 'error');
      return;
    }
    const name = prompt('Tên thành viên mới:');
    if (!name || !name.trim()) return;
    const avIndex = members.length % AVATARS.length;
    const av = AVATARS[avIndex];
    Storage.addMember({
      name: name.trim(),
      avatarId: av.id,
      avatar: av.emoji,
      avatarImg: av.img,
      color: MEMBER_COLORS[members.length % MEMBER_COLORS.length]
    });
    this.renderMemberManagement();
    this.renderFamilyAvatars();
    this.showToast('Đã thêm thành viên ✅');
  },

  // ==================== NOTIFICATIONS ====================
  renderReminderSettings() {
    const list = document.getElementById('reminderMemberList');
    const statusEl = document.getElementById('notificationStatus');
    const btn = document.getElementById('enableNotifBtn');
    if (!list) return;

    // Show notification permission status
    if ('Notification' in window) {
      const perm = Notification.permission;
      if (perm === 'granted') {
        if (statusEl) { statusEl.style.display = 'block'; statusEl.className = 'notification-status granted'; statusEl.textContent = '✅ Đã bật thông báo'; }
        if (btn) btn.style.display = 'none';
      } else if (perm === 'denied') {
        if (statusEl) { statusEl.style.display = 'block'; statusEl.className = 'notification-status denied'; statusEl.textContent = '❌ Thông báo bị chặn. Vào cài đặt trình duyệt để bật lại.'; }
        if (btn) btn.style.display = 'none';
      } else {
        if (statusEl) statusEl.style.display = 'none';
        if (btn) btn.style.display = 'block';
      }
    } else {
      if (statusEl) { statusEl.style.display = 'block'; statusEl.className = 'notification-status denied'; statusEl.textContent = 'Trình duyệt không hỗ trợ thông báo'; }
      if (btn) btn.style.display = 'none';
    }

    // Render member reminder rows
    const members = Storage.getMembers();
    const reminders = JSON.parse(localStorage.getItem(REMINDER_KEY) || '{}');
    list.innerHTML = members.map(m => {
      const reminder = reminders[m.id] || { enabled: false, time: '21:00' };
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      return `
        <div class="reminder-member-row">
          <div class="member-info">
            <img src="${imgSrc}" alt="${m.name}">
            <span>${m.name}</span>
          </div>
          <div class="reminder-time">
            <span>Nhắc lúc</span>
            <input type="time" value="${reminder.time}" data-member="${m.id}" class="reminder-time-input">
          </div>
          <label class="reminder-toggle">
            <input type="checkbox" ${reminder.enabled ? 'checked' : ''} data-member="${m.id}" class="reminder-toggle-input">
            <span class="slider"></span>
          </label>
        </div>`;
    }).join('');
  },

  async requestNotificationPermission() {
    if (!('Notification' in window)) {
      this.showToast('Trình duyệt không hỗ trợ', 'error');
      return;
    }
    const perm = await Notification.requestPermission();
    if (perm === 'granted') {
      this.showToast('Đã bật thông báo! 🔔');
      // Register service worker for background checks
      this.registerServiceWorker();
    } else {
      this.showToast('Thông báo bị từ chối', 'error');
    }
    this.renderReminderSettings();
  },

  async registerServiceWorker() {
    if ('serviceWorker' in navigator) {
      try {
        await navigator.serviceWorker.register('./sw.js');
      } catch (e) {
        console.warn('SW registration failed:', e);
      }
    }
  },

  saveReminderSettings(memberId, field, value) {
    const reminders = JSON.parse(localStorage.getItem(REMINDER_KEY) || '{}');
    if (!reminders[memberId]) reminders[memberId] = { enabled: false, time: '21:00' };
    reminders[memberId][field] = value;
    localStorage.setItem(REMINDER_KEY, JSON.stringify(reminders));
    // Schedule check
    this.scheduleReminderCheck();
  },

  scheduleReminderCheck() {
    // Clear existing timer
    if (this._reminderTimer) clearTimeout(this._reminderTimer);
    const reminders = JSON.parse(localStorage.getItem(REMINDER_KEY) || '{}');
    const now = new Date();
    let nearest = null;

    for (const memberId in reminders) {
      const r = reminders[memberId];
      if (!r.enabled) continue;
      const [h, min] = r.time.split(':').map(Number);
      const target = new Date(now);
      target.setHours(h, min, 0, 0);
      if (target <= now) target.setDate(target.getDate() + 1); // tomorrow
      if (!nearest || target < nearest.time) {
        nearest = { memberId, time: target };
      }
    }

    if (nearest) {
      const delay = nearest.time - now;
      this._reminderTimer = setTimeout(() => this.fireReminder(nearest.memberId), delay);
    }
  },

  fireReminder(memberId) {
    // Check if any transaction was added today by this member
    const today = new Date().toISOString().split('T')[0];
    const txs = Storage.getLocal();
    const hasTodayTx = txs.some(t => t.date === today && t.memberId === memberId);

    if (!hasTodayTx) {
      const member = Storage.getMemberById(memberId);
      const name = member ? member.name : 'Bạn';
      if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('CashFlow - Nhắc nhở 📝', {
          body: `${name} ơi, hôm nay chưa nhập giao dịch nào. Đừng quên ghi chép nhé!`,
          icon: member?.avatarImg || 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="90">💰</text></svg>'
        });
      }
    }
    // Reschedule for next day
    this.scheduleReminderCheck();
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

    if (editId) {
      document.getElementById('modalTitle').textContent = 'Sửa giao dịch';
      document.getElementById('editId').value = editId;
      const t = Storage.getLocal().find(x => x.id === editId);
      if (t) {
        this.selectedType = t.type;
        this.selectedCategory = t.category;
        this.selectedMemberId = t.memberId || null;
        this.selectedBeneficiaryId = t.beneficiaryId || 'family';
        document.getElementById('amount').value = formatNumberInput(t.amount);
        document.getElementById('date').value = t.date;
        document.getElementById('note').value = t.note || '';
        document.querySelectorAll('.type-btn').forEach(b => {
          b.classList.toggle('active', b.dataset.type === t.type);
        });
      }
    } else {
      document.getElementById('modalTitle').textContent = 'Thêm giao dịch';
      document.getElementById('editId').value = '';
      this.selectedType = 'expense';
      this.selectedCategory = null;
      const members = Storage.getMembers();
      this.selectedMemberId = members.length > 0 ? members[0].id : null;
      this.selectedBeneficiaryId = 'family';
      document.getElementById('amount').value = '';
      document.getElementById('date').value = new Date().toISOString().split('T')[0];
      document.querySelectorAll('.type-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.type === 'expense');
      });
    }
    this.renderMemberSelector();
    this.renderBeneficiarySelector();
    this.renderCategories();
    modal.classList.add('active');
    setTimeout(() => document.getElementById('amount')?.focus(), 100);
  },

  renderMemberSelector() {
    const c = document.getElementById('memberSelector'); if(!c) return;
    const members = Storage.getMembers();
    c.innerHTML = members.map(m => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      return `
      <button type="button" class="member-btn ${this.selectedMemberId === m.id ? 'active' : ''}" data-id="${m.id}" style="--member-color:${m.color}">
        <span class="member-avatar"><img src="${imgSrc}" alt="${m.name}" class="avatar-img"></span>
        <span>${m.name}</span>
      </button>`;
    }).join('');
  },

  renderBeneficiarySelector() {
    const c = document.getElementById('beneficiarySelector'); if(!c) return;
    const members = Storage.getMembers();
    const isFamily = !this.selectedBeneficiaryId || this.selectedBeneficiaryId === 'family';
    let html = `
      <button type="button" class="beneficiary-btn ${isFamily ? 'active' : ''}" data-id="family">
        <span>👨‍👩‍👧‍👦 Cả nhà</span>
      </button>
    `;
    html += members.map(m => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      const isActive = this.selectedBeneficiaryId === m.id;
      return `
        <button type="button" class="beneficiary-btn ${isActive ? 'active' : ''}" data-id="${m.id}" style="--member-color:${m.color}">
          <img src="${imgSrc}" alt="${m.name}" class="avatar-img-sm" style="width:16px;height:16px;border-radius:50%;">
          <span>${m.name}</span>
        </button>
      `;
    }).join('');
    c.innerHTML = html;
  },

  closeModal() {
    document.getElementById('transactionModal')?.classList.remove('active');
    this.editingId = null;
  },

  handleSubmit(e) {
    e.preventDefault();
    const rawAmount = document.getElementById('amount').value;
    const amount = parseNumberInput(rawAmount);
    const date = document.getElementById('date').value;
    const note = document.getElementById('note').value;
    if (!amount || amount <= 0) { this.showToast('Nhập số tiền hợp lệ','error'); return; }
    if (!this.selectedCategory) { this.showToast('Chọn danh mục','error'); return; }
    if (!date) { this.showToast('Chọn ngày','error'); return; }
    const beneficiaryId = this.selectedBeneficiaryId || 'family';
    const data = {
      type: this.selectedType,
      amount,
      category: this.selectedCategory,
      date,
      note,
      memberId: this.selectedMemberId,
      beneficiaryId
    };
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

  // ==================== KPI DRILLDOWN MODAL ====================
  openKPIModal(type) {
    const { year, month } = this.currentMonth;
    let transactions = Storage.getByMonth(year, month);
    if (this.dashboardScope === 'family') {
      transactions = transactions.filter(t => !t.beneficiaryId || t.beneficiaryId === 'family');
    } else if (this.dashboardScope === 'personal') {
      transactions = transactions.filter(t => t.beneficiaryId && t.beneficiaryId !== 'family');
    }

    const modal = document.getElementById('kpiDrilldownModal');
    const titleEl = document.getElementById('kpiDrilldownTitle');
    const subEl = document.getElementById('kpiDrilldownSubtitle');
    const totalValEl = document.getElementById('kpiTotalVal');
    const listEl = document.getElementById('kpiDrilldownList');
    if (!modal || !listEl) return;

    let items = [];
    let total = 0;
    let title = '';
    let subtitle = `Tháng ${month}/${year}`;
    if (this.dashboardScope === 'family') subtitle += ' · 👨‍👩‍👧‍👦 Chi tiêu chung';
    else if (this.dashboardScope === 'personal') subtitle += ' · 👤 Chi tiêu riêng';

    if (type === 'income') {
      title = '📥 Chi tiết Thu nhập';
      items = transactions.filter(t => t.type === 'income');
      total = items.reduce((s, t) => s + t.amount, 0);
    } else if (type === 'expense') {
      title = '📤 Chi tiết Chi tiêu';
      items = transactions.filter(t => t.type === 'expense');
      total = items.reduce((s, t) => s + t.amount, 0);
    } else if (type === 'balance') {
      title = '💎 Chi tiết Số dư';
      items = transactions;
      const inc = transactions.filter(t => t.type === 'income').reduce((s, t) => s + t.amount, 0);
      const exp = transactions.filter(t => t.type === 'expense').reduce((s, t) => s + t.amount, 0);
      total = inc - exp;
    } else if (type === 'savings') {
      title = '🎯 Tỷ lệ tiết kiệm';
      items = transactions.filter(t => t.type === 'expense');
      total = items.reduce((s, t) => s + t.amount, 0);
    }

    titleEl.textContent = title;
    subEl.textContent = subtitle;
    totalValEl.textContent = formatCurrency(total);
    totalValEl.className = 'kpi-total-val ' + (type === 'income' ? 'income' : (type === 'expense' ? 'expense' : ''));

    if (items.length === 0) {
      listEl.innerHTML = '<div class="kpi-empty-state">Chưa có giao dịch nào trong mục này 🍃</div>';
    } else {
      listEl.innerHTML = items.map(t => {
        const cat = getCategoryById(t.category);
        const payer = t.memberId ? Storage.getMemberById(t.memberId) : null;
        const bene = t.beneficiaryId && t.beneficiaryId !== 'family' ? Storage.getMemberById(t.beneficiaryId) : null;
        const isInc = t.type === 'income';
        
        const payerImg = payer ? (payer.avatarImg || (AVATARS.find(a => a.id === payer.avatarId) || AVATARS[0]).img) : '';
        const beneBadge = bene ? `👤 Cho ${bene.name}` : '👨‍👩‍👧‍👦 Cả nhà';

        return `
          <div class="kpi-drilldown-item">
            <div class="kpi-item-left">
              <div class="kpi-item-icon">${cat ? cat.emoji : '📦'}</div>
              <div class="kpi-item-info">
                <div class="kpi-item-title">${cat ? cat.label : 'Khác'}${t.note ? ' · ' + t.note : ''}</div>
                <div class="kpi-item-sub">
                  <span>${formatDate(t.date)}</span>
                  ${payer ? `<span class="kpi-item-badge"><img src="${payerImg}" style="width:12px;height:12px;border-radius:50%;vertical-align:middle;"> ${payer.name}</span>` : ''}
                  <span class="kpi-item-badge">${beneBadge}</span>
                </div>
              </div>
            </div>
            <div class="kpi-item-amount ${isInc ? 'income' : 'expense'}">${isInc ? '+' : '-'}${formatCurrency(t.amount)}</div>
          </div>
        `;
      }).join('');
    }

    modal.classList.add('active');
  },

  closeKPIModal() {
    document.getElementById('kpiDrilldownModal')?.classList.remove('active');
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
        document.getElementById('loanPrincipal').value = formatNumberInput(l.principal);
        document.getElementById('loanRate').value = l.interestRate;
        document.getElementById('loanTerm').value = l.termMonths;
        document.getElementById('loanMonthly').value = formatNumberInput(l.monthlyPayment);
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
      principal: parseNumberInput(document.getElementById('loanPrincipal').value),
      interestRate: parseFloat(document.getElementById('loanRate').value),
      termMonths: parseInt(document.getElementById('loanTerm').value),
      monthlyPayment: parseNumberInput(document.getElementById('loanMonthly').value),
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
    const bene = t.beneficiaryId && t.beneficiaryId !== 'family' ? Storage.getMemberById(t.beneficiaryId) : null;
    const memberBadge = member ? `<span class="family-avatar" style="border-color:${member.color};width:28px;height:28px;display:inline-flex"><img src="${member.avatarImg || (AVATARS.find(a => a.id === member.avatarId) || AVATARS[0]).img}" alt="${member.name}" class="avatar-img"></span>` : '';
    const beneTag = bene ? `<span style="font-size:0.7rem;color:var(--text-secondary);margin-left:4px;">(Cho ${bene.name})</span>` : '';

    return `<tr>
      <td>${formatDate(t.date)}</td>
      <td>${memberBadge}</td>
      <td><span class="badge">${t.type==='income'?'Thu':'Chi'}</span></td>
      <td>${cat?cat.emoji+' '+cat.label:'📦 Khác'} ${beneTag}</td>
      <td>${t.note||'—'}</td>
      <td class="amount-col ${t.type==='income'?'income':'expense'}">${t.type==='income'?'+':'-'}${formatCurrency(t.amount)}</td>
      <td><div class="action-btns"><button class="edit-btn" data-id="${t.id}">✏️</button><button class="delete-btn" data-id="${t.id}">🗑️</button></div></td>
    </tr>`;
  },

  renderRecentItem(t) {
    const cat = getCategoryById(t.category);
    const member = t.memberId ? Storage.getMemberById(t.memberId) : null;
    const bene = t.beneficiaryId && t.beneficiaryId !== 'family' ? Storage.getMemberById(t.beneficiaryId) : null;
    const isInc = t.type === 'income';
    const beneText = bene ? ` · Cho ${bene.name}` : '';

    return `<div class="recent-item">
      <div class="recent-icon">${cat?cat.emoji:'📦'}</div>
      <div class="recent-info">
        <div class="recent-title">${member?`<img src="${member.avatarImg || (AVATARS.find(a => a.id === member.avatarId) || AVATARS[0]).img}" class="avatar-img-sm"> `:''} ${cat?cat.label:'Khác'}${beneText}${t.note?' · '+t.note:''}</div>
        <div class="recent-date">${formatDate(t.date)}</div>
      </div>
      <div class="recent-amount ${isInc?'income':'expense'}">${isInc?'+':'-'}${formatCurrency(t.amount)}</div>
    </div>`;
  },

  // ==================== EVENT BINDINGS ====================
  bindEvents() {
    // Nav
    document.addEventListener('click', (e) => { const n=e.target.closest('[data-page]'); if(n){e.preventDefault();this.navigate(n.dataset.page);} });
    
    // Dashboard scope filter
    document.getElementById('dashboardScopeToggle')?.addEventListener('click', (e) => {
      const btn = e.target.closest('.scope-btn');
      if (btn) {
        this.dashboardScope = btn.dataset.scope;
        this.renderDashboard();
      }
    });

    // KPI Drilldown Cards click
    document.getElementById('kpiIncomeCard')?.addEventListener('click', () => this.openKPIModal('income'));
    document.getElementById('kpiExpenseCard')?.addEventListener('click', () => this.openKPIModal('expense'));
    document.getElementById('kpiBalanceCard')?.addEventListener('click', () => this.openKPIModal('balance'));
    document.getElementById('kpiSavingsCard')?.addEventListener('click', () => this.openKPIModal('savings'));
    document.getElementById('kpiDrilldownClose')?.addEventListener('click', () => this.closeKPIModal());
    document.getElementById('kpiDrilldownModal')?.addEventListener('click', (e) => {
      if (e.target.id === 'kpiDrilldownModal') this.closeKPIModal();
    });

    // Dashboard month nav
    document.getElementById('prevMonth')?.addEventListener('click', () => { this.currentMonth=navigateMonth(this.currentMonth.year,this.currentMonth.month,-1); this.renderDashboard(); });
    document.getElementById('nextMonth')?.addEventListener('click', () => { this.currentMonth=navigateMonth(this.currentMonth.year,this.currentMonth.month,1); this.renderDashboard(); });
    
    // Analytics month nav
    document.getElementById('analyticsPrevMonth')?.addEventListener('click', () => { this.analyticsMonth=navigateMonth(this.analyticsMonth.year,this.analyticsMonth.month,-1); this.renderAnalytics(); });
    document.getElementById('analyticsNextMonth')?.addEventListener('click', () => { this.analyticsMonth=navigateMonth(this.analyticsMonth.year,this.analyticsMonth.month,1); this.renderAnalytics(); });
    
    // Currency inputs auto-formatter (e.g. 5000 -> 5.000)
    ['amount', 'loanPrincipal', 'loanMonthly'].forEach(id => {
      const el = document.getElementById(id);
      if (el) {
        el.addEventListener('input', (e) => {
          e.target.value = formatNumberInput(e.target.value);
        });
      }
    });

    // Transaction modal
    document.getElementById('addTransactionBtn')?.addEventListener('click', (e) => { e.stopPropagation(); this.openModal(); });
    document.getElementById('fabAdd')?.addEventListener('click', () => this.openModal());
    document.getElementById('modalClose')?.addEventListener('click', () => this.closeModal());
    document.getElementById('transactionModal')?.addEventListener('click', (e) => { if(e.target.id==='transactionModal') this.closeModal(); });
    document.getElementById('transactionForm')?.addEventListener('submit', (e) => this.handleSubmit(e));
    document.querySelectorAll('.type-btn').forEach(b => b.addEventListener('click', (e) => { e.preventDefault(); this.selectType(e.currentTarget.dataset.type); }));
    document.getElementById('categoryGrid')?.addEventListener('click', (e) => { const b=e.target.closest('.category-btn'); if(b) this.selectCategory(b.dataset.id); });
    
    // Member selector in form (Ai chi?)
    document.getElementById('memberSelector')?.addEventListener('click', (e) => {
      const b = e.target.closest('.member-btn');
      if (b) { this.selectedMemberId = b.dataset.id; this.renderMemberSelector(); }
    });

    // Beneficiary selector in form (Chi cho ai?)
    document.getElementById('beneficiarySelector')?.addEventListener('click', (e) => {
      const b = e.target.closest('.beneficiary-btn');
      if (b) { this.selectedBeneficiaryId = b.dataset.id; this.renderBeneficiarySelector(); }
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
        this.setupMemberRows.push({ avatarId: AVATARS[this.setupMemberRows.length % AVATARS.length].id, name: '' });
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
    
    // Member management
    document.getElementById('memberManagementList')?.addEventListener('click', (e) => this.handleMemberManagementClick(e));
    document.getElementById('memberManagementList')?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target.classList.contains('member-manage-name-input')) {
        const id = e.target.dataset.id;
        if (e.target.value.trim()) {
          Storage.updateMember(id, { name: e.target.value.trim() });
          this._editingMemberId = null;
          this.renderMemberManagement();
          this.showToast('Đã cập nhật ✅');
          this.renderFamilyAvatars();
        }
      }
    });
    document.getElementById('addMemberBtn')?.addEventListener('click', () => this.handleAddMember());
    
    // Notification settings
    document.getElementById('enableNotifBtn')?.addEventListener('click', () => this.requestNotificationPermission());
    document.getElementById('reminderMemberList')?.addEventListener('change', (e) => {
      if (e.target.classList.contains('reminder-time-input')) {
        this.saveReminderSettings(e.target.dataset.member, 'time', e.target.value);
      }
      if (e.target.classList.contains('reminder-toggle-input')) {
        this.saveReminderSettings(e.target.dataset.member, 'enabled', e.target.checked);
      }
    });
    
    // Avatar Modal
    document.getElementById('avatarModalClose')?.addEventListener('click', () => this.closeAvatarModal());
    document.getElementById('avatarFileInput')?.addEventListener('change', (e) => this.handleAvatarFileSelect(e));
    document.getElementById('avatarPresetsGrid')?.addEventListener('click', (e) => this.handlePresetSelect(e));
    document.getElementById('confirmAvatarBtn')?.addEventListener('click', () => this.confirmAvatarSelection());
    
    // Escape key
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        this.closeModal();
        this.closeLoanModal();
        this.closeAvatarModal();
        this.closeKPIModal();
      }
    });
  }
};

document.addEventListener('DOMContentLoaded', () => App.init());
