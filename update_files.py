import os

# Update index.html
html_path = "index.html"
with open(html_path, "r", encoding="utf-8") as f:
    html = f.read()

# Setup Modal
html = html.replace('<div class="app">', '<div class="app">\n    <!-- Setup Modal -->\n    <div class="modal-overlay" id="setupModal">\n      <div class="modal setup-modal">\n        <div class="setup-header">\n          <span class="setup-emoji">👨‍👩‍👧‍👦</span>\n          <h2>Chào mừng đến CashFlow!</h2>\n          <p>Hãy thiết lập gia đình của bạn</p>\n        </div>\n        <form id="setupForm">\n          <div class="form-group">\n            <label for="familyName">Tên gia đình</label>\n            <input type="text" id="familyName" placeholder="VD: Gia đình Thức" required>\n          </div>\n          <div class="setup-members" id="setupMembers">\n            <h3>Thành viên</h3>\n            <div class="setup-member-list" id="setupMemberList">\n              <!-- JS renders member inputs here -->\n            </div>\n            <button type="button" class="btn-ghost" id="addSetupMember">+ Thêm thành viên</button>\n          </div>\n          <button type="submit" class="btn-primary btn-submit">🚀 Bắt đầu sử dụng</button>\n        </form>\n      </div>\n    </div>')

# Navigation
nav_old = """      <ul class="nav-menu">
        <li class="nav-item active" data-page="dashboard">
          <span class="nav-icon">📊</span>
          <span class="nav-label">Tổng quan</span>
        </li>
        <li class="nav-item" data-page="transactions">
          <span class="nav-icon">📋</span>
          <span class="nav-label">Giao dịch</span>
        </li>
        <li class="nav-item" data-page="analytics">
          <span class="nav-icon">📈</span>
          <span class="nav-label">Phân tích</span>
        </li>
        <li class="nav-item" data-page="settings">
          <span class="nav-icon">⚙️</span>
          <span class="nav-label">Cài đặt</span>
        </li>
      </ul>"""
nav_new = """      <ul class="nav-menu">
        <li class="nav-item active" data-page="dashboard">
          <span class="nav-icon">📊</span>
          <span class="nav-label">Tổng quan</span>
        </li>
        <li class="nav-item" data-page="transactions">
          <span class="nav-icon">📋</span>
          <span class="nav-label">Giao dịch</span>
        </li>
        <li class="nav-item" data-page="loans">
          <span class="nav-icon">💳</span>
          <span class="nav-label">Khoản vay</span>
        </li>
        <li class="nav-item" data-page="analytics">
          <span class="nav-icon">📈</span>
          <span class="nav-label">Phân tích</span>
        </li>
        <li class="nav-item" data-page="settings">
          <span class="nav-icon">⚙️</span>
          <span class="nav-label">Cài đặt</span>
        </li>
      </ul>"""
html = html.replace(nav_old, nav_new)

# Dashboard Header
dash_header_old = """        <div class="page-header">
          <h1>Tổng quan</h1>
          <div class="month-selector">
            <button class="btn-ghost" id="prevMonth">◀</button>
            <span id="currentMonth">Tháng 8, 2026</span>
            <button class="btn-ghost" id="nextMonth">▶</button>
          </div>
        </div>"""
dash_header_new = """        <div class="page-header">
          <div class="family-greeting">
            <h1 id="greetingText">Xin chào, Gia đình! 👋</h1>
            <div class="family-avatars" id="familyAvatars">
              <!-- JS renders member avatars -->
            </div>
          </div>
          <div class="month-selector">
            <button class="btn-ghost" id="prevMonth">◀</button>
            <span id="currentMonth">Tháng 8, 2026</span>
            <button class="btn-ghost" id="nextMonth">▶</button>
          </div>
        </div>
        
        <!-- Mood Card -->
        <div class="mood-card" id="moodCard">
          <div class="mood-emoji" id="moodEmoji">😊</div>
          <div class="mood-info">
            <div class="mood-label" id="moodLabel">Tốt</div>
            <div class="mood-message" id="moodMessage">Tốt lắm! Duy trì nhé!</div>
          </div>
        </div>

        <!-- Smart Tips -->
        <div class="tips-container" id="tipsContainer" style="display:none;">
          <!-- JS renders tips -->
        </div>"""
html = html.replace(dash_header_old, dash_header_new)

# Loans Page
loans_page = """
      <!-- PAGE 5: Loans -->
      <section class="page" id="page-loans">
        <div class="page-header">
          <h1>Khoản vay</h1>
          <button class="btn-primary" id="addLoanBtn">+ Thêm khoản vay</button>
        </div>
        
        <!-- Loan Summary Cards -->
        <div class="loan-summary-grid">
          <div class="summary-card">
            <div class="card-icon">💰</div>
            <div class="card-info">
              <span class="card-label">Tổng dư nợ</span>
              <span class="card-value" id="totalDebt">0 ₫</span>
            </div>
          </div>
          <div class="summary-card">
            <div class="card-icon">💸</div>
            <div class="card-info">
              <span class="card-label">Trả hàng tháng</span>
              <span class="card-value" id="monthlyPaymentTotal">0 ₫</span>
            </div>
          </div>
          <div class="summary-card">
            <div class="card-icon">📊</div>
            <div class="card-info">
              <span class="card-label">Lãi đã trả</span>
              <span class="card-value" id="totalInterestPaid">0 ₫</span>
            </div>
          </div>
        </div>
        
        <!-- Loan List -->
        <div class="loans-list" id="loansList">
          <div class="empty-state">
            <span class="empty-icon">💳</span>
            <p>Chưa có khoản vay nào</p>
          </div>
        </div>
      </section>
"""
html = html.replace('      <!-- PAGE 3: Analytics -->', loans_page + '\n      <!-- PAGE 3: Analytics -->')

# Transaction Modal Member Selector
member_selector = """
          <!-- Member Selector -->
          <div class="form-group">
            <label>Ai chi?</label>
            <div class="member-selector" id="memberSelector">
              <!-- JS renders member buttons -->
            </div>
          </div>
"""
html = html.replace('          <!-- Amount -->', member_selector + '\n          <!-- Amount -->')

# Loan Modal
loan_modal = """
    <!-- Loan Modal -->
    <div class="modal-overlay" id="loanModal">
      <div class="modal">
        <div class="modal-header">
          <h2 id="loanModalTitle">Thêm khoản vay</h2>
          <button class="btn-close" id="loanModalClose">✕</button>
        </div>
        <form id="loanForm">
          <input type="hidden" id="editLoanId" value="">
          <div class="form-group">
            <label>Loại khoản vay</label>
            <div class="loan-type-grid" id="loanTypeGrid">
              <!-- JS renders loan type buttons -->
            </div>
          </div>
          <div class="form-group">
            <label for="loanName">Tên khoản vay</label>
            <input type="text" id="loanName" placeholder="VD: Vay mua nhà" required>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label for="loanPrincipal">Số tiền vay (₫)</label>
              <input type="number" id="loanPrincipal" placeholder="0" min="0" required>
            </div>
            <div class="form-group">
              <label for="loanRate">Lãi suất (%/năm)</label>
              <input type="number" id="loanRate" placeholder="8" min="0" step="0.1" required>
            </div>
          </div>
          <div class="form-row">
            <div class="form-group">
              <label for="loanTerm">Kỳ hạn (tháng)</label>
              <input type="number" id="loanTerm" placeholder="240" min="1" required>
            </div>
            <div class="form-group">
              <label for="loanMonthly">Trả hàng tháng (₫)</label>
              <input type="number" id="loanMonthly" placeholder="0" min="0" required>
            </div>
          </div>
          <div class="form-group">
            <label for="loanStartDate">Ngày bắt đầu</label>
            <input type="date" id="loanStartDate" required>
          </div>
          <div class="form-group">
            <label for="loanNote">Ghi chú</label>
            <input type="text" id="loanNote" placeholder="VD: Ngân hàng ABC">
          </div>
          <button type="submit" class="btn-primary btn-submit">Lưu khoản vay</button>
        </form>
      </div>
    </div>
"""
html = html.replace('    <!-- Toast -->', loan_modal + '\n    <!-- Toast -->')

with open(html_path, "w", encoding="utf-8") as f:
    f.write(html)


# Update css/style.css
css_path = "css/style.css"
with open(css_path, "r", encoding="utf-8") as f:
    css = f.read()

old_root = """:root {
  /* Dark Mode DEFAULT */
  --bg-primary: #0a0e1a;
  --bg-secondary: #111827;
  --surface: rgba(255, 255, 255, 0.05);
  --surface-hover: rgba(255, 255, 255, 0.08);
  --border: rgba(255, 255, 255, 0.1);
  --border-hover: rgba(255, 255, 255, 0.2);
  --text-primary: #ffffff;
  --text-secondary: rgba(255, 255, 255, 0.6);
  --text-muted: rgba(255, 255, 255, 0.35);
  --primary: #6c5ce7;
  --primary-glow: rgba(108, 92, 231, 0.3);
  --income: #00b894;
  --income-bg: rgba(0, 184, 148, 0.15);
  --expense: #e17055;
  --expense-bg: rgba(225, 112, 85, 0.15);
  --danger: #ff6b6b;
  --warning: #feca57;
  --shadow: rgba(0, 0, 0, 0.3);
}

html[data-theme='light'] {
  --bg-primary: #f0f2f5;
  --bg-secondary: #ffffff;
  --surface: rgba(255, 255, 255, 0.8);
  --surface-hover: rgba(255, 255, 255, 0.95);
  --border: rgba(0, 0, 0, 0.08);
  --border-hover: rgba(0, 0, 0, 0.15);
  --text-primary: #1a1a2e;
  --text-secondary: #636e72;
  --text-muted: #b2bec3;
  --primary: #6c5ce7;
  --primary-glow: rgba(108, 92, 231, 0.2);
  --income: #00b894;
  --income-bg: rgba(0, 184, 148, 0.1);
  --expense: #e17055;
  --expense-bg: rgba(225, 112, 85, 0.1);
  --shadow: rgba(0, 0, 0, 0.08);
}"""

new_root = """:root {
  --bg-primary: #0f0b1e;
  --bg-secondary: #1a1333;
  --surface: rgba(255, 255, 255, 0.06);
  --surface-hover: rgba(255, 255, 255, 0.1);
  --border: rgba(255, 255, 255, 0.1);
  --border-hover: rgba(255, 255, 255, 0.2);
  --text-primary: #ffffff;
  --text-secondary: rgba(255, 255, 255, 0.65);
  --text-muted: rgba(255, 255, 255, 0.35);
  --primary: #8b5cf6;
  --primary-glow: rgba(139, 92, 246, 0.3);
  --accent-pink: #ec4899;
  --accent-amber: #f59e0b;
  --income: #10b981;
  --income-bg: rgba(16, 185, 129, 0.15);
  --expense: #ef4444;
  --expense-bg: rgba(239, 68, 68, 0.15);
  --danger: #ef4444;
  --warning: #f59e0b;
  --shadow: rgba(0, 0, 0, 0.3);
  --family-gradient: linear-gradient(135deg, #8b5cf6, #ec4899, #f59e0b);
}

html[data-theme='light'] {
  --bg-primary: #faf5ff;
  --bg-secondary: #ffffff;
  --surface: rgba(255, 255, 255, 0.8);
  --surface-hover: rgba(255, 255, 255, 0.95);
  --border: rgba(0, 0, 0, 0.08);
  --border-hover: rgba(0, 0, 0, 0.15);
  --text-primary: #1e1b4b;
  --text-secondary: #4f46e5;
  --text-muted: #9ca3af;
  --primary: #8b5cf6;
  --primary-glow: rgba(139, 92, 246, 0.2);
  --accent-pink: #ec4899;
  --accent-amber: #f59e0b;
  --income: #10b981;
  --income-bg: rgba(16, 185, 129, 0.1);
  --expense: #ef4444;
  --expense-bg: rgba(239, 68, 68, 0.1);
  --danger: #ef4444;
  --warning: #f59e0b;
  --shadow: rgba(0, 0, 0, 0.08);
  --family-gradient: linear-gradient(135deg, #8b5cf6, #ec4899, #f59e0b);
}"""

css = css.replace(old_root, new_root)

new_css = """

/* NEW FAMILY UI STYLES */
.family-greeting h1 {
  font-size: 1.5rem;
  background: var(--family-gradient);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
}

.family-avatars {
  display: flex;
  gap: 8px;
  margin-top: 8px;
}

.family-avatar {
  width: 36px;
  height: 36px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.2rem;
  border: 2px solid;
  cursor: pointer;
  transition: transform 0.2s;
}

.family-avatar:hover { transform: scale(1.15); }
.family-avatar.active { box-shadow: 0 0 0 3px var(--primary-glow); }

.mood-card {
  display: flex;
  align-items: center;
  gap: 16px;
  padding: 20px 24px;
  border-radius: 16px;
  margin-bottom: 24px;
  border: 1px solid var(--border);
  transition: all 0.3s;
}

.mood-emoji {
  font-size: 3rem;
  animation: moodBounce 2s ease infinite;
}

@keyframes moodBounce {
  0%, 100% { transform: scale(1); }
  50% { transform: scale(1.1); }
}

.mood-label {
  font-size: 1.1rem;
  font-weight: 700;
}

.mood-message {
  color: var(--text-secondary);
  font-size: 0.95rem;
  margin-top: 4px;
}

.tips-container {
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-bottom: 24px;
}

.tip-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 12px 16px;
  border-radius: 12px;
  font-size: 0.9rem;
  background: var(--surface);
  border: 1px solid var(--border);
  animation: fadeIn 0.3s ease;
}

.tip-item.positive { border-color: rgba(16, 185, 129, 0.3); }
.tip-item.warning { border-color: rgba(239, 68, 68, 0.3); }
.tip-item.info { border-color: rgba(139, 92, 246, 0.3); }

.member-selector {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
}

.member-btn {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 4px;
  padding: 10px 14px;
  border-radius: 12px;
  background: var(--surface);
  border: 2px solid var(--border);
  cursor: pointer;
  transition: all 0.2s;
  font-size: 0.85rem;
  color: var(--text-primary);
}

.member-btn:hover { background: var(--surface-hover); }
.member-btn.active { border-color: var(--primary); background: var(--primary-glow); }

.member-btn .member-avatar {
  font-size: 1.8rem;
}

.loan-summary-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 16px;
  margin-bottom: 24px;
}

.loans-list {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.loan-card {
  background: var(--surface);
  backdrop-filter: blur(20px);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 20px;
  transition: all 0.3s;
}

.loan-card:hover {
  border-color: var(--border-hover);
  transform: translateY(-2px);
}

.loan-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 16px;
}

.loan-title {
  display: flex;
  align-items: center;
  gap: 10px;
}

.loan-title .loan-emoji {
  font-size: 1.5rem;
}

.loan-title h4 {
  margin: 0;
  font-size: 1.1rem;
}

.loan-title .loan-note {
  font-size: 0.85rem;
  color: var(--text-muted);
}

.loan-progress {
  margin: 16px 0;
}

.loan-progress-bar {
  width: 100%;
  height: 10px;
  background: var(--surface-hover);
  border-radius: 5px;
  overflow: hidden;
}

.loan-progress-fill {
  height: 100%;
  background: var(--family-gradient);
  border-radius: 5px;
  transition: width 0.8s ease;
}

.loan-progress-text {
  display: flex;
  justify-content: space-between;
  font-size: 0.85rem;
  color: var(--text-secondary);
  margin-top: 6px;
}

.loan-stats {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 12px;
  margin-top: 16px;
}

.loan-stat {
  padding: 10px;
  background: var(--surface-hover);
  border-radius: 10px;
}

.loan-stat-label {
  font-size: 0.8rem;
  color: var(--text-muted);
}

.loan-stat-value {
  font-size: 1rem;
  font-weight: 600;
  margin-top: 2px;
}

.loan-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}

.loan-type-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 8px;
}

.loan-type-btn {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 4px;
  padding: 10px;
  background: var(--surface);
  border: 2px solid var(--border);
  border-radius: 10px;
  cursor: pointer;
  font-size: 0.8rem;
  color: var(--text-primary);
  transition: all 0.2s;
}

.loan-type-btn:hover { background: var(--surface-hover); }
.loan-type-btn.active { border-color: var(--primary); background: var(--primary-glow); }
.loan-type-btn .loan-type-emoji { font-size: 1.5rem; }

.form-row {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 16px;
}

.setup-modal {
  max-width: 500px;
  width: 90%;
}

.setup-header {
  text-align: center;
  margin-bottom: 24px;
}

.setup-emoji {
  font-size: 3rem;
  display: block;
  margin-bottom: 12px;
}

.setup-header h2 {
  background: var(--family-gradient);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
}

.setup-member-list {
  display: flex;
  flex-direction: column;
  gap: 12px;
  margin: 16px 0;
}

.setup-member-row {
  display: flex;
  gap: 10px;
  align-items: center;
}

.setup-member-row .avatar-picker {
  width: 44px;
  height: 44px;
  border-radius: 50%;
  background: var(--surface);
  border: 2px solid var(--border);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.4rem;
  cursor: pointer;
  flex-shrink: 0;
  transition: all 0.2s;
}

.setup-member-row .avatar-picker:hover,
.setup-member-row .avatar-picker.selected {
  border-color: var(--primary);
}

.setup-member-row input {
  flex: 1;
}

.setup-member-row .remove-member-btn {
  background: none;
  border: none;
  color: var(--text-muted);
  cursor: pointer;
  font-size: 1.2rem;
  padding: 4px;
}

.setup-member-row .remove-member-btn:hover {
  color: var(--danger);
}

/* Avatar Picker Popover */
.avatar-popover {
  position: absolute;
  background: var(--bg-secondary);
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 8px;
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 4px;
  z-index: 100;
  box-shadow: 0 8px 30px var(--shadow);
}

.avatar-option {
  width: 40px;
  height: 40px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.3rem;
  cursor: pointer;
  border: none;
  background: none;
  transition: all 0.2s;
}

.avatar-option:hover {
  background: var(--surface-hover);
  transform: scale(1.15);
}

@media (max-width: 768px) {
  .loan-summary-grid { grid-template-columns: 1fr; }
  .form-row { grid-template-columns: 1fr; }
  .loan-type-grid { grid-template-columns: repeat(3, 1fr); }
  .loan-stats { grid-template-columns: 1fr; }
  .family-greeting h1 { font-size: 1.2rem; }
  .mood-card { flex-direction: column; text-align: center; }
  .mood-emoji { font-size: 2.5rem; }
}

"""
css += new_css

with open(css_path, "w", encoding="utf-8") as f:
    f.write(css)

print("Files updated successfully")
