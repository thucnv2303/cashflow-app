// Categories definition
const CATEGORIES = {
  expense: [
    { id: 'food', emoji: '🍜', label: 'Ăn uống' },
    { id: 'transport', emoji: '🚗', label: 'Di chuyển' },
    { id: 'shopping', emoji: '🛒', label: 'Mua sắm' },
    { id: 'entertainment', emoji: '🎮', label: 'Giải trí' },
    { id: 'health', emoji: '💊', label: 'Y tế' },
    { id: 'education', emoji: '📚', label: 'Học tập' },
    { id: 'bills', emoji: '🧾', label: 'Hóa đơn' },
    { id: 'other', emoji: '📦', label: 'Khác' }
  ],
  income: [
    { id: 'salary', emoji: '💰', label: 'Lương' },
    { id: 'freelance', emoji: '💻', label: 'Freelance' },
    { id: 'investment', emoji: '📈', label: 'Đầu tư' },
    { id: 'gift', emoji: '🎁', label: 'Quà tặng' },
    { id: 'other_income', emoji: '💵', label: 'Khác' }
  ]
};

// Chart color palette
const CHART_COLORS = [
  '#6c5ce7', '#00b894', '#e17055', '#fdcb6e', 
  '#0984e3', '#e84393', '#00cec9', '#fab1a0',
  '#a29bfe', '#55efc4', '#ff7675', '#ffeaa7'
];

// Avatar System - Cute illustrated avatars & photo upload support
const AVATAR_STYLE = 'adventurer'; // cute illustrated style
function getAvatarUrl(seed) {
  return `https://api.dicebear.com/9.x/${AVATAR_STYLE}/svg?seed=${encodeURIComponent(seed)}&backgroundColor=transparent`;
}

const AVATARS = [
  { id: 'avatar_dad', emoji: '👨', label: 'Bố', img: getAvatarUrl('dad-family') },
  { id: 'avatar_mom', emoji: '👩', label: 'Mẹ', img: getAvatarUrl('mom-family') },
  { id: 'avatar_son', emoji: '👦', label: 'Con trai', img: getAvatarUrl('boy-family') },
  { id: 'avatar_daughter', emoji: '👧', label: 'Con gái', img: getAvatarUrl('girl-family') },
  { id: 'avatar_grandpa', emoji: '👴', label: 'Ông', img: getAvatarUrl('grandpa-family') },
  { id: 'avatar_grandma', emoji: '👵', label: 'Bà', img: getAvatarUrl('grandma-family') },
  { id: 'avatar_baby_boy', emoji: '👶', label: 'Bé trai', img: getAvatarUrl('baby-boy-fun') },
  { id: 'avatar_baby_girl', emoji: '👧', label: 'Bé gái', img: getAvatarUrl('baby-girl-fun') },
  { id: 'avatar_young_man', emoji: '🧑', label: 'Anh', img: getAvatarUrl('young-man-cool') },
  { id: 'avatar_young_woman', emoji: '👱‍♀️', label: 'Chị', img: getAvatarUrl('young-woman-style') },
  { id: 'avatar_cat', emoji: '🐱', label: 'Mèo cưng', img: 'https://api.dicebear.com/9.x/bottts/svg?seed=cat-cute&backgroundColor=transparent' },
  { id: 'avatar_dog', emoji: '🐶', label: 'Cún cưng', img: 'https://api.dicebear.com/9.x/bottts/svg?seed=dog-cute&backgroundColor=transparent' }
];

const MEMBER_COLORS = ['#e77d3e', '#e8658b', '#d9983a', '#2d9d6f', '#5b8def', '#9333ea', '#06b6d4', '#d94f4f'];

// Client-side image resize and square crop to lightweight Base64 string (~2KB)
function processImageFile(file, maxSize = 80) {
  return new Promise((resolve, reject) => {
    if (!file || !file.type.startsWith('image/')) {
      reject(new Error('Vui lòng chọn file hình ảnh hợp lệ'));
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement('canvas');
        canvas.width = maxSize;
        canvas.height = maxSize;
        const ctx = canvas.getContext('2d');

        // Center crop to square
        const minDim = Math.min(img.width, img.height);
        const startX = (img.width - minDim) / 2;
        const startY = (img.height - minDim) / 2;

        ctx.drawImage(img, startX, startY, minDim, minDim, 0, 0, maxSize, maxSize);
        const dataUrl = canvas.toDataURL('image/jpeg', 0.7);
        resolve(dataUrl);
      };
      img.onerror = () => reject(new Error('Không thể đọc file ảnh'));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error('Lỗi đọc file'));
    reader.readAsDataURL(file);
  });
}

// Thousand separator formatting for input fields (e.g. 5000 -> 5.000)
function formatNumberInput(value) {
  if (value === null || value === undefined || value === '') return '';
  const cleanStr = value.toString().replace(/\D/g, '');
  if (!cleanStr) return '';
  const num = parseInt(cleanStr, 10);
  return isNaN(num) ? '' : new Intl.NumberFormat('vi-VN').format(num);
}

function parseNumberInput(value) {
  if (!value) return 0;
  const cleanStr = value.toString().replace(/\D/g, '');
  if (!cleanStr) return 0;
  const num = parseInt(cleanStr, 10);
  return isNaN(num) ? 0 : num;
}

// Loan Emojis
const LOAN_TYPES = [
  { id: 'house', emoji: '🏠', label: 'Vay mua nhà' },
  { id: 'car', emoji: '🚗', label: 'Trả góp xe' },
  { id: 'education', emoji: '🎓', label: 'Vay học phí' },
  { id: 'personal', emoji: '💰', label: 'Vay cá nhân' },
  { id: 'business', emoji: '💼', label: 'Vay kinh doanh' },
  { id: 'credit_card', emoji: '💳', label: 'Thẻ tín dụng' },
  { id: 'other_loan', emoji: '📄', label: 'Khoản vay khác' },
];

// Default Category Budgets (VND per month)
const DEFAULT_BUDGETS = {
  food: 6000000,
  shopping: 2000000,
  transport: 1000000,
  bills: 2500000,
  house: 3000000,
  health: 1000000,
  entertainment: 1500000,
  education: 1500000,
  other_expense: 1000000
};

// Calculate remaining days in viewed month
function getDaysRemainingInMonth(year, month) {
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;
  const currentDay = now.getDate();
  const daysInViewedMonth = new Date(year, month, 0).getDate();

  if (year < currentYear || (year === currentYear && month < currentMonth)) {
    return 0; // Tháng đã qua
  }
  if (year > currentYear || (year === currentYear && month > currentMonth)) {
    return daysInViewedMonth; // Tháng tương lai
  }
  // Tháng hiện tại
  return Math.max(1, daysInViewedMonth - currentDay + 1);
}

// Calculate budget health, percentage, remaining, and daily burn rate allowance
function calcCategoryBudgetInfo(spent, budget, daysRemaining) {
  const safeBudget = budget > 0 ? budget : 0;
  const percent = safeBudget > 0 ? Math.round((spent / safeBudget) * 100) : 0;
  const remaining = Math.max(0, safeBudget - spent);
  const isOver = spent > safeBudget && safeBudget > 0;
  const overAmount = isOver ? spent - safeBudget : 0;
  const dailyAllowance = (daysRemaining > 0 && remaining > 0) ? Math.floor(remaining / daysRemaining) : 0;

  let status = 'safe'; // green (< 75%)
  if (percent >= 100 || isOver) status = 'over'; // red (>= 100%)
  else if (percent >= 75) status = 'warning'; // yellow (75% - 99%)

  return {
    spent,
    budget: safeBudget,
    percent,
    remaining,
    isOver,
    overAmount,
    dailyAllowance,
    daysRemaining,
    status
  };
}

// Format number as VND currency
function formatCurrency(amount) {
  return new Intl.NumberFormat('vi-VN', {
    style: 'currency',
    currency: 'VND',
    maximumFractionDigits: 0
  }).format(amount);
}

// Format date to Vietnamese format (DD/MM/YYYY)
function formatDate(dateStr) {
  const date = new Date(dateStr);
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

// Format month label: "Tháng 8, 2026"
function formatMonthLabel(year, month) {
  return `Tháng ${month}, ${year}`;
}

// Get current year/month
function getCurrentMonth() {
  const now = new Date();
  return {
    year: now.getFullYear(),
    month: now.getMonth() + 1
  };
}

// Calculate % change between two values
function calcChange(current, previous) {
  if (previous === 0) return current > 0 ? 100 : 0;
  return ((current - previous) / previous) * 100;
}

// Generate unique ID
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// Preset Emojis for Custom Categories
const CUSTOM_CATEGORY_PRESET_EMOJIS = ['📈', '🏖️', '🐶', '🎁', '🏦', '🚗', '💍', '🎮', '👗', '🍼', '💻', '🏋️', '☕', '👶', '🍕', '🏠', '✈️', '🎨', '🛡️', '📦'];

// Get all expense categories including custom
function getExpenseCategories() {
  const custom = (typeof Storage !== 'undefined' && Storage.getCustomCategories) ? Storage.getCustomCategories() : [];
  return [
    ...CATEGORIES.expense,
    ...custom.filter(c => !CATEGORIES.expense.some(def => def.id === c.id))
  ];
}

// Get all categories flat (expense + income + custom)
function getAllCategories() {
  return [...getExpenseCategories(), ...CATEGORIES.income];
}

// Get category by id (with fallback)
function getCategoryById(id) {
  if (!id) return null;
  const all = getAllCategories();
  const found = all.find(c => c.id === id);
  if (found) return found;
  return { id, emoji: '📦', label: id };
}

// Get month range (start date, end date) for filtering
function getMonthRange(year, month) {
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 0);
  
  const formatDateString = (d) => {
    const y = d.getFullYear();
    const m = (d.getMonth() + 1).toString().padStart(2, '0');
    const day = d.getDate().toString().padStart(2, '0');
    return `${y}-${m}-${day}`;
  };
  
  return {
    start: formatDateString(start),
    end: formatDateString(end)
  };
}

// Navigate month (add/subtract months)
function navigateMonth(year, month, delta) {
  let newMonth = month + delta;
  let newYear = year;
  
  if (newMonth > 12) {
    newMonth = 1;
    newYear += 1;
  } else if (newMonth < 1) {
    newMonth = 12;
    newYear -= 1;
  }
  
  return { year: newYear, month: newMonth };
}

// Mood Analysis
function analyzeMood(savingsRate, expenseChange) {
  if (savingsRate >= 30) return {
    emoji: '🤩', label: 'Tuyệt vời',
    message: 'Tuyệt vời! Gia đình đang tiết kiệm rất tốt!',
    level: 'excellent', color: '#10b981', bgColor: 'rgba(16, 185, 129, 0.15)'
  };
  if (savingsRate >= 20) return {
    emoji: '😊', label: 'Tốt',
    message: 'Tốt lắm! Duy trì thói quen này nhé!',
    level: 'good', color: '#10b981', bgColor: 'rgba(16, 185, 129, 0.1)'
  };
  if (savingsRate >= 10) return {
    emoji: '😐', label: 'Tạm ổn',
    message: 'Ổn thôi, cố gắng cắt giảm chi tiêu thêm nhé.',
    level: 'ok', color: '#f59e0b', bgColor: 'rgba(245, 158, 11, 0.1)'
  };
  if (savingsRate >= 0) return {
    emoji: '😰', label: 'Cẩn thận',
    message: 'Cẩn thận! Chi tiêu gần bằng thu nhập rồi đấy.',
    level: 'warning', color: '#ef4444', bgColor: 'rgba(239, 68, 68, 0.1)'
  };
  return {
    emoji: '🔥', label: 'Báo động',
    message: 'Báo động! Đang chi nhiều hơn thu, cần điều chỉnh gấp!',
    level: 'danger', color: '#ef4444', bgColor: 'rgba(239, 68, 68, 0.15)'
  };
}

// Smart Tips Generator
function generateSmartTips(currentTransactions, previousTransactions) {
  const tips = [];
  
  const currExpense = currentTransactions.filter(t => t.type === 'expense').reduce((sum, t) => sum + t.amount, 0);
  const prevExpense = previousTransactions.filter(t => t.type === 'expense').reduce((sum, t) => sum + t.amount, 0);
  const currIncome = currentTransactions.filter(t => t.type === 'income').reduce((sum, t) => sum + t.amount, 0);

  // Positive: Total expense decreased
  if (prevExpense > 0 && currExpense < prevExpense) {
    tips.push({
      emoji: '🎉',
      text: 'Tháng này bạn đã chi tiêu ít hơn tháng trước. Tuyệt vời!',
      type: 'positive'
    });
  }

  // Info: No income recorded yet
  if (currIncome === 0) {
    tips.push({
      emoji: '💡',
      text: 'Chưa có khoản thu nhập nào được ghi nhận trong tháng này. Hãy nhớ cập nhật nhé!',
      type: 'info'
    });
  }

  // Warning: Category increased > 30%
  const currCatTotals = {};
  currentTransactions.filter(t => t.type === 'expense').forEach(t => {
    currCatTotals[t.category] = (currCatTotals[t.category] || 0) + t.amount;
  });
  
  const prevCatTotals = {};
  previousTransactions.filter(t => t.type === 'expense').forEach(t => {
    prevCatTotals[t.category] = (prevCatTotals[t.category] || 0) + t.amount;
  });

  for (const catId in currCatTotals) {
    const prev = prevCatTotals[catId] || 0;
    const curr = currCatTotals[catId];
    if (prev > 0) {
      const inc = ((curr - prev) / prev) * 100;
      if (inc > 30) {
        const cat = getCategoryById(catId);
        tips.push({
          emoji: '⚠️',
          text: `Chi tiêu cho ${cat ? cat.label : 'danh mục này'} đã tăng ${Math.round(inc)}% so với tháng trước!`,
          type: 'warning'
        });
      }
    }
  }

  // Warning: Large transaction
  const largeTransactions = currentTransactions.filter(t => t.type === 'expense' && t.amount > currIncome * 0.3 && currIncome > 0);
  if (largeTransactions.length > 0) {
    tips.push({
      emoji: '🚨',
      text: 'Có khoản chi tiêu rất lớn (hơn 30% thu nhập). Hãy cân nhắc kỹ trước khi chi.',
      type: 'warning'
    });
  }

  // Danger/Warning: Savings status
  if (currExpense > currIncome && currIncome > 0) {
    tips.push({
      emoji: '📉',
      text: 'Nguy hiểm! Bạn đang tiêu nhiều hơn số tiền kiếm được.',
      type: 'danger'
    });
  } else if (currIncome > 0 && ((currIncome - currExpense)/currIncome) < 0.1) {
    tips.push({
      emoji: '🐷',
      text: 'Bạn đang tiết kiệm được rất ít. Cố gắng giữ lại ít nhất 10% thu nhập nhé.',
      type: 'warning'
    });
  }

  // Info: Fallback if no tips
  if (tips.length === 0) {
    tips.push({
      emoji: '📝',
      text: 'Ghi chép chi tiêu mỗi ngày giúp bạn quản lý tài chính tốt hơn.',
      type: 'info'
    });
  }

  return tips;
}

// Loan Calculations
function calculateLoanStatus(loan) {
  const { principal, interestRate, termMonths, monthlyPayment, startDate } = loan;
  
  const start = new Date(startDate);
  const now = new Date();
  
  let monthsPaid = (now.getFullYear() - start.getFullYear()) * 12 + (now.getMonth() - start.getMonth());
  if (now.getDate() < start.getDate()) {
    monthsPaid--;
  }
  monthsPaid = Math.max(0, monthsPaid);
  monthsPaid = Math.min(monthsPaid, termMonths); 
  
  const monthlyRate = (interestRate / 100) / 12;
  
  let currentBalance = principal;
  let totalInterestPaid = 0;
  
  for (let i = 0; i < monthsPaid; i++) {
    const interest = currentBalance * monthlyRate;
    const principalPaid = monthlyPayment - interest;
    totalInterestPaid += interest;
    currentBalance -= principalPaid;
    if (currentBalance < 0) currentBalance = 0;
  }
  
  const currentMonthlyInterest = currentBalance * monthlyRate;
  const currentMonthlyPrincipal = monthlyPayment - currentMonthlyInterest;
  
  const estimatedEndDate = new Date(start);
  estimatedEndDate.setMonth(start.getMonth() + termMonths);
  
  const progressPercent = termMonths > 0 ? (monthsPaid / termMonths) * 100 : 100;
  
  return {
    monthsPaid,
    totalPaid: monthsPaid * monthlyPayment,
    remainingBalance: Math.max(0, currentBalance),
    totalInterestPaid,
    progressPercent: Math.min(100, Math.max(0, progressPercent)),
    estimatedEndDate,
    monthlyInterest: Math.max(0, currentMonthlyInterest),
    monthlyPrincipal: Math.max(0, currentMonthlyPrincipal)
  };
}
