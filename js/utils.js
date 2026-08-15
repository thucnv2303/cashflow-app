// Categories definition
const CATEGORIES = {
  expense: [
    { id: 'food', emoji: '🍜', label: 'Ăn uống' },
    { id: 'housing', emoji: '🏠', label: 'Nhà cửa' },
    { id: 'transport', emoji: '🚗', label: 'Di chuyển' },
    { id: 'shopping', emoji: '🛒', label: 'Mua sắm' },
    { id: 'health', emoji: '💊', label: 'Y tế' },
    { id: 'entertainment', emoji: '🎮', label: 'Giải trí' },
    { id: 'education', emoji: '📚', label: 'Học tập' },
    { id: 'other_expense', emoji: '📦', label: 'Khác' }
  ],
  income: [
    { id: 'salary', emoji: '💰', label: 'Lương' },
    { id: 'investment', emoji: '📈', label: 'Đầu tư' },
    { id: 'freelance', emoji: '💻', label: 'Freelance' },
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

// Get category by id
function getCategoryById(id) {
  for (const type in CATEGORIES) {
    const category = CATEGORIES[type].find(c => c.id === id);
    if (category) return category;
  }
  return null;
}

// Get all categories flat
function getAllCategories() {
  return [...CATEGORIES.expense, ...CATEGORIES.income];
}

// Get month range (start date, end date) for filtering
function getMonthRange(year, month) {
  const start = new Date(year, month - 1, 1);
  const end = new Date(year, month, 0); // Last day of the month
  
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
