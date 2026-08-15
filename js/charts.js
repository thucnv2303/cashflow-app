const Charts = {
  instances: {},  // Store chart instances for cleanup
  
  // Set Chart.js defaults for dark/light theme
  setTheme(isDark) {
    if (typeof Chart === 'undefined') return;
    Chart.defaults.color = isDark ? 'rgba(255, 255, 255, 0.6)' : '#636e72';
    Chart.defaults.font.family = 'Inter, sans-serif';
    Chart.defaults.plugins.tooltip.backgroundColor = isDark ? 'rgba(0, 0, 0, 0.8)' : 'rgba(255, 255, 255, 0.9)';
    Chart.defaults.plugins.tooltip.titleColor = isDark ? '#fff' : '#2d3436';
    Chart.defaults.plugins.tooltip.bodyColor = isDark ? '#fff' : '#2d3436';
    Chart.defaults.plugins.tooltip.borderColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)';
    Chart.defaults.plugins.tooltip.borderWidth = 1;
    Chart.defaults.plugins.tooltip.cornerRadius = 8;
  },
  
  // Destroy a chart instance
  destroy(id) {
    if (this.instances[id]) {
      this.instances[id].destroy();
      delete this.instances[id];
    }
  },
  
  // Destroy all charts
  destroyAll() {
    Object.keys(this.instances).forEach(id => this.destroy(id));
  },
  
  // 1. Category Doughnut Chart
  renderCategoryChart(canvasId, transactions) {
    this.destroy(canvasId);
    const canvas = document.getElementById(canvasId);
    if (!canvas || typeof Chart === 'undefined') return;
    
    const expenses = transactions.filter(t => t.type === 'expense');
    
    if (expenses.length === 0) {
      // Empty state can be handled by HTML overlay
      return;
    }
    
    const categoryTotals = {};
    expenses.forEach(t => {
      categoryTotals[t.category] = (categoryTotals[t.category] || 0) + t.amount;
    });
    
    const labels = [];
    const data = [];
    Object.entries(categoryTotals)
      .sort(([,a], [,b]) => b - a)
      .forEach(([catId, amount]) => {
        const cat = getCategoryById(catId);
        labels.push(cat ? cat.label : catId);
        data.push(amount);
      });
      
    const totalAmount = data.reduce((a,b) => a+b, 0);
      
    this.instances[canvasId] = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: data,
          backgroundColor: CHART_COLORS,
          borderWidth: 0,
          hoverOffset: 4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '70%',
        spacing: 3,
        plugins: {
          legend: {
            position: 'right',
            labels: { boxWidth: 12, usePointStyle: true, padding: 15 }
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                const value = context.raw;
                const percentage = ((value / totalAmount) * 100).toFixed(1);
                return ` ${context.label}: ${formatCurrency(value)} (${percentage}%)`;
              }
            }
          }
        }
      }
    });
  },
  
  // 2. Monthly Income vs Expense Bar Chart (last 6 months)
  renderMonthlyChart(canvasId, monthsData) {
    this.destroy(canvasId);
    const canvas = document.getElementById(canvasId);
    if (!canvas || typeof Chart === 'undefined') return;
    
    const labels = monthsData.map(d => `Th${d.month}`);
    const incomeData = monthsData.map(d => 
      d.transactions.filter(t => t.type === 'income').reduce((sum, t) => sum + t.amount, 0)
    );
    const expenseData = monthsData.map(d => 
      d.transactions.filter(t => t.type === 'expense').reduce((sum, t) => sum + t.amount, 0)
    );
    
    this.instances[canvasId] = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          {
            label: 'Thu nhập',
            data: incomeData,
            backgroundColor: '#00b894',
            borderRadius: 4
          },
          {
            label: 'Chi tiêu',
            data: expenseData,
            backgroundColor: '#ff7675',
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          tooltip: {
            callbacks: {
              label: (context) => ` ${context.dataset.label}: ${formatCurrency(context.raw)}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            grid: { color: 'rgba(150, 150, 150, 0.1)' },
            ticks: {
              callback: (value) => value === 0 ? '0' : formatCurrency(value).replace(' ₫', '')
            }
          },
          x: {
            grid: { display: false }
          }
        }
      }
    });
  },
  
  // 3. Cashflow Trend Line Chart
  renderTrendChart(canvasId, monthsData) {
    this.destroy(canvasId);
    const canvas = document.getElementById(canvasId);
    if (!canvas || typeof Chart === 'undefined') return;
    
    const labels = monthsData.map(d => `Th${d.month}`);
    const netData = monthsData.map(d => {
      const income = d.transactions.filter(t => t.type === 'income').reduce((sum, t) => sum + t.amount, 0);
      const expense = d.transactions.filter(t => t.type === 'expense').reduce((sum, t) => sum + t.amount, 0);
      return income - expense;
    });
    
    const ctx = canvas.getContext('2d');
    const gradient = ctx.createLinearGradient(0, 0, 0, 300);
    gradient.addColorStop(0, 'rgba(108, 92, 231, 0.4)');
    gradient.addColorStop(1, 'rgba(108, 92, 231, 0)');
    
    this.instances[canvasId] = new Chart(canvas, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [{
          label: 'Dòng tiền thuần',
          data: netData,
          borderColor: '#6c5ce7',
          backgroundColor: gradient,
          fill: true,
          tension: 0.4,
          pointRadius: 0,
          pointHoverRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          tooltip: {
            callbacks: {
              label: (context) => ` ${context.dataset.label}: ${formatCurrency(context.raw)}`
            }
          }
        },
        scales: {
          y: {
            grid: { color: 'rgba(150, 150, 150, 0.1)' },
            ticks: {
              callback: (value) => value === 0 ? '0' : formatCurrency(value).replace(' ₫', '')
            }
          },
          x: {
            grid: { display: false }
          }
        }
      }
    });
  },
  
  // 4. Category Trend Chart (for analytics page)
  renderCategoryTrendChart(canvasId, monthsData) {
    this.destroy(canvasId);
    const canvas = document.getElementById(canvasId);
    if (!canvas || typeof Chart === 'undefined') return;
    
    // Find top 5 expense categories overall
    const allExpenses = monthsData.flatMap(d => d.transactions.filter(t => t.type === 'expense'));
    const catTotals = {};
    allExpenses.forEach(t => {
      catTotals[t.category] = (catTotals[t.category] || 0) + t.amount;
    });
    
    const topCategories = Object.entries(catTotals)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(entry => entry[0]);
      
    const labels = monthsData.map(d => `Tháng ${d.month}`);
    
    const datasets = topCategories.map((catId, index) => {
      const data = monthsData.map(d => {
        return d.transactions
          .filter(t => t.type === 'expense' && t.category === catId)
          .reduce((sum, t) => sum + t.amount, 0);
      });
      
      const catInfo = getCategoryById(catId);
      
      return {
        label: catInfo ? catInfo.label : catId,
        data: data,
        borderColor: CHART_COLORS[index % CHART_COLORS.length],
        backgroundColor: CHART_COLORS[index % CHART_COLORS.length],
        tension: 0.4,
        fill: false,
        pointRadius: 0,
        pointHoverRadius: 5
      };
    });
    
    this.instances[canvasId] = new Chart(canvas, {
      type: 'line',
      data: {
        labels: labels,
        datasets: datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: 'index',
          intersect: false,
        },
        plugins: {
          tooltip: {
            callbacks: {
              label: (context) => ` ${context.dataset.label}: ${formatCurrency(context.raw)}`
            }
          }
        },
        scales: {
          y: {
            stacked: false,
            grid: { color: 'rgba(150, 150, 150, 0.1)' },
            ticks: {
               callback: (value) => value === 0 ? '0' : formatCurrency(value).replace(' ₫', '')
            }
          },
          x: {
            grid: { display: false }
          }
        }
      }
    });
  }
};
