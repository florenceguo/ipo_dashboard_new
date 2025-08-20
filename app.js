// 全局变量
let excelData = {};
let charts = {};

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    loadLatestExcelData();
});

// 自动检测并加载最新的Excel文件
async function loadLatestExcelData() {
    try {
        updateLoadingProgress('正在检测最新数据文件...');
        
        // 尝试加载不同可能的Excel文件名
        const possibleFiles = [
            'excel_data_optimized.json', // 优先加载优化后的文件
            'excel_data_summary.json',   // 备用文件
            'newest_data.json'           // 其他可能的文件名
        ];
        
        let dataLoaded = false;
        
        for (const filename of possibleFiles) {
            try {
                const response = await fetch(filename + '?t=' + new Date().getTime());
                if (response.ok) {
                    updateLoadingProgress('正在解析数据...');
                    excelData = await response.json();
                    console.log('数据加载成功:', filename);
                    dataLoaded = true;
                    break;
                }
            } catch (error) {
                console.log(`${filename} 不存在，尝试下一个`);
            }
        }
        
        if (!dataLoaded) {
            // 如果JSON文件不存在，尝试直接读取Excel文件
            await loadExcelDirectly();
        } else {
            initializeDashboard();
        }
        
    } catch (error) {
        console.error('加载数据失败:', error);
        showError('数据加载失败: ' + error.message);
    }
}

// 直接读取Excel文件
async function loadExcelDirectly() {
    updateLoadingProgress('正在查找Excel文件...');
    
    // 根据当前日期生成可能的Excel文件名
    const today = new Date();
    const dateStr = today.toISOString().split('T')[0]; // YYYY-MM-DD格式
    
    const possibleExcelFiles = [
        `新股周度统计${dateStr.replace(/-/g, '-')}.xlsx`,
        `新股周度统计${dateStr.replace(/-/g, '-')}v1.xlsx`,
        '新股周度统计2025-08-11v1.xlsx', // 您当前的文件名
        '新股周度统计2025-08-11.xlsx'
    ];
    
    // 提示用户如何更新数据
    showError(`
        <div>
            <h4>📊 数据文件检测</h4>
            <p>未找到JSON数据文件，请按以下步骤更新数据：</p>
            <ol style="text-align: left; margin: 10px 0;">
                <li>修改 <code>ipoinfo.py</code> 中的日期范围</li>
                <li>运行 <code>python ipoinfo.py</code> 生成新的Excel文件</li>
                <li>运行 <code>python analyze_excel.py</code> 生成JSON文件</li>
                <li>点击 "🔄 刷新数据" 按钮</li>
            </ol>
            <p><strong>当前查找的文件：</strong></p>
            <ul style="text-align: left; font-size: 0.9em;">
                ${possibleExcelFiles.map(file => `<li>${file}</li>`).join('')}
            </ul>
        </div>
    `);
}

// 刷新数据
async function refreshData() {
    updateLoadingProgress('正在刷新数据...');
    document.getElementById('loading').style.display = 'block';
    document.getElementById('content').style.display = 'none';
    
    // 清除缓存
    localStorage.removeItem('ipo_dashboard_data');
    localStorage.removeItem('ipo_dashboard_cache_time');
    
    // 重新加载
    await loadLatestExcelData();
}

// 初始化仪表板
function initializeDashboard() {
    console.log('🚀 initializeDashboard 被调用');
    
    // 更新数据时间
    updateDataTime();
    
    // 隐藏加载界面
    document.getElementById('loading').style.display = 'none';
    document.getElementById('content').style.display = 'block';
    
    // 初始化图表
    setTimeout(() => {
        console.log('⏰ 延时执行开始...');
        initializeCharts();
        console.log('📈 图表初始化完成，开始更新统计...');
        updateStatistics();
        console.log('🎯 仪表板初始化完全完成');
    }, 100);
}

// 更新数据时间显示
function updateDataTime() {
    const now = new Date();
    const timeStr = now.toLocaleString('zh-CN', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    });
    document.getElementById('dataUpdateTime').textContent = `数据更新时间: ${timeStr}`;
}

// 更新加载进度
function updateLoadingProgress(message) {
    const loadingElement = document.getElementById('loading');
    if (loadingElement) {
        loadingElement.innerHTML = `📊 ${message}`;
    }
}

// 显示错误信息
function showError(message) {
    document.getElementById('loading').innerHTML = `
        <div class="error">
            ${message}
        </div>
    `;
}

// 初始化所有图表
function initializeCharts() {
    try {
        createWeeklyReturnsChart();
        createBeijingExchangeChart();
        createMainBoardLotteryChart();
        createBeijingExchangeLotteryChart();
        createIssuanceChart();
        createSectorPerformanceChart();
        createPeRatioChart();
        createIpoHeatmapChart();
        updateComprehensiveStats();
    } catch (error) {
        console.error('图表初始化失败:', error);
        showError('图表初始化失败: ' + error.message);
    }
}

// 创建新股周度年化收益图表
function createWeeklyReturnsChart() {
    const ctx = document.getElementById('weeklyReturnsChart').getContext('2d');
    const data = excelData['周度收益']?.data || [];
    
    const labels = data.map(item => item.week_label);
    const returns = data.map(item => (item['年化收益贡献'] * 100).toFixed(2));
    
    charts.weeklyReturns = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: '年化收益贡献 (%)',
                data: returns,
                backgroundColor: 'rgba(52, 152, 219, 0.8)',
                borderColor: '#3498db',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { display: false },
                datalabels: {
                    display: function(context) {
                        // 只显示最后3个数据点的标签
                        return context.dataIndex >= context.dataset.data.length - 3;
                    },
                    anchor: 'end',
                    align: 'top',
                    formatter: function(value) {
                        return value + '%';
                    },
                    color: '#2c3e50',
                    font: {
                        weight: 'bold',
                        size: 12
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    }
                }
            }
        },
        plugins: [ChartDataLabels] // 启用数据标签插件
    });
}

// 创建北交所月度年化收益图表
function createBeijingExchangeChart() {
    const ctx = document.getElementById('beijingExchangeChart').getContext('2d');
    const data = excelData['北交所']?.data || [];
    
    const labels = data.map(item => {
        const date = new Date(item.listing_date);
        return `${date.getFullYear()}年${date.getMonth() + 1}月`;
    });
    const returns = data.map(item => (item['北交所年化收益贡献'] * 100).toFixed(2));
    
    charts.beijingExchange = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: '年化收益贡献 (%)',
                data: returns,
                backgroundColor: 'rgba(230, 126, 34, 0.8)',
                borderColor: '#e67e22',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { display: false },
                datalabels: {
                    display: function(context) {
                        // 只显示最后3个数据点的标签
                        return context.dataIndex >= context.dataset.data.length - 3;
                    },
                    anchor: 'end',
                    align: 'top',
                    formatter: function(value) {
                        return value + '%';
                    },
                    color: '#2c3e50',
                    font: {
                        weight: 'bold',
                        size: 12
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    }
                }
            }
        },
        plugins: [ChartDataLabels] // 启用数据标签插件
    });
}

// 其他图表创建函数...
// 主板中签率图表
function createMainBoardLotteryChart() {
    const ctx = document.getElementById('mainBoardLotteryChart').getContext('2d');
    const data = excelData['中签率统计']?.data || [];
    
    // 使用所有中签率统计数据（主要是主板数据）
    const labels = data.map(item => item.week_label);
    const lotteryA = data.map(item => (parseFloat(item.lottery_a) * 100 || 0));
    const lotteryB = data.map(item => (parseFloat(item.lottery_b) * 100 || 0));
    
    // 计算平均值
    const avgA = lotteryA.length > 0 ? lotteryA.reduce((a, b) => a + b, 0) / lotteryA.length : 0;
    const avgB = lotteryB.length > 0 ? lotteryB.reduce((a, b) => a + b, 0) / lotteryB.length : 0;
    
    // 计算A/B中签率比的均值
    const validA2BRatios = data.filter(item => item.lottery_a2b && !isNaN(parseFloat(item.lottery_a2b)));
    const avgA2B = validA2BRatios.length > 0 
        ? validA2BRatios.reduce((sum, item) => sum + parseFloat(item.lottery_a2b), 0) / validA2BRatios.length 
        : 0;
    
    document.getElementById('mainBoardAvgLotteryA').textContent = avgA.toFixed(4) + '%';
    document.getElementById('mainBoardAvgLotteryB').textContent = avgB.toFixed(4) + '%';
    document.getElementById('mainBoardAvgA2BRatio').textContent = avgA2B.toFixed(3);
    
    charts.mainBoardLottery = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: '主板A类中签率 (%)',
                    data: lotteryA,
                    borderColor: '#3498db',
                    backgroundColor: 'rgba(52, 152, 219, 0.1)',
                    fill: false,
                    tension: 0.4,
                    pointRadius: 4,
                    pointHoverRadius: 6
                },
                {
                    label: '主板B类中签率 (%)',
                    data: lotteryB,
                    borderColor: '#e74c3c',
                    backgroundColor: 'rgba(231, 76, 60, 0.1)',
                    fill: false,
                    tension: 0.4,
                    pointRadius: 4,
                    pointHoverRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { position: 'top' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + context.parsed.y.toFixed(4) + '%';
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value.toFixed(4) + '%';
                        }
                    },
                    title: {
                        display: true,
                        text: '中签率 (%)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: '时间'
                    }
                }
            }
        }
    });
}

// 北交所中签率图表
function createBeijingExchangeLotteryChart() {
    const ctx = document.getElementById('beijingExchangeLotteryChart').getContext('2d');
    const rawData = excelData['原始数据']?.data || [];
    
    // 筛选北交所数据 - 只筛选ipo_board='北交所'
    const beijingData = rawData.filter(item => item.ipo_board === '北交所');
    
    console.log('=== 北交所数据调试 ===');
    console.log('原始数据总数:', rawData.length);
    console.log('北交所数据总数:', beijingData.length);
    
    // 按时间排序（从早到晚）
    beijingData.sort((a, b) => new Date(a.listing_date) - new Date(b.listing_date));
    
    console.log('=== 所有北交所数据的lottery_online字段 ===');
    beijingData.forEach((item, index) => {
        console.log(`${index + 1}. ${item.sec_name} (${item.listing_date}): lottery_online = ${item.lottery_online}`);
    });
    
    // 创建标签和数据
    const labels = beijingData.map(item => {
        const date = new Date(item.listing_date);
        return `${item.sec_name}\n${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}`;
    });
    
    const lotteryOnline = beijingData.map(item => {
        const value = parseFloat(item.lottery_online) * 100 || 0;
        console.log(`${item.sec_name}: ${item.lottery_online} -> ${value}%`);
        return value;
    });
    
    console.log('图表标签数量:', labels.length);
    console.log('图表数据数量:', lotteryOnline.length);
    console.log('时间范围:', labels[0], '到', labels[labels.length - 1]);
    
    // 计算平均值
    const avgOnline = lotteryOnline.length > 0 
        ? lotteryOnline.reduce((a, b) => a + b, 0) / lotteryOnline.length 
        : 0;
    
    document.getElementById('beijingAvgLotteryOnline').textContent = avgOnline.toFixed(4) + '%';
    document.getElementById('beijingLotteryCount').textContent = beijingData.length;
    
    charts.beijingExchangeLottery = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: '北交所网上中签率 (%)',
                    data: lotteryOnline,
                    borderColor: '#9b59b6',
                    backgroundColor: 'rgba(155, 89, 182, 0.1)',
                    fill: false,
                    tension: 0.4,
                    pointRadius: 4,
                    pointHoverRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { position: 'top' },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + context.parsed.y.toFixed(4) + '%';
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value.toFixed(4) + '%';
                        }
                    },
                    title: {
                        display: true,
                        text: '中签率 (%)'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: '时间'
                    }
                }
            }
        }
    });
}

function createIssuanceChart() {
    const ctx = document.getElementById('issuanceChart').getContext('2d');
    const data = excelData['发行统计']?.data || [];
    
    const labels = data.map(item => item.week_label);
    const stockCount = data.map(item => item.stock_count || 0);
    const raisedFund = data.map(item => parseFloat((item.total_raised_fund || 0).toFixed(2)));
    
    charts.issuance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: '发行数量',
                    data: stockCount,
                    backgroundColor: 'rgba(155, 89, 182, 0.8)',
                    borderColor: '#9b59b6',
                    borderWidth: 1,
                    yAxisID: 'y'
                },
                {
                    label: '募资金额 (亿)',
                    data: raisedFund,
                    type: 'line',
                    borderColor: '#f39c12',
                    backgroundColor: 'rgba(243, 156, 18, 0.3)',
                    fill: false,
                    tension: 0.4,
                    borderWidth: 3,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            scales: {
                y: {
                    type: 'linear',
                    display: true,
                    position: 'left',
                    beginAtZero: true,
                    title: { display: true, text: '发行数量 (只)' }
                },
                y1: {
                    type: 'linear',
                    display: true,
                    position: 'right',
                    beginAtZero: true,
                    title: { display: true, text: '募资金额 (亿元)' },
                    grid: { drawOnChartArea: false }
                }
            }
        }
    });
}

function createSectorPerformanceChart() {
    const ctx = document.getElementById('sectorPerformanceChart').getContext('2d');
    const rawData = excelData['板块涨跌幅']?.data || [];
    
    if (rawData.length === 0) return;
    
    const sectors = ['上证主板', '深证主板', '科创板', '创业板', '北交所'];
    
    // 如果数据包含month_label，使用月度数据
    if (rawData[0]?.month_label) {
        const months = [...new Set(rawData.map(item => item.month_label))].sort();
        
        const datasets = sectors.map((sector, index) => {
            const colors = ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF'];
            const data = months.map(month => {
                const monthData = rawData.find(item => item.month_label === month);
                const value = monthData ? monthData[sector] : null;
                return (value !== null && value !== undefined && !isNaN(value) && value !== 0) 
                    ? (value * 100) : null;
            });
            
            return {
                label: sector,
                data: data,
                borderColor: colors[index],
                backgroundColor: colors[index] + '20',
                fill: false,
                tension: 0.4,
                spanGaps: false,
                pointRadius: data.map(val => val === null ? 0 : 4),
                pointHoverRadius: data.map(val => val === null ? 0 : 6)
            };
        });
        
        charts.sectorPerformance = new Chart(ctx, {
            type: 'line',
            data: { labels: months, datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'top', labels: { usePointStyle: true } },
                    tooltip: {
                        filter: function(tooltipItem) { return tooltipItem.parsed.y !== null; },
                        callbacks: {
                            label: function(context) {
                                return context.dataset.label + ': ' + context.parsed.y.toFixed(2) + '%';
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { callback: function(value) { return value + '%'; } },
                        title: { display: true, text: '月度平均涨跌幅 (%)' }
                    },
                    x: { title: { display: true, text: '月份' } }
                },
                interaction: { intersect: false, mode: 'index' }
            }
        });
    } else {
        // 兼容周度数据，转换为月度显示
        // 将week_label转换为月份
        const monthlyData = {};
        
        rawData.forEach(item => {
            if (!item.week_label) return;
            
            // 从week_label提取年月 (如 24W02 -> 2024-01)
            const yearSuffix = item.week_label.substring(0, 2);
            const weekNum = parseInt(item.week_label.substring(3));
            const year = 2000 + parseInt(yearSuffix);
            const month = Math.ceil(weekNum / 4.33); // 大概每月4.33周
            const monthKey = `${year}-${month.toString().padStart(2, '0')}`;
            
            if (!monthlyData[monthKey]) {
                monthlyData[monthKey] = {};
                sectors.forEach(sector => {
                    monthlyData[monthKey][sector] = [];
                });
            }
            
            sectors.forEach(sector => {
                const value = item[sector];
                if (value !== null && value !== undefined && !isNaN(value)) {
                    monthlyData[monthKey][sector].push(value);
                }
            });
        });
        
        // 计算每个月的平均值
        const months = Object.keys(monthlyData).sort();
        
        const datasets = sectors.map((sector, index) => {
            const colors = ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF'];
            const data = months.map(month => {
                const values = monthlyData[month][sector];
                if (values.length === 0) return null;
                const avg = values.reduce((a, b) => a + b, 0) / values.length;
                return avg * 100;
            });
            
            return {
                label: sector,
                data: data,
                borderColor: colors[index],
                backgroundColor: colors[index] + '20',
                fill: false,
                tension: 0.4,
                spanGaps: false,
                pointRadius: data.map(val => val === null ? 0 : 4),
                pointHoverRadius: data.map(val => val === null ? 0 : 6)
            };
        });
        
        charts.sectorPerformance = new Chart(ctx, {
            type: 'line',
            data: { labels: months, datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'top', labels: { usePointStyle: true } },
                    tooltip: {
                        filter: function(tooltipItem) { return tooltipItem.parsed.y !== null; },
                        callbacks: {
                            label: function(context) {
                                return context.dataset.label + ': ' + context.parsed.y.toFixed(2) + '%';
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: { callback: function(value) { return value + '%'; } },
                        title: { display: true, text: '月度平均涨跌幅 (%)' }
                    },
                    x: { title: { display: true, text: '月份' } }
                },
                interaction: { intersect: false, mode: 'index' }
            }
        });
    }
}

function createPeRatioChart() {
    const ctx = document.getElementById('peRatioChart').getContext('2d');
    const data = excelData['市盈率统计']?.data || [];
    
    const labels = data.map(item => item.week_label);
    const ipoPe = data.map(item => item.ipo_pe ? item.ipo_pe.toFixed(1) : null);
    const industryPe = data.map(item => item.industry_pe ? item.industry_pe.toFixed(1) : null);
    
    charts.peRatio = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'IPO市盈率',
                    data: ipoPe,
                    borderColor: '#8e44ad',
                    backgroundColor: 'rgba(142, 68, 173, 0.1)',
                    fill: false,
                    tension: 0.4,
                    spanGaps: true
                },
                {
                    label: '行业市盈率',
                    data: industryPe,
                    borderColor: '#34495e',
                    backgroundColor: 'rgba(52, 73, 94, 0.1)',
                    fill: false,
                    tension: 0.4,
                    spanGaps: true
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'top' } },
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: '市盈率' }
                }
            }
        }
    });
}

// 更新统计数据
function updateStatistics() {
    console.log('🔄 updateStatistics 函数被调用');
    console.log('excelData keys:', Object.keys(excelData || {}));
    
    try {
        // 周度收益统计
        const weeklyData = excelData['周度收益']?.data || [];
        if (weeklyData.length > 0) {
            const returns = weeklyData.map(item => item['年化收益贡献']);
            const avgReturn = (returns.reduce((a, b) => a + b, 0) / returns.length * 100).toFixed(2);
            const maxReturn = (Math.max(...returns) * 100).toFixed(2);
            
            document.getElementById('avgWeeklyReturn').textContent = avgReturn + '%';
            document.getElementById('maxWeeklyReturn').textContent = maxReturn + '%';
        }
        
        // 北交所统计
        const bjData = excelData['北交所']?.data || [];
        if (bjData.length > 0) {
            const returns = bjData.map(item => item['北交所年化收益贡献']);
            const avgReturn = (returns.reduce((a, b) => a + b, 0) / returns.length * 100).toFixed(2);
            const totalReturn = (returns.reduce((a, b) => a + b, 0) * 100).toFixed(2);
            
            document.getElementById('avgBjReturn').textContent = avgReturn + '%';
            document.getElementById('totalBjReturn').textContent = totalReturn + '%';
        }
        
        // 其他统计...
        console.log('📊 开始调用子统计函数...');
        updateLotteryStats();
        updateIssuanceStats();
        updateSectorStats();
        updatePeStats();
        console.log('✅ 所有统计函数调用完成');
        
    } catch (error) {
        console.error('❌ 统计数据更新失败:', error);
    }
}

function updateLotteryStats() {
    // 这个函数现在不需要了，因为中签率统计已经在各自的图表创建函数中处理
    // 主板中签率在 createMainBoardLotteryChart 中处理
    // 北交所中签率在 createBeijingExchangeLotteryChart 中处理
    console.log('📊 updateLotteryStats: 中签率统计已在图表创建函数中处理');
}

function updateIssuanceStats() {
    const issuanceData = excelData['发行统计']?.data || [];
    console.log('发行统计数据:', issuanceData.length, '条记录');
    
    if (issuanceData.length > 0) {
        const totalStocks = issuanceData.reduce((a, b) => a + (b.stock_count || 0), 0);
        const totalRaised = issuanceData.reduce((a, b) => a + (b.total_raised_fund || 0), 0);
        
        console.log('总发行数量:', totalStocks);
        console.log('总募资金额:', totalRaised.toFixed(1), '亿');
        
        const stocksElement = document.getElementById('totalStocks');
        const raisedElement = document.getElementById('totalRaised');
        
        console.log('stocksElement:', stocksElement);
        console.log('raisedElement:', raisedElement);
        
        if (stocksElement) stocksElement.textContent = totalStocks;
        if (raisedElement) raisedElement.textContent = totalRaised.toFixed(1) + '亿';
    } else {
        console.log('❌ 发行统计数据为空');
    }
}

function updateSectorStats() {
    const sectorData = excelData['板块涨跌幅']?.data || [];
    console.log('板块涨跌幅数据:', sectorData.length, '条记录');
    
    if (sectorData.length > 0) {
        const sectors = ['上证主板', '深证主板', '科创板', '创业板', '北交所'];
        let bestSector = '';
        let bestReturn = -Infinity; // 初始化为负无穷，以处理负收益情况
        
        sectors.forEach(sector => {
            const validReturns = sectorData
                .map(item => item[sector])
                .filter(val => val !== null && val !== undefined && !isNaN(val));
            
            if (validReturns.length > 0) {
                const avgReturn = validReturns.reduce((a, b) => a + b, 0) / validReturns.length;
                console.log(`${sector}: 平均涨幅 ${avgReturn.toFixed(2)}%`);
                if (avgReturn > bestReturn) {
                    bestReturn = avgReturn;
                    bestSector = sector;
                }
            }
        });
        
        console.log('最佳板块:', bestSector, '涨幅:', bestReturn.toFixed(2) + '%');
        
        const sectorElement = document.getElementById('bestSector');
        const returnElement = document.getElementById('bestSectorReturn');
        
        console.log('sectorElement:', sectorElement);
        console.log('returnElement:', returnElement);
        
        if (bestSector && sectorElement && returnElement) {
            sectorElement.textContent = bestSector;
            returnElement.textContent = (bestReturn * 100).toFixed(2) + '%';
        }
    } else {
        console.log('❌ 板块涨跌幅数据为空');
    }
}

function updatePeStats() {
    const peData = excelData['市盈率统计']?.data || [];
    if (peData.length > 0) {
        const ipoPeValues = peData.map(item => item.ipo_pe).filter(val => val !== null && val !== undefined);
        const industryPeValues = peData.map(item => item.industry_pe).filter(val => val !== null && val !== undefined);
        
        if (ipoPeValues.length > 0) {
            const avgIpoPe = (ipoPeValues.reduce((a, b) => a + b, 0) / ipoPeValues.length).toFixed(1);
            document.getElementById('avgIpoPe').textContent = avgIpoPe;
        }
        
        if (industryPeValues.length > 0) {
            const avgIndustryPe = (industryPeValues.reduce((a, b) => a + b, 0) / industryPeValues.length).toFixed(1);
            document.getElementById('avgIndustryPe').textContent = avgIndustryPe;
        }
    }
}

// 收益测算工具功能
function toggleDropdown() {
    const options = document.getElementById('boardOptions');
    options.classList.toggle('show');
}

function toggleAll(checkbox) {
    const allCheckboxes = document.querySelectorAll('#boardOptions input[type="checkbox"]');
    allCheckboxes.forEach(cb => {
        if (cb.value !== 'all') {
            cb.checked = checkbox.checked;
        }
    });
    updateSelection();
}

function updateSelection() {
    const checkboxes = document.querySelectorAll('#boardOptions input[type="checkbox"]:not([value="all"])');
    const allCheckbox = document.querySelector('#boardOptions input[value="all"]');
    const selectedBoards = document.getElementById('selectedBoards');
    
    const checkedBoxes = Array.from(checkboxes).filter(cb => cb.checked);
    
    // 更新全选状态
    allCheckbox.checked = checkedBoxes.length === checkboxes.length;
    
    // 更新显示文本
    if (checkedBoxes.length === 0) {
        selectedBoards.textContent = '请选择板块';
    } else if (checkedBoxes.length === checkboxes.length) {
        selectedBoards.textContent = '全部板块';
    } else if (checkedBoxes.length === 1) {
        selectedBoards.textContent = checkedBoxes[0].value;
    } else {
        selectedBoards.textContent = `已选择${checkedBoxes.length}个板块`;
    }
}

function getSelectedBoards() {
    const checkboxes = document.querySelectorAll('#boardOptions input[type="checkbox"]:not([value="all"])');
    return Array.from(checkboxes)
        .filter(cb => cb.checked)
        .map(cb => cb.value);
}

function calculateReturns() {
    const netAssetsWan = parseFloat(document.getElementById('netAssets').value) || 0;
    const riskFreeRate = parseFloat(document.getElementById('riskFreeRate').value) || 0;
    const selectedBoards = getSelectedBoards();
    
    if (netAssetsWan <= 0) {
        alert('请输入有效的净资产金额');
        return;
    }
    
    if (selectedBoards.length === 0) {
        alert('请至少选择一个板块');
        return;
    }
    
    // 转换单位：万元 -> 元
    const aum = netAssetsWan * 10000;
    console.log('💡 单位转换:', {
        输入的万元: netAssetsWan,
        转换后的元: aum.toLocaleString(),
        对应亿元: (aum / 100000000).toFixed(2) + '亿'
    });
    // 转换单位：百分比 -> 小数 (如1.4% -> 0.014)
    const rf = riskFreeRate / 100;
    
    // 为了匹配Python测试，使用特定时间范围：2025-01-01 到 2025-07-21
    const testStartDate = new Date('2025-01-01');
    const testEndDate = new Date('2025-07-21');
    
    console.log('🧪 测试参数 (匹配Python):', {
        时间范围: '2025-01-01 到 2025-07-21',
        资产规模: (aum / 100000000).toFixed(2) + '亿元',
        选择板块: selectedBoards.length > 0 ? selectedBoards : '全部板块'
    });
    
    // 调用rt_estimation逻辑，使用指定时间范围
    const result = rtEstimation(aum, selectedBoards, rf, testStartDate, testEndDate);
    
    // 更新结果显示
    document.getElementById('expectedReturn').textContent = (result.totalReturn * 100).toFixed(2) + '%';
    document.getElementById('expectedProfit').textContent = '¥' + Math.round(aum * result.totalReturn).toLocaleString();
    document.getElementById('riskAdjustedReturn').textContent = ((result.totalReturn - rf) * 100).toFixed(2) + '%';
    document.getElementById('recommendedAllocation').textContent = calculateRecommendedAllocation(result.totalReturn * 100, riskFreeRate) + '%';
    
    // 显示计算详情
    showRtEstimationBreakdown(selectedBoards, result, netAssetsWan, riskFreeRate);
}

// rt_estimation逻辑实现 - 完全按照Python函数逻辑
function rtEstimation(aum, selectedBoards, rf, startDate = null, endDate = null) {
    const rawData = excelData['原始数据']?.data || [];
    
    console.log('🔍 rt_estimation开始 - 完全按照Python逻辑');
    console.log('📥 输入参数:', {
        aum: aum.toLocaleString() + '元',
        selectedBoards: selectedBoards,
        rf: rf,
        startDate: startDate?.toISOString().split('T')[0],
        endDate: endDate?.toISOString().split('T')[0]
    });
    
    if (rawData.length === 0) {
        console.log('❌ 原始数据为空');
        return { ipoReturn: 0, freecashReturn: 0, totalReturn: rf, validCount: 0, totalRtamt: 0 };
    }
    
    // 步骤1: 时间筛选 - ipoinfo_sub = ipoinfo[(ipoinfo.listing_date>=startdate) & (ipoinfo.listing_date<=enddate)]
    let ipoinfo_sub = rawData;
    if (startDate && endDate) {
        ipoinfo_sub = rawData.filter(item => {
            const listingDate = new Date(item.listing_date);
            return listingDate >= startDate && listingDate <= endDate;
        });
        console.log('📅 时间筛选:', {
            原始数据: rawData.length,
            筛选后: ipoinfo_sub.length,
            时间范围: `${startDate.toISOString().split('T')[0]} 到 ${endDate.toISOString().split('T')[0]}`
        });
    }
    
    // 步骤2: 计算rtamt - ipoinfo_sub['rtamt'] = ipoinfo_sub.apply(lambda x: min(aum,x.offline_maxbuyamt)*x.pctchg*x.lottery_b,axis = 1)
    let totalRtamt = 0;
    let validCount = 0;
    
    ipoinfo_sub.forEach((item, index) => {
        if (item.offline_maxbuyamt !== null && item.offline_maxbuyamt !== undefined && 
            item.pctchg !== null && item.pctchg !== undefined &&
            item.lottery_b !== null && item.lottery_b !== undefined) {
            
            const offlineMaxbuyamt = parseFloat(item.offline_maxbuyamt);
            const pctchg = parseFloat(item.pctchg);
            const lotteryB = parseFloat(item.lottery_b);
            
            // 完全按照Python公式: min(aum, offline_maxbuyamt) * pctchg * lottery_b
            // 注意：这里需要确认pctchg和lottery_b的单位
            const rtamt = Math.min(aum, offlineMaxbuyamt) * pctchg * lotteryB;
            totalRtamt += rtamt;
            validCount++;
            
            if (validCount <= 5) {
                console.log(`✅ rtamt计算${validCount}:`, {
                    sec_name: item.sec_name,
                    min_aum_offline: Math.min(aum, offlineMaxbuyamt).toLocaleString(),
                    pctchg: pctchg,
                    lottery_b: lotteryB,
                    rtamt: rtamt.toFixed(2)
                });
            }
        }
    });
    
    // 步骤3: 板块筛选 - if len(board)>0: ipoinfo_sub = ipoinfo_sub[ipoinfo_sub.ipo_board.isin(board)]
    if (selectedBoards.length > 0) {
        // 重新筛选并重新计算rtamt
        const boardFilteredData = ipoinfo_sub.filter(item => selectedBoards.includes(item.ipo_board));
        
        totalRtamt = 0;
        validCount = 0;
        
        boardFilteredData.forEach(item => {
            if (item.offline_maxbuyamt !== null && item.offline_maxbuyamt !== undefined && 
                item.pctchg !== null && item.pctchg !== undefined &&
                item.lottery_b !== null && item.lottery_b !== undefined) {
                
                const offlineMaxbuyamt = parseFloat(item.offline_maxbuyamt);
                const pctchg = parseFloat(item.pctchg);
                const lotteryB = parseFloat(item.lottery_b);
                
                const rtamt = Math.min(aum, offlineMaxbuyamt) * pctchg * lotteryB;
                totalRtamt += rtamt;
                validCount++;
            }
        });
        
        console.log('📋 板块筛选后:', {
            筛选前: ipoinfo_sub.length,
            筛选后: boardFilteredData.length,
            选择板块: selectedBoards,
            有效记录: validCount,
            总rtamt: totalRtamt.toFixed(2)
        });
        
        ipoinfo_sub = boardFilteredData;
    }
    
    // 步骤4: 计算时间跨度
    const delta = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24));
    
    // 步骤5: 计算iport - iport = (365/delta)*(ipoinfo_sub.rtamt.sum()/aum)
    const iport = (365 / delta) * (totalRtamt / aum);
    
    // 步骤6: 计算freecash_rt - freecash_rt = rf*(aum-1.2*80000000)/aum
    const freecash_rt = rf * (aum - 1.2 * 80000000) / aum;
    
    // 步骤7: 计算tot_rt - tot_rt = iport+freecash_rt
    const tot_rt = iport + freecash_rt;
    
    console.log('🎯 Python公式计算结果:', {
        delta天数: delta,
        '365/delta': (365/delta).toFixed(4),
        'rtamt.sum()': totalRtamt.toFixed(2),
        'rtamt.sum()/aum': (totalRtamt/aum).toFixed(8),
        'iport': (iport * 100).toFixed(4) + '%',
        'freecash_rt': (freecash_rt * 100).toFixed(4) + '%',
        'tot_rt': (tot_rt * 100).toFixed(4) + '%'
    });
    
    return {
        ipoReturn: iport,
        freecashReturn: freecash_rt,
        totalReturn: tot_rt,
        validCount: validCount,
        totalRtamt: totalRtamt,
        startDate: startDate,
        endDate: endDate,
        deltaYears: delta / 365
    };
}

function calculateBoardReturns(selectedBoards) {
    const boardReturns = {};
    
    // 从周度收益数据计算各板块收益
    const weeklyData = excelData['周度收益']?.data || [];
    const sectorData = excelData['板块涨跌幅']?.data || [];
    
    selectedBoards.forEach(board => {
        let returns = [];
        
        // 从板块涨跌幅数据获取收益
        if (sectorData.length > 0) {
            returns = sectorData
                .map(item => item[board])
                .filter(val => val !== null && val !== undefined && !isNaN(val));
        }
        
        // 如果没有数据，使用默认值
        if (returns.length === 0) {
            const defaultReturns = {
                '上证主板': 8.5,
                '深证主板': 9.2,
                '科创板': 12.8,
                '创业板': 11.5,
                '北交所': 15.3
            };
            boardReturns[board] = defaultReturns[board] || 10.0;
        } else {
            // 计算年化收益率
            const avgReturn = returns.reduce((a, b) => a + b, 0) / returns.length;
            boardReturns[board] = avgReturn * 100; // 转换为百分比
        }
    });
    
    return boardReturns;
}

function calculateExpectedReturn(boardReturns, netAssets) {
    const boards = Object.keys(boardReturns);
    const returns = Object.values(boardReturns);
    
    // 根据净资产规模调整收益预期
    let scaleAdjustment = 1.0;
    if (netAssets < 100) {
        scaleAdjustment = 0.8; // 小资金中签率低
    } else if (netAssets > 1000) {
        scaleAdjustment = 1.2; // 大资金中签率相对稳定
    }
    
    // 计算加权平均收益
    const avgReturn = returns.reduce((a, b) => a + b, 0) / returns.length;
    
    return avgReturn * scaleAdjustment;
}

function calculateRecommendedAllocation(expectedReturn, riskFreeRate) {
    const excessReturn = expectedReturn - riskFreeRate;
    
    if (excessReturn <= 0) {
        return 0;
    } else if (excessReturn < 3) {
        return 30;
    } else if (excessReturn < 6) {
        return 50;
    } else if (excessReturn < 10) {
        return 70;
    } else {
        return 85;
    }
}

function showRtEstimationBreakdown(selectedBoards, result, netAssetsWan, riskFreeRate) {
    const breakdown = document.getElementById('calculationBreakdown');
    
    let html = `
        <p><strong>计算基础:</strong></p>
        <ul>
            <li>净资产规模: ${netAssetsWan}万元 (${(netAssetsWan * 10000).toLocaleString()}元)</li>
            <li>无风险利率: ${riskFreeRate}% (${(riskFreeRate/100).toFixed(4)})</li>
            <li>选择板块: ${selectedBoards.length > 0 ? selectedBoards.join(', ') : '全部板块'}</li>
            <li>数据时间范围: ${result.startDate ? result.startDate.toISOString().split('T')[0] : '未知'} 至 ${result.endDate ? result.endDate.toISOString().split('T')[0] : '未知'}</li>
            <li>时间跨度: ${result.deltaYears ? result.deltaYears.toFixed(2) : '1.00'}年</li>
            <li>有效IPO数量: ${result.validCount}个</li>
        </ul>
        
        <p><strong>收益构成:</strong></p>
        <ul>
            <li>打新收益率: ${(result.ipoReturn * 100).toFixed(2)}%</li>
            <li>闲置资金收益率: ${(result.freecashReturn * 100).toFixed(2)}%</li>
            <li>总收益率: ${(result.totalReturn * 100).toFixed(2)}%</li>
        </ul>
        
        <p><strong>收益计算:</strong></p>
        <ul>
            <li>总打新收益金额: ¥${Math.round(result.totalRtamt).toLocaleString()}</li>
            <li>年化打新收益: ¥${Math.round(netAssetsWan * 10000 * result.ipoReturn).toLocaleString()}</li>
            <li>闲置资金收益: ¥${Math.round(netAssetsWan * 10000 * result.freecashReturn).toLocaleString()}</li>
        </ul>
    `;
    
    breakdown.innerHTML = html;
}

function showCalculationBreakdown(selectedBoards, boardReturns, netAssets, expectedReturn, riskFreeRate) {
    const breakdown = document.getElementById('calculationBreakdown');
    
    let html = `
        <p><strong>计算基础:</strong></p>
        <ul>
            <li>净资产规模: ${netAssets}万元</li>
            <li>无风险利率: ${riskFreeRate}%</li>
            <li>选择板块: ${selectedBoards.join(', ')}</li>
        </ul>
        
        <p><strong>各板块预期收益率:</strong></p>
        <ul>
    `;
    
    selectedBoards.forEach(board => {
        html += `<li>${board}: ${boardReturns[board].toFixed(2)}%</li>`;
    });
    
    html += `
        </ul>
        
        <p><strong>风险提示:</strong></p>
        <ul>
            <li>以上收益率基于历史数据计算，不代表未来表现</li>
            <li>新股投资存在破发风险，请谨慎投资</li>
            <li>建议分散投资，控制单一资产比例</li>
        </ul>
    `;
    
    breakdown.innerHTML = html;
}

// 点击外部关闭下拉框
document.addEventListener('click', function(event) {
    const dropdown = document.getElementById('boardOptions');
    const header = document.querySelector('.multi-select-header');
    
    if (dropdown && header && !header.contains(event.target) && !dropdown.contains(event.target)) {
        dropdown.classList.remove('show');
    }
});

// IPO发行热力日历图表（分板块展示）
function createIpoHeatmapChart() {
    const ctx = document.getElementById('ipoHeatmapChart').getContext('2d');
    const rawData = excelData['原始数据']?.data || [];
    
    if (rawData.length === 0) return;
    
    // 定义板块和颜色
    const boards = ['上证主板', '深证主板', '科创板', '创业板', '北交所'];
    const boardColors = {
        '上证主板': 'rgba(255, 99, 132, 0.8)',
        '深证主板': 'rgba(54, 162, 235, 0.8)', 
        '科创板': 'rgba(255, 206, 86, 0.8)',
        '创业板': 'rgba(75, 192, 192, 0.8)',
        '北交所': 'rgba(153, 102, 255, 0.8)'
    };
    
    // 按月和板块统计IPO数量
    const monthlyBoardStats = {};
    let totalPeakCount = 0;
    let peakMonth = '';
    
    rawData.forEach(item => {
        if (!item.listing_date || !item.ipo_board) return;
        
        const date = new Date(item.listing_date);
        const monthKey = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
        const board = item.ipo_board;
        
        if (!monthlyBoardStats[monthKey]) {
            monthlyBoardStats[monthKey] = {};
            boards.forEach(b => monthlyBoardStats[monthKey][b] = 0);
        }
        
        if (boards.includes(board)) {
            monthlyBoardStats[monthKey][board]++;
        }
    });
    
    // 计算每月总数和峰值
    const months = Object.keys(monthlyBoardStats).sort();
    months.forEach(month => {
        const monthTotal = Object.values(monthlyBoardStats[month]).reduce((sum, count) => sum + count, 0);
        if (monthTotal > totalPeakCount) {
            totalPeakCount = monthTotal;
            peakMonth = month;
        }
    });
    
    // 更新统计信息
    document.getElementById('peakMonth').textContent = peakMonth || '-';
    document.getElementById('peakCount').textContent = totalPeakCount || '-';
    
    // 准备图表数据集
    const datasets = boards.map(board => ({
        label: board,
        data: months.map(month => monthlyBoardStats[month][board] || 0),
        backgroundColor: boardColors[board],
        borderColor: boardColors[board].replace('0.8', '1'),
        borderWidth: 1
    }));
    
    charts.ipoHeatmap = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: months,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { 
                    position: 'top',
                    labels: {
                        usePointStyle: true,
                        padding: 15
                    }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    callbacks: {
                        title: function(context) {
                            return context[0].label + '月IPO发行情况';
                        },
                        label: function(context) {
                            return `${context.dataset.label}: ${context.parsed.y}只`;
                        },
                        footer: function(tooltipItems) {
                            const total = tooltipItems.reduce((sum, item) => sum + item.parsed.y, 0);
                            return `总计: ${total}只`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    title: {
                        display: true,
                        text: '月份'
                    }
                },
                y: {
                    stacked: true,
                    beginAtZero: true,
                    ticks: { stepSize: 1 },
                    title: {
                        display: true,
                        text: '发行数量'
                    }
                }
            },
            interaction: {
                mode: 'index',
                intersect: false
            }
        }
    });
}

// 更新综合统计数据
function updateComprehensiveStats() {
    const rawData = excelData['原始数据']?.data || [];
    const lotteryData = excelData['中签率统计']?.data || [];
    
    if (rawData.length === 0) return;
    
    // 计算总IPO数量
    const totalCount = rawData.length;
    document.getElementById('totalIpoCount').textContent = totalCount;
    
    // 计算总募资金额
    const totalRaised = rawData.reduce((sum, item) => {
        const amount = parseFloat(item.actual_raised_fund) || 0;
        return sum + amount;
    }, 0);
    document.getElementById('totalRaisedAmount').textContent = totalRaised.toFixed(1);
    
    // 计算平均首日涨幅
    const validReturns = rawData.filter(item => 
        item.pctchg !== null && item.pctchg !== undefined && !isNaN(item.pctchg)
    );
    const avgReturn = validReturns.length > 0 
        ? validReturns.reduce((sum, item) => sum + parseFloat(item.pctchg), 0) / validReturns.length
        : 0;
    document.getElementById('avgFirstDayReturn').textContent = avgReturn.toFixed(2) + '%';
    
    // 计算平均中签率
    if (lotteryData.length > 0) {
        const avgLottery = lotteryData.reduce((sum, item) => {
            const rateA = parseFloat(item.lottery_a) * 100 || 0;
            const rateB = parseFloat(item.lottery_b) * 100 || 0;
            return sum + (rateA + rateB) / 2;
        }, 0) / lotteryData.length;
        document.getElementById('avgLotteryRate').textContent = avgLottery.toFixed(4) + '%';
    }
    
    // 计算上涨成功率
    const positiveReturns = validReturns.filter(item => parseFloat(item.pctchg) > 0);
    const successRate = validReturns.length > 0 
        ? (positiveReturns.length / validReturns.length * 100)
        : 0;
    document.getElementById('successRate').textContent = successRate.toFixed(1) + '%';
    
    // 计算平均市盈率
    const validPE = rawData.filter(item => 
        item.ipo_pe && !isNaN(parseFloat(item.ipo_pe)) && parseFloat(item.ipo_pe) > 0
    );
    const avgPE = validPE.length > 0 
        ? validPE.reduce((sum, item) => sum + parseFloat(item.ipo_pe), 0) / validPE.length
        : 0;
    document.getElementById('avgPE').textContent = avgPE.toFixed(1);
}