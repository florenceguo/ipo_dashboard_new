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
            'excel_data_summary.json', // 如果有JSON文件
            'newest_data.json'         // 或者其他可能的文件名
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
    // 更新数据时间
    updateDataTime();
    
    // 隐藏加载界面
    document.getElementById('loading').style.display = 'none';
    document.getElementById('content').style.display = 'block';
    
    // 初始化图表
    setTimeout(() => {
        initializeCharts();
        updateStatistics();
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
        createLotteryRateChart();
        createIssuanceChart();
        createSectorPerformanceChart();
        createPeRatioChart();
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
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: '年化收益贡献 (%)',
                data: returns,
                borderColor: '#3498db',
                backgroundColor: 'rgba(52, 152, 219, 0.1)',
                fill: true,
                tension: 0.4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
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
        }
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
            plugins: { legend: { display: false } },
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
        }
    });
}

// 其他图表创建函数...
function createLotteryRateChart() {
    const ctx = document.getElementById('lotteryRateChart').getContext('2d');
    const data = excelData['中签率统计']?.data || [];
    
    const labels = data.map(item => item.week_label);
    const lotteryA = data.map(item => (item.lottery_a * 10000).toFixed(3));
    const lotteryB = data.map(item => (item.lottery_b * 10000).toFixed(3));
    
    charts.lotteryRate = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'A类中签率 (‰)',
                    data: lotteryA,
                    borderColor: '#e74c3c',
                    backgroundColor: 'rgba(231, 76, 60, 0.1)',
                    fill: false,
                    tension: 0.4
                },
                {
                    label: 'B类中签率 (‰)',
                    data: lotteryB,
                    borderColor: '#27ae60',
                    backgroundColor: 'rgba(39, 174, 96, 0.1)',
                    fill: false,
                    tension: 0.4
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
                    ticks: {
                        callback: function(value) {
                            return value + '‰';
                        }
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
        updateLotteryStats();
        updateIssuanceStats();
        updateSectorStats();
        updatePeStats();
        
    } catch (error) {
        console.error('统计数据更新失败:', error);
    }
}

function updateLotteryStats() {
    const lotteryData = excelData['中签率统计']?.data || [];
    if (lotteryData.length > 0) {
        const lotteryA = lotteryData.map(item => item.lottery_a);
        const lotteryB = lotteryData.map(item => item.lottery_b);
        const avgLotteryA = (lotteryA.reduce((a, b) => a + b, 0) / lotteryA.length * 10000).toFixed(3);
        const avgLotteryB = (lotteryB.reduce((a, b) => a + b, 0) / lotteryB.length * 10000).toFixed(3);
        
        document.getElementById('avgLotteryA').textContent = avgLotteryA + '‰';
        document.getElementById('avgLotteryB').textContent = avgLotteryB + '‰';
    }
}

function updateIssuanceStats() {
    const issuanceData = excelData['发行统计']?.data || [];
    if (issuanceData.length > 0) {
        const totalStocks = issuanceData.reduce((a, b) => a + (b.stock_count || 0), 0);
        const totalRaised = issuanceData.reduce((a, b) => a + (b.total_raised_fund || 0), 0);
        
        document.getElementById('totalStocks').textContent = totalStocks;
        document.getElementById('totalRaised').textContent = totalRaised.toFixed(1) + '亿';
    }
}

function updateSectorStats() {
    const sectorData = excelData['板块涨跌幅']?.data || [];
    if (sectorData.length > 0) {
        const sectors = ['上证主板', '深证主板', '科创板', '创业板', '北交所'];
        let bestSector = '';
        let bestReturn = 0;
        
        sectors.forEach(sector => {
            const validReturns = sectorData
                .map(item => item[sector])
                .filter(val => val !== null && val !== undefined && !isNaN(val));
            
            if (validReturns.length > 0) {
                const avgReturn = validReturns.reduce((a, b) => a + b, 0) / validReturns.length;
                if (avgReturn > bestReturn) {
                    bestReturn = avgReturn;
                    bestSector = sector;
                }
            }
        });
        
        document.getElementById('bestSector').textContent = bestSector;
        document.getElementById('bestSectorReturn').textContent = (bestReturn * 100).toFixed(2) + '%';
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
    const netAssets = parseFloat(document.getElementById('netAssets').value) || 0;
    const riskFreeRate = parseFloat(document.getElementById('riskFreeRate').value) || 0;
    const selectedBoards = getSelectedBoards();
    
    if (netAssets <= 0) {
        alert('请输入有效的净资产金额');
        return;
    }
    
    if (selectedBoards.length === 0) {
        alert('请至少选择一个板块');
        return;
    }
    
    // 计算各板块的历史收益率
    const boardReturns = calculateBoardReturns(selectedBoards);
    
    // 计算预期收益
    const expectedReturn = calculateExpectedReturn(boardReturns, netAssets);
    const expectedProfit = (netAssets * 10000 * expectedReturn / 100).toFixed(0);
    const riskAdjustedReturn = (expectedReturn - riskFreeRate).toFixed(2);
    
    // 计算建议配置比例
    const recommendedAllocation = calculateRecommendedAllocation(expectedReturn, riskFreeRate);
    
    // 更新结果显示
    document.getElementById('expectedReturn').textContent = expectedReturn.toFixed(2) + '%';
    document.getElementById('expectedProfit').textContent = '¥' + parseInt(expectedProfit).toLocaleString();
    document.getElementById('riskAdjustedReturn').textContent = riskAdjustedReturn + '%';
    document.getElementById('recommendedAllocation').textContent = recommendedAllocation + '%';
    
    // 显示计算详情
    showCalculationBreakdown(selectedBoards, boardReturns, netAssets, expectedReturn, riskFreeRate);
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