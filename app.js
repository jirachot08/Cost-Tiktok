/**
 * Cost Dashboard - TikTok & Facebook Ads
 * Pulls data from Google Sheets and displays interactive charts and tables
 */

// Configuration
const CONFIG = {
    spreadsheetId: '1cFnle7-557TOR3Sv8HqriV6C1eBUJ-qj48AvEZoVzzs',
    sheets: {
        summary: {
            gid: '312126150',
            name: 'ผลรวมทั้งหมด',
            theme: 'summary-theme',
            colors: {
                primary: '#9333ea',
                secondary: '#f59e0b',
                gradient: ['rgba(147, 51, 234, 0.8)', 'rgba(245, 158, 11, 0.8)']
            }
        },
        tiktok: {
            gid: '80338864',
            name: 'TikTok',
            theme: 'tiktok-theme',
            colors: {
                primary: '#fe2c55',
                secondary: '#25f4ee',
                gradient: ['rgba(254, 44, 85, 0.8)', 'rgba(37, 244, 238, 0.8)']
            }
        },
        facebook: {
            gid: '0',
            name: 'Facebook',
            theme: 'facebook-theme',
            colors: {
                primary: '#1877f2',
                secondary: '#42b72a',
                gradient: ['rgba(24, 119, 242, 0.8)', 'rgba(66, 183, 42, 0.8)']
            }
        },
        // Budget sheets by month (TikTok)
        budgetTiktok: {
            JAN: { gid: '1427937661', name: 'Budget TikTok JAN' },
            FEB: { gid: '784237547', name: 'Budget TikTok FEB' },
            MAR: { gid: '782118916', name: 'Budget TikTok MAR' },
            APR: { gid: '1838576178', name: 'Budget TikTok APR' },
            MAY: { gid: '681531460', name: 'Budget TikTok MAY' },
            JUN: { gid: '1330523840', name: 'Budget TikTok JUN' }
        },
        // Budget sheets by month (Facebook)
        budgetFacebook: {
            FEB: { gid: 'BUDGET_FB_FEB_GID', name: 'Budget Facebook FEB' },
            MAR: { gid: 'BUDGET_FB_MAR_GID', name: 'Budget Facebook MAR' },
            APR: { gid: 'BUDGET_FB_APR_GID', name: 'Budget Facebook APR' },
            MAY: { gid: 'BUDGET_FB_MAY_GID', name: 'Budget Facebook MAY' },
            JUN: { gid: 'BUDGET_FB_JUN_GID', name: 'Budget Facebook JUN' }
        }
    }
};

// State
let currentPlatform = 'summary';
let selectedBudgetMonth = 'JAN';  // Current selected budget month
let platformData = {
    summary: null,
    tiktok: null,
    facebook: null
};
let budgetData = null;  // Budget data from sheet
let charts = {
    daily: null,
    product: null
};
let refreshIntervalId = null;
let scheduledRefreshTimeoutId = null;
const SCHEDULED_REFRESH_HOUR = 12; // Refresh at 12:00 noon
const SCHEDULED_REFRESH_MINUTE = 0;

// DOM Elements
const elements = {
    loadingIndicator: document.getElementById('loadingIndicator'),
    dashboardContent: document.getElementById('dashboardContent'),
    errorState: document.getElementById('errorState'),
    errorMessage: document.getElementById('errorMessage'),
    tabButtons: document.querySelectorAll('.tab-btn'),
    totalCost: document.getElementById('totalCost'),
    monthlyCost: document.getElementById('monthlyCost'),
    productCount: document.getElementById('productCount'),
    daysCount: document.getElementById('daysCount'),
    lastUpdated: document.getElementById('lastUpdated'),
    topProductsList: document.getElementById('topProductsList'),
    tableHeader: document.getElementById('tableHeader'),
    tableBody: document.getElementById('tableBody'),
    searchInput: document.getElementById('searchInput'),
    monthFilter: document.getElementById('monthFilter'),
    productMonthFilter: document.getElementById('productMonthFilter'),
    dailyChart: document.getElementById('dailyChart'),
    productChart: document.getElementById('productChart'),
    refreshBtn: document.getElementById('refreshBtn'),
    countdown: document.getElementById('countdown'),
    autoRefreshStatus: document.getElementById('autoRefreshStatus'),
    // Date Range Picker
    datePickerToggle: document.getElementById('datePickerToggle'),
    dateRangeModal: document.getElementById('dateRangeModal'),
    dateRangeOverlay: document.getElementById('dateRangeOverlay'),
    selectedDateRange: document.getElementById('selectedDateRange'),
    calendar1: document.getElementById('calendar1'),
    calendar2: document.getElementById('calendar2'),
    calMonth1Label: document.getElementById('calMonth1Label'),
    calMonth2Label: document.getElementById('calMonth2Label'),
    prevMonth: document.getElementById('prevMonth'),
    nextMonth: document.getElementById('nextMonth'),
    applyDateRange: document.getElementById('applyDateRange'),
    cancelDateRange: document.getElementById('cancelDateRange'),
    // Budget Tracker
    budgetSection: document.getElementById('budgetSection'),
    budgetStatus: document.getElementById('budgetStatus'),
    budgetSpent: document.getElementById('budgetSpent'),
    budgetTotal: document.getElementById('budgetTotal'),
    budgetProgressBar: document.getElementById('budgetProgressBar'),
    budgetPercentage: document.getElementById('budgetPercentage'),
    dailyBudget: document.getElementById('dailyBudget'),
    daysPassed: document.getElementById('daysPassed'),
    daysRemaining: document.getElementById('daysRemaining'),
    shouldSpent: document.getElementById('shouldSpent'),
    categoryBudgets: document.getElementById('categoryBudgets'),
    budgetMonthSelect: document.getElementById('budgetMonthSelect'),
    budgetMonthSelectHeader: document.getElementById('budgetMonthSelectHeader'),
    // Budget View Page
    budgetViewContent: document.getElementById('budgetViewContent'),
    budgetViewTitle: document.getElementById('budgetViewTitle'),
    budgetViewMonth: document.getElementById('budgetViewMonth'),
    budgetViewTotal: document.getElementById('budgetViewTotal'),
    budgetViewSpent: document.getElementById('budgetViewSpent'),
    budgetViewRemaining: document.getElementById('budgetViewRemaining'),
    budgetViewDaily: document.getElementById('budgetViewDaily'),
    budgetViewStatus: document.getElementById('budgetViewStatus'),
    budgetViewProgressBar: document.getElementById('budgetViewProgressBar'),
    budgetViewPercent: document.getElementById('budgetViewPercent'),
    budgetViewShouldSpent: document.getElementById('budgetViewShouldSpent'),
    budgetViewDaysRemaining: document.getElementById('budgetViewDaysRemaining'),
    budgetViewCategories: document.getElementById('budgetViewCategories')
};


// Global filter state
let globalDateFilter = {
    startDate: null,
    endDate: null,
    tempStartDate: null,
    tempEndDate: null
};

// Calendar state
let calendarState = {
    currentMonth: new Date().getMonth(),
    currentYear: new Date().getFullYear()
};

// Initialize
document.addEventListener('DOMContentLoaded', init);

async function init() {
    setupEventListeners();
    await loadAllData();
}

function setupEventListeners() {
    // Tab switching
    elements.tabButtons.forEach(btn => {
        btn.addEventListener('click', () => switchPlatform(btn.dataset.platform));
    });

    // Search and filter
    elements.searchInput.addEventListener('input', debounce(filterTable, 300));
    elements.monthFilter.addEventListener('change', filterTable);

    // Product month filter
    elements.productMonthFilter.addEventListener('change', updateProductsByMonth);

    // Refresh button
    elements.refreshBtn.addEventListener('click', manualRefresh);

    // Date Range Picker
    elements.datePickerToggle.addEventListener('click', toggleDatePicker);
    elements.applyDateRange.addEventListener('click', applyDateRangeFilter);
    elements.cancelDateRange.addEventListener('click', closeDatePicker);
    elements.prevMonth.addEventListener('click', () => navigateMonth(-1));
    elements.nextMonth.addEventListener('click', () => navigateMonth(1));

    // Quick filter options
    document.querySelectorAll('.quick-filter-option').forEach(btn => {
        btn.addEventListener('click', () => handleQuickFilter(btn.dataset.range));
    });

    // Close modal when clicking overlay
    elements.dateRangeOverlay.addEventListener('click', closeDatePicker);

    // Budget month selector (Category section)
    if (elements.budgetMonthSelect) {
        elements.budgetMonthSelect.addEventListener('change', async (e) => {
            await handleBudgetMonthChange(e.target.value);
        });
    }

    // Budget month selector (Header section)
    if (elements.budgetMonthSelectHeader) {
        elements.budgetMonthSelectHeader.addEventListener('change', async (e) => {
            await handleBudgetMonthChange(e.target.value);
        });
    }
}

// Handle budget month change and sync both dropdowns
async function handleBudgetMonthChange(month) {
    selectedBudgetMonth = month;

    // Sync both dropdowns
    if (elements.budgetMonthSelect) {
        elements.budgetMonthSelect.value = month;
    }
    if (elements.budgetMonthSelectHeader) {
        elements.budgetMonthSelectHeader.value = month;
    }

    // Clear cache for this platform and month to force refetch
    if (budgetDataCache[currentPlatform]) {
        budgetDataCache[currentPlatform][selectedBudgetMonth] = null;
    }

    await updateBudgetTracker();
}

// Data Loading
async function loadAllData() {
    showLoading();

    try {
        // Load all platforms in parallel
        const [summaryData, tiktokData, facebookData] = await Promise.all([
            fetchSheetData('summary'),
            fetchSheetData('tiktok'),
            fetchSheetData('facebook')
        ]);

        // Use special parser for summary sheet (different structure)
        platformData.summary = parseSummaryCSVData(summaryData);
        platformData.tiktok = parseCSVData(tiktokData);
        platformData.facebook = parseCSVData(facebookData);

        showDashboard();
        updateDashboard();
        updateLastUpdated();

        // Start scheduled refresh at noon daily
        startScheduledRefresh();
    } catch (error) {
        console.error('Error loading data:', error);
        showError(error.message);
    }
}

// Auto-refresh data from Google Sheets
async function refreshData() {
    try {
        console.log('🔄 Auto-refreshing data...');

        // Fetch new data in background
        const [summaryData, tiktokData, facebookData] = await Promise.all([
            fetchSheetData('summary'),
            fetchSheetData('tiktok'),
            fetchSheetData('facebook')
        ]);

        // Use special parser for summary sheet
        platformData.summary = parseSummaryCSVData(summaryData);
        platformData.tiktok = parseCSVData(tiktokData);
        platformData.facebook = parseCSVData(facebookData);

        // Update dashboard without showing loading
        updateDashboard();
        updateLastUpdated();

        console.log('✅ Data refreshed successfully');
    } catch (error) {
        console.warn('⚠️ Auto-refresh failed:', error.message);
        // Don't show error to user, just log it
    }
}

function startScheduledRefresh() {
    // Clear existing scheduled refresh
    stopScheduledRefresh();

    // Update display with next refresh time
    updateNextRefreshDisplay();

    // Check every minute if it's time to refresh
    refreshIntervalId = setInterval(() => {
        const now = new Date();
        const currentHour = now.getHours();
        const currentMinute = now.getMinutes();

        // Refresh at scheduled time (noon)
        if (currentHour === SCHEDULED_REFRESH_HOUR && currentMinute === SCHEDULED_REFRESH_MINUTE) {
            console.log('🕛 Scheduled refresh triggered at noon');
            refreshData();
            updateNextRefreshDisplay();
        }

        // Update display every minute
        updateNextRefreshDisplay();
    }, 60000); // Check every minute

    console.log(`⏰ Scheduled refresh set for ${SCHEDULED_REFRESH_HOUR}:${String(SCHEDULED_REFRESH_MINUTE).padStart(2, '0')} daily`);
}

function stopScheduledRefresh() {
    if (refreshIntervalId) {
        clearInterval(refreshIntervalId);
        refreshIntervalId = null;
    }
    if (scheduledRefreshTimeoutId) {
        clearTimeout(scheduledRefreshTimeoutId);
        scheduledRefreshTimeoutId = null;
    }
}

function updateNextRefreshDisplay() {
    if (!elements.countdown || !elements.autoRefreshStatus) return;

    const now = new Date();
    const nextRefresh = new Date();

    // Set next refresh time to noon
    nextRefresh.setHours(SCHEDULED_REFRESH_HOUR, SCHEDULED_REFRESH_MINUTE, 0, 0);

    // If noon has passed today, set for tomorrow
    if (now >= nextRefresh) {
        nextRefresh.setDate(nextRefresh.getDate() + 1);
    }

    // Calculate time difference
    const diffMs = nextRefresh - now;
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffMinutes = Math.floor((diffMs % (1000 * 60 * 60)) / (1000 * 60));

    // Update display
    if (diffHours > 0) {
        elements.countdown.textContent = `${diffHours}ชม. ${diffMinutes}น.`;
    } else {
        elements.countdown.textContent = `${diffMinutes} นาที`;
    }

    // Update status text
    elements.autoRefreshStatus.innerHTML = `🕛 รีเฟรชถัดไป: <span id="countdown">${elements.countdown.textContent}</span>`;
}

async function manualRefresh() {
    // Add refreshing state to button
    elements.refreshBtn.classList.add('refreshing');
    elements.refreshBtn.textContent = '🔄 กำลังโหลด...';

    try {
        await refreshData();
        updateNextRefreshDisplay();
    } finally {
        // Remove refreshing state
        elements.refreshBtn.classList.remove('refreshing');
        elements.refreshBtn.textContent = '🔄 รีเฟรช';
    }
}

// Global Date Filter Functions

const monthNamesThai = {
    0: 'มกราคม', 1: 'กุมภาพันธ์', 2: 'มีนาคม',
    3: 'เมษายน', 4: 'พฤษภาคม', 5: 'มิถุนายน',
    6: 'กรกฎาคม', 7: 'สิงหาคม', 8: 'กันยายน',
    9: 'ตุลาคม', 10: 'พฤศจิกายน', 11: 'ธันวาคม'
};

function toggleDatePicker() {
    const modal = elements.dateRangeModal;
    if (modal.classList.contains('hidden')) {
        openDatePicker();
    } else {
        closeDatePicker();
    }
}

function openDatePicker() {
    // Initialize temp dates with current selection
    globalDateFilter.tempStartDate = globalDateFilter.startDate;
    globalDateFilter.tempEndDate = globalDateFilter.endDate;

    elements.dateRangeOverlay.classList.remove('hidden');
    elements.dateRangeModal.classList.remove('hidden');
    renderCalendars();
    updateQuickFilterHighlight();
}

function closeDatePicker() {
    elements.dateRangeOverlay.classList.add('hidden');
    elements.dateRangeModal.classList.add('hidden');
    globalDateFilter.tempStartDate = null;
    globalDateFilter.tempEndDate = null;
}

function navigateMonth(direction) {
    calendarState.currentMonth += direction;
    if (calendarState.currentMonth < 0) {
        calendarState.currentMonth = 11;
        calendarState.currentYear--;
    } else if (calendarState.currentMonth > 11) {
        calendarState.currentMonth = 0;
        calendarState.currentYear++;
    }
    renderCalendars();
}

function renderCalendars() {
    const month1 = calendarState.currentMonth;
    const year1 = calendarState.currentYear;
    const month2 = month1 === 11 ? 0 : month1 + 1;
    const year2 = month1 === 11 ? year1 + 1 : year1;

    elements.calMonth1Label.textContent = `${monthNamesThai[month1]} ${year1}`;
    elements.calMonth2Label.textContent = `${monthNamesThai[month2]} ${year2}`;

    elements.calendar1.innerHTML = renderCalendarGrid(year1, month1);
    elements.calendar2.innerHTML = renderCalendarGrid(year2, month2);

    // Add click listeners to days
    document.querySelectorAll('.calendar-day:not(.empty):not(.disabled)').forEach(day => {
        day.addEventListener('click', () => handleDayClick(day.dataset.date));
    });
}

function renderCalendarGrid(year, month) {
    const weekdays = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let html = '<div class="calendar-grid">';

    // Weekday headers
    weekdays.forEach(day => {
        html += `<div class="calendar-weekday">${day}</div>`;
    });

    // Empty cells before first day
    for (let i = 0; i < firstDay; i++) {
        html += '<div class="calendar-day empty"></div>';
    }

    // Days
    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const date = new Date(year, month, day);
        date.setHours(0, 0, 0, 0);

        let classes = ['calendar-day'];

        // Today
        if (date.getTime() === today.getTime()) {
            classes.push('today');
        }

        // Selected range
        if (globalDateFilter.tempStartDate && globalDateFilter.tempEndDate) {
            const start = new Date(globalDateFilter.tempStartDate);
            const end = new Date(globalDateFilter.tempEndDate);
            start.setHours(0, 0, 0, 0);
            end.setHours(0, 0, 0, 0);

            if (date.getTime() === start.getTime()) {
                classes.push('range-start');
            } else if (date.getTime() === end.getTime()) {
                classes.push('range-end');
            } else if (date > start && date < end) {
                classes.push('in-range');
            }
        } else if (globalDateFilter.tempStartDate) {
            const start = new Date(globalDateFilter.tempStartDate);
            start.setHours(0, 0, 0, 0);
            if (date.getTime() === start.getTime()) {
                classes.push('selected');
            }
        }

        html += `<div class="${classes.join(' ')}" data-date="${dateStr}">${day}</div>`;
    }

    html += '</div>';
    return html;
}

function handleDayClick(dateStr) {
    if (!globalDateFilter.tempStartDate || globalDateFilter.tempEndDate) {
        // Start new selection
        globalDateFilter.tempStartDate = dateStr;
        globalDateFilter.tempEndDate = null;
    } else {
        // Complete selection
        if (dateStr < globalDateFilter.tempStartDate) {
            globalDateFilter.tempEndDate = globalDateFilter.tempStartDate;
            globalDateFilter.tempStartDate = dateStr;
        } else {
            globalDateFilter.tempEndDate = dateStr;
        }
    }
    renderCalendars();
    updateQuickFilterHighlight();
}

function handleQuickFilter(range) {
    const today = new Date();
    let startDate, endDate;

    switch (range) {
        case 'lifetime':
            startDate = null;
            endDate = null;
            break;
        case 'today':
            startDate = formatDateISO(today);
            endDate = formatDateISO(today);
            break;
        case 'yesterday':
            const yesterday = new Date(today);
            yesterday.setDate(yesterday.getDate() - 1);
            startDate = formatDateISO(yesterday);
            endDate = formatDateISO(yesterday);
            break;
        case 'last7days':
            const last7 = new Date(today);
            last7.setDate(last7.getDate() - 6);
            startDate = formatDateISO(last7);
            endDate = formatDateISO(today);
            break;
        case 'last14days':
            const last14 = new Date(today);
            last14.setDate(last14.getDate() - 13);
            startDate = formatDateISO(last14);
            endDate = formatDateISO(today);
            break;
        case 'last30days':
            const last30 = new Date(today);
            last30.setDate(last30.getDate() - 29);
            startDate = formatDateISO(last30);
            endDate = formatDateISO(today);
            break;
        case 'thisweek':
            const dayOfWeek = today.getDay();
            const thisWeekStart = new Date(today);
            thisWeekStart.setDate(today.getDate() - dayOfWeek);
            startDate = formatDateISO(thisWeekStart);
            endDate = formatDateISO(today);
            break;
        case 'lastweek':
            const lastWeekEnd = new Date(today);
            lastWeekEnd.setDate(today.getDate() - today.getDay() - 1);
            const lastWeekStart = new Date(lastWeekEnd);
            lastWeekStart.setDate(lastWeekEnd.getDate() - 6);
            startDate = formatDateISO(lastWeekStart);
            endDate = formatDateISO(lastWeekEnd);
            break;
        case 'thismonth':
            const thisMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
            startDate = formatDateISO(thisMonthStart);
            endDate = formatDateISO(today);
            break;
        case 'lastmonth':
            const lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0);
            const lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
            startDate = formatDateISO(lastMonthStart);
            endDate = formatDateISO(lastMonthEnd);
            break;
    }

    globalDateFilter.tempStartDate = startDate;
    globalDateFilter.tempEndDate = endDate;
    renderCalendars();
    updateQuickFilterHighlight();
}

function updateQuickFilterHighlight() {
    document.querySelectorAll('.quick-filter-option').forEach(btn => {
        btn.classList.remove('active');
    });

    if (!globalDateFilter.tempStartDate && !globalDateFilter.tempEndDate) {
        document.querySelector('.quick-filter-option[data-range="lifetime"]')?.classList.add('active');
    }
}

function formatDateISO(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function applyDateRangeFilter() {
    globalDateFilter.startDate = globalDateFilter.tempStartDate;
    globalDateFilter.endDate = globalDateFilter.tempEndDate;

    closeDatePicker();
    updateSelectedDateRangeDisplay();
    applyGlobalDateFilter();
}

function updateSelectedDateRangeDisplay() {
    let displayText = 'ทั้งหมด';

    if (globalDateFilter.startDate && globalDateFilter.endDate) {
        if (globalDateFilter.startDate === globalDateFilter.endDate) {
            displayText = formatDate(globalDateFilter.startDate);
        } else {
            displayText = `${formatDate(globalDateFilter.startDate)} - ${formatDate(globalDateFilter.endDate)}`;
        }
    }

    elements.selectedDateRange.textContent = displayText;
}

function applyGlobalDateFilter() {
    const data = platformData[currentPlatform];
    if (!data) return;

    // Filter data based on global date filter
    const filteredData = getFilteredData(data);

    // Update all dashboard components with filtered data
    updateSummaryCardsFiltered(filteredData);
    updateTopProducts(filteredData, 'all');
    updateChartsFiltered(filteredData);
    filterTable();
}

function getFilteredData(data) {
    // If no date range set, return original data
    if (!globalDateFilter.startDate || !globalDateFilter.endDate) {
        return data;
    }

    const startDate = new Date(globalDateFilter.startDate);
    const endDate = new Date(globalDateFilter.endDate);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);

    const filteredDailyData = data.dailyData.filter(day => {
        const dayDate = new Date(day.date);
        dayDate.setHours(12, 0, 0, 0);
        return dayDate >= startDate && dayDate <= endDate;
    });

    // Calculate new totals
    const filteredProductTotals = {};
    data.headers.forEach(header => {
        filteredProductTotals[header] = 0;
    });

    filteredDailyData.forEach(day => {
        data.headers.forEach(header => {
            filteredProductTotals[header] += (day.values[header] || 0);
        });
    });

    return {
        ...data,
        dailyData: filteredDailyData,
        productTotals: filteredProductTotals,
        totalCost: Object.values(filteredProductTotals).reduce((a, b) => a + b, 0)
    };
}

function updateSummaryCardsFiltered(data) {
    const totalCost = data.totalCost || Object.values(data.productTotals).reduce((a, b) => a + b, 0);

    // For filtered data, show filtered total as "monthly" 
    let periodCost = totalCost;

    elements.totalCost.textContent = formatCurrency(totalCost);
    elements.monthlyCost.textContent = formatCurrency(periodCost);
    elements.productCount.textContent = data.headers.length;
    elements.daysCount.textContent = data.dailyData.length;
}

function updateChartsFiltered(data) {
    // Update daily chart
    if (charts.daily) {
        charts.daily.data.labels = data.dailyData.map(d => formatDate(d.date));
        charts.daily.data.datasets[0].data = data.dailyData.map(d => {
            return data.headers.reduce((sum, h) => sum + (d.values[h] || 0), 0);
        });
        charts.daily.update('none');
    }

    // Update product chart (top 10)
    if (charts.product) {
        const sorted = Object.entries(data.productTotals)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10);

        charts.product.data.labels = sorted.map(item => item[0]);
        charts.product.data.datasets[0].data = sorted.map(item => item[1]);
        charts.product.update('none');
    }
}

// Budget Tracker Functions
// ==========================================

const MONTH_NAMES_TH = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

// Budget data storage (by platform and month)
let budgetDataCache = {
    tiktok: {},  // { FEB: data, MAR: data, ... }
    facebook: {} // { FEB: data, MAR: data, ... }
};

// Fetch budget data from Google Sheet
async function fetchBudgetData(platform) {
    const sheetKey = platform === 'tiktok' ? 'budgetTiktok' : 'budgetFacebook';
    const budgetSheets = CONFIG.sheets[sheetKey];

    console.log('[Budget] Fetching for:', platform, 'Month:', selectedBudgetMonth);

    // Get sheet config for selected month
    const monthSheet = budgetSheets ? budgetSheets[selectedBudgetMonth] : null;
    console.log('[Budget] Month sheet:', monthSheet);

    if (!monthSheet || monthSheet.gid.includes('GID')) {
        console.log('[Budget] Using default - no sheet configured');
        return getDefaultBudgetForPlatform(platform);
    }

    // Add cache-busting timestamp to force fresh data
    const cacheBuster = Date.now();
    const originalUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.spreadsheetId}/export?format=csv&gid=${monthSheet.gid}&_=${cacheBuster}`;

    // Try multiple fetch methods (same as fetchSheetData)
    const methods = [
        // Method 1: Direct fetch
        async () => {
            const response = await fetchWithTimeout(originalUrl, 10000);
            if (!response.ok) throw new Error('Direct fetch failed');
            return await response.text();
        },
        // Method 2: Using corsproxy.io
        async () => {
            const proxyUrl = `https://corsproxy.io/?${encodeURIComponent(originalUrl)}`;
            console.log('[Budget] Using proxy:', proxyUrl);
            const response = await fetchWithTimeout(proxyUrl, 15000);
            if (!response.ok) throw new Error('Corsproxy fetch failed');
            return await response.text();
        }
    ];

    for (const method of methods) {
        try {
            const csvText = await method();
            if (csvText && csvText.length > 10) {
                console.log('[Budget] CSV received, lines:', csvText.split('\n').length);
                const result = parseBudgetCSV(csvText, platform);
                console.log('[Budget] Parsed categories:', result.categories.length, 'TotalBudget:', result.totalBudget);
                return result;
            }
        } catch (e) {
            console.warn('[Budget] Fetch method failed:', e.message);
        }
    }

    console.log('[Budget] All methods failed, using default');
    return getDefaultBudgetForPlatform(platform);
}

function parseBudgetCSV(csvText, platform) {
    const lines = csvText.trim().split('\n');
    console.log('[Budget Parse] Total lines:', lines.length);
    console.log('[Budget Parse] Raw first line:', lines[0]);
    console.log('[Budget Parse] Raw CSV preview:', csvText.substring(0, 500));

    if (lines.length < 2) return getDefaultBudgetForPlatform(platform);

    // Helper function to parse CSV line with quoted values
    function parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim().replace(/"/g, ''));
                current = '';
            } else if (char !== '\r') {
                current += char;
            }
        }
        result.push(current.trim().replace(/"/g, ''));
        return result;
    }

    const headers = parseCSVLine(lines[0]);
    console.log('[Budget Parse] Headers:', headers);

    // Find column indices for styles
    const colIndices = {
        product: 0,
        abx: headers.findIndex(h => h.toLowerCase().includes('abx')),
        ace: headers.findIndex(h => h.toLowerCase().includes('ace')),
        sale: headers.findIndex(h => h.toLowerCase().includes('sale')),
        review: headers.findIndex(h => h.toLowerCase().includes('review')),
        branding: headers.findIndex(h => h.toLowerCase().includes('branding'))
    };
    console.log('[Budget Parse] Column indices:', colIndices);

    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();

    const categories = [];
    let totalABX = 0;
    let totalACE = 0;

    for (let i = 1; i < lines.length; i++) {
        const values = parseCSVLine(lines[i]);
        if (values.length < 2) continue;

        const product = values[0];
        if (product === 'SUM' || !product) continue;

        // Debug first few rows
        if (i <= 3) {
            console.log(`[Budget Parse] Row ${i}:`, values);
        }

        // Parse ABX budget
        const abxBudget = colIndices.abx >= 0 ?
            (parseFloat(values[colIndices.abx]?.replace(/,/g, '')) || 0) : 0;

        // Parse ACE budget (might be in different positions)
        const aceBudget = colIndices.ace >= 0 ?
            (parseFloat(values[colIndices.ace]?.replace(/,/g, '')) || 0) : 0;

        // Parse style percentages
        const parsePercent = (val) => {
            if (!val) return 0;
            const numStr = val.replace('%', '').replace(/,/g, '').trim();
            const num = parseFloat(numStr);
            // If value is like "60%" return 0.6, if like "0.6" return 0.6
            return num > 1 ? num / 100 : num;
        };

        const salePercent = colIndices.sale >= 0 ? parsePercent(values[colIndices.sale]) : 0;
        const reviewPercent = colIndices.review >= 0 ? parsePercent(values[colIndices.review]) : 0;
        const brandingPercent = colIndices.branding >= 0 ? parsePercent(values[colIndices.branding]) : 0;

        // Calculate style amounts (based on ABX budget)
        const styles = [];
        if (salePercent > 0) {
            styles.push({ name: 'Sale', percent: salePercent, amount: abxBudget * salePercent });
        }
        if (reviewPercent > 0) {
            styles.push({ name: 'Review', percent: reviewPercent, amount: abxBudget * reviewPercent });
        }
        if (brandingPercent > 0) {
            styles.push({ name: 'Branding', percent: brandingPercent, amount: abxBudget * brandingPercent });
        }

        if (abxBudget > 0 || aceBudget > 0) {
            categories.push({
                code: product,
                name: product,
                abx: abxBudget,
                ace: aceBudget,
                budget: abxBudget + aceBudget,
                styles: styles
            });
        }

        totalABX += abxBudget;
        totalACE += aceBudget;
    }

    console.log('[Budget Parse] Result: categories:', categories.length, 'totalBudget:', totalABX + totalACE);

    return {
        month: currentMonth,
        year: currentYear,
        daysInMonth: daysInMonth,
        totalBudget: totalABX + totalACE,
        totalABX: totalABX,
        totalACE: totalACE,
        categories: categories
    };
}

function getDefaultBudgetForPlatform(platform) {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();

    if (platform === 'tiktok') {
        return {
            month: currentMonth,
            year: currentYear,
            daysInMonth: daysInMonth,
            totalBudget: 50454250,
            totalABX: 28590000,
            totalACE: 21864250,
            categories: [
                { code: 'L3', name: 'L3', abx: 7310000, ace: 5848000, budget: 13158000 },
                { code: 'L4', name: 'L4', abx: 5250000, ace: 3150000, budget: 8400000 },
                { code: 'L6', name: 'L6', abx: 1575000, ace: 2100000, budget: 3675000 },
                { code: 'L10', name: 'L10', abx: 1530000, ace: 2805000, budget: 4335000 },
                { code: 'L8', name: 'L8', abx: 937500, ace: 1718750, budget: 2656250 },
                { code: 'L7', name: 'L7', abx: 675000, ace: 1125000, budget: 1800000 },
                { code: 'L19', name: 'L19', abx: 1080000, ace: 1080000, budget: 2160000 },
                { code: 'C4', name: 'C4', abx: 540000, ace: 900000, budget: 1440000 },
                { code: 'S1', name: 'S1', abx: 100000, ace: 0, budget: 100000 },
                { code: 'S2', name: 'S2', abx: 140000, ace: 0, budget: 140000 },
                { code: 'S3', name: 'S3', abx: 525000, ace: 0, budget: 525000 },
                { code: 'S4', name: 'S4', abx: 320000, ace: 0, budget: 320000 },
                { code: 'A1', name: 'A1', abx: 840000, ace: 840000, budget: 1680000 },
                { code: 'L11', name: 'L11', abx: 0, ace: 200000, budget: 200000 },
                { code: 'L13', name: 'L13', abx: 960000, ace: 960000, budget: 1920000 },
                { code: 'L14', name: 'L14', abx: 60000, ace: 140000, budget: 200000 },
                { code: 'C1', name: 'C1', abx: 0, ace: 175000, budget: 175000 },
                { code: 'C2', name: 'C2', abx: 0, ace: 210000, budget: 210000 },
                { code: 'C3', name: 'C3', abx: 225000, ace: 225000, budget: 450000 },
                { code: 'T5ABC', name: 'T5ABC', abx: 0, ace: 75000, budget: 75000 },
                { code: 'T6', name: 'T6', abx: 0, ace: 75000, budget: 75000 },
                { code: 'L3L4L6', name: 'L3L4L6', abx: 550000, ace: 100000, budget: 650000 },
                { code: 'L4S3', name: 'L4S3', abx: 600000, ace: 0, budget: 600000 },
                { code: 'L6S4', name: 'L6S4', abx: 175000, ace: 0, budget: 175000 },
                { code: 'S1234', name: 'S1234', abx: 157500, ace: 0, budget: 157500 },
                { code: 'L3L4L6L10', name: 'L3L4L6L10', abx: 1500000, ace: 0, budget: 1500000 },
                { code: 'L4L6L7C2C4', name: 'L4L6L7C2C4', abx: 840000, ace: 0, budget: 840000 },
                { code: 'L9L19', name: 'L9L19', abx: 300000, ace: 0, budget: 300000 },
                { code: 'C1C2', name: 'C1C2', abx: 250000, ace: 0, budget: 250000 },
                { code: 'L7L11', name: 'L7L11', abx: 0, ace: 137500, budget: 137500 },
                { code: 'T5ABC 6A', name: 'T5ABC 6A', abx: 150000, ace: 0, budget: 150000 },
                { code: 'BR', name: 'BR', abx: 2000000, ace: 0, budget: 2000000 }
            ]
        };
    } else {
        return {
            month: currentMonth,
            year: currentYear,
            daysInMonth: daysInMonth,
            totalBudget: 30000000,
            totalABX: 15000000,
            totalACE: 15000000,
            categories: [
                { code: 'FB-ADS', name: 'Facebook Ads', abx: 15000000, ace: 0, budget: 15000000 }
            ]
        };
    }
}

// Get budget data (uses cache if available)
async function getBudgetData(platform) {
    // Always fetch fresh data (cache disabled)
    console.log('[Budget] Fetching fresh data for:', platform, 'Month:', selectedBudgetMonth);
    return await fetchBudgetData(platform);
}


async function updateBudgetTracker() {
    // Don't show budget for summary view
    if (currentPlatform === 'summary') {
        if (elements.budgetSection) {
            elements.budgetSection.classList.add('hidden');
        }
        return;
    }

    // Get budget data (async)
    const budget = await getBudgetData(currentPlatform);

    if (!budget || !elements.budgetSection) {
        if (elements.budgetSection) {
            elements.budgetSection.classList.add('hidden');
        }
        return;
    }

    // Show budget section
    elements.budgetSection.classList.remove('hidden');

    // Get current platform data for actual spending
    const data = platformData[currentPlatform];
    const now = new Date();

    // Get month index from selected budget month
    const monthMap = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11 };
    const selectedMonthIndex = monthMap[selectedBudgetMonth] ?? now.getMonth();
    const selectedYear = 2026; // Hardcoded for now, could be dynamic
    const daysInSelectedMonth = new Date(selectedYear, selectedMonthIndex + 1, 0).getDate();

    // Calculate days passed based on selected month vs current date
    let daysPassed, daysRemaining;
    if (now.getFullYear() === selectedYear && now.getMonth() === selectedMonthIndex) {
        // Current month - use current day
        daysPassed = now.getDate();
        daysRemaining = daysInSelectedMonth - daysPassed;
    } else if (now.getFullYear() > selectedYear || (now.getFullYear() === selectedYear && now.getMonth() > selectedMonthIndex)) {
        // Past month - all days passed
        daysPassed = daysInSelectedMonth;
        daysRemaining = 0;
    } else {
        // Future month - no days passed yet
        daysPassed = 0;
        daysRemaining = daysInSelectedMonth;
    }

    // Calculate actual spent (from selected month data)
    let actualSpent = 0;
    if (data && data.dailyData) {
        const selectedMonthStr = `${selectedYear}-${String(selectedMonthIndex + 1).padStart(2, '0')}`;
        data.dailyData.forEach(d => {
            if (d.date.startsWith(selectedMonthStr)) {
                actualSpent += d.total || Object.values(d.values).reduce((sum, v) => sum + v, 0);
            }
        });
    }

    // Calculate percentages and targets
    const totalBudget = budget.totalBudget;
    const percentage = totalBudget > 0 ? (actualSpent / totalBudget) * 100 : 0;
    const dailyBudget = totalBudget / daysInSelectedMonth;
    const shouldHaveSpent = dailyBudget * daysPassed;
    const isOver = actualSpent > shouldHaveSpent;

    // Update month display (removed - using dropdown now)

    // Update status
    elements.budgetStatus.textContent = isOver ? 'เกินงบ' : 'ใต้งบ';
    elements.budgetStatus.className = `budget-status ${isOver ? 'over' : 'under'}`;

    // Update amounts
    elements.budgetSpent.textContent = formatCurrency(actualSpent);
    elements.budgetTotal.textContent = formatCurrency(totalBudget);

    // Update progress bar
    const progressWidth = Math.min(percentage, 100);
    elements.budgetProgressBar.style.width = `${progressWidth}%`;
    elements.budgetProgressBar.className = `progress-bar ${percentage > 100 ? 'over' : ''}`;

    // Update percentage
    elements.budgetPercentage.textContent = `${percentage.toFixed(1)}%`;

    // Update daily info
    elements.dailyBudget.textContent = formatCurrency(dailyBudget);
    elements.daysPassed.textContent = `${daysPassed} วัน`;
    elements.daysRemaining.textContent = `${daysRemaining} วัน`;
    elements.shouldSpent.textContent = formatCurrency(shouldHaveSpent);

    // Update category budgets
    updateCategoryBudgets(budget.categories, data);
}

function updateCategoryBudgets(categories, data) {
    if (!elements.categoryBudgets || !categories) return;

    const now = new Date();
    const currentMonthStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const currentDay = now.getDate();
    const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();

    let html = '';

    categories.forEach(cat => {
        // Calculate spent for this category (simplified - would need proper mapping)
        let spent = 0;
        const dailyBudget = cat.budget / daysInMonth;
        const shouldHaveSpent = dailyBudget * currentDay;

        // For now, simulate some spending based on actual data percentage
        if (data && data.dailyData) {
            // Use a portion of the budget as simulated spending
            const dayData = data.dailyData.filter(d => d.date.startsWith(currentMonthStr));
            if (dayData.length > 0) {
                // Distribute actual spending proportionally to budget
                const totalActual = dayData.reduce((sum, d) => sum + (d.total || 0), 0);
                const budgetRatio = cat.budget / categories.reduce((sum, c) => sum + c.budget, 0);
                spent = totalActual * budgetRatio * 0.5; // Simulated
            }
        }

        const percentage = cat.budget > 0 ? (spent / cat.budget) * 100 : 0;
        const isOver = percentage > (currentDay / daysInMonth) * 100;

        // Format ABX and ACE amounts
        const abxAmount = cat.abx || 0;
        const aceAmount = cat.ace || 0;
        const dailyAbx = abxAmount / daysInMonth;
        const dailyAce = aceAmount / daysInMonth;

        // Generate ABX styles HTML
        let abxStylesHtml = '';
        if (cat.styles && cat.styles.length > 0) {
            abxStylesHtml = `
                <div class="budget-type-styles">
                    ${cat.styles.map(style => `
                        <div class="style-item style-${style.name.toLowerCase()}">
                            <span class="style-name">${style.name}</span>
                            <span class="style-percent">${Math.round(style.percent * 100)}%</span>
                            <span class="style-amount">${formatCurrency(style.amount)}</span>
                        </div>
                    `).join('')}
                </div>
            `;
        }

        html += `
            <div class="category-card ${cat.styles && cat.styles.length > 0 ? 'has-styles' : ''}">
                <div class="category-header">
                    <span class="category-name">${cat.name}</span>
                    <span class="category-code">${cat.code}</span>
                </div>
                <div class="category-budget-sections">
                    <!-- ABX Section -->
                    <div class="budget-section abx-section">
                        <div class="section-title">
                            <span class="title-label">ABX</span>
                            <span class="title-amount">${formatCurrency(abxAmount)}</span>
                            <span class="title-daily">฿${formatCompactNumber(dailyAbx)}/วัน</span>
                        </div>
                        ${abxStylesHtml}
                    </div>
                    <!-- ACE Section -->
                    ${aceAmount > 0 ? `
                    <div class="budget-section ace-section">
                        <div class="section-title">
                            <span class="title-label">ACE</span>
                            <span class="title-amount">${formatCurrency(aceAmount)}</span>
                            <span class="title-daily">฿${formatCompactNumber(dailyAce)}/วัน</span>
                        </div>
                    </div>
                    ` : ''}
                </div>
                <div class="category-total">
                    <span class="total-label">รวม:</span>
                    <span class="total-amount">${formatCurrency(cat.budget)}</span>
                </div>
                <div class="category-progress">
                    <div class="bar ${isOver ? 'over' : ''}" style="width: ${Math.min(percentage, 100)}%"></div>
                </div>
            </div>
        `;
    });

    elements.categoryBudgets.innerHTML = html;
}

// Update Budget View Page (for dedicated Budget tabs)
async function updateBudgetView(platform) {
    const budget = await getBudgetData(platform);

    if (!budget) return;

    const data = platformData[platform];
    const now = new Date();
    const currentDay = now.getDate();

    // Calculate actual spent
    let actualSpent = 0;
    if (data && data.dailyData) {
        const currentMonthStr = `${budget.year}-${String(budget.month + 1).padStart(2, '0')}`;
        data.dailyData.forEach(d => {
            if (d.date.startsWith(currentMonthStr)) {
                actualSpent += d.total || Object.values(d.values).reduce((sum, v) => sum + v, 0);
            }
        });
    }

    const totalBudget = budget.totalBudget;
    const remaining = totalBudget - actualSpent;
    const percentage = totalBudget > 0 ? (actualSpent / totalBudget) * 100 : 0;
    const dailyBudget = totalBudget / budget.daysInMonth;
    const daysRemaining = budget.daysInMonth - currentDay;
    const shouldHaveSpent = dailyBudget * currentDay;
    const isOver = actualSpent > shouldHaveSpent;

    // Update header
    const platformName = platform === 'tiktok' ? 'TikTok' : 'Facebook';
    elements.budgetViewTitle.textContent = `💰 Budget ${platformName}`;
    elements.budgetViewMonth.textContent = `${MONTH_NAMES_TH[budget.month]} ${budget.year}`;

    // Update overview cards
    elements.budgetViewTotal.textContent = formatCurrency(totalBudget);
    elements.budgetViewSpent.textContent = formatCurrency(actualSpent);
    elements.budgetViewRemaining.textContent = formatCurrency(remaining);
    elements.budgetViewDaily.textContent = formatCurrency(dailyBudget);

    // Update main progress
    elements.budgetViewStatus.textContent = isOver ? 'เกินงบ' : 'ใต้งบ';
    elements.budgetViewStatus.className = `budget-status ${isOver ? 'over' : 'under'}`;

    const progressWidth = Math.min(percentage, 100);
    elements.budgetViewProgressBar.style.width = `${progressWidth}%`;
    elements.budgetViewProgressBar.className = `progress-bar-large ${percentage > 100 ? 'over' : ''}`;

    elements.budgetViewPercent.textContent = `${percentage.toFixed(1)}%`;
    elements.budgetViewShouldSpent.textContent = formatCurrency(shouldHaveSpent);
    elements.budgetViewDaysRemaining.textContent = `${daysRemaining} วัน`;

    // Update categories
    updateBudgetViewCategories(budget.categories, data, budget.daysInMonth, currentDay);
}

function updateBudgetViewCategories(categories, data, daysInMonth, currentDay) {
    if (!elements.budgetViewCategories || !categories) return;

    const now = new Date();
    const currentMonthStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

    let html = '';

    categories.forEach(cat => {
        const dailyBudget = cat.budget / daysInMonth;
        const shouldHaveSpent = dailyBudget * currentDay;

        // Calculate spent for category
        let spent = 0;
        if (data && data.dailyData) {
            const dayData = data.dailyData.filter(d => d.date.startsWith(currentMonthStr));
            if (dayData.length > 0) {
                const totalActual = dayData.reduce((sum, d) => sum + (d.total || 0), 0);
                const budgetRatio = cat.budget / categories.reduce((sum, c) => sum + c.budget, 0);
                spent = totalActual * budgetRatio * 0.5;
            }
        }

        const percentage = cat.budget > 0 ? (spent / cat.budget) * 100 : 0;
        const isOver = percentage > (currentDay / daysInMonth) * 100;
        const remaining = cat.budget - spent;

        html += `
            <div class="budget-category-item">
                <div class="cat-header">
                    <span class="cat-name">${cat.name}</span>
                    <span class="cat-code">${cat.code}</span>
                </div>
                <div class="cat-amounts">
                    <span class="spent">${formatCurrency(spent)}</span>
                    <span class="budget">/ ${formatCurrency(cat.budget)}</span>
                </div>
                <div class="cat-progress">
                    <div class="bar ${isOver ? 'over' : ''}" style="width: ${Math.min(percentage, 100)}%"></div>
                </div>
                <div class="cat-daily">
                    <span>งบต่อวัน: ${formatCurrency(dailyBudget)}</span>
                    <span>เหลือ: ${formatCurrency(remaining)}</span>
                </div>
            </div>
        `;
    });

    elements.budgetViewCategories.innerHTML = html;
}

async function fetchSheetData(platform) {
    const sheet = CONFIG.sheets[platform];
    const originalUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.spreadsheetId}/export?format=csv&gid=${sheet.gid}`;

    // Try multiple methods to fetch data
    const methods = [
        // Method 1: Direct fetch (works if no CORS issues)
        async () => {
            const response = await fetchWithTimeout(originalUrl, 10000);
            if (!response.ok) throw new Error('Direct fetch failed');
            return await response.text();
        },
        // Method 2: Using allorigins.win CORS proxy
        async () => {
            const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(originalUrl)}`;
            const response = await fetchWithTimeout(proxyUrl, 15000);
            if (!response.ok) throw new Error('Proxy fetch failed');
            return await response.text();
        },
        // Method 3: Using corsproxy.io
        async () => {
            const proxyUrl = `https://corsproxy.io/?${encodeURIComponent(originalUrl)}`;
            const response = await fetchWithTimeout(proxyUrl, 15000);
            if (!response.ok) throw new Error('Corsproxy fetch failed');
            return await response.text();
        }
    ];

    for (const method of methods) {
        try {
            const result = await method();
            if (result && result.includes('MONTH') || result.includes('วันที่')) {
                console.log(`Successfully loaded ${platform} data`);
                return result;
            }
        } catch (e) {
            console.warn(`Fetch method failed for ${platform}:`, e.message);
        }
    }

    // Fallback: Return sample data if all methods fail
    console.warn(`All fetch methods failed for ${platform}, using sample data`);
    return getSampleData(platform);
}

async function fetchWithTimeout(url, timeout) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeout);

    try {
        const response = await fetch(url, { signal: controller.signal });
        clearTimeout(timeoutId);
        return response;
    } catch (error) {
        clearTimeout(timeoutId);
        throw error;
    }
}

function getSampleData(platform) {
    // Sample data for when Google Sheets is not accessible
    const sampleTiktok = `COST TT ADS,,,,,,,,,,,,,,,
MONTH,A1,C1,C2,C3,C4,D2,D3,L10,L3,L4,L6,L7,รวม
JAN,"1,834,932","389,582","426,796","846,501","1,900,554","1,673,912","2,880,870","5,926,906","14,038,204","10,329,806","6,081,524","1,540,457","47,869,044"
FEB,"16,999","11,460","11,210","23,077","63,472","15,324","44,395","186,421","564,559","319,742","181,330","43,967","1,481,956"
SUM,"1,851,931","401,042","438,006","869,578","1,964,026","1,689,236","2,925,265","6,113,327","14,602,763","10,649,548","6,262,854","1,584,424","49,351,000"

วันที่,A1,C1,C2,C3,C4,D2,D3,L10,L3,L4,L6,L7,รวม
2026-01-01,"58,200","17,754","37,584","20,290","78,430","45,459","57,997","218,150","486,204","451,304","234,574","66,845","1,772,791"
2026-01-02,"58,200","13,230","10,198","20,398","62,655","45,449","58,000","161,727","486,696","344,006","194,743","35,553","1,490,855"
2026-01-03,"58,200","13,070","10,197","20,398","62,644","45,444","57,999","161,729","486,048","343,333","194,075","35,553","1,488,690"
2026-01-04,"58,200","12,730","10,198","20,400","62,658","45,456","57,993","161,735","486,760","343,445","194,808","35,553","1,489,936"
2026-01-05,"58,200","12,675","10,197","20,399","62,421","45,451","57,997","161,728","486,754","343,060","187,967","35,552","1,482,401"`;

    const sampleFacebook = `COST FB ADS,,,,,,,,,,,,,,,
MONTH,A1,C1,C2,C3,C4,D2,D3,L10,L3,L4,L6,L7,รวม
JAN,"413,058","103,171","107,928","486,939","514,343","156,201","178,918","1,591,058","1,964,248","2,442,401","1,787,429","484,652","10,230,346"
FEB,"0","0","0","0","0","0","0","0","0","0","0","0","0"
SUM,"413,058","103,171","107,928","486,939","514,343","156,201","178,918","1,591,058","1,964,248","2,442,401","1,787,429","484,652","10,230,346"

วันที่,A1,C1,C2,C3,C4,D2,D3,L10,L3,L4,L6,L7,รวม
2026-01-01,"14,264","3,165","4,104","18,351","16,049","6,591","6,054","43,056","60,607","90,105","67,242","17,923","347,511"
2026-01-02,"13,284","3,465","5,276","15,513","16,458","6,580","6,133","40,539","57,195","86,383","63,883","16,441","331,150"
2026-01-03,"13,056","3,372","5,329","16,721","17,022","6,469","6,028","49,201","66,579","88,273","64,094","18,646","354,790"
2026-01-04,"16,082","3,715","5,214","17,043","19,004","7,535","7,112","58,718","72,453","102,546","71,955","17,725","399,102"
2026-01-05,"12,890","3,156","4,036","13,543","14,478","7,107","6,525","49,400","59,615","79,008","52,292","14,591","316,641"`;

    return platform === 'tiktok' ? sampleTiktok : sampleFacebook;
}

function parseCSVData(csvText) {
    const lines = csvText.split('\n').map(line => {
        // Parse CSV properly handling quoted values with commas
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];

            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else if (char !== '\r') {
                current += char;
            }
        }
        result.push(current.trim());

        return result;
    });

    // Find header row (row with dates starting with YYYY)
    let headerRowIndex = -1;
    let dataStartIndex = -1;

    for (let i = 0; i < lines.length; i++) {
        if (lines[i][0] === 'วันที่' || lines[i][0] === 'Date') {
            headerRowIndex = i;
            dataStartIndex = i + 1;
            break;
        }
    }

    if (headerRowIndex === -1) {
        console.warn('Header row not found, using default structure');
        return { headers: [], dailyData: [], monthlyData: {}, productTotals: {} };
    }

    // Extract product codes from header
    const headers = lines[headerRowIndex].slice(1).filter(h => h && h !== 'รวม' && h !== 'Sum');

    // Parse daily data
    const dailyData = [];
    for (let i = dataStartIndex; i < lines.length; i++) {
        const row = lines[i];
        if (row[0] && /^\d{4}-\d{2}-\d{2}$/.test(row[0])) {
            const dateStr = row[0];
            const values = {};
            let dailyTotal = 0;

            for (let j = 0; j < headers.length; j++) {
                const value = parseNumber(row[j + 1]);
                values[headers[j]] = value;
                dailyTotal += value;
            }

            dailyData.push({
                date: dateStr,
                values: values,
                total: dailyTotal
            });
        }
    }

    // Calculate product totals
    const productTotals = {};
    headers.forEach(header => {
        productTotals[header] = dailyData.reduce((sum, day) => sum + (day.values[header] || 0), 0);
    });

    // Parse monthly summaries
    const monthlyData = {};
    const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];

    for (let i = 0; i < lines.length; i++) {
        const row = lines[i];
        if (months.includes(row[0]?.toUpperCase())) {
            const monthTotal = row.slice(1).reduce((sum, val) => sum + parseNumber(val), 0);
            if (monthTotal > 0) {
                monthlyData[row[0].toUpperCase()] = monthTotal;
            }
        }
    }

    return {
        headers,
        dailyData,
        monthlyData,
        productTotals
    };
}

// Special parser for ADS_SUM sheet (monthly structure)
function parseSummaryCSVData(csvText) {
    const lines = csvText.split('\n').map(line => {
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];

            if (char === '"') {
                inQuotes = !inQuotes;
            } else if (char === ',' && !inQuotes) {
                result.push(current.trim());
                current = '';
            } else if (char !== '\r') {
                current += char;
            }
        }
        result.push(current.trim());

        return result;
    });

    // First row should be headers: MONTH, A1, C1, C2, ...
    if (lines.length === 0 || lines[0][0] !== 'MONTH') {
        console.warn('Summary sheet has unexpected format');
        return { headers: [], dailyData: [], monthlyData: {}, productTotals: {} };
    }

    // Extract product codes from header (skip MONTH and Sum)
    const headers = lines[0].slice(1).filter(h => h && h !== 'Sum' && h !== 'รวม');

    // Month name mapping for sorting
    const monthOrder = {
        'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04',
        'MAY': '05', 'JUN': '06', 'JUL': '07', 'AUG': '08',
        'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'
    };

    const monthNamesThai = {
        'JAN': 'มกราคม', 'FEB': 'กุมภาพันธ์', 'MAR': 'มีนาคม',
        'APR': 'เมษายน', 'MAY': 'พฤษภาคม', 'JUN': 'มิถุนายน',
        'JUL': 'กรกฎาคม', 'AUG': 'สิงหาคม', 'SEP': 'กันยายน',
        'OCT': 'ตุลาคม', 'NOV': 'พฤศจิกายน', 'DEC': 'ธันวาคม'
    };

    // Parse monthly data as "daily" data (for compatibility)
    const dailyData = [];
    const productTotals = {};

    // Initialize product totals
    headers.forEach(header => {
        productTotals[header] = 0;
    });

    for (let i = 1; i < lines.length; i++) {
        const row = lines[i];
        const monthName = row[0]?.toUpperCase();

        // Skip empty rows and SUM row
        if (!monthName || monthName === 'SUM' || !monthOrder[monthName]) continue;

        // Check if row has any data
        let hasData = false;
        for (let j = 1; j < row.length && j <= headers.length; j++) {
            if (parseNumber(row[j]) > 0) {
                hasData = true;
                break;
            }
        }

        if (!hasData) continue;

        const values = {};
        let monthTotal = 0;

        for (let j = 0; j < headers.length; j++) {
            const value = parseNumber(row[j + 1]);
            values[headers[j]] = value;
            monthTotal += value;
            productTotals[headers[j]] += value;
        }

        // Use a fake date for sorting (current year - month)
        const currentYear = new Date().getFullYear();
        const fakeDate = `${currentYear}-${monthOrder[monthName]}-01`;

        dailyData.push({
            date: fakeDate,
            displayName: monthNamesThai[monthName] || monthName,
            values: values,
            total: monthTotal,
            isMonthly: true
        });
    }

    // Sort by date
    dailyData.sort((a, b) => a.date.localeCompare(b.date));

    return {
        headers,
        dailyData,
        monthlyData: {},
        productTotals,
        isSummary: true
    };
}

function parseNumber(str) {
    if (!str) return 0;
    // Remove quotes, commas, and parse
    const cleaned = str.replace(/["',]/g, '').trim();
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
}

// Display Functions
function showLoading() {
    elements.loadingIndicator.classList.remove('hidden');
    elements.dashboardContent.classList.add('hidden');
    elements.errorState.classList.add('hidden');
}

function showDashboard() {
    elements.loadingIndicator.classList.add('hidden');
    elements.dashboardContent.classList.remove('hidden');
    elements.errorState.classList.add('hidden');
}

function showError(message) {
    elements.loadingIndicator.classList.add('hidden');
    elements.dashboardContent.classList.add('hidden');
    elements.errorState.classList.remove('hidden');
    elements.errorMessage.textContent = message;
}

function switchPlatform(platform) {
    currentPlatform = platform;

    // Update theme
    document.body.classList.remove('tiktok-theme', 'facebook-theme', 'summary-theme');
    if (CONFIG.sheets[platform]) {
        document.body.classList.add(CONFIG.sheets[platform].theme);
    }

    // Update active tab
    elements.tabButtons.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.platform === platform);
    });

    // Always show dashboard content
    elements.dashboardContent.classList.remove('hidden');
    if (elements.budgetViewContent) {
        elements.budgetViewContent.classList.add('hidden');
    }
    updateDashboard();
}


function updateDashboard() {
    const data = platformData[currentPlatform];
    if (!data) return;

    updateSelectedDateRangeDisplay();

    updateSummaryCards(data);
    updateProductMonthFilter(data);
    updateTopProducts(data);
    updateTable(data);
    updateCharts(data);
    updateMonthFilter(data);
    updateBudgetTracker();
}

function updateSummaryCards(data) {
    const totalCost = Object.values(data.productTotals).reduce((sum, val) => sum + val, 0);
    const currentMonth = new Date().toLocaleString('en', { month: 'short' }).toUpperCase();
    const monthlyCost = data.monthlyData[currentMonth] ||
        data.dailyData.reduce((sum, day) => sum + day.total, 0);

    elements.totalCost.textContent = formatCurrency(totalCost);
    elements.monthlyCost.textContent = formatCurrency(monthlyCost);
    elements.productCount.textContent = data.headers.length;
    elements.daysCount.textContent = data.dailyData.length;
}

function updateTopProducts(data, selectedMonth = 'all') {
    let productTotals;

    if (selectedMonth === 'all') {
        // Use cumulative totals
        productTotals = data.productTotals;
    } else {
        // Calculate totals for selected month only
        productTotals = {};
        data.headers.forEach(header => {
            productTotals[header] = 0;
        });

        data.dailyData.forEach(day => {
            const monthNum = day.date.split('-')[1];
            if (monthNum === selectedMonth) {
                data.headers.forEach(header => {
                    productTotals[header] += (day.values[header] || 0);
                });
            }
        });
    }

    // Show ALL products sorted from highest to lowest cost
    const sorted = Object.entries(productTotals)
        .sort((a, b) => b[1] - a[1]);

    // Calculate total for selected period
    const periodTotal = sorted.reduce((sum, item) => sum + item[1], 0);

    const monthNames = {
        '01': 'มกราคม', '02': 'กุมภาพันธ์', '03': 'มีนาคม',
        '04': 'เมษายน', '05': 'พฤษภาคม', '06': 'มิถุนายน',
        '07': 'กรกฎาคม', '08': 'สิงหาคม', '09': 'กันยายน',
        '10': 'ตุลาคม', '11': 'พฤศจิกายน', '12': 'ธันวาคม'
    };

    const periodLabel = selectedMonth === 'all' ? 'สะสมทั้งหมด' : monthNames[selectedMonth];

    elements.topProductsList.innerHTML = `
        <div class="period-summary">
            <span class="period-label">📊 ${periodLabel}: </span>
            <span class="period-total">${formatCurrency(periodTotal)}</span>
        </div>
    ` + sorted.map((item, index) => `
        <div class="top-product-item">
            <div class="product-rank">${index + 1}</div>
            <div class="product-info">
                <div class="product-code">${item[0]}</div>
                <div class="product-cost">${formatCurrency(item[1])}</div>
            </div>
        </div>
    `).join('');
}

function updateProductsByMonth() {
    const data = platformData[currentPlatform];
    if (!data) return;

    const selectedMonth = elements.productMonthFilter.value;
    updateTopProducts(data, selectedMonth);
}

function updateProductMonthFilter(data) {
    const months = new Set();
    data.dailyData.forEach(day => {
        const monthNum = day.date.split('-')[1];
        months.add(monthNum);
    });

    const monthNames = {
        '01': 'มกราคม', '02': 'กุมภาพันธ์', '03': 'มีนาคม',
        '04': 'เมษายน', '05': 'พฤษภาคม', '06': 'มิถุนายน',
        '07': 'กรกฎาคม', '08': 'สิงหาคม', '09': 'กันยายน',
        '10': 'ตุลาคม', '11': 'พฤศจิกายน', '12': 'ธันวาคม'
    };

    elements.productMonthFilter.innerHTML = '<option value="all">ทุกเดือน (สะสม)</option>';
    [...months].sort().forEach(month => {
        elements.productMonthFilter.innerHTML += `<option value="${month}">${monthNames[month] || month}</option>`;
    });
}

function updateTable(data) {
    // Update header
    const headerCells = ['<th>วันที่</th>'];
    data.headers.forEach(header => {
        headerCells.push(`<th>${header}</th>`);
    });
    headerCells.push('<th>รวม</th>');
    elements.tableHeader.innerHTML = headerCells.join('');

    // Update body
    renderTableBody(data.dailyData, data.headers);
}

function renderTableBody(dailyData, headers, filter = '') {
    const searchTerm = filter.toLowerCase();
    const selectedMonth = elements.monthFilter.value;

    let filteredHeaders = headers;
    if (searchTerm) {
        filteredHeaders = headers.filter(h => h.toLowerCase().includes(searchTerm));
    }

    let filteredData = dailyData;
    if (selectedMonth !== 'all') {
        filteredData = dailyData.filter(day => {
            const monthNum = parseInt(day.date.split('-')[1]);
            return monthNum === parseInt(selectedMonth);
        });
    }

    // Calculate totals for each column
    const columnTotals = {};
    let grandTotal = 0;

    headers.forEach(header => {
        columnTotals[header] = 0;
    });

    filteredData.forEach(day => {
        headers.forEach(header => {
            const shouldShow = !searchTerm || header.toLowerCase().includes(searchTerm);
            if (shouldShow) {
                columnTotals[header] += (day.values[header] || 0);
            }
        });
        grandTotal += searchTerm
            ? filteredHeaders.reduce((sum, h) => sum + (day.values[h] || 0), 0)
            : (day.total || 0);
    });

    const rows = filteredData.map(day => {
        let rowHtml = `<td>${formatDate(day.date)}</td>`;

        headers.forEach(header => {
            const shouldShow = !searchTerm || header.toLowerCase().includes(searchTerm);
            if (shouldShow) {
                const value = day.values[header] || 0;
                rowHtml += `<td>${value > 0 ? formatCurrency(value) : '-'}</td>`;
            }
        });

        const filteredTotal = (searchTerm ?
            filteredHeaders.reduce((sum, h) => sum + (day.values[h] || 0), 0) :
            day.total);
        rowHtml += `<td style="font-weight: 600; color: var(--current-primary);">${formatCurrency(filteredTotal)}</td>`;

        return `<tr>${rowHtml}</tr>`;
    });

    // Add total row
    let totalRowHtml = `<td style="font-weight: 700; background: var(--glass-bg); color: var(--current-primary);">รวมทั้งหมด</td>`;

    headers.forEach(header => {
        const shouldShow = !searchTerm || header.toLowerCase().includes(searchTerm);
        if (shouldShow) {
            const total = columnTotals[header];
            totalRowHtml += `<td style="font-weight: 600; background: var(--glass-bg);">${total > 0 ? formatCurrency(total) : '-'}</td>`;
        }
    });
    totalRowHtml += `<td style="font-weight: 700; background: linear-gradient(135deg, var(--current-primary), var(--current-secondary)); color: white; border-radius: 0 0 8px 0;">${formatCurrency(grandTotal)}</td>`;

    rows.push(`<tr class="total-row">${totalRowHtml}</tr>`);

    elements.tableBody.innerHTML = rows.join('');

    // Update header if filtering
    if (searchTerm) {
        const headerCells = ['<th>วันที่</th>'];
        filteredHeaders.forEach(header => {
            headerCells.push(`<th>${header}</th>`);
        });
        headerCells.push('<th>รวม</th>');
        elements.tableHeader.innerHTML = headerCells.join('');
    }
}

function filterTable() {
    const data = platformData[currentPlatform];
    if (!data) return;

    // Apply global date filter first
    const filteredData = getFilteredData(data);

    const searchTerm = elements.searchInput.value;
    renderTableBody(filteredData.dailyData, filteredData.headers, searchTerm);
}

function updateMonthFilter(data) {
    const months = new Set();
    data.dailyData.forEach(day => {
        const monthNum = day.date.split('-')[1];
        months.add(monthNum);
    });

    const monthNames = {
        '01': 'มกราคม', '02': 'กุมภาพันธ์', '03': 'มีนาคม',
        '04': 'เมษายน', '05': 'พฤษภาคม', '06': 'มิถุนายน',
        '07': 'กรกฎาคม', '08': 'สิงหาคม', '09': 'กันยายน',
        '10': 'ตุลาคม', '11': 'พฤศจิกายน', '12': 'ธันวาคม'
    };

    elements.monthFilter.innerHTML = '<option value="all">ทุกเดือน</option>';
    [...months].sort().forEach(month => {
        elements.monthFilter.innerHTML += `<option value="${month}">${monthNames[month] || month}</option>`;
    });
}

function updateCharts(data) {
    const colors = CONFIG.sheets[currentPlatform].colors;

    // Destroy existing charts
    if (charts.daily) charts.daily.destroy();
    if (charts.product) charts.product.destroy();

    // Daily Chart
    const dailyLabels = data.dailyData.map(d => formatDate(d.date));
    const dailyValues = data.dailyData.map(d => d.total);

    charts.daily = new Chart(elements.dailyChart, {
        type: 'line',
        data: {
            labels: dailyLabels,
            datasets: [{
                label: 'ค่าใช้จ่ายรายวัน',
                data: dailyValues,
                borderColor: colors.primary,
                backgroundColor: createGradient(elements.dailyChart, colors.gradient[0]),
                fill: true,
                tension: 0.4,
                pointRadius: 3,
                pointBackgroundColor: colors.primary,
                pointBorderColor: '#fff',
                pointBorderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: colors.primary,
                    borderWidth: 1,
                    callbacks: {
                        label: function (context) {
                            return formatCurrency(context.raw);
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        color: 'rgba(255, 255, 255, 0.1)'
                    },
                    ticks: {
                        color: '#a0a0b0',
                        maxTicksLimit: 10
                    }
                },
                y: {
                    grid: {
                        color: 'rgba(255, 255, 255, 0.1)'
                    },
                    ticks: {
                        color: '#a0a0b0',
                        callback: function (value) {
                            return formatCompactNumber(value);
                        }
                    }
                }
            },
            interaction: {
                mode: 'nearest',
                axis: 'x',
                intersect: false
            }
        }
    });

    // Product Chart (Top 10)
    const sortedProducts = Object.entries(data.productTotals)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const productLabels = sortedProducts.map(p => p[0]);
    const productValues = sortedProducts.map(p => p[1]);

    charts.product = new Chart(elements.productChart, {
        type: 'bar',
        data: {
            labels: productLabels,
            datasets: [{
                label: 'ค่าใช้จ่ายรวม',
                data: productValues,
                backgroundColor: createBarGradients(elements.productChart, colors, productValues.length),
                borderRadius: 8,
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    borderColor: colors.primary,
                    borderWidth: 1,
                    callbacks: {
                        label: function (context) {
                            return formatCurrency(context.raw);
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        color: 'rgba(255, 255, 255, 0.1)'
                    },
                    ticks: {
                        color: '#a0a0b0',
                        callback: function (value) {
                            return formatCompactNumber(value);
                        }
                    }
                },
                y: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#a0a0b0'
                    }
                }
            }
        }
    });
}

function createGradient(canvas, color) {
    const ctx = canvas.getContext('2d');
    const gradient = ctx.createLinearGradient(0, 0, 0, 300);
    gradient.addColorStop(0, color);
    gradient.addColorStop(1, 'rgba(0, 0, 0, 0)');
    return gradient;
}

function createBarGradients(canvas, colors, count) {
    const gradients = [];
    const ctx = canvas.getContext('2d');

    for (let i = 0; i < count; i++) {
        const gradient = ctx.createLinearGradient(0, 0, 200, 0);
        gradient.addColorStop(0, colors.primary);
        gradient.addColorStop(1, colors.secondary);
        gradients.push(gradient);
    }

    return gradients;
}

// Utility Functions
function formatCurrency(value) {
    return new Intl.NumberFormat('th-TH', {
        style: 'currency',
        currency: 'THB',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(value);
}

function formatCompactNumber(value) {
    if (value >= 1000000) {
        return (value / 1000000).toFixed(1) + 'M';
    } else if (value >= 1000) {
        return (value / 1000).toFixed(0) + 'K';
    }
    return value;
}

function formatDate(dateStr) {
    const [year, month, day] = dateStr.split('-');
    return `${day}/${month}`;
}

function updateLastUpdated() {
    const now = new Date();
    elements.lastUpdated.textContent = `อัพเดทล่าสุด: ${now.toLocaleString('th-TH')}`;
}

function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

