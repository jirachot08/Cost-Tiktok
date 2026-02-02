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
        }
    }
};

// State
let currentPlatform = 'summary';
let platformData = {
    summary: null,
    tiktok: null,
    facebook: null
};
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
    startDate: document.getElementById('startDate'),
    endDate: document.getElementById('endDate'),
    applyDateRange: document.getElementById('applyDateRange'),
    selectedDateRange: document.getElementById('selectedDateRange'),
    clearDateFilter: document.getElementById('clearDateFilter')
};

// Global filter state
let globalDateFilter = {
    startDate: null,
    endDate: null
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

    // Date range filter
    elements.applyDateRange.addEventListener('click', applyDateRangeFilter);
    elements.clearDateFilter.addEventListener('click', clearGlobalDateFilter);
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

function applyDateRangeFilter() {
    const startDateValue = elements.startDate.value;
    const endDateValue = elements.endDate.value;

    if (!startDateValue && !endDateValue) {
        alert('กรุณาเลือกวันที่อย่างน้อย 1 วัน');
        return;
    }

    // If only start date, use it as both start and end
    if (startDateValue && !endDateValue) {
        globalDateFilter.startDate = startDateValue;
        globalDateFilter.endDate = startDateValue;
    }
    // If only end date, use it as both start and end
    else if (!startDateValue && endDateValue) {
        globalDateFilter.startDate = endDateValue;
        globalDateFilter.endDate = endDateValue;
    }
    // Both dates provided
    else {
        // Ensure start is before or equal to end
        if (startDateValue > endDateValue) {
            globalDateFilter.startDate = endDateValue;
            globalDateFilter.endDate = startDateValue;
        } else {
            globalDateFilter.startDate = startDateValue;
            globalDateFilter.endDate = endDateValue;
        }
    }

    updateSelectedDateRangeDisplay();
    applyGlobalDateFilter();
}

function clearGlobalDateFilter() {
    globalDateFilter.startDate = null;
    globalDateFilter.endDate = null;

    elements.startDate.value = '';
    elements.endDate.value = '';

    updateSelectedDateRangeDisplay();

    // Refresh with all data
    const data = platformData[currentPlatform];
    if (data) {
        updateSummaryCards(data);
        updateTopProducts(data, 'all');
        updateCharts(data);
        filterTable();
    }
}

function initializeDateInputs(data) {
    if (!data || !data.dailyData || data.dailyData.length === 0) return;

    // Get min and max dates from data
    const dates = data.dailyData.map(d => d.date).sort();
    const minDate = dates[0];
    const maxDate = dates[dates.length - 1];

    // Set min/max attributes on date inputs
    elements.startDate.min = minDate;
    elements.startDate.max = maxDate;
    elements.endDate.min = minDate;
    elements.endDate.max = maxDate;
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
    document.body.classList.remove('tiktok-theme', 'facebook-theme');
    document.body.classList.add(CONFIG.sheets[platform].theme);

    // Update active tab
    elements.tabButtons.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.platform === platform);
    });

    // Update dashboard
    updateDashboard();
}

function updateDashboard() {
    const data = platformData[currentPlatform];
    if (!data) return;

    // Initialize date range inputs
    initializeDateInputs(data);
    updateSelectedDateRangeDisplay();

    updateSummaryCards(data);
    updateProductMonthFilter(data);
    updateTopProducts(data);
    updateTable(data);
    updateCharts(data);
    updateMonthFilter(data);
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
