/**
 * Manual Trade Portfolio Dashboard — Live Google Sheets Integration
 * Sheet ID: 1TPs4U18P3MDsQTCOGP2n0OdrptCR10uZFI-ISGosJZ8
 */

const SHEET_ID = '1TPs4U18P3MDsQTCOGP2n0OdrptCR10uZFI-ISGosJZ8';
const AUTO_REFRESH_MS = 60 * 60 * 1000; // 60 minutes

const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_NAMES_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// ===== State =====
let allMonthsData = {}; // { 'Jul': [ { dateStr, dayName, grossPnl, grossRoi, expenses, netPnl, netRoi, capital } ] }
let allTradesChronological = [];
let selectedMonth = null;
let searchQuery = '';
let sortCol = 'date';
let sortDir = 'desc';
let autoRefreshTimer = null;
let charts = { bar: null, equity: null };

// ===== DOM Utility =====
const $ = id => document.getElementById(id);

// ===== Init =====
document.addEventListener('DOMContentLoaded', () => {
    initClock();
    populateMonthSelector();
    bindEvents();
    fetchAllData();
    startAutoRefresh();
});

// ===== Clock =====
function initClock() {
    const update = () => {
        const now = new Date();
        const formatted = now.toLocaleString('en-IN', {
            timeZone: 'Asia/Kolkata',
            weekday: 'short', day: 'numeric', month: 'short', year: 'numeric',
            hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true
        });
        if ($('liveClock')) $('liveClock').textContent = formatted;
    };
    update();
    setInterval(update, 1000);
}

// ===== Month Selector =====
function populateMonthSelector() {
    const sel = $('monthSelect');
    if (!sel) return;
    sel.innerHTML = '';

    const now = new Date();
    const currentMonthIdx = now.getMonth(); // 0-based index

    MONTH_NAMES.forEach((name, i) => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = `${name}-26`;
        if (i === currentMonthIdx) opt.selected = true;
        sel.appendChild(opt);
    });

    selectedMonth = MONTH_NAMES[currentMonthIdx];
}

// ===== Events =====
function bindEvents() {
    if ($('btnRefresh')) {
        $('btnRefresh').addEventListener('click', () => {
            $('btnRefresh').classList.add('spinning');
            fetchAllData().finally(() => {
                setTimeout(() => $('btnRefresh').classList.remove('spinning'), 600);
            });
        });
    }

    if ($('monthSelect')) {
        $('monthSelect').addEventListener('change', (e) => {
            selectedMonth = e.target.value;
            renderAll();
        });
    }

    if ($('searchInput')) {
        $('searchInput').addEventListener('input', (e) => {
            searchQuery = e.target.value.toLowerCase().trim();
            renderTable();
        });
    }

    if ($('btnExport')) $('btnExport').addEventListener('click', exportCSV);
    if ($('btnDismissError')) {
        $('btnDismissError').addEventListener('click', () => {
            $('errorBanner').style.display = 'none';
        });
    }

    // Keyboard shortcut (Ctrl+E or Cmd+E to export CSV)
    document.addEventListener('keydown', (e) => {
        if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'e') {
            e.preventDefault();
            exportCSV();
        }
    });
}

// ===== Data Parsing Helpers =====
function parseNum(val) {
    if (!val || val === '-' || val === '- ') return 0;
    if (typeof val === 'number') return val;
    return parseFloat(val.replace(/[₹,%\s]/g, '')) || 0;
}

function formatINR(num) {
    if (num === undefined || num === null || isNaN(num)) return '₹0';
    if (num === 0) return '₹0';
    const abs = Math.abs(num);
    const sign = num < 0 ? '-' : '';
    if (abs >= 10000000) return sign + '₹' + (abs / 10000000).toFixed(2) + ' Cr';
    if (abs >= 100000) return sign + '₹' + (abs / 100000).toFixed(2) + ' L';
    return sign + '₹' + abs.toLocaleString('en-IN', { maximumFractionDigits: 0 });
}

function formatPct(num) {
    if (num === undefined || num === null || isNaN(num)) return '0.00%';
    const sign = num > 0 ? '+' : '';
    return `${sign}${num.toFixed(2)}%`;
}

function parseCSV(text) {
    const rows = [];
    let current = '';
    let inQuotes = false;
    let row = [];

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (inQuotes) {
            if (ch === '"' && text[i + 1] === '"') {
                current += '"';
                i++;
            } else if (ch === '"') {
                inQuotes = false;
            } else {
                current += ch;
            }
        } else {
            if (ch === '"') {
                inQuotes = true;
            } else if (ch === ',') {
                row.push(current.trim());
                current = '';
            } else if (ch === '\n' || (ch === '\r' && text[i + 1] === '\n')) {
                row.push(current.trim());
                current = '';
                if (row.some(c => c !== '')) rows.push(row);
                row = [];
                if (ch === '\r') i++;
            } else {
                current += ch;
            }
        }
    }
    row.push(current.trim());
    if (row.some(c => c !== '')) rows.push(row);
    return rows;
}

// Parses date strings like "1/07/2026 - Wed" or "31/07/2026 - Fri" or "2026-07-31"
function parseDateStr(str) {
    if (!str) return null;
    const parts = str.split('-')[0].trim().split('/');
    if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);
        return new Date(year, month, day);
    }
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
}

// ===== Google Sheet Fetching =====
function buildCsvUrl(sheetName) {
    return `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
}

async function fetchSheet(sheetName) {
    const url = buildCsvUrl(sheetName);
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Failed to fetch ${sheetName}: ${resp.status}`);
    const text = await resp.text();
    return parseCSV(text);
}

async function fetchAllData() {
    if ($('loadingOverlay')) $('loadingOverlay').classList.remove('hidden');
    if ($('errorBanner')) $('errorBanner').style.display = 'none';

    allMonthsData = {};
    allTradesChronological = [];

    try {
        const monthFetchPromises = MONTH_NAMES.map(async (shortName, idx) => {
            const fullName = MONTH_NAMES_FULL[idx];
            const candidateNames = [
                `${shortName} 26 Algo PnL`,
                `${shortName} 26`,
                `${shortName} 2026`,
                `${fullName} 2026`,
                `${shortName}-26`
            ];

            for (const sheetName of candidateNames) {
                try {
                    const rows = await fetchSheet(sheetName);
                    if (rows && rows.length > 1) {
                        const parsed = parseSheetRows(rows, shortName);
                        if (parsed.trades.length > 0) {
                            allMonthsData[shortName] = parsed;
                            return;
                        }
                    }
                } catch (e) {
                    // Try next candidate sheet name silently
                }
            }
        });

        await Promise.all(monthFetchPromises);

        // Compile all trades into chronological order
        Object.values(allMonthsData).forEach(m => {
            if (m.trades) allTradesChronological.push(...m.trades);
        });

        allTradesChronological.sort((a, b) => a.parsedDate - b.parsedDate);

        // Calculate cumulative net equity and peak drawdown chronologically
        let cumulativeNetPnl = 0;
        let peakEquity = 0;
        let runningCapital = 100000;

        allTradesChronological.forEach(trade => {
            if (trade.capital > 0) runningCapital = trade.capital;
            cumulativeNetPnl += trade.netPnl;
            trade.cumulativeNetPnl = cumulativeNetPnl;
            trade.currentEquity = runningCapital + cumulativeNetPnl;
            
            if (cumulativeNetPnl > peakEquity) {
                peakEquity = cumulativeNetPnl;
            }
            trade.drawdown = cumulativeNetPnl - peakEquity; // <= 0
        });

        if ($('lastUpdated')) {
            $('lastUpdated').textContent = `Updated ${new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}`;
        }

        renderAll();
        showToast('Live sheet data synced', 'success');
    } catch (err) {
        console.error('Fetch error:', err);
        if ($('errorMessage')) {
            $('errorMessage').textContent = `Failed to load data: ${err.message}. Check network or sheet permissions.`;
        }
        if ($('errorBanner')) $('errorBanner').style.display = 'flex';
        showToast('Failed to fetch sheet data', 'error');
    } finally {
        if ($('loadingOverlay')) $('loadingOverlay').classList.add('hidden');
    }
}

function parseSheetRows(rows, monthShort) {
    let capital = 100000;
    const trades = [];

    // Right side capital check (Cols O, P, Q - index 14, 15, 16)
    rows.forEach(r => {
        for (let colIdx = 14; colIdx < r.length; colIdx++) {
            if (r[colIdx] && r[colIdx].trim().toLowerCase() === 'capital') {
                // Next row or column might have the capital
                for (let nextCol = colIdx; nextCol < r.length; nextCol++) {
                    const num = parseNum(r[nextCol]);
                    if (num > 10000) { capital = num; break; }
                }
            }
        }
    });

    // Header row is index 0: Date, Today's PnL, Today's ROI, Expenses, Expense %, Month PnL, Month ROI
    for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        const rawDate = r[0] ? r[0].trim() : '';

        // Ignore summary total row or empty rows
        if (!rawDate || rawDate.toLowerCase().startsWith('total')) continue;

        const parsedDate = parseDateStr(rawDate);
        if (!parsedDate) continue;

        const grossPnl = parseNum(r[1]);
        const grossRoi = parseNum(r[2]);
        const expenses = parseNum(r[3]);
        const expensePct = parseNum(r[4]);
        
        let netPnl = parseNum(r[5]);
        // If netPnl column is 0 or empty but gross/expenses exist, calculate it:
        if (netPnl === 0 && (grossPnl !== 0 || expenses !== 0)) {
            netPnl = grossPnl - expenses;
        }

        let netRoi = parseNum(r[6]);
        if (netRoi === 0 && grossRoi !== 0) {
            netRoi = grossRoi - expensePct;
        }

        // Extract day name from "1/07/2026 - Wed" or Date object
        let dayName = '';
        if (rawDate.includes('-')) {
            const split = rawDate.split('-');
            if (split.length > 1) dayName = split[1].trim();
        }
        if (!dayName) {
            dayName = parsedDate.toLocaleDateString('en-US', { weekday: 'short' });
        }

        trades.push({
            dateStr: rawDate,
            parsedDate: parsedDate,
            dayName: dayName,
            grossPnl: grossPnl,
            grossRoi: grossRoi,
            expenses: expenses,
            expensePct: expensePct,
            netPnl: netPnl,
            netRoi: netRoi,
            capital: capital,
            monthShort: monthShort,
            cumulativeNetPnl: 0,
            drawdown: 0
        });
    }

    return { capital, trades };
}

// ===== Auto Refresh =====
function startAutoRefresh() {
    if (autoRefreshTimer) clearInterval(autoRefreshTimer);
    autoRefreshTimer = setInterval(fetchAllData, AUTO_REFRESH_MS);
}

// ===== Main Render =====
function renderAll() {
    renderSummaryCards();
    renderStatsRow();
    renderDaywise();
    renderTable();
    renderAnalytics();
    renderHeatmap();
    renderMonthlyGrid();
    if ($('selectedMonthName')) $('selectedMonthName').textContent = `${selectedMonth}-26`;
}

// ===== Summary Cards =====
function renderSummaryCards() {
    const monthObj = allMonthsData[selectedMonth];
    const monthTrades = monthObj ? monthObj.trades : [];

    // Today's P&L (Last traded day in selected month)
    let todayPnl = 0;
    let todayDateStr = '';
    for (let i = monthTrades.length - 1; i >= 0; i--) {
        if (monthTrades[i].grossPnl !== 0 || monthTrades[i].expenses !== 0) {
            todayPnl = monthTrades[i].netPnl;
            todayDateStr = monthTrades[i].dateStr;
            break;
        }
    }

    // Monthly Net P&L
    let monthlyNetPnl = 0;
    monthTrades.forEach(t => { monthlyNetPnl += t.netPnl; });

    // Yearly Net P&L
    let yearlyNetPnl = 0;
    allTradesChronological.forEach(t => { yearlyNetPnl += t.netPnl; });

    // Capital Deployed
    let capitalDeployed = monthObj ? monthObj.capital : 100000;
    if (capitalDeployed <= 0) capitalDeployed = 100000;

    // Current Drawdown (drawdown on the last available traded day overall)
    let currentDD = 0;
    let maxDD = 0;
    if (allTradesChronological.length > 0) {
        const lastTrade = allTradesChronological[allTradesChronological.length - 1];
        currentDD = lastTrade.drawdown || 0;
    }
    allTradesChronological.forEach(t => {
        if (t.drawdown < maxDD) maxDD = t.drawdown;
    });

    // Update Cards DOM
    setCardValue('todayPnl', todayPnl, formatINR(todayPnl));
    if ($('todayPnlSub')) $('todayPnlSub').textContent = todayDateStr ? `on ${todayDateStr}` : 'No trading data';

    setCardValue('monthlyPnl', monthlyNetPnl, formatINR(monthlyNetPnl));
    if ($('monthlyPnlSub')) {
        const roi = capitalDeployed > 0 ? (monthlyNetPnl / capitalDeployed * 100) : 0;
        $('monthlyPnlSub').textContent = `${formatPct(roi)} of capital`;
    }

    setCardValue('yearlyPnl', yearlyNetPnl, formatINR(yearlyNetPnl));
    if ($('yearlyPnlSub')) $('yearlyPnlSub').textContent = `Jan to ${selectedMonth} 2026`;

    if ($('capitalDeployed')) $('capitalDeployed').textContent = formatINR(capitalDeployed);
    if ($('capitalSub')) {
        const activeDays = monthTrades.filter(t => t.grossPnl !== 0 || t.expenses !== 0).length;
        $('capitalSub').textContent = `${activeDays} active trading days`;
    }

    const currentDDPct = capitalDeployed > 0 ? (currentDD / capitalDeployed * 100) : 0;
    const maxDDPct = capitalDeployed > 0 ? (maxDD / capitalDeployed * 100) : 0;

    setCardValue('currentDD', currentDD, `${formatINR(currentDD)} (${currentDDPct.toFixed(2)}%)`);
    if ($('ddSub')) $('ddSub').textContent = `Max DD: ${formatINR(maxDD)} (${maxDDPct.toFixed(2)}%)`;
}

function setCardValue(elId, num, text) {
    const el = $(elId);
    if (!el) return;
    el.textContent = text;
    el.className = 'card-value ' + (num > 0 ? 'positive' : num < 0 ? 'negative' : '');
}

// ===== Stats Row =====
function renderStatsRow() {
    let winDays = 0;
    let lossDays = 0;
    let maxDDVal = 0;
    let totalExpenses = 0;
    let yearlyNetPnl = 0;
    let latestCapital = 100000;

    allTradesChronological.forEach(t => {
        if (t.netPnl > 0) winDays++;
        else if (t.netPnl < 0) lossDays++;
        if (t.drawdown < maxDDVal) maxDDVal = t.drawdown;
        totalExpenses += t.expenses;
        yearlyNetPnl += t.netPnl;
        if (t.capital > 0) latestCapital = t.capital;
    });

    const totalTradedDays = winDays + lossDays;
    const winRatio = totalTradedDays > 0 ? Math.round((winDays / totalTradedDays) * 100) : 0;
    const totalRoi = latestCapital > 0 ? (yearlyNetPnl / latestCapital * 100) : 0;
    const currentEquity = latestCapital + yearlyNetPnl;

    if ($('winDays')) {
        $('winDays').textContent = winDays;
        $('winDays').className = 'stat-value positive';
    }
    if ($('lossDays')) {
        $('lossDays').textContent = lossDays;
        $('lossDays').className = 'stat-value negative';
    }
    if ($('winRatio')) {
        $('winRatio').textContent = totalTradedDays > 0 ? `${winRatio}%` : '—';
        $('winRatio').className = 'stat-value ' + (winRatio >= 50 ? 'positive' : 'negative');
    }
    if ($('maxDD')) {
        $('maxDD').textContent = maxDDVal !== 0 ? formatINR(maxDDVal) : '—';
        $('maxDD').className = 'stat-value negative';
    }
    if ($('roi')) {
        $('roi').textContent = formatPct(totalRoi);
        $('roi').className = 'stat-value ' + (totalRoi >= 0 ? 'positive' : 'negative');
    }
    if ($('totalExpenses')) {
        $('totalExpenses').textContent = formatINR(totalExpenses);
        $('totalExpenses').className = 'stat-value';
    }
    if ($('currentEquity')) {
        $('currentEquity').textContent = formatINR(currentEquity);
        $('currentEquity').className = 'stat-value';
    }
}

// ===== Day-wise Performance (Mon-Fri) =====
function renderDaywise() {
    const grid = $('daywiseGrid');
    if (!grid) return;
    grid.innerHTML = '';

    const dayMap = { 'Mon': 'Monday', 'Tue': 'Tuesday', 'Wed': 'Wednesday', 'Thu': 'Thursday', 'Fri': 'Friday' };
    const dayTotals = { 'Monday': 0, 'Tuesday': 0, 'Wednesday': 0, 'Thursday': 0, 'Friday': 0 };

    allTradesChronological.forEach(t => {
        for (const [short, full] of Object.entries(dayMap)) {
            if (t.dayName.toLowerCase().includes(short.toLowerCase())) {
                dayTotals[full] += t.netPnl;
                break;
            }
        }
    });

    const values = Object.values(dayTotals);
    const maxAbs = Math.max(...values.map(v => Math.abs(v)), 1);

    Object.entries(dayTotals).forEach(([day, val]) => {
        const cls = val > 0 ? 'positive' : val < 0 ? 'negative' : '';
        const barPct = Math.round((Math.abs(val) / maxAbs) * 100);

        const card = document.createElement('div');
        card.className = 'daywise-card';
        card.innerHTML = `
            <div class="daywise-day">${day}</div>
            <div class="daywise-value ${cls}">${val !== 0 ? formatINR(val) : '—'}</div>
            <div class="daywise-bar">
                <div class="daywise-bar-fill ${val >= 0 ? 'positive' : 'negative'}" style="width: ${barPct}%"></div>
            </div>
        `;
        grid.appendChild(card);
    });
}

// ===== Daily Trade Log Table =====
function renderTable() {
    const thead = $('tableHead');
    const tbody = $('tableBody');
    if (!thead || !tbody) return;

    const monthObj = allMonthsData[selectedMonth];
    let monthTrades = monthObj ? [...monthObj.trades] : [];

    // Filter by search query
    if (searchQuery) {
        monthTrades = monthTrades.filter(t => 
            t.dateStr.toLowerCase().includes(searchQuery) ||
            t.dayName.toLowerCase().includes(searchQuery)
        );
    }

    // Build Table Header
    thead.innerHTML = '';
    const headerRow = document.createElement('tr');
    const cols = [
        { key: 'date', label: 'Date' },
        { key: 'day', label: 'Day' },
        { key: 'grossPnl', label: 'Gross P&L' },
        { key: 'expenses', label: 'Expenses' },
        { key: 'netPnl', label: 'Net P&L' },
        { key: 'netRoi', label: 'ROI %' },
        { key: 'cummPnl', label: 'Cumm Net P&L' },
        { key: 'drawdown', label: 'Drawdown' },
    ];

    cols.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col.label;
        th.dataset.key = col.key;
        if (col.key === sortCol) {
            th.classList.add(sortDir === 'asc' ? 'sort-asc' : 'sort-desc');
        }
        th.addEventListener('click', () => {
            if (sortCol === col.key) {
                sortDir = sortDir === 'asc' ? 'desc' : 'asc';
            } else {
                sortCol = col.key;
                sortDir = 'desc';
            }
            renderTable();
        });
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // Sort Trades
    monthTrades.sort((a, b) => {
        let va, vb;
        if (sortCol === 'date') { va = a.parsedDate; vb = b.parsedDate; }
        else if (sortCol === 'day') { va = a.dayName; vb = b.dayName; }
        else if (sortCol === 'cummPnl') { va = a.cumulativeNetPnl; vb = b.cumulativeNetPnl; }
        else { va = a[sortCol] || 0; vb = b[sortCol] || 0; }
        if (va < vb) return sortDir === 'asc' ? -1 : 1;
        if (va > vb) return sortDir === 'asc' ? 1 : -1;
        return 0;
    });

    // Build Table Body
    tbody.innerHTML = '';

    if (monthTrades.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = cols.length;
        td.textContent = 'No trade entries for this month';
        td.style.textAlign = 'center';
        td.style.padding = '36px';
        td.style.color = 'var(--text-muted)';
        tr.appendChild(td);
        tbody.appendChild(tr);
        return;
    }

    monthTrades.forEach(t => {
        const tr = document.createElement('tr');
        addCell(tr, t.dateStr, '');
        addCell(tr, t.dayName, '');
        addCell(tr, formatINR(t.grossPnl), t.grossPnl > 0 ? 'positive' : t.grossPnl < 0 ? 'negative' : 'zero');
        addCell(tr, t.expenses > 0 ? formatINR(t.expenses) : '—', 'zero');
        addCell(tr, formatINR(t.netPnl), t.netPnl > 0 ? 'positive' : t.netPnl < 0 ? 'negative' : 'zero');
        addCell(tr, formatPct(t.netRoi), t.netRoi > 0 ? 'positive' : t.netRoi < 0 ? 'negative' : 'zero');
        addCell(tr, formatINR(t.cumulativeNetPnl), t.cumulativeNetPnl > 0 ? 'positive' : t.cumulativeNetPnl < 0 ? 'negative' : 'zero');
        addCell(tr, t.drawdown < 0 ? formatINR(t.drawdown) : '—', t.drawdown < 0 ? 'negative' : 'zero');
        tbody.appendChild(tr);
    });
}

function addCell(tr, text, cls) {
    const td = document.createElement('td');
    td.textContent = text;
    if (cls) td.className = cls;
    tr.appendChild(td);
}

// ===== CSV Export =====
function exportCSV() {
    const monthObj = allMonthsData[selectedMonth];
    if (!monthObj || !monthObj.trades || monthObj.trades.length === 0) {
        showToast('No data to export for selected month', 'error');
        return;
    }

    const headers = ['Date', 'Day', 'Gross P&L', 'Expenses', 'Net P&L', 'ROI %', 'Cumm Net P&L', 'Drawdown'];
    const rows = monthObj.trades.map(t => [
        `"${t.dateStr}"`,
        `"${t.dayName}"`,
        t.grossPnl,
        t.expenses,
        t.netPnl,
        t.netRoi.toFixed(2),
        t.cumulativeNetPnl,
        t.drawdown
    ]);

    const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `Manual_Trading_${selectedMonth}_2026.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast(`Exported ${selectedMonth}-26 trades to CSV`, 'success');
}

// ===== Analytics Charts =====
function renderAnalytics() {
    if ($('barChartMonth')) $('barChartMonth').textContent = `${selectedMonth}-26`;
    if ($('equityChartMonth')) $('equityChartMonth').textContent = `${selectedMonth}-26`;

    const monthObj = allMonthsData[selectedMonth];
    const monthTrades = monthObj ? [...monthObj.trades] : [];
    monthTrades.sort((a, b) => a.parsedDate - b.parsedDate);

    const labels = monthTrades.map(t => t.dateStr.split('-')[0].trim());
    const netPnls = monthTrades.map(t => t.netPnl);
    const cummPnls = monthTrades.map(t => t.cumulativeNetPnl);

    // 1. Daily Net P&L Bar Chart
    const barCtx = $('dailyBarChart');
    if (barCtx) {
        if (charts.bar) charts.bar.destroy();
        charts.bar = new Chart(barCtx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Net P&L (₹)',
                    data: netPnls,
                    backgroundColor: netPnls.map(v => v >= 0 ? '#10b981' : '#ef4444'),
                    borderRadius: 6,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => `Net P&L: ${formatINR(ctx.raw)}`
                        }
                    }
                },
                scales: {
                    x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#64748b' } },
                    y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#64748b', callback: v => formatINR(v) } }
                }
            }
        });
    }

    // 2. Cumulative Equity Line Chart
    const equityCtx = $('equityLineChart');
    if (equityCtx) {
        if (charts.equity) charts.equity.destroy();
        charts.equity = new Chart(equityCtx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Cumulative Net P&L',
                    data: cummPnls,
                    borderColor: '#6366f1',
                    backgroundColor: 'rgba(99, 102, 241, 0.12)',
                    fill: true,
                    tension: 0.3,
                    pointRadius: 4,
                    pointHoverRadius: 6,
                    pointBackgroundColor: '#818cf8'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => `Cumm Net P&L: ${formatINR(ctx.raw)}`
                        }
                    }
                },
                scales: {
                    x: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#64748b' } },
                    y: { grid: { color: 'rgba(255,255,255,0.05)' }, ticks: { color: '#64748b', callback: v => formatINR(v) } }
                }
            }
        });
    }
}

// ===== 2026 Win/Loss Calendar Heatmap =====
function renderHeatmap() {
    const container = $('calendarHeatmap');
    if (!container) return;
    container.innerHTML = '';

    // Create a date lookup map for all trades
    const dateMap = {};
    allTradesChronological.forEach(t => {
        const year = t.parsedDate.getFullYear();
        const month = t.parsedDate.getMonth();
        const date = t.parsedDate.getDate();
        const key = `${year}-${month}-${date}`;
        dateMap[key] = t;
    });

    MONTH_NAMES.forEach((mName, mIdx) => {
        const col = document.createElement('div');
        col.className = 'heatmap-month-col';

        const label = document.createElement('div');
        label.className = 'heatmap-month-label';
        label.textContent = mName;
        col.appendChild(label);

        const daysGrid = document.createElement('div');
        daysGrid.className = 'heatmap-days-grid';

        const daysInMonth = new Date(2026, mIdx + 1, 0).getDate();

        for (let day = 1; day <= daysInMonth; day++) {
            const box = document.createElement('div');
            box.className = 'day-box';
            
            const key = `2026-${mIdx}-${day}`;
            const trade = dateMap[key];

            if (trade && (trade.grossPnl !== 0 || trade.expenses !== 0)) {
                const val = trade.netPnl;
                let colorClass = 'scale-neutral';
                if (val > 2000) colorClass = 'scale-profit-3';
                else if (val > 500) colorClass = 'scale-profit-2';
                else if (val > 0) colorClass = 'scale-profit-1';
                else if (val < -2000) colorClass = 'scale-loss-3';
                else if (val < -500) colorClass = 'scale-loss-2';
                else if (val < 0) colorClass = 'scale-loss-1';

                box.classList.add(colorClass);
                box.title = `${trade.dateStr}: Net P&L ${formatINR(val)}`;
            } else {
                box.classList.add('scale-neutral');
                box.title = `${day} ${mName} 2026: No Trade Data`;
            }

            daysGrid.appendChild(box);
        }

        col.appendChild(daysGrid);
        container.appendChild(col);
    });
}

// ===== Monthly Summary Grid =====
function renderMonthlyGrid() {
    const grid = $('heatmapGrid');
    if (!grid) return;
    grid.innerHTML = '';

    MONTH_NAMES.forEach(mName => {
        const monthObj = allMonthsData[mName];
        const trades = monthObj ? monthObj.trades : [];
        let netPnl = 0;
        let expenses = 0;
        let winCount = 0;
        let lossCount = 0;

        trades.forEach(t => {
            netPnl += t.netPnl;
            expenses += t.expenses;
            if (t.netPnl > 0) winCount++;
            else if (t.netPnl < 0) lossCount++;
        });

        const statusClass = netPnl > 0 ? 'profit' : netPnl < 0 ? 'loss' : 'neutral';
        const statusText = netPnl > 0 ? 'PROFIT' : netPnl < 0 ? 'LOSS' : 'NEUTRAL';

        const card = document.createElement('div');
        card.className = 'month-summary-card';
        card.innerHTML = `
            <div class="month-card-header">
                <span class="month-card-name">${mName} 2026</span>
                <span class="month-card-status ${statusClass}">${statusText}</span>
            </div>
            <div class="month-card-pnl ${statusClass}">${trades.length > 0 ? formatINR(netPnl) : '—'}</div>
            <div class="month-card-details">
                <span>Wins: ${winCount} | Losses: ${lossCount}</span>
                <span>Exp: ${formatINR(expenses)}</span>
            </div>
        `;
        grid.appendChild(card);
    });
}

// ===== Toast Notification System =====
function showToast(message, type = 'info') {
    const container = $('toastContainer');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
        <span>${type === 'success' ? '✓' : type === 'error' ? '✕' : 'ℹ'}</span>
        <span>${message}</span>
    `;

    container.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
    }, 3500);
}
