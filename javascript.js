// ====================================================================
// 📊 GLOBAL DATA STATE UTILITIES
// ====================================================================
let globalData = []; 
let fileType = ""; 
let _skipChartRedraw = false; // perf: skip expensive chart redraw on row toggles
let genericHeaders = []; // Para sa dynamic generic mapping
let parsedEmployeesData = {};
let currentActiveEmployee = "";
let manualAttendanceLogs = []; // {id, name, date, time, status} — source of truth ng Live Log Feed
let manualLogsCounter = 0;
let _manualLogIdSeq = 0; // unique id generator para sa manual logs (di nagbabago kahit may delete)
let chartInstancePie = null;
let chartInstanceBar3D = null;
// Live Excel Modal State
let _liveExcelWorkbook = null;
let _liveExcelSelectedSheet = null;
let _liveExcelFile = null;
// 🔴 LIVE LOCAL FILE WATCH STATE — File System Access API (Chrome/Edge lang)
let _pendingLiveFileHandle = null; // di pa committed na handle, galing sa picker, naghihintay sa "I-load sa Dashboard"
let _liveFileHandle = null;        // ACTIVE na FileSystemFileHandle na pinapanood (null = walang live sync)
let _liveFileLastModified = null;  // huling nakitang lastModified timestamp, para malaman kung nag-iba
let _livePollIntervalId = null;    // setInterval id ng polling loop
let _liveFileFailCount = 0;        // bilang ng magkakasunod na error bago tuluyang itigil
const LIVE_POLL_MS = 4000;         // bawat ilang segundo titingnan kung na-save ulit ang file
const DAYS_OF_WEEK = ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"];

// ====================================================================
// 💾 PERSISTENCE LAYER — auto-save sa browser (localStorage) para hindi
// nawawala ang na-upload na data kapag nag-reload ng page. Ang "Tanggalin
// ang Data" button ang gagamitin para sadyang burahin ang naka-save na data.
// ====================================================================
const STORAGE_KEY = 'mtoSystemMonitoring_savedState_v1';
let _saveDebounceTimer = null;

function saveAppState() {
    try {
        const isAttendanceTabActive = !document.getElementById('attendanceView').classList.contains('hidden');
        const titleEl = document.getElementById('tableTitle');
        const state = {
            globalData,
            fileType,
            genericHeaders,
            currentFileTitle: titleEl ? titleEl.innerText : '',
            parsedEmployeesData,
            currentActiveEmployee,
            manualAttendanceLogs,
            manualLogsCounter,
            _manualLogIdSeq,
            activeView: isAttendanceTabActive ? 'attendance' : 'dashboard'
        };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
        // Karaniwang dahilan: puno na ang storage ng browser (mahigpit ang limitasyon, ~5-10MB).
        console.warn('Hindi na-save ang data sa browser storage:', e);
    }
}

// Debounced version — gamit para sa madalas na pag-type (hal. number inputs)
// para hindi mag-save sa kada letrang i-type ng user.
function saveAppStateDebounced() {
    clearTimeout(_saveDebounceTimer);
    _saveDebounceTimer = setTimeout(saveAppState, 400);
}

function loadAppState() {
    let raw;
    try {
        raw = localStorage.getItem(STORAGE_KEY);
    } catch (e) {
        console.warn('Hindi ma-access ang browser storage:', e);
        return;
    }
    if (!raw) return;

    let state;
    try {
        state = JSON.parse(raw);
    } catch (e) {
        console.warn('Sira ang naka-save na data, lalaktawan:', e);
        return;
    }

    globalData = Array.isArray(state.globalData) ? state.globalData : [];
    fileType = state.fileType || "";
    genericHeaders = Array.isArray(state.genericHeaders) ? state.genericHeaders : [];
    parsedEmployeesData = state.parsedEmployeesData && typeof state.parsedEmployeesData === 'object' ? state.parsedEmployeesData : {};
    currentActiveEmployee = state.currentActiveEmployee || "";
    manualAttendanceLogs = Array.isArray(state.manualAttendanceLogs) ? state.manualAttendanceLogs : [];
    manualLogsCounter = Number(state.manualLogsCounter) || 0;
    _manualLogIdSeq = Number(state._manualLogIdSeq) || 0;

    // Ibalik ang pamagat ng file sa header ng table
    if (state.currentFileTitle) {
        const titleEl = document.getElementById('tableTitle');
        if (titleEl) titleEl.innerText = state.currentFileTitle;
    }

    // Ibalik ang Dashboard Monitoring data (kung meron)
    if (globalData.length > 0) {
        document.getElementById('searchBar').disabled = false;
        document.getElementById('searchBar').placeholder = "Mag-type rito para mag-filter ng data...";
        buildDashboard(globalData);
    }

    // Ibalik ang Timesheet / Attendance file data (kung meron)
    if (Object.keys(parsedEmployeesData).length > 0) {
        refreshEditorUI();
    }

    // Ibalik ang Manual Attendance Logger feed (kung meron)
    renderManualLogs();

    // Ibalik ang huling aktibong tab (Dashboard o Attendance)
    if (state.activeView === 'attendance') {
        document.getElementById('btnAttendance').click();
    }
}

// Tinatanggal lahat ng naka-save na CSV/Excel data (dashboard, timesheet, at manual logs).
// Ginagamit ng "🗑️ Tanggalin ang Data" button sa sidebar.
function clearAppState() {
    const meronLaman = globalData.length > 0 || Object.keys(parsedEmployeesData).length > 0 || manualAttendanceLogs.length > 0;
    const msg = meronLaman
        ? "Sigurado ka bang nais mong tanggalin ang LAHAT ng na-upload na CSV/Excel data, kasama ang manual attendance logs? Hindi na ito mababawi."
        : "Walang naka-save na data sa ngayon. Sigurado ka bang nais mong magpatuloy?";
    if (!confirm(msg)) return;

    try {
        localStorage.removeItem(STORAGE_KEY);
    } catch (e) {
        console.warn('Hindi na-clear ang browser storage:', e);
        alert('⚠️ Hindi na-clear ang data. Posibleng naka-disable ang storage ng browser mo (hal. private/incognito mode o mahigpit na privacy settings).');
        return;
    }
    alert('✅ Tagumpay! Nabura na ang lahat ng naka-save na CSV/Excel data. Mag-re-reload na ngayon ang page.');
    // Reload para sigurado ang malinis na pagsisimula ng buong UI (charts, table, forms, atbp.)
    location.reload();
}

// ====================================================================
// 🌓 THEME CONTROL LAYER
// ====================================================================
function toggleTheme() {
    const html = document.documentElement;
    const btnIcon = document.getElementById('themeBtnIcon');
    const btnText = document.getElementById('themeBtnText');
    
    if (html.classList.contains('dark')) {
        html.classList.remove('dark');
        html.classList.add('light');
        btnIcon.innerText = "🌙";
        btnText.innerText = "LIGHT MODE";
        localStorage.setItem('theme', 'light');
        updateHighchartsTheme(false);
    } else {
        html.classList.remove('light');
        html.classList.add('dark');
        btnIcon.innerText = "☀️";
        btnText.innerText = "DARK MODE";
        localStorage.setItem('theme', 'dark');
        updateHighchartsTheme(true);
    }
}

function updateHighchartsTheme(isDark) {
    const color = isDark ? '#cbd5e1' : '#334155';
    if (window.Highcharts) {
        Highcharts.setOptions({
            legend: { itemStyle: { color: color } },
            xAxis: { labels: { style: { color: color } } },
            yAxis: { labels: { style: { color: color } } }
        });
    }
    if(globalData.length > 0 && fileType !== "attendance") {
        document.getElementById('searchBar').dispatchEvent(new Event('input'));
    }
}

// ====================================================================
// 🗺️ TAB NAVIGATION CONTROL INTERFACE
// ====================================================================
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('btnDashboard').addEventListener('click', function() {
        this.className = "w-full flex items-center space-x-2 py-2.5 px-4 rounded-xl bg-orange-600 text-white font-bold transition-all shadow-md shadow-orange-600/20 cursor-pointer";
        document.getElementById('btnAttendance').className = "w-full flex items-center space-x-2 py-2.5 px-4 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-bold transition-all hover:bg-orange-600 hover:text-white cursor-pointer";
        document.getElementById('dashboardView').classList.remove('hidden');
        document.getElementById('attendanceView').classList.add('hidden');
        saveAppState();
    });

    document.getElementById('btnAttendance').addEventListener('click', function() {
        this.className = "w-full flex items-center space-x-2 py-2.5 px-4 rounded-xl bg-orange-600 text-white font-bold transition-all shadow-md shadow-orange-600/20 cursor-pointer";
        document.getElementById('btnDashboard').className = "w-full flex items-center space-x-2 py-2.5 px-4 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-bold transition-all hover:bg-orange-600 hover:text-white cursor-pointer";
        document.getElementById('attendanceView').classList.remove('hidden');
        document.getElementById('dashboardView').classList.add('hidden');
        saveAppState();
    });

    const savedTheme = localStorage.getItem('theme') || 'light';
    const html = document.documentElement;
    if(savedTheme === 'dark') {
        html.classList.remove('light');
        html.classList.add('dark');
        document.getElementById('themeBtnIcon').innerText = "☀️";
        document.getElementById('themeBtnText').innerText = "LIGHT MODE";
        updateHighchartsTheme(true);
    } else {
        document.getElementById('themeBtnIcon').innerText = "🌙";
        document.getElementById('themeBtnText').innerText = "DARK MODE";
        updateHighchartsTheme(false);
    }
    
    const exportBtn = document.getElementById('btnExportDashboard');
    if (exportBtn) {
        exportBtn.addEventListener('click', handleDashboardExport);
    }

    const clearDataBtn = document.getElementById('clearDataBtn');
    if (clearDataBtn) {
        clearDataBtn.addEventListener('click', clearAppState);
    }


    // 📂 Live Excel Modal — minsan-lang-mag-upload (snapshot) na file input handler
    const liveExcelInput = document.getElementById('liveExcelFileInput');
    if (liveExcelInput) {
        liveExcelInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            _pendingLiveFileHandle = null; // hindi ito live-link — plain snapshot lang
            _liveExcelFile = file;
            const reader = new FileReader();
            reader.onload = function(evt) {
                const data = new Uint8Array(evt.target.result);
                _liveExcelWorkbook = XLSX.read(data, { type: 'array', cellDates: true });
                document.getElementById('liveExcelFileName').innerHTML =
                    '<span class="text-emerald-600 dark:text-emerald-400 font-black">✅ ' + file.name + '</span> <span class="text-slate-400 font-normal">(isang beses lang)</span>';
                renderSheetButtons(_liveExcelWorkbook.SheetNames);
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // 🔁 Ibalik ang dating na-upload na data (kung meron) — dito malulutas
    // ang problema na nawawala ang laman ng website pag nag-reload.
    loadAppState();
});

// ====================================================================
// 🛠️ DYNAMIC PARSING DATA UTILITIES
// ====================================================================
function parseCSVLine(text) {
    let insideQuote = false, entries = [''], index = 0;
    for (let i = 0; i < text.length; i++) {
        let char = text[i];
        if (char === '"') { insideQuote = !insideQuote; } 
        else if (char === ',' && !insideQuote) { entries[++index] = ''; } 
        else { entries[index] += char; }
    }
    return entries;
}

function cleanNumber(val) {
    if (!val) return 0;
    let clean = val.replace(/"/g, '').replace(/,/g, '').trim();
    return parseInt(clean) || 0;
}

// ====================================================================
// 📅 LATEST DATE CARD UPDATER + DATE PICKER SEARCH
// ====================================================================

// Global: all unique YYYY-MM-DD dates parsed from current data
let _allParsedDates = {};  // map: "YYYY-MM-DD" -> count of records

// Helper: parse any date string into a YYYY-MM-DD string, or null
function parseDateToISO(str) {
    if (!str || str === '-' || str === '' || str === 'N/A') return null;
    let s = String(str).trim();
    if (!s) return null;
    // Already YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    // ISO with time component: 2026-06-15T00:00:00.000Z
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.split('T')[0];
    // Excel serial number (pure 4-5 digit number, e.g. 46180 = 2026-06-15)
    // Dates in modern Excel are typically in the 40000-80000 range
    if (/^\d{4,5}$/.test(s)) {
        let serial = parseInt(s);
        // Convert Excel serial to Unix timestamp: (serial - 25569) days since Jan 1 1970
        let d = new Date((serial - 25569) * 86400 * 1000);
        if (!isNaN(d.getTime())) {
            let y = d.getUTCFullYear();
            if (y >= 1990 && y <= 2100) {
                return y + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' +
                    String(d.getUTCDate()).padStart(2, '0');
            }
        }
    }
    // Try native Date parse (handles ISO, RFC2822, many locale formats)
    let d = new Date(s);
    if (!isNaN(d.getTime()) && d.getFullYear() >= 1990 && d.getFullYear() <= 2100) {
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' +
            String(d.getDate()).padStart(2, '0');
    }
    // Try MM/DD/YYYY, M/D/YYYY, or DD/MM/YYYY (Filipino/European format)
    let slashParts = s.split('/');
    if (slashParts.length === 3) {
        let raw0 = slashParts[0].trim(), raw1 = slashParts[1].trim(), raw2 = slashParts[2].trim();
        let y = raw2.length === 2 ? '20' + raw2 : raw2;
        // MM/DD/YYYY attempt (default)
        let d1 = new Date(`${y}-${raw0.padStart(2,'0')}-${raw1.padStart(2,'0')}`);
        if (!isNaN(d1.getTime()) && d1.getFullYear() >= 1990 && d1.getFullYear() <= 2100) {
            return d1.getFullYear() + '-' + String(d1.getMonth() + 1).padStart(2,'0') + '-' +
                String(d1.getDate()).padStart(2,'0');
        }
        // DD/MM/YYYY fallback (when first part > 12, it must be the day)
        if (parseInt(raw0) > 12) {
            let d2 = new Date(`${y}-${raw1.padStart(2,'0')}-${raw0.padStart(2,'0')}`);
            if (!isNaN(d2.getTime()) && d2.getFullYear() >= 1990 && d2.getFullYear() <= 2100) {
                return d2.getFullYear() + '-' + String(d2.getMonth() + 1).padStart(2,'0') + '-' +
                    String(d2.getDate()).padStart(2,'0');
            }
        }
    }
    return null;
}

// Get ISO date string from a row depending on fileType
function getISODateFromRow(row) {
    // Direct field mappings for known file types
    if (fileType === 'raw' || fileType === 'crossdock') {
        let d = parseDateToISO(row.date);
        if (d) return d;
    }
    if (fileType === 'generic') {
        for (let h of genericHeaders) {
            if (h.toLowerCase().includes('date') || h.toLowerCase().includes('time')) {
                let d = parseDateToISO(row[h]);
                if (d) return d;
            }
        }
    }
    // Universal fallback: scan all string fields whose KEY NAME suggests a date.
    // This ensures ANY file type (progress, device, custom) works if it has a date column.
    const skipKeys = new Set(['_idx', 'qty', 'allocated', 'forecast', 'shipped',
                               'discrepancy', 'percentage', 'specRows', 'specCols',
                               'dailyRate', 'otHours', 'allowance']);
    const dateHints = ['date', 'time', 'dt', 'day', 'when', 'created', 'updated',
                       'ship', 'receiv', 'deliver', 'dispatch', 'sent', 'arrival'];
    for (let key of Object.keys(row)) {
        if (skipKeys.has(key)) continue;
        let keyLow = key.toLowerCase();
        if (!dateHints.some(hint => keyLow.includes(hint))) continue;
        let val = String(row[key] || '').trim();
        if (!val || val === '-' || val.length < 5) continue;
        let d = parseDateToISO(val);
        if (d && d >= '1990-01-01' && d <= '2100-12-31') return d;
    }
    return null;
}

function updateLatestDateCard(data) {
    // Build a fresh date map from ALL globalData (not just filtered)
    _allParsedDates = {};
    globalData.forEach(row => {
        let iso = getISODateFromRow(row);
        if (iso) _allParsedDates[iso] = (_allParsedDates[iso] || 0) + 1;
    });

    const allDates = Object.keys(_allParsedDates).sort();
    const picker = document.getElementById('datePicker');

    if (allDates.length === 0) {
        // No date column in this file type
        document.getElementById('latestDateLabel').innerText = '—';
        document.getElementById('latestDateCount').innerText = '—';
        picker.disabled = true;
        picker.value = '';
        document.getElementById('clearDateBtn').classList.add('hidden');
        return;
    }

    // Enable date picker — max is always TODAY so user can always pick the current date
    const todayISO = new Date().toISOString().split('T')[0];
    picker.disabled = false;
    picker.min = allDates[0];                 // earliest date in the file
    picker.max = todayISO;                    // always allow picking today

    // Default: prefer today if file has data for today, otherwise use latest in file
    const latestISO = allDates[allDates.length - 1];
    const preferredISO = _allParsedDates[todayISO] ? todayISO : latestISO;
    if (!picker.value) {
        picker.value = preferredISO;
    }

    // Display the currently selected date info
    const displayISO = picker.value || preferredISO;
    const displayCount = _allParsedDates[displayISO] || 0;

    // Format: M/D/YYYY
    const parts = displayISO.split('-');
    const displayLabel = parseInt(parts[1]) + '/' + parseInt(parts[2]) + '/' + parts[0];
    document.getElementById('latestDateLabel').innerText = displayLabel;
    document.getElementById('latestDateCount').innerText = displayCount > 0
        ? displayCount.toLocaleString() + ' record/box' + (displayCount !== 1 ? 'es' : '')
        : '0 — Walang record sa date na ito';
}

// Called when user picks a date from the date picker
function searchByDate(isoDate) {
    if (!isoDate || !_allParsedDates) return;

    const count = _allParsedDates[isoDate] || 0;
    const parts = isoDate.split('-');
    const displayLabel = parseInt(parts[1]) + '/' + parseInt(parts[2]) + '/' + parts[0];

    document.getElementById('latestDateLabel').innerText = displayLabel;
    document.getElementById('latestDateCount').innerText = count > 0
        ? count.toLocaleString() + ' record/box' + (count !== 1 ? 'es' : '')
        : '0 — Walang record sa date na ito';

    // Show clear button
    document.getElementById('clearDateBtn').classList.remove('hidden');

    // Filter the table to show only rows matching this date
    if (count > 0) {
        const filtered = globalData.filter(row => getISODateFromRow(row) === isoDate);
        _skipChartRedraw = true;
        buildDashboard(filtered);
        _skipChartRedraw = false;
    } else {
        // No records — show empty table
        _skipChartRedraw = true;
        buildDashboard([]);
        _skipChartRedraw = false;
    }
}

// Clear the date filter and restore full data
function clearDateSearch() {
    const picker = document.getElementById('datePicker');
    picker.value = '';
    document.getElementById('clearDateBtn').classList.add('hidden');
    // Restore full table — updateLatestDateCard inside buildDashboard
    // will re-apply the correct default (today if data exists, else latest in file)
    buildDashboard(globalData);
}

// ====================================================================
// 📊 SUMMARY CARDS UPDATER (Pending, Done, Local, HighValue)
// Returns the computed totals so the pie chart uses the EXACT same numbers.
// ====================================================================
function updateSummaryCards(data) {
    let totalPending = 0, totalDone = 0, totalLocal = 0, totalHighValue = 0;

    if (fileType === "raw") {
        data.forEach(row => {
            let statusLow = row.status.toLowerCase();
            let isDone = statusLow.includes('done') || (statusLow.includes('ship') && !statusLow.includes('pending'));
            if (isDone) totalDone++; else totalPending++;
            if (row.type === "HighValue") totalHighValue++; else totalLocal++;
        });
    } else if (fileType === "progress") {
        data.forEach(row => {
            let isDone = row.percentage >= 100 || row.discrepancy <= 0;
            if (isDone) totalDone++; else totalPending++;
        });
    } else if (fileType === "crossdock") {
        data.forEach(row => {
            let isDone = row.statusReceived.toLowerCase().includes('done');
            if (isDone) totalDone++; else totalPending++;
        });
    } else if (fileType === "device") {
        data.forEach(row => {
            let statusLow = (row.status || '').toLowerCase();
            if (statusLow === 'done') totalDone++; else totalPending++;
        });
    } else if (fileType === "generic") {
        totalDone = 0; totalPending = data.length;
    } else if (fileType === "cancelled") {
        // done = rows with picker assigned; pending = rows without picker
        data.forEach(row => {
            if (row.picker && row.picker !== '') totalDone++; else totalPending++;
        });
    }

    document.getElementById('summaryPending').innerText = totalPending.toLocaleString();
    document.getElementById('summaryDone').innerText = totalDone.toLocaleString();
    document.getElementById('summaryHighValue').innerText = (fileType === "raw") ? totalHighValue.toLocaleString() : '—';

    // Return values so pie chart uses EXACTLY the same numbers as the cards
    return { totalPending, totalDone, totalLocal, totalHighValue };
}


function toggleType(index) {
    let row = globalData[index];
    row.type = (row.type === "Local") ? "HighValue" : "Local";
    _skipChartRedraw = true;
    let query = document.getElementById('searchBar').value;
    if (query) { document.getElementById('searchBar').dispatchEvent(new Event('input')); } else { buildDashboard(globalData); }
    _skipChartRedraw = false;
    saveAppState();
}

function buildDashboard(data) {
    const tbody = document.getElementById('tableBody');
    const thead = document.getElementById('tableHeader');
    
    if (data.length === 0) { 
        tbody.innerHTML = `<tr><td class="py-8 text-center text-slate-400">Walang natagpuang tugma.</td></tr>`; 
        return; 
    }
    
    let doneCount = 0, pendingCount = 0;
    let whMap = {}, batchMap = {};
    let sumForecast = 0, sumAllocated = 0, sumShipped = 0, sumDiscrepancy = 0;
    let htmlBuffer = [];
    
    const textClass = "text-slate-800 dark:text-slate-200";
    const fontMuted = "text-slate-500 dark:text-slate-400";
    
    if (fileType === "progress") {
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400"><th class="py-3.5 px-4">WW</th><th class="py-3.5 px-4">Month</th><th class="py-3.5 px-4">Warehouse</th><th class="py-3.5 px-4">MTO Batch Name</th><th class="py-3.5 px-4 text-right">Forecast</th><th class="py-3.5 px-4 text-right">Allocated</th><th class="py-3.5 px-4 text-right">Total Shipped</th><th class="py-3.5 px-4 text-right">Discrepancy</th><th class="py-3.5 px-4 text-center">Status</th><th class="py-3.5 px-4 text-center">Percentage</th></tr>`;
        
        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            sumForecast += row.forecast; sumAllocated += row.allocated; sumShipped += row.shipped; sumDiscrepancy += row.discrepancy;
            if (row.wh) whMap[row.wh] = (whMap[row.wh] || 0) + row.shipped;
            if (row.batch) batchMap[row.batch] = (batchMap[row.batch] || 0) + row.shipped;
            
            let isDone = row.percentage >= 100 || row.discrepancy <= 0;
            if (isDone) doneCount++; else pendingCount++;
            
            let origIndex = row._idx;
            let statusBadge = isDone ? 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>` : 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;
            
            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40"><td class="py-2.5 px-4 font-mono text-xs ${fontMuted}">${row.ww}</td><td class="py-2.5 px-4 font-bold ${textClass}">${row.month}</td><td class="py-2.5 px-4 font-black text-orange-600 dark:text-orange-400">${row.wh}</td><td class="py-2.5 px-4 font-medium ${textClass} truncate max-w-[150px]">${row.batch}</td><td class="py-2.5 px-4 text-right font-mono ${fontMuted}">${row.forecast.toLocaleString()}</td><td class="py-2.5 px-4 text-right font-mono text-blue-600 dark:text-blue-400">${row.allocated.toLocaleString()}</td><td class="py-2.5 px-4 text-right font-mono ${textClass} font-bold">${row.shipped.toLocaleString()}</td><td class="py-2.5 px-4 text-right font-mono text-red-500">${row.discrepancy.toLocaleString()}</td><td class="py-2.5 px-4 text-center">${statusBadge}</td><td class="py-2.5 px-4 text-center font-mono font-bold text-emerald-600 dark:text-emerald-400">${row.percentage}%</td></tr>`);
        }
        updateMetricsDisplay(sumForecast, "Total Forecast Volume", sumAllocated, "Total Allocated Qty", sumShipped, "Total Shipped (Done)", sumDiscrepancy, "Total Discrepancy Volume");
        generateHighchartsGraphs(doneCount, pendingCount, whMap, batchMap);

    } else if (fileType === "raw") {
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400"><th class="py-3.5 px-4">Box / MTB ID</th><th class="py-3.5 px-4 text-center">Type</th><th class="py-3.5 px-4">Series / Serial Number</th><th class="py-3.5 px-4">Destination Warehouse</th><th class="py-3.5 px-4 text-right">Qty</th><th class="py-3.5 px-4 text-center">Status</th><th class="py-3.5 px-4 text-center">Shipped Date</th><th class="py-3.5 px-4">Remarks</th></tr>`;
        
        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            let statusLow = row.status.toLowerCase();
            let isDone = statusLow.includes('done') || (statusLow.includes('ship') && !statusLow.includes('pending'));
            if (isDone) doneCount++; else pendingCount++;
            if (row.wh) whMap[row.wh] = (whMap[row.wh] || 0) + 1;
            if (row.series) batchMap[row.series] = (batchMap[row.series] || 0) + 1;
            
            let origIndex = row._idx;
            let statusBadge = isDone ? 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>` : 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;
            
            let typeColor = row.type === "HighValue" ? "bg-purple-100 text-purple-800 border-purple-300" : "bg-sky-100 text-sky-800 border-sky-300";
            let typeBadge = `<button onclick="toggleType(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-md text-[11px] font-black tracking-wide border transition-all ${typeColor}">${row.type}</button>`;
            
            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40"><td class="py-2.5 px-4 font-mono font-bold ${textClass}">${row.mtb}</td><td class="py-2.5 px-4 text-center">${typeBadge}</td><td class="py-2.5 px-4 ${fontMuted} text-xs font-mono">${row.series || ''}</td><td class="py-2.5 px-4 text-orange-600 dark:text-orange-400 font-bold">${row.wh || ''}</td><td class="py-2.5 px-4 text-right font-mono ${fontMuted}">${row.qty ? row.qty.toLocaleString() : '1'}</td><td class="py-2.5 px-4 text-center">${statusBadge}</td><td class="py-2.5 px-4 text-center ${fontMuted} font-mono text-xs">${row.date || ''}</td><td class="py-2.5 px-4 ${fontMuted} text-xs italic truncate max-w-[120px]" title="${row.remarks || ''}">${row.remarks || ''}</td></tr>`);
        }
        updateMetricsDisplay(data.length, "Total Loaded Boxes", doneCount, "Boxes Shipped (Done)", pendingCount, "Pending for Ship", "N/A", "Discrepancy (Raw)");
        generateHighchartsGraphs(doneCount, pendingCount, whMap, batchMap);
        
    } else if (fileType === "crossdock") {
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400"><th class="py-3.5 px-4">Date</th><th class="py-3.5 px-4">MTB ID</th><th class="py-3.5 px-4">SKU Name</th><th class="py-3.5 px-4 text-right">Req Qty</th><th class="py-3.5 px-4 text-right">Act Qty</th><th class="py-3.5 px-4">Pallet</th><th class="py-3.5 px-4 text-center">Checked By</th><th class="py-3.5 px-4 text-center">Status</th></tr>`;
        
        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            sumAllocated += row.qty; sumShipped += row.actualQty;
            let isReceived = row.statusReceived.toLowerCase().includes('done');
            if (isReceived) doneCount++; else pendingCount++;
            if (row.pallet) whMap[row.pallet] = (whMap[row.pallet] || 0) + row.actualQty;
            if (row.skuName) batchMap[row.skuName] = (batchMap[row.skuName] || 0) + row.actualQty;
            
            let origIndex = row._idx;
            let statusBadge = isReceived ? 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>` : 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;
            
            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40"><td class="py-2.5 px-4 font-mono text-xs ${fontMuted}">${row.date}</td><td class="py-2.5 px-4 font-bold font-mono ${textClass}">${row.mtb}</td><td class="py-2.5 px-4 ${textClass} font-medium truncate max-w-[200px]" title="${row.skuName}">${row.skuName}</td><td class="py-2.5 px-4 text-right font-mono ${fontMuted}">${row.qty.toLocaleString()}</td><td class="py-2.5 px-4 text-right font-mono text-blue-600 dark:text-blue-400 font-bold">${row.actualQty.toLocaleString()}</td><td class="py-2.5 px-4 text-xs ${fontMuted} truncate max-w-[150px]">${row.pallet}</td><td class="py-2.5 px-2 text-center"><input type="text" value="${row.checkedBy}" oninput="updateCheckerName(${origIndex}, this.value)" class="w-full text-center font-bold bg-slate-100 dark:bg-slate-900 border border-slate-300 dark:border-slate-700 text-slate-800 dark:text-slate-100 rounded px-1.5 py-1 text-xs focus:bg-white focus:border-indigo-500 focus:outline-none" /></td><td class="py-2.5 px-4 text-center">${statusBadge}</td></tr>`);
        }
        updateMetricsDisplay(data.length, "Total Line Items", sumAllocated, "Total Ordered Qty", sumShipped, "Total Actual Box Qty", (sumAllocated - sumShipped), "Variance / Gap Qty");
        generateHighchartsGraphs(doneCount, pendingCount, whMap, batchMap);
        
    } else if (fileType === "device") {
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400">
            <th class="py-3.5 px-4">Device ID</th>
            <th class="py-3.5 px-4">Device Type</th>
            <th class="py-3.5 px-4 text-center">Specification (Rows)</th>
            <th class="py-3.5 px-4 text-center">Specification (Columns)</th>
            <th class="py-3.5 px-4 text-center">Status</th>
            <th class="py-3.5 px-4 text-center">Device Progress</th>
            <th class="py-3.5 px-4 text-center">Template ID</th>
        </tr>`;

        let progressMap = {}, typeMap = {};
        const isDarkMode = document.documentElement.classList.contains('dark');
        const inputStyle = isDarkMode
            ? 'background:#1e293b;color:#e2e8f0;border:1px solid #475569;'
            : 'background:#f8fafc;color:#1e293b;border:1px solid #cbd5e1;';

        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            let statusLow = row.status.toLowerCase();
            let isDone = statusLow === 'done';
            if (isDone) doneCount++; else pendingCount++;

            typeMap[row.deviceType] = (typeMap[row.deviceType] || 0) + 1;
            progressMap[row.progress] = (progressMap[row.progress] || 0) + 1;

            let origIndex = row._idx;

            // --- STATUS: clickable button cycling Normal → Done → Pending → Normal ---
            let statusColor, statusLabel;
            if (statusLow === 'normal') {
                statusColor = 'background:#e0f2fe;color:#0369a1;border:1px solid #7dd3fc;';
                statusLabel = 'Normal';
            } else if (statusLow === 'done') {
                statusColor = 'background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;';
                statusLabel = 'Done';
            } else {
                statusColor = 'background:#fef3c7;color:#92400e;border:1px solid #fcd34d;';
                statusLabel = row.status || 'Pending';
            }
            let statusBadge = `<button onclick="cycleDeviceStatus(${origIndex})" style="${statusColor}padding:2px 10px;border-radius:9999px;font-size:11px;font-weight:700;cursor:pointer;">${statusLabel}</button>`;

            // --- SPEC ROWS: editable number input ---
            let specRowInput = `<input type="number" min="0" value="${row.specRows !== undefined ? row.specRows : 0}" onchange="updateDeviceField(${origIndex},'specRows',this.value)" style="${inputStyle}width:70px;border-radius:6px;padding:2px 6px;font-size:12px;font-family:monospace;text-align:center;" />`;

            // --- SPEC COLS: editable number input ---
            let specColInput = `<input type="number" min="0" value="${row.specCols !== undefined ? row.specCols : 0}" onchange="updateDeviceField(${origIndex},'specCols',this.value)" style="${inputStyle}width:70px;border-radius:6px;padding:2px 6px;font-size:12px;font-family:monospace;text-align:center;" />`;

            // --- DEVICE PROGRESS: editable text input (type to change) ---
            let progressInput = `<input type="text" value="${row.progress || ''}" onchange="updateDeviceField(${origIndex},'progress',this.value)" style="${inputStyle}width:140px;border-radius:6px;padding:2px 8px;font-size:12px;font-weight:600;" placeholder="Type progress..." />`;

            // --- TEMPLATE ID: show exactly as loaded from Excel/CSV ---
            let tplVal = (row.templateId && row.templateId !== '-' && row.templateId.trim() !== '') ? row.templateId : '—';
            let tplDisplay = tplVal !== '—'
                ? `<span style="background:#ede9fe;color:#5b21b6;border:1px solid #c4b5fd;border-radius:6px;padding:2px 8px;font-size:11px;font-family:monospace;font-weight:700;">${tplVal}</span>`
                : `<span style="color:#94a3b8;font-size:12px;">—</span>`;

            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40">
                <td class="py-2.5 px-4 font-mono font-bold ${textClass}">${row.deviceId}</td>
                <td class="py-2.5 px-4 font-medium ${textClass}">${row.deviceType}</td>
                <td class="py-2 px-4 text-center">${specRowInput}</td>
                <td class="py-2 px-4 text-center">${specColInput}</td>
                <td class="py-2.5 px-4 text-center">${statusBadge}</td>
                <td class="py-2 px-4 text-center">${progressInput}</td>
                <td class="py-2.5 px-4 text-center">${tplDisplay}</td>
            </tr>`);
        }

        updateMetricsDisplay(
            data.length,   "Total Devices",
            doneCount,     "Completed Devices",
            pendingCount,  "Devices In Progress",
            Object.keys(typeMap).length, "Device Types"
        );
        generateHighchartsGraphs(doneCount, pendingCount, typeMap, progressMap);

    } else if (fileType === "generic") {
        let headerRowHtml = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400">`;
        genericHeaders.forEach(h => {
            headerRowHtml += `<th class="py-3.5 px-4 capitalize">${h}</th>`;
        });
        headerRowHtml += `</tr>`;
        thead.innerHTML = headerRowHtml;

        data.forEach(row => {
            let rowHtml = `<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40">`;
            genericHeaders.forEach(h => {
                rowHtml += `<td class="py-2.5 px-4 text-xs ${textClass}">${row[h] || '-'}</td>`;
            });
            rowHtml += `</tr>`;
            htmlBuffer.push(rowHtml);
        });
        
        updateMetricsDisplay(data.length, "Total Records", genericHeaders.length, "Total Columns", 0, "Custom File Active", 0, "No Discrepancy Calc");
        clearHighchartsGraphs();


    } else if (fileType === "cancelled") {
        // ✅ MTO Monitoring / Cancelled — all 11 columns with correct headers
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400 text-xs uppercase tracking-wide">
            <th class="py-3.5 px-3">WHs</th>
            <th class="py-3.5 px-3">MT Order</th>
            <th class="py-3.5 px-3">Picking ID</th>
            <th class="py-3.5 px-3">SKU</th>
            <th class="py-3.5 px-3">SKU Name</th>
            <th class="py-3.5 px-3 text-right">Qty</th>
            <th class="py-3.5 px-3 text-right">Picked Qty</th>
            <th class="py-3.5 px-3 text-right">Checked Qty</th>
            <th class="py-3.5 px-3 text-right">Lack Item</th>
            <th class="py-3.5 px-3">Reason</th>
            <th class="py-3.5 px-3">Picker</th>
        </tr>`;

        let whsMap2 = {}, reasonMap2 = {};
        let sumQty2 = 0, sumLack2 = 0;

        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            sumQty2  += row.qty || 0;
            sumLack2 += row.lackItem || 0;

            if (row.whs)    whsMap2[row.whs]       = (whsMap2[row.whs]       || 0) + 1;
            if (row.reason) reasonMap2[row.reason]  = (reasonMap2[row.reason] || 0) + 1;

            let hasPicker  = row.picker  && row.picker  !== '';
            if (hasPicker) doneCount++; else pendingCount++;

            // Reason badge color mapping
            const reasonColors = {
                'NO OTHER LOCATION'  : 'bg-orange-100 text-orange-800 border-orange-300',
                'EXPIRY ITEMS'       : 'bg-red-100    text-red-800    border-red-300',
                'INVALID PICKING ID' : 'bg-purple-100 text-purple-800 border-purple-300',
                'LIQUOR'             : 'bg-sky-100    text-sky-800    border-sky-300',
                'INCOMPLETE CHECKING': 'bg-yellow-100 text-yellow-800 border-yellow-300',
                'SMALL ITEMS'        : 'bg-teal-100   text-teal-800   border-teal-300',
                'DAMAGED ITEMS'      : 'bg-rose-100   text-rose-800   border-rose-300',
                'LACKING ITEM/LOST'  : 'bg-pink-100   text-pink-800   border-pink-300'
            };
            let reasonKey   = (row.reason || '').toUpperCase().trim();
            let reasonClass = reasonColors[reasonKey] || 'bg-slate-100 text-slate-700 border-slate-300';
            let reasonBadge = row.reason
                ? `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold border ${reasonClass} whitespace-nowrap">${row.reason}</span>`
                : `<span class="${fontMuted} text-xs">—</span>`;

            let pickerBadge = hasPicker
                ? `<span class="px-2 py-0.5 rounded-md text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">${row.picker}</span>`
                : `<span class="${fontMuted} text-xs italic">—</span>`;

            let lackDisplay = row.lackItem > 0
                ? `<span class="font-mono font-bold text-red-500">${row.lackItem.toLocaleString()}</span>`
                : `<span class="${fontMuted} text-xs">—</span>`;

            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40">
                <td class="py-2 px-3 font-black text-orange-600 dark:text-orange-400 text-xs">${row.whs || '—'}</td>
                <td class="py-2 px-3 font-mono font-bold ${textClass} text-xs">${row.mtOrder}</td>
                <td class="py-2 px-3 font-mono ${fontMuted} text-xs">${row.pickingId || '—'}</td>
                <td class="py-2 px-3 ${fontMuted} text-xs font-mono truncate max-w-[100px]" title="${row.sku}">${row.sku || '—'}</td>
                <td class="py-2 px-3 ${textClass} text-xs truncate max-w-[180px]" title="${row.skuName}">${row.skuName || '—'}</td>
                <td class="py-2 px-3 text-right font-mono ${fontMuted} text-xs">${row.qty > 0 ? row.qty.toLocaleString() : '—'}</td>
                <td class="py-2 px-3 text-right font-mono text-blue-600 dark:text-blue-400 text-xs">${row.pickedQty > 0 ? row.pickedQty.toLocaleString() : '—'}</td>
                <td class="py-2 px-3 text-right font-mono text-indigo-600 dark:text-indigo-400 text-xs">${row.checkedQty > 0 ? row.checkedQty.toLocaleString() : '—'}</td>
                <td class="py-2 px-3 text-right">${lackDisplay}</td>
                <td class="py-2 px-3">${reasonBadge}</td>
                <td class="py-2 px-3 text-center">${pickerBadge}</td>
            </tr>`);
        }

        let uniqueReasons = Object.keys(reasonMap2).length;
        updateMetricsDisplay(data.length, "Total Cancelled Lines", sumQty2, "Total Qty", sumLack2, "Total Lack Items", uniqueReasons, "Cancellation Reasons");
        generateHighchartsGraphs(doneCount, pendingCount, whsMap2, reasonMap2);
    }
    
    tbody.innerHTML = htmlBuffer.join('');
    document.getElementById('rowCount').innerText = `May kabuuang ${data.length} na talaan ang aktibo sa system.`;

    // ✅ Update the 4 summary cards AND capture values for pie chart sync
    const cardTotals = updateSummaryCards(data);

    // ✅ Update the latest date card
    updateLatestDateCard(data);

    // ✅ Re-draw pie chart so it ALWAYS matches the cards exactly
    if (fileType !== "generic") {
        refreshPieChartFromCards(cardTotals);
    }
}

function updateCheckerName(index, bagongPangalan) { globalData[index].checkedBy = bagongPangalan.toUpperCase(); saveAppStateDebounced(); }

function toggleStatus(index) {
    let row = globalData[index];
    if (fileType === "progress") {
        if (row.percentage < 100 || row.discrepancy > 0) {
            row.shipped = row.allocated; row.discrepancy = 0; row.percentage = 100;
        } else {
            row.shipped = 0; row.discrepancy = row.allocated; row.percentage = 0;
        }
    } else if (fileType === "raw") {
        let statusLow = row.status.toLowerCase();
        if (statusLow.includes('done') || (statusLow.includes('ship') && !statusLow.includes('pending'))) {
            row.status = "Pending";
        } else {
            row.status = "Done"; let ngayon = new Date(); row.date = ngayon.toISOString().split('T')[0]; 
        }
    } else if (fileType === "crossdock") {
        row.statusReceived = row.statusReceived.toLowerCase() === 'done' ? "Pending" : "Done";
    }
    // Skip chart redraw on status toggle — only rebuild table + cards
    _skipChartRedraw = true;
    let query = document.getElementById('searchBar').value;
    if (query) { document.getElementById('searchBar').dispatchEvent(new Event('input')); } else { buildDashboard(globalData); }
    _skipChartRedraw = false;
    saveAppState();
}
// ====================================================================
// 🖊️ DEVICE TABLE INTERACTIVE CONTROLS
// ====================================================================
// Cycle device Status: Normal → Done → Pending → Normal
function cycleDeviceStatus(index) {
    let row = globalData[index];
    let current = (row.status || 'Normal').toLowerCase();
    if (current === 'normal') {
        row.status = 'Done';
    } else if (current === 'done') {
        row.status = 'Pending';
    } else {
        row.status = 'Normal';
    }
    _skipChartRedraw = true;
    let query = document.getElementById('searchBar').value;
    if (query) { document.getElementById('searchBar').dispatchEvent(new Event('input')); } else { buildDashboard(globalData); }
    _skipChartRedraw = false;
    saveAppState();
}

// Update any device field (specRows, specCols, progress) without full re-render
function updateDeviceField(index, field, value) {
    globalData[index][field] = value;
    saveAppStateDebounced();
}

function updateMetricsDisplay(m1v, m1l, m2v, m2l, m3v, m3l, m4v, m4l) {
    document.getElementById('metric1Value').innerText = m1v.toLocaleString(); document.getElementById('metric1Label').innerText = m1l;
    document.getElementById('metric2Value').innerText = m2v.toLocaleString(); document.getElementById('metric2Label').innerText = m2l;
    document.getElementById('metric3Value').innerText = m3v.toLocaleString(); document.getElementById('metric3Label').innerText = m3l;
    document.getElementById('metric4Value').innerText = m4v.toLocaleString(); document.getElementById('metric4Label').innerText = m4l;
}

function clearHighchartsGraphs() {
    ['chart1Container', 'chart2Container', 'chart3Container'].forEach(id => {
        const el = document.getElementById(id);
        if(el) el.innerHTML = `<div class="h-full flex items-center justify-center text-xs text-slate-400 italic">Walang visual graph para sa custom list file.</div>`;
    });
}


// ====================================================================
// 🥧 PIE CHART SYNC — always mirrors the 4 summary cards exactly
// ====================================================================
function refreshPieChartFromCards(cardTotals) {
    if (_skipChartRedraw) return; // perf: skip on status toggle
    const isDark = document.documentElement.classList.contains('dark');
    const lblColor = isDark ? '#94a3b8' : '#475569';

    let pieData = [];

    if (fileType === "raw") {
        // For Raw files: show Done/Pending AND Local/HighValue in one pie
        pieData = [
            { name: '✅ Done',        y: cardTotals.totalDone,       color: '#10b981' },
            { name: '⏳ Pending',     y: cardTotals.totalPending,    color: '#f59e0b' },
            { name: '📦 Local',       y: cardTotals.totalLocal,      color: '#38bdf8' },
            { name: '💎 High Value',  y: cardTotals.totalHighValue,  color: '#a855f7' }
        ].filter(d => d.y > 0);
    } else {
        // For all other file types: show Done vs Pending only
        pieData = [
            { name: '✅ Done / Received', y: cardTotals.totalDone,    color: '#10b981' },
            { name: '⏳ Pending',         y: cardTotals.totalPending, color: '#f59e0b' }
        ].filter(d => d.y > 0);
    }

    Highcharts.chart('chart1Container', {
        chart: { type: 'pie', options3d: { enabled: true, alpha: 45, beta: 0 }, backgroundColor: 'transparent' },
        title: { text: '📊 Status Overview', style: { fontSize: '12px', fontWeight: 'bold', color: lblColor } },
        plotOptions: {
            pie: {
                depth: 35,
                allowPointSelect: true,
                cursor: 'pointer',
                dataLabels: {
                    enabled: true,
                    format: '<b>{point.name}</b><br>{point.y} ({point.percentage:.0f}%)',
                    style: { fontSize: '9px', color: lblColor, fontWeight: 'bold' }
                },
                showInLegend: true
            }
        },
        legend: { itemStyle: { fontSize: '10px', color: lblColor } },
        credits: { enabled: false },
        series: [{ name: 'Count', data: pieData }]
    });
}

// ====================================================================
// 📊 HIGHCHARTS VISUALIZATION ENGINE
// ====================================================================
function generateHighchartsGraphs(doneCount, pendingCount, whMap, batchMap) {
    if (_skipChartRedraw) return; // perf: skip on status toggle
    const isDark = document.documentElement.classList.contains('dark');
    const lblColor = isDark ? '#94a3b8' : '#475569';

    // chart1 is now handled by refreshPieChartFromCards() to keep it in sync with the cards

    let whData = Object.keys(whMap).map(key => ({ name: key, y: whMap[key] })).slice(0, 10);
    Highcharts.chart('chart2Container', {
        chart: { type: 'pie', options3d: { enabled: true, alpha: 60, beta: 0 }, backgroundColor: 'transparent' },
        title: {
            text: fileType === "crossdock" ? '📊 Pallet Load Mix' : fileType === "cancelled" ? '🏬 WHs Distribution' : '🚚 Warehouse Mix',
            style: { fontSize: '13px', fontWeight: 'bold', color: lblColor }
        },
        plotOptions: { pie: { innerSize: 70, depth: 35, allowPointSelect: true, cursor: 'pointer', dataLabels: { enabled: false } } },
        credits: { enabled: false },
        series: [{ name: 'Share Volume', data: whData }]
    });

    let sortedBatches = Object.keys(batchMap).sort((a,b) => batchMap[b] - batchMap[a]).slice(0, 7);
    let batchValues = sortedBatches.map(b => batchMap[b]);
    Highcharts.chart('chart3Container', {
        chart: { type: 'column', options3d: { enabled: true, alpha: 15, beta: 15, depth: 50, viewDistance: 25 }, backgroundColor: 'transparent' },
        title: { text: null }, 
        xAxis: { categories: sortedBatches, labels: { style: { fontSize: '9px', color: lblColor } } },
        yAxis: { title: { text: null }, labels: { style: { color: lblColor } } },
        plotOptions: { column: { depth: 25, color: '#6366f1' } }, 
        legend: { enabled: false }, 
        credits: { enabled: false }, 
        series: [{ name: 'Volume', data: batchValues }]
    });
}
// ====================================================================
// 🔄 CORE DATA PROCESSOR — reusable parser for main uploader + Live Excel modal
// ====================================================================
function processFileContents(contents, fileName) {
    lines = contents.split(/\r?\n/).filter(l => l.trim() !== "");
    if(lines.length === 0) {
        // Ang file ay ganap na walang laman — ipakita blangkong dashboard
        globalData = [];
        document.getElementById('tableTitle').innerText = "📁 " + fileName;
        document.getElementById('searchBar').disabled = false;
        document.getElementById('searchBar').placeholder = "Mag-type rito para mag-filter ng data...";
        document.getElementById('btnDashboard').click(); buildDashboard([]);
        saveAppState();
        return;
    }

    let isAttendanceFile = false;
    for (let i = 0; i < Math.min(lines.length, 5); i++) {
        let lowLine = lines[i].toLowerCase();
        if (lowLine.includes('attendance') || lowLine.includes('number of ot') || lowLine.includes('daily rate')) {
            isAttendanceFile = true; break;
        }
    }

    if (isAttendanceFile) {
        fileType = "attendance";
        parseAttendanceCSV(contents); 
        document.getElementById('btnAttendance').click(); 
        saveAppState();
        return;
    }
    
    let startRow = 0;
    let foundMatch = false;
    for (let i = 0; i < lines.length; i++) {
        let lowerLine = lines[i].toLowerCase();
        if (lowerLine.includes('mto batch') || lowerLine.includes('box id') || lowerLine.includes('mtb') || lowerLine.includes('sku name') || lowerLine.includes('type') || lowerLine.includes('series') || lowerLine.includes('destination') || lowerLine.includes('high value') || lowerLine.includes('device id') || lowerLine.includes('device type') || lowerLine.includes('device progress')) {
            startRow = i; 
            foundMatch = true;
            break;
        }
    }
    
    let headers = parseCSVLine(lines[startRow]).map(h => h.trim());
    let lowerHeaders = headers.map(h => h.toLowerCase());
    let loadedRows = [];
    
    if (foundMatch && lowerHeaders.some(h => h.includes('device id') || h.includes('device type') || h.includes('device progress'))) {
        fileType = "device"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let deviceIdIdx    = lowerHeaders.findIndex(h => h.includes('device id'));
        let deviceTypeIdx  = lowerHeaders.findIndex(h => h.includes('device type'));
        let specRowsIdx    = lowerHeaders.findIndex(h => h.includes('specification') && h.includes('row'));
        let specColsIdx    = lowerHeaders.findIndex(h => h.includes('specification') && h.includes('col'));
        let statusIdx      = lowerHeaders.findIndex(h => h.includes('status'));
        let progressIdx    = lowerHeaders.findIndex(h => h.includes('device progress') || h.includes('progress'));
        let templateIdIdx  = lowerHeaders.findIndex(h => h.includes('template id') || h.includes('template'));

        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim());
            if (!columns[deviceIdIdx] || columns[deviceIdIdx].trim() === "") continue;
            loadedRows.push({
                deviceId:   columns[deviceIdIdx]   ? columns[deviceIdIdx].trim()   : '-',
                deviceType: columns[deviceTypeIdx] ? columns[deviceTypeIdx].trim() : '-',
                specRows:   columns[specRowsIdx]   ? columns[specRowsIdx].trim()   : '0',
                specCols:   columns[specColsIdx]   ? columns[specColsIdx].trim()   : '0',
                status:     columns[statusIdx]     ? columns[statusIdx].trim()     : 'Normal',
                progress:   columns[progressIdx]   ? columns[progressIdx].trim()   : '-',
                templateId: columns[templateIdIdx] ? columns[templateIdIdx].trim() : '-'
            });
        }
    } else if (foundMatch && lowerHeaders.some(h => h.includes('sku name')) && lowerHeaders.some(h => h.includes('checked by'))) {
        fileType = "crossdock"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let dateIdx = lowerHeaders.findIndex(h => h.includes('date')), skuNameIdx = lowerHeaders.findIndex(h => h.includes('sku name'));
        let qtyIdx = lowerHeaders.findIndex(h => h === 'qty' || h.includes('req qty')), mtbIdx = lowerHeaders.findIndex(h => h.includes('mtb'));
        let actQtyIdx = lowerHeaders.findIndex(h => h.includes('actual qty') || h.includes('actual box')), palletIdx = lowerHeaders.findIndex(h => h.includes('pallet')), checkerIdx = lowerHeaders.findIndex(h => h.includes('checked by')), statusRecIdx = lowerHeaders.findIndex(h => h.includes('status if received'));
        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim()); if (!columns[mtbIdx]) continue;
            loadedRows.push({ date: columns[dateIdx] ? columns[dateIdx].trim() : '', skuName: columns[skuNameIdx] ? columns[skuNameIdx].trim() : '', qty: cleanNumber(columns[qtyIdx]), mtb: columns[mtbIdx].trim().replace(/"/g, ''), actualQty: cleanNumber(columns[actQtyIdx]), pallet: columns[palletIdx] ? columns[palletIdx].trim() : '', checkedBy: columns[checkerIdx] && columns[checkerIdx].trim() !== "" ? columns[checkerIdx].trim().toUpperCase() : '', statusReceived: columns[statusRecIdx] ? columns[statusRecIdx].trim() : 'Pending' });
        }
    } else if (foundMatch && lowerHeaders.some(h => h.includes('forecast') || h.includes('mto batch'))) {
        fileType = "progress"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let whIdx = lowerHeaders.findIndex(h => h.includes('wh') || h.includes('warehouse')), wwIdx = lowerHeaders.findIndex(h => h.includes('work week') || h.includes('ww')), monthIdx = lowerHeaders.findIndex(h => h.includes('month')), batchIdx = lowerHeaders.findIndex(h => h.includes('mto batch') || h.includes('batch')), forecastIdx = lowerHeaders.findIndex(h => h.includes('forecast')), allocatedIdx = lowerHeaders.findIndex(h => h.includes('allocated')), shippedIdx = lowerHeaders.findIndex(h => h.includes('total shipped') || h.includes('shipped')), discIdx = lowerHeaders.findIndex(h => h.includes('discrepancy')), pctIdx = lowerHeaders.findIndex(h => h.includes('percentage') || h.includes('rate'));
        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim()); if (!columns[batchIdx] || columns[batchIdx].toLowerCase().includes('overall')) continue;
            loadedRows.push({ wh: columns[whIdx] ? columns[whIdx].trim().toUpperCase() : '', ww: columns[wwIdx] ? columns[wwIdx].trim() : '-', month: columns[monthIdx] ? columns[monthIdx].trim() : '-', batch: columns[batchIdx].trim().replace(/"/g, ''), forecast: cleanNumber(columns[forecastIdx]), allocated: cleanNumber(columns[allocatedIdx]), shipped: cleanNumber(columns[shippedIdx]), discrepancy: cleanNumber(columns[discIdx]), percentage: parseInt(columns[pctIdx] ? columns[pctIdx].replace('%', '') : '0') || 0 });
        }
    } else if (foundMatch && lowerHeaders.some(h => h.includes('lack item') || h.includes('lack') || h.includes('picking id'))) {
        // ✅ NEW: MTO Monitoring / Cancelled file type
        fileType = "cancelled"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let whsIdxC    = lowerHeaders.findIndex(h => h === 'whs' || h.includes('wh'));
        let mtOrderIdx = lowerHeaders.findIndex(h => h === 'mt order' || h.includes('mt order'));
        let pickIdIdx  = lowerHeaders.findIndex(h => h.includes('picking id') || h.includes('picking'));
        let skuIdxC    = lowerHeaders.findIndex(h => h === 'sku');
        let skuNmIdx   = lowerHeaders.findIndex(h => h === 'sku name' || (h.includes('sku') && h.includes('name')));
        let qty2Idx    = lowerHeaders.findIndex(h => h === 'qty');
        let pickedIdx  = lowerHeaders.findIndex(h => h.includes('picked qty') || h.includes('picked'));
        let checkedIdx = lowerHeaders.findIndex(h => h.includes('checked qty') || (h.includes('checked') && !h.includes('checked by')));
        let lackIdx    = lowerHeaders.findIndex(h => h.includes('lack item') || h.includes('lack'));
        let reasonIdx  = lowerHeaders.findIndex(h => h.includes('reason'));
        let pickerIdx  = lowerHeaders.findIndex(h => h.includes('picker'));

        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim());
            if (!columns[mtOrderIdx] || columns[mtOrderIdx].trim() === '') continue;
            loadedRows.push({
                whs:        columns[whsIdxC]    ? columns[whsIdxC].trim().toUpperCase()          : '',
                mtOrder:    columns[mtOrderIdx] ? columns[mtOrderIdx].trim()                     : '',
                pickingId:  columns[pickIdIdx]  ? columns[pickIdIdx].trim()                      : '',
                sku:        columns[skuIdxC]    ? columns[skuIdxC].trim().replace(/"/g,'')        : '',
                skuName:    columns[skuNmIdx]   ? columns[skuNmIdx].trim().replace(/"/g,'')       : '',
                qty:        cleanNumber(columns[qty2Idx]),
                pickedQty:  cleanNumber(columns[pickedIdx]),
                checkedQty: cleanNumber(columns[checkedIdx]),
                lackItem:   cleanNumber(columns[lackIdx]),
                reason:     columns[reasonIdx]  ? columns[reasonIdx].trim()                      : '',
                picker:     columns[pickerIdx]  ? columns[pickerIdx].trim().toUpperCase()         : ''
            });
        }
    } else if (foundMatch) {
        fileType = "raw"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let typeIdx = lowerHeaders.findIndex(h => h === 'type' || h.includes('class') || h.includes('high value'));
        let mtbIdx = lowerHeaders.findIndex(h => h === 'mtb' || h.includes('box id') || h.includes('mtb id') || h.includes('box'));
        let seriesIdx = lowerHeaders.findIndex(h => h.includes('series') || h.includes('pallet') || h.includes('serial') || h.includes('number'));
        let whIdx = lowerHeaders.findIndex(h => h.includes('wh') || h.includes('destination') || h.includes('warehouse') || h.includes('dest'));
        let statusIdx = lowerHeaders.findIndex(h => h.includes('status') || h.includes('shipped)'));
        let dateIdx = lowerHeaders.findIndex(h => h.includes('date') || h.includes('shipped date') || h.includes('time'));
        let qtyIdx = lowerHeaders.findIndex(h => h === 'qty' || h.includes('quantity') || h.includes('pcs'));
        let remarksIdx = lowerHeaders.findIndex(h => h.includes('remark') || h.includes('note'));
        
        if (mtbIdx === -1) mtbIdx = lowerHeaders.findIndex(h => h.includes('mtb')) !== -1 ? lowerHeaders.findIndex(h => h.includes('mtb')) : 1;
        if (seriesIdx === -1) seriesIdx = lowerHeaders.findIndex(h => h.includes('series')) !== -1 ? lowerHeaders.findIndex(h => h.includes('series')) : 2;
        if (whIdx === -1) whIdx = lowerHeaders.findIndex(h => h.includes('wh') || h.includes('dest')) !== -1 ? lowerHeaders.findIndex(h => h.includes('wh') || h.includes('dest')) : 3;
        if (statusIdx === -1) statusIdx = lowerHeaders.findIndex(h => h.includes('status')) !== -1 ? lowerHeaders.findIndex(h => h.includes('status')) : 4;
        
        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim()); if (!columns[mtbIdx] || columns[mtbIdx].trim() === "" || columns[mtbIdx].toLowerCase().includes('type')) continue;
            
            let rawTypeValue = typeIdx !== -1 && columns[typeIdx] ? columns[typeIdx].trim().toUpperCase() : '';
            let rawType = "Local";
            if (rawTypeValue.includes("HIGH VALUE")) {
                rawType = "HighValue";
            }
            
            loadedRows.push({ type: rawType, mtb: columns[mtbIdx].trim().replace(/"/g, ''), series: columns[seriesIdx] ? columns[seriesIdx].trim().replace(/"/g, '') : '', wh: columns[whIdx] ? columns[whIdx].trim().toUpperCase() : '', qty: qtyIdx !== -1 && columns[qtyIdx] ? cleanNumber(columns[qtyIdx]) : 1, status: columns[statusIdx] ? columns[statusIdx].trim() : 'Pending', date: dateIdx !== -1 && columns[dateIdx] ? columns[dateIdx].trim() : '', remarks: remarksIdx !== -1 && columns[remarksIdx] ? columns[remarksIdx].trim() : '' });
        }
    } else {
        fileType = "generic";
        document.getElementById('tableTitle').innerText = "📁 " + fileName;
        headers = parseCSVLine(lines[0]).map(h => h.trim().replace(/"/g, ''));
        genericHeaders = headers;
        
        for (let i = 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim());
            if (columns.length === 0 || (columns.length === 1 && columns[0] === "")) continue;
            let rowObj = {};
            headers.forEach((h, index) => {
                rowObj[h] = columns[index] ? columns[index].trim().replace(/"/g, '') : '';
            });
            loadedRows.push(rowObj);
        }
    }
    
    if (loadedRows.length > 0) {
        // Stamp stable index on every row — avoids expensive indexOf() calls
        loadedRows.forEach((r, i) => { r._idx = i; });
        globalData = loadedRows;
        // Reset date picker on new file upload
        const dp = document.getElementById('datePicker');
        dp.value = '';
        dp.disabled = true;
        document.getElementById('clearDateBtn').classList.add('hidden');
        document.getElementById('latestDateLabel').innerText = '—';
        document.getElementById('latestDateCount').innerText = '—';
        document.getElementById('searchBar').disabled = false;
        document.getElementById('searchBar').placeholder = "Mag-type rito para mag-filter ng data...";
        document.getElementById('btnDashboard').click(); buildDashboard(globalData);
        saveAppState();
    } else {
        // Walang data rows — ipakita ang blangkong dashboard (huwag mag-alert)
        globalData = [];
        const dp = document.getElementById('datePicker');
        dp.value = ''; dp.disabled = true;
        document.getElementById('clearDateBtn').classList.add('hidden');
        document.getElementById('latestDateLabel').innerText = '—';
        document.getElementById('latestDateCount').innerText = '—';
        document.getElementById('searchBar').disabled = false;
        document.getElementById('searchBar').placeholder = "Mag-type rito para mag-filter ng data...";
        document.getElementById('btnDashboard').click(); buildDashboard([]);
        saveAppState();
    }

}

// ====================================================================
// 📊 LIVE EXCEL MODAL — multi-sheet Excel selector for Attendance & Dashboard
// ====================================================================

// Shared helper: i-convert lahat ng date cells sa isang sheet papuntang YYYY-MM-DD
// string bago i-export sa CSV. Ginagamit ng main uploader, snapshot upload, AT
// ng live-watch polling loop — iisa lang ang logic para walang magkakaibang resulta.
function convertSheetDatesToISO(sheet) {
    if (!sheet['!ref']) return;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = sheet[addr];
            if (cell && cell.t === 'd' && cell.v instanceof Date) {
                const d = cell.v;
                const iso = d.getFullYear() + '-' +
                    String(d.getMonth() + 1).padStart(2, '0') + '-' +
                    String(d.getDate()).padStart(2, '0');
                cell.t = 's'; cell.v = iso; cell.w = iso;
            }
        }
    }
}

// Shared helper: gumawa ng mga pindutan para sa bawat sheet sa workbook, at
// awtomatikong piliin ang una. Ginagamit pareho ng live-link at snapshot path.
function renderSheetButtons(sheetNames) {
    const container = document.getElementById('liveExcelSheetButtons');
    container.innerHTML = '';
    sheetNames.forEach((name, idx) => {
        const btn = document.createElement('button');
        btn.dataset.sheet = name;
        btn.className = idx === 0
            ? 'px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-600 text-white border border-emerald-600 cursor-pointer'
            : 'px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 hover:bg-emerald-500 hover:text-white border border-slate-300 dark:border-slate-600 cursor-pointer transition-colors';
        btn.innerText = '📄 ' + name;
        btn.onclick = () => selectLiveExcelSheet(name);
        container.appendChild(btn);
    });
    document.getElementById('liveExcelSheetSelector').classList.remove('hidden');
    selectLiveExcelSheet(sheetNames[0]);
}

function openLiveExcelModal() {
    // Reset the modal to a clean state
    const fileInput = document.getElementById('liveExcelFileInput');
    if (fileInput) fileInput.value = '';
    document.getElementById('liveExcelFileName').textContent = 'Wala pang napiling file';
    document.getElementById('liveExcelSheetSelector').classList.add('hidden');
    document.getElementById('liveExcelSheetButtons').innerHTML = '';
    document.getElementById('liveExcelPreview').classList.add('hidden');
    document.getElementById('liveExcelLoadBtn').disabled = true;
    _liveExcelWorkbook = null;
    _liveExcelSelectedSheet = null;
    _pendingLiveFileHandle = null; // huwag dalhin ang dating di-pa-committed na pick
    document.getElementById('liveExcelModal').classList.remove('hidden');
}

function closeLiveExcelModal() {
    document.getElementById('liveExcelModal').classList.add('hidden');
}

function selectLiveExcelSheet(sheetName) {
    _liveExcelSelectedSheet = sheetName;
    // Highlight selected button
    const btns = document.getElementById('liveExcelSheetButtons').querySelectorAll('button');
    btns.forEach(btn => {
        if (btn.dataset.sheet === sheetName) {
            btn.className = 'px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-600 text-white border border-emerald-600 cursor-pointer';
        } else {
            btn.className = 'px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 hover:bg-emerald-500 hover:text-white border border-slate-300 dark:border-slate-600 cursor-pointer transition-colors';
        }
    });
    // Show preview stats
    const sheet = _liveExcelWorkbook.Sheets[sheetName];
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    const rowCount = range ? (range.e.r - range.s.r) : 0;
    const colCount = range ? (range.e.c - range.s.c + 1) : 0;
    document.getElementById('liveExcelPreviewText').innerText = '📄 Sheet: "' + sheetName + '"';
    document.getElementById('liveExcelPreviewRows').innerText = rowCount.toLocaleString() + ' rows × ' + colCount + ' columns';
    document.getElementById('liveExcelPreview').classList.remove('hidden');
    document.getElementById('liveExcelLoadBtn').disabled = false;
}

function loadSelectedExcelSheet() {
    if (!_liveExcelWorkbook || !_liveExcelSelectedSheet || !_liveExcelFile) return;
    const sheet = _liveExcelWorkbook.Sheets[_liveExcelSelectedSheet];
    convertSheetDatesToISO(sheet);
    const contents = XLSX.utils.sheet_to_csv(sheet);
    const fileName = _liveExcelFile.name + ' › ' + _liveExcelSelectedSheet;
    closeLiveExcelModal();

    const badge = document.getElementById('liveExcelActiveBadge');

    if (_pendingLiveFileHandle) {
        // ✅ May live-linked handle — i-commit bilang ACTIVE at simulan ang polling
        _liveFileHandle = _pendingLiveFileHandle;
        _liveFileLastModified = _liveExcelFile.lastModified;
        _liveFileFailCount = 0;
        clearLivePolling();
        _livePollIntervalId = setInterval(checkLiveFileForChanges, LIVE_POLL_MS);
        if (badge) { badge.textContent = '🔴 ' + _liveExcelSelectedSheet; badge.classList.remove('hidden'); }
        updateLiveSyncIndicator(true);
    } else {
        // Plain snapshot lang — itigil ang dating live sync (kung meron) dahil
        // static na ngayon ang ipinapakita sa dashboard
        clearLivePolling();
        _liveFileHandle = null;
        if (badge) { badge.textContent = _liveExcelSelectedSheet; badge.classList.remove('hidden'); }
        updateLiveSyncIndicator(false);
    }
    _pendingLiveFileHandle = null;

    processFileContents(contents, fileName);
}

// ====================================================================
// 🔴 LIVE LOCAL FILE WATCH — File System Access API (Chrome/Edge desktop lang)
// Hindi tulad ng regular <input type="file">, ang showOpenFilePicker() ay
// nagbibigay ng FileSystemFileHandle na pwedeng paulit-ulit basahin ulit nang
// hindi na kailangang mag-upload — kaya posible ang totoong "live" na pag-sync.
// ====================================================================
async function linkLiveLocalFile() {
    if (!('showOpenFilePicker' in window)) {
        alert('⚠️ Hindi suportado ng browser mo ang Live Local File feature.\n\nGumamit ng Google Chrome o Microsoft Edge (desktop) para gumana ito. Hindi ito available sa Safari o Firefox — pwede mo pa ring gamitin ang "Minsan Lang Mag-upload" option.');
        return;
    }
    try {
        const [handle] = await window.showOpenFilePicker({
            types: [{
                description: 'Excel Files',
                accept: {
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                    'application/vnd.ms-excel': ['.xls']
                }
            }],
            excludeAcceptAllOption: false,
            multiple: false
        });

        const file = await handle.getFile();
        const buf = await file.arrayBuffer();
        const data = new Uint8Array(buf);

        _pendingLiveFileHandle = handle; // hindi pa ACTIVE — magko-commit lang pagka-click ng "I-load sa Dashboard"
        _liveExcelFile = file;
        _liveExcelWorkbook = XLSX.read(data, { type: 'array', cellDates: true });

        document.getElementById('liveExcelFileName').innerHTML =
            '<span class="text-red-600 dark:text-red-400 font-black">🔴 ' + file.name + '</span> <span class="text-slate-400 font-normal">(live-linked)</span>';

        renderSheetButtons(_liveExcelWorkbook.SheetNames);
    } catch (err) {
        if (err.name === 'AbortError') return; // kinansela lang ng user ang picker — tahimik na huminto
        console.error('Live file link error:', err);
        let msg = '⚠️ Hindi na-link ang file.';
        if (err.name === 'SecurityError' || err.name === 'NotAllowedError') {
            msg += '\n\nKaramihan, ito ay dahil binuksan ang dashboard sa pamamagitan lang ng dobleng-click sa index.html. Kailangan itong i-serve via local server (hal. "Live Server" extension sa VS Code, o python -m http.server) o i-host online para gumana ang Live Local File link.';
        }
        alert(msg);
    }
}

function clearLivePolling() {
    if (_livePollIntervalId) {
        clearInterval(_livePollIntervalId);
        _livePollIntervalId = null;
    }
}

// Tinatawag bawat ilang segundo habang naka-live-link. Tinitingnan ang lastModified
// ng file — kung walang pagbabago, wala ring ginagawa (mabilis at di nakaka-abala).
async function checkLiveFileForChanges() {
    if (!_liveFileHandle) { clearLivePolling(); return; }
    try {
        const file = await _liveFileHandle.getFile();
        _liveFileFailCount = 0; // matagumpay na nabasa — i-reset ang error counter
        if (file.lastModified === _liveFileLastModified) return; // walang bagong pagbabago
        _liveFileLastModified = file.lastModified;

        const buf = await file.arrayBuffer();
        const data = new Uint8Array(buf);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });

        if (!workbook.Sheets[_liveExcelSelectedSheet]) {
            console.warn('Live-linked sheet hindi na natagpuan pagkatapos ng update:', _liveExcelSelectedSheet);
            return; // baka na-rename o natanggal — laktawan ang cycle na ito, panoorin pa rin
        }
        const sheet = workbook.Sheets[_liveExcelSelectedSheet];
        convertSheetDatesToISO(sheet);
        const contents = XLSX.utils.sheet_to_csv(sheet);
        const fileName = file.name + ' › ' + _liveExcelSelectedSheet;
        processFileContents(contents, fileName);
        flashLiveSyncIndicator();
    } catch (err) {
        console.warn('Live file check failed:', err);
        _liveFileFailCount++;
        if (_liveFileFailCount >= 3) {
            // Tatlong magkakasunod na error (hal. na-delete o na-move ang file) — itigil na
            clearLivePolling();
            _liveFileHandle = null;
            updateLiveSyncIndicator(false, true);
        }
    }
}

function stopLiveFileWatchManually() {
    clearLivePolling();
    _liveFileHandle = null;
    _liveFileLastModified = null;
    updateLiveSyncIndicator(false);
}

function updateLiveSyncIndicator(isLive, isError) {
    const row = document.getElementById('liveSyncStatusRow');
    const text = document.getElementById('liveSyncStatusText');
    if (!row || !text) return;
    if (isError) {
        row.classList.remove('hidden');
        text.innerText = '⚠️ Nawalan ng koneksyon sa file';
        return;
    }
    if (isLive) {
        row.classList.remove('hidden');
        text.innerText = 'Live syncing • ' + (_liveExcelSelectedSheet || '');
    } else {
        row.classList.add('hidden');
    }
}

function flashLiveSyncIndicator() {
    const text = document.getElementById('liveSyncStatusText');
    if (!text) return;
    const stamp = new Date().toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    text.innerText = '✅ Na-update kanina • ' + stamp;
    setTimeout(() => {
        if (_liveFileHandle) text.innerText = 'Live syncing • ' + (_liveExcelSelectedSheet || '');
    }, 2500);
}

// ====================================================================
// 📂 CORE FILE ROUTER (EXCEL & CSV ACCURATE CONVERSION)
// ====================================================================
document.getElementById('csvFileInput').addEventListener('change', function(e) {
    let file = e.target.files[0]; if (!file) return;
    let reader = new FileReader();
    let isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
    
    reader.onload = function(evt) {
        let contents = "";
        if (isExcel) {
            let data = new Uint8Array(evt.target.result);
            let workbook = XLSX.read(data, { type: 'array', cellDates: true });
            let sheet = workbook.Sheets[workbook.SheetNames[0]];
            convertSheetDatesToISO(sheet);
            contents = XLSX.utils.sheet_to_csv(sheet);
        } else { contents = evt.target.result.replace(/^\uFEFF/, ''); }
        processFileContents(contents, file.name);
    };
    if (isExcel) { reader.readAsArrayBuffer(file); } else { reader.readAsText(file); }
});


// Dynamic Client Filter Control — debounced to avoid lag on fast typing
let _searchDebounceTimer = null;
document.getElementById('searchBar').addEventListener('input', function(e) {
    clearTimeout(_searchDebounceTimer);
    _searchDebounceTimer = setTimeout(function() {
    let query = e.target.value.toLowerCase().trim(); if (!query) { buildDashboard(globalData); return; }
    let filtered = globalData.filter(row => {
        if (fileType === "progress") return row.wh.toLowerCase().includes(query) || row.batch.toLowerCase().includes(query) || row.month.toLowerCase().includes(query);
        if (fileType === "raw") return row.mtb.toLowerCase().includes(query) || row.wh.toLowerCase().includes(query) || row.series.toLowerCase().includes(query) || row.type.toLowerCase().includes(query) || (row.remarks && row.remarks.toLowerCase().includes(query));
        if (fileType === "crossdock") return row.mtb.toLowerCase().includes(query) || row.skuName.toLowerCase().includes(query) || row.checkedBy.toLowerCase().includes(query);
        if (fileType === "device") return row.deviceId.toLowerCase().includes(query) || row.deviceType.toLowerCase().includes(query) || row.status.toLowerCase().includes(query) || row.progress.toLowerCase().includes(query) || row.templateId.toLowerCase().includes(query);
        if (fileType === "cancelled") return row.mtOrder.toLowerCase().includes(query) || row.whs.toLowerCase().includes(query) || row.skuName.toLowerCase().includes(query) || row.pickingId.toLowerCase().includes(query) || (row.reason && row.reason.toLowerCase().includes(query)) || (row.picker && row.picker.toLowerCase().includes(query)) || (row.sku && row.sku.toLowerCase().includes(query));
        if (fileType === "generic") {
            return genericHeaders.some(h => String(row[h]).toLowerCase().includes(query));
        }
    });
    buildDashboard(filtered);
    }, 180); // debounce: wait 180ms after user stops typing
});

// ====================================================================
// 📅 ATTENDANCE TIMESHEET SYSTEM CONTROLLER
// ====================================================================
function parseAttendanceCSV(text) {
    const lines = text.split(/\r?\n/).map(line => parseCSVLine(line.trim())).filter(l => l.length > 0 && l.some(cell => cell.trim() !== ""));
    if (lines.length < 3) return;
    const headerRow = lines[0];
    const columnHeaders = lines[2] ? lines[2].map(h => h.toLowerCase().trim()) : [];
    parsedEmployeesData = {};
    
    let columnsPerEmployee = 7; 
    for (let i = 1; i < headerRow.length; i++) {
        if (headerRow[i] && headerRow[i].toUpperCase().includes('ATTENDANCE')) { columnsPerEmployee = i; break; }
    }

    for (let i = 0; i < headerRow.length; i += columnsPerEmployee) { 
        let empName = headerRow[i] ? headerRow[i].replace('ATTENDANCE', '').trim().toUpperCase() : "";
        if (!empName) continue;
        parsedEmployeesData[empName] = [];
        
        let otOffset = 5, allowanceOffset = -1; 
        for (let offset = 0; offset < columnsPerEmployee; offset++) {
            let hName = columnHeaders[i + offset] || "";
            if (hName.includes('allowance')) allowanceOffset = offset;
            if (hName.includes('of ot') || hName.includes('ot')) otOffset = offset;
        }

        for (let j = 3; j < lines.length; j++) {
            const row = lines[j]; if (!row || row.length <= i || !row[i] || row[i].trim() === "" || row[i].toLowerCase().includes('date')) continue; 
            let rawDate = row[i].trim(); let formattedDate = convertToHTMLDate(rawDate);
            const day = row[i+1] ? row[i+1].trim().toUpperCase() : "MONDAY";
            const shift = row[i+2] ? row[i+2].trim() : "PM SHIFT";
            const remarks = row[i+3] ? row[i+3].trim().toUpperCase() : "PRESENT";
            let dailyRate = parseFloat(row[i+4]) || 0;
            let allowance = allowanceOffset !== -1 && row[i + allowanceOffset] ? parseFloat(row[i + allowanceOffset]) || 0 : 0;
            let otHours = row[i + otOffset] ? parseFloat(row[i + otOffset]) || 0 : 0;
            
            if (dailyRate === 0 && remarks === "PRESENT") dailyRate = 650;
            if (dailyRate === 0 && remarks === "DOUBLE PAY") dailyRate = 1250;
            
            parsedEmployeesData[empName].push({ date: formattedDate, day, shift, remarks, dailyRate, allowance, otHours });
        }
    }
    refreshEditorUI();
}

function convertToHTMLDate(dateStr) {
    if(dateStr.includes('/')) {
        const parts = dateStr.split('/');
        if(parts.length === 3) return `${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`;
    }
    return dateStr;
}

function convertToCSVDate(dateStr) {
    if(dateStr.includes('-')) {
        const parts = dateStr.split('-');
        if(parts.length === 3) return `${parts[1]}/${parts[2]}/${parts[0]}`;
    }
    return dateStr;
}

function addNewEmployee() {
    const nameInput = document.getElementById('newEmployeeName'); const name = nameInput.value.trim().toUpperCase();
    if (name === "" || parsedEmployeesData[name]) return;
    parsedEmployeesData[name] = []; const ngayon = new Date();
    parsedEmployeesData[name].push({ date: ngayon.toISOString().split('T')[0], day: DAYS_OF_WEEK[ngayon.getDay()], shift: "PM SHIFT", remarks: "PRESENT", dailyRate: 650, allowance: 0, otHours: 0 });
    nameInput.value = ""; refreshEditorUI(); switchEmployeeTab(name);
    saveAppState();
}

function deleteEmployeeName(empName, event) {
    if (event) event.stopPropagation();
    if (confirm(`Sigurado ka bang nais mong burahin si ${empName}?`)) {
        delete parsedEmployeesData[empName]; const empNames = Object.keys(parsedEmployeesData);
        if (currentActiveEmployee === empName) currentActiveEmployee = empNames.length > 0 ? empNames[0] : "";
        if (empNames.length === 0) {
            // Keep rowActionsContainer and editorActions visible so user can still add employees
            document.getElementById('employeeTabs').innerHTML = "";
            document.getElementById('graphsSection').classList.add('hidden');
            document.getElementById('csvDataTable').classList.add('hidden');
            document.getElementById('computationSummary').classList.add('hidden');
        } else { refreshEditorUI(); }
        saveAppState();
    }
}
function refreshEditorUI() {
    const empNames = Object.keys(parsedEmployeesData); if(empNames.length === 0) return;
    // rowActionsContainer and editorActions are always visible (shown by default in HTML)
    document.getElementById('graphsSection').classList.remove('hidden');
    document.getElementById('graphsSection').classList.add('grid');
    renderEmployeeTabs();
}

function renderEmployeeTabs() {
    const tabsContainer = document.getElementById('employeeTabs'); tabsContainer.innerHTML = "";
    Object.keys(parsedEmployeesData).forEach((name, index) => {
        if (!currentActiveEmployee && index === 0) currentActiveEmployee = name;
        const isActive = currentActiveEmployee === name;
        const wrapper = document.createElement('div');
        wrapper.className = `inline-flex items-center rounded-lg overflow-hidden border ${isActive ? 'bg-orange-600 text-white border-orange-700 font-bold' : 'bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-slate-300 dark:border-slate-700 hover:bg-slate-200'}`;
        
        const tabBtn = document.createElement('button'); tabBtn.className = 'px-3 py-1.5 text-xs font-bold cursor-pointer'; tabBtn.innerText = name;
        tabBtn.onclick = () => switchEmployeeTab(name);
        
        const deleteBtn = document.createElement('button'); deleteBtn.className = `px-2 py-1.5 text-xs font-bold cursor-pointer ${isActive ? 'text-orange-200 hover:text-white' : 'text-slate-400 hover:text-red-500'}`;
        deleteBtn.innerHTML = '✕'; deleteBtn.onclick = (e) => deleteEmployeeName(name, e);
        
        wrapper.appendChild(tabBtn); wrapper.appendChild(deleteBtn); tabsContainer.appendChild(wrapper);
    });
    if (currentActiveEmployee) displayEmployeeData(currentActiveEmployee);
}

function switchEmployeeTab(empName) { currentActiveEmployee = empName; renderEmployeeTabs(); displayEmployeeData(empName); }

// 🔴 TINANGGAL ANG PETSA, ARAW, SIPET, AT GINAWANG "DELETE" ANG AKSYON DITO:
function displayEmployeeData(empName) {
    const table = document.getElementById('csvDataTable'); 
    
    let thead = table.querySelector('thead');
    if (!thead) {
        thead = document.createElement('thead');
        table.insertBefore(thead, table.firstChild);
    }
    thead.innerHTML = `
        <tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400 text-xs font-bold uppercase tracking-wider">
            <th class="p-3">Date</th>
            <th class="p-3">Day</th>
            <th class="p-3">Shift</th>
            <th class="p-3">Status / Remarks</th>
            <th class="p-3">Daily Rate (₱)</th>
            <th class="p-3">Allowance (₱)</th>
            <th class="p-3">OT (Hours)</th>
            <th class="p-3 text-center">DELETE</th>
        </tr>
    `;

    const tbody = document.getElementById('csvDataBody'); 
    tbody.innerHTML = ""; 
    
    const isDark = document.documentElement.classList.contains('dark');
    const inputBg = isDark ? "bg-slate-900 text-slate-100 border-slate-700" : "bg-white text-slate-800 border-slate-300";
    let attBuffer = [];
    
    (parsedEmployeesData[empName] || []).forEach((rec, index) => {
        let dayOptions = DAYS_OF_WEEK.map(d => `<option value="${d}" ${rec.day === d ? 'selected' : ''}>${d}</option>`).join('');
        
        attBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40">
            <td class="p-2"><input type="date" class="w-full text-xs p-1.5 ${inputBg} border rounded-lg focus:outline-none focus:ring-1 focus:ring-orange-500" value="${rec.date}" onchange="updateCell('${empName}', ${index}, 'date', this.value)"></td>
            <td class="p-2"><select class="w-full text-xs p-1.5 ${inputBg} border rounded-lg focus:outline-none focus:ring-1 focus:ring-orange-500">${dayOptions}</select></td>
            <td class="p-2"><select class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-bold focus:outline-none focus:ring-1 focus:ring-orange-500" onchange="updateCell('${empName}', ${index}, 'shift', this.value)"><option value="AM SHIFT" ${rec.shift === 'AM SHIFT' ? 'selected' : ''}>AM SHIFT</option><option value="PM SHIFT" ${rec.shift === 'PM SHIFT' ? 'selected' : ''}>PM SHIFT</option></select></td>
            <td class="p-2"><select class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-bold focus:outline-none focus:ring-1 focus:ring-orange-500" onchange="updateRemarksEvent('${empName}', ${index}, this)"><option value="PRESENT" ${rec.remarks === 'PRESENT' ? 'selected' : ''}>PRESENT</option><option value="ABSENT" ${rec.remarks === 'ABSENT' ? 'selected' : ''}>ABSENT</option><option value="RESTDAY" ${rec.remarks === 'RESTDAY' ? 'selected' : ''}>RESTDAY</option><option value="DOUBLE PAY" ${rec.remarks === 'DOUBLE PAY' ? 'selected' : ''}>DOUBLE PAY</option><option value="CDO" ${rec.remarks === 'CDO' ? 'selected' : ''}>CDO</option></select></td>
            <td class="p-2"><input type="number" class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-mono text-slate-700 dark:text-slate-300" value="${rec.dailyRate}" oninput="updateCell('${empName}', ${index}, 'dailyRate', this.value)"></td>
            <td class="p-2"><input type="number" class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-mono font-bold text-emerald-600 dark:text-emerald-400" value="${rec.allowance || 0}" oninput="updateCell('${empName}', ${index}, 'allowance', this.value)"></td>
            <td class="p-2"><input type="number" class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-mono text-slate-700 dark:text-slate-300" value="${rec.otHours}" oninput="updateCell('${empName}', ${index}, 'otHours', this.value)"></td>
            <td class="p-2 text-center"><button onclick="deleteRow('${empName}', ${index})" class="bg-red-500 text-white font-bold p-1.5 rounded-lg hover:bg-red-600 text-xs transition-colors cursor-pointer shadow-sm">🗑️</button></td>
        </tr>`);
    });
    
    tbody.innerHTML = attBuffer.join('');
    table.classList.remove('hidden'); 
    document.getElementById('computationSummary').classList.remove('hidden');
    calculateTotals(empName); 
    updateVisualGraphs(parsedEmployeesData[empName] || []);
}

function deleteRow(empName, index) { if (confirm("Nais mo bang burahin ang hanay na ito?")) { parsedEmployeesData[empName].splice(index, 1); displayEmployeeData(empName); saveAppState(); } }

// ====================================================================
// ⚙️ SYSTEM CALCULATIONS & GRAPH REFRESHERS
// ====================================================================
function updateCell(empName, index, field, value) {
    if (field === 'dailyRate' || field === 'otHours' || field === 'allowance') { parsedEmployeesData[empName][index][field] = Number(value) || 0; } 
    else { parsedEmployeesData[empName][index][field] = value; }
    if (field === 'date') {
        let d = new Date(value);
        if(!isNaN(d.getTime())) { parsedEmployeesData[empName][index]['day'] = DAYS_OF_WEEK[d.getDay()]; displayEmployeeData(empName); saveAppState(); return; }
    }
    calculateTotals(empName); updateVisualGraphs(parsedEmployeesData[empName]);
    saveAppStateDebounced();
}

function updateRemarksEvent(empName, index, selectElement) {
    const val = selectElement.value; parsedEmployeesData[empName][index]['remarks'] = val;
    if (val === 'PRESENT') parsedEmployeesData[empName][index]['dailyRate'] = 650;
    else if (val === 'DOUBLE PAY') parsedEmployeesData[empName][index]['dailyRate'] = 1250;
    else if (['ABSENT', 'RESTDAY', 'CDO'].includes(val)) { parsedEmployeesData[empName][index]['dailyRate'] = 0; parsedEmployeesData[empName][index]['allowance'] = 0; }
    displayEmployeeData(empName);
    saveAppState();
}

function addNewRowToCurrentEmployee() {
    if(!currentActiveEmployee) return; const records = parsedEmployeesData[currentActiveEmployee];
    let newDate = "2026-06-01", newDay = "MONDAY";
    if(records.length > 0) {
        let d = new Date(records[records.length - 1].date); d.setDate(d.getDate() + 1);
        if(!isNaN(d.getTime())) { newDate = d.toISOString().split('T')[0]; newDay = DAYS_OF_WEEK[d.getDay()]; }
    }
    parsedEmployeesData[currentActiveEmployee].push({ date: newDate, day: newDay, shift: "PM SHIFT", remarks: "PRESENT", dailyRate: 650, allowance: 0, otHours: 0 });
    displayEmployeeData(currentActiveEmployee);
    saveAppState();
}

function checkAttendance(status) {
    const nameInput = document.getElementById('employeeName'); const name = nameInput.value.trim().toUpperCase();
    if (name === "") { alert("Pakiusap, ilagay ang iyong pangalan."); return; }

    const defaultTime = new Date().toTimeString().substring(0, 5);
    _manualLogIdSeq++;
    manualAttendanceLogs.unshift({
        id: _manualLogIdSeq,
        name: name,
        date: new Date().toLocaleDateString('en-PH'),
        time: defaultTime,
        status: status
    });
    manualLogsCounter++;

    renderManualLogs();
    nameInput.value = "";
    saveAppState();
}

// Muling ginuhit ang buong Live Log Feed mula sa manualAttendanceLogs array.
// Ito ang gumagawang posible na ma-restore ang feed kapag nag-reload ang page.
function renderManualLogs() {
    const tbody = document.getElementById('attendanceLog');
    if (!tbody) return;

    if (manualAttendanceLogs.length === 0) {
        tbody.innerHTML = `<tr id="noRecordRow"><td colspan="5" class="py-12 text-center text-slate-400 dark:text-slate-500">Walang manual record sa ngayon.</td></tr>`;
        document.getElementById('totalLogs').innerText = manualLogsCounter;
        updateActiveDashboardCardDirectly();
        return;
    }

    const isDark = document.documentElement.classList.contains('dark');
    const selectBg = isDark ? "bg-slate-900 border-slate-700 text-slate-100" : "bg-white border-slate-300 text-slate-800";
    let buffer = [];
    manualAttendanceLogs.forEach(log => {
        buffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/30 transition-colors">
            <td class="p-2 font-bold text-slate-800 dark:text-slate-200">${log.name}</td>
            <td class="p-2 text-slate-600 dark:text-slate-400">${log.date}</td>
            <td class="p-2"><input type="time" class="${selectBg} border text-xs p-1 rounded" value="${log.time}" onchange="updateManualLogField(${log.id}, 'time', this.value)"></td>
            <td class="p-2"><select class="${selectBg} border text-xs p-1 rounded" onchange="updateManualLogField(${log.id}, 'status', this.value)"><option value="In" ${log.status === 'In' ? 'selected' : ''}>Time In</option><option value="Out" ${log.status === 'Out' ? 'selected' : ''}>Time Out</option></select></td>
            <td class="p-2 text-center"><button onclick="deleteManualLogRow(${log.id})" class="bg-red-500 text-white p-1 rounded text-[10px] cursor-pointer">🗑️</button></td>
        </tr>`);
    });

    tbody.innerHTML = buffer.join('');
    document.getElementById('totalLogs').innerText = manualLogsCounter;
    updateActiveDashboardCardDirectly();
}

// Para sa pag-edit ng time o status sa isang manual log row
function updateManualLogField(id, field, value) {
    const log = manualAttendanceLogs.find(l => l.id === id);
    if (log) { log[field] = value; }
    if (field === 'status') updateActiveDashboardCardDirectly();
    saveAppStateDebounced();
}

function deleteManualLogRow(id) {
    if (confirm("Nais mo bang burahin ang manual log na ito?")) {
        manualAttendanceLogs = manualAttendanceLogs.filter(l => l.id !== id);
        manualLogsCounter = Math.max(0, manualLogsCounter - 1);
        renderManualLogs();
        saveAppState();
    }
}

function updateActiveDashboardCardDirectly() {
    let activeInCount = 0; document.querySelectorAll('#attendanceLog select').forEach(sel => { if(sel.value === 'In') activeInCount++; });
    document.getElementById('activeUsers').innerText = activeInCount;
}

function calculateTotals(empName) {
    const records = parsedEmployeesData[empName] || []; 
    let daysPresent = 0, totalBasic = 0, totalAllowance = 0, totalOTHours = 0, estimatedOTPay = 0;
    
    records.forEach(rec => {
        if (rec.remarks === "PRESENT" || rec.remarks === "DOUBLE PAY") daysPresent++;
        totalBasic += Number(rec.dailyRate) || 0; totalAllowance += Number(rec.allowance) || 0; totalOTHours += Number(rec.otHours) || 0;
        if ((Number(rec.otHours) || 0) > 0) {
            let r = String(rec.shift).toUpperCase().includes("AM") ? 93 : 101;
            estimatedOTPay += r * (Number(rec.otHours) || 0);
        }
    });
    
    let grossSalary = totalBasic + totalAllowance + estimatedOTPay;
    document.getElementById('sumDaysPresent').innerText = daysPresent;
    document.getElementById('sumBasicSalary').innerText = "₱" + totalBasic.toLocaleString('en-US', {minimumFractionDigits: 2});
    document.getElementById('sumOTHours').innerText = totalOTHours + " hrs";
    document.getElementById('sumGrossSalary').innerText = "₱" + grossSalary.toLocaleString('en-US', {minimumFractionDigits: 2});
}

function updateVisualGraphs(records) {
    let remarksCount = { "PRESENT": 0, "ABSENT": 0, "RESTDAY": 0, "DOUBLE PAY": 0, "CDO": 0 };
    let labelsBar = [], dataRateBar = [], dataOTBar = [];
    records.forEach(rec => {
        if (remarksCount[rec.remarks] !== undefined) remarksCount[rec.remarks]++;
        labelsBar.push(rec.date.substring(5));
        dataRateBar.push(rec.dailyRate + (rec.allowance || 0));
        dataOTBar.push(rec.otHours);
    });

    const isDark = document.documentElement.classList.contains('dark');
    const lblColor = isDark ? '#94a3b8' : '#475569';

    // ✅ HIGHCHARTS 3D PIE — same style as dashboard pie chart
    const REMARK_COLORS = {
        "PRESENT":    '#10b981',  // green
        "ABSENT":     '#ef4444',  // red
        "RESTDAY":    '#f59e0b',  // amber
        "DOUBLE PAY": '#3b82f6',  // blue
        "CDO":        '#a855f7'   // purple
    };
    const REMARK_ICONS = {
        "PRESENT":    '✅',
        "ABSENT":     '❌',
        "RESTDAY":    '🌙',
        "DOUBLE PAY": '💰',
        "CDO":        '📋'
    };

    const pieData = Object.entries(remarksCount)
        .filter(([, count]) => count > 0)
        .map(([label, count]) => ({
            name: REMARK_ICONS[label] + ' ' + label,
            y: count,
            color: REMARK_COLORS[label]
        }));

    Highcharts.chart('percentageChart', {
        chart: {
            type: 'pie',
            options3d: { enabled: true, alpha: 45, beta: 0 },
            backgroundColor: 'transparent',
            margin: [30, 10, 60, 10],
            spacing: [4, 4, 4, 4]
        },
        title: {
            text: '📊 Attendance Breakdown',
            style: { fontSize: '11px', fontWeight: 'bold', color: lblColor },
            margin: 8
        },
        plotOptions: {
            pie: {
                depth: 35,
                allowPointSelect: true,
                cursor: 'pointer',
                center: ['50%', '48%'],
                size: '72%',
                dataLabels: {
                    enabled: true,
                    distance: 22,
                    formatter: function() {
                        return '<b>' + this.point.name + '</b><br>' + this.percentage.toFixed(0) + '%';
                    },
                    style: {
                        fontSize: '8px',
                        fontWeight: 'bold',
                        color: lblColor,
                        textOutline: 'none'
                    },
                    connectorWidth: 1,
                    connectorColor: lblColor
                },
                showInLegend: false
            }
        },
        credits: { enabled: false },
        series: [{ name: 'Days', data: pieData }]
    });

    // ✅ Bar chart stays using Chart.js (income + OT trend)
    if (chartInstanceBar3D) chartInstanceBar3D.destroy();
    const axisColor = isDark ? '#94a3b8' : '#475569';
    chartInstanceBar3D = new Chart(document.getElementById('trend3dChart').getContext('2d'), {
        type: 'bar',
        data: {
            labels: labelsBar,
            datasets: [
                { label: 'Income Base (₱)', data: dataRateBar, backgroundColor: '#3b82f6' },
                { label: 'OT Hours', data: dataOTBar, backgroundColor: '#f59e0b', yAxisID: 'y1' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y:  { position: 'left',  ticks: { color: axisColor } },
                y1: { position: 'right', grid: { drawOnChartArea: false }, ticks: { color: axisColor } },
                x:  { ticks: { color: axisColor } }
            },
            plugins: { legend: { labels: { color: axisColor } } }
        }
    });
}

// ====================================================================
// 📥 EXPORT EXTENSION UTILITIES
// ====================================================================
function exportToCSV() {
    const empNames = Object.keys(parsedEmployeesData); if(empNames.length === 0) return;
    let maxRows = 0; empNames.forEach(name => { if(parsedEmployeesData[name].length > maxRows) maxRows = parsedEmployeesData[name].length; });
    let csvLines = [];
    
    let r1 = []; empNames.forEach(n => r1.push(`${n} ATTENDANCE`, "", "", "", "", "", "", "")); csvLines.push(r1.join(','));
    let r2 = []; empNames.forEach(() => r2.push("UPDATED ATTENDANCE RECORD", "", "", "", "", "", "", "")); csvLines.push(r2.join(','));
    let r3 = []; empNames.forEach(() => r3.push("Date", "Day", "Shift", "Remarks", "Daily rate", "Allowance", "Number of OT", "")); csvLines.push(r3.join(','));

    for(let j = 0; j < maxRows; j++) {
        let dRow = [];
        empNames.forEach(name => {
            const rec = parsedEmployeesData[name][j];
            if(rec) dRow.push(convertToCSVDate(rec.date), rec.day, rec.shift, rec.remarks, rec.dailyRate, (rec.allowance || 0), rec.otHours, "");
            else dRow.push("", "", "", "", "", "", "", "");
        });
        csvLines.push(dRow.join(',')); 
    }

    const blob = new Blob([csvLines.join("\n")], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob); const link = document.createElement("a");
    link.setAttribute("href", url); link.setAttribute("download", "MTO_UPDATED_ATTENDANCE.csv");
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
}

function handleDashboardExport() {
    if (!globalData || globalData.length === 0) { alert("Walang data na pwedeng i-export sa ngayon. Pakiusap, mag-upload muna ng file."); return; }
    let csvLines = []; let fileName = "Dashboard_Export.csv";
    if (fileType === "progress") {
        csvLines.push("Work Week,Month,Warehouse,MTO Batch Name,Forecast,Allocated,Total Shipped,Discrepancy,Percentage");
        globalData.forEach(row => { csvLines.push(`"${row.ww}","${row.month}","${row.wh}","${row.batch}",${row.forecast},${row.allocated},${row.shipped},${row.discrepancy},${row.percentage}%`); });
        fileName = "MTO_Progress_Matrix_2026.csv";
    } else if (fileType === "raw") {
        csvLines.push("Box / MTB ID,Type,Series / Serial Number,Destination Warehouse,Qty,Status,Shipped Date,Remarks");
        globalData.forEach(row => { csvLines.push(`"${row.mtb}","${row.type}","${row.series}","${row.wh}",${row.qty},"${row.status}","${row.date}","${row.remarks || ''}"`); });
        fileName = "MTO_Raw_Pallet_Tracking.csv";
    } else if (fileType === "crossdock") {
        csvLines.push("Date,MTB ID,SKU Name,Req Qty,Act Qty,Pallet,Checked By,Status");
        globalData.forEach(row => { csvLines.push(`"${row.date}","${row.mtb}","${row.skuName}",${row.qty},${row.actualQty},"${row.pallet}","${row.checkedBy}","${row.statusReceived}"`); });
        fileName = "CrossDock_Transmittal_Hub.csv";
    } else if (fileType === "device") {
        csvLines.push("Device ID,Device Type,Specification(Rows),Specification(Columns),Status,Device Progress,Template ID");
        globalData.forEach(row => { csvLines.push(`"${row.deviceId}","${row.deviceType}","${row.specRows}","${row.specCols}","${row.status}","${row.progress}","${row.templateId}"`); });
        fileName = "Device_Management_Export.csv";
    } else if (fileType === "generic") {
        csvLines.push(genericHeaders.join(","));
        globalData.forEach(row => {
            let line = genericHeaders.map(h => `"${row[h] || ''}"`).join(",");
            csvLines.push(line);
        });
        fileName = "Custom_Dashboard_Export.csv";
    } else if (fileType === "cancelled") {
        csvLines.push("WHs,MT Order,Picking ID,SKU,SKU Name,Qty,PICKED QTY,CHECKED QTY,LACK ITEM,Reason,Picker");
        globalData.forEach(row => {
            csvLines.push(`"${row.whs}","${row.mtOrder}","${row.pickingId}","${row.sku}","${row.skuName.replace(/"/g,"'")}",${row.qty || 0},${row.pickedQty || ''},${row.checkedQty || ''},${row.lackItem || ''},"${row.reason}","${row.picker}"`);
        });
        fileName = "MTO_Cancelled_Monitoring_Export.csv";
    }
    
    const blob = new Blob(["\ufeff" + csvLines.join("\n")], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob); const link = document.createElement("a");
    link.setAttribute("href", url); link.setAttribute("download", fileName);
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
}