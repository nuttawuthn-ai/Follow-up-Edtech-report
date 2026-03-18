const CONFIG = {
  apiBaseUrl: 'https://script.google.com/macros/s/AKfycbz_WGKpm3WHqeHq61tNnDjyhpjYyxwAMf7tui3x6zfDv47i6BADWNDqRjMUVZAKVbJnmQ/exec',
  fallbackDemo: false,
  autoRefreshMs: 0,
  searchDebounceMs: 250
};

const monthNamesThai = {
  1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน',
  5: 'พฤษภาคม', 6: 'มิถุนายน', 7: 'กรกฎาคม', 8: 'สิงหาคม',
  9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
};

let appState = {
  dashboardData: null,
  availablePeriods: [],
  selectedYear: null,
  selectedMonth: null,
  search: '',
  autoRefreshTimer: null,
  searchTimer: null,
  requestToken: 0,
  dataCache: {},
  periodsLoaded: false
};

const el = {
  yearSelect:      document.getElementById('yearSelect'),
  monthSelect:     document.getElementById('monthSelect'),
  searchInput:     document.getElementById('searchInput'),
  refreshBtn:      document.getElementById('refreshBtn'),
  exportCsvBtn:    document.getElementById('exportCsvBtn'),
  lastUpdated:     document.getElementById('lastUpdated'),
  deadlineValue:   document.getElementById('deadlineValue'),
  normalCount:     document.getElementById('normalCount'),
  lateCount:       document.getElementById('lateCount'),
  missingCount:    document.getElementById('missingCount'),
  totalCount:      document.getElementById('totalCount'),
  summaryLabel:    document.getElementById('summaryLabel'),
  personTableBody: document.getElementById('personTableBody'),
  simpleChart:     document.getElementById('simpleChart'),
  statusText:      document.getElementById('statusText'),
  asOfDate:        document.getElementById('asOfDate'),
  ruleText:        document.getElementById('ruleText'),
  // % bars
  normalPctFill:   document.getElementById('normalPctFill'),
  normalPctLabel:  document.getElementById('normalPctLabel'),
  latePctFill:     document.getElementById('latePctFill'),
  latePctLabel:    document.getElementById('latePctLabel'),
  missingPctFill:  document.getElementById('missingPctFill'),
  missingPctLabel: document.getElementById('missingPctLabel'),
  // completion bar
  completionPct:   document.getElementById('completionPct'),
  compSegNormal:   document.getElementById('compSegNormal'),
  compSegLate:     document.getElementById('compSegLate'),
  compSegMissing:  document.getElementById('compSegMissing'),
  legendNormal:    document.getElementById('legendNormal'),
  legendLate:      document.getElementById('legendLate'),
  legendMissing:   document.getElementById('legendMissing')
};

// ─── Helpers ───────────────────────────────────────────────

function periodKey(year, month) { return `${year}-${month}`; }

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;').replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;').replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function formatDateTime(value) {
  if (!value) return '-';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString('th-TH', {
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
    timeZone: 'Asia/Bangkok'
  });
}

function normalizeStatus(status) {
  const s = String(status || '').trim().toLowerCase();
  if (s === 'ปกติ' || s === 'normal') return 'ปกติ';
  if (s === 'ล่าช้า' || s === 'late' || s === 'delay' || s === 'delayed') return 'ล่าช้า';
  if (['ยังไม่ส่ง','ยังไม่ได้ส่ง','missing','not submitted','pending'].includes(s)) return 'ยังไม่ส่ง';
  return 'ยังไม่ส่ง';
}

function getBadgeClass(status) {
  const s = normalizeStatus(status);
  if (s === 'ปกติ') return 'normal';
  if (s === 'ล่าช้า') return 'late';
  return 'missing';
}

function getRowClass(status) {
  const s = normalizeStatus(status);
  if (s === 'ปกติ') return 'row-normal';
  if (s === 'ล่าช้า') return 'row-late';
  return 'row-missing';
}

function pct(count, total) {
  if (!total) return 0;
  return Math.round((count / total) * 100);
}

function setStatus(message)      { if (el.statusText)   el.statusText.textContent = message; }
function setLastUpdated(message) { if (el.lastUpdated)  el.lastUpdated.textContent = message; }

function createOptions(select, values, formatter = v => v) {
  if (!select) return;
  select.innerHTML = '';
  values.forEach(value => {
    const opt = document.createElement('option');
    opt.value = String(value);
    opt.textContent = formatter(value);
    select.appendChild(opt);
  });
}

// ─── Skeleton loading ────────────────────────────────────────

function showSkeletonLoading() {
  // stat cards: แสดง skeleton แทนตัวเลข
  [el.normalCount, el.lateCount, el.missingCount, el.totalCount].forEach(el => {
    if (el) el.innerHTML = '<span class="skeleton"></span>';
  });
  // reset % bars
  [el.normalPctFill, el.latePctFill, el.missingPctFill].forEach(fill => {
    if (fill) fill.style.width = '0%';
  });
  [el.normalPctLabel, el.latePctLabel, el.missingPctLabel].forEach(lbl => {
    if (lbl) lbl.textContent = '-';
  });
  // completion bar
  if (el.compSegNormal)  el.compSegNormal.style.width  = '0%';
  if (el.compSegLate)    el.compSegLate.style.width    = '0%';
  if (el.compSegMissing) el.compSegMissing.style.width = '0%';
  if (el.completionPct)  el.completionPct.textContent  = '-';
  // table
  if (el.personTableBody) {
    el.personTableBody.innerHTML = `
      <tr><td colspan="6" style="padding:16px 10px;">
        <span class="skeleton" style="width:40%;height:14px;margin-bottom:10px;"></span>
        <span class="skeleton" style="width:70%;height:14px;margin-bottom:10px;"></span>
        <span class="skeleton" style="width:55%;height:14px;"></span>
      </td></tr>`;
  }
}

// ─── API ─────────────────────────────────────────────────────

async function fetchJson(url) {
  const res  = await fetch(url, { method: 'GET', cache: 'no-store' });
  const text = await res.text();
  let json = null;
  try { json = JSON.parse(text); } catch { throw new Error(`API ตอบกลับไม่ใช่ JSON: ${text.slice(0, 200)}`); }
  if (!res.ok)    throw new Error(`โหลดข้อมูลไม่สำเร็จ: ${res.status}`);
  if (!json.ok)   throw new Error(json.message || 'API error');
  return json;
}

async function fetchAvailablePeriods() {
  const url  = `${CONFIG.apiBaseUrl}?action=years-months&_ts=${Date.now()}`;
  const json = await fetchJson(url);
  return Array.isArray(json.data) ? json.data : [];
}

async function fetchDashboardData(year, month) {
  const url  = `${CONFIG.apiBaseUrl}?action=dashboard-data&year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}&_ts=${Date.now()}`;
  const json = await fetchJson(url);
  return json.data;
}

// ─── Period helpers ──────────────────────────────────────────

function getUniqueYears(periods) {
  return [...new Set(periods.map(p => Number(p.year)))].sort((a, b) => b - a);
}

function getMonthsForYear(periods, year) {
  return periods
    .filter(p => Number(p.year) === Number(year))
    .map(p => Number(p.month))
    .sort((a, b) => a - b);
}

function buildFallbackPeriods() {
  const now = new Date();
  return [{ year: now.getFullYear(), month: now.getMonth() + 1, monthThai: monthNamesThai[now.getMonth() + 1] || '' }];
}

function ensureSelectedPeriod() {
  if (!appState.availablePeriods.length) appState.availablePeriods = buildFallbackPeriods();
  const years = getUniqueYears(appState.availablePeriods);
  if (!years.length) { const n = new Date(); appState.selectedYear = n.getFullYear(); appState.selectedMonth = n.getMonth() + 1; return; }
  if (!appState.selectedYear || !years.includes(Number(appState.selectedYear))) appState.selectedYear = years[0];
  const months = getMonthsForYear(appState.availablePeriods, appState.selectedYear);
  if (!months.length) { appState.selectedMonth = 1; return; }
  if (!appState.selectedMonth || !months.includes(Number(appState.selectedMonth))) appState.selectedMonth = months[months.length - 1] || months[0];
}

function renderPeriodSelectors() {
  ensureSelectedPeriod();
  const years  = getUniqueYears(appState.availablePeriods);
  const months = getMonthsForYear(appState.availablePeriods, appState.selectedYear);
  createOptions(el.yearSelect,  years,  y => String(Number(y) + 543));
  createOptions(el.monthSelect, months, m => `${m} - ${monthNamesThai[m] || ''}`);
  if (el.yearSelect)  el.yearSelect.value  = String(appState.selectedYear);
  if (el.monthSelect) el.monthSelect.value = String(appState.selectedMonth);
}

// ─── Render ──────────────────────────────────────────────────

function renderChart(normalCount, lateCount, missingCount, total) {
  if (!el.simpleChart) return;
  const bars = [
    { label: 'ปกติ',     value: normalCount,  cls: 'success' },
    { label: 'ล่าช้า',   value: lateCount,    cls: 'danger'  },
    { label: 'ยังไม่ส่ง', value: missingCount, cls: 'warning' }
  ];
  el.simpleChart.innerHTML = '';
  bars.forEach(item => {
    const p   = total ? Math.round((item.value / total) * 100) : 0;
    const row = document.createElement('div');
    row.className = 'bar-row';
    row.innerHTML = `
      <div>${item.label}</div>
      <div class="bar-track"><div class="bar-fill ${item.cls}" style="width:${p}%"></div></div>
      <div>${item.value}</div>`;
    el.simpleChart.appendChild(row);
  });
}

function renderTable(records) {
  if (!el.personTableBody) return;
  el.personTableBody.innerHTML = '';
  if (!records.length) {
    el.personTableBody.innerHTML = '<tr><td colspan="6" class="empty-note">ไม่พบข้อมูล</td></tr>';
    return;
  }
  records.forEach(item => {
    const tr = document.createElement('tr');
    tr.className = getRowClass(item.status);   // ★ row highlight
    const fileCell = item.fileUrl
      ? `<a class="file-link" href="${item.fileUrl}" target="_blank" rel="noopener">เปิดไฟล์</a>`
      : '-';
    tr.innerHTML = `
      <td>${escapeHtml(item.name || '-')}</td>
      <td><span class="badge ${getBadgeClass(item.status)}">${escapeHtml(normalizeStatus(item.status))}</span></td>
      <td>${escapeHtml(item.submittedAtDisplay || formatDateTime(item.submittedAt || ''))}</td>
      <td>${escapeHtml(item.deadline || '-')}</td>
      <td>${fileCell}</td>
      <td>${escapeHtml(item.note || '-')}</td>`;
    el.personTableBody.appendChild(tr);
  });
}

function normalizeRows(rows) {
  return (Array.isArray(rows) ? rows : []).map(r => ({
    ...r,
    name: String(r.name || '').trim(),
    status: normalizeStatus(r.status),
    submittedAt: r.submittedAt || '',
    submittedAtDisplay: r.submittedAtDisplay || '',
    deadline: r.deadline || '',
    fileUrl: r.fileUrl || '',
    note: r.note || ''
  }));
}

function getCountsFromApiOrRows(payload, filteredRows, isSearching) {
  const apiCounts = payload?.counts || {};
  const rows = Array.isArray(filteredRows) ? filteredRows : [];
  if (!isSearching) {
    return {
      normal:  Number.isFinite(Number(apiCounts.normal))  ? Number(apiCounts.normal)  : rows.filter(r => r.status === 'ปกติ').length,
      late:    Number.isFinite(Number(apiCounts.late))    ? Number(apiCounts.late)    : rows.filter(r => r.status === 'ล่าช้า').length,
      missing: Number.isFinite(Number(apiCounts.missing)) ? Number(apiCounts.missing) : rows.filter(r => r.status === 'ยังไม่ส่ง').length,
      total:   Number.isFinite(Number(apiCounts.total))   ? Number(apiCounts.total)   : rows.length
    };
  }
  return {
    normal:  rows.filter(r => r.status === 'ปกติ').length,
    late:    rows.filter(r => r.status === 'ล่าช้า').length,
    missing: rows.filter(r => r.status === 'ยังไม่ส่ง').length,
    total:   rows.length
  };
}

// ★ อัปเดต % bars ใน stat cards
function renderStatPctBars(counts) {
  const { normal, late, missing, total } = counts;
  const pNormal  = pct(normal,  total);
  const pLate    = pct(late,    total);
  const pMissing = pct(missing, total);

  if (el.normalPctFill)   el.normalPctFill.style.width   = `${pNormal}%`;
  if (el.normalPctLabel)  el.normalPctLabel.textContent   = `${pNormal}% ของทั้งหมด`;
  if (el.latePctFill)     el.latePctFill.style.width      = `${pLate}%`;
  if (el.latePctLabel)    el.latePctLabel.textContent     = `${pLate}% ของทั้งหมด`;
  if (el.missingPctFill)  el.missingPctFill.style.width   = `${pMissing}%`;
  if (el.missingPctLabel) el.missingPctLabel.textContent  = `${pMissing}% ของทั้งหมด`;
}

// ★ อัปเดต completion bar
function renderCompletionBar(counts) {
  const { normal, late, missing, total } = counts;
  const pNormal  = pct(normal,  total);
  const pLate    = pct(late,    total);
  const pMissing = pct(missing, total);

  if (el.compSegNormal)  el.compSegNormal.style.width  = `${pNormal}%`;
  if (el.compSegLate)    el.compSegLate.style.width    = `${pLate}%`;
  if (el.compSegMissing) el.compSegMissing.style.width = `${pMissing}%`;
  if (el.completionPct)  el.completionPct.textContent  = `${pNormal}%`;
  if (el.legendNormal)   el.legendNormal.textContent   = `${pNormal}%`;
  if (el.legendLate)     el.legendLate.textContent     = `${pLate}%`;
  if (el.legendMissing)  el.legendMissing.textContent  = `${pMissing}%`;
}

function renderFilteredView() {
  if (!appState.dashboardData) return;

  const search       = (appState.search || '').toLowerCase();
  const rows         = normalizeRows(appState.dashboardData.rows);
  const isSearching  = Boolean(search);
  const filteredRows = rows.filter(r => !search || String(r.name || '').toLowerCase().includes(search));
  const counts       = getCountsFromApiOrRows(appState.dashboardData, filteredRows, isSearching);

  if (el.deadlineValue) el.deadlineValue.textContent = appState.dashboardData.deadline || '-';
  if (el.normalCount)   el.normalCount.textContent   = counts.normal;
  if (el.lateCount)     el.lateCount.textContent     = counts.late;
  if (el.missingCount)  el.missingCount.textContent  = counts.missing;
  if (el.totalCount)    el.totalCount.textContent    = counts.total;
  if (el.asOfDate)      el.asOfDate.textContent      = new Date().toLocaleString('th-TH');
  if (el.ruleText && appState.dashboardData.ruleText) el.ruleText.textContent = appState.dashboardData.ruleText;

  if (el.summaryLabel) {
    el.summaryLabel.textContent =
      appState.dashboardData.reportLabel ||
      `${monthNamesThai[appState.selectedMonth] || appState.selectedMonth} ${Number(appState.selectedYear) + 543}`;
  }

  renderStatPctBars(counts);     // ★ % bars
  renderCompletionBar(counts);   // ★ completion bar
  renderTable(filteredRows);     // ★ row highlight อยู่ใน renderTable
  renderChart(counts.normal, counts.late, counts.missing, counts.total);
}

// ─── Cache + load ─────────────────────────────────────────────

async function loadDashboardForSelectedPeriod() {
  const token = ++appState.requestToken;
  const key   = periodKey(appState.selectedYear, appState.selectedMonth);
  if (appState.dataCache[key]) {
    if (token !== appState.requestToken) return false;
    appState.dashboardData = appState.dataCache[key];
    return true;
  }
  const data = await fetchDashboardData(appState.selectedYear, appState.selectedMonth);
  if (token !== appState.requestToken) return false;
  appState.dataCache[key] = data;
  appState.dashboardData  = data;
  return true;
}

async function switchPeriod() {
  try {
    const key      = periodKey(appState.selectedYear, appState.selectedMonth);
    const isCached = Boolean(appState.dataCache[key]);

    if (!isCached) {
      // ★ แสดง skeleton ระหว่างรอข้อมูลใหม่
      setStatus('สถานะระบบ: กำลังโหลดข้อมูล');
      showSkeletonLoading();
    }

    const loaded = await loadDashboardForSelectedPeriod();
    if (loaded === false) return;

    setLastUpdated(`อัปเดตล่าสุด: ${new Date().toLocaleString('th-TH')}`);
    setStatus('สถานะระบบ: พร้อมใช้งาน');
    renderFilteredView();
  } catch (error) {
    console.error(error);
    setStatus('สถานะระบบ: โหลดข้อมูลไม่สำเร็จ');
    setLastUpdated(error.message || 'Unknown error');
  }
}

// ─── Export CSV ──────────────────────────────────────────────

function exportCsv() {
  if (!appState.dashboardData) return;

  const rows     = normalizeRows(appState.dashboardData.rows);
  const label    = appState.dashboardData.reportLabel ||
                   `${monthNamesThai[appState.selectedMonth] || appState.selectedMonth}_${Number(appState.selectedYear) + 543}`;
  const headers  = ['ชื่อ', 'สถานะ', 'วันที่ส่งจริง', 'วันกำหนดส่ง', 'ไฟล์รายงาน', 'หมายเหตุ'];

  const csvRows  = [
    headers.join(','),
    ...rows.map(r => [
      `"${(r.name || '').replace(/"/g, '""')}"`,
      `"${normalizeStatus(r.status)}"`,
      `"${r.submittedAtDisplay || formatDateTime(r.submittedAt) || ''}"`,
      `"${r.deadline || ''}"`,
      `"${r.fileUrl || ''}"`,
      `"${(r.note || '').replace(/"/g, '""')}"`
    ].join(','))
  ];

  const bom  = '\uFEFF';
  const blob = new Blob([bom + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `รายงาน_${label}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// ─── Events ──────────────────────────────────────────────────

function debounce(fn, wait) {
  return function(...args) {
    clearTimeout(appState.searchTimer);
    appState.searchTimer = setTimeout(() => fn.apply(this, args), wait);
  };
}

function bindEvents() {
  if (el.yearSelect) {
    el.yearSelect.addEventListener('change', async () => {
      appState.selectedYear = Number(el.yearSelect.value);
      renderPeriodSelectors();
      await switchPeriod();
    });
  }

  if (el.monthSelect) {
    el.monthSelect.addEventListener('change', async () => {
      appState.selectedMonth = Number(el.monthSelect.value);
      await switchPeriod();
    });
  }

  if (el.searchInput) {
    const debouncedSearch = debounce(() => {
      appState.search = el.searchInput.value.trim();
      renderFilteredView();
    }, CONFIG.searchDebounceMs);
    el.searchInput.addEventListener('input', debouncedSearch);
  }

  if (el.refreshBtn) {
    el.refreshBtn.addEventListener('click', async () => {
      appState.dataCache    = {};
      appState.periodsLoaded = false;
      await initApp(true, true);
    });
  }

  if (el.exportCsvBtn) {
    el.exportCsvBtn.addEventListener('click', exportCsv);
  }
}

function startAutoRefresh() {
  if (appState.autoRefreshTimer) clearInterval(appState.autoRefreshTimer);
  if (CONFIG.autoRefreshMs > 0) {
    appState.autoRefreshTimer = setInterval(async () => {
      try {
        const key = periodKey(appState.selectedYear, appState.selectedMonth);
        delete appState.dataCache[key];
        await switchPeriod();
      } catch (err) { console.error('Auto refresh error:', err); }
    }, CONFIG.autoRefreshMs);
  }
}

// ─── Init ─────────────────────────────────────────────────────

async function initApp(isRefresh = false, keepCurrentSelection = false) {
  try {
    setStatus('สถานะระบบ: กำลังโหลดข้อมูล');
    if (isRefresh) setLastUpdated('กำลังรีเฟรช...');

    if (!appState.periodsLoaded || isRefresh) {
      const periods = await fetchAvailablePeriods();
      appState.availablePeriods = periods.length ? periods : buildFallbackPeriods();
      appState.periodsLoaded    = true;
    }

    if (!keepCurrentSelection) { appState.selectedYear = null; appState.selectedMonth = null; }

    ensureSelectedPeriod();
    renderPeriodSelectors();

    showSkeletonLoading();
    const loaded = await loadDashboardForSelectedPeriod();
    if (loaded === false) return;

    setLastUpdated(`อัปเดตล่าสุด: ${new Date().toLocaleString('th-TH')}`);
    setStatus('สถานะระบบ: พร้อมใช้งาน');
    renderFilteredView();
  } catch (error) {
    console.error(error);
    setStatus('สถานะระบบ: โหลดข้อมูลไม่สำเร็จ');
    setLastUpdated(error.message || 'Unknown error');
  }
}

bindEvents();
startAutoRefresh();
initApp().catch(console.error);
