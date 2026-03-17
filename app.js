const CONFIG = {
  apiBaseUrl: 'PASTE_YOUR_WEBAPP_URL_HERE',
  fallbackDemo: false,
  autoRefreshMs: 0,
  searchDebounceMs: 250
};

const monthNamesThai = {
  1: 'มกราคม',
  2: 'กุมภาพันธ์',
  3: 'มีนาคม',
  4: 'เมษายน',
  5: 'พฤษภาคม',
  6: 'มิถุนายน',
  7: 'กรกฎาคม',
  8: 'สิงหาคม',
  9: 'กันยายน',
  10: 'ตุลาคม',
  11: 'พฤศจิกายน',
  12: 'ธันวาคม'
};

let appState = {
  dashboardData: null,
  availablePeriods: [],
  selectedYear: null,
  selectedMonth: null,
  search: '',
  autoRefreshTimer: null,
  searchTimer: null,
  requestToken: 0
};

const el = {
  yearSelect: document.getElementById('yearSelect'),
  monthSelect: document.getElementById('monthSelect'),
  searchInput: document.getElementById('searchInput'),
  refreshBtn: document.getElementById('refreshBtn'),
  runSystemCheckBtn: document.getElementById('runSystemCheckBtn'),
  lastUpdated: document.getElementById('lastUpdated'),
  deadlineValue: document.getElementById('deadlineValue'),
  normalCount: document.getElementById('normalCount'),
  lateCount: document.getElementById('lateCount'),
  missingCount: document.getElementById('missingCount'),
  totalCount: document.getElementById('totalCount'),
  summaryLabel: document.getElementById('summaryLabel'),
  personTableBody: document.getElementById('personTableBody'),
  simpleChart: document.getElementById('simpleChart'),
  statusText: document.getElementById('statusText'),
  asOfDate: document.getElementById('asOfDate'),
  ruleText: document.getElementById('ruleText'),
  diagPageUrl: document.getElementById('diagPageUrl'),
  diagApiUrl: document.getElementById('diagApiUrl'),
  diagOverallStatus: document.getElementById('diagOverallStatus'),
  systemCheckBody: document.getElementById('systemCheckBody')
};

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function formatDateTime(value) {
  if (!value) return '-';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString('th-TH', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    timeZone: 'Asia/Bangkok'
  });
}

function normalizeStatus(status) {
  const s = String(status || '').trim().toLowerCase();

  if (s === 'ปกติ' || s === 'normal') return 'ปกติ';
  if (s === 'ล่าช้า' || s === 'late' || s === 'delay' || s === 'delayed') return 'ล่าช้า';
  if (
    s === 'ยังไม่ส่ง' ||
    s === 'ยังไม่ได้ส่ง' ||
    s === 'missing' ||
    s === 'not submitted' ||
    s === 'pending'
  ) return 'ยังไม่ส่ง';

  return 'ยังไม่ส่ง';
}

function getBadgeClass(status) {
  const s = normalizeStatus(status);
  if (s === 'ปกติ') return 'normal';
  if (s === 'ล่าช้า') return 'late';
  return 'missing';
}

function setStatus(message) {
  if (el.statusText) el.statusText.textContent = message;
}

function setLastUpdated(message) {
  if (el.lastUpdated) el.lastUpdated.textContent = message;
}

function createOptions(select, values, formatter = v => v) {
  if (!select) return;
  select.innerHTML = '';

  values.forEach(value => {
    const option = document.createElement('option');
    option.value = String(value);
    option.textContent = formatter(value);
    select.appendChild(option);
  });
}

async function fetchJson(url) {
  const res = await fetch(url, {
    method: 'GET',
    cache: 'no-store'
  });

  const text = await res.text();
  let json = null;

  try {
    json = JSON.parse(text);
  } catch (err) {
    throw new Error(`API ตอบกลับไม่ใช่ JSON: ${text.slice(0, 200)}`);
  }

  if (!res.ok) {
    throw new Error(`โหลดข้อมูลไม่สำเร็จ: ${res.status}`);
  }

  if (!json.ok) {
    throw new Error(json.message || 'API error');
  }

  return json;
}

async function fetchAvailablePeriods() {
  if (!CONFIG.apiBaseUrl || CONFIG.apiBaseUrl.includes('https://script.google.com/macros/s/AKfycbz_WGKpm3WHqeHq61tNnDjyhpjYyxwAMf7tui3x6zfDv47i6BADWNDqRjMUVZAKVbJnmQ/exec')) {
    throw new Error('ยังไม่ได้ตั้งค่า Apps Script Web App URL');
  }

  const url = `${CONFIG.apiBaseUrl}?action=years-months&_ts=${Date.now()}`;
  const json = await fetchJson(url);
  return Array.isArray(json.data) ? json.data : [];
}

async function fetchDashboardData(year, month) {
  if (!CONFIG.apiBaseUrl || CONFIG.apiBaseUrl.includes('https://script.google.com/macros/s/AKfycbz_WGKpm3WHqeHq61tNnDjyhpjYyxwAMf7tui3x6zfDv47i6BADWNDqRjMUVZAKVbJnmQ/exec')) {
    throw new Error('ยังไม่ได้ตั้งค่า Apps Script Web App URL');
  }

  const url = `${CONFIG.apiBaseUrl}?action=dashboard-data&year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}&_ts=${Date.now()}`;
  const json = await fetchJson(url);
  return json.data;
}

function getUniqueYears(periods) {
  return [...new Set(periods.map(item => Number(item.year)))].sort((a, b) => b - a);
}

function getMonthsForYear(periods, year) {
  return periods
    .filter(item => Number(item.year) === Number(year))
    .map(item => Number(item.month))
    .sort((a, b) => a - b);
}

function buildFallbackPeriods() {
  const now = new Date();
  return [{
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    monthThai: monthNamesThai[now.getMonth() + 1] || ''
  }];
}

function ensureSelectedPeriod() {
  if (!appState.availablePeriods.length) {
    appState.availablePeriods = buildFallbackPeriods();
  }

  const years = getUniqueYears(appState.availablePeriods);

  if (!years.length) {
    const now = new Date();
    appState.selectedYear = now.getFullYear();
    appState.selectedMonth = now.getMonth() + 1;
    return;
  }

  if (!appState.selectedYear || !years.includes(Number(appState.selectedYear))) {
    appState.selectedYear = years[0];
  }

  const months = getMonthsForYear(appState.availablePeriods, appState.selectedYear);

  if (!months.length) {
    appState.selectedMonth = 1;
    return;
  }

  if (!appState.selectedMonth || !months.includes(Number(appState.selectedMonth))) {
    appState.selectedMonth = months[months.length - 1] || months[0];
  }
}

function renderPeriodSelectors() {
  ensureSelectedPeriod();

  const years = getUniqueYears(appState.availablePeriods);
  const months = getMonthsForYear(appState.availablePeriods, appState.selectedYear);

  createOptions(el.yearSelect, years, y => String(Number(y) + 543));
  createOptions(el.monthSelect, months, m => `${m} - ${monthNamesThai[m] || ''}`);

  if (el.yearSelect) el.yearSelect.value = String(appState.selectedYear);
  if (el.monthSelect) el.monthSelect.value = String(appState.selectedMonth);
}

function renderChart(normalCount, lateCount, missingCount, total) {
  if (!el.simpleChart) return;

  const bars = [
    { label: 'ปกติ', value: normalCount, cls: 'success' },
    { label: 'ล่าช้า', value: lateCount, cls: 'danger' },
    { label: 'ยังไม่ส่ง', value: missingCount, cls: 'warning' }
  ];

  el.simpleChart.innerHTML = '';

  bars.forEach(item => {
    const pct = total ? Math.round((item.value / total) * 100) : 0;
    const row = document.createElement('div');
    row.className = 'bar-row';
    row.innerHTML = `
      <div>${item.label}</div>
      <div class="bar-track">
        <div class="bar-fill ${item.cls}" style="width:${pct}%"></div>
      </div>
      <div>${item.value}</div>
    `;
    el.simpleChart.appendChild(row);
  });
}

function renderTable(records) {
  if (!el.personTableBody) return;

  el.personTableBody.innerHTML = '';

  records.forEach(item => {
    const tr = document.createElement('tr');
    const fileCell = item.fileUrl
      ? `<a class="file-link" href="${item.fileUrl}" target="_blank" rel="noopener">เปิดไฟล์</a>`
      : '-';

    tr.innerHTML = `
      <td>${escapeHtml(item.name || '-')}</td>
      <td><span class="badge ${getBadgeClass(item.status)}">${escapeHtml(normalizeStatus(item.status))}</span></td>
      <td>${escapeHtml(item.submittedAtDisplay || formatDateTime(item.submittedAt || ''))}</td>
      <td>${escapeHtml(item.deadline || '-')}</td>
      <td>${fileCell}</td>
      <td>${escapeHtml(item.note || '-')}</td>
    `;

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
    const normal = Number.isFinite(Number(apiCounts.normal))
      ? Number(apiCounts.normal)
      : rows.filter(r => r.status === 'ปกติ').length;

    const late = Number.isFinite(Number(apiCounts.late))
      ? Number(apiCounts.late)
      : rows.filter(r => r.status === 'ล่าช้า').length;

    const missing = Number.isFinite(Number(apiCounts.missing))
      ? Number(apiCounts.missing)
      : rows.filter(r => r.status === 'ยังไม่ส่ง').length;

    const total = Number.isFinite(Number(apiCounts.total))
      ? Number(apiCounts.total)
      : rows.length;

    return { normal, late, missing, total };
  }

  return {
    normal: rows.filter(r => r.status === 'ปกติ').length,
    late: rows.filter(r => r.status === 'ล่าช้า').length,
    missing: rows.filter(r => r.status === 'ยังไม่ส่ง').length,
    total: rows.length
  };
}

function renderFilteredView() {
  if (!appState.dashboardData) return;

  const search = (appState.search || '').toLowerCase();
  const rows = normalizeRows(appState.dashboardData.rows);
  const isSearching = Boolean(search);

  const filteredRows = rows.filter(r => {
    if (!search) return true;
    return String(r.name || '').toLowerCase().includes(search);
  });

  const counts = getCountsFromApiOrRows(appState.dashboardData, filteredRows, isSearching);

  if (el.deadlineValue) el.deadlineValue.textContent = appState.dashboardData.deadline || '-';
  if (el.normalCount) el.normalCount.textContent = counts.normal;
  if (el.lateCount) el.lateCount.textContent = counts.late;
  if (el.missingCount) el.missingCount.textContent = counts.missing;
  if (el.totalCount) el.totalCount.textContent = counts.total;
  if (el.asOfDate) el.asOfDate.textContent = new Date().toLocaleString('th-TH');
  if (el.ruleText && appState.dashboardData.ruleText) el.ruleText.textContent = appState.dashboardData.ruleText;

  if (el.summaryLabel) {
    el.summaryLabel.textContent =
      appState.dashboardData.reportLabel ||
      `${monthNamesThai[appState.selectedMonth] || appState.selectedMonth} ${Number(appState.selectedYear) + 543}`;
  }

  renderTable(filteredRows);
  renderChart(counts.normal, counts.late, counts.missing, counts.total);
}

async function loadDashboardForSelectedPeriod() {
  const token = ++appState.requestToken;
  const data = await fetchDashboardData(appState.selectedYear, appState.selectedMonth);

  if (token !== appState.requestToken) return false;
  appState.dashboardData = data;
  return true;
}

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
      await initApp(false, true);
    });
  }

  if (el.monthSelect) {
    el.monthSelect.addEventListener('change', async () => {
      appState.selectedMonth = Number(el.monthSelect.value);
      await initApp(false, true);
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
      await initApp(true, true);
    });
  }

  if (el.runSystemCheckBtn) {
    el.runSystemCheckBtn.addEventListener('click', runSystemCheck);
  }
}

function startAutoRefresh() {
  if (appState.autoRefreshTimer) clearInterval(appState.autoRefreshTimer);

  if (CONFIG.autoRefreshMs > 0) {
    appState.autoRefreshTimer = setInterval(async () => {
      try {
        await initApp(true, true);
      } catch (err) {
        console.error('Auto refresh error:', err);
      }
    }, CONFIG.autoRefreshMs);
  }
}

function diagStatusClass(status) {
  if (status === 'PASS') return 'diag-ok';
  if (status === 'WARN') return 'diag-warn';
  if (status === 'FAIL') return 'diag-err';
  return '';
}

function renderSystemCheckRows(rows) {
  if (!el.systemCheckBody) return;
  el.systemCheckBody.innerHTML = rows.map(row => `
    <tr>
      <td>${escapeHtml(row.name)}</td>
      <td class="${diagStatusClass(row.status)}">${escapeHtml(row.status)}</td>
      <td>${escapeHtml(row.detail)}</td>
      <td>${escapeHtml(row.ms)}</td>
    </tr>
  `).join('');
}

async function timedFetchJson(url, timeoutMs = 15000) {
  const started = performance.now();
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, {
      method: 'GET',
      cache: 'no-store',
      signal: controller.signal
    });

    const text = await res.text();
    let json = null;

    try {
      json = JSON.parse(text);
    } catch (e) {
      json = null;
    }

    return {
      ok: res.ok,
      status: res.status,
      elapsedMs: Math.round(performance.now() - started),
      text,
      json
    };
  } finally {
    clearTimeout(timer);
  }
}

async function runSystemCheck() {
  if (el.diagPageUrl) el.diagPageUrl.textContent = window.location.href;
  if (el.diagApiUrl) el.diagApiUrl.textContent = CONFIG.apiBaseUrl || '-';
  if (el.diagOverallStatus) el.diagOverallStatus.textContent = 'กำลังตรวจ...';

  const rows = [];
  const pushRow = (name, status, detail, ms = '-') => {
    rows.push({ name, status, detail, ms: typeof ms === 'number' ? `${ms} ms` : ms });
    renderSystemCheckRows(rows);
  };

  try {
    pushRow('หน้าเว็บ', 'PASS', 'JavaScript ทำงานปกติ');

    const ping = await timedFetchJson(`${CONFIG.apiBaseUrl}?action=ping&_ts=${Date.now()}`);
    if (!ping.ok || !ping.json || !ping.json.ok) {
      pushRow('API Ping', 'FAIL', ping.text || `HTTP ${ping.status}`, ping.elapsedMs);
    } else {
      pushRow('API Ping', 'PASS', ping.json.message || 'API ใช้งานได้', ping.elapsedMs);
    }

    const ym = await timedFetchJson(`${CONFIG.apiBaseUrl}?action=years-months&_ts=${Date.now()}`);
    if (!ym.ok || !ym.json || !ym.json.ok) {
      pushRow('years-months', 'FAIL', ym.text || `HTTP ${ym.status}`, ym.elapsedMs);
    } else {
      pushRow('years-months', 'PASS', `พบ ${Array.isArray(ym.json.data) ? ym.json.data.length : 0} งวด`, ym.elapsedMs);
    }

    const dash = await timedFetchJson(`${CONFIG.apiBaseUrl}?action=dashboard-data&year=${appState.selectedYear}&month=${appState.selectedMonth}&_ts=${Date.now()}`);
    if (!dash.ok || !dash.json || !dash.json.ok) {
      pushRow('dashboard-data', 'FAIL', dash.text || `HTTP ${dash.status}`, dash.elapsedMs);
      if (el.diagOverallStatus) el.diagOverallStatus.textContent = 'พบปัญหา';
    } else {
      const count = dash.json.data && Array.isArray(dash.json.data.rows) ? dash.json.data.rows.length : 0;
      pushRow('dashboard-data', 'PASS', `โหลด rows ได้ ${count} รายการ`, dash.elapsedMs);
      if (el.diagOverallStatus) el.diagOverallStatus.textContent = 'ผ่านทั้งหมด';
    }
  } catch (error) {
    pushRow('System Check', 'FAIL', error.message || String(error));
    if (el.diagOverallStatus) el.diagOverallStatus.textContent = 'พบปัญหา';
  }
}

async function initApp(isRefresh = false, keepCurrentSelection = false) {
  try {
    setStatus('สถานะระบบ: กำลังโหลดข้อมูล');

    if (isRefresh) setLastUpdated('กำลังรีเฟรช...');

    const periods = await fetchAvailablePeriods();
    appState.availablePeriods = periods.length ? periods : buildFallbackPeriods();

    if (!keepCurrentSelection) {
      appState.selectedYear = null;
      appState.selectedMonth = null;
    }

    ensureSelectedPeriod();
    renderPeriodSelectors();

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
initApp().then(() => runSystemCheck()).catch(console.error);
