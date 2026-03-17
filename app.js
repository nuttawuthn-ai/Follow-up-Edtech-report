const CONFIG = {
  apiBaseUrl: 'https://script.google.com/macros/s/AKfycbzHoaAp5eV9qSrbDVi9468lXyqpkD0-_VaBo_fMrB2Jkk-ni3v4AcdH7ActZTmrCw1ipg/exec',
  fallbackDemo: false
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
  search: ''
};

const el = {
  yearSelect: document.getElementById('yearSelect'),
  monthSelect: document.getElementById('monthSelect'),
  searchInput: document.getElementById('searchInput'),
  refreshBtn: document.getElementById('refreshBtn'),
  lastUpdated: document.getElementById('lastUpdated'),
  deadlineValue: document.getElementById('deadlineValue'),
  normalCount: document.getElementById('normalCount'),
  lateCount: document.getElementById('lateCount'),
  missingCount: document.getElementById('missingCount'),
  totalCount: document.getElementById('totalCount'),
  summaryLabel: document.getElementById('summaryLabel'),
  personTableBody: document.getElementById('personTableBody'),
  simpleChart: document.getElementById('simpleChart'),
  statusText: document.getElementById('statusText')
};

function formatDateTime(value) {
  return value || '-';
}

function getBadgeClass(status) {
  if (status === 'ปกติ') return 'normal';
  if (status === 'ล่าช้า') return 'late';
  return 'missing';
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
  const res = await fetch(url, { method: 'GET' });

  if (!res.ok) {
    throw new Error(`โหลดข้อมูลไม่สำเร็จ: ${res.status}`);
  }

  const json = await res.json();

  if (!json.ok) {
    throw new Error(json.message || 'API error');
  }

  return json;
}

async function fetchAvailablePeriods() {
  if (!CONFIG.apiBaseUrl || CONFIG.apiBaseUrl.includes('PASTE_YOUR')) {
    throw new Error('ยังไม่ได้ตั้งค่า Apps Script Web App URL');
  }

  const json = await fetchJson(`${CONFIG.apiBaseUrl}?action=years-months`);
  return Array.isArray(json.data) ? json.data : [];
}

async function fetchDashboardData(year, month) {
  if (!CONFIG.apiBaseUrl || CONFIG.apiBaseUrl.includes('PASTE_YOUR')) {
    throw new Error('ยังไม่ได้ตั้งค่า Apps Script Web App URL');
  }

  const url =
    `${CONFIG.apiBaseUrl}?action=dashboard-data&year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}`;

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
  return [
    {
      year: now.getFullYear(),
      month: now.getMonth() + 1,
      monthThai: monthNamesThai[now.getMonth() + 1] || ''
    }
  ];
}

function ensureSelectedPeriod() {
  if (!appState.availablePeriods.length) {
    appState.availablePeriods = buildFallbackPeriods();
  }

  const years = getUniqueYears(appState.availablePeriods);

  if (!appState.selectedYear || !years.includes(Number(appState.selectedYear))) {
    appState.selectedYear = years[0];
  }

  const months = getMonthsForYear(appState.availablePeriods, appState.selectedYear);

  if (!appState.selectedMonth || !months.includes(Number(appState.selectedMonth))) {
    appState.selectedMonth = months[months.length - 1] || months[0];
  }
}

function renderPeriodSelectors() {
  ensureSelectedPeriod();

  const years = getUniqueYears(appState.availablePeriods);
  const months = getMonthsForYear(appState.availablePeriods, appState.selectedYear);

  createOptions(el.yearSelect, years);
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
      <td>${item.name || '-'}</td>
      <td><span class="badge ${getBadgeClass(item.status)}">${item.status || '-'}</span></td>
      <td>${formatDateTime(item.submittedAt)}</td>
      <td>${item.deadline || '-'}</td>
      <td>${fileCell}</td>
      <td>${item.note || '-'}</td>
    `;

    el.personTableBody.appendChild(tr);
  });
}

function renderFilteredView() {
  if (!appState.dashboardData) return;

  const search = (appState.search || '').toLowerCase();
  const rows = Array.isArray(appState.dashboardData.rows) ? appState.dashboardData.rows : [];

  const filteredRows = rows.filter(r => {
    if (!search) return true;
    return String(r.name || '').toLowerCase().includes(search);
  });

  const normalCount = filteredRows.filter(r => r.status === 'ปกติ').length;
  const lateCount = filteredRows.filter(r => r.status === 'ล่าช้า').length;
  const missingCount = filteredRows.filter(r => r.status === 'ยังไม่ส่ง').length;
  const total = filteredRows.length;

  if (el.deadlineValue) el.deadlineValue.textContent = appState.dashboardData.deadline || '-';
  if (el.normalCount) el.normalCount.textContent = normalCount;
  if (el.lateCount) el.lateCount.textContent = lateCount;
  if (el.missingCount) el.missingCount.textContent = missingCount;
  if (el.totalCount) el.totalCount.textContent = total;
  if (el.summaryLabel) {
    el.summaryLabel.textContent =
      `${monthNamesThai[appState.selectedMonth] || appState.selectedMonth} ${appState.selectedYear}`;
  }

  renderTable(filteredRows);
  renderChart(normalCount, lateCount, missingCount, total);
}

async function loadDashboardForSelectedPeriod() {
  const data = await fetchDashboardData(appState.selectedYear, appState.selectedMonth);
  appState.dashboardData = data;
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
    el.searchInput.addEventListener('input', () => {
      appState.search = el.searchInput.value.trim();
      renderFilteredView();
    });
  }

  if (el.refreshBtn) {
    el.refreshBtn.addEventListener('click', async () => {
      await initApp(true, true);
    });
  }
}

async function initApp(isRefresh = false, keepCurrentSelection = false) {
  try {
    if (el.statusText) el.statusText.textContent = 'สถานะระบบ: กำลังโหลดข้อมูล';
    if (isRefresh && el.lastUpdated) {
      el.lastUpdated.textContent = 'กำลังรีเฟรช...';
    }

    const periods = await fetchAvailablePeriods();
    appState.availablePeriods = periods.length ? periods : buildFallbackPeriods();

    if (!keepCurrentSelection) {
      appState.selectedYear = null;
      appState.selectedMonth = null;
    }

    ensureSelectedPeriod();
    renderPeriodSelectors();

    await loadDashboardForSelectedPeriod();

    if (el.lastUpdated) {
      el.lastUpdated.textContent = `อัปเดตล่าสุด: ${new Date().toLocaleString('th-TH')}`;
    }
    if (el.statusText) {
      el.statusText.textContent = 'สถานะระบบ: พร้อมใช้งาน';
    }

    renderFilteredView();
  } catch (error) {
    console.error(error);
    if (el.statusText) {
      el.statusText.textContent = 'สถานะระบบ: โหลดข้อมูลไม่สำเร็จ';
    }
    if (el.lastUpdated) {
      el.lastUpdated.textContent = error.message || 'Unknown error';
    }
  }
}

bindEvents();
initApp();
