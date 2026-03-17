const CONFIG = {
  apiBaseUrl: 'https://script.google.com/macros/s/AKfycbz_WGKpm3WHqeHq61tNnDjyhpjYyxwAMf7tui3x6zfDv47i6BADWNDqRjMUVZAKVbJnmQ/exec',
  fallbackDemo: false,
  autoRefreshMs: 0
};

const appState = {
  periods: [],
  selectedYear: '',
  selectedMonth: '',
  dashboardData: null
};

const el = {
  yearSelect: document.getElementById('yearSelect'),
  monthSelect: document.getElementById('monthSelect'),
  refreshBtn: document.getElementById('refreshBtn'),
  runSystemCheckBtn: document.getElementById('runSystemCheckBtn'),
  systemStatus: document.getElementById('systemStatus'),
  summaryTitle: document.getElementById('summaryTitle'),
  totalCount: document.getElementById('totalCount'),
  normalCount: document.getElementById('normalCount'),
  lateCount: document.getElementById('lateCount'),
  missingCount: document.getElementById('missingCount'),
  tableContainer: document.getElementById('tableContainer'),
  asOfDate: document.getElementById('asOfDate'),
  deadlineValue: document.getElementById('deadlineValue'),
  diagPageUrl: document.getElementById('diagPageUrl'),
  diagApiUrl: document.getElementById('diagApiUrl'),
  diagOverallStatus: document.getElementById('diagOverallStatus'),
  systemCheckBody: document.getElementById('systemCheckBody')
};

function setStatus(text) {
  if (el.systemStatus) {
    el.systemStatus.textContent = text;
  }
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function thaiNowString() {
  return new Date().toLocaleString('th-TH', {
    dateStyle: 'long',
    timeStyle: 'medium',
    timeZone: 'Asia/Bangkok'
  });
}

function formatThaiDate(dateValue) {
  if (!dateValue) return '-';
  const d = new Date(dateValue);
  if (Number.isNaN(d.getTime())) return '-';
  return d.toLocaleDateString('th-TH', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    timeZone: 'Asia/Bangkok'
  });
}

function formatThaiDateTime(dateValue) {
  if (!dateValue) return '-';
  const d = new Date(dateValue);
  if (Number.isNaN(d.getTime())) return '-';
  return d.toLocaleString('th-TH', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    timeZone: 'Asia/Bangkok'
  });
}

function parseDateSafe(value) {
  if (!value) return null;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}

function toGregorianYear(yearValue) {
  const y = Number(yearValue);
  if (!Number.isFinite(y)) return y;
  return y > 2400 ? y - 543 : y;
}

function getSecondFriday(year, month) {
  const firstDay = new Date(year, month - 1, 1);
  const firstDayDow = firstDay.getDay();
  const friday = 5;
  const offsetToFirstFriday = (friday - firstDayDow + 7) % 7;
  const firstFridayDate = 1 + offsetToFirstFriday;
  const secondFridayDate = firstFridayDate + 7;
  return new Date(year, month - 1, secondFridayDate, 23, 59, 59, 999);
}

function getNextMonth(year, month) {
  if (month === 12) {
    return { year: year + 1, month: 1 };
  }
  return { year, month: month + 1 };
}

function getReportDeadline(reportYear, reportMonth) {
  const y = toGregorianYear(reportYear);
  const m = Number(reportMonth);
  const next = getNextMonth(y, m);
  return getSecondFriday(next.year, next.month);
}

function monthNameThai(month) {
  const names = [
    '', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  return names[Number(month)] || String(month);
}

function displayThaiYear(year) {
  const y = Number(year);
  if (!Number.isFinite(y)) return String(year || '');
  return y > 2400 ? String(y) : String(y + 543);
}

function deriveStatus(row, selectedYear, selectedMonth) {
  const submitted = parseDateSafe(row.submittedAt || row.submitted_at || row.submittedDate || row.lastSubmittedAt);
  const deadline = getReportDeadline(selectedYear, selectedMonth);

  if (!submitted) {
    return {
      status: 'ยังไม่ส่ง',
      deadlineText: formatThaiDate(deadline),
      note: row.note || 'ยังไม่พบข้อมูลการส่ง'
    };
  }

  if (submitted.getTime() <= deadline.getTime()) {
    return {
      status: 'ปกติ',
      deadlineText: formatThaiDate(deadline),
      note: row.note || 'ส่งภายในกำหนด'
    };
  }

  return {
    status: 'ล่าช้า',
    deadlineText: formatThaiDate(deadline),
    note: row.note || 'ส่งเกินกำหนด'
  };
}

function getStatusBadgeClass(status) {
  if (status === 'ปกติ') return 'badge badge-normal';
  if (status === 'ล่าช้า') return 'badge badge-late';
  if (status === 'ยังไม่ส่ง') return 'badge badge-missing';
  return 'badge badge-other';
}

async function fetchJson(url, timeoutMs = 20000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, {
      method: 'GET',
      cache: 'no-store',
      signal: controller.signal
    });

    const text = await response.text();
    let json = null;

    try {
      json = JSON.parse(text);
    } catch (err) {
      throw new Error(`API ตอบกลับไม่ใช่ JSON: ${text.slice(0, 200)}`);
    }

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    return json;
  } finally {
    clearTimeout(timer);
  }
}

async function fetchAvailablePeriods() {
  const url = `${CONFIG.apiBaseUrl}?action=years-months&_ts=${Date.now()}`;
  return fetchJson(url);
}

async function fetchDashboardData(year, month) {
  const url = `${CONFIG.apiBaseUrl}?action=dashboard-data&year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}&_ts=${Date.now()}`;
  return fetchJson(url);
}

function fillPeriodSelectors(periods) {
  appState.periods = Array.isArray(periods) ? periods : [];

  el.yearSelect.innerHTML = '';
  el.monthSelect.innerHTML = '';

  const uniqueYears = [...new Set(appState.periods.map(p => String(p.year)))];

  if (uniqueYears.length === 0) {
    el.yearSelect.innerHTML = '<option value="">ไม่พบข้อมูลปี</option>';
    el.monthSelect.innerHTML = '<option value="">ไม่พบข้อมูลเดือน</option>';
    return;
  }

  uniqueYears.forEach(year => {
    const option = document.createElement('option');
    option.value = year;
    option.textContent = displayThaiYear(year);
    el.yearSelect.appendChild(option);
  });

  if (!appState.selectedYear) {
    appState.selectedYear = uniqueYears[0];
  }

  el.yearSelect.value = appState.selectedYear;
  fillMonthSelectorByYear(appState.selectedYear);
}

function fillMonthSelectorByYear(year) {
  const months = appState.periods
    .filter(p => String(p.year) === String(year))
    .map(p => Number(p.month))
    .sort((a, b) => a - b);

  el.monthSelect.innerHTML = '';

  months.forEach(month => {
    const option = document.createElement('option');
    option.value = String(month);
    option.textContent = monthNameThai(month);
    el.monthSelect.appendChild(option);
  });

  if (!months.length) {
    el.monthSelect.innerHTML = '<option value="">ไม่พบข้อมูลเดือน</option>';
    appState.selectedMonth = '';
    return;
  }

  if (!months.includes(Number(appState.selectedMonth))) {
    appState.selectedMonth = String(months[0]);
  }

  el.monthSelect.value = appState.selectedMonth;
}

function renderSummary(rows, payload) {
  const total = rows.length;
  const normal = rows.filter(r => r.status === 'ปกติ').length;
  const late = rows.filter(r => r.status === 'ล่าช้า').length;
  const missing = rows.filter(r => r.status === 'ยังไม่ส่ง').length;

  el.totalCount.textContent = String(total);
  el.normalCount.textContent = String(normal);
  el.lateCount.textContent = String(late);
  el.missingCount.textContent = String(missing);

  const selectedYearDisplay = displayThaiYear(appState.selectedYear);
  const selectedMonthDisplay = monthNameThai(appState.selectedMonth);
  el.summaryTitle.textContent = `งวดรายงาน ${selectedMonthDisplay} ${selectedYearDisplay}`;

  const deadline = payload?.deadline || (rows[0] && rows[0].deadline) || formatThaiDate(getReportDeadline(appState.selectedYear, appState.selectedMonth));
  el.deadlineValue.textContent = deadline || '-';
  el.asOfDate.textContent = thaiNowString();
}

function renderTable(rows) {
  if (!rows.length) {
    el.tableContainer.innerHTML = '<div class="empty-box">ไม่พบข้อมูลในงวดที่เลือก</div>';
    return;
  }

  const html = `
    <table>
      <thead>
        <tr>
          <th>ลำดับ</th>
          <th>ชื่อ</th>
          <th>วันที่ส่ง</th>
          <th>วันครบกำหนด</th>
          <th>สถานะ</th>
          <th>หมายเหตุ</th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((row, index) => `
          <tr>
            <td>${index + 1}</td>
            <td>${escapeHtml(row.name || row.staffName || '-')}</td>
            <td>${escapeHtml(row.submittedAtDisplay || row.submittedAt || '-')}</td>
            <td>${escapeHtml(row.deadline || '-')}</td>
            <td><span class="${getStatusBadgeClass(row.status)}">${escapeHtml(row.status || '-')}</span></td>
            <td>${escapeHtml(row.note || '-')}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
  el.tableContainer.innerHTML = html;
}

function normalizeRowsFromPayload(payload) {
  const rawRows = Array.isArray(payload?.rows)
    ? payload.rows
    : Array.isArray(payload?.data?.rows)
      ? payload.data.rows
      : [];

  return rawRows.map(row => {
    const derived = deriveStatus(row, appState.selectedYear, appState.selectedMonth);
    const submittedSource = row.submittedAt || row.submitted_at || row.submittedDate || row.lastSubmittedAt || '';
    return {
      ...row,
      name: row.name || row.staffName || row.staff || row.fullName || '-',
      submittedAt: submittedSource || '',
      submittedAtDisplay: submittedSource ? formatThaiDateTime(submittedSource) : '-',
      deadline: row.deadline || derived.deadlineText,
      status: row.status || derived.status,
      note: row.note || derived.note
    };
  });
}

async function loadDashboardData() {
  if (!appState.selectedYear || !appState.selectedMonth) {
    setStatus('สถานะระบบ: ยังไม่ได้เลือกปี/เดือน');
    return;
  }

  try {
    setStatus('สถานะระบบ: กำลังโหลดข้อมูล...');
    const payload = await fetchDashboardData(appState.selectedYear, appState.selectedMonth);

    if (payload.ok === false) {
      throw new Error(payload.message || payload.error || 'โหลดข้อมูลไม่สำเร็จ');
    }

    appState.dashboardData = payload.data || payload;
    const rows = normalizeRowsFromPayload(appState.dashboardData);

    renderSummary(rows, appState.dashboardData);
    renderTable(rows);

    setStatus('สถานะระบบ: พร้อมใช้งาน และคำนวณสถานะล่าสุดแล้ว');
  } catch (error) {
    console.error(error);
    setStatus(`สถานะระบบ: โหลดข้อมูลไม่สำเร็จ - ${error.message}`);
    el.tableContainer.innerHTML = `<div class="empty-box">โหลดข้อมูลไม่สำเร็จ<br>${escapeHtml(error.message)}</div>`;
  }
}

function bindEvents() {
  el.yearSelect.addEventListener('change', async (event) => {
    appState.selectedYear = event.target.value;
    fillMonthSelectorByYear(appState.selectedYear);
    await loadDashboardData();
  });

  el.monthSelect.addEventListener('change', async (event) => {
    appState.selectedMonth = event.target.value;
    await loadDashboardData();
  });

  el.refreshBtn.addEventListener('click', async () => {
    await loadDashboardData();
  });

  el.runSystemCheckBtn.addEventListener('click', runSystemCheck);
}

function diagStatusClass(status) {
  if (status === 'PASS') return 'diag-ok';
  if (status === 'WARN') return 'diag-warn';
  if (status === 'FAIL') return 'diag-err';
  return 'diag-muted';
}

function renderSystemCheckRows(rows) {
  el.systemCheckBody.innerHTML = rows.map(row => `
    <tr>
      <td>${escapeHtml(row.name)}</td>
      <td class="${diagStatusClass(row.status)}">${escapeHtml(row.status)}</td>
      <td>${escapeHtml(row.detail)}</td>
      <td>${escapeHtml(row.ms)}</td>
    </tr>
  `).join('');
}

function updateOverallDiagStatus(rows) {
  const hasFail = rows.some(r => r.status === 'FAIL');
  const hasWarn = rows.some(r => r.status === 'WARN');

  if (hasFail) {
    el.diagOverallStatus.textContent = 'พบปัญหา';
    el.diagOverallStatus.className = 'diag-err';
  } else if (hasWarn) {
    el.diagOverallStatus.textContent = 'ผ่านแบบมีคำเตือน';
    el.diagOverallStatus.className = 'diag-warn';
  } else {
    el.diagOverallStatus.textContent = 'ผ่านทั้งหมด';
    el.diagOverallStatus.className = 'diag-ok';
  }
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
  const rows = [];
  const pushRow = (name, status, detail, ms = '-') => {
    rows.push({
      name,
      status,
      detail,
      ms: typeof ms === 'number' ? `${ms} ms` : ms
    });
    renderSystemCheckRows(rows);
    updateOverallDiagStatus(rows);
  };

  renderSystemCheckRows([{ name: 'เริ่มต้น', status: 'INFO', detail: 'กำลังตรวจระบบ...', ms: '-' }]);
  el.diagOverallStatus.textContent = 'กำลังตรวจ...';
  el.diagOverallStatus.className = 'diag-muted';

  try {
    pushRow('หน้าเว็บ', 'PASS', 'JavaScript ทำงานปกติ');

    if (!CONFIG.apiBaseUrl) {
      pushRow('API Base URL', 'FAIL', 'ไม่พบ CONFIG.apiBaseUrl');
      return;
    }
    pushRow('API Base URL', 'PASS', CONFIG.apiBaseUrl);

    const pingUrl = `${CONFIG.apiBaseUrl}?action=ping&_ts=${Date.now()}`;
    try {
      const ping = await timedFetchJson(pingUrl);
      if (!ping.ok) {
        pushRow('API Ping', 'FAIL', `HTTP ${ping.status}`, ping.elapsedMs);
      } else if (!ping.json) {
        pushRow('API Ping', 'FAIL', 'ตอบกลับไม่ใช่ JSON', ping.elapsedMs);
      } else if (ping.json.ok === true) {
        pushRow('API Ping', 'PASS', ping.json.message || 'API ใช้งานได้', ping.elapsedMs);
      } else {
        pushRow('API Ping', 'WARN', JSON.stringify(ping.json), ping.elapsedMs);
      }
    } catch (error) {
      pushRow('API Ping', 'FAIL', error.message || String(error));
    }

    const ymUrl = `${CONFIG.apiBaseUrl}?action=years-months&_ts=${Date.now()}`;
    let ymData = null;

    try {
      const ym = await timedFetchJson(ymUrl);
      if (!ym.ok) {
        pushRow('years-months', 'FAIL', `HTTP ${ym.status}`, ym.elapsedMs);
      } else if (!ym.json) {
        pushRow('years-months', 'FAIL', 'ตอบกลับไม่ใช่ JSON', ym.elapsedMs);
      } else if (ym.json.ok === true && Array.isArray(ym.json.data)) {
        ymData = ym.json.data;
        pushRow('years-months', 'PASS', `พบข้อมูล ${ym.json.data.length} รายการ`, ym.elapsedMs);
      } else {
        pushRow('years-months', 'WARN', JSON.stringify(ym.json), ym.elapsedMs);
      }
    } catch (error) {
      pushRow('years-months', 'FAIL', error.message || String(error));
    }

    let year = appState.selectedYear;
    let month = appState.selectedMonth;

    if ((!year || !month) && Array.isArray(ymData) && ymData.length > 0) {
      year = String(ymData[0].year || '');
      month = String(ymData[0].month || '');
      pushRow('ปี/เดือนที่ใช้ตรวจ', 'WARN', `ใช้ค่าแรกจาก API: ${year}/${month}`);
    } else if (year && month) {
      pushRow('ปี/เดือนที่ใช้ตรวจ', 'PASS', `${year}/${month}`);
    } else {
      pushRow('ปี/เดือนที่ใช้ตรวจ', 'FAIL', 'ไม่พบปี/เดือนสำหรับทดสอบ');
      return;
    }

    const dashUrl = `${CONFIG.apiBaseUrl}?action=dashboard-data&year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}&_ts=${Date.now()}`;

    try {
      const dash = await timedFetchJson(dashUrl);
      if (!dash.ok) {
        pushRow('dashboard-data', 'FAIL', `HTTP ${dash.status}`, dash.elapsedMs);
      } else if (!dash.json) {
        pushRow('dashboard-data', 'FAIL', 'ตอบกลับไม่ใช่ JSON', dash.elapsedMs);
      } else if (dash.json.ok === true && dash.json.data) {
        const keys = Object.keys(dash.json.data || {});
        pushRow('dashboard-data', 'PASS', `โหลดสำเร็จ: ${keys.join(', ') || '(ไม่มี key)'}`, dash.elapsedMs);
      } else {
        pushRow('dashboard-data', 'WARN', JSON.stringify(dash.json), dash.elapsedMs);
      }
    } catch (error) {
      pushRow('dashboard-data', 'FAIL', error.message || String(error));
    }
  } catch (error) {
    pushRow('System Check', 'FAIL', error.message || String(error));
  }
}

async function initApp() {
  try {
    el.diagPageUrl.textContent = window.location.href;
    el.diagApiUrl.textContent = CONFIG.apiBaseUrl || '-';
    el.asOfDate.textContent = thaiNowString();

    bindEvents();

    setStatus('สถานะระบบ: กำลังโหลดรายการปี/เดือน...');
    const periodsPayload = await fetchAvailablePeriods();

    if (periodsPayload.ok === false) {
      throw new Error(periodsPayload.message || periodsPayload.error || 'โหลดปี/เดือนไม่สำเร็จ');
    }

    const periods = Array.isArray(periodsPayload.data) ? periodsPayload.data : [];
    fillPeriodSelectors(periods);

    await loadDashboardData();
    await runSystemCheck();
  } catch (error) {
    console.error(error);
    setStatus(`สถานะระบบ: โหลดข้อมูลไม่สำเร็จ - ${error.message}`);
    el.tableContainer.innerHTML = `<div class="empty-box">เริ่มต้นระบบไม่สำเร็จ<br>${escapeHtml(error.message)}</div>`;
  }
}

window.addEventListener('error', (event) => {
  console.error('Global Error:', event.error || event.message);
});

window.addEventListener('unhandledrejection', (event) => {
  console.error('Unhandled Promise Rejection:', event.reason);
});

document.addEventListener('DOMContentLoaded', initApp);
