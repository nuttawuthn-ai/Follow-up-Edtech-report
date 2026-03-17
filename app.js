(function () {
  'use strict';

 const CONFIG = {
  apiBaseUrl: 'https://script.google.com/macros/s/AKfycbz_WGKpm3WHqeHq61tNnDjyhpjYyxwAMf7tui3x6zfDv47i6BADWNDqRjMUVZAKVbJnmQ/exec',
  fallbackDemo: false,
  autoRefreshMs: 60000
};

  const ACTION_CANDIDATES = {
    periods: ['years-months', 'yearsMonths', 'getYearsMonths', 'get-years-months'],
    dashboard: ['dashboard-data', 'dashboardData', 'getDashboardData', 'get-dashboard-data']
  };

  const monthNamesThai = {
    1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน', 5: 'พฤษภาคม', 6: 'มิถุนายน',
    7: 'กรกฎาคม', 8: 'สิงหาคม', 9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
  };

  const appState = { dashboardData: null, availablePeriods: [], selectedYear: null, selectedMonth: null, search: '', autoRefreshTimer: null, debugLogs: [] };
  const el = {
    yearSelect: document.getElementById('yearSelect'), monthSelect: document.getElementById('monthSelect'), searchInput: document.getElementById('searchInput'), refreshBtn: document.getElementById('refreshBtn'),
    lastUpdated: document.getElementById('lastUpdated'), deadlineValue: document.getElementById('deadlineValue'), normalCount: document.getElementById('normalCount'), lateCount: document.getElementById('lateCount'),
    missingCount: document.getElementById('missingCount'), totalCount: document.getElementById('totalCount'), summaryLabel: document.getElementById('summaryLabel'), personTableBody: document.getElementById('personTableBody'),
    simpleChart: document.getElementById('simpleChart'), statusText: document.getElementById('statusText'), rowCountLabel: document.getElementById('rowCountLabel'), errorPanel: document.getElementById('errorPanel'),
    errorText: document.getElementById('errorText'), debugPanel: document.getElementById('debugPanel'), debugLog: document.getElementById('debugLog'), debugPageUrl: document.getElementById('debugPageUrl'),
    debugApiUrl: document.getElementById('debugApiUrl'), debugLastStep: document.getElementById('debugLastStep'), toggleDebugBtn: document.getElementById('toggleDebugBtn'), copyDebugBtn: document.getElementById('copyDebugBtn'), clearDebugBtn: document.getElementById('clearDebugBtn')
  };

  function nowStamp() { return new Date().toLocaleString('th-TH'); }
  function safeJson(v) { try { return JSON.stringify(v, null, 2); } catch { return String(v); } }
  function logDebug(step, detail) {
    const line = `[${nowStamp()}] ${step}${detail ? `\n${detail}` : ''}`;
    appState.debugLogs.push(line); if (appState.debugLogs.length > 300) appState.debugLogs.shift();
    if (el.debugLastStep) el.debugLastStep.textContent = step;
    if (el.debugLog) el.debugLog.textContent = appState.debugLogs.join('\n\n');
    console.log(step, detail || '');
  }
  function showError(message, detail) {
    if (el.errorPanel) el.errorPanel.classList.remove('hidden');
    if (el.errorText) el.errorText.textContent = detail ? `${message}\n\n${detail}` : message;
    logDebug('ERROR', detail ? `${message}\n${detail}` : message);
  }
  function clearError() { if (el.errorPanel) el.errorPanel.classList.add('hidden'); if (el.errorText) el.errorText.textContent = '-'; }
  function setStatus(message) { if (el.statusText) el.statusText.textContent = message; logDebug('STATUS', message); }
  function setLastUpdated(message) { if (el.lastUpdated) el.lastUpdated.textContent = message; }
  function formatDateTime(value) { return value || '-'; }
  function escapeHtml(value) { return String(value ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;'); }
  function getBadgeClass(status) { if (status === 'ปกติ') return 'normal'; if (status === 'ล่าช้า') return 'late'; return 'missing'; }
  function createOptions(select, values, formatter = v => v) { if (!select) return; select.innerHTML=''; values.forEach(value=>{ const option=document.createElement('option'); option.value=String(value); option.textContent=formatter(value); select.appendChild(option);}); }

  async function fetchText(url) {
    logDebug('FETCH_START', url);
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), CONFIG.requestTimeoutMs);
    try {
      const res = await fetch(url, { method: 'GET', cache: 'no-store', mode: 'cors', signal: controller.signal, headers: { Accept: 'application/json, text/plain, */*' } });
      const rawText = await res.text();
      logDebug('FETCH_RESPONSE', `status=${res.status} ok=${res.ok} type=${res.headers.get('content-type') || '-'}`);
      logDebug('FETCH_RAW', rawText.slice(0, 700));
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return rawText;
    } catch (error) {
      if (error.name === 'AbortError') throw new Error(`คำขอใช้เวลานานเกิน ${CONFIG.requestTimeoutMs / 1000} วินาที`);
      throw error;
    } finally { clearTimeout(timeoutId); }
  }

  async function fetchJsonByActions(actionList, extraParams = {}) {
    let lastError = null;
    for (const action of actionList) {
      const params = new URLSearchParams({ action, _ts: Date.now().toString(), ...Object.fromEntries(Object.entries(extraParams).map(([k, v]) => [k, String(v)])) });
      const url = `${CONFIG.apiBaseUrl}?${params.toString()}`;
      try {
        const rawText = await fetchText(url);
        let json;
        try { json = JSON.parse(rawText); } catch (parseError) { throw new Error(`API ไม่ได้ส่ง JSON ที่ถูกต้องสำหรับ action=${action}: ${parseError.message}`); }
        if (json && json.ok) {
          logDebug('ACTION_MATCHED', action);
          return json;
        }
        const message = json && json.message ? String(json.message) : 'Unknown API error';
        if (/Unknown action/i.test(message)) {
          logDebug('ACTION_REJECTED', `${action} -> ${message}`);
          lastError = new Error(`action=${action}: ${message}`);
          continue;
        }
        throw new Error(message);
      } catch (error) {
        if (/Unknown action/i.test(error.message || '')) {
          logDebug('ACTION_REJECTED', `${action} -> ${error.message}`);
          lastError = error;
          continue;
        }
        throw error;
      }
    }
    throw lastError || new Error('ไม่พบ action ที่ Apps Script รองรับ');
  }

  async function fetchAvailablePeriods() {
    const json = await fetchJsonByActions(ACTION_CANDIDATES.periods);
    return Array.isArray(json.data) ? json.data : [];
  }
  async function fetchDashboardData(year, month) {
    const json = await fetchJsonByActions(ACTION_CANDIDATES.dashboard, { year, month });
    return json.data;
  }

  function getUniqueYears(periods) { return [...new Set(periods.map(item => Number(item.year)).filter(Boolean))].sort((a, b) => b - a); }
  function getMonthsForYear(periods, year) { return periods.filter(item => Number(item.year) === Number(year)).map(item => Number(item.month)).filter(Boolean).sort((a,b)=>a-b); }
  function buildFallbackPeriods() { const now = new Date(); return [{ year: now.getFullYear(), month: now.getMonth()+1, monthThai: monthNamesThai[now.getMonth()+1] || '' }]; }
  function ensureSelectedPeriod() {
    if (!appState.availablePeriods.length) appState.availablePeriods = buildFallbackPeriods();
    const years = getUniqueYears(appState.availablePeriods);
    if (!years.length) { const now = new Date(); appState.selectedYear = now.getFullYear(); appState.selectedMonth = now.getMonth()+1; return; }
    if (!appState.selectedYear || !years.includes(Number(appState.selectedYear))) appState.selectedYear = years[0];
    const months = getMonthsForYear(appState.availablePeriods, appState.selectedYear);
    if (!months.length) { appState.selectedMonth = 1; return; }
    if (!appState.selectedMonth || !months.includes(Number(appState.selectedMonth))) appState.selectedMonth = months[months.length - 1] || months[0];
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
    const bars = [{ label: 'ปกติ', value: normalCount, cls: 'success' }, { label: 'ล่าช้า', value: lateCount, cls: 'danger' }, { label: 'ยังไม่ส่ง', value: missingCount, cls: 'warning' }];
    el.simpleChart.innerHTML='';
    bars.forEach(item => { const pct = total ? Math.round((item.value/total)*100) : 0; const row = document.createElement('div'); row.className = 'bar-row'; row.innerHTML = `<div>${escapeHtml(item.label)}</div><div class="bar-track"><div class="bar-fill ${item.cls}" style="width:${pct}%"></div></div><div>${item.value}</div>`; el.simpleChart.appendChild(row); });
  }
  function renderTable(records) {
    if (!el.personTableBody) return;
    el.personTableBody.innerHTML='';
    if (!records.length) { const tr=document.createElement('tr'); tr.innerHTML='<td colspan="6" class="empty-row">ไม่พบข้อมูล</td>'; el.personTableBody.appendChild(tr); return; }
    records.forEach(item => { const tr=document.createElement('tr'); const safeUrl=item.fileUrl ? String(item.fileUrl) : ''; const fileCell = safeUrl ? `<a class="file-link" href="${escapeHtml(safeUrl)}" target="_blank" rel="noopener">เปิดไฟล์</a>` : '-'; tr.innerHTML = `<td>${escapeHtml(item.name || '-')}</td><td><span class="badge ${getBadgeClass(item.status)}">${escapeHtml(item.status || '-')}</span></td><td>${escapeHtml(formatDateTime(item.submittedAt))}</td><td>${escapeHtml(item.deadline || '-')}</td><td>${fileCell}</td><td>${escapeHtml(item.note || '-')}</td>`; el.personTableBody.appendChild(tr); });
  }
  function renderFilteredView() {
    if (!appState.dashboardData) return;
    const search=(appState.search || '').toLowerCase();
    const rows=Array.isArray(appState.dashboardData.rows) ? appState.dashboardData.rows : [];
    const filteredRows=rows.filter(r => !search || String(r.name || '').toLowerCase().includes(search));
    const normalCount=filteredRows.filter(r => r.status === 'ปกติ').length;
    const lateCount=filteredRows.filter(r => r.status === 'ล่าช้า').length;
    const missingCount=filteredRows.filter(r => r.status === 'ยังไม่ส่ง').length;
    const total=filteredRows.length;
    if (el.deadlineValue) el.deadlineValue.textContent = appState.dashboardData.deadline || '-';
    if (el.normalCount) el.normalCount.textContent = normalCount;
    if (el.lateCount) el.lateCount.textContent = lateCount;
    if (el.missingCount) el.missingCount.textContent = missingCount;
    if (el.totalCount) el.totalCount.textContent = total;
    if (el.rowCountLabel) el.rowCountLabel.textContent = `${total} รายการ`;
    if (el.summaryLabel) el.summaryLabel.textContent = `${monthNamesThai[appState.selectedMonth] || appState.selectedMonth} ${appState.selectedYear}`;
    renderTable(filteredRows); renderChart(normalCount, lateCount, missingCount, total);
  }
  function setDiagMeta() {
  const pageEl = document.getElementById('diagPageUrl');
  const apiEl = document.getElementById('diagApiUrl');
  const statusEl = document.getElementById('diagOverallStatus');

  if (pageEl) pageEl.textContent = window.location.href;
  if (apiEl) apiEl.textContent = CONFIG.apiBaseUrl || '-';
  if (statusEl) statusEl.textContent = 'พร้อมตรวจ';
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function diagStatusClass(status) {
  if (status === 'PASS') return 'diag-ok';
  if (status === 'WARN') return 'diag-warn';
  if (status === 'FAIL') return 'diag-err';
  return 'diag-muted';
}

function renderSystemCheckRows(rows) {
  const tbody = document.getElementById('systemCheckBody');
  if (!tbody) return;

  tbody.innerHTML = rows.map(row => `
    <tr>
      <td>${escapeHtml(row.name)}</td>
      <td class="${diagStatusClass(row.status)}">${escapeHtml(row.status)}</td>
      <td>${escapeHtml(row.detail)}</td>
      <td>${escapeHtml(row.ms)}</td>
    </tr>
  `).join('');
}

function updateOverallDiagStatus(rows) {
  const el = document.getElementById('diagOverallStatus');
  if (!el) return;

  const hasFail = rows.some(r => r.status === 'FAIL');
  const hasWarn = rows.some(r => r.status === 'WARN');

  if (hasFail) {
    el.textContent = 'พบปัญหา';
    el.className = 'diag-err';
  } else if (hasWarn) {
    el.textContent = 'ผ่านแบบมีคำเตือน';
    el.className = 'diag-warn';
  } else {
    el.textContent = 'ผ่านทั้งหมด';
    el.className = 'diag-ok';
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

function getSelectedYearMonthForDiag() {
  const yearEl = document.getElementById('yearSelect');
  const monthEl = document.getElementById('monthSelect');

  return {
    year: yearEl?.value || '',
    month: monthEl?.value || ''
  };
}

async function runSystemCheck() {
  const rows = [];
  const baseUrl = CONFIG.apiBaseUrl;

  const pushRow = (name, status, detail, ms = '-') => {
    rows.push({ name, status, detail, ms: typeof ms === 'number' ? `${ms} ms` : ms });
    renderSystemCheckRows(rows);
    updateOverallDiagStatus(rows);
  };

  renderSystemCheckRows([
    { name: 'เริ่มต้น', status: 'INFO', detail: 'กำลังตรวจระบบ...', ms: '-' }
  ]);

  const statusEl = document.getElementById('diagOverallStatus');
  if (statusEl) {
    statusEl.textContent = 'กำลังตรวจ...';
    statusEl.className = 'diag-muted';
  }

  try {
    pushRow('หน้าเว็บ', 'PASS', 'JavaScript ทำงานปกติ');

    if (!baseUrl) {
      pushRow('API Base URL', 'FAIL', 'ไม่พบ CONFIG.apiBaseUrl');
      return;
    } else {
      pushRow('API Base URL', 'PASS', baseUrl);
    }

    const pingUrl = `${baseUrl}?action=ping&_ts=${Date.now()}`;
    try {
      const r = await timedFetchJson(pingUrl);
      if (!r.ok) {
        pushRow('API Ping', 'FAIL', `HTTP ${r.status}`, r.elapsedMs);
      } else if (!r.json) {
        pushRow('API Ping', 'FAIL', 'ตอบกลับไม่ใช่ JSON', r.elapsedMs);
      } else if (r.json.ok === true) {
        pushRow('API Ping', 'PASS', r.json.message || 'API ใช้งานได้', r.elapsedMs);
      } else {
        pushRow('API Ping', 'WARN', JSON.stringify(r.json), r.elapsedMs);
      }
    } catch (err) {
      pushRow('API Ping', 'FAIL', err.message || String(err));
    }

    const ymUrl = `${baseUrl}?action=years-months&_ts=${Date.now()}`;
    let ymData = null;

    try {
      const r = await timedFetchJson(ymUrl);
      if (!r.ok) {
        pushRow('years-months', 'FAIL', `HTTP ${r.status}`, r.elapsedMs);
      } else if (!r.json) {
        pushRow('years-months', 'FAIL', 'ตอบกลับไม่ใช่ JSON', r.elapsedMs);
      } else if (r.json.ok === true && Array.isArray(r.json.data)) {
        ymData = r.json.data;
        pushRow('years-months', 'PASS', `พบช่วงข้อมูล ${r.json.data.length} รายการ`, r.elapsedMs);
      } else {
        pushRow('years-months', 'WARN', JSON.stringify(r.json), r.elapsedMs);
      }
    } catch (err) {
      pushRow('years-months', 'FAIL', err.message || String(err));
    }

    const selected = getSelectedYearMonthForDiag();
    let year = selected.year;
    let month = selected.month;

    if ((!year || !month) && Array.isArray(ymData) && ymData.length > 0) {
      const first = ymData[0];
      year = first.year || '';
      month = first.month || '';
      pushRow('ปี/เดือนที่ใช้ตรวจ', 'WARN', `ไม่มีค่าที่เลือกอยู่ จึงใช้ค่าแรกจาก API: ${year}/${month}`);
    } else if (year && month) {
      pushRow('ปี/เดือนที่ใช้ตรวจ', 'PASS', `${year}/${month}`);
    } else {
      pushRow('ปี/เดือนที่ใช้ตรวจ', 'FAIL', 'ไม่พบปี/เดือนสำหรับทดสอบ');
      return;
    }

    const dashUrl = `${baseUrl}?action=dashboard-data&year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}&_ts=${Date.now()}`;

    try {
      const r = await timedFetchJson(dashUrl);
      if (!r.ok) {
        pushRow('dashboard-data', 'FAIL', `HTTP ${r.status}`, r.elapsedMs);
      } else if (!r.json) {
        pushRow('dashboard-data', 'FAIL', 'ตอบกลับไม่ใช่ JSON', r.elapsedMs);
      } else if (r.json.ok === true && r.json.data) {
        const keys = Object.keys(r.json.data || {});
        pushRow('dashboard-data', 'PASS', `โหลดสำเร็จ โครงข้อมูล: ${keys.join(', ') || '(ไม่มี key)'}`, r.elapsedMs);
      } else {
        pushRow('dashboard-data', 'WARN', JSON.stringify(r.json), r.elapsedMs);
      }
    } catch (err) {
      pushRow('dashboard-data', 'FAIL', err.message || String(err));
    }

  } catch (err) {
    pushRow('System Check', 'FAIL', err.message || String(err));
  }
}

function bindSystemCheckTools() {
  setDiagMeta();

  const runBtn = document.getElementById('runSystemCheckBtn');
  const toggleBtn = document.getElementById('toggleSystemCheckBtn');
  const panel = document.getElementById('systemCheckPanel');

  if (runBtn) {
    runBtn.addEventListener('click', runSystemCheck);
  }

  if (toggleBtn && panel) {
    toggleBtn.addEventListener('click', () => {
      panel.style.display = panel.style.display === 'none' ? '' : 'none';
    });
  }
}
  async function loadDashboardForSelectedPeriod() { appState.dashboardData = await fetchDashboardData(appState.selectedYear, appState.selectedMonth); }
  function startAutoRefresh() { if (appState.autoRefreshTimer) clearInterval(appState.autoRefreshTimer); if (CONFIG.autoRefreshMs > 0) appState.autoRefreshTimer = setInterval(() => initApp(true, true).catch(err => showError('Auto refresh error', err.message || String(err))), CONFIG.autoRefreshMs); }
  function bindEvents() {
    if (el.yearSelect) el.yearSelect.addEventListener('change', async () => { appState.selectedYear = Number(el.yearSelect.value); renderPeriodSelectors(); await initApp(false, true); });
    if (el.monthSelect) el.monthSelect.addEventListener('change', async () => { appState.selectedMonth = Number(el.monthSelect.value); await initApp(false, true); });
    if (el.searchInput) el.searchInput.addEventListener('input', () => { appState.search = el.searchInput.value.trim(); renderFilteredView(); });
    if (el.refreshBtn) el.refreshBtn.addEventListener('click', async () => { await initApp(true, true); });
    if (el.toggleDebugBtn && el.debugPanel) el.toggleDebugBtn.addEventListener('click', () => el.debugPanel.classList.toggle('hidden'));
    if (el.clearDebugBtn) el.clearDebugBtn.addEventListener('click', () => { appState.debugLogs = []; logDebug('DEBUG_LOG_CLEARED', '-'); });
    if (el.copyDebugBtn) el.copyDebugBtn.addEventListener('click', async () => { try { await navigator.clipboard.writeText(appState.debugLogs.join('\n\n')); logDebug('DEBUG_LOG_COPIED', 'copied to clipboard'); } catch (error) { showError('ไม่สามารถคัดลอก log ได้', error.message || String(error)); } });
    window.addEventListener('error', event => showError('JavaScript runtime error', `${event.message}\n${event.filename || ''}:${event.lineno || ''}:${event.colno || ''}`));
    window.addEventListener('unhandledrejection', event => showError('Unhandled Promise rejection', event.reason instanceof Error ? `${event.reason.message}\n${event.reason.stack || ''}` : safeJson(event.reason)));
  }
  async function initApp(isRefresh = false, keepCurrentSelection = false) {
    try {
      clearError(); setStatus('สถานะระบบ: กำลังโหลดข้อมูล'); if (isRefresh) setLastUpdated('กำลังรีเฟรช...');
      const periods = await fetchAvailablePeriods();
      appState.availablePeriods = periods.length ? periods : buildFallbackPeriods();
      if (!keepCurrentSelection) { appState.selectedYear = null; appState.selectedMonth = null; }
      ensureSelectedPeriod(); renderPeriodSelectors(); await loadDashboardForSelectedPeriod();
      setLastUpdated(`อัปเดตล่าสุด: ${nowStamp()}`); setStatus('สถานะระบบ: พร้อมใช้งาน'); renderFilteredView();
    } catch (error) {
      setStatus('สถานะระบบ: โหลดข้อมูลไม่สำเร็จ'); setLastUpdated(error.message || 'Unknown error');
      showError('โหลดข้อมูลไม่สำเร็จ', `${error.message || String(error)}\n\nถ้ายังขึ้น Unknown action ทุก action แปลว่าต้องแก้ Apps Script ฝั่ง backend หรือ redeploy เวอร์ชันล่าสุด`);
    }
  }
  function bootstrapDebugMeta() {
    if (el.debugPageUrl) el.debugPageUrl.textContent = window.location.href;
    if (el.debugApiUrl) el.debugApiUrl.textContent = CONFIG.apiBaseUrl;
    if (el.debugLog) el.debugLog.textContent = [`Version: ${CONFIG.version}`, `Actions(periods): ${ACTION_CANDIDATES.periods.join(', ')}`, `Actions(dashboard): ${ACTION_CANDIDATES.dashboard.join(', ')}`, `เวลาเริ่มต้น: ${nowStamp()}`].join('\n');
    logDebug('BOOTSTRAP', `version=${CONFIG.version}`);
  }
  bindEvents(); bootstrapDebugMeta(); startAutoRefresh(); initApp();
})();
