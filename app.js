(function () {
  'use strict';

  const CONFIG = {
    apiBaseUrl: 'https://script.google.com/macros/s/AKfycbzHoaAp5eV9qSrbDVi9468lXyqpkD0-_VaBo_fMrB2Jkk-ni3v4AcdH7ActZTmrCw1ipg/exec',
    fallbackDemo: false,
    autoRefreshMs: 120000,
    debug: true,
    requestTimeoutMs: 25000,
    version: 'debug-github-pages-2026-03-17-01'
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

  const appState = {
    dashboardData: null,
    availablePeriods: [],
    selectedYear: null,
    selectedMonth: null,
    search: '',
    autoRefreshTimer: null,
    debugLogs: []
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
    statusText: document.getElementById('statusText'),
    rowCountLabel: document.getElementById('rowCountLabel'),
    errorPanel: document.getElementById('errorPanel'),
    errorText: document.getElementById('errorText'),
    debugPanel: document.getElementById('debugPanel'),
    debugLog: document.getElementById('debugLog'),
    debugPageUrl: document.getElementById('debugPageUrl'),
    debugApiUrl: document.getElementById('debugApiUrl'),
    debugLastStep: document.getElementById('debugLastStep'),
    toggleDebugBtn: document.getElementById('toggleDebugBtn'),
    copyDebugBtn: document.getElementById('copyDebugBtn'),
    clearDebugBtn: document.getElementById('clearDebugBtn')
  };

  function nowStamp() {
    return new Date().toLocaleString('th-TH');
  }

  function safeJson(value) {
    try {
      return JSON.stringify(value, null, 2);
    } catch (_) {
      return String(value);
    }
  }

  function logDebug(step, detail) {
    const line = `[${nowStamp()}] ${step}${detail ? `\n${detail}` : ''}`;
    appState.debugLogs.push(line);
    if (appState.debugLogs.length > 200) appState.debugLogs.shift();

    if (CONFIG.debug) {
      console.log(step, detail || '');
    }

    if (el.debugLastStep) el.debugLastStep.textContent = step;
    if (el.debugLog) el.debugLog.textContent = appState.debugLogs.join('\n\n');
  }

  function showError(message, detail) {
    const combined = detail ? `${message}\n\n${detail}` : message;
    if (el.errorPanel) el.errorPanel.classList.remove('hidden');
    if (el.errorText) el.errorText.textContent = combined;
    logDebug('ERROR', combined);
  }

  function clearError() {
    if (el.errorPanel) el.errorPanel.classList.add('hidden');
    if (el.errorText) el.errorText.textContent = '-';
  }

  function setStatus(message) {
    if (el.statusText) el.statusText.textContent = message;
    logDebug('STATUS', message);
  }

  function setLastUpdated(message) {
    if (el.lastUpdated) el.lastUpdated.textContent = message;
  }

  function formatDateTime(value) {
    return value || '-';
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function getBadgeClass(status) {
    if (status === 'ปกติ') return 'normal';
    if (status === 'ล่าช้า') return 'late';
    return 'missing';
  }

  function createOptions(select, values, formatter = (v) => v) {
    if (!select) return;

    select.innerHTML = '';

    values.forEach((value) => {
      const option = document.createElement('option');
      option.value = String(value);
      option.textContent = formatter(value);
      select.appendChild(option);
    });
  }

  async function fetchJson(url) {
    logDebug('FETCH_START', url);

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), CONFIG.requestTimeoutMs);

    try {
      const res = await fetch(url, {
        method: 'GET',
        cache: 'no-store',
        mode: 'cors',
        signal: controller.signal,
        headers: {
          Accept: 'application/json, text/plain, */*'
        }
      });

      logDebug('FETCH_RESPONSE', `status=${res.status} ok=${res.ok} type=${res.headers.get('content-type') || '-'}`);

      if (!res.ok) {
        throw new Error(`โหลดข้อมูลไม่สำเร็จ: HTTP ${res.status}`);
      }

      const rawText = await res.text();
      logDebug('FETCH_RAW', rawText.slice(0, 1200));

      let json;
      try {
        json = JSON.parse(rawText);
      } catch (parseError) {
        throw new Error(`API ไม่ได้ส่ง JSON ที่ถูกต้อง\n${parseError.message}\n\nตัวอย่างข้อมูลที่ได้รับ:\n${rawText.slice(0, 500)}`);
      }

      if (!json || typeof json !== 'object') {
        throw new Error('API ส่งข้อมูลกลับมาไม่ใช่ object');
      }

      if (!json.ok) {
        throw new Error(json.message || 'API ตอบกลับ ok=false');
      }

      logDebug('FETCH_JSON_OK', safeJson(json).slice(0, 1200));
      return json;
    } catch (error) {
      if (error.name === 'AbortError') {
        throw new Error(`คำขอใช้เวลานานเกิน ${CONFIG.requestTimeoutMs / 1000} วินาที`);
      }
      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  async function fetchAvailablePeriods() {
    if (!CONFIG.apiBaseUrl || CONFIG.apiBaseUrl.includes('PASTE_YOUR')) {
      throw new Error('ยังไม่ได้ตั้งค่า Apps Script Web App URL');
    }

    const url = `${CONFIG.apiBaseUrl}?action=years-months&_ts=${Date.now()}`;
    const json = await fetchJson(url);
    return Array.isArray(json.data) ? json.data : [];
  }

  async function fetchDashboardData(year, month) {
    if (!CONFIG.apiBaseUrl || CONFIG.apiBaseUrl.includes('PASTE_YOUR')) {
      throw new Error('ยังไม่ได้ตั้งค่า Apps Script Web App URL');
    }

    const url = `${CONFIG.apiBaseUrl}?action=dashboard-data&year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}&_ts=${Date.now()}`;
    const json = await fetchJson(url);
    return json.data;
  }

  function getUniqueYears(periods) {
    return [...new Set(periods.map((item) => Number(item.year)).filter(Boolean))].sort((a, b) => b - a);
  }

  function getMonthsForYear(periods, year) {
    return periods
      .filter((item) => Number(item.year) === Number(year))
      .map((item) => Number(item.month))
      .filter(Boolean)
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

    createOptions(el.yearSelect, years);
    createOptions(el.monthSelect, months, (m) => `${m} - ${monthNamesThai[m] || ''}`);

    if (el.yearSelect) el.yearSelect.value = String(appState.selectedYear);
    if (el.monthSelect) el.monthSelect.value = String(appState.selectedMonth);

    logDebug('RENDER_SELECTORS', `year=${appState.selectedYear} month=${appState.selectedMonth}`);
  }

  function renderChart(normalCount, lateCount, missingCount, total) {
    if (!el.simpleChart) return;

    const bars = [
      { label: 'ปกติ', value: normalCount, cls: 'success' },
      { label: 'ล่าช้า', value: lateCount, cls: 'danger' },
      { label: 'ยังไม่ส่ง', value: missingCount, cls: 'warning' }
    ];

    el.simpleChart.innerHTML = '';

    bars.forEach((item) => {
      const pct = total ? Math.round((item.value / total) * 100) : 0;
      const row = document.createElement('div');
      row.className = 'bar-row';
      row.innerHTML = `
        <div>${escapeHtml(item.label)}</div>
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

    if (!records.length) {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="6" class="empty-row">ไม่พบข้อมูล</td>';
      el.personTableBody.appendChild(tr);
      return;
    }

    records.forEach((item) => {
      const tr = document.createElement('tr');
      const safeUrl = item.fileUrl ? String(item.fileUrl) : '';
      const fileCell = safeUrl
        ? `<a class="file-link" href="${escapeHtml(safeUrl)}" target="_blank" rel="noopener">เปิดไฟล์</a>`
        : '-';

      tr.innerHTML = `
        <td>${escapeHtml(item.name || '-')}</td>
        <td><span class="badge ${getBadgeClass(item.status)}">${escapeHtml(item.status || '-')}</span></td>
        <td>${escapeHtml(formatDateTime(item.submittedAt))}</td>
        <td>${escapeHtml(item.deadline || '-')}</td>
        <td>${fileCell}</td>
        <td>${escapeHtml(item.note || '-')}</td>
      `;

      el.personTableBody.appendChild(tr);
    });
  }

  function renderFilteredView() {
    if (!appState.dashboardData) {
      logDebug('RENDER_SKIP', 'dashboardData is empty');
      return;
    }

    const search = (appState.search || '').toLowerCase();
    const rows = Array.isArray(appState.dashboardData.rows) ? appState.dashboardData.rows : [];

    const filteredRows = rows.filter((r) => {
      if (!search) return true;
      return String(r.name || '').toLowerCase().includes(search);
    });

    const normalCount = filteredRows.filter((r) => r.status === 'ปกติ').length;
    const lateCount = filteredRows.filter((r) => r.status === 'ล่าช้า').length;
    const missingCount = filteredRows.filter((r) => r.status === 'ยังไม่ส่ง').length;
    const total = filteredRows.length;

    if (el.deadlineValue) el.deadlineValue.textContent = appState.dashboardData.deadline || '-';
    if (el.normalCount) el.normalCount.textContent = normalCount;
    if (el.lateCount) el.lateCount.textContent = lateCount;
    if (el.missingCount) el.missingCount.textContent = missingCount;
    if (el.totalCount) el.totalCount.textContent = total;
    if (el.rowCountLabel) el.rowCountLabel.textContent = `${total} รายการ`;

    if (el.summaryLabel) {
      el.summaryLabel.textContent = `${monthNamesThai[appState.selectedMonth] || appState.selectedMonth} ${appState.selectedYear}`;
    }

    renderTable(filteredRows);
    renderChart(normalCount, lateCount, missingCount, total);
    logDebug('RENDER_DONE', `rows=${rows.length} filtered=${filteredRows.length}`);
  }

  async function loadDashboardForSelectedPeriod() {
    const data = await fetchDashboardData(appState.selectedYear, appState.selectedMonth);
    if (!data || typeof data !== 'object') {
      throw new Error('dashboard-data ไม่มีข้อมูลที่ถูกต้อง');
    }
    appState.dashboardData = data;
    logDebug('DASHBOARD_DATA_OK', safeJson(data).slice(0, 1200));
  }

  function startAutoRefresh() {
    if (appState.autoRefreshTimer) {
      clearInterval(appState.autoRefreshTimer);
    }

    if (CONFIG.autoRefreshMs > 0) {
      appState.autoRefreshTimer = setInterval(async () => {
        try {
          logDebug('AUTO_REFRESH', `every ${CONFIG.autoRefreshMs}ms`);
          await initApp(true, true);
        } catch (err) {
          console.error('Auto refresh error:', err);
          showError('Auto refresh error', err.message || String(err));
        }
      }, CONFIG.autoRefreshMs);
    }
  }

  function bindEvents() {
    if (el.yearSelect) {
      el.yearSelect.addEventListener('change', async () => {
        appState.selectedYear = Number(el.yearSelect.value);
        logDebug('EVENT_YEAR_CHANGE', `year=${appState.selectedYear}`);
        renderPeriodSelectors();
        await initApp(false, true);
      });
    }

    if (el.monthSelect) {
      el.monthSelect.addEventListener('change', async () => {
        appState.selectedMonth = Number(el.monthSelect.value);
        logDebug('EVENT_MONTH_CHANGE', `month=${appState.selectedMonth}`);
        await initApp(false, true);
      });
    }

    if (el.searchInput) {
      el.searchInput.addEventListener('input', () => {
        appState.search = el.searchInput.value.trim();
        logDebug('EVENT_SEARCH', appState.search || '(empty)');
        renderFilteredView();
      });
    }

    if (el.refreshBtn) {
      el.refreshBtn.addEventListener('click', async () => {
        logDebug('EVENT_REFRESH_CLICK', 'manual refresh');
        await initApp(true, true);
      });
    }

    if (el.toggleDebugBtn && el.debugPanel) {
      el.toggleDebugBtn.addEventListener('click', () => {
        el.debugPanel.classList.toggle('hidden');
      });
    }

    if (el.clearDebugBtn) {
      el.clearDebugBtn.addEventListener('click', () => {
        appState.debugLogs = [];
        logDebug('DEBUG_LOG_CLEARED', '-');
      });
    }

    if (el.copyDebugBtn) {
      el.copyDebugBtn.addEventListener('click', async () => {
        const text = appState.debugLogs.join('\n\n');
        try {
          await navigator.clipboard.writeText(text);
          logDebug('DEBUG_LOG_COPIED', 'copied to clipboard');
        } catch (error) {
          showError('ไม่สามารถคัดลอก log ได้', error.message || String(error));
        }
      });
    }

    window.addEventListener('error', (event) => {
      showError('JavaScript runtime error', `${event.message}\n${event.filename || ''}:${event.lineno || ''}:${event.colno || ''}`);
    });

    window.addEventListener('unhandledrejection', (event) => {
      const reason = event.reason instanceof Error ? `${event.reason.message}\n${event.reason.stack || ''}` : safeJson(event.reason);
      showError('Unhandled Promise rejection', reason);
    });
  }

  async function initApp(isRefresh = false, keepCurrentSelection = false) {
    try {
      clearError();
      setStatus('สถานะระบบ: กำลังโหลดข้อมูล');
      if (isRefresh) setLastUpdated('กำลังรีเฟรช...');

      const periods = await fetchAvailablePeriods();
      appState.availablePeriods = periods.length ? periods : buildFallbackPeriods();
      logDebug('PERIODS_OK', safeJson(appState.availablePeriods));

      if (!keepCurrentSelection) {
        appState.selectedYear = null;
        appState.selectedMonth = null;
      }

      ensureSelectedPeriod();
      renderPeriodSelectors();
      await loadDashboardForSelectedPeriod();

      setLastUpdated(`อัปเดตล่าสุด: ${nowStamp()}`);
      setStatus('สถานะระบบ: พร้อมใช้งาน');
      renderFilteredView();
    } catch (error) {
      console.error(error);
      setStatus('สถานะระบบ: โหลดข้อมูลไม่สำเร็จ');
      setLastUpdated(error.message || 'Unknown error');
      showError('โหลดข้อมูลไม่สำเร็จ', `${error.message || String(error)}\n\nโปรดตรวจสอบว่า Apps Script ถูก deploy เป็น Web App และอนุญาตให้เข้าถึงแบบ Anyone.`);
    }
  }

  function bootstrapDebugMeta() {
    if (el.debugPageUrl) el.debugPageUrl.textContent = window.location.href;
    if (el.debugApiUrl) el.debugApiUrl.textContent = CONFIG.apiBaseUrl;
    if (el.debugLog) {
      el.debugLog.textContent = [
        `Version: ${CONFIG.version}`,
        `User agent: ${navigator.userAgent}`,
        `เวลาเริ่มต้น: ${nowStamp()}`,
        `กำลังเริ่มระบบ...`
      ].join('\n');
    }
    logDebug('BOOTSTRAP', `version=${CONFIG.version}`);
  }

  bindEvents();
  bootstrapDebugMeta();
  startAutoRefresh();
  initApp();
})();
