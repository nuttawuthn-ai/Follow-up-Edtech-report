const CONFIG = {
  apiBaseUrl: 'https://script.google.com/a/macros/ku.th/s/AKfycbzHoaAp5eV9qSrbDVi9468lXyqpkD0-_VaBo_fMrB2Jkk-ni3v4AcdH7ActZTmrCw1ipg/exec'
};

const monthNamesThai = {
  1: 'มกราคม', 2: 'กุมภาพันธ์', 3: 'มีนาคม', 4: 'เมษายน', 5: 'พฤษภาคม', 6: 'มิถุนายน',
  7: 'กรกฎาคม', 8: 'สิงหาคม', 9: 'กันยายน', 10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม'
};

let appState = {
  data: null,
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
  if (!value) return '-';
  return value;
}

function getBadgeClass(status) {
  if (status === 'ปกติ') return 'normal';
  if (status === 'ล่าช้า') return 'late';
  return 'missing';
}

function createOptions(select, values, formatter = v => v) {
  select.innerHTML = '';
  values.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = formatter(value);
    select.appendChild(option);
  });
}

async function fetchDashboardData() {
  if (!CONFIG.apiBaseUrl || CONFIG.apiBaseUrl.includes('PASTE_YOUR')) {
    if (CONFIG.fallbackDemo) return getDemoData();
    throw new Error('ยังไม่ได้ตั้งค่า Apps Script Web App URL');
  }

  const res = await fetch(`${CONFIG.apiBaseUrl}?action=dashboard-data`, { method: 'GET' });
  if (!res.ok) throw new Error(`โหลดข้อมูลไม่สำเร็จ: ${res.status}`);
  return await res.json();
}

function getDemoData() {
  return {
    ok: true,
    lastUpdated: new Date().toLocaleString('th-TH'),
    availableYears: [2026],
    availableMonths: [1,2,3,4,5,6,7,8,9,10,11,12],
    currentYear: 2026,
    currentMonth: 3,
    records: [
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'สิงห์ทอง ครองพงษ์', deadline: '14/03/2026', submittedAt: '12/03/2026 10:30:00', status: 'ปกติ', fileUrl: '#', note: '' },
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'ธนวิตร์ พัฒนะ', deadline: '14/03/2026', submittedAt: '16/03/2026 09:15:00', status: 'ล่าช้า', fileUrl: '#', note: 'ส่งหลังครบกำหนด' },
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'พงศธร ช้างโรจน์', deadline: '14/03/2026', submittedAt: '', status: 'ยังไม่ส่ง', fileUrl: '', note: '' },
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'ภควดี ปุริโต', deadline: '14/03/2026', submittedAt: '11/03/2026 14:00:00', status: 'ปกติ', fileUrl: '#', note: '' },
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'วลฌา อาศัยผล', deadline: '14/03/2026', submittedAt: '13/03/2026 16:45:00', status: 'ปกติ', fileUrl: '#', note: '' },
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'หทัยรัตน์ ศรีสุภะ', deadline: '14/03/2026', submittedAt: '', status: 'ยังไม่ส่ง', fileUrl: '', note: '' },
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'สุภัทรา นวชินกุล', deadline: '14/03/2026', submittedAt: '14/03/2026 11:20:00', status: 'ปกติ', fileUrl: '#', note: '' },
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'ไพลิน จิตเจริญสมุทร', deadline: '14/03/2026', submittedAt: '17/03/2026 08:05:00', status: 'ล่าช้า', fileUrl: '#', note: '' },
      { year: 2026, month: 3, monthThai: 'มีนาคม', name: 'ณัฐวุฒิ นันทปรีชา', deadline: '14/03/2026', submittedAt: '10/03/2026 13:00:00', status: 'ปกติ', fileUrl: '#', note: '' }
    ]
  };
}

function renderChart(normalCount, lateCount, missingCount, total) {
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
      <div class="bar-track"><div class="bar-fill ${item.cls}" style="width:${pct}%"></div></div>
      <div>${item.value}</div>
    `;
    el.simpleChart.appendChild(row);
  });
}

function renderTable(records) {
  el.personTableBody.innerHTML = '';
  records.forEach(item => {
    const tr = document.createElement('tr');
    const fileCell = item.fileUrl
      ? `<a class="file-link" href="${item.fileUrl}" target="_blank" rel="noopener">เปิดไฟล์</a>`
      : '-';
    tr.innerHTML = `
      <td>${item.name}</td>
      <td><span class="badge ${getBadgeClass(item.status)}">${item.status}</span></td>
      <td>${formatDateTime(item.submittedAt)}</td>
      <td>${item.deadline || '-'}</td>
      <td>${fileCell}</td>
      <td>${item.note || '-'}</td>
    `;
    el.personTableBody.appendChild(tr);
  });
}

function renderFilteredView() {
  if (!appState.data) return;

  const { selectedYear, selectedMonth, search } = appState;
  const records = appState.data.records
    .filter(r => Number(r.year) === Number(selectedYear) && Number(r.month) === Number(selectedMonth))
    .filter(r => !search || r.name.toLowerCase().includes(search.toLowerCase()));

  const normalCount = records.filter(r => r.status === 'ปกติ').length;
  const lateCount = records.filter(r => r.status === 'ล่าช้า').length;
  const missingCount = records.filter(r => r.status === 'ยังไม่ส่ง').length;
  const total = records.length;
  const first = records[0];

  el.deadlineValue.textContent = first?.deadline || '-';
  el.normalCount.textContent = normalCount;
  el.lateCount.textContent = lateCount;
  el.missingCount.textContent = missingCount;
  el.totalCount.textContent = total;
  el.summaryLabel.textContent = `${monthNamesThai[selectedMonth] || selectedMonth} ${selectedYear}`;

  renderTable(records);
  renderChart(normalCount, lateCount, missingCount, total);
}

function bindEvents() {
  el.yearSelect.addEventListener('change', () => {
    appState.selectedYear = Number(el.yearSelect.value);
    renderFilteredView();
  });

  el.monthSelect.addEventListener('change', () => {
    appState.selectedMonth = Number(el.monthSelect.value);
    renderFilteredView();
  });

  el.searchInput.addEventListener('input', () => {
    appState.search = el.searchInput.value.trim();
    renderFilteredView();
  });

  el.refreshBtn.addEventListener('click', async () => {
    await initApp(true);
  });
}

async function initApp(isRefresh = false) {
  try {
    el.statusText.textContent = 'สถานะระบบ: กำลังโหลดข้อมูล';
    if (isRefresh) el.lastUpdated.textContent = 'กำลังรีเฟรช...';

    const data = await fetchDashboardData();
    appState.data = data;
    appState.selectedYear = appState.selectedYear || data.currentYear || data.availableYears?.[0];
    appState.selectedMonth = appState.selectedMonth || data.currentMonth || data.availableMonths?.[0];

    createOptions(el.yearSelect, data.availableYears || []);
    createOptions(el.monthSelect, data.availableMonths || [], m => `${m} - ${monthNamesThai[m] || ''}`);

    el.yearSelect.value = appState.selectedYear;
    el.monthSelect.value = appState.selectedMonth;
    el.lastUpdated.textContent = `อัปเดตล่าสุด: ${data.lastUpdated || '-'}`;
    el.statusText.textContent = 'สถานะระบบ: พร้อมใช้งาน';

    renderFilteredView();
  } catch (error) {
    console.error(error);
    el.statusText.textContent = 'สถานะระบบ: โหลดข้อมูลไม่สำเร็จ';
    el.lastUpdated.textContent = error.message;
  }
}

bindEvents();
initApp();
