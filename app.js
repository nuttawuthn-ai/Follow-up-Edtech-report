<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Follow-up EdTech Report Dashboard</title>
  <link rel="stylesheet" href="styles.css" />
</head>
<body>
  <div class="container">
    <header class="page-header">
      <h1>Dashboard ติดตามการส่งรายงาน</h1>
      <p id="statusText">สถานะระบบ: กำลังโหลดข้อมูล</p>
      <p id="lastUpdated">-</p>
    </header>

    <section class="toolbar">
      <div class="field">
        <label for="yearSelect">ปี</label>
        <select id="yearSelect"></select>
      </div>

      <div class="field">
        <label for="monthSelect">เดือน</label>
        <select id="monthSelect"></select>
      </div>

      <div class="field">
        <label for="searchInput">ค้นหาชื่อ</label>
        <input id="searchInput" type="text" placeholder="พิมพ์ชื่อบุคลากร" />
      </div>

      <div class="field button-wrap">
        <button id="refreshBtn" type="button">รีเฟรชข้อมูล</button>
      </div>
    </section>

    <section class="summary-header">
      <h2 id="summaryLabel">-</h2>
      <p>กำหนดส่ง: <span id="deadlineValue">-</span></p>
    </section>

    <section class="stats-grid">
      <div class="card">
        <h3>ปกติ</h3>
        <div id="normalCount">0</div>
      </div>

      <div class="card">
        <h3>ล่าช้า</h3>
        <div id="lateCount">0</div>
      </div>

      <div class="card">
        <h3>ยังไม่ส่ง</h3>
        <div id="missingCount">0</div>
      </div>

      <div class="card">
        <h3>รวมทั้งหมด</h3>
        <div id="totalCount">0</div>
      </div>
    </section>

    <section class="chart-section">
      <h3>กราฟสรุปสถานะ</h3>
      <div id="simpleChart"></div>
    </section>

    <section class="table-section">
      <h3>รายละเอียดรายบุคคล</h3>
      <table>
        <thead>
          <tr>
            <th>ชื่อ</th>
            <th>สถานะ</th>
            <th>วันที่ส่งจริง</th>
            <th>วันกำหนดส่ง</th>
            <th>ไฟล์รายงาน</th>
            <th>หมายเหตุ</th>
          </tr>
        </thead>
        <tbody id="personTableBody"></tbody>
      </table>
    </section>
  </div>

  <script src="app.js"></script>
</body>
</html>
