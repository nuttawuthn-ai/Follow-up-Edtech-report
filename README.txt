ไฟล์ชุดนี้พร้อมอัปขึ้น GitHub Pages ทันที

ให้นำไฟล์ 3 ตัวนี้ขึ้นไปไว้ที่ root ของ repository เดียวกัน
1) index.html
2) app.js
3) styles.css

สำคัญมาก
- ชื่อไฟล์ต้องตรงตามนี้เท่านั้น
- ห้ามใช้ index_debug.html หรือ app_debug.js บน GitHub จริง
- ให้ rename ไฟล์ดังนี้ก่อนอัป
  - index_debug.html -> index.html
  - app_debug.js -> app.js

สิ่งที่เพิ่มในเวอร์ชัน debug
- แสดง Error Panel บนหน้าเว็บทันที
- มี Debug Console ดู log การโหลดข้อมูล
- ตรวจ timeout ของ API
- ตรวจกรณี API ไม่ส่ง JSON
- ใส่ _ts กัน cache
- จับ JavaScript runtime error และ unhandled promise rejection
- ปุ่มคัดลอก log ได้

วิธีตรวจหลังอัปขึ้น GitHub Pages
1) เปิดหน้าเว็บ
2) กดปุ่ม “เปิด/ปิด Debug”
3) ดูข้อความใน Error Panel และ Debug Console
4) ถ้ายังไม่ขึ้น ให้ตรวจว่า Apps Script Web App เปิดสิทธิ์ Anyone แล้ว
5) ทดสอบเปิด Apps Script URL ตรง ๆ ดังนี้
   ?action=years-months
   ?action=dashboard-data&year=2026&month=3

ถ้าเปิด URL API แล้วไม่คืน JSON
แปลว่าปัญหาอยู่ฝั่ง Google Apps Script ไม่ใช่ GitHub Pages
