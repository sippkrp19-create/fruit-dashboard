# Fruit Farm Dashboard

เว็บแดชบอร์ดสวนผลไม้สำหรับดูภาพรวมรายรับ รายจ่าย สุทธิ เงินค้าง และข้อมูลจาก Google Sheets

## Google Sheets Source

เว็บดึงข้อมูลจาก Google Sheets:

https://docs.google.com/spreadsheets/d/1YEsURhIz4oyM9sVebBrXfHId9TRn1BKOVEQVIqvVBHA/edit?usp=sharing

ชีตข้อมูลหลัก: `ฐานข้อมูล`

ชีตรายการ dropdown: `รายการดรอปดาวน์`

## Web Dashboard

โฟลเดอร์เว็บ:

```text
web-dashboard/
```

ไฟล์หลัก:

- `web-dashboard/index.html`
- `web-dashboard/styles.css`
- `web-dashboard/app.js`

## วิธีเปิดใช้งานบนเครื่อง

```powershell
cd "C:\Users\sippk\OneDrive\เอกสาร\ทำเว็ปเพจ\แดชบอร์ด\web-dashboard"
C:\Users\sippk\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe -m http.server 8765 --bind 127.0.0.1
```

แล้วเปิด:

```text
http://127.0.0.1:8765/
```

## ฟีเจอร์หลัก

- KPI รายรับ รายจ่าย สุทธิ ปริมาณกก. ราคาเฉลี่ย และเงินค้าง
- ตัวกรองปี เดือน สินค้า สายพันธุ์/เกรด วิธีรับ/จ่ายเงิน และสถานะ
- กราฟวงกลมเลือกมุมมองได้: รายรับ / รายจ่าย / สุทธิ
- กราฟวงกลมรองรับกรณีมีข้อมูลชนิดเดียว
- ชนิดสินค้าหลักแสดงเสมอ: ทุเรียน, ลองกอง, มังคุด
- ตัวเลือกเกรดหลักแสดงเสมอ และตรวจข้อมูลจริงจาก Google Sheets เมื่อเลือก
- รวมค่า `โอน` เป็น `เงินโอน`
- กราฟรายเดือน คลิกเดือนเพื่อกรองทั้งหน้า
- กล่องเงินสด เงินโอน ค้างรับ ค้างจ่าย
- รายการต้องติดตาม
- ตารางรายการล่าสุด
- พื้นหลังภาพผลไม้จริงพร้อมโทนเขียวทอง
- รองรับมือถือ

## เงื่อนไขสำคัญ

Google Sheets ต้องแชร์แบบ `Anyone with the link can view` เพื่อให้เว็บดึงข้อมูลจาก browser ได้

