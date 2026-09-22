# STOCK // RADIAL

PWA สำหรับดู stock สินค้าจาก `stock.xlsx` และแปลงรายงาน PDF จาก Oracle ผ่าน PythonAnywhere

## Frontend

ไฟล์หน้าเว็บอยู่ที่ root ของ repo นี้ และเหมาะกับ GitHub Pages:

- `index.html`
- `app.js`
- `style.css`
- `manifest.json`
- `service-worker.js`
- `stock.xlsx`
- `icons/`

GitHub Pages URL:

```text
https://jaym1383.github.io/stock/
```

## Backend

ไฟล์ backend สำหรับ PythonAnywhere อยู่ใน `backend/`

Endpoint หลัก:

```text
POST https://adizjust.pythonanywhere.com/api/convert
```

รับ `multipart/form-data` field ชื่อ `file` เป็น PDF แล้วคืน `stock.xlsx` พร้อมบันทึกเป็นสต็อกชุดกลาง
การอัปโหลด `.xlsx`, `.xls` หรือ `.csv` จากหน้าเว็บจะส่งเป็น XLSX ไปที่ `POST /api/upload-stock` เช่นกัน

หน้าเว็บโหลดข้อมูลชุดกลางจาก `GET /api/stock` และตรวจ `GET /api/stock-version` ทุก 10 วินาทีขณะเปิดแท็บอยู่
แท็บที่ถูกซ่อนจะหยุดตรวจและตรวจทันทีเมื่อกลับมาใช้งาน
เมื่อมีคนอัปโหลดข้อมูลใหม่ หน้าที่เปิดอยู่จะอัปเดตอัตโนมัติภายในรอบตรวจถัดไป
ก่อนมีการอัปโหลดครั้งแรก เว็บจะใช้ `stock.xlsx` ที่มากับ GitHub Pages
ข้อมูลชุดกลางถูกเก็บใน `~/stock/data/stock.xlsx` บน PythonAnywhere

## PythonAnywhere Setup

หลัง clone repo ไปที่ PythonAnywhere แล้ว ให้ติดตั้ง dependencies:

```bash
cd ~/stock/backend
python3.13 -m pip install --user -r requirements.txt
```

ตั้งค่า WSGI file ของ PythonAnywhere ให้ชี้มาที่ `backend/pythonanywhere_wsgi.py` หรือคัดลอกเนื้อหาไฟล์นั้นไปใส่ใน WSGI config ของ web app

จากนั้น Reload web app แล้วทดสอบ:

```text
https://adizjust.pythonanywhere.com/health
```

## Verification

ทดสอบล่าสุดแล้ว:

- GitHub Pages เปิดได้ที่ `https://jaym1383.github.io/stock/`
- PythonAnywhere `/health` ตอบ `{"ok": true}`
- อัปโหลด PDF ผ่านหน้าเว็บจริงสำเร็จ และ backend คืน `stock.xlsx` กลับมาให้ UI อ่านทันที
- ทดสอบ API ข้อมูลชุดกลางด้วยสองไคลเอนต์ และตรวจว่าไฟล์เสียไม่ทับข้อมูลเดิม
