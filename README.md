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

หน้าเว็บโหลดข้อมูลชุดกลางจาก `GET /api/stock` เมื่อเปิดหน้าเว็บหรือกดปุ่ม `เริ่มใหม่`
ปุ่มนี้กลับไปหน้าแรกและดึงสต็อกล่าสุดจาก PythonAnywhere; ไม่มีการตรวจข้อมูลเป็นระยะ
ระหว่างโหลดหรืออัปโหลด ปุ่มที่ส่งคำขอจะถูกปิดชั่วคราว และเว้น 3 วินาทีก่อนกดซ้ำ
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

## Daily stock display and Notify

`POST /api/stock/quantity` (also `/api/sell-out`) accepts `path`, `quantity`,
`expected_quantity`, `version`, `day`, and a unique `request_id`. It records a
sell-out (minus) or receipt (plus) in `data/daily-stock.sqlite3`. The imported
`stock.xlsx` is never modified. Repeated requests with the same ID do not create
another sale. A stale product quantity, import version, or business day returns
HTTP 409 so the user can review the current stock before submitting again.

`GET /api/daily-stock` returns the current business day, stock version, product
display overrides and today's shared Notify events. The frontend checks every
five seconds while visible and refreshes when returning to the tab. Products
sold down to zero remain visible as zero. Notify shows each movement's actual
server timestamp. The import warning shows the filename and full Thai date/time,
including seconds; PDF imports capture the time the conversion response arrives.

At 00:00 Asia/Bangkok, the API switches to a fresh business day: display quantities
return to the imported workbook and Notify displays only the new day's events.
Older events remain in SQLite for audit. A new import starts a new set of display
overrides, while earlier movements from today remain visible in Notify.
Offline snapshots are read-only and cannot submit shared movements.

Deploy both the frontend and `backend/app.py`; reload the PythonAnywhere web app.
No new Python dependency is needed (SQLite is built in). Keep `data/stock.xlsx`,
`data/stock.metadata.json`, and `data/daily-stock.sqlite3` when deploying. On Linux,
a file lock plus SQLite transactions serialize submissions across workers.
