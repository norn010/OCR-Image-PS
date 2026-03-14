# PDF OCR Web (Typhoon OCR)

โปรเจกต์นี้ใช้สำหรับดึงข้อมูลจากเอกสาร PDF และรูปภาพด้วย Typhoon OCR API และบันทึกลงฐานข้อมูล SQL Server โดยแบ่งการทำงานเป็น 2 ส่วนคือ **Backend (Flask)** และ **Frontend (React + Vite + Tailwind)**

## 1. การติดตั้งและรัน Backend (Flask)

Backend ทำหน้าที่สื่อสารกับ Typhoon OCR API, แปลงข้อมูล, สร้างไฟล์ Excel/Word และเชื่อมต่อกับ SQL Server

```powershell
# 1. เข้าไปที่โฟลเดอร์โปรเจกต์
cd "C:\Users\EAKSAHA\Downloads\หนอน\OCR-TTB-main\OCR-TTB-main"

# 2. สร้างและใช้งาน Virtual Environment (ถ้ายังไม่มี)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 3. ติดตั้ง Dependencies
pip install -r requirements.txt

# 4. ตั้งค่า API Key และ Database (สร้างไฟล์ .env)
# ดูตัวอย่างการตั้งค่าในหัวข้อ "การตั้งค่า .env"

# 5. รันเซิร์ฟเวอร์
python app.py
```
Backend จะรันอยู่ที่ `http://127.0.0.1:5000`

## 2. การติดตั้งและรัน Frontend (React + Vite)

Frontend เป็นหน้าต่างผู้ใช้ (UI) ที่ปรับปรุงใหม่ให้ทันสมัยและใช้งานง่ายขึ้น

```powershell
# 1. เปิด Terminal ใหม่ (คู่กับ Backend) และเข้าไปที่โฟลเดอร์ frontend
cd "C:\Users\EAKSAHA\Downloads\หนอน\OCR-TTB-main\OCR-TTB-main\frontend"

# 2. ติดตั้ง Dependencies (ทำครั้งแรกครั้งเดียว)
npm install

# 3. รันระบบ Frontend
npm run dev
```
เว็บแอปพลิเคชันจะเปิดขึ้นมาที่ `http://localhost:5173` (หรือพอร์ตอื่นตามที่ Vite กำหนด) 
**ให้เข้าใช้งานผ่าน URL ของ Frontend นี้เท่านั้น** (ตัว Frontend จะส่ง request ไปหา Backend ให้เองผ่าน proxy)

---

## การตั้งค่า `.env` (ฝั่ง Backend)

สร้างไฟล์ `.env` ที่โฟลเดอร์หลัก (ระดับเดียวกับ `app.py`) แล้วกำหนดค่าดังนี้:

```env
# ตั้งค่า Typhoon API Key
TYPHOON_API_KEY=YOUR_KEY

# ตั้งค่าฐานข้อมูล SQL Server
SQLSERVER_CONNECTION_STRING=mssql+pyodbc://@LAPTOP-V2TJ4I1J\SQLEXPRESS/ExcelTtbDB?driver=ODBC+Driver+17+for+SQL+Server&trusted_connection=yes&TrustServerCertificate=yes
```

## ฟีเจอร์หลัก
- **อัปโหลดไฟล์**: รองรับ PDF, PNG, JPG และรหัสผ่าน PDF
- **OCR Real-time**: ดึงข้อมูลและแสดงผลลัพธ์แบบเรียลไทม์ทีละหน้า
- **แก้ไขข้อมูล**: สามารถตรวจสอบและแก้ไขผล OCR ที่ไม่ถูกต้องได้โดยตรงจากหน้าเว็บ
- **Export**: กดดาวน์โหลดข้อมูลที่ดึงได้ในรูปแบบไฟล์ Excel (`.xlsx`) และ Word (`.docx`)
- **Database**: ซิงค์ข้อมูลที่ได้ลงตาราง SQL Server อัตโนมัติ
