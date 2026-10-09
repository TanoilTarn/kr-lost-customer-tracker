# KR Lost Customer Tracker

สคริปต์ Python สำหรับหาลูกค้าเส้นทาง TH → Korea ที่หายไปหรือยอดลดลงในแต่ละเดือน (นับตามเดือนของ ETD) แล้วสร้างไฟล์ HTML ให้เซลกรอกผลการเช็ค และรวมคำตอบกลับเป็น Excel

## ติดตั้ง
```bash
pip install -r requirements.txt
```

## 1) สร้างไฟล์ส่งให้เซล
```bash
python3 kr_tracker.py build KR_May_to_Oct.xls
```
ได้ไฟล์ `dist/KR_Lost_Customer_Tracker.html` และสรุปลูกค้าที่หายไปของเดือนล่าสุดบนหน้าจอ
(ใช้ `-o ชื่อไฟล์.html` เพื่อกำหนดชื่อเอง)

## 2) เซลกรอกแล้วส่งกลับ
เปิดไฟล์ในเบราว์เซอร์ → เลือก Sales + ใส่ชื่อผู้เช็ค → เลือกผลการเช็ค/พิมพ์หมายเหตุ → กด **บันทึกไฟล์ (ส่งกลับ)** → ส่งไฟล์ที่ได้กลับมา

## 2.1) ให้เซลกด "ส่งข้อมูล" จากลิงก์ได้เลย (Google Sheet)
1. สร้าง Google Sheet ใหม่ → Extensions → Apps Script → วางโค้ดจาก `apps_script/Code.gs` → Save
2. Deploy → New deployment → Web app · Execute as: **Me** · Who has access: **Anyone** → Deploy → คัดลอก URL (ลงท้าย `/exec`)
3. สร้างไฟล์ใหม่พร้อม URL (จำไว้ใน `config.json` ครั้งต่อไปไม่ต้องใส่อีก)
```bash
python3 kr_tracker.py build KR_May_to_Oct.xls --sheet-url https://script.google.com/macros/s/XXXX/exec
```
หน้าเว็บจะมีปุ่ม **ส่งข้อมูล** แทนปุ่มบันทึกไฟล์ คำตอบเข้า Google Sheet ของเจ้าของทันที
(ชีต `Feedback` = คำตอบล่าสุดต่อราย, ชีต `Log` = ประวัติทุกครั้งที่ส่ง) และทุกคนที่เปิดลิงก์จะเห็นคำตอบล่าสุด

## 3) รวมคำตอบ
```bash
python3 kr_tracker.py merge returned/*.html
```
ได้ไฟล์ในโฟลเดอร์ `merged/`
- `KR_feedback_<เวลา>.xlsx` : ชีต Feedback, สรุปตาม Sales, สรุปเหตุผล
- `KR_Lost_Customer_Tracker_merged_<เวลา>.html` : หน้า HTML ที่มีคำตอบของทุกคนรวมกัน

ถ้าลูกค้ารายเดียวกันมีหลายคำตอบ จะใช้คำตอบที่บันทึกล่าสุด

## เกณฑ์การแบ่งกลุ่ม
- หายไปเดือนนี้: เดือนก่อน > 0 และเดือนนี้ = 0
- ลดลง ≥50%: เดือนนี้ ≤ ครึ่งหนึ่งของเดือนก่อน
- หายไปก่อนหน้า: 0 ทั้งเดือนนี้และเดือนก่อน แต่เคยจองก่อนหน้านั้น
- ลูกค้าใหม่: มียอดเดือนนี้ และไม่เคยจองมาก่อน

## โครงสร้าง
- `kr_tracker.py` : สคริปต์หลัก (build / merge)
- `src/template.html` : หน้าเว็บ (ตาราง, ช่อง feedback, ปุ่มบันทึก/รวมไฟล์)
- `dist/` : ไฟล์ HTML ที่สร้างแล้ว
