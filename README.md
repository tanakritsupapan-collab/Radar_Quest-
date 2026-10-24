# Radar Quest Game — วิธีติดตั้ง

แอปเกมเก็บแต้ม: ผู้เล่นตามหาอุปกรณ์ที่มีเลข Asset/Serial ยกมาไว้ที่ ASSET ZONE ถ่ายรูป 3 รูป แล้วได้แต้ม
ระบบทำงานบน Google Apps Script + Google Sheet

## ไฟล์ในโฟลเดอร์
| ไฟล์ | ใช้ทำอะไร |
|---|---|
| `Code.gs` | โค้ดฝั่งเซิร์ฟเวอร์ |
| `Index.html` | หน้าแอป |
| `Asset_Master_SAP.xlsx` | ข้อมูล SAP — อัปโหลดขึ้น Google Drive ก่อนรัน importSAP |
| `appsscript.json` | ตั้งค่าโปรเจกต์ (Drive API v3, web app) |
| `media/` | ป้าย ASSET ZONE (A4), GIF สาธิต, วิดีโอเรื่องจำลอง, ภาพร่างหน้าจอ |

## ติดตั้ง (ครั้งแรก)
1. เปิด https://script.google.com → New project → ตั้งชื่อ **Radar Quest Game**
2. เปิด `Code.gs` ในโฟลเดอร์นี้ด้วย Notepad → Ctrl+A, Ctrl+C → วางทับใน Code.gs ของโปรเจกต์
   สร้างไฟล์ HTML ชื่อ `Index` แล้ววางเนื้อหาจาก `Index.html` แบบเดียวกัน → Ctrl+S
   อัปโหลด `Asset_Master_SAP.xlsx` และ `media/radar-quest-how-to-earn.gif` ขึ้น Google Drive (ไว้ที่ไหนก็ได้ ชื่อไฟล์ต้องตรง)
3. Services (+) → เพิ่ม **Drive API** (v3)
4. เลือกฟังก์ชัน **setup** → Run → อนุญาตสิทธิ์ (กด Select all)
   - Execution log จะแสดง **อีเมลแอดมิน + รหัสผ่านชั่วคราว** → เข้าแอปแล้วเปลี่ยนที่ Profile › Change password
   - ระบบสร้าง Google Sheet ชื่อ **Radar Quest Game Data** และโฟลเดอร์ Drive **Radar Quest Photos**
5. เลือกฟังก์ชัน **importSAP** → Run → ข้อมูล SAP เข้าแท็บ `Ref`
6. Deploy → New deployment → Web app → Execute as **Me**, Who has access **Anyone** → Deploy → ได้ลิงก์แอป

## แท็บใน Google Sheet
- **Ref** ข้อมูลจาก SAP (Asset13, Invent no, Asset Sup, Detail, Serial) + ช่องที่แอดมินแก้: Status, Owner, Admin Note
- **Updates** ทุกครั้งที่ผู้เล่นเจอของ (ห้อง, Too big to move, รูป 3 รูป, แต้ม) + ช่องกันโกง: Pick, Photo Age, Photo Hash, Flags, Check, Checked By, Checked At
  (ถ้าเป็น Sheet เก่า ระบบเพิ่มคอลัมน์ที่ขาดให้เองครั้งแรกที่มีคนใช้แอป)
- **NewItems** ของที่ไม่มีใน SAP (ส้ม = รอตรวจ, เขียว = อนุมัติ, เทา = ปฏิเสธ) + กอง / Remark
- **Users**, **Locations**, **Points** (ทุกการได้/หักแต้ม), **History** (ใครแก้อะไร เมื่อไร)

## กติกาแต้ม
- ย้ายมาที่ ASSET ZONE: **+10** · Too big to move: **+5**
- ของใหม่ที่ไม่มีใน SAP: ได้แต้มเมื่อแอดมิน Approve
- เครื่องหนึ่งอัปเดตได้ครั้งเดียว · แอดมินสุ่มตรวจและหักแต้มได้ (บันทึกเหตุผล)

## อ่านเลขจากป้าย (ตั้งแต่ Version 6)
1. ผู้เล่นถ่ายรูปป้ายครั้งเดียว → มือถืออ่าน **บาร์โค้ด** เองด้วย ZXing (ไม่อัปโหลดรูปไปอ่าน)
   - ป้าย AEROTHAI: บาร์โค้ด Code 93 = **Asset13** 13 หลัก (ป้าย `1-130-01872-0000` → `1130018720000`) จับคู่แบบตรงทุกหลักเท่านั้น
2. ไม่มีบาร์โค้ด → ใช้ **OCR** เดิมอัตโนมัติ (รอไม่เกิน 30 วินาที)
3. ไม่สำเร็จทั้งคู่ → พิมพ์เลขเอง
- รูปป้ายนี้กลายเป็นรูปที่ 1 (Barcode) ให้อัตโนมัติ กดถ่ายใหม่ได้
- ตอนส่ง: ถ้าเซิร์ฟเวอร์เป็นคนจับคู่รูปป้ายใบนั้นกับของชิ้นนั้นเอง (บาร์โค้ดหรือ OCR) จะไม่อ่าน OCR ซ้ำ ไม่งั้นตรวจด้วย OCR เหมือนเดิม
- คอลัมน์ `Scan Timing` ในแท็บ Updates: `how=barcode|photo|typed decode= ocr= lookup= total=` (มิลลิวินาที) ไว้เทียบความเร็ว
- Pick = `barcode` / `photo` (OCR) / `typed`

## โปรเจกต์ TEST
ทำสำเนาโปรเจกต์ แล้วตั้งชื่อให้มีคำว่า **TEST** → ลบ Script Properties `SHEET_ID`, `PHOTO_FOLDER_ID`, `EXPORT_FOLDER_ID` → รัน setup / importSAP
ระบบสร้าง Sheet / โฟลเดอร์ชื่อลงท้าย " TEST" และจะหยุดทำงานทันทีถ้าโปรเจกต์ TEST ยังชี้ไปที่ข้อมูลจริง

## กันโกง (Admin › Needs check)
ผู้เล่นได้แต้มทันทีเสมอ แต่ระบบติดธงให้แอดมินตรวจเมื่อ
- **label-mismatch / label-unreadable / label-unavailable** ผู้เล่นพิมพ์เลขเอง แล้วรูปป้าย (รูปที่ 1) อ่าน OCR ได้ไม่ตรงกับของที่เลือก หรืออ่านไม่ได้
- **old-photo** รูปถ่ายไว้นานเกิน 15 นาที
- **dup:รหัส** รูปเหมือน (หรือเกือบเหมือน) รูปที่เคยใช้ส่งของชิ้นอื่นแล้ว

แอดมินกด **OK** หรือ **Wrong** — Wrong หักแต้มได้พร้อมเหตุผล (บันทึกในแท็บ Points และ History)

## เครื่องมือแอดมินอื่น
- Reset การเจอของ (ของชิ้นนั้นให้คนอื่นหาใหม่ได้ เลือกดึงแต้มคืนได้)
- แก้ชื่อ / ซ่อนห้อง, ล้างป้าย "New" ของห้องที่ผู้เล่นเพิ่ม
- หน้าอุปกรณ์: แก้ Status (Used / Spare / Dispose), Owner, Admin Note, Removed, ย้ายห้องของการเจอ, เพิ่มอุปกรณ์เข้า Ref (รวมถึงจากของใหม่ใน NewItems)
- ปรับแต้มเองพร้อมเหตุผล

## GIF "How to earn" ในแอป
แอปดึงไฟล์ `radar-quest-how-to-earn.gif` ที่ใหม่ที่สุดใน Google Drive — จะเปลี่ยนรูป ให้อัปโหลดไฟล์ใหม่ชื่อเดิม (แอปเห็นภายใน ~1 ชั่วโมง)

## หมายเหตุ
- อีเมล activate / reset รหัส ส่งจากบัญชี Google ของเจ้าของสคริปต์ (Gmail ส่งได้ ~100 ฉบับ/วัน)
- รัน `importSAP` ซ้ำได้: ข้อมูล SAP ถูกแทนที่ แต่ Status / Owner / Note ที่แอดมินแก้ยังอยู่ และข้อมูลการเจอของ (Updates) ไม่ถูกแตะ
- Export (แท็บ Admin › More) ได้ไฟล์ ZIP = Excel + โฟลเดอร์รูป เก็บใน Drive "Radar Quest Exports"

## ลิงก์สำหรับผู้เล่น
- ใช้ลิงก์ web app โดยตรง (ลิงก์เดิมทุกเวอร์ชัน):
  https://script.google.com/macros/s/AKfycbwOiwDWOVYm3XZWBDG90B09d56DbbYmeqojHsadcKKLpmh6d8IKXjAi9cUK_pNWrSQm5A/exec
- ลิงก์ในอีเมลแจ้งอนุมัติพาไปลิงก์นี้ (`ENTRY_URL` ใน Code.gs ว่างไว้)
- หน้า Google Sites (sites.google.com/view/radar-quest) เลิกใช้แล้ว — บนมือถือยังขึ้น "Page not found"
- ถ้ามือถือขึ้น "Page not found" (มีหลายบัญชี Google, ลิงก์มี /u/1/): เปิดใน Chrome (⋮ › Open in Chrome) แล้ว Add to Home screen

## เริ่มเกมจริง (ล้างข้อมูลทดสอบ)
เปิด Apps Script → เลือกฟังก์ชัน **resetGame** → Run
- สำรอง Google Sheet ทั้งไฟล์ไว้ก่อน (ชื่อ "… backup วันเวลา")
- เก็บ: ข้อมูล SAP, บัญชีที่อนุมัติแล้วทั้งหมด (แต้มกลับเป็น 0), 2 ห้องเริ่มต้น
- ล้าง: การเจอของ, ของใหม่, ประวัติแต้ม, History, ห้องที่เพิ่มตอนทดสอบ, Status/Owner/Note ที่แอดมินแก้, ของที่แอดมินเพิ่มเอง, บัญชีที่ยังไม่อนุมัติ
- รูปทดสอบและไฟล์ Export ย้ายไปถังขยะของ Drive (กู้คืนได้ภายใน 30 วัน)
- setup / importSAP / resetGame รันได้จากหน้า Apps Script โดยเจ้าของเท่านั้น

## อัปเดตโค้ดครั้งต่อไป
วางโค้ดใหม่ → Ctrl+S → Deploy › Manage deployments › ✏️ › Version: New version › Deploy (ลิงก์เดิม)

เวอร์ชันที่ deploy ล่าสุด: **Version 6** (24 ต.ค. 2026 12:18) — บาร์โค้ด ZXing + OCR สำรอง (schema 4)
ย้อนกลับ: Deploy › Manage deployments › ✏️ › Version: **Version 5** (18 ต.ค. 21:58) › Deploy (ลิงก์เดิม)
