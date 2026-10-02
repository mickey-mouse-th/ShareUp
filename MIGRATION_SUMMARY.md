# ShareUp: สรุปการย้ายไปเว็บแอปใหม่ (สิ่งที่ทำแล้ว / ยังไม่ได้ทำ)

อัปเดตล่าสุด: 2026-10-01 · ครอบคลุม repo `ShareUp` (Apps Script เดิม) และ `ShareUp-web` (เว็บใหม่)

---

## 1. ภาพรวม

**เป้าหมาย:** ย้าย ShareUp จาก Google Apps Script + Google Sheets ไปเป็นเว็บแยก HTML / CSS / JS / backend ชัดเจน
ใช้ฐานข้อมูลเดิม (Neon Postgres) โฮสต์ฟรี และแก้ปัญหา Safari ล็อกอินหลุดทุกครั้งที่ปิดแอป

**ผลลัพธ์:** เว็บใหม่ออนไลน์และใช้ข้อมูลจริงครบแล้ว

| รายการ | ค่า |
|---|---|
| เว็บใหม่ | https://shareup-web.mickey01mickey.workers.dev |
| Repo เว็บใหม่ | https://github.com/mickey-mouse-th/ShareUp-web (private) |
| Repo แอปเดิม | https://github.com/mickey-mouse-th/ShareUp |
| Frontend + API | Cloudflare Workers (ฟรี 100,000 request/วัน) เสิร์ฟทั้งสองบนโดเมนเดียว |
| ฐานข้อมูล | Neon Postgres (ฟรี 0.5 GB, region สิงคโปร์) ตัวเดียวกับแอปเดิม |
| รูปสลิป | เก็บใน Neon (ไม่ใช้ R2 เพราะต้องผูกบัตร) |
| ที่เก็บ session | ตาราง `session` ใน Postgres + cookie HttpOnly |

**ทำไม Safari ถึงหายจากปัญหา:** Apps Script ฝังแอปใน iframe ต่างโดเมน Safari เลยล้าง storage ทุกครั้ง
เว็บใหม่เป็นโดเมนเดียวกับ API ใช้ cookie จริง (first-party) จึงไม่ถูกล้าง
**ยังไม่ได้ทดสอบบน Safari/iPhone จริง** (ดูหัวข้อ 6)

---

## 2. โครงสร้างโปรเจกต์ (ShareUp-web)

```
public/                frontend (HTML / CSS / JS ล้วน ไม่มี build step)
  index.html           เปลือกหน้า + sheet ทั้งหมด
  css/                 shared, auth, home, detail, sheets, admin
  js/api.js            ทุกการเรียก backend อยู่ที่นี่ที่เดียว
  js/router.js         URL แยกต่อหน้า
  js/main.js           เริ่มแอป, แท็บ, share mode
  js/pages/            auth, home, detail, txform, eventform, account, share, slips, admin, pdf
src/index.js           Worker (Hono)
src/routes/            auth, home, events, share (สาธารณะ), account, admin, receipts
src/services/          กติกา + SQL (settlement, transactions, share, settings, receipts, ...)
src/middleware/        session cookie, บังคับ JSON
db/schema.sql          schema (ตัวเดียวกับแอปเดิม)
db/migrations/         001_app_setting.sql, 002_receipt_image.sql
test/                  42 ข้อ (node --test)
wrangler.toml          config + ตัวแปร
```

---

## 3. สิ่งที่ทำแล้ว

### 3.1 ฟีเจอร์ (เทียบเท่าแอปเดิม ยกเว้นหัวข้อ 5)
- ล็อกอิน / ออกจากระบบ / จำการล็อกอิน (Remember me) / สมัครสมาชิก (ตอนนี้ปิดไว้)
- หน้าแรก: สถิติ, ตัวกรอง Pending / All / Settled, ค้นหา, สร้าง-แก้ชื่อ-ลบ event, จัดการเพื่อนใน event
- หน้า event: เพิ่ม / แก้ไข / ลบค่าใช้จ่าย (หารเท่ากัน หรือกำหนดเอง), ติ๊ก Paid, สรุปยอดโอนพร้อมติ๊กโอนแล้ว
- สลิป: แนบหลายรูป, ย่อรูปในเบราว์เซอร์ก่อนอัปโหลด (~150-200 KB), ตัวดูแบบเลื่อน, ดาวน์โหลด
- ลิงก์แชร์: view-only หรือ editable (editable ลบอะไรไม่ได้เลย), ผู้เข้าชมไม่ต้องล็อกอิน
- Export PDF (รองรับภาษาไทยผ่านการเรนเดอร์เป็นภาพ)
- โปรไฟล์: เปลี่ยนชื่อ / รูป / รหัสผ่าน (พร้อมรายการเงื่อนไขรหัสผ่านที่ติ๊กเขียวทีละข้อ)
- Admin: จัดการผู้ใช้ (เปิด-ปิดบัญชี, บทบาท, ลบ), ตั้งค่า (อายุ session, นโยบายรหัสผ่าน), ตัวแก้ธีมสี
- URL แยกต่อหน้า: `/`, `/events/:id`, `/admin`, `/admin/users|theme|settings` (รีเฟรชแล้วอยู่หน้าเดิม, ปุ่ม back ใช้ได้)
- ดีไซน์: ย่อทั้งระบบ ~15% ให้เหมาะกับมือถือ, ช่องกรอกคง 16px กัน iPhone ซูมเอง

### 3.2 ประสิทธิภาพ
วัดการโหลดข้อมูลหน้าแรกจากเครื่องไปยัง Neon: **median ~43 ms** (Apps Script ประมาณ 1-3 วินาทีต่อการเรียก ตามที่ประเมิน ยังไม่ได้วัดฝั่ง Apps Script จริง)

### 3.3 ความปลอดภัย (ที่ปรับให้ดีกว่าแอปเดิม)
- ทุกการเขียนเป็น SQL ชุดเดียวที่พิสูจน์ในตัวว่า event / เพื่อน / รายการเป็นของผู้ล็อกอิน
  (แอปเดิมมีหลายจุดที่แก้รายการด้วย id อย่างเดียว หรือไม่เช็กเจ้าของ event เลย)
- ปิดบัญชี / เปลี่ยนบทบาท มีผลทันที (ตรวจจาก DB ทุก request ไม่รอ session หมดอายุ)
- เปลี่ยนรหัสผ่านแล้ว session อื่นของบัญชีนั้นหลุดทันที
- ลิงก์แชร์: ไม่มี route ลบเลย (ไม่ใช่แค่ซ่อนปุ่ม)
- รูป: ตรวจชนิดจากไบต์จริง (รับ JPEG / PNG / WebP ปฏิเสธ SVG), จำกัดขนาด, ไม่เกิน 8 รูปต่อรายการ
- POST/PUT/PATCH/DELETE ต้องเป็น `application/json` (กัน CSRF แบบฟอร์ม), cookie `HttpOnly` + `SameSite=Lax`
- ปิดสมัครเอง (`ALLOW_REGISTRATION="false"`), ปิด preview URLs
- ค่าธีมที่บันทึกต้องเป็น `#rrggbb` เท่านั้น (กันฝัง CSS)

### 3.4 การย้ายข้อมูลจาก Google Sheets (ล้าง DB แล้วย้ายใหม่)

ผลตรวจ `compareSheetsAndDb()` หลังย้าย:

| รายการ | Sheets | Postgres |
|---|---|---|
| บัญชี / เพื่อน / event | 3 / 5 / 5 | 3 / 5 / 5 |
| รายการค่าใช้จ่าย / splits | 81 / 148 | 81 / 148 |
| ลิงก์แชร์ / การโอน / รูปสลิป | 4 / 1 / 64 | 4 / 1 / 64 |
| **ยอดรวม** | 56,301.90 | 56,301.90 |

ส่วนต่าง 2 ตัวที่อธิบายได้ ไม่กระทบเงิน: ผู้ร่วม event 9 → 8 (แถวชี้ไปเพื่อนที่ไม่มีอยู่จริง 1 แถว)
และ Paid 4 → 3 (น่าจะเป็นแถวเศษ/ซ้ำของรายการที่ลบไปแล้ว ยังไม่ได้ยืนยันสาเหตุ)
รูปสลิป 64 รูป (13.5 MB ≈ 2.7% ของ Neon ฟรี) ถูกคัดลอกจาก Drive เข้า DB แล้ว ไฟล์ใน Drive ยังอยู่ครบ

### 3.5 สิ่งที่เพิ่มให้แอป Apps Script เดิม (`Code.js`)
`compareSheetsAndDb()` (เทียบ Sheets กับ Postgres, อ่านอย่างเดียว) · `migrateSlipsToDb()` · `diagnoseShares()` ·
`migrateRestToDb()` ปรับเป็น transaction เดียว (พังแล้ว rollback เองและรันซ้ำได้) · อ่านรูปที่ id ขึ้นต้น `db:` ได้ ·
`.claspignore` (กัน `clasp push` อัปโหลดไฟล์ผิดที่)

---

## 4. การตัดสินใจสำคัญและเหตุผล

| เรื่อง | เลือก | เหตุผล |
|---|---|---|
| Hosting | Cloudflare Workers (ไม่ใช่ Render/Vercel) | ไม่มี cold start, ฟรีตลอด, โดเมนเดียวกับ API |
| เก็บรูป | ใน Neon (ไม่ใช่ R2) | R2 ต้องผูกบัตร และเกินโควตาคิดเงินอัตโนมัติ ไม่มีเพดานหยุดเอง |
| รหัสผ่าน | ยังเป็น SHA-256 ไม่มี salt | ให้แอปเดิมกับแอปใหม่ล็อกอินได้ทั้งคู่ (ดูหัวข้อ 6 ข้อ 3) |
| ตั้งค่า / ธีม | ตาราง `app_setting` (ใหม่) | Worker อ่าน Script Properties ของ Apps Script ไม่ได้ |
| Frontend | HTML / CSS / JS ล้วน + jQuery + select2 | พอร์ตจากโค้ดเดิมตรงตัว ไม่ต้องมี build |

---

## 5. สิ่งที่ยังไม่ได้ทำ

### ค้างจากโค้ด
- **Rate limit** ที่ล็อกอิน / ลิงก์แชร์ / อัปโหลดรูป (เสี่ยงที่สุดที่ยังเปิดอยู่)
- **Security headers** (CSP, กันฝังใน iframe, ห้ามเดาชนิดไฟล์)
- **ภาษาไทย** ทั้งแอปเป็นอังกฤษ
- **PWA** (ไอคอน / ติดตั้งบนหน้าจอโฮม)
- จำตัวกรอง Pending / All / Settled และคำค้นไว้ใน URL · การ์ด event ยังไม่ใช่ลิงก์จริง
- ฟอนต์ไทยใน PDF (โค้ดอ้าง Sarabun แต่ไม่ได้โหลด)
- ทดสอบอัตโนมัติฝั่งหน้าจอ (มีเฉพาะฝั่ง server) · CI / auto-deploy

### ค้างจากการใช้งานจริง
- แอป Apps Script เดิม: ยังอ่านจาก Sheets ซึ่งแท็บถูกเปลี่ยนชื่อแล้ว จึงเห็นข้อมูลว่าง
  ต้องตั้ง `DB_BACKEND = postgres` ใน Script Properties หรือเลิกใช้ (ตอนนี้ยังไม่ได้ตัดสินใจ)
- ลิงก์แชร์เก่าแบบ `script.google.com/...?share=TOKEN` ใช้ไม่ได้ในแอปใหม่โดยตรง
  (token เดิมใช้ได้ถ้าเปลี่ยนเป็น `<โดเมนใหม่>/?share=TOKEN`) ยังไม่ได้แจ้งใคร
- ล้างของสำรอง: ไฟล์สำรอง Google Sheets 4 ไฟล์, branch `backup-before-reset` ใน Neon,
  ไฟล์รูปใน Drive, ฟังก์ชันย้ายข้อมูลใน `Code.js` ยังเก็บไว้ทั้งหมดโดยตั้งใจ
- ยังไม่ได้ตั้งแจ้งเตือนเมื่อเว็บล่ม และยังไม่ได้ตรวจช่วงเวลากู้ข้อมูลย้อนหลังของ Neon ฟรี

---

## 6. ข้อควรระวัง / ข้อจำกัดปัจจุบัน

1. **ยังไม่ได้ยืนยันบนอุปกรณ์จริง:** การล็อกอินค้างบน Safari / iPhone, ขนาดปุ่มหลังย่อ (ปุ่มแก้ไข/ลบในการ์ดเหลือ 28px
   เล็กกว่าที่แนะนำสำหรับนิ้ว ~40-44px), ตัว Export PDF (ทดสอบในเบราว์เซอร์ที่ถูกซ่อนไม่จบ) และ
   การอัปโหลดรูป / ลิงก์แชร์ / หน้า Admin กับ DB จริง (ทดสอบด้วย API จำลองและ unit test เท่านั้น)
2. **`migrateRestToDb` เปลี่ยนชื่อแท็บใน Sheets เป็น `..._v1_backup_<เวลา>` เมื่อสำเร็จ** เป็นจุดตัดขาดตามที่ออกแบบไว้เดิม
   อย่ารันซ้ำ (มีตัวกันพลาด) และห้ามใช้ `TRUNCATE ... CASCADE` กับ `account` โดยไม่ตั้งใจ
   (จะลบตาราง `app_setting` ที่อ้าง `account` ตามไปด้วย ซึ่งเคยเกิดขึ้นครั้งหนึ่ง ครั้งนั้นตารางว่างอยู่แล้วจึงไม่เสียอะไร)
3. **รหัสผ่านยังเป็น SHA-256 ไม่มี salt** เพื่อให้สองแอปใช้ร่วมกันได้ เมื่อเลิกใช้ Apps Script
   ให้ตั้ง `UPGRADE_PASSWORD_HASH="true"` ใน `wrangler.toml` แล้ว deploy (ผู้ใช้จะถูกอัปเกรด hash ตอนล็อกอินครั้งถัดไป)
4. **Neon ฟรี 0.5 GB:** ยังไม่ได้ตรวจยอดใช้งานล่าสุดหลังย้ายใหม่ (ก่อนย้ายอยู่ที่ราว 42 MB และรูปที่ย้ายรอบนี้รวม 13.5 MB)
   รูปใหม่ ~190 KB ต่อรูป จุได้ราว 2,000+ รูป ควรเปิดดูที่ Neon console > Overview เป็นครั้งคราว
5. **Apps Script ที่ `clasp push` แล้วยังไม่ได้ deploy เวอร์ชันใหม่ของเว็บแอป** (ต้อง Deploy > Manage deployments > New version เอง)
6. การลบ event / บัญชี / รายการ ลบรูปในฐานข้อมูลตามไปด้วย แต่ไม่ลบไฟล์ต้นฉบับใน Google Drive

---

## 7. ความผิดพลาดระหว่างทาง (เพื่อไม่ให้เกิดซ้ำ)

| เหตุการณ์ | สาเหตุ | การแก้ |
|---|---|---|
| ผมบอกว่า R2 ใช้ฟรี | ไม่ได้บอกว่าต้องผูกบัตร และเกินโควตาคิดเงินอัตโนมัติ | เปลี่ยนไปเก็บใน Neon |
| `TRUNCATE ... CASCADE` ล้าง `app_setting` ด้วย | ตารางนั้นมี foreign key ไป `account` | ตรวจแล้วว่าว่างอยู่ ไม่เสียอะไร รอบถัดไปล้างเฉพาะตารางที่ต้องการ ไม่ใช้ CASCADE |
| ย้ายข้อมูลรอบแรกค้าง ("server error") | ใส่ข้อมูลทีละแถว ~230 คำสั่งใน autocommit | ปรับเป็น transaction เดียว (all-or-nothing) |
| รัน `migrateRestToDb` ซ้อนกัน 2 รอบ | กด Run ซ้ำ รอบที่สองชนกับ token ที่รอบแรก commit | rollback เองตามที่ออกแบบ ข้อมูลไม่ซ้ำ / ไม่เสีย |
| ช่องกรอกถูกย่อเหลือ 14px | สคริปต์ย่อ CSS ย่อทุกอย่างรวมถึงช่องกรอก | คืน 16px ด้วยกฎเฉพาะ (กัน iPhone ซูมเอง) |
| ฟอร์ม `letter-spacing` พัง (`-.3.5px`) | regex ย่อค่า px อ่านเลขทศนิยมผิด | แก้สคริปต์แล้วย่อใหม่ ตรวจว่าไม่มีค่าผิดรูป |

---

## 8. งานที่แนะนำทำต่อ (เรียงลำดับ)

1. ทดสอบบนมือถือจริง: Safari ล็อกอินค้าง + กดปุ่มเล็กในการ์ด
2. Rate limit + security headers (เสี่ยงต่ำ ปิดช่องที่ค้างอยู่)
3. จัดการแอป Apps Script เดิม (สลับเป็น Postgres หรือปิด) และแจ้งลิงก์ใหม่
4. ตั้งแจ้งเตือนเว็บล่ม + ตรวจการกู้ข้อมูล Neon
5. PWA + ภาษาไทย
6. เมื่อเลิกใช้ Apps Script: `UPGRADE_PASSWORD_HASH="true"` และล้างของสำรอง

---

## 9. คำสั่งที่ใช้บ่อย

```bash
# รันในเครื่อง (ต้องมี .dev.vars ที่มี DATABASE_URL, ไม่ถูก commit)
cd ShareUp-web && npm run dev          # http://localhost:8787  (เพิ่ม -- --ip 0.0.0.0 เพื่อเปิดจากมือถือ)
npm test                               # 42 ข้อ

# Deploy (ล็อกอิน wrangler แล้ว, secret DATABASE_URL ตั้งไว้แล้ว)
npm run deploy

# Apps Script
cd ShareUp && clasp push               # ดูรายการไฟล์ก่อนด้วย: clasp status
```

รัน SQL บน Neon: เลือก branch **production** ก่อนทุกครั้ง (branch `backup-before-reset` ใช้ดูอย่างเดียว ห้ามกด Reset from parent)

---

## 10. Commits สำคัญ

**ShareUp-web:** `75f3817` initial · `7b8a496` สลิปใน Postgres · `1dbb94d` URL แยกหน้า ·
`d352d77` ปิดสมัคร / preview URLs · `eacdbb1` ดีไซน์กะทัดรัด

**ShareUp:** `578068b` prototype วัดเวลา · `6bfebe4` อ่านรูป `db:` + `migrateSlipsToDb` · `5fe63e9` `.claspignore` ·
`dd85b3f` `compareSheetsAndDb` · `5b14fe2` ย้ายแบบ transaction เดียว · `00512cf` `diagnoseShares`
