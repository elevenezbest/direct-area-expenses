# Direct Area Expenses — คู่มือ Deploy & เชื่อม Google Sheets

> ## ✅ Deploy แล้ว (live)
> - **เว็บ:** https://elevenezbest.github.io/direct-area-expenses/
> - **Apps Script Web App:** `https://script.google.com/macros/s/AKfycbzPPi4bpJQ05EaQKQ-ejTRtBqncwZ19Ppd3G0ZCggI0pRvPnECkerorf32wonamnh2dWQ/exec`
> - **Repo:** https://github.com/elevenezbest/direct-area-expenses (GitHub Pages: branch `main` /docs)
> - แก้โค้ดแล้วอัปเดตเว็บ: `git add -A && git commit -m "..." && git push` (Pages rebuild อัตโนมัติ)
> - แก้ Apps Script แล้ว: Apps Script → Deploy → Manage deployments → ✏️ → Version: New version → Deploy (URL เดิม)


ต่อหลังบ้านให้ `Direct Area Expenses.dc.html` ใช้งานข้ามเครื่องจริง โดย**ไม่แก้ UI/flow เดิม**
หน้าเว็บคุยกับ Google Sheets ผ่าน **Apps Script Web App** (อ่าน DATA/เป้าหมายปีนี้/Admin M{n}, เขียน "ข้อมูลค่าใช้จ่าย")

```
docs/                         ← ตัวเว็บสำหรับ GitHub Pages / Netlify
  index.html                  ← แอป (สำเนาของ Direct Area Expenses.dc.html) — ตัวที่เว็บเสิร์ฟ
  Direct Area Expenses.dc.html← ไฟล์ต้นฉบับ (re-import กลับ Claude Design ได้)
  support.js                  ← DC runtime (โหลด React จาก unpkg เอง)
  uploads/                    ← asset
apps-script/Code.gs           ← โค้ดหลังบ้าน (วางใน Apps Script)
```

> **ไฟล์ที่เป็น source of truth คือ `docs/index.html`** (และสำเนา `.dc.html` ข้างกัน) — แก้ที่นี่ต่อไป

---

## ระบบ Login (ของจริง — เก็บในชีต)
- บัญชีเก็บในแผ่น **"ผู้ใช้งาน"** (สร้างอัตโนมัติในไฟล์ค่าใช้จ่าย) คอลัมน์: id · role · hub · name · password · firstLogin
- บัญชีเริ่มต้น 6 ราย: `bestdiarea` (แอดมิน) · `21BPL` `22PDT` `23AYU` `65WNO` `99BAG` (สมาชิกรายฮับ)
- **รหัสเข้าครั้งแรกของทุกคน = `123456`** → เข้าครั้งแรกระบบบังคับตั้งรหัสใหม่ของตัวเอง (≥6 ตัว)
- **ลืมรหัส:** แอดมิน → เมนู 👥 ผู้ใช้งาน → จัดการรหัส → "รีเซ็ต" (กลับเป็น 123456 + ให้ตั้งใหม่) หรือใช้ปุ่มในแอป
- เพิ่มผู้ใช้ใหม่: แอดมิน → เพิ่มผู้ใช้ (รหัสเริ่มต้น 123456) — บันทึกลงชีตจริง
- รหัสผ่านเก็บแบบ plaintext ในชีตส่วนตัวของคุณ (เหมาะกับใช้งานภายใน ไม่ใช่ระบบความปลอดภัยสูง)

## ขั้นที่ 1 — เตรียม Google Sheets (เจ้าของไฟล์ทำ)

1. **แชร์สิทธิ์** ทั้ง 2 สเปรดชีตเป็น **"ผู้มีลิงก์ดูได้"** (อ่านอย่างน้อย):
   - ค่าใช้จ่าย: `1F4djjhXTzhKrqXQhPVPDtjjccOfQHRFpQfvbbw-KCrs`
   - Performance: `18N5MmcIXpF0AS6DoVc1UxmvoBZkW4ViAi9nls4lsEP8`
2. ชีตที่ระบบใช้ (ตรวจแล้วว่ามีจริง ใช้ได้เลย):
   - ไฟล์ **ค่าใช้จ่าย** (`1F4d…`): `ข้อมูลค่าใช้จ่าย` (ว่าง พร้อมหัวคอลัมน์), `DATA` (ตารางอ้างอิง: hub map + ตัวเลือก type/detail)
   - ไฟล์ **HUB ADMIN PERFORMANCE** (`18N5…`): `Admin M{n}` (คะแนนแอดมิน), **`เป้าหมายปีนี้`** (เป้าหมาย Cost/parcel)
3. **`เป้าหมายปีนี้` มีจริงในไฟล์ PERF แล้ว — ไม่ต้องสร้างใหม่.** เป็นตารางแนวนอน 3 บล็อก (Area / HUB / Cost/parcel) ของหมวด ค่าใช้จ่ายทั่วไป (EXP) · วัสดุสิ้นเปลือง (CON) · ค่าน้ำ-ค่าไฟ (ELEC). ระบบ parse อัตโนมัติตามชื่อหัวบล็อก → เติม `rcTarget` รายฮับ (RENT ไม่มีในชีต → คงค่าเดิม, แถว AREA = ค่ารวม จะถูกข้าม). ค่าที่อ่านได้ยืนยันตรงกับชีตจริงทุกฮับแล้ว

---

## ขั้นที่ 2 — Deploy Apps Script Web App

1. เปิดชีต **ค่าใช้จ่าย** → เมนู **Extensions → Apps Script**
2. ลบโค้ดเดิม วางทั้งหมดจาก [`apps-script/Code.gs`](apps-script/Code.gs) → 💾 Save
3. **Deploy → New deployment** → ⚙️ เลือกชนิด **Web app**
   - **Execute as:** Me
   - **Who has access:** Anyone
4. กด **Deploy** → อนุญาตสิทธิ์ (ครั้งแรก Google จะเตือน "unverified" → Advanced → Go to project)
5. คัดลอก **Web app URL** ที่ลงท้าย `…/exec`
6. ทดสอบ: เปิด `…/exec?action=ping` ในเบราว์เซอร์ ควรได้ `{"ok":true,"pong":true}`

> แก้โค้ดภายหลังต้อง **Deploy → Manage deployments → ✏️ → Version: New** ทุกครั้ง URL ถึงจะอัปเดต

---

## ขั้นที่ 3 — ใส่ URL ลงหน้าเว็บ

เลือก **วิธีใดวิธีหนึ่ง**:

**ก) แก้ในไฟล์** — เปิด `docs/index.html` หา `WEBAPP_URL = ''` ใส่ URL:
```js
WEBAPP_URL = 'https://script.google.com/macros/s/AKfy.../exec';
```
(ถ้าต้องการ re-import Claude Design ด้วย ให้แก้ `docs/Direct Area Expenses.dc.html` ให้ตรงกัน)

**ข) ไม่แตะไฟล์แอป** — เพิ่มบรรทัดนี้ใน `<head>` ของ `docs/index.html` ก่อน `support.js`:
```html
<script>window.DA_WEBAPP_URL = 'https://script.google.com/macros/s/AKfy.../exec';</script>
```

> เว้นว่างไว้ = แอปรันแบบ prototype เดิม (mock/localStorage) ทุกอย่างไม่พัง

---

## ขั้นที่ 4 — Deploy หน้าเว็บ (GitHub Pages)

ติดตั้ง gh + git ไว้แล้ว ทำใน PowerShell ที่โฟลเดอร์นี้:

```powershell
# 1) ล็อกอิน GitHub (ครั้งเดียว)
& "C:\Program Files\GitHub CLI\gh.exe" auth login    # GitHub.com → HTTPS → web browser

# 2) สร้าง repo + push (โค้ด commit ไว้ให้แล้ว)
& "C:\Program Files\GitHub CLI\gh.exe" repo create direct-area-expenses --public --source . --remote origin --push

# 3) เปิด GitHub Pages: branch main / โฟลเดอร์ /docs
& "C:\Program Files\GitHub CLI\gh.exe" api -X POST repos/{owner}/direct-area-expenses/pages -f "source[branch]=main" -f "source[path]=/docs"
```
รอ ~1 นาที เว็บจะอยู่ที่ `https://<user>.github.io/direct-area-expenses/`

### ทางเลือก: Netlify (ลากวาง ไม่ต้อง CLI)
ลากโฟลเดอร์ `docs/` ทั้งโฟลเดอร์ไปวางที่ https://app.netlify.com/drop → ได้ URL ทันที

---

## การเชื่อมข้อมูล (ทำอะไรบ้าง)

| การทำงานในแอป | เมธอด | หลังบ้าน |
|---|---|---|
| โหลดตัวเลือก type/detail | `loadOptions()` | GET `?action=options` ← DATA |
| เติมยอดค่าใช้จ่าย 4 หมวด KPI | `loadCostData()` | GET `?action=cost` ← รวมแถว "ผ่าน" จาก ข้อมูลค่าใช้จ่าย (ตามฮับ/หมวด/วีค) |
| เติมเป้าหมาย Cost/parcel | `loadTargets()` | GET `?action=targets` ← เป้าหมายปีนี้ (ไฟล์ PERF) → rcTarget EXP/CON/ELEC |
| คะแนนแอดมิน | `loadPerf(n)` | GET `?action=perf&m=n` ← Admin M{n} (fallback gviz/seed) |
| บันทึกรายการเบิก | `submit()` | POST `append` → เพิ่มแถว (status "รอ") + จำ `sheetRow` |
| อนุมัติ / ไม่อนุมัติ | `approve()` / `rejectConfirm()` | POST `updateStatus` → แก้คอลัมน์ สถานะ |
| เพิ่มตัวเลือกใหม่ | `addOpt()` | POST `addOption` → เพิ่มเข้า DATA + เคลียร์ `pendingSheetOpts` |

**ข้อจำกัด / ที่เหลือ**
- การ map ประเภท→หมวด ใช้ `typeToCat` เดิม + fallback ยืดหยุ่น (รองรับชื่อใน DATA ที่ต่างเล็กน้อย เช่น "วัสดุสิ้นเปลือง" / "ค่าวัสดุสิ้นเปลือง")
- `ชื่อHUB` เขียนเป็นตัวย่อ (เช่น `21BPL`) ตรงกับ DATA; ตอนอ่านรวมยอด normalize กลับเป็น `HUB21BPL`
- การเขียนทั้งหมดเป็น background sync — ถ้า URL ว่าง/ออฟไลน์ แอปยังทำงาน optimistic เหมือนเดิม
- localStorage เดิม (`daExpSlideEdits`, `daExpOpts`, `daExpMonthCloseAt`) ยังคงไว้ ไม่ย้ายขึ้นชีต (นอกสโคป)
- ระบบล็อกอิน/หลักฐานไฟล์/LINE Notify ยังเป็น mock ตามเดิม
