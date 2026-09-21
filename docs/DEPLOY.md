# Deploy ขึ้น Google Apps Script อัตโนมัติ

ทุกครั้งที่ merge เข้า `main` GitHub Actions
(`.github/workflows/deploy-apps-script.yml`) จะ push โค้ดขึ้น Apps Script
project แล้วสร้าง version ใหม่ของ deployment ให้เอง

สิทธิ์การเข้าถึงเว็บแอปถูกบังคับเป็น **Anyone (ทุกคน เข้าได้โดยไม่ต้องล็อกอิน)**
และรันในสิทธิ์ของเจ้าของสคริปต์ ทุกครั้งที่ deploy

---

## ตั้งค่าครั้งแรก (ทำครั้งเดียว)

### 1. เปิด Apps Script API

เข้า https://script.google.com/home/usersettings แล้วเปิด **Google Apps Script API**
ถ้าไม่เปิด `clasp push` จะขึ้น error 403

### 2. หา Script ID

เปิด Apps Script project → **Project Settings (⚙)** → คัดลอก **Script ID**

### 3. หา Deployment ID ของตัวที่เครื่องสแกนใช้อยู่

เปิด Apps Script project → **Deploy → Manage deployments** → เลือก deployment
ที่เครื่องหน้างานใช้อยู่ → คัดลอก **Deployment ID**

> สำคัญ: ถ้าไม่ใส่ค่านี้ workflow จะ **สร้าง deployment ใหม่ซึ่งได้ URL ใหม่**
> เครื่องสแกนที่หน้างานยังชี้ URL เดิมและจะไม่เห็นการอัปเดตเลย

### 4. สร้าง credential ของ clasp

รันบนเครื่องตัวเอง (ต้องมี Node.js):

```bash
npm install -g @google/clasp@2.4.2
clasp login
```

ล็อกอินด้วยบัญชี Google ที่เป็นเจ้าของ Apps Script project จากนั้นเปิดไฟล์

- macOS / Linux: `~/.clasprc.json`
- Windows: `C:\Users\<ชื่อผู้ใช้>\.clasprc.json`

คัดลอก **เนื้อหาทั้งไฟล์**

> ไฟล์นี้มี refresh token ที่เข้าถึงบัญชี Google ได้ อย่า commit ลง repo
> และอย่าส่งให้ใครทางแชต ใส่เป็น repository secret เท่านั้น

### 5. ใส่ repository secrets

GitHub → repo นี้ → **Settings → Secrets and variables → Actions → New repository secret**

| ชื่อ secret | ค่า | จำเป็น |
|---|---|---|
| `CLASPRC_JSON` | เนื้อหาทั้งไฟล์ `~/.clasprc.json` จากข้อ 4 | ✅ |
| `APPS_SCRIPT_ID` | Script ID จากข้อ 2 | ✅ |
| `APPS_SCRIPT_DEPLOYMENT_ID` | Deployment ID จากข้อ 3 | แนะนำอย่างยิ่ง |

---

## การใช้งาน

**อัตโนมัติ** — merge เข้า `main` เมื่อมีการแก้ `index.html`, `scr/**` หรือ
`appsscript.json` แล้ว workflow จะทำงานเอง

**ด้วยมือ** — Actions → *Deploy to Apps Script* → **Run workflow**

เสร็จแล้วให้ refresh หน้าจอที่เครื่องสแกนแต่ละไลน์ เพื่อโหลดเวอร์ชันใหม่

---

## workflow ทำอะไรบ้าง

1. ตรวจว่ามี secret ครบ ถ้าไม่ครบจะหยุดพร้อมบอกว่าขาดตัวไหน
2. `clasp pull` ดึงโค้ดปัจจุบันจาก Apps Script มาเก็บเป็น **artifact
   `apps-script-backup-<sha>`** (เก็บไว้ 30 วัน) — ใช้กู้คืนได้ถ้า deploy ผิดพลาด
3. เตรียมไฟล์ลงโฟลเดอร์ `build/`
   - `index.html` → **`Index.html`** (ตัว `I` ใหญ่ เพราะ `doGet()` เรียก
     `createHtmlOutputFromFile('Index')` — ถ้าชื่อไม่ตรง เว็บแอปจะพัง)
   - `scr/backend.gs` → `backend.gs`
   - manifest: ใช้ของ project จริงเป็นฐาน (รักษา timezone / runtime /
     oauthScopes เดิมไว้) แล้วบังคับเฉพาะส่วน `webapp` ให้เป็น
     `access: ANYONE_ANONYMOUS` + `executeAs: USER_DEPLOYING`
4. `clasp push --force`
5. `clasp deploy -i <DEPLOYMENT_ID>` สร้าง version ใหม่ใน deployment เดิม
   (URL ไม่เปลี่ยน)

---

## ข้อควรระวัง

- **`clasp push --force` เขียนทับ project ทั้งหมด** ไฟล์ใดที่อยู่ใน Apps Script
  แต่ไม่มีในรีโปจะถูกลบ ก่อนเปิดใช้ครั้งแรกให้ `clasp pull` มาเทียบก่อนว่า
  `scr/backend.gs` ในรีโปตรงกับโค้ดที่รันอยู่จริง
- ตั้งแต่นี้ไป **อย่าแก้โค้ดในหน้าเว็บ Apps Script โดยตรง** เพราะ deploy
  รอบถัดไปจะทับทิ้ง ให้แก้ผ่าน PR แล้ว merge เข้า `main` แทน
- workflow นี้ทำงานเฉพาะตอน push เข้า `main` และตอนสั่งด้วยมือ — ไม่รันบน PR
  จึงไม่มีการ deploy จาก branch ที่ยังไม่ merge
- `ADMIN_PASSWORD` และ `CAP_SPREADSHEET_ID` ยัง hard-code อยู่ใน
  `scr/backend.gs` ซึ่ง repo นี้เป็น public — ควรย้ายไปเก็บใน Script Properties
  แยกต่างหาก (คนละเรื่องกับ workflow นี้)
