# Inventory Management System

> ระบบจัดการคลังอุปกรณ์แบบครบวงจร พร้อมระบบแจ้งเตือนอัตโนมัติทางอีเมล

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat&logo=google&logoColor=white)](https://script.google.com/)
[![Version](https://img.shields.io/badge/version-2.0-blue.svg)](https://github.com)
[![License](https://img.shields.io/badge/license-Educational-green.svg)](https://github.com)

**📧 Deployment Account**: `noreply.inventory.ims@gmail.com`  
**🗓️ Last Updated**: February 15, 2026  
**📦 Version**: 2.0 - Email Notifications & Auto-Archive

---

## 📋 สารบัญ

- [ภาพรวมระบบ](#-ภาพรวมระบบ)
- [ฟีเจอร์หลัก](#-ฟีเจอร์หลัก)
- [สิทธิ์การใช้งาน](#-สิทธิการใช้งาน)
- [เทคโนโลยี](#️-เทคโนโลยีที่ใช้)
- [โครงสร้างไฟล์](#-โครงสร้างไฟล์)
- [โครงสร้าง Database](#-โครงสร้าง-database)
- [คู่มือการติดตั้ง](#-คู่มือการติดตั้ง)
- [การใช้งาน](#-การใช้งาน)
- [ระบบอีเมล](#️-ระบบอีเมล)
- [Troubleshooting](#-troubleshooting)

---

## 🎯 ภาพรวมระบบ

ระบบจัดการคลังอุปกรณ์แบบ Web Application ที่พัฒนาด้วย Google Apps Script ใช้ Google Sheets เป็น Database มีระบบสิทธิ์การใช้งาน 3 ระดับ (Admin, User, Guest) พร้อมระบบแจ้งเตือนทางอีเมลและ Auto-archive อัตโนมัติ

### ✨ จุดเด่น

- ✅ **ฟรี 100%** - ไม่มีค่าใช้จ่าย ใช้ Google Services
- ✅ **ระบบอีเมลอัตโนมัติ** - แจ้งเตือนคำขอใหม่และรายการเกินกำหนด
- ✅ **Auto-archive** - ย้ายข้อมูลเก่าอัตโนมัติ
- ✅ **Partial Return** - คืนอุปกรณ์บางส่วนได้
- ✅ **Responsive Design** - ใช้งานได้ทุกอุปกรณ์
- ✅ **Audit Trail** - บันทึกทุก activity

---

## 🚀 ฟีเจอร์หลัก

### 🔐 Authentication & Authorization

- Login/Logout ด้วย Session Management (24 ชม.)
- Password hashing ด้วย SHA-256
- 3 ระดับสิทธิ์: **Admin**, **User**, **Guest**
- Email case-insensitive

### 📦 การจัดการอุปกรณ์ (Admin)

- ➕ เพิ่ม/แก้ไข/ลบ อุปกรณ์และอุปกรณ์เสริม
- 📊 ติดตามจำนวนคงคลัง (Total/Available)
- 🔍 ค้นหาและกรองตามสถานะ
- 🖼️ เพิ่มรูปภาพผ่าน URL
- 🔄 Many-to-many mapping ระหว่าง Items และ Accessories

### 📝 ระบบยืม-คืน

**สร้างคำขอ (User/Admin):**
- เลือกอุปกรณ์หลายรายการพร้อม accessories
- ระบุจำนวน วันที่ยืม วันที่คืน
- Shopping cart system
- แก้ไขคำขอได้ก่อนอนุมัติ

**จัดการคำขอ (Admin):**
- ✅ อนุมัติและจ่าย → หักสต็อกอัตโนมัติ
- 🔄 คืนทั้งหมด/บางส่วน → คืนสต็อกอัตโนมัติ
- ❌ ยกเลิกคำขอ

**สถานะคำขอ:**
```
Submit → Distributed → Partial_Returned → Returned
                    ↘ Cancelled
```

### 👥 จัดการผู้ใช้ (Admin)

- เพิ่ม/แก้ไข/ลบ users
- กำหนดสิทธิ์ (Admin/User/Guest)
- เปิด-ปิดการใช้งาน
- เปลี่ยนรหัสผ่าน
- ตั้งค่าผู้รับอีเมล (`Can_Send_Email`)

### 📧 ระบบแจ้งเตือนอีเมล

**1. แจ้งคำขอยืมใหม่**
- ส่งอัตโนมัติเมื่อมีคำขอใหม่
- แสดงรายละเอียดคำขอและรูปอุปกรณ์
- มีลิงก์เข้าสู่ระบบอนุมัติ

**2. แจ้งเตือนเกินกำหนด**
- ส่งทุกวันเวลา 08:30 น.
- แสดงจำนวนรายการที่เกินกำหนด
- แสดงสถานะคืนบางส่วน
- ระบุจำนวนค้างคืนของแต่ละรายการ

**รองรับ 2 ระบบ:**
- 📨 **Gmail (MailApp)** - ส่งฟรี 100 email/day
- 🚀 **Mailjet** - ส่งได้ไม่จำกัด (ต้อง verify email)

### ⏰ ระบบ Triggers อัตโนมัติ

- 🗄️ **Auto-archive** (02:00 ทุกวัน)
  - ย้ายคำขอเก่า (Cancelled/Returned > 7 วัน) ไป Historical sheets
  
- 📬 **Overdue email** (08:30 ทุกวัน)
  - ส่งอีเมลแจ้งเตือนรายการเกินกำหนด

### 📊 Dashboard & Reports (Admin)

- สถิติอุปกรณ์และคำขอ real-time
- Activity logs 3 ประเภท:
  - **Request Activity** - Submit, Distribute, Return, Cancel
  - **System Activity** - Login, Logout
  - **Inventory Activity** - Item/Accessory CRUD
- Historical requests viewer

---

## 👤 สิทธิ์การใช้งาน

| ฟีเจอร์ | Admin | User | Guest |
|---------|:-----:|:----:|:-----:|
| ดูรายการอุปกรณ์ | ✅ | ✅ | ✅ |
| สร้างคำขอยืม | ✅ | ✅ | ❌ |
| แก้ไขคำขอของตัวเอง | ✅ | ✅ | ❌ |
| อนุมัติ/จ่าย/คืนอุปกรณ์ | ✅ | ❌ | ❌ |
| จัดการอุปกรณ์/Accessories | ✅ | ❌ | ❌ |
| จัดการผู้ใช้ | ✅ | ❌ | ❌ |
| Dashboard & Activity Logs | ✅ | ❌ | ❌ |

---

## 🛠️ เทคโนโลยีที่ใช้

### Backend
- **Google Apps Script** - Server-side JavaScript
- **Google Sheets** - Database (NoSQL-like)

### Frontend
- **HTML5** - Structure
- **Tailwind CSS v4** - Styling & Responsive
- **Vanilla JavaScript** - Client logic
- **Moment.js** - Date/Time handling
- **Font Awesome** - Icons
- **SweetAlert2** - Beautiful alerts

### Email
- **MailApp** - Gmail API (default)
- **Mailjet** - Third-party email service (optional)

---

## 📁 โครงสร้างไฟล์

```
app-script/
├── 📄 code.gs              # Main logic (3,836 lines)
│   ├── CRUD operations
│   ├── Stock management
│   ├── Email functions
│   └── Dashboard & logs
│
├── 🔐 auth.gs              # Authentication (262 lines)
│   ├── Password hashing
│   ├── User authentication
│   └── Session management
│
├── 🗄️ archive.gs           # Auto-archive (106 lines)
│   └── Move old requests to Historical
│
├── ⏰ triggers.gs          # Triggers (386 lines)
│   ├── Setup/Remove triggers
│   └── Test & debug functions
│
├── 🎨 index.html           # Frontend UI (5,000+ lines)
│   ├── Tailwind CSS
│   └── Responsive design
│
└── 📖 README.md            # Documentation

**Total**: ~9,600 lines of code
```

---

## 📊 โครงสร้าง Database

### Google Sheets Structure (13 Sheets)

#### 1️⃣ **Users** - ผู้ใช้งาน
```
Email | Password | Permission | Active | Can_Send_Email
```

#### 2️⃣ **Items** - อุปกรณ์
```
Item_Id | Item_Name | Item_Desc | Total_Qty | Available_Qty | Image | Active | Created_By | Created_At | Modified_By | Modified_At
```

#### 3️⃣ **Accessories** - อุปกรณ์เสริม
```
Accessory_Id | Accessory_Name | Accessory_Desc | Total_Qty | Available_Qty | Active | Created_By | Created_At | Modified_By | Modified_At
```

#### 4️⃣ **Item_Accessory_Mapping** - ความสัมพันธ์
```
Mapping_Id | Item_Id | Accessory_Id | Created_By | Created_At | Active
```

#### 5️⃣ **Requests** - คำขอยืม
```
Request_Id | Requirer_Name | Status | Request_Date | Distributed_Date | Return_Date | Remark | Created_By | Created_At | Modified_By | Modified_At
```

#### 6️⃣ **Request_Item** - รายการอุปกรณ์ในคำขอ
```
Request_Id | Item_Index | Item_Id | Item_Name | Qty | Returned_Qty | Status
```

#### 7️⃣ **Request_Item_Accessory** - รายการอุปกรณ์เสริม
```
Request_Id | Item_Index | Accessory_Index | Accessory_Id | Accessory_Name | Qty | Returned_Qty | Status
```

#### 8️⃣ **Sessions** - Session management
```
Session_Id | Email | Permission | Created_At | Last_Activity
```

#### 9️⃣ **Activity Logs** - 3 Sheets
- **Request_Activity** - Request operations
- **System_Activity** - Login/Logout
- **Inventory_Activity** - Item/Accessory CRUD

```
Log_Id | Email | Activity | Action_At
```

#### 🔟 **Historical** - 3 Sheets (Archive)
- **Historical_Requests**
- **Historical_Request_Item**
- **Historical_Request_Item_Accessory**

*(โครงสร้างเหมือน Requests, Request_Item, Request_Item_Accessory)*

---

## 🚀 คู่มือการติดตั้ง

### Step 1: สร้าง Google Sheets

1. สร้าง Google Sheets ใหม่
2. สร้าง Sheets ทั้ง 13 sheets ตามโครงสร้างข้างต้น
   - หรือใช้ function `repairSheets()` สร้างอัตโนมัติ

### Step 2: Setup Apps Script

1. เปิด **Extensions → Apps Script**
2. สร้างไฟล์ทั้งหมด:
   - `code.gs`
   - `auth.gs`
   - `archive.gs`
   - `triggers.gs`
   - `index.html`
3. Copy code จาก repository

### Step 3: Deploy Web App

1. คลิก **Deploy → New deployment**
2. เลือก type: **Web app**
3. ตั้งค่า:
   ```
   Execute as: Me (noreply.inventory.ims@gmail.com)
   Who has access: Anyone
   ```
4. คลิก **Deploy**
5. **Copy Web App URL**

> ⚠️ **สำคัญ**: Deploy ด้วยบัญชี `noreply.inventory.ims@gmail.com` เพื่อส่งอีเมลจากบัญชีนี้

### Step 4: สร้าง Admin User

1. เปิด **Users sheet**
2. เพิ่มข้อมูล:
   ```
   Email: admin@admin.com
   Password: [รัน hashPassword("admin123")]
   Permission: Admin
   Active: TRUE
   Can_Send_Email: TRUE
   ```

### Step 5: ตั้งค่าอีเมล

#### ตัวเลือก A: Gmail (แนะนำ) 📨

ใช้ค่าเริ่มต้น - ไม่ต้องตั้งค่า:
```javascript
const EMAIL_USE_MAILJET = false;
```

**ข้อจำกัด**: 100 email/day

#### ตัวเลือก B: Mailjet 🚀

1. สมัคร [Mailjet](https://www.mailjet.com/)
2. รับ API Key & Secret
3. Verify email `noreply.inventory.ims@gmail.com`:
   - Account Settings → Senders & Domains
   - Add sender
   - Check email verification
4. แก้ไข `code.gs`:
   ```javascript
   const EMAIL_USE_MAILJET = true;
   const MAILJET_API_KEY = "your-key";
   const MAILJET_API_SECRET = "your-secret";
   const MAILJET_FROM_EMAIL = "noreply.inventory.ims@gmail.com";
   ```
5. Deploy ใหม่

**ข้อดี**: ไม่จำกัดจำนวน, Tracking, Better deliverability

### Step 6: ตั้งค่า Triggers

1. เปิด Apps Script Editor
2. รัน function: `setupAllTriggers()`
3. **Authorize permissions**:
   - Gmail
   - SpreadsheetApp
   - ScriptApp
4. ตรวจสอบ: รัน `listAllTriggers()`

**Triggers ที่สร้าง:**
- 🗄️ Archive: ทุกวัน 02:00 น.
- 📬 Overdue email: ทุกวัน 08:30 น.

### Step 7: ทดสอบระบบ

1. เปิด Web App URL
2. Login ด้วย `admin@admin.com` / `admin123`
3. สร้างคำขอยืมทดสอบ
4. ตรวจสอบอีเมลที่ users ที่ตั้งค่า `Can_Send_Email = TRUE`

✅ **เสร็จสิ้น!**

---

## 💼 การใช้งาน

### 🧑 User Workflow

1. **Login** เข้าสู่ระบบ
2. **Browse** ดูรายการอุปกรณ์
3. **Add to Cart** เลือกอุปกรณ์ที่ต้องการ
4. **Create Request** สร้างคำขอยืม
5. **Wait** รอ Admin อนุมัติ

### 👨‍💼 Admin Workflow

1. **Manage Items** จัดการอุปกรณ์และ accessories
2. **Review Requests** ตรวจสอบคำขอใหม่
3. **Approve & Distribute** อนุมัติและจ่ายอุปกรณ์
4. **Track Returns** บันทึกการคืน (เต็มจำนวนหรือบางส่วน)
5. **View Dashboard** ดูสถิติและ activity logs
6. **Manage Users** จัดการผู้ใช้และสิทธิ์

---

## ✉️ ระบบอีเมล

### Email Templates

#### 1. คำขอยืมใหม่

**Subject:**
```
[คำขอยืมใหม่] Request #XXX - ชื่อผู้ยืม
```

**Content:**
- ✅ รายละเอียดคำขอ (เลขที่, ผู้ยืม, วันที่)
- 🖼️ รายการอุปกรณ์พร้อมรูปภาพ
- 🔗 ลิงก์เข้าสู่ระบบอนุมัติ
- 📧 ผู้รับ: Users ที่ `Can_Send_Email = TRUE`

**ตัวอย่าง:**

![New Request Email](https://via.placeholder.com/600x400?text=New+Request+Email+Template)

#### 2. แจ้งเตือนเกินกำหนด

**Subject:**
```
[แจ้งเตือนเกินกำหนด] X รายการ
```

**Content:**
- 📊 สรุปจำนวนรายการ
- 📅 รายละเอียดแต่ละคำขอ (เกินกำหนดกี่วัน)
- 🔄 สถานะคืนบางส่วน (หากมี)
- 📦 รายการอุปกรณ์ค้างคืนพร้อมจำนวน
- ⏰ ส่ง: ทุกวัน 08:30 น.

**ตัวอย่าง:**

![Overdue Email](https://via.placeholder.com/600x400?text=Overdue+Email+Template)

### การตั้งค่าผู้รับอีเมล

1. เปิด **Users sheet**
2. ตั้งค่า column `Can_Send_Email`:
   - `TRUE` = รับอีเมล
   - `FALSE` หรือเว้นว่าง = ไม่รับอีเมล

---

## 🔧 Troubleshooting

### 🚫 อีเมลไม่ออก

**ตรวจสอบ:**
1. ✅ มี users ที่ตั้งค่า `Can_Send_Email = TRUE` หรือไม่?
2. ✅ รัน `testEmailSending()` เพื่อทดสอบ
3. ✅ ตรวจสอบ authorization (Gmail scope)
4. ✅ ดู Execution log ใน Apps Script Editor

**แก้ไข:**
```javascript
// ทดสอบการส่งอีเมล
testEmailSending();

// ตรวจสอบผู้รับ
getEmailRecipients();
```

### ⏰ Triggers ไม่ทำงาน

**ตรวจสอบ:**
1. ✅ รัน `listAllTriggers()` - มี triggers หรือไม่?
2. ✅ ดู Trigger executions (Apps Script → Triggers → Executions)
3. ✅ Check authorization

**แก้ไข:**
```javascript
// ลบและสร้างใหม่
removeAllTriggers();
setupAllTriggers();

// ตรวจสอบ
listAllTriggers();
```

### 🔄 ข้อมูลไม่อัพเดท

**แก้ไข:**
```javascript
// ปิด cache
const ENABLE_CACHE = false;
```

- รีเฟรชหน้าเว็บ (Ctrl+F5)
- ตรวจสอบชื่อ sheets ถูกต้อง

### 🔒 Authorization Issues

**วิธีแก้:**
1. รัน function ใน Apps Script Editor (ไม่ใช่ Web App)
2. คลิก "Review Permissions"
3. เลือกบัญชี Google
4. คลิก "Allow" ทุก scope

---

## 🔒 ความปลอดภัย

- ✅ **Password Hashing**: SHA-256
- ✅ **Session Management**: 24-hour expiration
- ✅ **Permission Control**: Role-based access
- ✅ **Input Validation**: Frontend & Backend
- ✅ **Audit Trail**: บันทึกทุก action
- ✅ **Email Normalization**: Case-insensitive

---

## 📈 System Statistics

| Metric | Value |
|--------|-------|
| **Total Lines** | ~9,600 |
| **Backend** | ~4,500 lines |
| **Frontend** | ~5,000 lines |
| **Sheets** | 13 sheets |
| **Functions** | 50+ |
| **Email Templates** | 2 |
| **Triggers** | 2 |
| **Supported Users** | Unlimited |

---

## 🔮 Future Enhancements

- [x] ~~Email notifications~~ ✅ Done
- [x] ~~Auto-archive~~ ✅ Done
- [ ] **Line Notify** integration
- [ ] **QR Code** tracking
- [ ] **Export reports** (PDF/Excel)
- [ ] **Image upload** (Google Drive)
- [ ] **Advanced search** & filters
- [ ] **Calendar view** booking
- [ ] **Mobile App** (PWA)
- [ ] **Custom email templates**

---

## 📝 Changelog

### Version 2.0 (Feb 15, 2026)
- ✨ Email notification system
- ✨ Auto-archive feature
- ✨ Mailjet integration
- ✨ Partial return support
- 🐛 Bug fixes & improvements

### Version 1.0
- 🎉 Initial release
- ✅ Basic CRUD operations
- ✅ Request management
- ✅ User authentication

---

## 📄 License

This project is for **educational purposes**.

---


## 🙏 Acknowledgments

- Google Apps Script Team
- Tailwind CSS
- Moment.js
- Font Awesome
- SweetAlert2
- Mailjet

---