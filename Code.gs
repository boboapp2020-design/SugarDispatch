/**
 * Sugar Delivery — Google Apps Script backend
 * เก็บข้อมูล User + บันทึกการส่งมอบน้ำตาล ลงใน Google Sheet
 *
 * ===== วิธีติดตั้ง (ทำครั้งเดียว ~3 นาที) =====
 * 1. เปิด Google Sheet ของคุณ:
 *    https://docs.google.com/spreadsheets/d/1uUsnIx34UE3p1_Z1GHGix8BLi9-eFoTtxBN4xBfIakg/edit
 * 2. เมนู Extensions (ส่วนขยาย) → Apps Script
 * 3. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดไฟล์นี้ลงไป → กด Save (💾)
 * 4. กดปุ่ม Deploy (ทำให้ใช้งานได้) → New deployment
 *    - ประเภท: Web app
 *    - Execute as: Me (ฉัน)
 *    - Who has access: Anyone (ทุกคน)   ← สำคัญมาก
 *    → กด Deploy → อนุญาตสิทธิ์บัญชี Google ของคุณ
 * 5. คัดลอก "Web app URL" (ลงท้ายด้วย /exec)
 * 6. เปิดแอป → หน้า Login → ⚙ สำหรับผู้ดูแลระบบ → วาง URL ในช่อง
 *    "เชื่อมต่อ Google Sheets" → กด บันทึก & ทดสอบการเชื่อมต่อ
 *
 * ระบบจะสร้างชีต "Users" และ "Records" ให้อัตโนมัติ
 * หมายเหตุ: ถ้าแก้โค้ดนี้ภายหลัง ต้อง Deploy → Manage deployments → Edit → New version
 */

var USERS_SHEET = "Users";
var RECORDS_SHEET = "Records";

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "";
  if (action === "getUsers") return jsonOut(getUsers());
  if (action === "getRecords") return jsonOut(getRecords());
  return jsonOut({ ok: true, service: "sugar-delivery", time: new Date() });
}

/* ================================================================
 * LINE Bot — พิมพ์ "check" ในไลน์ แล้วบอทตอบสรุปงานค้าง (ฟรี ไม่นับโควตา)
 *
 * ===== วิธีติดตั้ง LINE Bot (ทำครั้งเดียว) =====
 * 1. เข้า https://developers.line.biz/console/ (ล็อกอินด้วย LINE)
 *    → Create provider (ตั้งชื่ออะไรก็ได้) → Create Messaging API channel
 *    → ตั้งชื่อบอท เช่น "Sugar Dispatch Bot" → สร้าง
 * 2. แท็บ Messaging API → เลื่อนล่างสุด → Channel access token → กด Issue
 *    → คัดลอก token มาวางแทน PASTE_TOKEN_HERE ข้างล่างนี้ → Save → Deploy เวอร์ชันใหม่
 * 3. แท็บ Messaging API → Webhook URL → วาง URL ของ Web App (ลงท้าย /exec)
 *    → กด Verify (ต้องขึ้น Success) → เปิดสวิตช์ "Use webhook"
 * 4. ปิดข้อความตอบกลับอัตโนมัติ: LINE Official Account Manager
 *    (manager.line.biz) → ตั้งค่า → การตอบกลับ → ปิด "ตอบกลับอัตโนมัติ"
 * 5. สแกน QR ของบอท (แท็บ Messaging API) เพิ่มเพื่อน → พิมพ์ check
 * ================================================================ */
var LINE_TOKEN = "PASTE_TOKEN_HERE";

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  if (body.events) return handleLineWebhook(body);      // LINE Bot
  if (body.action === "saveRecord") saveRecord(body.record);
  else if (body.action === "saveUsers") saveUsers(body.users);
  else if (body.action === "savePhoto") savePhoto(body);
  else if (body.action === "deleteRecord") deleteRecord(body.uid);
  else if (body.action === "logEvent") logEvent(body.event);
  return jsonOut({ ok: true });
}

/* ---------- delete record (แอดมินลบข้อมูล) ---------- */
function deleteRecord(uid) {
  var sh = getSheet(RECORDS_SHEET, REC_HEADERS);
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(uid)) sh.deleteRow(i + 1);
  }
  var ps = getSheet(PHOTOS_SHEET, ["uid", "index", "url", "fileId", "updatedAt"]);
  var pd = ps.getDataRange().getValues();
  for (var j = pd.length - 1; j >= 1; j--) {
    if (String(pd[j][0]) === String(uid)) ps.deleteRow(j + 1);
  }
}

/* ---------- LINE Bot ---------- */
function handleLineWebhook(body) {
  body.events.forEach(function (ev) {
    if (ev.type === "message" && ev.message && ev.message.type === "text" && ev.replyToken) {
      var txt = String(ev.message.text).trim().toLowerCase();
      var triggers = ["check", "เช็ค", "เช็ก", "สถานะ", "status", "ตรวจ", "งาน"];
      if (triggers.indexOf(txt) >= 0) {
        lineReply(ev.replyToken, buildStatusReport());
      }
    }
  });
  return jsonOut({ ok: true });
}

function buildStatusReport() {
  var recs = getRecords();
  var c = { draft: 0, wait_check: 0, wait_approve: 0, returned: 0, rejected: 0, done: 0 };
  recs.forEach(function (r) {
    var s = r.status === "wait" ? "wait_check" : r.status;
    if (c[s] != null) c[s]++;
  });
  var pending = c.draft + c.wait_check + c.wait_approve + c.returned + c.rejected;

  var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var tRecs = recs.filter(function (r) { return r.date === today; });
  var tTons = 0;
  tRecs.forEach(function (r) { tTons += Number(r.tons) || 0; });

  // รายการที่ค้าง (สูงสุด 8 รายการ)
  var pendList = recs.filter(function (r) {
    var s = r.status === "wait" ? "wait_check" : r.status;
    return s !== "done";
  }).slice(0, 8);
  var stName = { draft: "ร่าง", wait_check: "รอตรวจสอบ", wait_approve: "รออนุมัติ", returned: "ส่งกลับแก้ไข", rejected: "ไม่ผ่านอนุมัติ", wait: "รอตรวจสอบ" };
  var lines = pendList.map(function (r) {
    return "• " + r.plate + " — " + (stName[r.status] || r.status) +
      (r.recordedBy ? " (" + r.recordedBy.username + ")" : "");
  });

  var now = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
  return "🚛 SUGAR DISPATCH\n" +
    "━━━━━━━━━━━━━━\n" +
    "📋 ค้างดำเนินการ: " + pending + " รายการ\n" +
    "  📝 ร่าง: " + c.draft + "\n" +
    "  🔎 รอตรวจสอบ: " + c.wait_check + "\n" +
    "  ⏳ รออนุมัติ: " + c.wait_approve + "\n" +
    "  ↩️ ส่งกลับแก้ไข: " + c.returned + "\n" +
    "  ❌ ไม่ผ่านอนุมัติ: " + c.rejected + "\n" +
    "✅ สำเร็จสะสม: " + c.done + " รายการ\n" +
    "━━━━━━━━━━━━━━\n" +
    "📅 วันนี้: " + tRecs.length + " คัน / " + (Math.round(tTons * 100) / 100) + " ตัน\n" +
    (lines.length ? "━━━━━━━━━━━━━━\nรายการค้าง:\n" + lines.join("\n") + "\n" : "") +
    "🕐 " + now;
}

function lineReply(replyToken, text) {
  if (!LINE_TOKEN || LINE_TOKEN === "PASTE_TOKEN_HERE") return;
  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + LINE_TOKEN },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: "text", text: text }] }),
    muteHttpExceptions: true,
  });
}

/* รันฟังก์ชันนี้ 1 ครั้งหลังวาง token เพื่อ (1) ให้ Google ขอสิทธิ์เชื่อมต่อภายนอก
 * (2) เช็คว่า token ถูกต้อง — ดูผลใน Execution log */
function testLineSetup() {
  var resp = UrlFetchApp.fetch("https://api.line.me/v2/bot/info", {
    headers: { Authorization: "Bearer " + LINE_TOKEN },
    muteHttpExceptions: true,
  });
  Logger.log(resp.getResponseCode() === 200
    ? "✅ token ใช้ได้ — บอทชื่อ: " + JSON.parse(resp.getContentText()).displayName
    : "❌ token ไม่ถูกต้อง: " + resp.getContentText());
}

/* ---------- audit log (ใครทำอะไร เมื่อไหร่) ---------- */
var AUDIT_SHEET = "AuditLog";
function logEvent(ev) {
  if (!ev) return;
  var sh = getSheet(AUDIT_SHEET, ["ts", "user", "name", "action", "target", "note"]);
  sh.appendRow([ev.ts, ev.user || "", ev.by || "", ev.action || "", ev.target || "", ev.note || ""]);
}

/* ---------- Photos (เก็บรูปใน Google Drive โฟลเดอร์ SugarDeliveryPhotos) ---------- */
var PHOTOS_SHEET = "Photos";

function getPhotoFolder_() {
  var it = DriveApp.getFoldersByName("SugarDeliveryPhotos");
  return it.hasNext() ? it.next() : DriveApp.createFolder("SugarDeliveryPhotos");
}

function savePhoto(b) {
  var m = String(b.dataUrl).match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!m) return;
  var name = b.uid + "-" + b.index + ".jpg";
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  var folder = getPhotoFolder_();
  var it = folder.getFilesByName(name);
  while (it.hasNext()) it.next().setTrashed(true);   // replace on re-upload
  var f = folder.createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = "https://drive.google.com/thumbnail?id=" + f.getId() + "&sz=w800";

  var sh = getSheet(PHOTOS_SHEET, ["uid", "index", "url", "fileId", "updatedAt"]);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(b.uid) && Number(data[i][1]) === Number(b.index)) {
      sh.getRange(i + 1, 1, 1, 5).setValues([[b.uid, b.index, url, f.getId(), new Date()]]);
      return;
    }
  }
  sh.appendRow([b.uid, b.index, url, f.getId(), new Date()]);
}

function getPhotoMap_() {
  var sh = getSheet(PHOTOS_SHEET, ["uid", "index", "url", "fileId", "updatedAt"]);
  var rows = sh.getDataRange().getValues();
  rows.shift();
  var map = {};
  rows.filter(function (r) { return r[0]; })
    .sort(function (a, b) { return Number(a[1]) - Number(b[1]); })
    .forEach(function (r) {
      var uid = String(r[0]);
      (map[uid] = map[uid] || []).push(String(r[2]));
    });
  return map;
}

/* ---------- Users ---------- */
function getUsers() {
  var sh = getSheet(USERS_SHEET, ["id", "username", "password", "name", "role"]);
  var rows = sh.getDataRange().getValues();
  rows.shift();
  return rows
    .filter(function (r) { return r[1]; })
    .map(function (r) {
      return {
        id: Number(r[0]),
        username: String(r[1]),
        password: String(r[2]),
        name: String(r[3]),
        role: String(r[4]),
      };
    });
}

function saveUsers(users) {
  var sh = getSheet(USERS_SHEET, ["id", "username", "password", "name", "role"]);
  sh.clearContents();
  sh.appendRow(["id", "username", "password", "name", "role"]);
  users.forEach(function (u) {
    sh.appendRow([u.id, u.username, u.password, u.name, u.role]);
  });
}

/* ---------- Records ---------- */
var REC_HEADERS = ["uid", "date", "plate", "customer", "transport", "sugarType", "tons", "lots", "status", "recordedBy", "updatedAt", "json"];

function getRecords() {
  var sh = getSheet(RECORDS_SHEET, REC_HEADERS);
  var rows = sh.getDataRange().getValues();
  rows.shift();
  var photoMap = getPhotoMap_();
  var out = [];
  rows.forEach(function (r) {
    if (!r[0]) return;
    try {
      var rec = JSON.parse(r[11]);
      rec.photoUrls = photoMap[String(rec.uid)] || [];
      out.push(rec);
    } catch (err) {}
  });
  return out;
}

function saveRecord(rec) {
  var sh = getSheet(RECORDS_SHEET, REC_HEADERS);
  var row = [
    rec.uid, rec.date, rec.plate, rec.customer, rec.transport || "",
    rec.sugarType || "", rec.tons || "", (rec.lots || []).length,
    rec.status, rec.recordedBy ? rec.recordedBy.name : "",
    new Date(), JSON.stringify(rec),
  ];
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(rec.uid)) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}
