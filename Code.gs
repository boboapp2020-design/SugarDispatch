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
  if (action === "getLineLog") return jsonOut(getLineLog());
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
 * 5. สแกน QR ของบอท (แท็บ Messaging API) เพิ่มเพื่อน → พิมพ์ Bot Check
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
      var txt = String(ev.message.text).trim().toLowerCase().replace(/\s+/g, " ");
      var triggers = ["bot check", "botcheck", "bot เช็ค", "บอทเช็ค"];
      if (triggers.indexOf(txt) >= 0) {
        lineReplyFlex(ev.replyToken, buildStatusFlex());
      }
    }
  });
  return jsonOut({ ok: true });
}

function statusCounts_() {
  var recs = getRecords();
  var c = { draft: 0, wait_check: 0, wait_approve: 0, returned: 0, rejected: 0, done: 0 };
  recs.forEach(function (r) {
    var s = r.status === "wait" ? "wait_check" : r.status;
    if (c[s] != null) c[s]++;
  });
  var today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  var tRecs = recs.filter(function (r) { return r.date === today; });
  var tTons = 0;
  tRecs.forEach(function (r) { tTons += Number(r.tons) || 0; });
  var stName = { draft: "ร่าง", wait_check: "รอตรวจสอบ", wait_approve: "รออนุมัติ", returned: "ส่งกลับแก้ไข", rejected: "ไม่ผ่านอนุมัติ", wait: "รอตรวจสอบ" };
  var pendList = recs.filter(function (r) { return r.status !== "done"; }).slice(0, 5)
    .map(function (r) {
      return "• " + r.plate + " — " + (stName[r.status] || r.status) +
        (r.recordedBy ? " (" + r.recordedBy.username + ")" : "");
    });
  return {
    c: c,
    pending: c.draft + c.wait_check + c.wait_approve + c.returned + c.rejected,
    tCount: tRecs.length,
    tTons: Math.round(tTons * 100) / 100,
    pendList: pendList,
  };
}

function buildStatusFlex() {
  var s = statusCounts_();
  var now = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
  function row(label, val, color) {
    return { type: "box", layout: "horizontal", margin: "sm", contents: [
      { type: "text", text: label, size: "sm", color: "#5C6B84", flex: 6 },
      { type: "text", text: String(val), size: "sm", color: color || "#1C2A42", align: "end", weight: "bold", flex: 4 },
    ]};
  }
  var bodyRows = [
    { type: "box", layout: "horizontal", contents: [
      { type: "text", text: "📋 ค้างดำเนินการ", size: "md", weight: "bold", color: "#1C2A42", flex: 6 },
      { type: "text", text: s.pending + " รายการ", size: "md", weight: "bold",
        color: s.pending > 0 ? "#C4453A" : "#2E8B57", align: "end", flex: 4 },
    ]},
    { type: "separator", margin: "md" },
    row("📝 ร่าง", s.c.draft),
    row("🔎 รอตรวจสอบ", s.c.wait_check, "#2C6E9B"),
    row("⏳ รออนุมัติ", s.c.wait_approve, "#C9992E"),
    row("↩️ ส่งกลับแก้ไข", s.c.returned, "#C4453A"),
    row("❌ ไม่ผ่านอนุมัติ", s.c.rejected, "#C4453A"),
    { type: "separator", margin: "md" },
    row("✅ สำเร็จสะสม", s.c.done + " รายการ", "#2E8B57"),
    row("📅 วันนี้", s.tCount + " คัน / " + s.tTons + " ตัน"),
  ];
  if (s.pendList.length) {
    bodyRows.push({ type: "separator", margin: "md" });
    bodyRows.push({ type: "text", text: "รายการค้างล่าสุด", size: "xs", weight: "bold", color: "#5C6B84", margin: "md" });
    s.pendList.forEach(function (line) {
      bodyRows.push({ type: "text", text: line, size: "xs", color: "#8DA0B5", wrap: true, margin: "sm" });
    });
  }
  return {
    type: "bubble",
    header: { type: "box", layout: "vertical", backgroundColor: "#16294B", paddingAll: "16px", contents: [
      { type: "text", text: "🚛 SUGAR DISPATCH", color: "#FFFFFF", weight: "bold", size: "lg" },
      { type: "text", text: "สรุปสถานะงาน · " + now, color: "#C9992E", size: "xs", margin: "sm" },
    ]},
    body: { type: "box", layout: "vertical", paddingAll: "16px", contents: bodyRows },
    footer: { type: "box", layout: "vertical", paddingAll: "12px", contents: [
      { type: "button", style: "primary", color: "#1F3A68", height: "sm",
        action: { type: "uri", label: "เปิดแอป SUGAR DISPATCH", uri: "https://boboapp2020-design.github.io/SugarDispatch/" } },
    ]},
  };
}

function lineReplyFlex(replyToken, bubble) {
  if (!LINE_TOKEN || LINE_TOKEN === "PASTE_TOKEN_HERE") {
    lineLog_("SKIP", "LINE_TOKEN ยังไม่ได้ใส่");
    return;
  }
  var resp = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + LINE_TOKEN },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: "flex", altText: "🚛 SUGAR DISPATCH — สรุปสถานะงาน", contents: bubble }],
    }),
    muteHttpExceptions: true,
  });
  lineLog_("FLEX " + resp.getResponseCode(), resp.getContentText().slice(0, 400));
  if (resp.getResponseCode() !== 200) {
    // การ์ดถูกปฏิเสธ — ส่งแบบข้อความธรรมดาแทน (replyToken ยังใช้ได้)
    var resp2 = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/reply", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + LINE_TOKEN },
      payload: JSON.stringify({
        replyToken: replyToken,
        messages: [{ type: "text", text: buildStatusText_() }],
      }),
      muteHttpExceptions: true,
    });
    lineLog_("TEXT-FALLBACK " + resp2.getResponseCode(), resp2.getContentText().slice(0, 400));
  }
}

function buildStatusText_() {
  var s = statusCounts_();
  var now = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
  return "🚛 SUGAR DISPATCH\n📋 ค้างดำเนินการ: " + s.pending + " รายการ\n" +
    "📝 ร่าง: " + s.c.draft + " | 🔎 รอตรวจ: " + s.c.wait_check + "\n" +
    "⏳ รออนุมัติ: " + s.c.wait_approve + " | ↩️ ส่งกลับ: " + (s.c.returned + s.c.rejected) + "\n" +
    "✅ สำเร็จ: " + s.c.done + "\n📅 วันนี้: " + s.tCount + " คัน / " + s.tTons + " ตัน\n" +
    (s.pendList.length ? s.pendList.join("\n") + "\n" : "") + "🕐 " + now;
}

var LINELOG_SHEET = "LineLog";
function lineLog_(status, detail) {
  try {
    var sh = getSheet(LINELOG_SHEET, ["ts", "status", "detail"]);
    sh.appendRow([new Date(), status, detail]);
  } catch (e) {}
}
function getLineLog() {
  var sh = getSheet(LINELOG_SHEET, ["ts", "status", "detail"]);
  var rows = sh.getDataRange().getValues();
  rows.shift();
  return rows.slice(-10).map(function (r) { return { ts: r[0], status: r[1], detail: r[2] }; });
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
