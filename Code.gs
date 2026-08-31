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

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
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
