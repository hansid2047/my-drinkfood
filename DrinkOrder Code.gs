// ═══════════════════════════════════════════════════════════════════
//  辦公室飲料或餐點揪團訂購系統 — Google Apps Script 後端
//
//  設定步驟：
//  1. 開啟 Google 試算表 > 擴充功能 > Apps Script
//  2. 貼上此程式碼，儲存
//  3. 執行 initSheets() 初始化工作表
//  4. 部署 > 新增部署 > 網路應用程式
//     - 以身分執行：我（你的帳號）
//     - 誰可以存取：所有人
//  5. 複製部署 URL 貼到 index.html 的 GAS_URL
// ═══════════════════════════════════════════════════════════════════

const SHEET_ORDERS  = "訂單紀錄";
const SHEET_SESSIONS = "揪團場次";

// ─── 初始化 ──────────────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(SHEET_ORDERS)) {
    const s = ss.insertSheet(SHEET_ORDERS);
    s.appendRow(["場次ID","場次名稱","訂購者姓名","飲料或餐點名稱","大小","冰塊","甜度","備註","金額(元)","訂購時間"]);
    styleHeader(s, 10, "#2D6A4F");
    s.setFrozenRows(1);
    [80,160,90,160,70,80,80,160,80,160].forEach((w,i)=>s.setColumnWidth(i+1,w));
  }

  if (!ss.getSheetByName(SHEET_SESSIONS)) {
    const s = ss.insertSheet(SHEET_SESSIONS);
    s.appendRow(["場次ID","場次名稱","主辦人","店家","截止時間","狀態","建立時間","菜單連結","菜單圖片"]);
    styleHeader(s, 9, "#1B4332");
    s.setFrozenRows(1);
    [100,180,90,130,140,70,160,260,260].forEach((w,i)=>s.setColumnWidth(i+1,w));
  }
}

function styleHeader(sheet, cols, color) {
  sheet.getRange(1,1,1,cols).setFontWeight("bold").setBackground(color).setFontColor("#ffffff");
}

// ─── JSONP 回應 ───────────────────────────────────────────────────
function doGet(e) {
  const p  = e.parameter || {};
  const cb = p.callback  || "";
  let result;
  try {
    const action = p.action || "";
    if      (action === "getSessions") result = getSessions();
    else if (action === "getOrders")   result = getOrders(p.sessionId);
    else if (action === "addOrder")    result = addOrder(p);
    else if (action === "addSession")  result = addSession(p);
    else if (action === "closeSession")result = closeSession(p.sessionId);
    else result = { ok:false, error:"unknown action" };
  } catch(err) {
    result = { ok:false, error: err.message };
  }

  const json = JSON.stringify(result);
  if (cb) {
    return ContentService.createTextOutput(`${cb}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ─── 場次 ─────────────────────────────────────────────────────────
function getSessions() {
  const s    = getSheet(SHEET_SESSIONS);
  const rows = s.getDataRange().getValues();
  if (rows.length <= 1) return { ok:true, data:[] };
  return {
    ok: true,
    data: rows.slice(1).map(r => ({
      sessionId: r[0], name: r[1], host: r[2], store: r[3],
      deadline: fmtDT(r[4]), status: r[5], createdAt: fmtDT(r[6]),
      menuUrl: r[7]||"", menuImg: r[8]||"",
    }))
  };
}

function addSession(p) {
  const s      = getSheet(SHEET_SESSIONS);
  const id     = "S" + new Date().getTime();
  const now    = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  s.appendRow([id, p.name||"", p.host||"", p.store||"", p.deadline||"", "開放中", now, p.menuUrl||"", p.menuImg||""]);
  return { ok:true, sessionId: id };
}

function closeSession(sessionId) {
  const s    = getSheet(SHEET_SESSIONS);
  const rows = s.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if (String(rows[i][0]) === String(sessionId)) {
      s.getRange(i+1, 6).setValue("已截止");
      s.getRange(i+1, 1, 1, 7).setBackground("#F3F4F6");
      return { ok:true };
    }
  }
  return { ok:false, error:"not found" };
}

// ─── 訂單 ─────────────────────────────────────────────────────────
function getOrders(sessionId) {
  const s    = getSheet(SHEET_ORDERS);
  const rows = s.getDataRange().getValues();
  if (rows.length <= 1) return { ok:true, data:[] };
  let data = rows.slice(1).map(r => ({
    sessionId: r[0], sessionName: r[1], name: r[2], drink: r[3],
    size: r[4], ice: r[5], sugar: r[6], note: r[7], price: r[8], orderedAt: r[9],
  }));
  if (sessionId) data = data.filter(o => String(o.sessionId) === String(sessionId));
  return { ok:true, data };
}

function addOrder(p) {
  const s   = getSheet(SHEET_ORDERS);
  const now = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  s.appendRow([
    p.sessionId   || "",
    p.sessionName || "",
    p.name        || "",
    p.drink       || "",
    p.size        || "",
    p.ice         || "",
    p.sugar       || "",
    p.note        || "",
    p.price       || "0",
    now,
  ]);
  // 交替底色
  const row   = s.getLastRow();
  s.getRange(row,1,1,10).setBackground(row%2===0?"#F0FDF4":"#FFFFFF");
  return { ok:true };
}

// ─── 工具 ─────────────────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s    = ss.getSheetByName(name);
  if (!s) { initSheets(); s = ss.getSheetByName(name); }
  return s;
}

function fmtDT(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd HH:mm");
  return String(val);
}
