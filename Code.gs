/*************************************************************
 * 巡檢通報系統 — Google Apps Script 後端（每專案一頁籤版）
 * ----------------------------------------------------------
 * 規則：
 *   - 同一個專案的巡檢紀錄 → 寫進「以該專案命名」的同一個頁籤
 *   - 不同專案 → 各自獨立頁籤
 *   - 未選專案 → 寫進「未分類」頁籤
 *   - 改善紀錄（單位/工法/狀態/照片）以摘要欄位併入同一列，
 *     完整多次改善歷程保留在 App 中
 *
 * 安裝：在試算表 → 擴充功能 → Apps Script 貼上本檔，
 *       或新建獨立專案亦可（已內建你的試算表 ID）。
 *       部署 → 網頁應用程式 → 執行身分:我自己 / 存取權:任何人
 *************************************************************/

// 你的試算表 ID（留空字串 '' 則改用「與本腳本綁定」的試算表）
var SHEET_ID = '1-7-VSmePtpsExX9l8LQVrK2cCjDuq5GKxrRI6CNLjbw';

var MEDIA_FOLDER_NAME = '巡檢照片';
var DEFAULT_TAB = '未分類';
var SHEET_PUMP = '抽水機台帳';
var SHEET_PCHK = '抽水機點檢';
var SHEET_PROJ = '專案清單';

var HEADERS_PROJ = ['id','專案名稱','開始日期','結束日期','備註','建立時間','更新時間','已刪除'];

var HEADERS_PUMP = ['設備編號','id','名稱','型式','驅動方式','口徑(吋)','抽水量(cms)','保管單位','維護廠商',
                    '站房/預佈位置','保養週期(天)','最後點檢','最後結果','下次保養到期','最後更新'];
var HEADERS_PCHK = ['項次','localId','設備編號','抽水機名稱','型式','驅動方式','綜合判定',
                    '異常項目','試車啟動','出水情形','運轉(分)','試車備註','點檢人員','填報單位',
                    '點檢時間','照片連結','最後更新'];

var HEADERS = ['項次','localId','所屬專案','行政區','地標/位置','緯度','經度','GPS精度(m)',
               '破損類別','破損狀況','其他描述','巡檢人員','填報單位','建立時間','狀態',
               '照片/影片連結','改善次數','致災原因','最新改善單位','最新改善工法','最新改善後狀態',
               '短期改善作為','中期改善作為','長期改善作為','改善期限','負責人',
               '最新改善日期','改善照片連結','最後更新'];

function doGet() { return json_({ ok: true, msg: '巡檢通報後端運作中' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var body = JSON.parse(e.postData.contents);
    if (body.type === 'ping')   return json_({ ok: true, msg: 'pong' });
    if (body.type === 'record') return json_({ ok: true, seq: saveRecord_(body.record) });
    if (body.type === 'pump')   { savePump_(body.pump); return json_({ ok: true }); }
    if (body.type === 'pcheck') return json_({ ok: true, seq: savePcheck_(body.check) });
    if (body.type === 'project'){ saveProject_(body.project); return json_({ ok: true }); }
    if (body.type === 'pull')   return json_({ ok: true, projects: readProjects_(), pumps: readPumps_() });
    return json_({ ok: false, error: '未知的請求類型' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ---------- 主要寫入：依專案路由到對應頁籤 ---------- */
function saveRecord_(r) {
  var ss = getSS_();
  var tabName = sanitizeTab_(r.projectName || DEFAULT_TAB);   // ← 專案名 = 頁籤名
  var sh = getOrCreateSheet_(ss, tabName, HEADERS);

  var folder = getMediaFolder_();
  var mediaLinks = saveMediaList_(folder, r.media, r.localId, '巡檢');

  // 取最後一次改善作為摘要欄位
  var imps = r.improvements || [];
  var last = imps.length ? imps[imps.length - 1] : null;
  var impLinks = last ? saveMediaList_(folder, last.media, r.localId + '_imp', '改善') : '';

  var rowIdx = findRowByLocalId_(sh, r.localId);  // 同頁籤內以 localId 防重複
  var seq;
  if (rowIdx === -1) {
    seq = nextSeq_(sh);                           // 項次：該專案頁籤內連號
    sh.appendRow(buildRow_(seq, r, mediaLinks, last, impLinks, imps.length));
  } else {
    seq = sh.getRange(rowIdx, 1).getValue() || nextSeq_(sh);
    sh.getRange(rowIdx, 1, 1, HEADERS.length)
      .setValues([buildRow_(seq, r, mediaLinks, last, impLinks, imps.length)]);
  }
  return seq;
}

function buildRow_(seq, r, mediaLinks, last, impLinks, impCount) {
  return [
    seq, r.localId, r.projectName || DEFAULT_TAB, r.area || '',
    r.landmark || '', r.lat || '', r.lng || '', r.acc || '',
    (r.categories || []).join('、'), r.condition || '', r.note || '', r.inspector || '', r.inspectorUnit || '',
    r.createdAt ? new Date(r.createdAt) : '', r.status || '',
    mediaLinks, impCount,
    last ? (last.cause || '') : '',
    last ? (last.unit || '') : '',
    last ? (last.method || '') : '',
    last ? (last.status || '') : '',
    last ? (last.shortTerm || '') : '',
    last ? (last.midTerm || '') : '',
    last ? (last.longTerm || '') : '',
    (last && last.deadline) ? last.deadline : '',
    last ? (last.assignee || '') : '',
    (last && last.date) ? new Date(last.date) : '',
    impLinks, new Date()
  ];
}

/* ---------- 媒體：base64 dataURL → Drive ---------- */
function saveMediaList_(folder, list, baseName, tag) {
  if (!list || !list.length) return '';
  var links = [];
  for (var i = 0; i < list.length; i++) {
    try {
      var m = list[i];
      var match = /^data:([^;]+);base64,(.*)$/.exec(m.data);
      if (!match) continue;
      var mime = match[1];
      var ext = mime.indexOf('video') === 0 ? '.mp4' : '.jpg';
      var blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mime,
                                   tag + '_' + baseName + '_' + (i + 1) + ext);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      links.push(file.getUrl());
    } catch (e) {}
  }
  return links.join('\n');
}

/* ---------- 抽水機台帳 ---------- */
function savePump_(p) {
  var ss = getSS_();
  var sh = getOrCreateSheet_(ss, SHEET_PUMP, HEADERS_PUMP);
  var row = [
    p.code || '', p.id, p.name || '', p.type || '', p.driver || '',
    p.sizeInch || '', p.capacityCms || '', p.unit || '', p.vendor || '', p.location || '',
    p.intervalDays || '', p.lastCheckAt ? new Date(p.lastCheckAt) : '', p.lastResult || '',
    p.nextDueAt ? new Date(p.nextDueAt) : '', new Date()
  ];
  var idx = findRowByCol_(sh, 2, p.id);   // 第2欄=id
  if (idx === -1) sh.appendRow(row);
  else sh.getRange(idx, 1, 1, HEADERS_PUMP.length).setValues([row]);
}

/* ---------- 抽水機點檢 ---------- */
function savePcheck_(c) {
  var ss = getSS_();
  var sh = getOrCreateSheet_(ss, SHEET_PCHK, HEADERS_PCHK);
  var folder = getMediaFolder_();

  // 異常項目彙整 + 收集異常照片
  var abnormal = [], media = [];
  (c.items || []).forEach(function (it) {
    if (it.val === '異常') {
      abnormal.push('• ' + it.name + (it.note ? '：' + it.note : ''));
      if (it.photo) media.push({ data: it.photo });
    }
  });
  (c.media || []).forEach(function (m) { media.push(m); });
  var links = saveMediaList_(folder, media, c.localId, '點檢');

  var idx = findRowByCol_(sh, 2, c.localId);
  var seq = (idx === -1) ? nextSeq_(sh) : (sh.getRange(idx, 1).getValue() || nextSeq_(sh));
  var row = [
    seq, c.localId, c.pumpCode || '', c.pumpName || '', c.pumpType || '', c.driver || '',
    c.result || '', abnormal.join('\n') || '（全部正常）',
    c.trialStart || '', c.trialOutflow || '', c.trialMinutes || '', c.trialNote || '',
    c.inspector || '', c.inspectorUnit || '', c.createdAt ? new Date(c.createdAt) : '',
    links, new Date()
  ];
  if (idx === -1) sh.appendRow(row);
  else sh.getRange(idx, 1, 1, HEADERS_PCHK.length).setValues([row]);
  return seq;
}

function findRowByCol_(sh, col, val) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var v = sh.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(val)) return i + 2;
  return -1;
}

/* ---------- 共用清單：專案 ---------- */
function saveProject_(p) {
  var sh = getOrCreateSheet_(getSS_(), SHEET_PROJ, HEADERS_PROJ);
  var row = [p.id, p.name || '', p.startDate || '', p.endDate || '', p.note || '',
             p.createdAt ? new Date(p.createdAt) : new Date(), new Date(), p.deleted ? '是' : ''];
  var idx = findRowByCol_(sh, 1, p.id);
  if (idx === -1) sh.appendRow(row);
  else sh.getRange(idx, 1, 1, HEADERS_PROJ.length).setValues([row]);
}

/* ---------- 拉取共用清單（專案 + 抽水機台帳） ---------- */
function readProjects_() {
  var sh = getSS_().getSheetByName(SHEET_PROJ);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS_PROJ.length).getValues();
  return rows.filter(function (r) { return r[0] && r[7] !== '是'; }).map(function (r) {
    return { id: String(r[0]), name: r[1], startDate: fmtD_(r[2]), endDate: fmtD_(r[3]),
             note: r[4], createdAt: r[5] ? new Date(r[5]).getTime() : Date.now() };
  });
}
function readPumps_() {
  var sh = getSS_().getSheetByName(SHEET_PUMP);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS_PUMP.length).getValues();
  return rows.filter(function (r) { return r[1]; }).map(function (r) {
    return { id: String(r[1]), code: r[0], name: r[2], type: r[3], driver: r[4],
             sizeInch: r[5], capacityCms: r[6], unit: r[7], vendor: r[8], location: r[9],
             intervalDays: r[10] || 30, lastCheckAt: r[11] ? new Date(r[11]).getTime() : null,
             lastResult: r[12] || '', nextDueAt: r[13] ? new Date(r[13]).getTime() : null };
  });
}
function fmtD_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    var m = ('0' + (d.getMonth() + 1)).slice(-2), day = ('0' + d.getDate()).slice(-2);
    return d.getFullYear() + '-' + m + '-' + day;
  }
  return String(d);
}

/* ---------- 工具 ---------- */
function getSS_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// 頁籤名稱合法化：移除 Google 試算表不允許的字元，限長
function sanitizeTab_(name) {
  var s = String(name).replace(/[:\\\/\?\*\[\]]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 90);
  return s || DEFAULT_TAB;
}

function getOrCreateSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#13263F').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(15, 220); // 照片連結欄加寬
  }
  return sh;
}

function findRowByLocalId_(sh, localId) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(localId)) return i + 2;
  return -1;
}

function nextSeq_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var seqs = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  seqs.forEach(function (s) { var n = parseInt(s[0], 10); if (n > max) max = n; });
  return max + 1;
}

function getMediaFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('MEDIA_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var it = DriveApp.getFoldersByName(MEDIA_FOLDER_NAME);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(MEDIA_FOLDER_NAME);
  props.setProperty('MEDIA_FOLDER_ID', folder.getId());
  return folder;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
