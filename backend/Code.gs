/**
 * 家庭二手庫 — Google Apps Script 後台
 *
 * 第一次使用：
 *   1. 改下面兩個設定（EDITORS、CLIENT_ID）
 *   2. 執行 setup()   → 建立分頁、照片資料夾
 *   3. 部署 → 新增部署作業 → 網頁應用程式（執行身分：我；存取權：所有人）
 *
 * 只有 EDITORS 裡的 Google 帳號能讀寫，沒有家庭密碼／唯讀模式。
 */

// ===== 設定 =====
const EDITORS = ['你的Gmail@gmail.com', '太太的Gmail@gmail.com'];   // 可以使用的 Google 帳號
const CLIENT_ID = 'xxxxxxxx.apps.googleusercontent.com';            // 沿用家庭防災庫的 OAuth 用戶端 ID 即可
const SESSION_DAYS = 60;                                             // 登入後幾天內不用重新登入

// ===== 分頁結構（欄位順序就是試算表的欄位順序）=====
const TABLES = {
  Members:  ['id', 'name'],
  Items:    ['id', 'cat', 'sub', 'name', 'brand', 'model', 'owner', 'status',
             'buyDate', 'listPrice', 'buyPrice', 'discNote', 'channel', 'origin', 'receipt', 'warrantyEnd', 'warrantyNote',
             'pack', 'packNote', 'cond', 'condNote', 'size', 'age', 'mfgDate', 'accident',
             'buyWhy', 'sellWhy', 'desc', 'askPrice', 'priceNote', 'listedAt',
             'soldPrice', 'soldDate', 'soldFee', 'soldVia', 'postFields', 'photos', 'updatedAt', 'updatedBy'],
  Wishlist: ['id', 'cat', 'name', 'owner', 'budget', 'refPrice', 'priority', 'via', 'buyWhy', 'note', 'link', 'createdAt'],
  Settings: ['id', 'value'],
  Photos:   ['id', 'thumb']
};
const FOLDER_NAME = '家庭二手庫照片';

// ===== 初始化 =====
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABLES).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const cols = TABLES[name];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // 全部存成純文字，避免日期、價格被自動轉格式
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  });
  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  const members = ss.getSheetByName('Members');
  if (members.getLastRow() < 2) {
    members.getRange(2, 1, 5, 2).setValues([
      ['dad', '爸爸'], ['mom', '媽媽'], ['son1', '哥哥'], ['son2', '弟弟'], ['girl', '妹妹']
    ]);
  }
  const settings = ss.getSheetByName('Settings');
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, 7, 2).setValues([
      ['meet', ''], ['ship', '7-11 賣貨便'], ['pay', '面交付現／轉帳'], ['priceNote', ''],
      ['footer', '有興趣請私訊，謝謝！'],
      ['postFields', 'brand,cond,pack,buyDate,channel,list,warr,spec,sellWhy,desc,deal'], ['emoji', '1']
    ]);
  }
  getFolder_();
  Logger.log('完成。Members 分頁的稱呼可以改；Settings 的 meet 填面交地點（寫到行政區即可）。');
}

/** 清掉所有登入（例如手機遺失時） */
function logoutEveryone() {
  const p = PropertiesService.getScriptProperties();
  Object.keys(p.getProperties()).filter(k => k.indexOf('sess_') === 0).forEach(k => p.deleteProperty(k));
  Logger.log('所有裝置都需要重新登入。');
}

// ===== 網頁入口 =====
function doGet() {
  return json_({ ok: true, app: '家庭二手庫' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '格式錯誤' }); }
  try {
    switch (req.action) {
      case 'loginGoogle': return json_(loginGoogle_(req.credential));
      case 'logout':      endSession_(req.session); return json_({ ok: true });
    }
    const sess = getSession_(req.session);
    if (!sess) return json_({ ok: false, error: 'AUTH', message: '登入已過期，請重新登入' });
    switch (req.action) {
      case 'load':   return json_(Object.assign({ ok: true, role: sess.role, who: sess.name }, loadAll_()));
      case 'photo':  return json_({ ok: true, data: getPhoto_(req.id) });
      case 'sync':   return json_(sync_(req.changes, sess));
      case 'upload': return json_(upload_(req.full, req.thumb, req.name));
    }
    return json_({ ok: false, error: '未知的動作' });
  } catch (err) {
    return json_({ ok: false, error: 'SERVER', message: String(err && err.message || err) });
  }
}

// ===== 登入 =====
function loginGoogle_(credential) {
  if (!credential) return { ok: false, error: '缺少登入資訊' };
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return { ok: false, error: 'Google 登入驗證失敗' };
  const t = JSON.parse(res.getContentText());
  if (t.aud !== CLIENT_ID) return { ok: false, error: '用戶端 ID 不符，請檢查 CLIENT_ID 設定' };
  if (String(t.email_verified) !== 'true') return { ok: false, error: 'Email 未驗證' };
  const email = String(t.email || '').toLowerCase();
  if (EDITORS.map(x => x.toLowerCase()).indexOf(email) < 0) return { ok: false, error: '這個帳號沒有使用權限：' + email };
  return Object.assign({ ok: true }, newSession_('editor', t.name || email, email));
}

function newSession_(role, name, email) {
  const p = PropertiesService.getScriptProperties();
  const now = Date.now();
  Object.entries(p.getProperties()).forEach(([k, v]) => {
    if (k.indexOf('sess_') === 0) { try { if (JSON.parse(v).exp < now) p.deleteProperty(k); } catch (e) { p.deleteProperty(k); } }
  });
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  p.setProperty('sess_' + sha256_(token), JSON.stringify({ role, name, email, exp: now + SESSION_DAYS * 864e5 }));
  return { session: token, role, name };
}

function getSession_(token) {
  if (!token) return null;
  const v = PropertiesService.getScriptProperties().getProperty('sess_' + sha256_(token));
  if (!v) return null;
  const s = JSON.parse(v);
  return s.exp > Date.now() && s.role === 'editor' ? s : null;
}

function endSession_(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('sess_' + sha256_(token));
}

// ===== 讀寫 =====
function loadAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = { tables: {} };
  Object.keys(TABLES).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { out.tables[name] = []; return; }
    const last = sh.getLastRow();
    if (last < 2) { out.tables[name] = []; return; }
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    const rows = sh.getRange(2, 1, last - 1, head.length).getDisplayValues();
    out.tables[name] = rows.filter(r => r[0] !== '').map(r => {
      const o = {}; head.forEach((h, i) => { if (h) o[h] = r[i]; }); return o;
    });
  });
  out.serverTime = new Date().toISOString();
  return out;
}

function sync_(changes, sess) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(changes || {}).forEach(name => {
      if (!TABLES[name] || name === 'Photos') return;
      const sh = ss.getSheetByName(name);
      const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
      const last = sh.getLastRow();
      const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]) : [];
      const ch = changes[name];
      (ch.upsert || []).forEach(row => {
        if (name === 'Items') { row.updatedAt = new Date().toISOString(); row.updatedBy = sess.name; }
        const values = [head.map(h => row[h] == null ? '' : String(row[h]))];
        const at = ids.indexOf(String(row.id));
        if (at >= 0) sh.getRange(at + 2, 1, 1, head.length).setValues(values);
        else { sh.getRange(sh.getLastRow() + 1, 1, 1, head.length).setNumberFormat('@').setValues(values); ids.push(String(row.id)); }
      });
      // 由下往上刪，避免列號位移
      (ch.del || []).map(id => ids.indexOf(String(id))).filter(i => i >= 0).sort((a, b) => b - a)
        .forEach(i => { sh.deleteRow(i + 2); ids.splice(i, 1); });
    });
    return { ok: true, serverTime: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}

// ===== 照片 =====
// 照片資料夾建在試算表所在的資料夾裡（例如 Drive「家庭二手庫」）
function getFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const parents = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const found = parent.getFoldersByName(FOLDER_NAME);
  const f = found.hasNext() ? found.next() : parent.createFolder(FOLDER_NAME);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

function upload_(full, thumb, name) {
  const m = /^data:(image\/\w+);base64,(.+)$/.exec(full || '');
  if (!m) return { ok: false, error: '照片格式錯誤' };
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (name || 'photo') + '.jpg');
  const file = getFolder_().createFile(blob);
  const id = file.getId();
  if (thumb) {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Photos');
    sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[id, thumb]]);
  }
  return { ok: true, id };
}

function getPhoto_(id) {
  const file = DriveApp.getFileById(id);
  // 只讀照片資料夾裡的檔案
  const parents = file.getParents(), folderId = getFolder_().getId();
  let inside = false;
  while (parents.hasNext()) if (parents.next().getId() === folderId) inside = true;
  if (!inside) throw new Error('找不到這張照片');
  const blob = file.getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

// ===== 工具 =====
function sha256_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
