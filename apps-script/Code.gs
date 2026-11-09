/**
 * 活動報到系統(Google 試算表 + Apps Script)
 *
 * 分頁:
 *   名單  — 貴賓/民眾基本資料(代號、姓名、單位、類別、桌次、報到禮)
 *   回函  — Google 表單回覆(代號、是否出席、出席人數、各人葷素)
 *   票券  — 由「回函」自動產生,每人一列(vip001、vip001-1 ...),含報到/領禮狀態與 QR
 *   設定  — 報到網址、工作人員密碼
 *   紀錄  — 每次報到/領禮/撤銷的流水帳
 */

var SHEET = { GUESTS: '名單', RSVP: '回函', TICKETS: '票券', SETTINGS: '設定', LOG: '紀錄' };
var MAX_PARTY = 4; // 本人 + 最多 3 位攜伴
var COLS = 14;
var TICKET_HEADERS = ['票號', '代號', '姓名', '單位', '類別', '葷素', '桌次', '報到禮',
  '報到時間', '報到處', '領禮時間', '領禮處', 'QR網址', 'QR圖'];

/* ───────────── 試算表選單 ───────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📋 報到系統')
    .addItem('① 建立分頁(第一次使用)', 'setup')
    .addItem('② 儲存報到網址(部署後)', 'menuSaveUrl')
    .addItem('③ 為名單產生兩種 QR(回函 / 報到)', 'menuFillQr')
    .addItem('④ 由回函產生票券(可重複按)', 'menuGenerate')
    .addItem('⑤ 查看報到網址與密碼', 'menuInfo')
    .addToUi();
}

function menuGenerate() {
  var msg = generateTickets_();
  SpreadsheetApp.getUi().alert(msg);
}

function menuSaveUrl() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('還沒有部署網頁應用程式。請先照教學「部署」,再回來按這個選單。'); return; }
  ss_().getSheetByName(SHEET.SETTINGS).getRange('B2').setValue(url);
  ui.alert('已儲存報到網址:\n' + url +
    (/\/exec$/.test(url) ? '' : '\n\n⚠ 網址結尾不是 /exec,請改貼「部署」後給你的網頁應用程式網址到「設定」分頁 B2。'));
}

/** 為「名單」每位貴賓產生:回函連結、回函 QR、報到 QR(報到 QR 不需等回函) */
function menuFillQr() {
  var ui = SpreadsheetApp.getUi();
  var sh = ss_().getSheetByName(SHEET.GUESTS);
  var last = sh.getLastRow();
  if (last < 2) { ui.alert('「名單」分頁還沒有資料。'); return; }
  sh.getRange(1, 8, 1, 5).setValues([['回函連結', '回函QR圖片網址', '報到QR圖片網址(純代號)', '回函QR圖', '報到QR圖']])
    .setFontWeight('bold').setBackground('#e8eaed');
  var codes = sh.getRange(2, 1, last - 1, 1).getValues();
  var S4 = "'" + SHEET.SETTINGS + "'!$B$4";
  var Q = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=';
  var rows = [];
  codes.forEach(function (c, i) {
    var r = i + 2;
    if (String(c[0]).trim() === '') { rows.push(['', '', '', '', '']); return; }
    rows.push([
      '=IF(' + S4 + '="","",SUBSTITUTE(' + S4 + ',"{CODE}",$A' + r + '))',
      '=IF($H' + r + '="","","' + Q + '"&ENCODEURL($H' + r + '))',
      '=IF($A' + r + '="","","' + Q + '"&ENCODEURL(UPPER($A' + r + ')))', // 報到 QR:內容就是代號本身,一律大寫(系統不分大小寫)
      '=IF($I' + r + '="","",IMAGE($I' + r + '))',
      '=IF($J' + r + '="","",IMAGE($J' + r + '))'
    ]);
  });
  sh.getRange(2, 8, rows.length, 5).setValues(rows);
  ui.alert('已為 ' + codes.filter(function (c) { return String(c[0]).trim(); }).length + ' 位貴賓產生 QR。\n' +
    (getSetting_('回函表單預填網址') ? '' : '\n⚠ 「設定」分頁的「回函表單預填網址」還是空的,回函連結/回函 QR 會是空白。'));
}

function menuInfo() {
  var base = getSetting_('報到網址');
  var k = boardKey_();
  SpreadsheetApp.getUi().alert(
    '【報到處(工讀生用)】\n' + base + '\n密碼:' + getSetting_('工作人員密碼') +
    '\n\n【大螢幕歡迎看板】\n' + base + '?page=board&k=' + k +
    '\n\n【貴賓查詢頁 / 電子報到憑證】\n' + base + '?page=me');
}

/* ───────────── 初始化 ───────────── */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SSID', ss.getId());
  ensureSheet_(ss, SHEET.GUESTS, ['代號', '姓名', '單位', '類別', '桌次', '報到禮', '備註']);
  ensureSheet_(ss, SHEET.RSVP, ['時間戳記', '代號', '是否出席', '出席人數', '本人葷素', '攜伴1葷素', '攜伴2葷素', '攜伴3葷素']);
  var t = ensureSheet_(ss, SHEET.TICKETS, TICKET_HEADERS);
  t.getRange('I:I').setNumberFormat('MM/dd HH:mm:ss');
  t.getRange('K:K').setNumberFormat('MM/dd HH:mm:ss');
  var s = ensureSheet_(ss, SHEET.SETTINGS, ['項目', '值']);
  if (s.getLastRow() < 3) {
    s.getRange(2, 1, 2, 2).setValues([['報到網址', ''], ['工作人員密碼', String(Math.floor(100000 + Math.random() * 900000))]]);
  }
  ensureSheet_(ss, SHEET.LOG, ['時間', '報到處', '動作', '票號', '結果']);
  ['回函表單預填網址', '回函來源試算表'].forEach(function (nm) {
    var v = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(function (r) { return String(r[0]).trim(); });
    if (v.indexOf(nm) < 0) s.getRange(s.getLastRow() + 1, 1, 1, 2).setValues([[nm, '']]);
  });
  SpreadsheetApp.getUi().alert('分頁已準備好。請在「設定」分頁查看工作人員密碼(可自行修改)。');
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ───────────── 共用小工具 ───────────── */

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SSID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase(); }

function getSetting_(name) {
  var sh = ss_().getSheetByName(SHEET.SETTINGS);
  if (!sh || sh.getLastRow() < 1) return '';
  var v = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === name) return String(v[i][1]).trim();
  return '';
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function fmtT_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm:ss');
  return v == null ? '' : String(v);
}

function extractCode_(q) {
  q = String(q == null ? '' : q).trim();
  var m = q.match(/[?&]t=([^&\s#]+)/);
  if (m) { try { return decodeURIComponent(m[1]); } catch (e) { return m[1]; } }
  return q;
}

/* ───────────── 由回函產生票券 ───────────── */

function nameKey_(s) {
  return String(s == null ? '' : s).normalize('NFKC').replace(/[\s　]/g, '').toLowerCase();
}

function readGuests_() {
  var sh = ss_().getSheetByName(SHEET.GUESTS);
  var out = { byCode: {}, byName: {} };
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
    var c = String(r[0]).trim();
    if (!c) return;
    var g = { code: c, name: String(r[1]).trim(), org: String(r[2]).trim(), type: String(r[3]).trim(), table: String(r[4]).trim() };
    out.byCode[norm_(c)] = g;
    var nk = nameKey_(g.name);
    if (nk) (out.byName[nk] = out.byName[nk] || []).push(g);
  });
  return out;
}

/** 回函資料來源:預設是本試算表的「回函」分頁;若「設定」填了「回函來源試算表」,則直接讀那份(表單回覆)的第一個分頁 */
function rsvpSheet_() {
  var src = getSetting_('回函來源試算表');
  if (!src) return ss_().getSheetByName(SHEET.RSVP);
  var m = src.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return SpreadsheetApp.openById(m ? m[1] : src).getSheets()[0];
}

/**
 * 讀回函。支援兩種表單:
 *  (A) 有「代號」欄:用代號對應名單
 *  (B) 沒有代號:用「姓名」(必要時加「服務單位」)對應名單
 * 人數欄:「攜伴人數」(不含本人,會自動 +1) 或「出席人數」(含本人)
 * 葷素欄:「本人…」「攜伴1…」「攜伴2…」「攜伴3…」(有幾個讀幾個)
 */
function readRsvps_(guests) {
  var sh = rsvpSheet_();
  var out = { map: {}, unmatched: [] };
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getDataRange().getValues();
  var head = vals[0].map(String);
  var col = function (pred) { for (var i = 0; i < head.length; i++) if (pred(head[i])) return i; return -1; };
  var has = function (k) { return function (h) { return h.indexOf(k) >= 0; }; };
  var cCode = col(has('代號')), cName = col(has('姓名')), cOrg = col(has('單位')), cAtt = col(has('是否出席'));
  var cComp = col(function (h) { return h.indexOf('攜伴') >= 0 && h.indexOf('人數') >= 0; });
  var cTotal = col(function (h) { return h.indexOf('人數') >= 0 && h.indexOf('攜伴') < 0; });
  var cDiet = [col(has('本人')), col(has('攜伴1')), col(has('攜伴2')), col(has('攜伴3'))];
  if ((cCode < 0 && cName < 0) || (cComp < 0 && cTotal < 0)) {
    throw new Error('回函的表頭需要有「姓名」(或「代號」),以及「攜伴人數」(或「出席人數」)。目前表頭:' + head.join('、'));
  }
  var DECLINE = /(否|不出席|不克|無法|婉拒|不參加)/;
  for (var i = 1; i < vals.length; i++) {
    var row = vals[i], g = null, label = '';
    var declined = cAtt >= 0 && DECLINE.test(String(row[cAtt]));
    if (cCode >= 0 && String(row[cCode]).trim()) {
      label = String(row[cCode]).trim();
      g = guests.byCode[norm_(label)] || null;
    }
    if (!g && cName >= 0 && String(row[cName]).trim()) {
      label = String(row[cName]).trim();
      var cands = guests.byName[nameKey_(label)] || [];
      if (cands.length > 1 && cOrg >= 0) {
        var org = nameKey_(row[cOrg]);
        var narrowed = cands.filter(function (c) {
          var o = nameKey_(c.org);
          return org && o && (o.indexOf(org) >= 0 || org.indexOf(o) >= 0);
        });
        if (narrowed.length) cands = narrowed;
      }
      if (cands.length === 1) g = cands[0];
      else if (cands.length > 1) { out.unmatched.push(label + '(名單有多位同名,請手動處理)'); continue; }
    }
    if (!g) { if (label && !declined) out.unmatched.push(label); continue; }
    var n;
    if (cComp >= 0) { var m = String(row[cComp]).match(/\d+/); n = 1 + (m ? parseInt(m[0], 10) : 0); }
    else { var t = String(row[cTotal]).match(/\d+/); n = t ? parseInt(t[0], 10) : 1; }
    if (declined) n = 0;
    n = Math.max(0, Math.min(n, MAX_PARTY));
    out.map[norm_(g.code)] = { // 同一人回函多次,以最後一筆為準
      code: g.code, n: n,
      diets: cDiet.map(function (c) { return c >= 0 ? String(row[c]).trim() : ''; })
    };
  }
  return out;
}

function ticketRow_(code, party, idx, r, diet) {
  var G = "'" + SHEET.GUESTS + "'!$A:$G";
  var lk = function (c) { return '=IFERROR(VLOOKUP($B' + r + ',' + G + ',' + c + ',FALSE)&"","")'; };
  var name = idx === 0 ? lk(2)
    : '=IFERROR(VLOOKUP($B' + r + ',' + G + ',2,FALSE)&" 的攜伴' + idx + '","")';
  var qrUrl = '=IF($A' + r + '="","","https://api.qrserver.com/v1/create-qr-code/?size=300x300&data="&ENCODEURL(UPPER($A' + r + ')))';
  var qrImg = '=IF($M' + r + '="","",IMAGE($M' + r + '))';
  return [code, party, name, lk(3), lk(4), diet, lk(5), lk(6), '', '', '', '', qrUrl, qrImg];
}

function generateTickets_() {
  return withLock_(function () {
    var ss = ss_();
    var guests = readGuests_();
    var rs = readRsvps_(guests);
    var rsvps = rs.map;
    var sh = ss.getSheetByName(SHEET.TICKETS);
    var existing = {};
    var last = Math.max(sh.getLastRow(), 1);
    if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) {
      if (r[0] !== '') existing[norm_(r[0])] = i + 2;
    });
    var rows = [], added = 0, updated = 0, warns = [];
    Object.keys(rsvps).forEach(function (k) {
      var r = rsvps[k];
      var party = r.code;
      for (var i = 0; i < r.n; i++) {
        var code = i === 0 ? party : party + '-' + i;
        var diet = r.diets[i] || '';
        var row = existing[norm_(code)];
        if (row) {
          if (diet) sh.getRange(row, 6).setValue(diet);
          updated++;
        } else {
          rows.push(ticketRow_(code, party, i, last + 1 + rows.length, diet));
          existing[norm_(code)] = -1;
          added++;
        }
      }
      // 人數減少時提醒,不自動刪除
      var p = norm_(party);
      Object.keys(existing).forEach(function (ek) {
        var idx = ek === p ? 0 : (ek.indexOf(p + '-') === 0 ? parseInt(ek.slice(p.length + 1), 10) : -1);
        if (idx >= r.n && existing[ek] > 0) warns.push('票 ' + ek + ' 超出回函人數(' + r.n + '人),請確認是否要手動刪除該列');
      });
    });
    // 名單上還沒回函的貴賓:先建立本人一張票,手持報到 QR 到場也能報到
    var noRsvp = 0;
    Object.keys(guests.byCode).forEach(function (k) {
      if (rsvps[k] || existing[k]) return;
      var gc = guests.byCode[k].code;
      rows.push(ticketRow_(gc, gc, 0, last + 1 + rows.length, ''));
      existing[k] = -1; added++; noRsvp++;
    });
    if (rows.length) sh.getRange(last + 1, 1, rows.length, COLS).setValues(rows);
    SpreadsheetApp.flush();
    rs.unmatched.slice().reverse().forEach(function (nm) {
      warns.unshift('回函中的「' + nm + '」在名單找不到對應的人,請確認姓名是否一致(需手動處理,否則不會產生票券)');
    });
    if (noRsvp) warns.push('有 ' + noRsvp + ' 位貴賓尚未回函,已先各建立「本人」一張票(葷素顯示未填)');
    return '新增票券 ' + added + ' 張,既有票券 ' + updated + ' 張(已更新葷素)。' +
      (warns.length ? '\n\n⚠ 注意(共 ' + warns.length + ' 項):\n' + warns.slice(0, 20).join('\n') : '');
  });
}

/* ───────────── 網頁 ───────────── */
/*
 * 速度設計:報到網頁一次載入整份名單(snapshot),掃描/搜尋都在手機本機完成;
 * 報到、領禮、撤銷在手機上先生效,再排隊送到這裡寫入試算表。
 * 這裡的工作:提供快照、依序寫入(加鎖)、擋掉重複與重送。
 */

/**
 * 同一個網址,依 ?page= 顯示不同畫面:
 *   (無)          報到處(工讀生用,要密碼)
 *   ?page=board   大螢幕歡迎看板(要看板金鑰 &k=…)
 *   ?page=me      貴賓查詢頁 / 電子報到憑證(輸入代號與姓名)
 */
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || '';
  var file = page === 'board' ? 'board' : (page === 'me' ? 'guest' : 'index');
  var title = { board: '貴賓報到看板', guest: '貴賓報到憑證', index: '活動報到系統' }[file];
  return HtmlService.createHtmlOutputFromFile(file).setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 網頁唯一入口:所有操作都要帶工作人員密碼 */
function api(pin, action, arg) {
  var expected = getSetting_('工作人員密碼');
  if (!expected) return { ok: false, auth: false, error: '尚未設定工作人員密碼(見「設定」分頁)' };
  if (String(pin == null ? '' : pin).trim() !== expected) return { ok: false, auth: false, error: '密碼錯誤' };
  arg = arg || {};
  switch (action) {
    case 'ping': return { ok: true };
    case 'snapshot': return snapshot_();
    case 'roster': return roster_();
    case 'process': return withLock_(function () { return process_(arg); });
    case 'undo': return withLock_(function () { return undo_(arg); });
    case 'addMember': return withLock_(function () { return addMember_(arg); });
    case 'setDiet': return withLock_(function () { return setDiet_(arg); });
    default: return { ok: false, error: '未知操作' };
  }
}

/* ───────────── 公開頁面(不使用工作人員密碼,只回傳必要欄位) ───────────── */

function boardKey_() {
  var k = getSetting_('大螢幕看板金鑰');
  if (k) return k;
  k = '';
  var chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  for (var i = 0; i < 10; i++) k += chars.charAt(Math.floor(Math.random() * chars.length));
  var sh = ss_().getSheetByName(SHEET.SETTINGS);
  sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setValues([['大螢幕看板金鑰', k]]);
  return k;
}

/** 大螢幕看板:要帶看板金鑰。只回傳:人數統計、最近報到的幾組(姓名、單位、桌次、人數) */
function boardData(key) {
  var expected = getSetting_('大螢幕看板金鑰');
  if (!expected || String(key == null ? '' : key) !== expected) return { ok: false, error: '看板金鑰不正確,請用「報到系統」選單⑤ 取得的看板網址。' };
  var list = readTickets_().list;
  var showNames = getSetting_('看板顯示姓名') !== '否';
  var primary = {}, parties = {}, inCount = 0, giftCount = 0;
  list.forEach(function (t) {
    var k = norm_(t.party);
    if (suffix_(t) === 0) primary[k] = t;
    if (t.inAt) {
      inCount++;
      var p = parties[k] = parties[k] || { key: k, party: t.party, n: 0, at: '' };
      p.n++;
      if (t.inAt > p.at) p.at = t.inAt;
    }
    if (t.giftAt) giftCount++;
  });
  var recent = Object.keys(parties).map(function (k) { return parties[k]; })
    .sort(function (a, b) { return a.at < b.at ? 1 : a.at > b.at ? -1 : 0; })
    .slice(0, 12)
    .map(function (p) {
      var h = primary[p.key] || {};
      return {
        id: p.party + '@' + p.at,
        name: showNames ? (h.name || p.party) : '貴賓',
        org: showNames ? (h.org || '') : '',
        table: h.table || '', n: p.n, at: p.at
      };
    });
  return { ok: true, total: list.length, in: inCount, gift: giftCount, recent: recent,
    now: Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), 'HH:mm:ss') };
}

/**
 * 貴賓查詢頁 / 電子報到憑證。必須同時輸入「代號」與「姓名」且相符,才回傳該組資料;
 * 只回傳這一組(本人與同行者)的姓名、桌次、報到狀態。錯誤時不透露是代號還是姓名錯。
 */
function guestLookup(code, name) {
  var key = norm_(extractCode_(code)), nk = nameKey_(name);
  if (!key || !nk) return { ok: false, error: '請輸入代號與姓名' };
  var list = readTickets_().list;
  var me = list.filter(function (t) { return norm_(t.code) === key && suffix_(t) === 0; })[0];
  if (!me || nameKey_(me.name) !== nk) {
    Utilities.sleep(600); // 稍微拖慢,避免被人一個個猜代號
    return { ok: false, error: '查無資料,請確認「代號」與「姓名」是否與邀請卡相同。' };
  }
  var members = list.filter(function (t) { return norm_(t.party) === norm_(me.party); })
    .sort(function (a, b) { return suffix_(a) - suffix_(b); })
    .map(function (t) { return { code: t.code, name: t.name, diet: t.diet, inAt: t.inAt, giftAt: t.giftAt }; });
  return { ok: true, party: { code: me.party, name: me.name, org: me.org, table: me.table, gift: me.gift, members: members } };
}

function readTickets_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEET.TICKETS);
  var tz = ss.getSpreadsheetTimeZone();
  var n = Math.max(sh.getLastRow() - 1, 0);
  var vals = n ? sh.getRange(2, 1, n, 12).getValues() : [];
  var list = [];
  vals.forEach(function (r, i) {
    if (String(r[0]).trim() === '') return;
    list.push({
      row: i + 2, code: String(r[0]).trim(), party: String(r[1]).trim(),
      name: String(r[2]), org: String(r[3]), type: String(r[4]), diet: String(r[5]),
      table: String(r[6]), gift: String(r[7]),
      inAt: fmtT_(r[8], tz), inDesk: String(r[9]), giftAt: fmtT_(r[10], tz), giftDesk: String(r[11])
    });
  });
  return { sh: sh, list: list };
}

function snapshot_() {
  var list = readTickets_().list.map(function (t) { delete t.row; return t; });
  return { ok: true, tickets: list, at: new Date().getTime() };
}

/**
 * 名冊:名單上每一位貴賓 + 回函狀態(yes 出席 / no 不克出席 / none 沒有回函)。
 * 「到了沒有」由報到網頁用手機上的報到資料即時計算,所以這裡不重複回傳。
 */
function roster_() {
  var guests = readGuests_();
  var rs = { map: {}, unmatched: [] }, rsvpError = '';
  try { rs = readRsvps_(guests); } catch (e) { rsvpError = String(e && e.message || e); }
  var rows = Object.keys(guests.byCode).map(function (k) {
    var g = guests.byCode[k], r = rs.map[k];
    return { code: g.code, name: g.name, org: g.org, type: g.type, table: g.table,
      rsvp: r ? (r.n > 0 ? 'yes' : 'no') : 'none' };
  });
  return { ok: true, rows: rows, unmatched: rs.unmatched, rsvpError: rsvpError };
}

function suffix_(t) { var m = t.code.match(/-(\d+)$/); return m ? parseInt(m[1], 10) : 0; }

function seen_(opId) {
  if (!opId) return false;
  var c = CacheService.getScriptCache();
  if (c.get('op:' + opId)) return true;
  return false;
}
function markSeen_(opId) {
  if (opId) CacheService.getScriptCache().put('op:' + opId, '1', 21600);
}

function writeLogs_(logs) {
  if (!logs.length) return;
  var sh = ss_().getSheetByName(SHEET.LOG);
  sh.getRange(sh.getLastRow() + 1, 1, logs.length, 5).setValues(logs);
}

function process_(arg) {
  if (seen_(arg.opId)) return { ok: true, replay: true, results: [] };
  var d = readTickets_();
  var desk = String(arg.desk || '?');
  var mode = arg.mode === 'in' ? 'in' : 'both';
  var now = new Date();
  var results = [], logs = [];
  (arg.codes || []).forEach(function (code) {
    var key = norm_(code);
    var t = d.list.filter(function (x) { return norm_(x.code) === key; })[0];
    if (!t) { results.push({ code: code, status: 'notfound' }); return; }
    var needIn = !t.inAt, needGift = mode === 'both' && !t.giftAt;
    if (needIn && needGift) d.sh.getRange(t.row, 9, 1, 4).setValues([[now, desk, now, desk]]);
    else if (needIn) d.sh.getRange(t.row, 9, 1, 2).setValues([[now, desk]]);
    else if (needGift) d.sh.getRange(t.row, 11, 1, 2).setValues([[now, desk]]);
    if (needIn || needGift) {
      var did = (needIn ? ['報到'] : []).concat(needGift ? ['領禮'] : []);
      results.push({ code: t.code, status: 'ok', did: did });
      logs.push([now, desk, did.join('+'), t.code, '成功']);
    } else {
      results.push({ code: t.code, status: 'dup', inAt: t.inAt, inDesk: t.inDesk, giftAt: t.giftAt, giftDesk: t.giftDesk });
      logs.push([now, desk, mode === 'in' ? '報到' : '報到+領禮', t.code, '重複(原 ' + t.inAt + ' 第' + t.inDesk + '處)']);
    }
  });
  writeLogs_(logs);
  markSeen_(arg.opId);
  return { ok: true, results: results };
}

function undo_(arg) {
  if (seen_(arg.opId)) return { ok: true, replay: true };
  var d = readTickets_();
  var key = norm_(arg.code);
  var t = d.list.filter(function (x) { return norm_(x.code) === key; })[0];
  if (!t) return { ok: false, error: '找不到票 ' + arg.code };
  var desk = String(arg.desk || '?');
  if (arg.what === 'gift') {
    d.sh.getRange(t.row, 11, 1, 2).clearContent();
    writeLogs_([[new Date(), desk, '撤銷領禮', t.code, '成功']]);
  } else {
    d.sh.getRange(t.row, 9, 1, 4).clearContent();
    writeLogs_([[new Date(), desk, '撤銷報到', t.code, '成功']]);
  }
  markSeen_(arg.opId);
  return { ok: true };
}

/** 補填或修改某位的葷素(報到時現場詢問) */
function setDiet_(arg) {
  if (seen_(arg.opId)) return { ok: true, replay: true };
  var d = readTickets_();
  var key = norm_(arg.code);
  var t = d.list.filter(function (x) { return norm_(x.code) === key; })[0];
  if (!t) return { ok: false, error: '找不到 ' + arg.code };
  var diet = String(arg.diet == null ? '' : arg.diet).slice(0, 60);
  d.sh.getRange(t.row, 6).setValue(diet);
  writeLogs_([[new Date(), String(arg.desk || '?'), '修改葷素', t.code, diet || '(清空)']]);
  markSeen_(arg.opId);
  return { ok: true };
}

function addMember_(arg) {
  var d = readTickets_();
  var key = norm_(arg.party);
  var mine = d.list.filter(function (t) { return norm_(t.party) === key; });
  if (!mine.length) return { ok: false, error: '找不到此組' };
  var max = 0;
  mine.forEach(function (t) { max = Math.max(max, suffix_(t)); });
  var idx = max + 1;
  var party = mine[0].party;
  var r = d.sh.getLastRow() + 1;
  d.sh.getRange(r, 1, 1, COLS).setValues([ticketRow_(party + '-' + idx, party, idx, r, String(arg.diet == null ? '' : arg.diet).slice(0, 60))]);
  SpreadsheetApp.flush();
  writeLogs_([[new Date(), String(arg.desk || '?'), '現場加人', party + '-' + idx, '成功']]);
  return { ok: true, code: party + '-' + idx };
}
