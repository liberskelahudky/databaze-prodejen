/**
 * ═══════════════════════════════════════════════════════════════════
 *  LL DATABÁZE PRODEJEN — Apps Script v4
 * ═══════════════════════════════════════════════════════════════════
 *
 *  CO JE NOVÉ oproti v3:
 *   • listy si vytvoří sám: DATA (úložiště), Prodejny (čitelný přehled), Log (historie změn)
 *   • ukládá VŠECHNA pole včetně nových (glory, terminály, typ váhy,
 *     vysavače, ledovač, ledový box, tříšť…) — nic se neztrácí
 *   • list "Prodejny" se po každém uložení přegeneruje jako čitelná
 *     tabulka (jen pro čtení/filtry — NEEDITOVAT, přepíše se!)
 *   • list "Log" = sdílená historie změn (kdo/kdy/co), čte ji appka
 *     přes tlačítko 🕘 Historie → „Sdílená (Sheets)"
 *
 *  JAK NASADIT (3 kroky):
 *   1. Otevři script.google.com → svůj projekt → smaž starý kód
 *      a vlož celý tento soubor.
 *   2. Deploy → Manage deployments → ✏️ u stávajícího deploymentu
 *      → Version: "New version" → Deploy.  (URL zůstane stejná,
 *      v appce se nic měnit nemusí!)
 *   3. V appce klikni ⬆ Uložit — tím se naplní nové úložiště DATA.
 *      (Do té doby appka hlásí „Sheets prázdné" a jede z lokálních dat.)
 */

var KEY = 'LL2026';
var CHUNK = 45000; // limit ~50k znaků na buňku

// ─── pomocné ───────────────────────────────────────────────────────
function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function _sheet(name) {
  var ss = _ss();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function _jsonp(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── úložiště: JSON blob po kusech v listu DATA, sloupec A ─────────
function _writeBlob(jsonStr) {
  var sh = _sheet('DATA');
  sh.clearContents();
  var rows = [];
  for (var i = 0; i < jsonStr.length; i += CHUNK) {
    rows.push([jsonStr.substring(i, i + CHUNK)]);
  }
  if (rows.length) sh.getRange(1, 1, rows.length, 1).setValues(rows);
}

function _readBlob() {
  var sh = _ss().getSheetByName('DATA');
  if (!sh || sh.getLastRow() === 0) return null;
  var vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  var jsonStr = vals.map(function(r){ return r[0]; }).join('');
  if (!jsonStr) return null;
  try { return JSON.parse(jsonStr); } catch (e) { return null; }
}

// ─── čitelný přehled "Prodejny" (jen pro lidi, přegenerovává se) ───
var COLS = [
  ['code','Číslo'], ['name','Prodejna'], ['type','Typ'], ['addr','Adresa'],
  ['tel','Telefon'], ['email','E-mail'], ['vedouci','Vedoucí'], ['vedouciTel','Tel. vedoucí'],
  ['wolt','Wolt'], ['pokladny','Pokladny'], ['terminaly','Terminály'], ['glory','Glory'],
  ['vahy','Váhy u kas'], ['typVahy','Typ váhy'], ['lcd','TV'],
  ['ledovac','Ledovač'], ['ledovacBox','Ledový box'],
  ['trist','Tříšť'], ['kiosek','Kiosek'], ['vyroba','Výroba'], ['eshop','E-shop'],
  ['vysavacVelky','Vysavač velký'], ['vysavacRucni','Vysavač ruční'],
  ['samoLednice','Sam. lednice'], ['seating','Posezení'], ['barrier','Bezbariér'],
  ['zmrzType','Zmrzlina typ'], ['zmrzHours','Zmrzlina hodiny'],
  ['hours','Otevírací doba'], ['leaseEnd','Konec nájmu'], ['landlord','Pronajímatel']
];

function _cellVal(s, field) {
  var v = s[field];
  if (v === true) return 'ANO';
  if (v === false || v === undefined || v === null) return '';
  if (field === 'hours' && Array.isArray(v)) return v.join(' | ');
  if (Array.isArray(v) || typeof v === 'object') return '';
  return v;
}

function _rebuildReadable(storesArr) {
  var sh = _sheet('Prodejny');
  sh.clearContents();
  var header = COLS.map(function(c){ return c[1]; });
  var rows = [header];
  storesArr.forEach(function(s) {
    rows.push(COLS.map(function(c){ return _cellVal(s, c[0]); }));
  });
  sh.getRange(1, 1, rows.length, header.length).setValues(rows);
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#F6F1E7');
  sh.setFrozenRows(1);
}

// ─── sdílená historie změn ─────────────────────────────────────────
function _appendLog(entries, author) {
  if (!entries || !entries.length) return;
  var sh = _sheet('Log');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4)
      .setValues([['Čas', 'Kdo', 'Prodejna', 'Změna']])
      .setFontWeight('bold').setBackground('#F6F1E7');
    sh.setFrozenRows(1);
  }
  var rows = entries.map(function(e) {
    return [new Date(e.t || Date.now()), e.who || author || '', e.store || '', e.txt || ''];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  // drž max ~500 řádků, ať list neroste donekonečna
  var max = 500, extra = sh.getLastRow() - 1 - max;
  if (extra > 0) sh.deleteRows(2, extra);
}

function _readLog(limit) {
  var sh = _ss().getSheetByName('Log');
  if (!sh || sh.getLastRow() < 2) return [];
  var n = Math.min(limit || 50, sh.getLastRow() - 1);
  var vals = sh.getRange(sh.getLastRow() - n + 1, 1, n, 4).getValues();
  return vals.map(function(r) {
    var t = (r[0] instanceof Date) ? r[0].getTime() : Date.now();
    return { t: t, who: String(r[1] || ''), store: String(r[2] || ''), txt: String(r[3] || '') };
  });
}

// ═══ GET: ?action=get | ?action=log  (+ &key= &callback=) ══════════
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || '';
  if (p.key !== KEY) return _jsonp({ ok: false, error: 'bad key' }, cb);

  if (p.action === 'get') {
    var blob = _readBlob();
    if (!blob) return _jsonp({ ok: true, stores: [], savedAt: 0 }, cb);
    return _jsonp({ ok: true, stores: blob.stores || [], savedAt: blob.savedAt || 0 }, cb);
  }
  if (p.action === 'log') {
    return _jsonp({ ok: true, log: _readLog(50) }, cb);
  }
  return _jsonp({ ok: true, info: 'LL Apps Script v4', actions: ['get', 'log'] }, cb);
}

// ═══ POST: uložení {key, author, log, data:{stores, savedAt}} ══════
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // dva současné zápisy se serializují — nic se nepřepíše napůl
  try {
    var payload;
    try { payload = JSON.parse(e.postData.contents); }
    catch (err) { return _jsonp({ ok: false, error: 'bad json' }); }

    if (payload.key !== KEY) return _jsonp({ ok: false, error: 'bad key' });
    var data = payload.data || {};
    if (!Array.isArray(data.stores)) return _jsonp({ ok: false, error: 'no stores' });

    // Pojistka proti přepsání novějších dat staršími:
    var prev = _readBlob();
    if (prev && prev.savedAt && data.savedAt && data.savedAt < prev.savedAt) {
      return _jsonp({ ok: false, error: 'stale', serverSavedAt: prev.savedAt,
        message: 'V Sheets jsou novější data — nejdřív Načíst ze Sheets.' });
    }

    _writeBlob(JSON.stringify({ stores: data.stores, savedAt: data.savedAt || Date.now() }));
    _rebuildReadable(data.stores);
    _appendLog(payload.log || [], payload.author || '');

    return _jsonp({ ok: true, savedAt: data.savedAt, stores: data.stores.length });
  } finally {
    lock.releaseLock();
  }
}
