// ============================================================
// LIBEŘSKÉ LAHŮDKY — Google Apps Script backend v3
// Vše přes doGet + JSONP — funguje z file:// bez CORS
// ============================================================

const SHEET_NAME = 'Prodejny';
const SECRET_KEY = 'LL2026';

function doGet(e) {
  var params = e.parameter || {};
  var cb = params.callback || '';

  if (params.key !== SECRET_KEY) {
    return jsonp(cb, { error: 'Unauthorized' });
  }

  var action = params.action || 'get';

  if (action === 'get') {
    return jsonp(cb, getStores());
  }

  if (action === 'save') {
    try {
      var stores = JSON.parse(decodeURIComponent(params.data || '[]'));
      saveStores(stores);
      return jsonp(cb, { ok: true });
    } catch(err) {
      return jsonp(cb, { error: err.toString() });
    }
  }

  return jsonp(cb, { error: 'Unknown action' });
}

// Zachovej doPost pro případ přímého volání
function doPost(e) {
  return doGet(e);
}

function jsonp(callback, data) {
  var json = JSON.stringify(data);
  var output = callback ? callback + '(' + json + ')' : json;
  var mime = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(output).setMimeType(mime);
}

function getStores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { stores: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { stores: [] };

  var headers = data[0];
  var stores = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var val = row[i];
      if (['hours','saints','exceptions','changelog'].indexOf(h) > -1) {
        try { obj[h] = JSON.parse(val || (h === 'hours' ? '[]' : '{}')); }
        catch(e) { obj[h] = h === 'hours' ? [] : {}; }
      } else if (['barrier','seating','eshop','kiosek','vyroba','okenko'].indexOf(h) > -1) {
        obj[h] = val === true || val === 'true' || val === 1;
      } else if (['zmrzlina','nfc','lcd','dvere','lednice','vahy'].indexOf(h) > -1) {
        obj[h] = parseInt(val) || 0;
      } else {
        obj[h] = val !== undefined && val !== null ? String(val) : '';
      }
    });
    return obj;
  });

  return { stores: stores, updated: new Date().toISOString() };
}

function saveStores(stores) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  var headers = [
    'id','name','type','addr','tel','email','vedouci','vedouciTel',
    'obchZast','wolt','gmaps','drive','landlord','landlordTel','leaseEnd','area',
    'mhd','parking','barrier','seating','eshop','kiosek','vyroba',
    'hours','zmrzlina','nfc','lcd','dvere','kavovar','lednice','vahy','pos',
    'okenko','zmrzHours','notes','saints','exceptions','changelog'
  ];

  var rows = [headers];
  stores.forEach(function(s) {
    rows.push(headers.map(function(h) {
      var val = s[h];
      if (['hours','saints','exceptions','changelog'].indexOf(h) > -1) {
        return JSON.stringify(val || (h === 'hours' ? [] : {}));
      }
      return val !== undefined && val !== null ? val : '';
    }));
  });

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#18160f').setFontColor('white').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 5); // Resize prvních 5 sloupců
}
