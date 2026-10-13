// Yasou POS - Google Apps Script
// עמודות בגיליון: orderId | timestamp | table | total | station | items | note | status | time
var SHEET_ID = 'YOUR_SHEET_ID';
var SHEET_NAME = 'Orders';
var HEADERS = ['orderId', 'timestamp', 'table', 'total', 'station', 'items', 'note', 'status', 'time'];
var MAX_ROWS_READ = 400;

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// קבלת הזמנה חדשה, או סימון הזמנה כבוצעה
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    var last = sheet.getLastRow();

    if (data.action === 'done') {
      var ids = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getValues() : [];
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0] === data.orderId) sheet.getRange(i + 2, 8).setValue('done');
      }
      return json_({ ok: true });
    }

    var orderId = data.orderId;
    var existing = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getValues() : [];
    for (var j = 0; j < existing.length; j++) {
      if (existing[j][0] === orderId) return json_({ ok: true, dup: true });
    }

    var readable = Utilities.formatDate(new Date(data.timestamp), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    ['cold', 'hot', 'bar'].forEach(function (st) {
      var items = data.split ? data.split[st] : null;
      if (!items || items.length === 0) return;
      var notes = items.filter(function (it) { return it.note; }).map(function (it) { return it.note; }).join(' | ');
      sheet.appendRow([orderId, data.timestamp, data.table, data.total, st, JSON.stringify(items), notes, 'new', readable]);
    });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, err: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// ?action=orders מחזיר את ההזמנות האחרונות בפורמט שהאפליקציה מכירה
function doGet(e) {
  try {
    if (!e || !e.parameter || e.parameter.action !== 'orders') {
      return ContentService.createTextOutput('Yasou POS OK').setMimeType(ContentService.MimeType.TEXT);
    }
    var sheet = getSheet_();
    var last = sheet.getLastRow();
    if (last < 2) return json_({ ok: true, orders: [] });

    var start = Math.max(2, last - MAX_ROWS_READ + 1);
    var rows = sheet.getRange(start, 1, last - start + 1, 8).getValues();
    var map = {};
    var order = [];
    rows.forEach(function (r) {
      var id = r[0];
      if (!id) return;
      if (!map[id]) {
        map[id] = {
          orderId: id,
          timestamp: Number(r[1]),
          table: Number(r[2]),
          total: Number(r[3]),
          status: 'new',
          items: [],
          split: { cold: [], hot: [], bar: [] }
        };
        order.push(id);
      }
      var items = [];
      try { items = JSON.parse(r[5]); } catch (x) {}
      var o = map[id];
      if (o.split[r[4]]) o.split[r[4]] = items;
      o.items = o.items.concat(items);
      if (r[7] === 'done') o.status = 'done';
    });
    var orders = order.map(function (id) { return map[id]; });
    orders.sort(function (a, b) { return b.timestamp - a.timestamp; });
    return json_({ ok: true, orders: orders });
  } catch (err) {
    return json_({ ok: false, err: String(err) });
  }
}
