var MAX_STORES = 20;

// フロントエンドはJSONP（GET + callbackパラメータ）で通信するため、
// 読み取りも書き込みもdoGetで処理する。doPostは互換用に残す。
function doGet(e) {
  var params = (e && e.parameter) || {};
  var result;
  if (params.action) {
    result = runAction(params.action, parsePayload(params.payload));
  } else {
    try {
      result = getDataset();
    } catch (err) {
      result = { error: errorMessage(err) };
    }
  }
  return jsonResponse(result, params.callback);
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'Invalid JSON body' });
  }
  return jsonResponse(runAction(body.action, body.payload || {}));
}

function parsePayload(raw) {
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch (err) {
    return {};
  }
}

function runAction(action, payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return { error: '他の処理が実行中です。少し待ってからもう一度お試しください' };
  }
  try {
    switch (action) {
      case 'addStore':
        addStore(payload);
        break;
      case 'removeStore':
        removeStore(payload);
        break;
      case 'submitRequest':
        submitRequest(payload);
        break;
      case 'submitResponse':
        submitResponse(payload);
        break;
      case 'approveResponse':
        approveResponse(payload);
        break;
      case 'rejectResponse':
        rejectResponse(payload);
        break;
      default:
        return { error: 'Unknown action: ' + action };
    }
    return getDataset();
  } catch (err) {
    return { error: errorMessage(err) };
  } finally {
    lock.releaseLock();
  }
}

function errorMessage(err) {
  return err && err.message ? err.message : String(err);
}

function jsonResponse(obj, callback) {
  var json = JSON.stringify(obj);
  // コールバック名は識別子のみ許可（JSONP経由のスクリプト注入を防ぐ）
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」が見つかりません。スプレッドシートの設定を確認してください');
  }
  return sheet;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function appendRowByHeaders(sheet, obj) {
  var headers = getHeaders(sheet);
  var row = headers.map(function(h) {
    return obj.hasOwnProperty(h) ? obj[h] : '';
  });
  sheet.appendRow(row);
}

function sheetToObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var rows = values.slice(1);
  var result = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var hasValue = false;
    for (var j = 0; j < row.length; j++) {
      if (row[j] !== '' && row[j] !== null) { hasValue = true; break; }
    }
    if (!hasValue) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    result.push(obj);
  }
  return result;
}

function formatDate(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

function getDataset() {
  var storesSheet = getSheet('Stores');
  var requestsSheet = getSheet('Requests');
  var responsesSheet = getSheet('Responses');

  var stores = sheetToObjects(storesSheet).map(function(s) {
    return { id: Number(s.id), name: String(s.name) };
  });

  var responsesRaw = sheetToObjects(responsesSheet).map(function(r) {
    return {
      requestId: Number(r.requestId),
      storeId: Number(r.storeId),
      count: Number(r.count),
      names: r.names ? String(r.names) : '',
      status: String(r.status)
    };
  });

  var requests = sheetToObjects(requestsSheet).map(function(r) {
    var id = Number(r.id);
    return {
      id: id,
      storeId: Number(r.storeId),
      date: formatDate(r.date),
      time: String(r.time),
      staffType: String(r.staffType),
      count: Number(r.count),
      note: r.note ? String(r.note) : '',
      status: String(r.status),
      createdAt: formatDate(r.createdAt),
      responses: responsesRaw.filter(function(resp) { return resp.requestId === id; }).map(function(resp) {
        return { storeId: resp.storeId, count: resp.count, names: resp.names, status: resp.status };
      })
    };
  });

  return { stores: stores, requests: requests };
}

function addStore(payload) {
  var name = payload.name ? String(payload.name).trim() : '';
  if (!name) throw new Error('店舗名を入力してください');

  var sheet = getSheet('Stores');
  var stores = sheetToObjects(sheet);
  if (stores.length >= MAX_STORES) throw new Error('最大' + MAX_STORES + '店舗まで登録できます');

  var maxId = 0;
  for (var i = 0; i < stores.length; i++) {
    if (String(stores[i].name) === name) throw new Error('同じ名前の店舗が既に登録されています');
    maxId = Math.max(maxId, Number(stores[i].id));
  }
  appendRowByHeaders(sheet, { id: maxId + 1, name: name });
}

function removeStore(payload) {
  var sheet = getSheet('Stores');
  var values = sheet.getDataRange().getValues();
  if (values.length <= 2) throw new Error('最低1店舗は必要です');
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(payload.id)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
  throw new Error('店舗が見つかりません');
}

function submitRequest(payload) {
  var storeId = Number(payload.storeId);
  var count = Number(payload.count);
  if (!payload.date) throw new Error('日付を入力してください');
  if (!(count >= 1)) throw new Error('必要人数が正しくありません');
  if (!findStore(storeId)) throw new Error('店舗情報が見つかりません');

  var sheet = getSheet('Requests');
  var requests = sheetToObjects(sheet);
  var maxId = 0;
  requests.forEach(function(r) { maxId = Math.max(maxId, Number(r.id)); });
  var createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  appendRowByHeaders(sheet, {
    id: maxId + 1,
    storeId: storeId,
    date: String(payload.date),
    time: String(payload.time || ''),
    staffType: String(payload.staffType || ''),
    count: count,
    note: payload.note ? String(payload.note) : '',
    status: 'open',
    createdAt: createdAt
  });
}

function findStore(storeId) {
  var stores = sheetToObjects(getSheet('Stores'));
  for (var i = 0; i < stores.length; i++) {
    if (Number(stores[i].id) === storeId) return stores[i];
  }
  return null;
}

function findRequest(requestId) {
  var requests = sheetToObjects(getSheet('Requests'));
  for (var i = 0; i < requests.length; i++) {
    if (Number(requests[i].id) === requestId) return requests[i];
  }
  return null;
}

function submitResponse(payload) {
  var requestId = Number(payload.requestId);
  var storeId = Number(payload.storeId);
  var count = Number(payload.count);
  if (!(count >= 1)) throw new Error('派遣人数が正しくありません');

  var request = findRequest(requestId);
  if (!request) throw new Error('依頼が見つかりません');
  if (String(request.status) === 'approved') throw new Error('この依頼は必要人数に達しています');
  if (Number(request.storeId) === storeId) throw new Error('自店舗の依頼には応答できません');

  var sheet = getSheet('Responses');
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var reqIdCol = headers.indexOf('requestId');
  var storeIdCol = headers.indexOf('storeId');
  var countCol = headers.indexOf('count');
  var namesCol = headers.indexOf('names');
  var statusCol = headers.indexOf('status');

  // 却下済みの応答があれば上書きして再応答、未却下の応答があればエラー
  var rejectedRow = -1;
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][reqIdCol]) === requestId && Number(values[i][storeIdCol]) === storeId) {
      if (String(values[i][statusCol]) === 'rejected') {
        rejectedRow = i + 1;
      } else {
        throw new Error('既に応答済みです');
      }
    }
  }

  var names = payload.names ? String(payload.names) : '';
  if (rejectedRow > 0) {
    sheet.getRange(rejectedRow, countCol + 1).setValue(count);
    if (namesCol >= 0) sheet.getRange(rejectedRow, namesCol + 1).setValue(names);
    sheet.getRange(rejectedRow, statusCol + 1).setValue('pending');
  } else {
    appendRowByHeaders(sheet, {
      requestId: requestId,
      storeId: storeId,
      count: count,
      names: names,
      status: 'pending'
    });
  }
  updateRequestStatus(requestId);
}

function approveResponse(payload) {
  setResponseStatus(payload.requestId, payload.storeId, 'approved');
  updateRequestStatus(payload.requestId);
}

function rejectResponse(payload) {
  setResponseStatus(payload.requestId, payload.storeId, 'rejected');
  updateRequestStatus(payload.requestId);
}

// 部分承認対応：承認済み人数の合計が必要人数に達したらapproved、
// 承認待ちの応答が残っていればresponded、それ以外は引き続きopen（募集中）
function updateRequestStatus(requestId) {
  var request = findRequest(Number(requestId));
  if (!request) throw new Error('依頼が見つかりません');
  var responses = sheetToObjects(getSheet('Responses'));
  var approvedTotal = 0;
  var hasPending = false;
  for (var i = 0; i < responses.length; i++) {
    if (Number(responses[i].requestId) !== Number(requestId)) continue;
    var st = String(responses[i].status);
    if (st === 'approved') approvedTotal += Number(responses[i].count);
    else if (st === 'pending') hasPending = true;
  }
  var status;
  if (approvedTotal >= Number(request.count)) status = 'approved';
  else if (hasPending) status = 'responded';
  else status = 'open';
  setRequestStatus(requestId, status);
}

function setRequestStatus(requestId, status) {
  var sheet = getSheet('Requests');
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var statusCol = headers.indexOf('status') + 1;
  var idCol = headers.indexOf('id');
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][idCol]) === Number(requestId)) {
      sheet.getRange(i + 1, statusCol).setValue(status);
      return;
    }
  }
  throw new Error('依頼が見つかりません');
}

function setResponseStatus(requestId, storeId, status) {
  var sheet = getSheet('Responses');
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var statusCol = headers.indexOf('status') + 1;
  var reqIdCol = headers.indexOf('requestId');
  var storeIdCol = headers.indexOf('storeId');
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][reqIdCol]) === Number(requestId) &&
        Number(values[i][storeIdCol]) === Number(storeId) &&
        String(values[i][statusCol - 1]) !== 'rejected') {
      sheet.getRange(i + 1, statusCol).setValue(status);
      return;
    }
  }
  throw new Error('応答が見つかりません');
}
