function doGet(e) {
  return jsonResponse(getDataset());
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action;
  var payload = body.payload || {};

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
      return jsonResponse({ error: 'Unknown action: ' + action });
  }

  return jsonResponse(getDataset());
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
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
        return { storeId: resp.storeId, count: resp.count, status: resp.status };
      })
    };
  });

  return { stores: stores, requests: requests };
}

function addStore(payload) {
  var sheet = getSheet('Stores');
  var stores = sheetToObjects(sheet);
  var maxId = 0;
  stores.forEach(function(s) { maxId = Math.max(maxId, Number(s.id)); });
  sheet.appendRow([maxId + 1, payload.name]);
}

function removeStore(payload) {
  var sheet = getSheet('Stores');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(payload.id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function submitRequest(payload) {
  var sheet = getSheet('Requests');
  var requests = sheetToObjects(sheet);
  var maxId = 0;
  requests.forEach(function(r) { maxId = Math.max(maxId, Number(r.id)); });
  var newId = maxId + 1;
  var createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([
    newId, payload.storeId, payload.date, payload.time,
    payload.staffType, payload.count, payload.note || '',
    'open', createdAt
  ]);
}

function submitResponse(payload) {
  var responsesSheet = getSheet('Responses');
  responsesSheet.appendRow([payload.requestId, payload.storeId, payload.count, 'pending']);
  setRequestStatus(payload.requestId, 'responded');
}

function approveResponse(payload) {
  setResponseStatus(payload.requestId, payload.storeId, 'approved');
  setRequestStatus(payload.requestId, 'approved');
}

function rejectResponse(payload) {
  setResponseStatus(payload.requestId, payload.storeId, 'rejected');

  var responsesSheet = getSheet('Responses');
  var values = responsesSheet.getDataRange().getValues();
  var hasNonRejected = false;
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][0]) === Number(payload.requestId) && String(values[i][3]) !== 'rejected') {
      hasNonRejected = true;
      break;
    }
  }
  if (!hasNonRejected) {
    setRequestStatus(payload.requestId, 'open');
  }
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
      break;
    }
  }
}

function setResponseStatus(requestId, storeId, status) {
  var sheet = getSheet('Responses');
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var statusCol = headers.indexOf('status') + 1;
  var reqIdCol = headers.indexOf('requestId');
  var storeIdCol = headers.indexOf('storeId');
  for (var i = 1; i < values.length; i++) {
    if (Number(values[i][reqIdCol]) === Number(requestId) && Number(values[i][storeIdCol]) === Number(storeId)) {
      sheet.getRange(i + 1, statusCol).setValue(status);
      break;
    }
  }
}
