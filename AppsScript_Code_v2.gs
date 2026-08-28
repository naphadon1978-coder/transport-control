// ===== Transport Control - Google Apps Script Backend (v2) =====
// วางโค้ดนี้ทั้งหมดแทนของเดิมใน Extensions > Apps Script > Code.gs

const FACTORY_STEPS = ["เตรียมสินค้าแล้ว","ตรวจนับสินค้าแล้ว","โหลดสินค้าเสร็จแล้ว","เปิดบิลแล้ว","รถออกจากบริษัทแล้ว"];

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Trucks');
  var data = sheet.getDataRange().getValues();
  var trucks = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    try {
      trucks.push(JSON.parse(data[i][1]));
    } catch (err) {
      // skip broken row
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, trucks: trucks }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var body = JSON.parse(e.postData.contents);

  if (body.action === 'saveTruck') {
    var sheet = ss.getSheetByName('Trucks');
    var data = sheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === body.truck.id) { rowIdx = i + 1; break; }
    }
    var json = JSON.stringify(body.truck);
    if (rowIdx === -1) {
      sheet.appendRow([body.truck.id, json, new Date()]);
    } else {
      sheet.getRange(rowIdx, 2).setValue(json);
      sheet.getRange(rowIdx, 3).setValue(new Date());
    }
    // เขียนแท็บ Report แบบอ่านง่ายให้อัตโนมัติด้วย
    upsertReportRow(ss, body.truck);
  }

  if (body.action === 'deleteTruck') {
    var sheet2 = ss.getSheetByName('Trucks');
    var data2 = sheet2.getDataRange().getValues();
    for (var j = 1; j < data2.length; j++) {
      if (data2[j][0] === body.id) { sheet2.deleteRow(j + 1); break; }
    }
    deleteReportRow(ss, body.id);
  }

  if (body.action === 'logEvent') {
    var evSheet = ss.getSheetByName('Events');
    evSheet.appendRow([new Date(), body.plate, body.status, body.stopIndex || '', body.customer || '']);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function upsertReportRow(ss, truck) {
  var sheet = ss.getSheetByName('Report');
  if (!sheet) return; // ถ้ายังไม่ได้สร้างแท็บ Report ไว้ ให้ข้ามไปเงียบๆ ไม่ error

  var stops = truck.stops || [];
  var history = truck.history || {};
  var mismatch = truck.assignedDriver && truck.verifiedDriver && truck.assignedDriver !== truck.verifiedDriver;
  var custList = stops.map(function (s) { return s.customer || '-'; }).join(' -> ');

  var row = [
    truck.id,
    truck.verifiedDriver || '',
    truck.assignedDriver || '',
    truck.verifiedDriver ? (mismatch ? 'ไม่ตรงกัน' : 'ตรงกัน') : '',
    custList
  ];

  FACTORY_STEPS.forEach(function (s) {
    row.push(history[s] ? history[s].time : '');
  });

  for (var k = 0; k < 3; k++) {
    var arriveLabel = 'ถึงลูกค้าจุดที่ ' + (k + 1);
    var doneLabel = 'ลงสินค้าเสร็จจุดที่ ' + (k + 1);
    row.push(history[arriveLabel] ? history[arriveLabel].time : '');
    row.push(history[doneLabel] ? history[doneLabel].time : '');
  }

  row.push(history['รถกำลังกลับบริษัท'] ? history['รถกำลังกลับบริษัท'].time : '');
  row.push(history['รถถึงบริษัทแล้ว'] ? history['รถถึงบริษัทแล้ว'].time : '');
  row.push(new Date());

  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === truck.id) { rowIdx = i + 1; break; }
  }
  if (rowIdx === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  }
}

function deleteReportRow(ss, plateId) {
  var sheet = ss.getSheetByName('Report');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === plateId) { sheet.deleteRow(i + 1); break; }
  }
}
