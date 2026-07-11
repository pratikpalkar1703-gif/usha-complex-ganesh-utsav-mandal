// ─────────────────────────────────────────────
//  Ganpati Decoration 2026 – Google Sheets API
// ─────────────────────────────────────────────

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Receipt #','Date','Wing','Floor','Room No.','Full Name','Mobile','Amount (₹)','Status','Payment Mode','Type']);
    const h = sheet.getRange(1,1,1,11);
    h.setBackground('#f97316'); h.setFontColor('#ffffff'); h.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

// ── POST: insert new row OR update existing row ──
function doPost(e) {
  try {
    const sheet = getSheet();
    const data  = JSON.parse(e.postData.contents);

    // ── UPDATE existing row ──
    if (data.action === 'update') {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const vals = sheet.getRange(2,1,lastRow-1,9).getValues();
        for (let i = 0; i < vals.length; i++) {
          if (String(vals[i][0]) === String(data.receiptNum)) {
            const row = i + 2;
            sheet.getRange(row,6).setValue(data.name);
            sheet.getRange(row,7).setValue(data.mobile);
            sheet.getRange(row,8).setValue(Number(data.amount));
            sheet.autoResizeColumns(1,9);
            return ContentService
              .createTextOutput(JSON.stringify({status:'updated'}))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({status:'not_found'}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── INSERT new row ──
    ensureHeader(sheet);

    // If this is a Collected entry, delete any prior "Not Available" row for same room
    if ((data.status || 'Collected') !== 'Not Available' && data.roomCode) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const vals = sheet.getRange(2,1,lastRow-1,9).getValues();
        // Iterate bottom-up so row indices stay valid after deleteRow
        for (let i = vals.length - 1; i >= 0; i--) {
          if (String(vals[i][4]) === String(data.roomCode) && String(vals[i][8]) === 'Not Available') {
            sheet.deleteRow(i + 2);
          }
        }
      }
    }

    sheet.appendRow([
      data.receiptNum,
      data.date,
      data.wing,
      data.floor,
      data.roomCode  || '—',
      data.name      || '—',
      data.mobile    || '—',
      data.status === 'Not Available' ? 'N/A' : Number(data.amount),
      data.status    || 'Collected',
      data.paymentMode || '—',
      data.type        || 'Resident'
    ]);
    sheet.autoResizeColumns(1,11);

    return ContentService
      .createTextOutput(JSON.stringify({status:'success'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status:'error',message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET ──
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // ── Update an existing row (GET so response is readable) ──
  if (action === 'update') {
    try {
      const receiptNum = e.parameter.receiptNum;
      const sheet      = getSheet();
      const lastRow    = sheet.getLastRow();
      if (lastRow > 1) {
        const vals = sheet.getRange(2,1,lastRow-1,9).getValues();
        for (let i = 0; i < vals.length; i++) {
          if (String(vals[i][0]).trim() === String(receiptNum).trim()) {
            const row = i + 2;
            sheet.getRange(row,6).setValue(e.parameter.name);
            sheet.getRange(row,7).setValue(e.parameter.mobile);
            sheet.getRange(row,8).setValue(Number(e.parameter.amount));
            sheet.autoResizeColumns(1,9);
            return ContentService
              .createTextOutput(JSON.stringify({status:'updated', row: row}))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({status:'not_found', receiptNum: receiptNum}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({status:'error', message: err.toString()}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── Check if room already has a Collected entry ──
  if (action === 'checkRoom') {
    try {
      const roomCode = e.parameter.room;
      const sheet    = getSheet();
      const lastRow  = sheet.getLastRow();
      const results  = [];
      if (lastRow > 1) {
        const vals = sheet.getRange(2,1,lastRow-1,9).getValues();
        vals.forEach(r => {
          if (String(r[4]) === roomCode && String(r[8]) !== 'Not Available') {
            results.push({receiptNum:String(r[0]),date:String(r[1]),name:String(r[5]),mobile:String(r[6]),amount:r[7]});
          }
        });
      }
      return ContentService
        .createTextOutput(JSON.stringify({status:'success',matches:results}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({status:'error',matches:[]}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── All rows for dashboard ──
  if (action === 'getData') {
    try {
      const sheet   = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService
          .createTextOutput(JSON.stringify({status:'success',data:[]}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const rows = sheet.getRange(2,1,lastRow-1,11).getValues();
      const data = rows.map(r => ({
        receiptNum:r[0],date:r[1],wing:r[2],floor:r[3],
        roomCode:r[4],name:r[5],mobile:r[6],amount:r[7],
        status:      r[8]  || 'Collected',
        paymentMode: r[9]  || '—',
        type:        r[10] || 'Resident'
      }));
      return ContentService
        .createTextOutput(JSON.stringify({status:'success',data:data}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({status:'error',message:err.toString()}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({status:'live',app:'Ganpati Decoration 2026'}))
    .setMimeType(ContentService.MimeType.JSON);
}
