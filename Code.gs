// ─────────────────────────────────────────────
//  Ganpati Decoration 2026 – Google Sheets API
// ─────────────────────────────────────────────

var DRIVE_FOLDER = 'Usha Complex Receipts 2026';

function testDriveAuth() {
  const folder = DriveApp.getRootFolder();
  Logger.log('Drive authorized! Root folder: ' + folder.getName());
}

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

// ── POST: insert new row OR update existing row OR upload PDF ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // ── UPLOAD PDF to Drive ──
    if (data.action === 'uploadPDF') {
      var step = 'init';
      try {
        step = 'decode';
        const pdfBytes = Utilities.base64Decode(data.pdf);
        const blob = Utilities.newBlob(pdfBytes, 'application/pdf',
                      'Receipt-' + data.data.receiptNum + '-' + (data.data.roomCode||'') + '.pdf');

        const roomCode = String(data.data.roomCode || '').trim();
        const wingMatch = roomCode.match(/^([A-La-l])-/i);
        const subFolderName = wingMatch
          ? wingMatch[1].toUpperCase() + ' Wing'
          : 'Non Residents';

        step = 'getOrCreateFolder';
        const parent = getOrCreateFolder(DRIVE_FOLDER);

        step = 'getOrCreateSubFolder';
        const folder = getOrCreateSubFolder(parent, subFolderName);

        step = 'checkExisting';
        const fileName = 'Receipt-' + data.data.receiptNum + '-' + (data.data.roomCode||'') + '.pdf';
        const existingFiles = folder.getFilesByName(fileName);
        if (existingFiles.hasNext()) {
          const existingFile = existingFiles.next();
          const existingLink = existingFile.getUrl();
          // Backfill column 12 if empty
          try {
            const sheet2 = getSheet();
            const lr2 = sheet2.getLastRow();
            if (lr2 > 1) {
              const vs = sheet2.getRange(2, 1, lr2 - 1, 12).getValues();
              for (let i = 0; i < vs.length; i++) {
                if (String(vs[i][0]) === String(data.data.receiptNum) && !vs[i][11]) {
                  sheet2.getRange(i + 2, 12).setValue(existingLink);
                  break;
                }
              }
            }
          } catch(_) {}
          return ContentService
            .createTextOutput(JSON.stringify({status:'ok', link: existingLink}))
            .setMimeType(ContentService.MimeType.JSON);
        }

        step = 'createFile';
        const file = folder.createFile(blob);

        step = 'setSharing';
        try {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch(shareErr) {
          Logger.log('setSharing failed (non-fatal): ' + shareErr.toString());
        }

        step = 'done';
        const link = file.getUrl();
        Logger.log('Uploaded: ' + file.getName() + ' → ' + subFolderName + ' → ' + link);

        // Save link to sheet column 12
        try {
          const sheet = getSheet();
          const lastRow = sheet.getLastRow();
          if (lastRow > 1) {
            const vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
            for (let i = 0; i < vals.length; i++) {
              if (String(vals[i][0]) === String(data.data.receiptNum)) {
                sheet.getRange(i + 2, 12).setValue(link);
                break;
              }
            }
          }
        } catch(sheetErr) {
          Logger.log('Could not save link to sheet: ' + sheetErr.toString());
        }

        return ContentService
          .createTextOutput(JSON.stringify({status:'ok', link: link}))
          .setMimeType(ContentService.MimeType.JSON);

      } catch(driveErr) {
        return ContentService
          .createTextOutput(JSON.stringify({status:'error', message:'[step:' + step + '] ' + driveErr.toString()}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    const sheet = getSheet();

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
      data.roomCode    || '—',
      data.name        || '—',
      data.mobile      || '—',
      data.status === 'Not Available' ? 'N/A' : Number(data.amount),
      data.status      || 'Collected',
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

  // ── GET LINK: look up Drive link for a receipt (no PDF needed) ──
  if (action === 'getLink') {
    try {
      const receiptNum = e.parameter.receiptNum || '';
      const roomCode   = e.parameter.roomCode   || '';
      const sheet      = getSheet();
      const lastRow    = sheet.getLastRow();

      // 1. Check column 12 in sheet first
      if (lastRow > 1) {
        const vals = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
        for (let i = 0; i < vals.length; i++) {
          if (String(vals[i][0]) === String(receiptNum) && vals[i][11]) {
            return ContentService
              .createTextOutput(JSON.stringify({status:'ok', link: vals[i][11]}))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      // 2. Not in sheet — search Drive by filename
      const wingMatch     = String(roomCode).match(/^([A-La-l])-/i);
      const subFolderName = wingMatch ? wingMatch[1].toUpperCase() + ' Wing' : 'Non Residents';
      const fileName      = 'Receipt-' + receiptNum + '-' + roomCode + '.pdf';
      const parent        = getOrCreateFolder(DRIVE_FOLDER);
      const subFolder     = getOrCreateSubFolder(parent, subFolderName);
      const files         = subFolder.getFilesByName(fileName);
      if (files.hasNext()) {
        const link = files.next().getUrl();
        // Save to sheet column 12 for next time
        if (lastRow > 1) {
          const vs = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
          for (let i = 0; i < vs.length; i++) {
            if (String(vs[i][0]) === String(receiptNum)) {
              sheet.getRange(i + 2, 12).setValue(link);
              break;
            }
          }
        }
        return ContentService
          .createTextOutput(JSON.stringify({status:'ok', link: link}))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // 3. Not found anywhere
      return ContentService
        .createTextOutput(JSON.stringify({status:'not_found'}))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({status:'error', message: err.toString()}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

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

  if (action === 'getData') {
    try {
      const sheet   = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService
          .createTextOutput(JSON.stringify({status:'success',data:[]}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const rows = sheet.getRange(2,1,lastRow-1,12).getValues();
      const data = rows.map(r => ({
        receiptNum:r[0],date:r[1],wing:r[2],floor:r[3],
        roomCode:r[4],name:r[5],mobile:r[6],amount:r[7],
        status:      r[8]  || 'Collected',
        paymentMode: r[9]  || '—',
        type:        r[10] || 'Resident',
        driveLink:   r[11] || ''
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

// ── DRIVE HELPERS ─────────────────────────────
function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function getOrCreateSubFolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}
