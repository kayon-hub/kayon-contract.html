/**
 * ═══════════════════════════════════════════════════════════════
 * KAYON STUDIO｜電子合約 Google Apps Script 後端
 * 功能：接收合約資料 → 儲存 Google Sheet → 產生 PDF → 存 Drive → 發 Email
 * ═══════════════════════════════════════════════════════════════
 */

// ─── 設定區 ───────────────────────────────────────────────────
const CONFIG = {
  // Google Sheet ID（從 Sheet URL 取得）
  SHEET_ID: 'YOUR_GOOGLE_SHEET_ID', // ← 請在此填入您的 Google Sheet ID

  // Google Drive 資料夾 ID（合約儲存位置）
  DRIVE_FOLDER_ID: 'YOUR_DRIVE_FOLDER_ID', // ← 請在此填入您的 Drive 資料夾 ID

  // KAYON STUDIO 管理員 Email
  ADMIN_EMAIL: 'info@karbonxgaiaentertainment.com',
  BOOKING_EMAIL: 'booking@karbonxgaiaentertainment.com',
  NOREPLY_EMAIL: 'no-reply@karbonxgaiaentertainment.com',

  // 工作表名稱
  SHEET_NAME: '合約紀錄',
  BIND_SHEET_NAME: '綁定紀錄',

  // 是否發送管理員通知
  NOTIFY_ADMIN: true,

  // 是否發送學員確認信
  NOTIFY_STUDENT: true,
};

// ─── CORS Headers ─────────────────────────────────────────────
function setCORSHeaders(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader('Access-Control-Allow-Origin', '*')
    .addHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
    .addHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ─── OPTIONS 預檢請求 ──────────────────────────────────────────
function doOptions(e) {
  return setCORSHeaders(ContentService.createTextOutput(''));
}

// ─── GET 請求處理 ──────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'ping') {
    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({
      status: 'ok',
      service: 'KAYON STUDIO Contract API',
      version: '2.1.0',
      timestamp: new Date().toISOString()
    })));
  }

  if (action === 'checkBind') {
    const uid = e.parameter.uid;
    const boundData = checkBindStatus(uid);
    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(boundData)));
  }

  return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Invalid action' })));
}

// ─── POST 請求處理 ─────────────────────────────────────────────
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    // 嘗試取得鎖定，最多等待 30 秒
    if (!lock.tryLock(30000)) {
      return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ success: false, message: '系統繁忙，請稍後再試' })));
    }

    const postData = JSON.parse(e.postData.contents);
    const action = postData.action || 'contract';

    if (action === 'bind') {
      const result = handleBind(postData);
      return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(result)));
    }

    // 預設處理合約簽署
    const lastRow = saveToSheet(postData);
    const driveUrl = saveToDrive(postData);
    
    // 如果有信箱且設定要通知
    if (CONFIG.NOTIFY_STUDENT && postData.email) {
      sendStudentConfirmation(postData, driveUrl);
    }
    
    if (CONFIG.NOTIFY_ADMIN) {
      sendAdminNotification(postData, driveUrl);
    }

    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ 
      success: true, 
      row: lastRow,
      driveUrl: driveUrl
    })));

  } catch (err) {
    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ 
      success: false, 
      error: err.toString() 
    })));
  } finally {
    lock.releaseLock();
  }
}

// ─── 檢查綁定狀態 ──────────────────────────────────────────────
function checkBindStatus(uid) {
  if (!uid || uid === 'unknown') return { bound: false };
  
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.BIND_SHEET_NAME);
  if (!sheet) return { bound: false };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === uid) {
      return { 
        bound: true, 
        name: data[i][1], 
        email: data[i][2],
        phone: data[i][3]
      };
    }
  }
  return { bound: false };
}

// ─── 處理 LINE 綁定 ────────────────────────────────────────────
function handleBind(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.BIND_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.BIND_SHEET_NAME);
    const headers = ['LINE UID', '姓名', '電子信箱', '聯絡電話', '顯示名稱', '頭像', '綁定時間'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#C9A84C').setFontColor('#000').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const rows = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.uid) {
      targetRow = i + 1;
      break;
    }
  }

  const rowData = [
    data.uid,
    data.name,
    data.email,
    data.phone,
    data.lineDisplayName,
    data.linePicture,
    new Date()
  ];

  if (targetRow > 0) {
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return { success: true };
}

// ─── 儲存至 Google Sheet ──────────────────────────────────────
function saveToSheet(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    const headers = [
      '合約編號', '簽署時間', '學員姓名', '課程方案',
      '聯絡電話', '電子信箱', 'LINE UID', 'LINE 顯示名稱',
      '法定代理人', '關係', '法代電話', '法代信箱',
      '身份文件', 'Drive 連結', '狀態'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#C9A84C').setFontColor('#000').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const row = [
    data.contractNo || '—',
    new Date(),
    data.name || '—',
    data.course || '—',
    data.phone || '—',
    data.email || '—',
    data.lineUid || '—',
    data.lineDisplayName || '—',
    data.guardianName || '—',
    data.guardianRel || '—',
    data.guardianPhone || '—',
    data.guardianEmail || '—',
    data.guardianIdUploaded ? '已上傳 ✓' : '未上傳',
    '處理中...',
    '已簽署'
  ];

  sheet.appendRow(row);
  return sheet.getLastRow();
}

// ─── 儲存合約至 Google Drive ──────────────────────────────────
function saveToDrive(data) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
    
    // 建立合約文字內容 (之後可擴充為 PDF)
    const content = `KAYON STUDIO 課程合約\n編號：${data.contractNo}\n姓名：${data.name}\n課程：${data.course}\n時間：${dateStr}`;
    const file = folder.createFile(`${data.contractNo}_${data.name}.txt`, content);
    
    // 如果有上傳簽名或圖片，這裡可以處理存檔
    // ...
    
    return file.getUrl();
  } catch (err) {
    return '儲存失敗：' + err.toString();
  }
}

// ─── 發送通知信 ──────────────────────────────────────────────
function sendStudentConfirmation(data, driveUrl) {
  const subject = `【KAYON STUDIO】課程合約簽署完成通知 - ${data.name}`;
  const body = `親愛的 ${data.name} 您好：\n\n感謝您簽署課程合約。\n\n課程方案：${data.course}\n合約編號：${data.contractNo}\n合約連結：${driveUrl}\n\n如有任何問題，歡迎透過 LINE 官方帳號聯繫我們。`;
  
  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    body: body,
    name: 'KAYON STUDIO'
  });
}

function sendAdminNotification(data, driveUrl) {
  const subject = `[通知] 新合約簽署：${data.name} - ${data.course}`;
  const body = `管理員您好，有新的學員完成合約簽署：\n\n姓名：${data.name}\n電話：${data.phone}\n課程：${data.course}\n合約連結：${driveUrl}`;
  
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: subject,
    body: body,
    name: 'KAYON STUDIO 系統通知'
  });
}
