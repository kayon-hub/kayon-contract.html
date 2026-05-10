/**
 * ═══════════════════════════════════════════════════════════════
 * KAYON STUDIO｜電子合約 Google Apps Script 後端
 * 功能：接收合約資料 → 儲存 Google Sheet → 產生 PDF → 存 Drive → 發 Email
 *
 * 部署方式：
 * 1. 前往 script.google.com → 新增專案
 * 2. 貼上此程式碼
 * 3. 部署 → 新增部署 → 類型選「網頁應用程式」
 * 4. 設定「誰可以存取」為「任何人」
 * 5. 複製部署 URL，貼入前端 GAS_URL 常數
 * ═══════════════════════════════════════════════════════════════
 */

// ─── 設定區 ───────────────────────────────────────────────────
const CONFIG = {
  // Google Sheet ID（從 Sheet URL 取得）
  SHEET_ID: 'YOUR_GOOGLE_SHEET_ID',

  // Google Drive 資料夾 ID（合約儲存位置）
  DRIVE_FOLDER_ID: 'YOUR_DRIVE_FOLDER_ID',

  // KAYON STUDIO 管理員 Email
  ADMIN_EMAIL: 'info@karbonxgaiaentertainment.com',
  BOOKING_EMAIL: 'booking@karbonxgaiaentertainment.com',
  NOREPLY_EMAIL: 'no-reply@karbonxgaiaentertainment.com',

  // 工作表名稱
  SHEET_NAME: '合約紀錄',

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

// ─── 主要 POST 接收器 ─────────────────────────────────────────
function doPost(e) {
  try {
    // 解析 JSON 資料
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      data = e.parameter;
    }

    Logger.log('收到資料：' + JSON.stringify(data));

    // 身份綁定請求
    if (data.action === 'bind') {
      saveBind(data);
      return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ success: true, message: '綁定成功' })));
    }

    // 藍新 Webhook
    if (data.Status !== undefined || data.TradeNo !== undefined) {
      handleNewebpayWebhook(data);
      return setCORSHeaders(ContentService.createTextOutput('OK'));
    }

    // 1. 儲存至 Google Sheet
    const rowData = saveToSheet(data);

    // 2. 產生合約文字記錄並儲存至 Drive
    const driveUrl = saveToDrive(data);

    // 3. 發送 Email 通知
    if (CONFIG.NOTIFY_ADMIN) {
      sendAdminNotification(data, driveUrl);
    }

    if (CONFIG.NOTIFY_STUDENT && data.email) {
      sendStudentConfirmation(data, driveUrl);
    }

    // 回傳成功
    const result = {
      success: true,
      contractNo: data.contractNo,
      driveUrl: driveUrl,
      timestamp: new Date().toISOString(),
      message: '合約已成功儲存'
    };

    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify(result))
    );

  } catch (err) {
    Logger.log('錯誤：' + err.toString());
    const error = {
      success: false,
      error: err.toString(),
      timestamp: new Date().toISOString()
    };
    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify(error))
    );
  }
}

// ─── GET 查詢（供測試用）─────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || 'ping';

  if (action === 'ping') {
    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify({
        status: 'ok',
        service: 'KAYON STUDIO Contract API',
        version: '2.0.0',
        timestamp: new Date().toISOString()
      }))
    );
  }

  if (action === 'checkBind' && e.parameter.uid) {
    const result = checkBind(e.parameter.uid);
    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify(result))
    );
  }

  if (action === 'query' && e.parameter.contractNo) {
    const record = queryContract(e.parameter.contractNo);
    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify(record))
    );
  }

  return setCORSHeaders(
    ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
  );
}

// ─── 儲存至 Google Sheet ──────────────────────────────────────
function saveToSheet(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // 若工作表不存在，建立並加入標題列
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    const headers = [
      '合約編號', '簽署時間', '學員姓名', '課程方案',
      '聯絡電話', '電子信箱', 'LINE UID', 'LINE 顯示名稱',
      '法定代理人', '關係', '法代電話（主）', '法代電話（備）', '法代信箱',
      '身份文件上傳', 'Drive 連結', '狀態', '備註'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // 格式化標題列
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#C9A84C');
    headerRange.setFontColor('#000000');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    sheet.setFrozenRows(1);
  }

  const now = new Date();
  const row = [
    data.contractNo || '—',
    now.toLocaleString('zh-TW'),
    data.name || '—',
    data.course || '—',
    data.phone || '—',
    data.email || '—',
    data.lineUid || '—',
    data.lineDisplayName || '—',
    data.guardianName || '—',
    data.guardianRel || '—',
    data.guardianPhone || '—',
    data.guardianPhone2 || '—',
    data.guardianEmail || '—',
    data.guardianIdUploaded ? '已上傳 ✓' : '未上傳',
    '產生中...', // Drive URL 稍後更新
    '已簽署',
    ''
  ];

  const lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);

  // 交替列背景色
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, row.length).setBackground('#1A1A1A');
  }

  Logger.log(`儲存至 Sheet 第 ${lastRow} 列`);
  return lastRow;
}

// ─── 更新 Sheet 中的 Drive URL ────────────────────────────────
function updateSheetDriveUrl(contractNo, driveUrl) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === contractNo) {
      sheet.getRange(i + 1, 13).setValue(driveUrl); // 欄 M = Drive 連結
      break;
    }
  }
}

// ─── 儲存合約至 Google Drive ──────────────────────────────────
function saveToDrive(data) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');

    // 建立合約文字內容
    const content = buildContractText(data, now);
    const fileName = `${data.contractNo}_${data.name}_${dateStr}.txt`;

    // 儲存文字版合約
    const file = folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
    file.setDescription(`KAYON STUDIO 課程合約 | ${data.name} | ${data.course} | ${dateStr}`);

    const fileUrl = file.getUrl();
    Logger.log(`Drive 儲存成功：${fileUrl}`);

    // 更新 Sheet 中的 Drive URL
    if (data.contractNo) {
      updateSheetDriveUrl(data.contractNo, fileUrl);
    }

    return fileUrl;

  } catch (err) {
    Logger.log('Drive 儲存錯誤：' + err.toString());
    return '儲存失敗：' + err.toString();
  }
}

// ─── 建立合約文字內容 ─────────────────────────────────────────
function buildContractText(data, signTime) {
  const line = '═'.repeat(60);
  const dash = '─'.repeat(60);

  return `
${line}
KAYON STUDIO｜KARBØN × GAIA ENTERTAINMENT
線上音樂創作課程合約書（電子版）
${line}

【合約基本資訊】
合約編號：${data.contractNo || '—'}
簽署時間：${Utilities.formatDate(signTime, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss')} (台灣時間)
課程方案：${data.course || '—'}

${dash}
【學員資訊】
姓名：${data.name || '—'}
聯絡電話：${data.phone || '—'}
電子信箱：${data.email || '—'}
LINE UID：${data.lineUid || '—'}
LINE 顯示名稱：${data.lineDisplayName || '—'}

${dash}
【法定代理人資訊（未成年適用）】
姓名：${data.guardianName || '無（成年學員）'}
關係：${data.guardianRel || '—'}
聯絡電話（主）：${data.guardianPhone || '—'}
聯絡電話（備）：${data.guardianPhone2 || '—'}
電子信箱：${data.guardianEmail || '—'}
身份證明文件：${data.guardianIdUploaded ? '已上傳 ✓' : '未上傳'}

${dash}
【合約條款摘要】

§1 付款規範
・課程方案與付款方式依雙方書面確認內容為準
・優惠方案不得同時並行
・完成付款後始完成課程預約

§2 課程進行規範
・使用 Google Meet 線上授課
・課程時間以 Google Calendar 為準
・遲到超過10分鐘視同缺席
・課程不得自行錄音錄影
・講師提供講義及數位教材

§3 改期與取消規定
・改期須透過 LINE 系統提前24小時操作
・堂數由系統自動管理
・當月未用堂數不累積
・課程開始後不提供退款

§4 智慧財產權
・課程內容屬 KAYON STUDIO 智慧財產
・未授權不得外流轉載
・學員作品著作權歸學員所有

§5 系統功能
・LINE 官方帳號預約管理
・Google Calendar 同步
・堂數系統自動管理

${dash}
【電子簽名法律聲明】

本合約依據中華民國《電子簽章法》第4條：
「依本法規定得以電子文件為意思表示者，其效力與書面文件相同。」

學員已於系統完成電子簽名，簽署時間戳記由系統自動記錄，
具備與紙本合約相同之完整法律效力。

${dash}
【KAYON STUDIO 聯絡資訊】
官方網站：https://karbonxgaiaentertainment.com
LINE 官方：https://liff.line.me/2009971827-k0iEMkIu?view=register
Email：kayon@karbonxgaiaentertainment.com

${line}
本文件由 KAYON STUDIO 電子合約系統自動產生
Generated at：${new Date().toISOString()}
${line}
`;
}

// ─── 發送管理員通知 ───────────────────────────────────────────
function sendAdminNotification(data, driveUrl) {
  const subject = `[合約簽署] ${data.contractNo} | ${data.name} | ${data.course}`;
  const body = `
KAYON STUDIO 電子合約系統通知

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
新合約已完成簽署
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

合約編號：${data.contractNo}
學員姓名：${data.name}
課程方案：${data.course}
聯絡電話：${data.phone}
電子信箱：${data.email}
LINE UID：${data.lineUid}
簽署時間：${data.time}

法定代理人：${data.guardianName || '無（成年學員）'}

Drive 合約連結：
${driveUrl}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
KAYON STUDIO 電子合約系統
  `.trim();

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, body, {
    cc: CONFIG.BOOKING_EMAIL,
    name: 'KAYON STUDIO 系統通知'
  });
  Logger.log('管理員通知已發送至：' + CONFIG.ADMIN_EMAIL);
}

// ─── 發送學員確認信 ───────────────────────────────────────────
function sendStudentConfirmation(data, driveUrl) {
  if (!data.email || !data.email.includes('@')) return;

  const subject = `【KAYON STUDIO】課程合約簽署確認｜${data.contractNo}`;
  const body = `
親愛的 ${data.name} 你好，

感謝您選擇 KAYON STUDIO！您的課程合約已成功完成電子簽署。

━━━━━━━━━━━━━━━━━━━━━━━━━
合約資訊
━━━━━━━━━━━━━━━━━━━━━━━━━
合約編號：${data.contractNo}
課程方案：${data.course}
簽署時間：${data.time}

合約電子版連結（僅供查閱）：
${driveUrl}

━━━━━━━━━━━━━━━━━━━━━━━━━
後續步驟
━━━━━━━━━━━━━━━━━━━━━━━━━
1. 請前往 LINE 官方帳號完成首堂課程預約
2. 講師將於 1-2 個工作天內與您確認課程時間
3. 課程前請確認 Google Meet 連結（將於課前24小時發送）

LINE 官方帳號：
https://liff.line.me/2009971827-k0iEMkIu?view=register

如有任何問題，歡迎透過 LINE 或 Email 聯繫：
kayon@karbonxgaiaentertainment.com

期待與你一起創作！🎵

KAYON
KAYON STUDIO｜KARBØN × GAIA ENTERTAINMENT
https://karbonxgaiaentertainment.com

---
This email was sent automatically by the system.
Please do not reply directly to this email.
如需協助請聯絡：booking@karbonxgaiaentertainment.com
  `.trim();

  GmailApp.sendEmail(data.email, subject, body, {
    from: CONFIG.NOREPLY_EMAIL,
    replyTo: CONFIG.BOOKING_EMAIL,
    name: 'KAYON STUDIO'
  });
  Logger.log('學員確認信已發送至：' + data.email);
}

// ─── 查詢合約（GET action=query）────────────────────────────
function queryContract(contractNo) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return { error: '工作表不存在' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === contractNo) {
        const record = {};
        headers.forEach((h, j) => record[h] = data[i][j]);
        return { found: true, record };
      }
    }

    return { found: false, contractNo };

  } catch (err) {
    return { error: err.toString() };
  }
}

// ─── 每月自動備份（可設定觸發器）────────────────────────────
function monthlyBackup() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

    const now = new Date();
    const monthStr = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM');
    const backupName = `KAYON_合約備份_${monthStr}`;

    // 匯出為 CSV
    const data = sheet.getDataRange().getValues();
    const csv = data.map(row =>
      row.map(cell => `"${String(cell).replace(/"/g,'""')}"`).join(',')
    ).join('\n');

    folder.createFile(`${backupName}.csv`, csv, MimeType.CSV);
    Logger.log(`每月備份完成：${backupName}`);

  } catch (err) {
    Logger.log('備份錯誤：' + err.toString());
  }
}

/**
 * ═══════════════════════════════════════════════════════════════
 * 安裝觸發器指令（在 Apps Script 編輯器執行一次即可）
 * ═══════════════════════════════════════════════════════════════
 *
 * function installTriggers() {
 *   // 每月1日凌晨2點執行備份
 *   ScriptApp.newTrigger('monthlyBackup')
 *     .timeBased()
 *     .onMonthDay(1)
 *     .atHour(2)
 *     .create();
 * }
 */

/* ════════════════════════════════════════
   學員綁定系統
════════════════════════════════════════ */

// 綁定工作表名稱
const BIND_SHEET = '學員綁定';

// LINE Messaging API
const LINE_CHANNEL_TOKEN = '+wzpaFedPdOffcJWLneo/h7SFPA6+Ohj+zo79KURh4fqMJ0KNM4PlEtpVLIwYIIA+XpNCU8N+pKbfq0bJuTr67Ntas00rD1d5gCEsuH45TQz0aes/BRfTj5skHTgjnhJXjRQ//jZ0JDdq2y8Cj4iHAdB04t89/1O/w1cDnyilFU='; // 已更新

/* ── 建立綁定工作表 ── */
function getBindSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(BIND_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BIND_SHEET);
    const headers = ['綁定時間','LINE UID','LINE 名稱','真實姓名','電子信箱','聯絡電話','付款狀態','合約狀態'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    const hr = sheet.getRange(1,1,1,headers.length);
    hr.setBackground('#C9A84C');hr.setFontColor('#000');hr.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ── 檢查是否已綁定 ── */
function checkBind(uid) {
  const sheet = getBindSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === uid) {
      return { bound: true, name: data[i][3], email: data[i][4] };
    }
  }
  return { bound: false };
}

/* ── 寫入綁定資料 ── */
function saveBind(data) {
  const sheet = getBindSheet();
  const row = [
    data.bindTime || new Date().toLocaleString('zh-TW'),
    data.uid, data.lineDisplayName, data.name,
    data.email, data.phone, '待付款', '待簽署'
  ];
  sheet.appendRow(row);
  Logger.log('學員綁定成功：' + data.name + ' / ' + data.email);
}

/* ── 藍新 Webhook 處理 ── */
function handleNewebpayWebhook(data) {
  Logger.log('收到藍新 webhook：' + JSON.stringify(data));

  const email = data.Email || data.RespondCode || '';
  const amount = data.Amt || data.TradeAmt || '';
  const tradeNo = data.MerchantOrderNo || data.TradeNo || '';
  const status = data.Status || '';

  if (status !== 'SUCCESS') {
    Logger.log('付款未成功，略過');
    return;
  }

  // 比對 email 找 LINE UID
  const sheet = getBindSheet();
  const sheetData = sheet.getDataRange().getValues();
  let foundRow = -1;
  let lineUid = null;
  let studentName = '';

  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][4].toLowerCase() === email.toLowerCase()) {
      foundRow = i + 1;
      lineUid = sheetData[i][1];
      studentName = sheetData[i][3];
      break;
    }
  }

  if (!lineUid) {
    Logger.log('找不到對應的 LINE UID，email：' + email);
    // 發管理員通知
    GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, '⚠️ 付款成功但找不到綁定學員',
      `付款信箱：${email}
金額：${amount}
訂單號：${tradeNo}
請手動確認並發送合約。`);
    return;
  }

  // 更新付款狀態
  if (foundRow > 0) {
    sheet.getRange(foundRow, 7).setValue('已付款 ✓');
  }

  // 發 LINE 推播
  sendLineMessage(lineUid, studentName, tradeNo, amount);

  // 寫入合約紀錄
  const contractNo = 'KS-' + new Date().getFullYear() + '-' + String(Math.floor(Math.random()*900+100)).padStart(3,'0');
  Logger.log('自動發送合約給：' + studentName + ' / LINE UID：' + lineUid);
}

/* ── LINE Messaging API 推播 ── */
function sendLineMessage(uid, name, tradeNo, amount) {
  const contractUrl = 'https://kayon-contract.karbonxgaiaentertainment.com';

  const message = {
    to: uid,
    messages: [
      {
        type: 'flex',
        altText: '✅ 付款成功！請點此完成合約簽署',
        contents: {
          type: 'bubble',
          size: 'kilo',
          header: {
            type: 'box',
            layout: 'vertical',
            contents: [{
              type: 'text',
              text: 'KAYONSTUDIO｜CLASS',
              color: '#C9A84C',
              size: 'xs',
              weight: 'bold',
              align: 'center'
            }],
            backgroundColor: '#0A0A0A',
            paddingAll: '12px'
          },
          body: {
            type: 'box',
            layout: 'vertical',
            spacing: 'sm',
            contents: [
              { type: 'text', text: '✅ 付款確認成功', weight: 'bold', size: 'lg', color: '#F0EDE6' },
              { type: 'text', text: `感謝 ${name} 的報名！`, size: 'sm', color: '#888888', margin: 'xs' },
              { type: 'separator', margin: 'md', color: '#252525' },
              {
                type: 'box', layout: 'vertical', margin: 'md', spacing: 'xs',
                contents: [
                  { type: 'box', layout: 'horizontal', contents: [
                    { type: 'text', text: '訂單號碼', size: 'xs', color: '#888888', flex: 2 },
                    { type: 'text', text: tradeNo || '—', size: 'xs', color: '#F0EDE6', flex: 3, align: 'end' }
                  ]},
                  { type: 'box', layout: 'horizontal', contents: [
                    { type: 'text', text: '付款金額', size: 'xs', color: '#888888', flex: 2 },
                    { type: 'text', text: amount ? '$' + amount : '—', size: 'xs', color: '#C9A84C', flex: 3, align: 'end', weight: 'bold' }
                  ]}
                ]
              },
              { type: 'separator', margin: 'md', color: '#252525' },
              { type: 'text', text: '請點下方按鈕完成電子合約簽署', size: 'xs', color: '#888888', margin: 'md', wrap: true }
            ],
            backgroundColor: '#111111',
            paddingAll: '16px'
          },
          footer: {
            type: 'box',
            layout: 'vertical',
            contents: [{
              type: 'button',
              action: { type: 'uri', label: '📋 立即簽署合約', uri: contractUrl },
              style: 'primary',
              color: '#C9A84C',
              height: 'sm'
            }],
            backgroundColor: '#111111',
            paddingAll: '12px'
          }
        }
      }
    ]
  };

  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN
    },
    payload: JSON.stringify(message),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
  Logger.log('LINE 推播結果：' + res.getResponseCode() + ' / ' + res.getContentText());
}
/* ════════════════════════════════════════ */
