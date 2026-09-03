// 社團法人台灣人車公益協會 — Google Apps Script Form Handler
// 部署為 Web App: 執行身分「我」，存取「所有人」
// 建立後取得 URL 填入 HTML 表單的 SCRIPT_URL

const SHEET_ID = '1i3ZnaKGuYpazZHXhZ_KF6-78UcyKp3Euts4u3ruUoMs'; // 填入你的 Google Sheets ID (URL 中的那串)
const JOIN_SHEET = '入會申請';
const DONATE_SHEET = '捐款記錄';
const BEACH_SHEET = '淨灘活動報名';
const AID_SHEET = '公益申請';      // 公益申請專用分頁（自動建立，與入會/捐款分開）
const MEETING_SHEET = '例會報名';   // 例會出席報名（自動建立）
const FUND_SHEET = '孤兒院募資';    // 孤兒院用車整理專案募資（自動建立）
const FUND_GOAL = 200000;          // 募資目標金額
const FUND_OK = '已通過';
const FUND_ACT_SHEET = '募資互動';   // 分享／集氣次數（自動建立）

const MAIL_FROM = '社團法人台灣人車公益協會 <no-reply@oldcarnewlife.org.tw>';
const MAIL_FROM_FALLBACK = '社團法人台灣人車公益協會';   // 用 Gmail 寄時只能改顯示名稱
const MAIL_ADMIN = ['carstory.alliance@gmail.com', 'soulbreakin@gmail.com'];   // 收通知的信箱，可多個
const MAIL_REPLY_TO = 'carstory.alliance@gmail.com';
           // 審核狀態填這三個字，進度條才會計入

// 照片/影片上傳存放位置。留空 = 自動在雲端根目錄建立同名資料夾。
const AID_FOLDER_ID = '';
const AID_FOLDER_NAME = '人車故事公益申請_上傳';
// 審閱者信箱（會把每筆申請的上傳子資料夾分享給這些人檢視）。留空 = 檔案僅擁有者可見。
const AID_REVIEWER_EMAILS = []; // 例：['carstory.alliance@gmail.com', 'someone@gmail.com']

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    if (data.type === 'join') {
      const sheet = ss.getSheetByName(JOIN_SHEET) || ss.insertSheet(JOIN_SHEET);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          '時間戳記', '姓名', '電話', 'Email', 'LINE ID',
          '公司/品牌', '職稱/身份', '所在縣市', '會員類型',
          'IG 帳號', 'Facebook 粉專', '推薦人',
          '付款方式', '匯款後五碼', '狀態'
        ]);
      }
      sheet.appendRow([
        timestamp,
        data.name, data.phone, data.email, data.lineId,
        data.company, data.role, data.city, data.memberType,
        data.ig || '', data.fb || '', data.referral || '',
        data.paymentMethod, data.transferCode || '',
        '待審核'
      ]);

    } else if (data.type === 'donate') {
      const sheet = ss.getSheetByName(DONATE_SHEET) || ss.insertSheet(DONATE_SHEET);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          '時間戳記', '姓名', '電話', 'Email',
          '身分證/統編', '通訊地址', '捐款金額',
          '捐款方式', '指定用途', '是否需要收據',
          '收據抬頭', '統一編號', '是否匿名', '匯款後五碼', '備註', '狀態'
        ]);
      }
      sheet.appendRow([
        timestamp,
        data.name, data.phone, data.email,
        data.idNumber || '', data.address || '', data.amount,
        data.paymentMethod, data.purpose, data.receipt,
        data.receiptName || '', data.receiptTaxId || '', data.anonymous, data.transferCode,
        data.note || '', '待確認'
      ]);

    } else if (data.type === 'beach') {
      const sheet = ss.getSheetByName(BEACH_SHEET) || ss.insertSheet(BEACH_SHEET);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          '時間戳記','報名方式','服務單位','姓名','職稱','行動電話','E-mail','參加人數',
          'S','M','L','XL','2XL','3XL','件數合計','匯款金額','匯款末五碼','狀態'
        ]);
      }
      sheet.appendRow([
        timestamp, data.regType || '個人', data.unit || '', data.name, data.title || '',
        data.phone, data.email, data.groupCount || 1,
        data.sizeS || 0, data.sizeM || 0, data.sizeL || 0, data.sizeXL || 0, data['size2XL'] || 0, data['size3XL'] || 0,
        data.shirtTotal || 0, data.amount || '', data.transferCode || '', '待確認'
      ]);

    } else if (data.type === 'meeting') {
      // 協會例會出席報名
      const sheet = ss.getSheetByName(MEETING_SHEET) || ss.insertSheet(MEETING_SHEET);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          '時間戳記', '場次', '姓名', '公司/品牌', '職稱', '電話', 'Email',
          '身份', '出席人數', '餐點需求', '匯款金額', '匯款末五碼', '備註', '狀態'
        ]);
      }
      // 電話欄固定純文字，否則 0900… 會被當成數字吃掉開頭的 0
      // （只設格式不夠：appendRow 仍會把純數字字串轉成數值，所以再加 ' 前綴強制文字）
      sheet.getRange('F:F').setNumberFormat('@');
      const phoneTxt = data.phone ? "'" + String(data.phone).trim() : '';
      sheet.appendRow([
        timestamp, data.session || '', data.name || '', data.company || '', data.title || '',
        phoneTxt, data.email || '', data.identity || '', data.count || 1,
        data.meal || '', data.amount || '', data.transferCode ? "'" + String(data.transferCode).trim() : '',
        data.note || '', '待確認匯款'
      ]);

    } else if (data.type === 'fund') {
      // 孤兒院用車整理專案 · 募資捐款
      const sheet = ss.getSheetByName(FUND_SHEET) || ss.insertSheet(FUND_SHEET);
      ensureFundHeader_(sheet);
      sheet.getRange('D:D').setNumberFormat('@');
      sheet.getRange('G:G').setNumberFormat('@');
      const amount = Number(String(data.amount || '').replace(/[^0-9.]/g, '')) || 0;
      sheet.appendRow([
        timestamp, data.name || '', data.company || '',
        data.phone ? "'" + String(data.phone).trim() : '',
        data.email || '', amount,
        data.transferCode ? "'" + String(data.transferCode).trim() : '',
        data.display || '', data.displayName || '', data.message || '',
        data.receipt || '', data.receiptTitle || '', data.taxId || '',
        data.note || '', '待審核'
      ]);
      try { fundMails_(data, amount, timestamp); } catch (e) { console.warn('[mail] ' + e); }

    } else if (data.type === 'fundact') {
      // 募資頁的「分享 / 集氣」按鈕計數
      const act = String(data.act || '').trim();
      if (act === 'share' || act === 'cheer') {
        const sheet = ss.getSheetByName(FUND_ACT_SHEET) || ss.insertSheet(FUND_ACT_SHEET);
        if (sheet.getLastRow() === 0) sheet.appendRow(['時間戳記', '動作', '來源']);
        sheet.appendRow([timestamp, act, data.from || '']);
      }

    } else if (data.type === 'aid') {
      // 人車故事公益計畫申請（維修/保養/翻新/送車）— 專用分頁，自動建立
      const sheet = ss.getSheetByName(AID_SHEET) || ss.insertSheet(AID_SHEET);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          '時間戳記', '申請類型', '姓名', '電話', 'Email', 'LINE ID', '所在縣市',
          '車輛廠牌/年份', '車況描述', '人車故事', '家庭/經濟狀況',
          '是否同意公開故事', '照片/影片連結', '狀態'
        ]);
      }

      // 上傳照片/影片到雲端，收集可點連結
      const fileEntries = []; // { label, url }
      if (data.files && data.files.length) {
        const folder = getAidFolder_();
        const subName = timestamp.replace(/[\/:]/g, '-') + '_' + (data.name || '申請者');
        const sub = folder.createFolder(subName);
        if (AID_REVIEWER_EMAILS && AID_REVIEWER_EMAILS.length) {
          try { sub.addViewers(AID_REVIEWER_EMAILS); } catch (x) {}
        }
        data.files.forEach(function (f, i) {
          try {
            const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType || 'application/octet-stream', f.name || 'file');
            const file = sub.createFile(blob);
            fileEntries.push({ label: (f.name || ('檔案' + (i + 1))), url: file.getUrl() });
          } catch (x) {
            fileEntries.push({ label: '(檔案處理失敗: ' + (f.name || '') + ')', url: '' });
          }
        });
      }

      // 純文字後備（RichText 建立失敗時仍看得到連結）
      const linksText = fileEntries.map(function (e) { return e.url ? (e.label + ' → ' + e.url) : e.label; }).join('\n');

      sheet.appendRow([
        timestamp, data.category || '', data.name || '', data.phone || '', data.email || '',
        data.lineId || '', data.city || '', data.vehicle || '', data.condition || '',
        data.story || '', data.situation || '', data.publicConsent || '', linksText, '待審核'
      ]);

      // 把「照片/影片連結」欄(第13欄)改成可點的超連結：每個檔名各自連到自己的檔案
      if (fileEntries.length) {
        try {
          const row = sheet.getLastRow();
          const display = fileEntries.map(function (e) { return e.label; }).join('\n');
          const rt = SpreadsheetApp.newRichTextValue().setText(display);
          let offset = 0;
          fileEntries.forEach(function (e) {
            const start = offset, end = offset + e.label.length;
            if (e.url) rt.setLinkUrl(start, end, e.url);
            offset = end + 1; // +1 換行字元
          });
          sheet.getRange(row, 13).setRichTextValue(rt.build());
        } catch (x) { /* 保留純文字後備 */ }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 取得（或建立）上傳資料夾
// 募資分頁表頭：個人收據要身分證字號、公司要統一編號，欄名改成兩者合併
function ensureFundHeader_(sheet) {
  const HEAD = ['時間戳記', '姓名', '公司/單位', '電話', 'Email', '捐款金額', '匯款末五碼',
    '芳名公開方式', '公開顯示名稱', '想說的話', '收據需求', '收據抬頭',
    '身分證字號／統一編號', '備註', '審核狀態'];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEAD);
    sheet.setFrozenRows(1);
    return;
  }
  // 既有分頁：把舊欄名換掉，資料不動
  const cur = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEAD.length)).getValues()[0];
  const i = cur.indexOf('統一編號');
  if (i >= 0) sheet.getRange(1, i + 1).setValue('身分證字號／統一編號');
}

/* ═══════════════ 寄信 ═══════════════
 * 優先走 Resend（寄件人可用協會網域，信譽較好），
 * 沒設定 RESEND_API_KEY 就退回 Gmail（MailApp），完全不設定也不會讓表單失敗。
 * 金鑰放在「專案設定 → 指令碼屬性」，鍵名 RESEND_API_KEY。
 */
function sendMail_(to, subject, html) {
  if (!to) return;
  const plain = mailToPlain_(html);
  const key = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');

  if (key) {
    try {
      const res = UrlFetchApp.fetch('https://api.resend.com/emails', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + key },
        payload: JSON.stringify({
          from: MAIL_FROM,
          to: Array.isArray(to) ? to : [to],
          subject: subject,
          html: html,
          text: plain,
          reply_to: MAIL_REPLY_TO
        }),
        muteHttpExceptions: true
      });
      if (res.getResponseCode() < 300) return;
      console.warn('[mail] Resend 失敗，改用 Gmail：' + res.getContentText());
    } catch (e) {
      console.warn('[mail] Resend 例外，改用 Gmail：' + e);
    }
  }

  // 後備：用試算表擁有者的 Gmail 寄（每日約 100 封額度）
  try {
    MailApp.sendEmail({
      to: Array.isArray(to) ? to.join(',') : to,
      subject: subject,
      htmlBody: html,
      body: plain,
      name: MAIL_FROM_FALLBACK,
      replyTo: MAIL_REPLY_TO
    });
  } catch (e) {
    console.warn('[mail] 寄信失敗（不影響表單寫入）：' + e);
  }
}

function mailToPlain_(html) {
  return String(html)
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<a [^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi, '$2（$1）')
    .replace(/<\/(p|div|h1|h2|tr|li)>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/\n{3,}/g, '\n\n').trim();
}

/** 信件外框：協會橘 + 白底，用 table 排版才不會在各家信箱跑版 */
function mailShell_(title, bodyHtml, footNote) {
  return '' +
  '<div style="background:#f4f5f7;padding:24px 12px;font-family:\'Noto Sans TC\',\'Microsoft JhengHei\',sans-serif;">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;margin:0 auto;' +
  'background:#fff;border-radius:14px;overflow:hidden;border:1px solid #e3e5ea;">' +
  '<tr><td style="padding:0;line-height:0;">' +
  '<img src="https://oldcarnewlife.org.tw/assets/mail-banner.jpg" width="560" ' +
  'alt="社團法人台灣人車公益協會" style="display:block;width:100%;max-width:560px;height:auto;border:0;">' +
  '</td></tr>' +
  '<tr><td style="padding:26px 24px 8px;font-size:19px;font-weight:900;color:#1a1a2e;">' + title + '</td></tr>' +
  '<tr><td style="padding:0 24px 22px;font-size:14px;line-height:1.9;color:#4a4d54;">' + bodyHtml + '</td></tr>' +
  '<tr><td style="padding:16px 24px;background:#fafbfc;border-top:1px solid #eef0f3;font-size:11.5px;line-height:1.8;color:#8a8e96;">' +
  (footNote || '') +
  '<br>社團法人台灣人車公益協會　·　OLD CAR × NEW LIFE' +
  '<br>內政部核准立案 台內團字第 1140047990 號　·　oldcarnewlife.org.tw' +
  '</td></tr></table></div>';
}

/** 有人送出募資表單：回信給贊助者、通知協會 */
function fundMails_(data, amount, timestamp) {
  const name = data.name || '朋友';
  const money = 'NT$ ' + Number(amount || 0).toLocaleString('en-US');

  if (data.email) {
    const body =
      '<p>' + name + ' 你好，</p>' +
      '<p>我們收到你的贊助資料了，謝謝你願意幫南投神國教會找一台能上山的車。</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:14px 0;' +
      'border:1px solid #eef0f3;border-radius:10px;">' +
      '<tr><td style="padding:10px 14px;color:#8a8e96;font-size:13px;">贊助金額</td>' +
      '<td style="padding:10px 14px;text-align:right;font-weight:900;color:#d97b1e;">' + money + '</td></tr>' +
      '<tr><td style="padding:10px 14px;color:#8a8e96;font-size:13px;border-top:1px solid #f2f3f5;">匯款末五碼</td>' +
      '<td style="padding:10px 14px;text-align:right;border-top:1px solid #f2f3f5;">' + (data.transferCode || '—') + '</td></tr>' +
      '<tr><td style="padding:10px 14px;color:#8a8e96;font-size:13px;border-top:1px solid #f2f3f5;">送出時間</td>' +
      '<td style="padding:10px 14px;text-align:right;border-top:1px solid #f2f3f5;">' + timestamp + '</td></tr>' +
      '</table>' +
      '<p>接下來由協會財務核對入帳，<b>核對完成後你的贊助就會出現在募資頁的進度條與芳名錄</b>。' +
      '需要收據的話，我們會另外跟你聯絡。</p>' +
      '<p><a href="https://oldcarnewlife.org.tw/fund/" style="display:inline-block;background:#d97b1e;color:#fff;' +
      'text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:900;">看目前募資進度</a></p>';
    sendMail_(data.email, '謝謝你的贊助——四驅車專案已收到你的資料',
      mailShell_('我們收到你的贊助資料了', body, '這封信由系統自動發送，回信可直接聯絡協會。'));
  }

  if (MAIL_ADMIN.length) {
    const rows = [
      ['姓名', data.name], ['公司／單位', data.company], ['電話', data.phone],
      ['Email', data.email], ['金額', money], ['匯款末五碼', data.transferCode],
      ['芳名方式', data.display], ['公開名稱', data.displayName],
      ['留言', data.message], ['收據', data.receipt], ['抬頭', data.receiptTitle],
      ['身分證／統編', data.taxId], ['備註', data.note]
    ].filter(function (r) { return r[1]; })
     .map(function (r) {
       return '<tr><td style="padding:7px 12px;color:#8a8e96;font-size:13px;white-space:nowrap;">' + r[0] +
              '</td><td style="padding:7px 12px;font-size:13px;">' + r[1] + '</td></tr>';
     }).join('');
    const body =
      '<p>募資頁有一筆新的贊助資料，<b>狀態是「待審核」</b>，請核對入帳後把試算表的審核狀態改成「已通過」。</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:12px 0;' +
      'border:1px solid #eef0f3;border-radius:10px;">' + rows + '</table>' +
      '<p><a href="https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit" ' +
      'style="display:inline-block;background:#1a1a2e;color:#fff;text-decoration:none;padding:11px 22px;' +
      'border-radius:9px;font-weight:900;">開啟試算表核對</a></p>';
    sendMail_(MAIL_ADMIN, '【待審核】四驅車專案收到贊助 ' + money + '　' + (data.name || ''),
      mailShell_('有一筆新的贊助待核對', body, '審核狀態改成「已通過」後，募資頁進度條就會計入。'));
  }
}

/* ── 測試用：在編輯器選這個函式按「執行」，會跳授權視窗，允許後立刻寄一封測試信 ──
 * 新增寄信功能後一定要跑這一次，否則 Web App 沒有寄信權限，信會靜默寄不出去。
 */
function testMail() {
  const to = 'soulbreakin@gmail.com';
  const body =
    '<p>這是一封測試信。</p>' +
    '<p>如果你看得到這封信，代表 Apps Script 的寄信權限已經開通，' +
    '募資表單送出後就會自動寄「收到你的贊助資料了」給贊助者、' +
    '並通知協會有一筆待核對。</p>' +
    '<p><a href="https://oldcarnewlife.org.tw/fund/" style="display:inline-block;background:#d97b1e;' +
    'color:#fff;text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:900;">看募資頁</a></p>';
  sendMail_(to, '【測試】協會寄信功能已開通', mailShell_('寄信功能測試', body, '這封信由 Apps Script 手動執行寄出。'));
  const key = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');
  Logger.log('已送出，寄送方式：' + (key ? 'Resend' : 'Gmail（MailApp）'));
}

/* ── 把目前「待審核」的贊助整理成一封信寄給財務 ──
 * 在編輯器選 mailPendingToFinance 按執行即可；名稱含「測試」的列會自動略過。
 */
const FINANCE_EMAIL = ['0814kimi@gmail.com'];      // 財務長 吳宗圍
const FINANCE_CC = ['soulbreakin@gmail.com'];      // 一併知會

function mailPendingToFinance() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(FUND_SHEET);
  if (!sh || sh.getLastRow() < 2) { Logger.log('沒有資料'); return; }

  const v = sh.getDataRange().getValues();
  const h = v[0];
  const iTs = h.indexOf('時間戳記'), iName = h.indexOf('姓名'), iComp = h.indexOf('公司/單位');
  const iAmt = h.indexOf('捐款金額'), iCode = h.indexOf('匯款末五碼'), iSt = h.indexOf('審核狀態');
  const iPhone = h.indexOf('電話'), iMail = h.indexOf('Email'), iRcpt = h.indexOf('收據需求');

  const DEAD = ['作廢', '退回', '取消', '無效'];
  let total = 0;
  const rows = [];
  for (let r = 1; r < v.length; r++) {
    const st = String(v[r][iSt] || '').trim();
    if (st === FUND_OK) continue;
    let dead = false;
    for (let k = 0; k < DEAD.length; k++) if (st.indexOf(DEAD[k]) >= 0) dead = true;
    if (dead) continue;
    const nm = String(v[r][iName] || '');
    if (nm.indexOf('測試') >= 0 || nm.indexOf('請刪除') >= 0) continue;   // 略過測試資料
    const amt = Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
    if (amt <= 0) continue;
    total += amt;
    const ts = v[r][iTs];
    const tsTxt = (ts instanceof Date)
      ? Utilities.formatDate(ts, 'Asia/Taipei', 'MM/dd HH:mm')
      : String(ts).replace(/:\d\d$/, '');
    rows.push(
      '<tr>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:13px;white-space:nowrap;">' + tsTxt + '</td>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:13px;">' + nm +
        (v[r][iComp] ? '<br><span style="color:#8a8e96;font-size:11.5px;">' + v[r][iComp] + '</span>' : '') + '</td>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:13px;text-align:right;font-weight:900;color:#d97b1e;white-space:nowrap;">' +
        'NT$ ' + amt.toLocaleString('en-US') + '</td>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:13px;text-align:center;letter-spacing:1px;">' +
        (v[r][iCode] || '—') + '</td>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:11.5px;color:#8a8e96;">' +
        (v[r][iRcpt] || '') + '<br>' + (v[r][iPhone] || '') + '</td>' +
      '</tr>');
  }

  if (!rows.length) {
    Logger.log('目前沒有待審核的贊助');
    return;
  }

  const body =
    '<p>Kimi 你好，</p>' +
    '<p>四驅車專案目前有 <b>' + rows.length + ' 筆</b>贊助等待核對入帳，合計 ' +
    '<b style="color:#d97b1e;">NT$ ' + total.toLocaleString('en-US') + '</b>。</p>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:14px 0;' +
    'border:1px solid #eef0f3;border-radius:10px;border-collapse:separate;">' +
    '<tr style="background:#fafbfc;">' +
    '<th style="padding:9px 10px;text-align:left;font-size:11.5px;color:#8a8e96;">時間</th>' +
    '<th style="padding:9px 10px;text-align:left;font-size:11.5px;color:#8a8e96;">贊助者</th>' +
    '<th style="padding:9px 10px;text-align:right;font-size:11.5px;color:#8a8e96;">金額</th>' +
    '<th style="padding:9px 10px;text-align:center;font-size:11.5px;color:#8a8e96;">末五碼</th>' +
    '<th style="padding:9px 10px;text-align:left;font-size:11.5px;color:#8a8e96;">收據／電話</th></tr>' +
    rows.join('') + '</table>' +
    '<p><b>核對方式：</b>比對永豐帳戶入帳的末五碼與金額，確認無誤後把試算表「孤兒院募資」分頁' +
    '最後一欄「審核狀態」改成 <b>已通過</b>（要一字不差），募資頁的進度條就會把它從灰色轉成橘色。</p>' +
    '<p><a href="https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit" ' +
    'style="display:inline-block;background:#1a1a2e;color:#fff;text-decoration:none;padding:11px 22px;' +
    'border-radius:9px;font-weight:900;">開啟試算表核對</a>　' +
    '<a href="https://oldcarnewlife.org.tw/fund/" style="display:inline-block;background:#d97b1e;color:#fff;' +
    'text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:900;">看募資頁</a></p>';

  const html = mailShell_('四驅車專案：' + rows.length + ' 筆贊助待核對',
    body, '這封信由協會表單系統整理發出，測試資料已自動排除。');
  sendMail_(FINANCE_EMAIL.concat(FINANCE_CC),
    '【待核對】四驅車專案　' + rows.length + ' 筆　NT$ ' + total.toLocaleString('en-US'), html);
  Logger.log('已寄出：' + rows.length + ' 筆，合計 ' + total);
}

/* ── 審核通過自動通知理監事 ──
 * 在試算表把「審核狀態」改成「已通過」時，自動寄信給理事長與常務理事。
 * ⚠️ 要先執行一次 setupTriggers() 安裝觸發器，這個功能才會運作。
 */
const BOARD_EMAIL = [
  'ho2010@kwax.tw',           // 理事長 曾聖凱（Kevin）
  'renewucar@gmail.com'       // 常務理事 吳建勳（阿勳）
];
const BOARD_CC = ['soulbreakin@gmail.com', 'carstory.alliance@gmail.com'];

/** 安裝觸發器：整個專案只需要執行這一次 */
function setupTriggers() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFundApproved') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFundApproved').forSpreadsheet(ss).onEdit().create();
  Logger.log('觸發器已安裝：改「審核狀態」為「已通過」時會自動通知理監事');
}

function onFundApproved(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== FUND_SHEET) return;
    if (String(e.value || '').trim() !== FUND_OK) return;          // 只在改成「已通過」時
    if (String(e.oldValue || '').trim() === FUND_OK) return;       // 本來就通過就不重複寄

    const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const col = e.range.getColumn();
    if (h[col - 1] !== '審核狀態') return;                          // 只認審核狀態那一欄

    const row = sh.getRange(e.range.getRow(), 1, 1, sh.getLastColumn()).getValues()[0];
    const g = function (name) { const i = h.indexOf(name); return i < 0 ? '' : row[i]; };
    const name = String(g('姓名') || '');
    if (name.indexOf('測試') >= 0 || name.indexOf('請刪除') >= 0) return;   // 測試資料不通知

    const amt = Number(String(g('捐款金額') || '').replace(/[^0-9.]/g, '')) || 0;
    const anon = String(g('芳名公開方式') || '').indexOf('匿名') >= 0;
    const shown = anon ? '匿名者' : (String(g('公開顯示名稱') || '').trim() || name);
    const msg = String(g('想說的話') || '').trim();

    // 目前累計
    const v = sh.getDataRange().getValues();
    const iAmt = h.indexOf('捐款金額'), iSt = h.indexOf('審核狀態'), iNm = h.indexOf('姓名');
    let raised = 0, count = 0;
    for (let r = 1; r < v.length; r++) {
      if (String(v[r][iSt] || '').trim() !== FUND_OK) continue;
      const nmR = String(v[r][iNm] || '');
      if (nmR.indexOf('測試') >= 0 || nmR.indexOf('請刪除') >= 0) continue;
      raised += Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
      count++;
    }
    const pct = Math.round(raised / FUND_GOAL * 1000) / 10;
    const left = Math.max(0, FUND_GOAL - raised);
    const nf = function (n) { return 'NT$ ' + Number(n).toLocaleString('en-US'); };

    const body =
      '<p>四驅車專案有一筆贊助<b>確認入帳</b>了。</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:12px 0;' +
      'border:1px solid #eef0f3;border-radius:10px;">' +
      '<tr><td style="padding:10px 14px;color:#8a8e96;font-size:13px;">贊助者</td>' +
      '<td style="padding:10px 14px;text-align:right;font-weight:700;">' + shown + '</td></tr>' +
      '<tr><td style="padding:10px 14px;color:#8a8e96;font-size:13px;border-top:1px solid #f2f3f5;">金額</td>' +
      '<td style="padding:10px 14px;text-align:right;border-top:1px solid #f2f3f5;font-weight:900;color:#d97b1e;">' +
        nf(amt) + '</td></tr>' +
      (msg ? '<tr><td style="padding:10px 14px;color:#8a8e96;font-size:13px;border-top:1px solid #f2f3f5;">留言</td>' +
             '<td style="padding:10px 14px;text-align:right;border-top:1px solid #f2f3f5;font-size:13px;">「' + msg + '」</td></tr>' : '') +
      '</table>' +
      '<div style="background:#fdf7ef;border:1px solid #f0dcc0;border-radius:10px;padding:14px 16px;margin:14px 0;">' +
      '<div style="font-size:13px;color:#8b7d6b;">目前累計</div>' +
      '<div style="font-size:26px;font-weight:900;color:#d97b1e;margin:4px 0;">' + nf(raised) +
      ' <span style="font-size:14px;color:#8b7d6b;font-weight:400;">/ ' + nf(FUND_GOAL) + '（' + pct + '%）</span></div>' +
      '<div style="font-size:13px;color:#4a4d54;">共 ' + count + ' 筆贊助' +
      (left > 0 ? '　·　距離目標還差 <b>' + nf(left) + '</b>' : '　·　<b>已達標</b>') + '</div></div>' +
      '<p><a href="https://oldcarnewlife.org.tw/fund/" style="display:inline-block;background:#d97b1e;color:#fff;' +
      'text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:900;">看募資頁</a></p>';

    sendMail_(BOARD_EMAIL.concat(BOARD_CC),
      '【入帳】四驅車專案　' + nf(amt) + '　累計 ' + nf(raised) + '（' + pct + '%）',
      mailShell_('有一筆贊助確認入帳', body, '審核狀態改成「已通過」時自動發出，募資頁進度條已同步更新。'));
  } catch (err) {
    console.warn('[onFundApproved] ' + err);
  }
}

function getAidFolder_() {
  if (AID_FOLDER_ID) return DriveApp.getFolderById(AID_FOLDER_ID);
  const it = DriveApp.getFoldersByName(AID_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(AID_FOLDER_NAME);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.stat === 'fund') return fundStats_();
  return jsonOut_({ status: 'ok', message: '人車故事公益協會 Form API' });
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 募資進度：只計入「審核狀態 = 已通過」的捐款
function fundStats_() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(FUND_SHEET);
    if (sh) { try { ensureFundHeader_(sh); } catch (e) {} }
    if (!sh || sh.getLastRow() < 2) {
      return jsonOut_({ goal: FUND_GOAL, raised: 0, donors: 0, pending: 0, pendingDonors: 0, shares: 0, cheers: 0, list: [] });
    }
    const v = sh.getDataRange().getValues();
    const h = v[0];
    const iAmt = h.indexOf('捐款金額'), iSt = h.indexOf('審核狀態');
    const iWay = h.indexOf('芳名公開方式'), iName = h.indexOf('公開顯示名稱');
    const iMsg = h.indexOf('想說的話'), iTs = h.indexOf('時間戳記');
    let raised = 0, donors = 0, pending = 0, pendingDonors = 0;
    const list = [];
    const DEAD = ['作廢', '退回', '取消', '無效'];   // 這些不列入待審核
    for (let r = 1; r < v.length; r++) {
      const st = String(v[r][iSt] || '').trim();
      if (st !== FUND_OK) {
        // 尚未審核通過的：計入待審核（灰色段）
        const pa = Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
        let dead = false;
        for (let k = 0; k < DEAD.length; k++) if (st.indexOf(DEAD[k]) >= 0) dead = true;
        if (pa > 0 && !dead) { pending += pa; pendingDonors++; }
        continue;
      }
      const amt = Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
      raised += amt; donors++;
      const way = String(v[r][iWay] || '');
      const anon = way.indexOf('匿名') >= 0;
      list.push({
        // 匿名者只隱藏姓名，金額、留言與時間照常公開
        name: anon ? '匿名者' : (String(v[r][iName] || '').trim() || '匿名者'),
        amount: amt,
        msg: String(v[r][iMsg] || '').trim(),
        ts: String(v[r][iTs] || '')
      });
    }
    let shares = 0, cheers = 0;
    const ash = ss.getSheetByName(FUND_ACT_SHEET);
    if (ash && ash.getLastRow() > 1) {
      const av = ash.getRange(2, 2, ash.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < av.length; i++) {
        const a = String(av[i][0] || '').trim();
        if (a === 'share') shares++;
        else if (a === 'cheer') cheers++;
      }
    }
    return jsonOut_({ goal: FUND_GOAL, raised: raised, donors: donors,
                      pending: pending, pendingDonors: pendingDonors,
                      shares: shares, cheers: cheers, list: list.reverse().slice(0, 60) });
  } catch (err) {
    return jsonOut_({ goal: FUND_GOAL, raised: 0, donors: 0, pending: 0, pendingDonors: 0, shares: 0, cheers: 0, list: [], error: String(err) });
  }
}
