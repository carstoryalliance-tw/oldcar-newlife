// 社團法人台灣人車公益協會 — Google Apps Script Form Handler
// 部署為 Web App: 執行身分「我」，存取「所有人」
// 建立後取得 URL 填入 HTML 表單的 SCRIPT_URL

const SHEET_ID = '1i3ZnaKGuYpazZHXhZ_KF6-78UcyKp3Euts4u3ruUoMs'; // 填入你的 Google Sheets ID (URL 中的那串)
const JOIN_SHEET = '入會申請';
const DONATE_SHEET = '捐款記錄';
const BEACH_SHEET = '淨灘活動報名';
const AID_SHEET = '公益申請';      // 公益申請專用分頁（自動建立，與入會/捐款分開）

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
function getAidFolder_() {
  if (AID_FOLDER_ID) return DriveApp.getFolderById(AID_FOLDER_ID);
  const it = DriveApp.getFoldersByName(AID_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(AID_FOLDER_NAME);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '人車故事公益協會 Form API' }))
    .setMimeType(ContentService.MimeType.JSON);
}
