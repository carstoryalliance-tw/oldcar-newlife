// 社團法人台灣人車公益協會 — Google Apps Script Form Handler
// 部署為 Web App: 執行身分「我」，存取「所有人」
// 建立後取得 URL 填入 HTML 表單的 SCRIPT_URL

const SHEET_ID = ''; // 填入你的 Google Sheets ID (URL 中的那串)
const JOIN_SHEET = '入會申請';
const DONATE_SHEET = '捐款記錄';
const BEACH_SHEET = '淨灘活動報名';

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

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '人車故事公益協會 Form API' }))
    .setMimeType(ContentService.MimeType.JSON);
}
