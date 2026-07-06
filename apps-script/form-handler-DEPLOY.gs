// 社團法人台灣人車公益協會 — Form Handler（全選貼上覆蓋舊版 → 部署 → 管理部署 → 編輯 → 新版本 → 部署）
const SHEET_ID = '1i3ZnaKGuYpazZHXhZ_KF6-78UcyKp3Euts4u3ruUoMs';
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
        sheet.appendRow(['時間戳記','姓名','電話','Email','LINE ID','公司/品牌','職稱/身份','所在縣市','會員類型','IG 帳號','Facebook 粉專','推薦人','付款方式','匯款後五碼','狀態']);
      }
      sheet.appendRow([timestamp,data.name,data.phone,data.email,data.lineId,data.company,data.role,data.city,data.memberType,data.ig||'',data.fb||'',data.referral||'',data.paymentMethod,data.transferCode||'','待審核']);

    } else if (data.type === 'donate') {
      const sheet = ss.getSheetByName(DONATE_SHEET) || ss.insertSheet(DONATE_SHEET);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['時間戳記','姓名','電話','Email','身分證/統編','通訊地址','捐款金額','捐款方式','指定用途','是否需要收據','收據抬頭','統一編號','是否匿名','匯款後五碼','備註','狀態']);
      }
      sheet.appendRow([timestamp,data.name,data.phone,data.email,data.idNumber||'',data.address||'',data.amount,data.paymentMethod,data.purpose,data.receipt,data.receiptName||'',data.receiptTaxId||'',data.anonymous,data.transferCode,data.note||'','待確認']);

    } else if (data.type === 'beach') {
      const sheet = ss.getSheetByName(BEACH_SHEET) || ss.insertSheet(BEACH_SHEET);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['時間戳記','單位','姓名','聯絡電話','E-mail','T恤尺寸','匯款金額','匯款末五碼','狀態']);
      }
      sheet.appendRow([timestamp,data.unit||'',data.name,data.phone,data.email,data.size,data.amount||'',data.transferCode||'','待確認']);
    }

    return ContentService.createTextOutput(JSON.stringify({status:'success'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status:'ok',message:'人車故事公益協會 Form API'})).setMimeType(ContentService.MimeType.JSON);
}

// 一次性：在 Apps Script 編輯器選此函式按「執行」，即可立即建立「淨灘活動報名」分頁、表頭與即時統計面板
function setupBeachSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(BEACH_SHEET) || ss.insertSheet(BEACH_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['時間戳記','單位','姓名','聯絡電話','E-mail','T恤尺寸','匯款金額','匯款末五碼','狀態']);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,150);
    sheet.setColumnWidths(3,3,110);
  }
  buildBeachStats_(sheet);
}

// 在報名分頁右側（K、L 欄）建立即時統計面板，公式會隨新報名自動更新
function buildBeachStats_(sheet) {
  const stats = [
    ['📊 報名統計', ''],
    ['總報名人數', '=COUNTA(C2:C)'],
    ['已填匯款末五碼', '=COUNTIF(H2:H,"?*")'],
    ['應收金額合計', '=SUM(G2:G)'],
    ['', ''],
    ['尺寸統計', ''],
    ['S',   '=COUNTIF(F2:F,"S")'],
    ['M',   '=COUNTIF(F2:F,"M")'],
    ['L',   '=COUNTIF(F2:F,"L")'],
    ['XL',  '=COUNTIF(F2:F,"XL")'],
    ['2XL', '=COUNTIF(F2:F,"2XL")'],
    ['3XL', '=COUNTIF(F2:F,"3XL")'],
  ];
  sheet.getRange(1, 11, stats.length, 2).setValues(stats);
  sheet.getRange(1, 11, 1, 2).merge().setFontWeight('bold').setFontSize(12)
       .setBackground('#e8a020').setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.getRange(6, 11, 1, 2).merge().setFontWeight('bold')
       .setBackground('#f2f2f2').setHorizontalAlignment('center');
  sheet.getRange(1, 11, stats.length, 1).setFontWeight('bold');
  sheet.setColumnWidth(11, 130);
  sheet.setColumnWidth(12, 90);
}
