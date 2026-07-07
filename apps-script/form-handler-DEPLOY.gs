// 社團法人台灣人車公益協會 — Form Handler（全選貼上覆蓋舊版 → 存檔 → 執行 setupBeachSheet → 管理部署 → 編輯 → 新版本 → 部署）
const SHEET_ID = '1i3ZnaKGuYpazZHXhZ_KF6-78UcyKp3Euts4u3ruUoMs';
const JOIN_SHEET = '入會申請';
const DONATE_SHEET = '捐款記錄';
const BEACH_SHEET = '淨灘活動報名';
const BEACH_HEADERS = ['時間戳記','報名方式','服務單位','姓名','職稱','行動電話','E-mail','參加人數','S','M','L','XL','2XL','3XL','件數合計','匯款金額','匯款末五碼','狀態'];

// 各尺寸 T恤庫存上限（共 113 件）
const BEACH_STOCK = { S:1, M:9, L:18, XL:45, '2XL':31, '3XL':15 };

// 統計目前各尺寸已售出（讀 I:N 欄，第 9~14 欄）
function beachSold_(sheet) {
  const sold = { S:0, M:0, L:0, XL:0, '2XL':0, '3XL':0 };
  const last = sheet.getLastRow();
  if (last < 2) return sold;
  const v = sheet.getRange(2, 9, last - 1, 6).getValues();
  for (let i = 0; i < v.length; i++) {
    sold.S    += Number(v[i][0]) || 0;
    sold.M    += Number(v[i][1]) || 0;
    sold.L    += Number(v[i][2]) || 0;
    sold.XL   += Number(v[i][3]) || 0;
    sold['2XL'] += Number(v[i][4]) || 0;
    sold['3XL'] += Number(v[i][5]) || 0;
  }
  return sold;
}

// 各尺寸剩餘可報名數
function beachRemaining_(sheet) {
  const sold = beachSold_(sheet);
  const rem = {};
  for (const k in BEACH_STOCK) rem[k] = BEACH_STOCK[k] - (sold[k] || 0);
  return rem;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

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
      return jsonOut_({ status: 'success' });

    } else if (data.type === 'donate') {
      const sheet = ss.getSheetByName(DONATE_SHEET) || ss.insertSheet(DONATE_SHEET);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['時間戳記','姓名','電話','Email','身分證/統編','通訊地址','捐款金額','捐款方式','指定用途','是否需要收據','收據抬頭','統一編號','是否匿名','匯款後五碼','備註','狀態']);
      }
      sheet.appendRow([timestamp,data.name,data.phone,data.email,data.idNumber||'',data.address||'',data.amount,data.paymentMethod,data.purpose,data.receipt,data.receiptName||'',data.receiptTaxId||'',data.anonymous,data.transferCode,data.note||'','待確認']);
      return jsonOut_({ status: 'success' });

    } else if (data.type === 'beach') {
      const sheet = ss.getSheetByName(BEACH_SHEET) || ss.insertSheet(BEACH_SHEET);
      if (sheet.getLastRow() === 0) sheet.appendRow(BEACH_HEADERS);

      // 用鎖避免兩人同時送出造成超賣
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        const want = {
          S:  Number(data.sizeS)  || 0, M:  Number(data.sizeM)  || 0, L:  Number(data.sizeL)  || 0,
          XL: Number(data.sizeXL) || 0, '2XL': Number(data['size2XL']) || 0, '3XL': Number(data['size3XL']) || 0
        };
        const rem = beachRemaining_(sheet);
        const shortages = [];
        for (const k in want) {
          if (want[k] > rem[k]) shortages.push({ size: k, want: want[k], left: Math.max(0, rem[k]) });
        }
        if (shortages.length) {
          return jsonOut_({ status: 'error', code: 'stock', shortages: shortages, stock: rem });
        }
        sheet.appendRow([
          timestamp, data.regType||'個人', data.unit||'', data.name, data.title||'',
          data.phone, data.email, data.groupCount||1,
          want.S, want.M, want.L, want.XL, want['2XL'], want['3XL'],
          data.shirtTotal||0, data.amount||'', data.transferCode||'', '待確認'
        ]);
      } finally {
        lock.releaseLock();
      }
      return jsonOut_({ status: 'success' });
    }

    return jsonOut_({ status: 'error', message: 'unknown type' });
  } catch(err) {
    return jsonOut_({ status: 'error', message: err.toString() });
  }
}

// GET ?action=stock → 回傳各尺寸剩餘數，供報名頁即時顯示
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'stock') {
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(BEACH_SHEET);
      const rem = sheet ? beachRemaining_(sheet) : Object.assign({}, BEACH_STOCK);
      for (const k in rem) if (rem[k] < 0) rem[k] = 0;
      return jsonOut_({ status: 'ok', stock: rem, capacity: BEACH_STOCK });
    } catch(err) {
      return jsonOut_({ status: 'error', message: err.toString() });
    }
  }
  return jsonOut_({ status: 'ok', message: '人車故事公益協會 Form API' });
}

// 一次性：選此函式按「執行」，建立「淨灘活動報名」分頁、表頭與即時統計面板
// ※ 若分頁已存在但欄位是舊版，請先在試算表刪除該分頁再執行本函式
function setupBeachSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(BEACH_SHEET) || ss.insertSheet(BEACH_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(BEACH_HEADERS);
    sheet.getRange(1,1,1,BEACH_HEADERS.length).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1,150);
    sheet.setColumnWidth(3,150);
    sheet.setColumnWidth(7,200);
  }
  buildBeachStats_(sheet);
}

// 在報名分頁右側（T、U 欄）建立即時統計 + 各尺寸庫存剩餘，公式自動更新
function buildBeachStats_(sheet) {
  const stats = [
    ['📊 報名統計', ''],
    ['報名筆數', '=COUNTA(D2:D)'],
    ['總參加人數', '=SUM(H2:H)'],
    ['T恤總件數', '=SUM(O2:O)'],
    ['應收金額合計', '=SUM(P2:P)'],
    ['', ''],
    ['尺寸統計（已售 / 上限）', ''],
    ['S',   '=SUM(I2:I)&" / "&' + BEACH_STOCK.S],
    ['M',   '=SUM(J2:J)&" / "&' + BEACH_STOCK.M],
    ['L',   '=SUM(K2:K)&" / "&' + BEACH_STOCK.L],
    ['XL',  '=SUM(L2:L)&" / "&' + BEACH_STOCK.XL],
    ['2XL', '=SUM(M2:M)&" / "&' + BEACH_STOCK['2XL']],
    ['3XL', '=SUM(N2:N)&" / "&' + BEACH_STOCK['3XL']],
    ['', ''],
    ['尺寸剩餘（可報名）', ''],
    ['S 剩',   '=' + BEACH_STOCK.S + '-SUM(I2:I)'],
    ['M 剩',   '=' + BEACH_STOCK.M + '-SUM(J2:J)'],
    ['L 剩',   '=' + BEACH_STOCK.L + '-SUM(K2:K)'],
    ['XL 剩',  '=' + BEACH_STOCK.XL + '-SUM(L2:L)'],
    ['2XL 剩', '=' + BEACH_STOCK['2XL'] + '-SUM(M2:M)'],
    ['3XL 剩', '=' + BEACH_STOCK['3XL'] + '-SUM(N2:N)'],
  ];
  sheet.getRange(1, 20, stats.length, 2).setValues(stats);
  sheet.getRange(1, 20, 1, 2).merge().setFontWeight('bold').setFontSize(12)
       .setBackground('#e8a020').setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.getRange(7, 20, 1, 2).merge().setFontWeight('bold')
       .setBackground('#f2f2f2').setHorizontalAlignment('center');
  sheet.getRange(15, 20, 1, 2).merge().setFontWeight('bold')
       .setBackground('#fff3d6').setHorizontalAlignment('center');
  sheet.getRange(1, 20, stats.length, 1).setFontWeight('bold');
  sheet.setColumnWidth(20, 160);
  sheet.setColumnWidth(21, 90);
}
