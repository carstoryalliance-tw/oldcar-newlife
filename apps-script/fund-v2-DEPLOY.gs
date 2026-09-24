/*  募資表單 V2（兩步驟）— 測試用獨立 Apps Script
 *  ─────────────────────────────────────────────────────────
 *  這支是「另一個專案」，跟現行的 form-handler 完全分開：
 *    · 現行 Web App 一行都不用改，線上 /fund/ 不受影響
 *    · 寫進「同一份試算表」的新分頁，不碰任何既有分頁
 *
 *  部署步驟（Jay 手動做一次）：
 *    1. script.google.com → 新增專案，命名「募資V2（測試）」
 *    2. 把這整份貼進 Code.gs，存檔
 *    3. 部署 → 新增部署作業 → 類型「網頁應用程式」
 *       執行身分：我　／　存取權：所有人
 *    4. 複製那串 /exec 網址，填進 fund-test/index.html 的 SCRIPT_URL
 *    5.（選用）專案設定 → 指令碼屬性 → 加 RESEND_API_KEY，沒加就用 Gmail 寄
 *
 *  兩步驟流程：
 *    步驟①  姓名 / 公司 / 電話 / Email      → 立刻寫一列，進度「① 只留聯絡方式」
 *    步驟②  金額 / 末五碼 / 芳名 / 收據…    → 用登記編號回頭更新同一列，狀態「待審核」
 *
 *  ★ 為什麼步驟①就要寫進試算表：
 *    這樣「看了帳號卻沒填完」的人會留在表上（進度停在①），
 *    有名字有電話可以主動聯繫 —— 正好就是現在對不到帳的那群人。
 */

const SHEET_ID = '1i3ZnaKGuYpazZHXhZ_KF6-78UcyKp3Euts4u3ruUoMs';
const FUND2_SHEET = '募資V2';          // 新分頁，自動建立
const FUND2_ACT_SHEET = '募資V2互動';  // 分享／集氣，跟正式版分開計
const FUND_GOAL = 200000;
const FUND_OK = '已通過';

const MAIL_FROM = '社團法人台灣人車公益協會 <no-reply@oldcarnewlife.org.tw>';
const MAIL_FROM_FALLBACK = '社團法人台灣人車公益協會';
const MAIL_ADMIN = ['carstory.alliance@gmail.com', 'soulbreakin@gmail.com'];
const MAIL_REPLY_TO = 'carstory.alliance@gmail.com';

// 測試期間信件標題加註記，正式轉換時把這行改成 ''
const MAIL_TAG = '【V2測試】';

// 步驟①確認信裡「回來繼續填」的連結，正式轉換時改成 /fund/
const FORM_URL = 'https://oldcarnewlife.org.tw/fund-test/';
const BANK_HTML = '<b>永豐銀行（新莊副都心）</b><br>' +
  '銀行代碼　807<br>帳號　132-01-80110-2656<br>戶名　社團法人台灣人車公益協會';

const HEAD2 = ['登記編號', '時間戳記', '姓名', '公司/單位', '電話', 'Email',
  '捐款金額', '匯款末五碼', '芳名公開方式', '公開顯示名稱', '想說的話',
  '收據需求', '收據抬頭', '身分證字號／統一編號', '備註',
  '審核狀態', '填表進度', '完成時間'];

const STEP1_LABEL = '① 只留聯絡方式';
const STEP2_LABEL = '② 已填匯款資料';

/* ═══════════════ 進出口 ═══════════════ */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const now = new Date();
    const timestamp = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

    if (data.type === 'fund2') {
      const sheet = getFund2Sheet_(ss);
      const step = Number(data.step || 0);
      if (step === 1) return jsonOut_(fund2Step1_(sheet, data, timestamp));
      if (step === 2) return jsonOut_(fund2Step2_(sheet, data, timestamp));
      return jsonOut_({ status: 'error', message: 'step 必須是 1 或 2' });

    } else if (data.type === 'fund2act') {
      const act = String(data.act || '').trim();
      if (act === 'share' || act === 'cheer') {
        const sh = ss.getSheetByName(FUND2_ACT_SHEET) || ss.insertSheet(FUND2_ACT_SHEET);
        if (sh.getLastRow() === 0) sh.appendRow(['時間戳記', '動作', '來源']);
        sh.appendRow([timestamp, act, data.from || '']);
      }
      return jsonOut_({ status: 'ok' });
    }

    return jsonOut_({ status: 'error', message: '未知的 type：' + data.type });

  } catch (err) {
    console.error('[doPost] ' + err);
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.stat === 'fund2') return fund2Stats_();
  return jsonOut_({ status: 'ok', message: '募資 V2（測試）API' });
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═══════════════ 分頁 ═══════════════ */

function getFund2Sheet_(ss) {
  let sh = ss.getSheetByName(FUND2_SHEET);
  if (!sh) {
    sh = ss.insertSheet(FUND2_SHEET);
    sh.appendRow(HEAD2);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD2);
    sh.setFrozenRows(1);
  }
  // 電話、末五碼、登記編號強制文字，避免開頭的 0 被吃掉
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('E:E').setNumberFormat('@');
  sh.getRange('H:H').setNumberFormat('@');
  return sh;
}

function colIndex_(sheet, title) {
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return head.indexOf(title) + 1;   // 1-based，找不到回 0
}

/** 用登記編號找列號，找不到回 0 */
function findRowByPledge_(sheet, pledgeId) {
  if (!pledgeId) return 0;
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(pledgeId).trim()) return i + 2;
  }
  return 0;
}

/* ═══════════════ 步驟① 只留聯絡方式 ═══════════════ */

function fund2Step1_(sheet, data, timestamp) {
  const pledgeId = String(data.pledgeId || '').trim();
  if (!pledgeId) return { status: 'error', message: '缺少登記編號' };

  // 同一個編號重送（使用者重新整理）就更新，不要長出第二列
  const exist = findRowByPledge_(sheet, pledgeId);
  const row = [
    pledgeId, timestamp,
    data.name || '', data.company || '',
    data.phone ? "'" + String(data.phone).trim() : '',
    data.email || '',
    '', '', '', '', '', '', '', '', '',
    '', STEP1_LABEL, ''
  ];

  if (exist) {
    sheet.getRange(exist, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  // 把帳號寄給他，離開了也能回來把資料填完
  try { fund2Step1Mail_(data, pledgeId); }
  catch (e) { console.warn('[mail] step1 ' + e); }

  return { status: 'ok', pledgeId: pledgeId, step: 1 };
}

/* ═══════════════ 步驟② 匯款資料 ═══════════════ */

function fund2Step2_(sheet, data, timestamp) {
  const pledgeId = String(data.pledgeId || '').trim();
  const amount = Number(String(data.amount || '').replace(/[^0-9.]/g, '')) || 0;
  let row = findRowByPledge_(sheet, pledgeId);

  // 步驟①沒寫成功（網路掉了之類）也要收得下，補一列完整的
  if (!row) {
    sheet.appendRow([
      pledgeId || ('F' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss')),
      timestamp, data.name || '', data.company || '',
      data.phone ? "'" + String(data.phone).trim() : '',
      data.email || '',
      amount,
      data.transferCode ? "'" + String(data.transferCode).trim() : '',
      data.display || '', data.displayName || '', data.message || '',
      data.receipt || '', data.receiptTitle || '', data.taxId || '',
      data.note || '', '待審核', STEP2_LABEL, timestamp
    ]);
  } else {
    // 只覆蓋步驟②的欄位，步驟①填的聯絡方式保留
    const set = function (title, value) {
      const c = colIndex_(sheet, title);
      if (c) sheet.getRange(row, c).setValue(value);
    };
    set('捐款金額', amount);
    set('匯款末五碼', data.transferCode ? "'" + String(data.transferCode).trim() : '');
    set('芳名公開方式', data.display || '');
    set('公開顯示名稱', data.displayName || '');
    set('想說的話', data.message || '');
    set('收據需求', data.receipt || '');
    set('收據抬頭', data.receiptTitle || '');
    set('身分證字號／統一編號', data.taxId || '');
    set('備註', data.note || '');
    set('審核狀態', '待審核');
    set('填表進度', STEP2_LABEL);
    set('完成時間', timestamp);
    // 步驟②可能改了姓名／電話，一併更新
    if (data.name) set('姓名', data.name);
    if (data.company) set('公司/單位', data.company);
    if (data.phone) set('電話', "'" + String(data.phone).trim());
    if (data.email) set('Email', data.email);
  }

  try { fund2Mails_(data, amount, timestamp, pledgeId); }
  catch (e) { console.warn('[mail] ' + e); }

  return { status: 'ok', pledgeId: pledgeId, step: 2 };
}

/* ═══════════════ 進度統計 ═══════════════ */

function fund2Stats_() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(FUND2_SHEET);
    const empty = { goal: FUND_GOAL, raised: 0, donors: 0, pending: 0, pendingDonors: 0,
                    unfinished: 0, shares: 0, cheers: 0, list: [] };
    if (!sh || sh.getLastRow() < 2) return jsonOut_(empty);

    const v = sh.getDataRange().getValues();
    const h = v[0];
    const iAmt = h.indexOf('捐款金額'), iSt = h.indexOf('審核狀態');
    const iWay = h.indexOf('芳名公開方式'), iName = h.indexOf('公開顯示名稱');
    const iMsg = h.indexOf('想說的話'), iTs = h.indexOf('時間戳記');
    const iProg = h.indexOf('填表進度');

    let raised = 0, donors = 0, pending = 0, pendingDonors = 0, unfinished = 0;
    const list = [];
    const DEAD = ['作廢', '退回', '取消', '無效'];

    for (let r = 1; r < v.length; r++) {
      const prog = String(v[r][iProg] || '').trim();
      // 只留了聯絡方式、還沒填匯款資料的：單獨算一個數字，不進金額
      if (prog === STEP1_LABEL) { unfinished++; continue; }

      const st = String(v[r][iSt] || '').trim();
      if (st !== FUND_OK) {
        const pa = Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
        let dead = false;
        for (let k = 0; k < DEAD.length; k++) if (st.indexOf(DEAD[k]) >= 0) dead = true;
        if (pa > 0 && !dead) { pending += pa; pendingDonors++; }
        continue;
      }
      const amt = Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
      raised += amt; donors++;
      const anon = String(v[r][iWay] || '').indexOf('匿名') >= 0;
      list.push({
        name: anon ? '匿名者' : (String(v[r][iName] || '').trim() || '匿名者'),
        amount: amt,
        msg: String(v[r][iMsg] || '').trim(),
        ts: String(v[r][iTs] || '')
      });
    }

    let shares = 0, cheers = 0;
    const ash = ss.getSheetByName(FUND2_ACT_SHEET);
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
                      unfinished: unfinished, shares: shares, cheers: cheers,
                      list: list.reverse().slice(0, 60) });
  } catch (err) {
    return jsonOut_({ goal: FUND_GOAL, raised: 0, donors: 0, pending: 0, pendingDonors: 0,
                      unfinished: 0, shares: 0, cheers: 0, list: [], error: String(err) });
  }
}

/* ═══════════════ 寄信 ═══════════════ */

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
          subject: subject, html: html, text: plain,
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

  try {
    MailApp.sendEmail({
      to: Array.isArray(to) ? to.join(',') : to,
      subject: subject, htmlBody: html, body: plain,
      name: MAIL_FROM_FALLBACK, replyTo: MAIL_REPLY_TO
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
  '<br>衛生福利部勸募許可 衛部救字第 1151363585 號　·　勸募期間 115.09.23–116.09.19' +
  '<br>內政部核准立案 台內團字第 1140047990 號　·　oldcarnewlife.org.tw' +
  '</td></tr></table></div>';
}

/** 步驟①完成：把匯款帳號寄給他，附一條能回來續填的連結 */
function fund2Step1Mail_(data, pledgeId) {
  if (!data.email) return;
  const name = data.name || '朋友';
  const back = FORM_URL + '?p=' + encodeURIComponent(pledgeId);
  sendMail_(data.email, MAIL_TAG + '魚池神國教會專案 · 匯款帳號在這裡',
    mailShell_(name + '，這是匯款帳號',
      '<p>謝謝你願意支持。匯款資訊如下：</p>' +
      '<p style="background:#f7f8fa;border-radius:10px;padding:14px 16px;line-height:2;">' + BANK_HTML + '</p>' +
      '<p><b>你的登記編號　' + pledgeId + '</b></p>' +
      '<p>匯款之後，<b>請回到下面這頁把金額與匯款末五碼填完</b>，' +
      '財務才對得起帳，你的贊助也才會出現在進度條和芳名錄上。</p>' +
      '<p><a href="' + back + '" style="display:inline-block;background:#d97b1e;color:#fff;' +
      'text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:700;">回去把資料填完</a></p>' +
      '<p style="font-size:12.5px;color:#8a8e96;">＊ 依公益勸募條例，本專案只收這一個專戶的款項，' +
      '請勿使用郵政劃撥或其他帳戶，以免無法計入本專案。</p>',
      '這封信是系統自動發送，有問題直接回覆即可。'));
}

/** 步驟②完成：回信給贊助者、通知協會 */
function fund2Mails_(data, amount, timestamp, pledgeId) {
  const name = data.name || '朋友';
  const money = 'NT$ ' + Number(amount || 0).toLocaleString('en-US');

  if (data.email) {
    sendMail_(data.email, MAIL_TAG + '謝謝你支持魚池神國教會專案',
      mailShell_(name + '，我們收到你的資料了',
        '<p>你填的贊助資料已經送到協會這邊：</p>' +
        '<p><b>金額</b>　' + money + '<br>' +
        '<b>匯款末五碼</b>　' + (data.transferCode || '') + '<br>' +
        '<b>登記編號</b>　' + pledgeId + '</p>' +
        '<p>財務會用末五碼跟銀行帳目核對，核對完成後，你的贊助就會出現在募資頁的進度條與芳名錄上。</p>' +
        '<p>需要收據的話我們會另外跟你聯絡。真的很謝謝你。</p>',
        '這封信是系統自動發送，有問題直接回覆即可。'));
  }

  sendMail_(MAIL_ADMIN, MAIL_TAG + '新的募資登記：' + name + '　' + money,
    mailShell_('有人完成募資表單（V2）',
      '<p><b>登記編號</b>　' + pledgeId + '<br>' +
      '<b>姓名</b>　' + name + '<br>' +
      '<b>公司／單位</b>　' + (data.company || '—') + '<br>' +
      '<b>電話</b>　' + (data.phone || '—') + '<br>' +
      '<b>Email</b>　' + (data.email || '—') + '<br>' +
      '<b>金額</b>　' + money + '<br>' +
      '<b>匯款末五碼</b>　' + (data.transferCode || '—') + '<br>' +
      '<b>芳名方式</b>　' + (data.display || '—') + '<br>' +
      '<b>收據</b>　' + (data.receipt || '—') + '　' + (data.receiptTitle || '') + '　' + (data.taxId || '') + '<br>' +
      '<b>留言</b>　' + (data.message || '—') + '</p>' +
      '<p><a href="https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit">打開試算表核對</a></p>',
      '送出時間 ' + timestamp));
}

/* ═══════════════ 對帳小工具 ═══════════════ */

/** 只留了聯絡方式、沒回來填匯款資料的人 —— 執行後看紀錄 */
function listUnfinished() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(FUND2_SHEET);
  if (!sh || sh.getLastRow() < 2) { console.log('沒有資料'); return; }
  const v = sh.getDataRange().getValues();
  const h = v[0];
  const iProg = h.indexOf('填表進度'), iName = h.indexOf('姓名');
  const iPhone = h.indexOf('電話'), iMail = h.indexOf('Email'), iTs = h.indexOf('時間戳記');
  const out = [];
  for (let r = 1; r < v.length; r++) {
    if (String(v[r][iProg]).trim() !== STEP1_LABEL) continue;
    out.push([v[r][iTs], v[r][iName], v[r][iPhone], v[r][iMail]].join('　'));
  }
  console.log(out.length ? ('卡在步驟①的有 ' + out.length + ' 人：\n' + out.join('\n')) : '沒有人卡在步驟①');
}
