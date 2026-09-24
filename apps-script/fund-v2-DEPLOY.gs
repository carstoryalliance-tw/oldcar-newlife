/*  ⚠️ 這份已經併入 form-handler-DEPLOY.gs，不要再單獨貼到 Apps Script。
 *
 *  2026-09-24：原本想用兩個獨立專案（表單總機 + 募資V2），實際operation下來
 *  兩邊來回貼很容易貼錯 —— 曾經把 V2 貼進表單專案，害入會／捐款／公益申請／
 *  淨灘／例會五個表單靜默失敗。現在整個協會只有一支程式碼、一個部署網址。
 *
 *  要改募資邏輯，改 form-handler-DEPLOY.gs 裡「募資 V2」那一段。
 *  這份保留只是為了對照歷史。
 */

/*  募資表單 V2（兩步驟）— 正式版 Apps Script
 *  ─────────────────────────────────────────────────────────
 *  這支是「另一個專案」，跟現行的 form-handler 完全分開：
 *    · 現行 Web App 一行都不用改，線上 /fund/ 不受影響
 *    · 寫進「同一份試算表」的新分頁，不碰任何既有分頁
 *
 *  更新方式：把這整份貼回同一個 Apps Script 專案，存檔後
 *    「管理部署作業 → 編輯（鉛筆）→ 版本：新版本 → 部署」
 *    ⚠️ 一定要走「編輯現有部署」，網址才不會變。按「新增部署作業」會拿到新網址，
 *       線上 /fund/ 就連不到了。
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

/* 理監事／財務的收件名單。
 * ⚠️ 這個 repo 是公開的（GitHub Pages），不要把私人信箱寫在這裡。
 *    在「專案設定 → 指令碼屬性」新增一列：
 *      鍵　BOARD_EMAILS
 *      值　aaa@gmail.com,bbb@gmail.com,ccc@gmail.com   ← 逗號分隔
 *    沒設定就退回寄給 MAIL_ADMIN，不會讓通知整個消失。 */
function financeEmails_() {
  const raw = PropertiesService.getScriptProperties().getProperty('FINANCE_EMAILS');
  if (!raw) return MAIL_ADMIN;
  const list = raw.split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
  return list.length ? list : MAIL_ADMIN;
}

function boardEmails_() {
  const raw = PropertiesService.getScriptProperties().getProperty('BOARD_EMAILS');
  if (!raw) return MAIL_ADMIN;
  const list = raw.split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
  return list.length ? list : MAIL_ADMIN;
}

// 正式上線，信件標題不再加註記
const MAIL_TAG = '';

// 步驟①確認信裡「回來繼續填」的連結
const FORM_URL = 'https://oldcarnewlife.org.tw/fund/';
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
  return jsonOut_({ status: 'ok', message: '魚池神國教會專案 · 募資 API' });
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
    mailShell_('有人完成募資表單',
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

/* ═══════════════ 審核通過通知 ═══════════════ */

/**
 * 安裝觸發器 —— ⚠️ 這支只要執行一次。
 * 裝好之後，把「募資V2」的審核狀態改成「已通過」，就會自動發兩封信：
 *   · 給捐款者：你的款項確認收到了
 *   · 給理監事：有一筆入帳，附目前累計進度
 */
function setupTriggersV2() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFundApprovedV2') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFundApprovedV2').forSpreadsheet(ss).onEdit().create();
  console.log('✅ 觸發器已安裝：把「' + FUND2_SHEET + '」的審核狀態改成「' + FUND_OK + '」時，會通知捐款者與理監事');
}

function onFundApprovedV2(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== FUND2_SHEET) return;
    if (String(e.value || '').trim() !== FUND_OK) return;        // 只在改成「已通過」時
    if (String(e.oldValue || '').trim() === FUND_OK) return;     // 本來就通過，不重複寄

    const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (h[e.range.getColumn() - 1] !== '審核狀態') return;        // 只認審核狀態那一欄

    const row = sh.getRange(e.range.getRow(), 1, 1, sh.getLastColumn()).getValues()[0];
    const g = function (name) { const i = h.indexOf(name); return i < 0 ? '' : row[i]; };

    const name = String(g('姓名') || '');
    if (isTestRow_(name)) return;                                 // 測試資料不發信

    const amt = Number(String(g('捐款金額') || '').replace(/[^0-9.]/g, '')) || 0;
    const anon = String(g('芳名公開方式') || '').indexOf('匿名') >= 0;
    const shown = anon ? '匿名者' : (String(g('公開顯示名稱') || '').trim() || name);
    const msg = String(g('想說的話') || '').trim();
    const email = String(g('Email') || '').trim();
    const receipt = String(g('收據需求') || '').trim();
    const pledgeId = String(g('登記編號') || '').trim();

    const st = fundTotals_(sh, h);
    const nf = function (n) { return 'NT$ ' + Number(n).toLocaleString('en-US'); };

    // ── 1. 給捐款者 ──
    if (email) {
      const needReceipt = receipt.indexOf('不需要') < 0 && receipt !== '';
      sendMail_(email, '你的贊助已確認入帳　·　魚池神國教會專案',
        mailShell_((name || '朋友') + '，款項我們確認收到了',
          '<p>謝謝你。你的贊助已經由財務核對完成，<b>正式計入這個專案</b>。</p>' +
          '<p><b>金額</b>　' + nf(amt) + '<br>' +
          '<b>登記編號</b>　' + pledgeId + '<br>' +
          '<b>芳名錄顯示</b>　' + shown + '</p>' +
          '<div style="background:#f2faf5;border:1px solid #bfe0cc;border-radius:10px;padding:14px 16px;margin:14px 0;">' +
          '<div style="font-size:13px;color:#4a7a5c;">這個專案目前累計</div>' +
          '<div style="font-size:24px;font-weight:900;color:#2f6b47;margin:4px 0;">' + nf(st.raised) +
          ' <span style="font-size:13px;color:#4a7a5c;font-weight:400;">/ ' + nf(FUND_GOAL) + '（' + st.pct + '%）</span></div>' +
          '<div style="font-size:13px;color:#4a5a50;">共 ' + st.count + ' 筆贊助' +
          (st.left > 0 ? '　·　距離目標還差 ' + nf(st.left) : '　·　<b>已達標</b>') + '</div></div>' +
          (needReceipt
            ? '<p>你勾選了需要收據，我們會另外跟你聯絡確認寄送方式。</p>'
            : '') +
          '<p>找車、驗車與整理的進度會更新在募資頁上，歡迎隨時回來看。</p>' +
          '<p><a href="' + FORM_URL + '" style="display:inline-block;background:#d97b1e;color:#fff;' +
          'text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:700;">看專案進度</a></p>',
          '這封信是系統自動發送，有問題直接回覆即可。'));
    }

    // ── 2. 給理監事 ──
    sendMail_(boardEmails_(),
      '【入帳】魚池神國教會專案　' + nf(amt) + '　累計 ' + nf(st.raised) + '（' + st.pct + '%）',
      mailShell_('有一筆贊助確認入帳',
        '<p>魚池神國教會專案有一筆贊助<b>確認入帳</b>了。</p>' +
        '<p><b>贊助者</b>　' + shown + '<br>' +
        '<b>金額</b>　' + nf(amt) + '<br>' +
        '<b>登記編號</b>　' + pledgeId +
        (msg ? '<br><b>留言</b>　「' + msg + '」' : '') + '</p>' +
        '<div style="background:#fdf7ef;border:1px solid #f0dcc0;border-radius:10px;padding:14px 16px;margin:14px 0;">' +
        '<div style="font-size:13px;color:#8b7d6b;">目前累計</div>' +
        '<div style="font-size:26px;font-weight:900;color:#d97b1e;margin:4px 0;">' + nf(st.raised) +
        ' <span style="font-size:14px;color:#8b7d6b;font-weight:400;">/ ' + nf(FUND_GOAL) + '（' + st.pct + '%）</span></div>' +
        '<div style="font-size:13px;color:#4a4d54;">共 ' + st.count + ' 筆贊助' +
        (st.left > 0 ? '　·　距離目標還差 <b>' + nf(st.left) + '</b>' : '　·　<b>已達標</b>') + '</div></div>' +
        '<p><a href="' + FORM_URL + '" style="display:inline-block;background:#d97b1e;color:#fff;' +
        'text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:900;">看募資頁</a></p>',
        '審核狀態改成「已通過」時自動發出，募資頁進度條已同步更新。'));

  } catch (err) {
    console.warn('[onFundApprovedV2] ' + err);
  }
}

/** 測試資料判斷，統一一個地方 */
function isTestRow_(name) {
  const n = String(name || '');
  return n.indexOf('測試') >= 0 || n.indexOf('請刪除') >= 0 || n.indexOf('檢查') >= 0;
}

/** 已通過的累計金額／筆數／百分比（排除測試列） */
function fundTotals_(sh, h) {
  const v = sh.getDataRange().getValues();
  const iAmt = h.indexOf('捐款金額'), iSt = h.indexOf('審核狀態'), iNm = h.indexOf('姓名');
  let raised = 0, count = 0;
  for (let r = 1; r < v.length; r++) {
    if (String(v[r][iSt] || '').trim() !== FUND_OK) continue;
    if (isTestRow_(v[r][iNm])) continue;
    raised += Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
    count++;
  }
  return {
    raised: raised, count: count,
    pct: Math.round(raised / FUND_GOAL * 1000) / 10,
    left: Math.max(0, FUND_GOAL - raised)
  };
}

/* ═══════════════ 舊分頁搬遷 ═══════════════ */

const V1_SHEET = '孤兒院募資';   // 舊版募資分頁（名稱是早期專案留下的）
const V1_ACT_SHEET = '募資互動';   // 舊版的分享／集氣計數

/** 搬遷用指紋：時間＋姓名＋金額。時間欄可能是字串也可能是 Date，統一正規化。 */
function fp_(ts, name, amt) {
  let t = '';
  if (ts instanceof Date) {
    t = Utilities.formatDate(ts, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  } else {
    t = String(ts || '').trim();
    // 舊資料若被 Sheets 轉成日期再讀回字串，格式可能不同，能轉就統一
    const d = new Date(t);
    if (!isNaN(d.getTime()) && t.length > 10) {
      t = Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    }
  }
  const a = Number(String(amt || '').replace(/[^0-9.]/g, '')) || 0;
  return t + '|' + String(name || '').trim() + '|' + a;
}

/**
 * 把舊分頁的捐款人搬進「募資V2」。
 * ★ 可以重複執行 —— 已經搬過的會跳過，不會變成兩筆。
 *   判斷方式：登記編號用 'V1-<舊分頁的列號>'，搬之前先檢查在不在。
 *
 * 執行前建議先跑 previewMigrate()（唯讀）看會搬幾筆。
 */
function migrateFromV1() {
  const r = migrate_(false);
  console.log('捐款資料：新增 ' + r.moved + ' 筆，略過（已搬過）' + r.skipped + ' 筆，空白列 ' + r.blank + ' 列');
  const a = migrateActs_(false);
  console.log('分享／集氣：新增 ' + a.moved + ' 筆，略過 ' + a.skipped + ' 筆' +
              '（分享 ' + a.shares + '、集氣 ' + a.cheers + '）');
  console.log('※ 頁面上的「贊助」數字是從捐款資料算出來的，搬完就會自己對。');
  return { fund: r, acts: a };
}

/** 唯讀預演：只印出會搬什麼，不動任何資料 */
function previewMigrate() {
  const r = migrate_(true);
  console.log('【預演】捐款資料會搬 ' + r.moved + ' 筆，略過 ' + r.skipped + ' 筆，空白 ' + r.blank + ' 列');
  console.log(r.rows.join(String.fromCharCode(10)));
  const a = migrateActs_(true);
  console.log('【預演】分享／集氣會搬 ' + a.moved + ' 筆（分享 ' + a.shares + '、集氣 ' + a.cheers + '），略過 ' + a.skipped + ' 筆');
  return { fund: r, acts: a };
}

function migrate_(dryRun) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const src = ss.getSheetByName(V1_SHEET);
  if (!src) throw new Error('找不到舊分頁「' + V1_SHEET + '」');
  const dst = getFund2Sheet_(ss);

  const v = src.getDataRange().getValues();
  if (v.length < 2) return { moved: 0, skipped: 0, blank: 0, rows: [] };
  const h = v[0];
  const col = function (title) { return h.indexOf(title); };

  const iTs = col('時間戳記'), iName = col('姓名'), iComp = col('公司/單位');
  const iPhone = col('電話'), iMail = col('Email'), iAmt = col('捐款金額');
  const iCode = col('匯款末五碼'), iWay = col('芳名公開方式'), iDisp = col('公開顯示名稱');
  const iMsg = col('想說的話'), iRcpt = col('收據需求'), iTitle = col('收據抬頭');
  const iTax = col('身分證字號／統一編號'), iNote = col('備註'), iSt = col('審核狀態');

  // 目的分頁已經有的資料，用「時間＋姓名＋金額」指紋判斷有沒有搬過。
  // 刻意不用列號 —— 你在舊分頁刪一列或排序一次，列號就全位移了，
  // 那樣重跑會把同一個人搬成兩筆。
  const have = {};
  if (dst.getLastRow() > 1) {
    const dv = dst.getDataRange().getValues();
    const dh = dv[0];
    const dTs = dh.indexOf('時間戳記'), dName = dh.indexOf('姓名'), dAmt = dh.indexOf('捐款金額');
    for (let k = 1; k < dv.length; k++) {
      have[fp_(dv[k][dTs], dv[k][dName], dv[k][dAmt])] = true;
    }
  }

  let moved = 0, skipped = 0, blank = 0;
  const rows = [];
  const out = [];

  for (let r = 1; r < v.length; r++) {
    const name = iName >= 0 ? String(v[r][iName] || '').trim() : '';
    const amt = iAmt >= 0 ? (Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0) : 0;
    if (!name && !amt) { blank++; continue; }        // 整列空的就跳過

    if (have[fp_(v[r][iTs], name, amt)]) { skipped++; continue; }
    const pledgeId = 'V1-' + (r + 1);                // r+1 = 舊分頁當下的列號，僅供人工對照

    const pick = function (i) { return i >= 0 ? (v[r][i] || '') : ''; };
    const phone = String(pick(iPhone)).trim();
    const code = String(pick(iCode)).trim();

    out.push([
      pledgeId, pick(iTs), name, pick(iComp),
      phone ? "'" + phone : '',
      pick(iMail), amt,
      code ? "'" + code : '',
      pick(iWay), pick(iDisp), pick(iMsg),
      pick(iRcpt), pick(iTitle), pick(iTax), pick(iNote),
      String(pick(iSt) || '').trim(),   // 審核狀態原樣保留，已通過的搬過去還是已通過
      STEP2_LABEL,                      // 舊表的人本來就填完整了
      pick(iTs)
    ]);
    rows.push(pledgeId + '　' + name + '　' + amt + '　' + String(pick(iSt)).trim());
    moved++;
  }

  if (!dryRun && out.length) {
    dst.getRange(dst.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  }
  return { moved: moved, skipped: skipped, blank: blank, rows: rows };
}

/**
 * 搬分享／集氣計數（舊「募資互動」→ 新「募資V2互動」）。
 * 同樣可以重複執行。
 *
 * ⚠️ 這裡不能只用指紋去重 —— 同一秒可能有兩個人都按了分享，
 *    那兩筆的「時間＋動作＋來源」一模一樣，用指紋會被當成同一筆而少搬。
 *    所以改成比對「每個指紋在兩邊各出現幾次」，只補差額。
 */
function migrateActs_(dryRun) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const src = ss.getSheetByName(V1_ACT_SHEET);
  const out = { moved: 0, skipped: 0, shares: 0, cheers: 0 };
  if (!src || src.getLastRow() < 2) return out;

  let dst = ss.getSheetByName(FUND2_ACT_SHEET);
  if (!dst) { dst = ss.insertSheet(FUND2_ACT_SHEET); dst.appendRow(['時間戳記', '動作', '來源']); }
  if (dst.getLastRow() === 0) dst.appendRow(['時間戳記', '動作', '來源']);

  const key = function (row) {
    const t = (row[0] instanceof Date)
      ? Utilities.formatDate(row[0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss')
      : String(row[0] || '').trim();
    return t + '|' + String(row[1] || '').trim() + '|' + String(row[2] || '').trim();
  };

  // 目的分頁每個指紋已經有幾筆
  const dstCount = {};
  if (dst.getLastRow() > 1) {
    const dv = dst.getRange(2, 1, dst.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < dv.length; i++) {
      const k = key(dv[i]);
      dstCount[k] = (dstCount[k] || 0) + 1;
    }
  }

  const sv = src.getRange(2, 1, src.getLastRow() - 1, 3).getValues();
  const used = {};
  const rows = [];
  for (let i = 0; i < sv.length; i++) {
    const act = String(sv[i][1] || '').trim();
    if (act !== 'share' && act !== 'cheer') continue;
    const k = key(sv[i]);
    used[k] = (used[k] || 0) + 1;
    if ((dstCount[k] || 0) >= used[k]) { out.skipped++; continue; }   // 這一筆已經搬過了
    rows.push([sv[i][0], act, sv[i][2] || '']);
    out.moved++;
    if (act === 'share') out.shares++; else out.cheers++;
  }

  if (!dryRun && rows.length) {
    dst.getRange(dst.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  }
  return out;
}

/**
 * 盯搬遷期間有沒有人又填了舊表單。
 * 列出舊分頁裡「還沒搬進新分頁」的資料 —— 正常情況下搬完就該是 0 筆。
 * 不是 0 就再跑一次 migrateFromV1()（可以重複執行，不會重複搬）。
 */
function checkV1Leftover() {
  const r = migrate_(true);     // 唯讀預演
  const a = migrateActs_(true);
  if (r.moved === 0 && a.moved === 0) {
    console.log('✅ 舊分頁沒有漏掉的資料，捐款與分享／集氣全部都在新分頁了');
  } else {
    if (r.moved > 0) {
      console.log('⚠️ 捐款資料還有 ' + r.moved + ' 筆沒搬過來：');
      console.log(r.rows.join(String.fromCharCode(10)));
    }
    if (a.moved > 0) {
      console.log('⚠️ 分享／集氣還有 ' + a.moved + ' 筆沒搬（分享 ' + a.shares + '、集氣 ' + a.cheers + '）');
    }
    console.log('→ 跑一次 migrateFromV1() 就會補進去');
  }
  return { fund: r, acts: a };
}

/** 搬完對一下：兩邊的金額與筆數應該一致 */
function compareV1V2() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sum = function (sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return { rows: 0, total: 0, ok: 0, okTotal: 0 };
    const v = sh.getDataRange().getValues(), h = v[0];
    const iAmt = h.indexOf('捐款金額'), iSt = h.indexOf('審核狀態');
    let rows = 0, total = 0, ok = 0, okTotal = 0;
    for (let r = 1; r < v.length; r++) {
      const a = Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
      if (!a) continue;
      rows++; total += a;
      if (String(v[r][iSt] || '').trim() === FUND_OK) { ok++; okTotal += a; }
    }
    let last = '';
    if (v.length > 1) {
      const iTs = h.indexOf('時間戳記');
      const lv = v[v.length - 1][iTs];
      last = (lv instanceof Date)
        ? Utilities.formatDate(lv, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss')
        : String(lv || '');
    }
    return { rows: rows, total: total, ok: ok, okTotal: okTotal, last: last };
  };
  const a = sum(V1_SHEET), b = sum(FUND2_SHEET);
  console.log('舊分頁「' + V1_SHEET + '」：' + a.rows + ' 筆 / 共 ' + a.total +
              '（已通過 ' + a.ok + ' 筆 / ' + a.okTotal + '）　最後一筆 ' + a.last);
  console.log('新分頁「' + FUND2_SHEET + '」：' + b.rows + ' 筆 / 共 ' + b.total +
              '（已通過 ' + b.ok + ' 筆 / ' + b.okTotal + '）　最後一筆 ' + b.last);
  // 分享／集氣也對一下
  const acts = function (sheetName) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return { share: 0, cheer: 0 };
    const av = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
    let share = 0, cheer = 0;
    for (let i = 0; i < av.length; i++) {
      const x = String(av[i][0] || '').trim();
      if (x === 'share') share++; else if (x === 'cheer') cheer++;
    }
    return { share: share, cheer: cheer };
  };
  const a2 = acts(V1_ACT_SHEET), b2 = acts(FUND2_ACT_SHEET);
  console.log('分享／集氣　舊：' + a2.share + ' / ' + a2.cheer +
              '　新：' + b2.share + ' / ' + b2.cheer);

  const moneyOk = b.okTotal >= a.okTotal;
  const actOk = (b2.share >= a2.share && b2.cheer >= a2.cheer);
  console.log(moneyOk ? '✅ 新分頁的已通過金額沒有短少' : '⚠️ 金額比舊分頁少，請檢查');
  console.log(actOk ? '✅ 分享／集氣數字沒有短少' : '⚠️ 分享或集氣比舊分頁少，再跑一次 migrateFromV1()');
  return { v1: a, v2: b, v1acts: a2, v2acts: b2 };
}

/* ═══════════════ 手動寄信：財務／理監事 ═══════════════ */

/** 時間欄可能是 Date 也可能是字串，統一格式化 */
function tsText_(ts, fmt) {
  if (ts instanceof Date) return Utilities.formatDate(ts, 'Asia/Taipei', fmt);
  const t = String(ts || '');
  return fmt === 'MM/dd' ? t.slice(5, 10) : t.slice(5).replace(/:\d\d$/, '');
}

function nf_(n) { return 'NT$ ' + Number(n || 0).toLocaleString('en-US'); }

/**
 * 手動執行：把待核對的贊助清單寄給財務。
 * 收件人設在指令碼屬性 FINANCE_EMAILS（逗號分隔），沒設就寄給 MAIL_ADMIN。
 */
function mailPendingToFinance() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(FUND2_SHEET);
  if (!sh || sh.getLastRow() < 2) { console.log('沒有資料'); return; }

  const v = sh.getDataRange().getValues();
  const h = v[0];
  const iTs = h.indexOf('時間戳記'), iName = h.indexOf('姓名'), iComp = h.indexOf('公司/單位');
  const iAmt = h.indexOf('捐款金額'), iCode = h.indexOf('匯款末五碼'), iSt = h.indexOf('審核狀態');
  const iPhone = h.indexOf('電話'), iRcpt = h.indexOf('收據需求');
  const iPid = h.indexOf('登記編號'), iProg = h.indexOf('填表進度');

  const DEAD = ['作廢', '退回', '取消', '無效'];
  let total = 0;
  const rows = [];
  const waiting = [];   // 只留了聯絡方式、還沒回來填匯款資料的人

  for (let r = 1; r < v.length; r++) {
    const nm = String(v[r][iName] || '');
    if (isTestRow_(nm)) continue;

    if (String(v[r][iProg] || '').trim() === STEP1_LABEL) {
      waiting.push({ nm: nm, phone: v[r][iPhone] || '', ts: v[r][iTs] });
      continue;
    }

    const st = String(v[r][iSt] || '').trim();
    if (st === FUND_OK) continue;
    let dead = false;
    for (let k = 0; k < DEAD.length; k++) if (st.indexOf(DEAD[k]) >= 0) dead = true;
    if (dead) continue;

    const amt = Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
    if (amt <= 0) continue;
    total += amt;

    rows.push(
      '<tr>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:13px;white-space:nowrap;">' + tsText_(v[r][iTs], 'MM/dd HH:mm') + '</td>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:13px;">' + nm +
        (v[r][iComp] ? '<br><span style="color:#8a8e96;font-size:11.5px;">' + v[r][iComp] + '</span>' : '') +
        '<br><span style="color:#b0b3b8;font-size:11px;">' + (v[r][iPid] || '') + '</span></td>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:13px;text-align:right;font-weight:900;color:#d97b1e;white-space:nowrap;">' +
        nf_(amt) + '</td>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:13px;text-align:center;letter-spacing:1px;">' +
        (v[r][iCode] || '—') + '</td>' +
      '<td style="padding:9px 10px;border-top:1px solid #f2f3f5;font-size:11.5px;color:#8a8e96;">' +
        (v[r][iRcpt] || '') + '<br>' + (v[r][iPhone] || '') + '</td>' +
      '</tr>');
  }

  if (!rows.length && !waiting.length) { console.log('目前沒有待核對的贊助'); return; }

  let body = '';
  if (rows.length) {
    body +=
      '<p>魚池神國教會專案目前有 <b>' + rows.length + ' 筆</b>贊助等待核對入帳，合計 ' +
      '<b style="color:#d97b1e;">' + nf_(total) + '</b>。</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:14px 0;' +
      'border:1px solid #eef0f3;border-radius:10px;border-collapse:separate;">' +
      '<tr style="background:#fafbfc;">' +
      '<th style="padding:9px 10px;text-align:left;font-size:11.5px;color:#8a8e96;">時間</th>' +
      '<th style="padding:9px 10px;text-align:left;font-size:11.5px;color:#8a8e96;">贊助者／登記編號</th>' +
      '<th style="padding:9px 10px;text-align:right;font-size:11.5px;color:#8a8e96;">金額</th>' +
      '<th style="padding:9px 10px;text-align:center;font-size:11.5px;color:#8a8e96;">末五碼</th>' +
      '<th style="padding:9px 10px;text-align:left;font-size:11.5px;color:#8a8e96;">收據／電話</th></tr>' +
      rows.join('') + '</table>' +
      '<p><b>核對方式：</b>比對永豐勸募專戶入帳的末五碼與金額，確認無誤後把試算表' +
      '「<b>' + FUND2_SHEET + '</b>」分頁的「審核狀態」改成 <b>' + FUND_OK + '</b>（要一字不差）。' +
      '改完系統會自動通知捐款人與理監事，募資頁進度條也會從灰色轉成橘色。</p>';
  } else {
    body += '<p>目前沒有待核對的贊助。</p>';
  }

  // ★ 帳上有不明入帳時，答案通常在這份名單裡
  if (waiting.length) {
    let wr = '';
    for (let i = 0; i < waiting.length; i++) {
      wr += '<tr><td style="padding:8px 10px;border-top:1px solid #f2f3f5;font-size:12.5px;white-space:nowrap;color:#8a8e96;">' +
            tsText_(waiting[i].ts, 'MM/dd HH:mm') + '</td>' +
            '<td style="padding:8px 10px;border-top:1px solid #f2f3f5;font-size:12.5px;">' + waiting[i].nm + '</td>' +
            '<td style="padding:8px 10px;border-top:1px solid #f2f3f5;font-size:12.5px;">' + waiting[i].phone + '</td></tr>';
    }
    body +=
      '<div style="background:#f7f8fa;border:1px solid #e3e5ea;border-radius:10px;padding:14px 16px;margin:18px 0 6px;">' +
      '<div style="font-size:13.5px;font-weight:900;color:#1a1a2e;margin-bottom:4px;">' +
      '另有 ' + waiting.length + ' 人留了聯絡方式，但還沒回來填匯款資料</div>' +
      '<div style="font-size:12.5px;color:#8a8e96;line-height:1.8;">' +
      '他們拿到帳號了，可能已經匯款但沒填表。<b>帳上如果有對不到人的入帳，答案通常就在這份名單裡</b>' +
      ' —— 直接打電話問是最快的。</div>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin-top:10px;">' + wr + '</table>' +
      '</div>';
  }

  body +=
    '<p style="margin-top:16px;"><a href="https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit" ' +
    'style="display:inline-block;background:#1a1a2e;color:#fff;text-decoration:none;padding:11px 22px;' +
    'border-radius:9px;font-weight:900;">開啟試算表核對</a>　' +
    '<a href="' + FORM_URL + '" style="display:inline-block;background:#d97b1e;color:#fff;' +
    'text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:900;">看募資頁</a></p>';

  const subj = rows.length
    ? '【待核對】魚池神國教會專案　' + rows.length + ' 筆　' + nf_(total)
    : '【待核對】魚池神國教會專案　有 ' + waiting.length + ' 人尚未完成登記';

  sendMail_(financeEmails_(), subj,
    mailShell_('魚池神國教會專案：待核對清單', body, '由協會表單系統整理發出，測試資料已自動排除。'));
  console.log('已寄給財務：待核對 ' + rows.length + ' 筆 / ' + total +
              '，未完成登記 ' + waiting.length + ' 人');
}

/** 手動執行：把目前募資進度與已入帳名單，寄給理監事＋財務 */
function mailProgressToAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(FUND2_SHEET);
  if (!sh || sh.getLastRow() < 2) { console.log('沒有資料'); return; }
  const v = sh.getDataRange().getValues();
  const h = v[0];
  const iTs = h.indexOf('時間戳記'), iNm = h.indexOf('姓名'), iAmt = h.indexOf('捐款金額');
  const iSt = h.indexOf('審核狀態'), iWay = h.indexOf('芳名公開方式'), iShow = h.indexOf('公開顯示名稱');
  const iMsg = h.indexOf('想說的話'), iProg = h.indexOf('填表進度');

  let raised = 0, pending = 0, pcount = 0, waiting = 0;
  const rows = [];
  for (let r = 1; r < v.length; r++) {
    if (isTestRow_(v[r][iNm])) continue;
    if (String(v[r][iProg] || '').trim() === STEP1_LABEL) { waiting++; continue; }
    const amt = Number(String(v[r][iAmt] || '').replace(/[^0-9.]/g, '')) || 0;
    if (amt <= 0) continue;
    const st = String(v[r][iSt] || '').trim();
    if (st !== FUND_OK) { pending += amt; pcount++; continue; }
    raised += amt;
    const anon = String(v[r][iWay] || '').indexOf('匿名') >= 0;
    const who = anon ? '匿名者' : (String(v[r][iShow] || '').trim() || String(v[r][iNm] || ''));
    const msg = String(v[r][iMsg] || '').trim();
    rows.push(
      '<tr><td style="padding:9px 12px;border-top:1px solid #f2f3f5;font-size:13px;color:#8a8e96;white-space:nowrap;">' +
      tsText_(v[r][iTs], 'MM/dd') + '</td>' +
      '<td style="padding:9px 12px;border-top:1px solid #f2f3f5;font-size:13px;">' + who +
      (msg ? '<br><span style="color:#8a8e96;font-size:11.5px;">「' + msg + '」</span>' : '') + '</td>' +
      '<td style="padding:9px 12px;border-top:1px solid #f2f3f5;font-size:13px;text-align:right;font-weight:900;color:#d97b1e;white-space:nowrap;">' +
      nf_(amt) + '</td></tr>');
  }
  const pct = Math.round(raised / FUND_GOAL * 1000) / 10;
  const left = Math.max(0, FUND_GOAL - raised);
  const barW = Math.min(100, pct);

  const body =
    '<div style="background:#fdf7ef;border:1px solid #f0dcc0;border-radius:12px;padding:16px 18px;margin-bottom:16px;">' +
    '<div style="font-size:13px;color:#8b7d6b;">目前募得</div>' +
    '<div style="font-size:30px;font-weight:900;color:#d97b1e;margin:4px 0 8px;">' + nf_(raised) +
    ' <span style="font-size:14px;color:#8b7d6b;font-weight:400;">/ ' + nf_(FUND_GOAL) + '</span></div>' +
    '<div style="height:12px;background:#eceef2;border-radius:8px;overflow:hidden;">' +
    '<div style="height:12px;width:' + barW + '%;background:#d97b1e;border-radius:8px;"></div></div>' +
    '<div style="margin-top:8px;font-size:13px;color:#4a4d54;">已完成 <b>' + pct + '%</b>　·　共 ' + rows.length + ' 筆' +
    (left > 0 ? '　·　還差 <b>' + nf_(left) + '</b>' : '　·　<b>已達標</b>') +
    (pending > 0 ? '<br><span style="color:#8a8e96;">另有 ' + pcount + ' 筆待核對，' + nf_(pending) + '</span>' : '') +
    (waiting > 0 ? '<br><span style="color:#8a8e96;">' + waiting + ' 人留了聯絡方式但尚未完成登記</span>' : '') +
    '</div></div>' +
    '<p><b>已入帳名單</b></p>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;margin:8px 0 14px;' +
    'border:1px solid #eef0f3;border-radius:10px;">' + rows.join('') + '</table>' +
    '<p style="font-size:12px;color:#8a8e96;">勸募許可 衛部救字第 1151363585 號　·　期間 115.09.23–116.09.19</p>' +
    '<p><a href="' + FORM_URL + '" style="display:inline-block;background:#d97b1e;color:#fff;' +
    'text-decoration:none;padding:11px 22px;border-radius:9px;font-weight:900;">看募資頁</a></p>';

  const to = boardEmails_().concat(financeEmails_());
  sendMail_(to, '【魚池神國教會專案】目前募得 ' + nf_(raised) + '（' + pct + '%）',
    mailShell_('募資進度回報', body, '由協會表單系統整理發出，測試資料已排除。'));
  console.log('已寄給：' + to.join(', '));
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
