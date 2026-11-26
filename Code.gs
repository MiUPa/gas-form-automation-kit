/**
 * Code.gs — メイン処理
 *
 * Google フォームの送信をトリガーに、次の3つを自動で行います。
 *   1. 回答を整形して「受付ログ」シートに記録
 *   2. 申込者へ自動返信メールを送信
 *   3. 管理者へ通知（メール / Slack / LINE、設定したものだけ）
 *
 * 実行の入口は Setup.gs の setupAutomation() で登録される
 * インストール型トリガー（onFormSubmit）です。
 */

/**
 * フォーム送信時のトリガーハンドラ。
 * @param {Object} e スプレッドシートの onFormSubmit イベント
 */
function handleFormSubmit(e) {
  const record = eventToRecord_(e);
  processRecord_(record);
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LAST_PROCESSED_ROW', String(e.range.getRow()));
}

/**
 * 未処理の回答をまとめて処理します（トリガーが止まっていた期間の穴埋め用）。
 * カスタムメニュー「フォーム自動化 > 未処理分をまとめて処理」から実行します。
 */
function processPendingRows() {
  const sheet = getResponseSheet_();
  const lastRow = sheet.getLastRow();
  const props = PropertiesService.getScriptProperties();
  let processed = parseInt(props.getProperty('LAST_PROCESSED_ROW') || '1', 10);
  if (processed < 1) processed = 1;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let count = 0;
  for (let row = processed + 1; row <= lastRow; row++) {
    const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
    if (values.every((v) => v === '')) continue;
    processRecord_(rowToRecord_(headers, values));
    processed = row;
    count++;
  }
  props.setProperty('LAST_PROCESSED_ROW', String(processed));
  SpreadsheetApp.getActiveSpreadsheet().toast(count + ' 件を処理しました', 'フォーム自動化');
}

/** 1件分の処理: ログ記録 → 自動返信 → 通知 */
function processRecord_(record) {
  const status = { reply: '不要', slack: '対象外', line: '対象外' };

  const email = findEmail_(record);
  if (email) {
    try {
      sendReplyMail_(record, email);
      status.reply = '送信済み';
    } catch (err) {
      status.reply = '失敗: ' + err.message;
    }
  }

  const text = buildNotificationText_(record);
  if (getSecret_('SLACK_WEBHOOK_URL')) {
    status.slack = notifySlack_(text) ? '送信済み' : '失敗';
  }
  if (getSecret_('LINE_CHANNEL_ACCESS_TOKEN') && getSecret_('LINE_TO_ID')) {
    status.line = notifyLine_(text) ? '送信済み' : '失敗';
  }
  if (CONFIG.ADMIN_EMAIL) {
    try {
      MailApp.sendEmail(CONFIG.ADMIN_EMAIL, CONFIG.NOTIFY_TITLE, text);
    } catch (err) {
      console.log('管理者メール通知に失敗: ' + err.message);
    }
  }

  appendLog_(record, status);
}

/** トリガーイベントを { timestamp, items: [{header, value}] } に変換 */
function eventToRecord_(e) {
  const items = [];
  const headers = Object.keys(e.namedValues);
  headers.forEach((h) => {
    items.push({ header: h, value: (e.namedValues[h] || []).join(', ') });
  });
  return { timestamp: e.values[0] ? new Date(e.values[0]) : new Date(), items: items };
}

/** シートの1行を record 形式に変換（まとめ処理用） */
function rowToRecord_(headers, values) {
  const items = headers.map((h, i) => ({ header: String(h), value: formatValue_(values[i]) }));
  const ts = values[0] instanceof Date ? values[0] : new Date();
  return { timestamp: ts, items: items };
}

function formatValue_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  if (Array.isArray(v)) return v.join(', ');
  return v == null ? '' : String(v);
}

/** record からメールアドレスを探します */
function findEmail_(record) {
  const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  let target = null;
  if (CONFIG.EMAIL_COLUMN_HEADER) {
    target = record.items.find((i) => i.header === CONFIG.EMAIL_COLUMN_HEADER);
  } else {
    target = record.items.find((i) => /メール|email/i.test(i.header) && emailRe.test(i.value.trim()));
  }
  if (!target) return '';
  const v = target.value.trim();
  return emailRe.test(v) ? v : '';
}

/** 通知メッセージ本文を組み立てます */
function buildNotificationText_(record) {
  const lines = [CONFIG.NOTIFY_TITLE, '受付日時: ' + formatValue_(record.timestamp), ''];
  record.items.forEach((i) => {
    if (i.header === 'タイムスタンプ') return;
    lines.push('■ ' + i.header + '\n' + (i.value || '（未入力）'));
  });
  return lines.join('\n');
}

/** 受付ログシートへ1行追加します（なければ作成し、見出しを整形します） */
function appendLog_(record, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  const questionHeaders = record.items
    .map((i) => i.header)
    .filter((h) => h !== 'タイムスタンプ');
  const headers = ['受付日時'].concat(questionHeaders, ['自動返信', 'Slack通知', 'LINE通知']);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatLogHeader_(sheet, headers.length);
  }
  const row = [formatValue_(record.timestamp)]
    .concat(questionHeaders.map((h) => {
      const item = record.items.find((i) => i.header === h);
      return item ? item.value : '';
    }))
    .concat([status.reply, status.slack, status.line]);
  sheet.appendRow(row);
}

function formatLogHeader_(sheet, colCount) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setFontWeight('bold');
  range.setBackground('#F0EEE6');
  sheet.setFrozenRows(1);
}

/** 回答シートを取得します */
function getResponseSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CONFIG.RESPONSE_SHEET_NAME) {
    const named = ss.getSheetByName(CONFIG.RESPONSE_SHEET_NAME);
    if (named) return named;
  }
  const sheets = ss.getSheets();
  const found = sheets.find((s) => s.getName().indexOf('フォームの回答') === 0);
  if (found) return found;
  const fallback = sheets.find((s) => s.getName() !== CONFIG.LOG_SHEET_NAME);
  if (!fallback) throw new Error('回答シートが見つかりません。Config.gs の RESPONSE_SHEET_NAME を指定してください。');
  return fallback;
}
