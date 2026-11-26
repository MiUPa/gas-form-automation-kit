/**
 * Setup.gs — セットアップ用関数とカスタムメニュー
 *
 * 導入時は、このファイルの setupAutomation() を Apps Script エディタから
 * 1回実行してください（初回は Google の承認ダイアログが出ます）。
 */

/** スプレッドシートを開いたときに「フォーム自動化」メニューを追加します */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('フォーム自動化')
    .addItem('セットアップ（トリガー登録）', 'setupAutomation')
    .addItem('テスト通知を送る', 'testNotifications')
    .addItem('未処理分をまとめて処理', 'processPendingRows')
    .addSeparator()
    .addItem('トリガーを削除（停止）', 'removeAutomation')
    .addToUi();
}

/**
 * セットアップ: 送信トリガーを登録し、受付ログシートを用意します。
 * 何度実行してもトリガーは二重登録されません。
 */
function setupAutomation() {
  removeTriggers_();
  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(CONFIG.LOG_SHEET_NAME)) {
    const sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sheet.getRange(1, 1).setValue('受付日時');
    formatLogHeader_(sheet, 1);
  }
  PropertiesService.getScriptProperties()
    .setProperty('LAST_PROCESSED_ROW', String(getResponseSheet_().getLastRow()));

  SpreadsheetApp.getUi().alert(
    'セットアップ完了',
    'フォームが送信されるたびに、自動返信と通知が動くようになりました。\n' +
      'まず「フォーム自動化 > テスト通知を送る」で通知先を確認してください。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** トリガーを削除して自動化を止めます */
function removeAutomation() {
  removeTriggers_();
  SpreadsheetApp.getActiveSpreadsheet().toast('送信トリガーを削除しました', 'フォーム自動化');
}

function removeTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'handleFormSubmit')
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

/** 設定済みの通知先にテストメッセージを送ります */
function testNotifications() {
  const text = CONFIG.NOTIFY_TITLE + '\nこれはテスト通知です。設定が正しく完了していれば、このメッセージが届きます。';
  const results = [];
  if (getSecret_('SLACK_WEBHOOK_URL')) {
    results.push('Slack: ' + (notifySlack_(text) ? '送信成功' : '送信失敗（ログを確認）'));
  } else {
    results.push('Slack: 未設定（スキップ）');
  }
  if (getSecret_('LINE_CHANNEL_ACCESS_TOKEN') && getSecret_('LINE_TO_ID')) {
    results.push('LINE: ' + (notifyLine_(text) ? '送信成功' : '送信失敗（ログを確認）'));
  } else {
    results.push('LINE: 未設定（スキップ）');
  }
  if (CONFIG.ADMIN_EMAIL) {
    MailApp.sendEmail(CONFIG.ADMIN_EMAIL, CONFIG.NOTIFY_TITLE + '（テスト）', text);
    results.push('管理者メール: 送信しました');
  }
  SpreadsheetApp.getUi().alert('テスト通知の結果', results.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
