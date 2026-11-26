/**
 * Notify.gs — Slack / LINE 通知
 *
 * 送信先の設定はスクリプト プロパティ（公開したくない値なのでコードには書きません）:
 *   SLACK_WEBHOOK_URL          … Slack Incoming Webhook の URL
 *   LINE_CHANNEL_ACCESS_TOKEN  … LINE Messaging API のチャネルアクセストークン
 *   LINE_TO_ID                 … LINE の送信先（ユーザーID または グループID）
 *
 * ※ LINE Notify は 2025年3月末でサービス終了したため、本キットは
 *    LINE 公式アカウントの Messaging API（プッシュ送信）を使います。
 */

/** Slack へ通知します。成功したら true */
function notifySlack_(text) {
  const url = getSecret_('SLACK_WEBHOOK_URL');
  if (!url) return false;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true,
  });
  const ok = res.getResponseCode() === 200;
  if (!ok) console.log('Slack通知に失敗: HTTP ' + res.getResponseCode() + ' ' + res.getContentText());
  return ok;
}

/** LINE へ通知します。成功したら true */
function notifyLine_(text) {
  const token = getSecret_('LINE_CHANNEL_ACCESS_TOKEN');
  const to = getSecret_('LINE_TO_ID');
  if (!token || !to) return false;
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      to: to,
      messages: [{ type: 'text', text: text.slice(0, 4500) }],
    }),
    muteHttpExceptions: true,
  });
  const ok = res.getResponseCode() === 200;
  if (!ok) console.log('LINE通知に失敗: HTTP ' + res.getResponseCode() + ' ' + res.getContentText());
  return ok;
}
