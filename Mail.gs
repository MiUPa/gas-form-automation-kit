/**
 * Mail.gs — 自動返信メール
 *
 * Config.gs の REPLY_SUBJECT / REPLY_BODY テンプレートを使います。
 * テンプレート内の {{列の見出し}} はその項目の回答に置き換わります。
 * {{全項目}} は「見出し: 回答」の一覧、{{タイムスタンプ}} は受付日時です。
 */

function sendReplyMail_(record, to) {
  const body = renderTemplate_(CONFIG.REPLY_BODY, record);
  const options = {};
  if (CONFIG.REPLY_FROM_NAME) options.name = CONFIG.REPLY_FROM_NAME;
  MailApp.sendEmail(to, CONFIG.REPLY_SUBJECT, body, options);
}

function renderTemplate_(template, record) {
  let out = template;
  out = out.replace(/\{\{タイムスタンプ\}\}/g, formatValue_(record.timestamp));
  out = out.replace(/\{\{全項目\}\}/g, allItemsText_(record));
  record.items.forEach((i) => {
    const pattern = new RegExp('\\{\\{' + escapeRegExp_(i.header) + '\\}\\}', 'g');
    out = out.replace(pattern, i.value || '');
  });
  // 未入力で残ったプレースホルダは空にします
  out = out.replace(/\{\{[^}]+\}\}/g, '');
  return out;
}

function allItemsText_(record) {
  return record.items
    .filter((i) => i.header !== 'タイムスタンプ')
    .map((i) => i.header + ': ' + (i.value || '（未入力）'))
    .join('\n');
}

function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
