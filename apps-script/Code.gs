/**
 * AngelNote Gmail reminders.
 * Run setupDailyGmailReminder once from the Apps Script editor while signed in
 * with the same Google account used to sign in to AngelNote.
 */
const ANGELNOTE = Object.freeze({
  projectId: 'daymark-10-30',
  databaseId: '(default)',
  collectionId: 'planner',
  timeZone: 'Asia/Taipei',
  sendHour: 9,
  maxLeadDays: 365,
  sentKeyPrefix: 'ANGELNOTE_SENT_',
  sentKeyRetentionDays: 400
});

/** Authorizes the required services, verifies Firestore access, and installs the daily trigger. */
function setupDailyGmailReminder() {
  const ownerEmail = getOwnerEmail_();
  const documents = queryPlannerDocuments_(ownerEmail, 1);
  if (!documents.length) {
    throw new Error('尚未找到此帳號的 AngelNote 資料。請先登入網站，等待同步完成後再設定排程。');
  }
  PropertiesService.getScriptProperties().setProperty('ANGELNOTE_OWNER_EMAIL', ownerEmail);

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'sendDueGmailReminders')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('sendDueGmailReminders')
    .timeBased()
    .atHour(ANGELNOTE.sendHour)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone(ANGELNOTE.timeZone)
    .create();
  console.log(`Daily AngelNote Gmail reminder scheduled for ${ownerEmail}.`);
}

/** Sends one daily digest for enabled, incomplete tasks whose lead time is today. */
function sendDueGmailReminders() {
  const ownerEmail = PropertiesService.getScriptProperties().getProperty('ANGELNOTE_OWNER_EMAIL') || getOwnerEmail_();
  const today = Utilities.formatDate(new Date(), ANGELNOTE.timeZone, 'yyyy-MM-dd');
  const documents = queryPlannerDocuments_(ownerEmail);
  const dueItems = [];

  documents.forEach(document => {
    const root = decodeFields_(document.fields || {});
    const data = root.data || {};
    if (normalizeEmail_(data.reminderOwnerEmail) !== normalizeEmail_(ownerEmail)) return;

    const items = [...(Array.isArray(data.tasks) ? data.tasks : []), ...(Array.isArray(data.extra) ? data.extra : [])];
    items.forEach(item => {
      if (!item || item.done === true || item.gmailReminderEnabled !== true) return;
      if (item.kind && item.kind !== 'task' && item.kind !== 'deadline') return;
      const dueDate = String(item.date || '');
      const leadDays = Number(item.gmailReminderDays);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDate) || !Number.isInteger(leadDays) || leadDays < 0 || leadDays > ANGELNOTE.maxLeadDays) return;
      if (daysBetween_(today, dueDate) !== leadDays) return;

      const identity = `${document.name}|${String(item.id || item.title || '')}|${dueDate}`;
      const key = ANGELNOTE.sentKeyPrefix + digestKey_(identity);
      dueItems.push({ item, dueDate, key });
    });
  });

  const properties = PropertiesService.getScriptProperties();
  pruneSentKeys_(properties);
  const newItems = dueItems.filter(entry => !properties.getProperty(entry.key));
  if (!newItems.length) return;

  const dateLabel = Utilities.formatDate(new Date(`${today}T00:00:00+08:00`), ANGELNOTE.timeZone, 'yyyy 年 M 月 d 日');
  const lines = newItems.map(({ item, dueDate }) => {
    const dueLabel = Utilities.formatDate(new Date(`${dueDate}T00:00:00+08:00`), ANGELNOTE.timeZone, 'M 月 d 日');
    const details = [item.time ? `時間：${item.time}` : '', item.category ? `分類：${item.category}` : '', item.description ? `說明：${item.description}` : ''].filter(Boolean);
    return `・${item.title || '未命名待辦'}（${dueLabel}）${details.length ? `\n  ${details.join('｜')}` : ''}`;
  });
  const subject = `AngelNote 待辦提醒｜${dateLabel}`;
  const body = [`AngelNote 待辦提醒（${dateLabel}）`, '', ...lines, '', '完成後可回 AngelNote 勾選任務。'].join('\n');

  // Mark only after MailApp returns successfully so failed sends can retry tomorrow.
  MailApp.sendEmail(ownerEmail, subject, body, { name: 'AngelNote' });
  const sentAt = new Date().toISOString();
  const sentProperties = {};
  newItems.forEach(entry => { sentProperties[entry.key] = sentAt; });
  properties.setProperties(sentProperties, false);
}

/** Reads only the current user's planner document(s) through the Firestore REST API. */
function queryPlannerDocuments_(ownerEmail, limit) {
  const parent = `projects/${ANGELNOTE.projectId}/databases/${ANGELNOTE.databaseId}/documents`;
  const structuredQuery = {
    from: [{ collectionId: ANGELNOTE.collectionId, allDescendants: true }],
    where: {
      fieldFilter: {
        field: { fieldPath: 'data.reminderOwnerEmail' },
        op: 'EQUAL',
        value: { stringValue: ownerEmail }
      }
    }
  };
  if (limit) structuredQuery.limit = limit;

  const response = UrlFetchApp.fetch(`https://firestore.googleapis.com/v1/${parent}:runQuery`, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    payload: JSON.stringify({ structuredQuery }),
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error(`Firestore read failed (${status}): ${response.getContentText()}`);
  }
  return JSON.parse(response.getContentText())
    .filter(row => row.document)
    .map(row => row.document);
}

function getOwnerEmail_() {
  const email = normalizeEmail_(Session.getEffectiveUser().getEmail());
  if (!email) throw new Error('Apps Script 無法取得目前 Google 帳號，請確認你使用個人帳號登入。');
  return email;
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function daysBetween_(fromDate, toDate) {
  const from = fromDate.split('-').map(Number);
  const to = toDate.split('-').map(Number);
  return Math.round((Date.UTC(to[0], to[1] - 1, to[2]) - Date.UTC(from[0], from[1] - 1, from[2])) / 86400000);
}

function digestKey_(value) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value)).replace(/=+$/, '').slice(0, 40);
}

function pruneSentKeys_(properties) {
  const all = properties.getProperties();
  const cutoff = Date.now() - ANGELNOTE.sentKeyRetentionDays * 86400000;
  Object.keys(all).forEach(key => {
    if (!key.startsWith(ANGELNOTE.sentKeyPrefix)) return;
    const timestamp = Date.parse(all[key]);
    if (Number.isFinite(timestamp) && timestamp < cutoff) properties.deleteProperty(key);
  });
}

function decodeFields_(fields) {
  return Object.fromEntries(Object.entries(fields).map(([key, value]) => [key, decodeValue_(value)]));
}

function decodeValue_(value) {
  if (Object.prototype.hasOwnProperty.call(value, 'stringValue')) return value.stringValue;
  if (Object.prototype.hasOwnProperty.call(value, 'integerValue')) return Number(value.integerValue);
  if (Object.prototype.hasOwnProperty.call(value, 'doubleValue')) return Number(value.doubleValue);
  if (Object.prototype.hasOwnProperty.call(value, 'booleanValue')) return value.booleanValue;
  if (Object.prototype.hasOwnProperty.call(value, 'nullValue')) return null;
  if (value.mapValue) return decodeFields_(value.mapValue.fields || {});
  if (value.arrayValue) return (value.arrayValue.values || []).map(decodeValue_);
  if (value.timestampValue) return value.timestampValue;
  return null;
}

/** Optional cleanup: removes this project's daily trigger and stops automatic sending. */
function removeDailyGmailReminder() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'sendDueGmailReminders')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}
