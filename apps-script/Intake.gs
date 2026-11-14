/**
 * Recipe email intake — runs in the SEPARATE intake Google account (not your personal one).
 *
 * It looks at that account's own inbox every 10 minutes and pushes each new email to the recipe app's main script,
 * where it waits in the Review tab. It can only see the intake account's mail; your personal inbox is never involved.
 *
 * One-time setup, signed in to the INTAKE account:
 *   1. script.google.com > New project. Paste this file in. Save.
 *   2. Project Settings (gear) > Script properties > add two:
 *        APP_URL   = the recipe app's web app URL (the long ".../exec" address)
 *        PASSCODE  = the shared passcode
 *   3. Select the function "setup" and click Run. Approve the permissions. (Google may warn "unverified app":
 *      Advanced > Go to ... > Allow. It is your own script.)
 *   4. Done. Forward or send recipe emails to the intake account's address.
 */

/** Run once: checks the settings, creates the 10-minute timer, and collects any mail already waiting. */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('APP_URL') || !p.getProperty('PASSCODE')) {
    throw new Error('First add the script properties APP_URL and PASSCODE (Project Settings > Script properties).');
  }
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'collect') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('collect').timeBased().everyMinutes(10).create();
  collect();
  Logger.log('Set up. New emails are collected every 10 minutes.');
}

/**
 * Recipe-looking attachments (screenshots, photos, PDFs, Word files, text files), including pictures pasted into the
 * email body. Up to 3 per email, each under 6 MB. Tiny images (logos, signature icons) are skipped.
 */
function files_(message) {
  const ok = /^(image\/(jpeg|png|webp|gif)|application\/pdf|application\/msword|application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document|text\/plain)$/;
  const out = [];
  message.getAttachments({ includeInlineImages: true, includeAttachments: true }).forEach(function (a) {
    const mime = String(a.getContentType()).toLowerCase().split(';')[0];
    const size = a.getSize();
    if (out.length >= 3 || !ok.test(mime) || size > 6 * 1024 * 1024) return;
    if (/^image\//.test(mime) && size < 20 * 1024) return;
    out.push({ name: a.getName(), mime: mime, data: Utilities.base64Encode(a.getBytes()) });
  });
  return out;
}

/** Sends new inbox emails to the app, then labels and archives them so they are only sent once. */
function collect() {
  const p = PropertiesService.getScriptProperties();
  const url = p.getProperty('APP_URL'), key = p.getProperty('PASSCODE');
  if (!url || !key) return;
  const sent = GmailApp.getUserLabelByName('recipe-sent') || GmailApp.createLabel('recipe-sent');
  const address = Session.getEffectiveUser().getEmail();
  const threads = GmailApp.search('in:inbox -label:recipe-sent', 0, 10);
  for (let i = 0; i < threads.length; i++) {
    const t = threads[i];
    const m = t.getMessages()[0];
    const body = {
      key: key, action: 'intakeadd', address: address,
      item: { id: m.getId(), from: m.getFrom(), subject: m.getSubject(), date: m.getDate().toISOString(), plain: m.getPlainBody().slice(0, 60000), files: files_(m) }
    };
    let result;
    try {
      const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'text/plain', payload: JSON.stringify(body), muteHttpExceptions: true });
      result = JSON.parse(res.getContentText());
    } catch (e) { Logger.log('Could not reach the app: ' + e.message); return; }   // try again on the next run
    if (result && result.error === 'auth') { Logger.log('The passcode is wrong. Fix PASSCODE in Script properties.'); return; }
    if (result && result.error === 'queue is full') { Logger.log('The Review queue is full. Post or dismiss some, then it will continue.'); return; }
    if (result && result.ok) { t.addLabel(sent); t.moveToArchive(); }
  }
}
