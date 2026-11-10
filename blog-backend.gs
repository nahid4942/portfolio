/**
 * Blog backend (Google Apps Script).
 * Setup:
 * 1. Create a Google Sheet, then Extensions > Apps Script, and paste this file.
 * 2. Change NEW_PASSWORD inside setPassword(), select setPassword, click Run, accept the permissions.
 *    Then you can delete the password text from the script. Only a hash is stored.
 * 3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Copy the /exec URL
 *    into APPS_URL in blog.html and blog-admin.html.
 * After editing this script later, use Deploy > Manage deployments > Edit > New version.
 */
const HEAD = ['id', 'type', 'category', 'title', 'excerpt', 'content', 'image_url', 'youtube_url', 'author', 'published_at'];
const FOLDER = 'Blog media';

function setPassword() {
  const NEW_PASSWORD = 'change-this-password';
  if (NEW_PASSWORD.length < 8) throw new Error('Use at least 8 characters');
  PropertiesService.getScriptProperties().setProperty('PW', hash_(NEW_PASSWORD));
}
function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'blog-salt:' + s));
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); if (name === 'posts') s.appendRow(HEAD); }
  return s;
}
function readPosts_() {
  const v = sheet_('posts').getDataRange().getValues().slice(1);
  return v.filter(r => r[0]).map(r => {
    const o = {};
    HEAD.forEach((k, i) => o[k] = r[i] instanceof Date ? r[i].toISOString() : String(r[i] === null ? '' : r[i]));
    return o;
  }).sort((a, b) => new Date(b.published_at) - new Date(a.published_at));
}
function readSettings_() {
  const t = sheet_('settings').getRange('A1').getValue();
  try { return t ? JSON.parse(t) : {}; } catch (e) { return {}; }
}
function doGet() {
  return out_({ posts: readPosts_(), settings: readSettings_() });
}
function authed_(t) {
  return t && CacheService.getScriptCache().get('tok_' + t) === '1';
}
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'Bad request' }); }
  if (d.action === 'login') {
    const pw = PropertiesService.getScriptProperties().getProperty('PW');
    if (!pw || hash_(d.password || '') !== pw) { Utilities.sleep(1000); return out_({ ok: false, error: 'Wrong password' }); }
    const t = Utilities.getUuid() + Utilities.getUuid();
    CacheService.getScriptCache().put('tok_' + t, '1', 21600);
    return out_({ ok: true, token: t });
  }
  if (!authed_(d.token)) return out_({ ok: false, error: 'Session expired. Sign in again.' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.action === 'savePost') {
      const s = sheet_('posts'), p = d.post;
      let row = 0;
      if (p.id) {
        const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
        row = ids.indexOf(String(p.id)) + 1;
      }
      if (!row) { p.id = Utilities.getUuid(); row = s.getLastRow() + 1; }
      s.getRange(row, 1, 1, HEAD.length).setNumberFormat('@').setValues([HEAD.map(k => p[k] || '')]);
      return out_({ ok: true, id: p.id });
    }
    if (d.action === 'deletePost') {
      const s = sheet_('posts');
      const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
      const i = ids.indexOf(String(d.id));
      if (i > 0) s.deleteRow(i + 1);
      return out_({ ok: true });
    }
    if (d.action === 'saveSettings') {
      const c = sheet_('settings').getRange('A1');
      c.setNumberFormat('@').setValue(JSON.stringify(d.settings));
      return out_({ ok: true });
    }
    if (d.action === 'upload') {
      const it = DriveApp.getFoldersByName(FOLDER);
      const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
      const blob = Utilities.newBlob(Utilities.base64Decode(d.data), d.type || 'image/jpeg', Date.now() + '-' + String(d.name || 'image').replace(/[^\w.]/g, '_'));
      const f = folder.createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return out_({ ok: true, url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1600' });
    }
    if (d.action === 'changePassword') {
      if (!d.newPassword || d.newPassword.length < 8) return out_({ ok: false, error: 'Use at least 8 characters' });
      PropertiesService.getScriptProperties().setProperty('PW', hash_(d.newPassword));
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'Unknown action' });
  } finally {
    lock.releaseLock();
  }
}
