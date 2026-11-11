/**
 * ONE Google Apps Script backend for BOTH the blog and the portfolio.
 * Setup:
 * 1. Create (or reuse) a Google Sheet, then Extensions > Apps Script. Replace ALL code with this file.
 * 2. In setPassword() set NEW_EMAIL (portfolio admin login) and NEW_PASSWORD (min 8 chars, used by BOTH admins).
 *    Select setPassword, click Run, accept permissions. Afterwards delete the password text. Only a hash is stored.
 * 3. Select authorize, click Run (grants Sheets + Drive access).
 * 4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Copy the /exec URL and paste it as
 *    APPS_URL in: blog.html, blog-admin.html and data.js (the portfolio files use data.js).
 * After editing this script later: Deploy > Manage deployments > Edit > New version.
 *
 * Sheets used: "posts" (blog), "settings" (blog settings JSON), "portfolio" (portfolio JSON, split in chunks).
 * Uploads go to the Drive folder "Site media" (images and PDFs, link-viewable).
 */
const HEAD = ['id', 'type', 'category', 'title', 'excerpt', 'content', 'image_url', 'youtube_url', 'author', 'published_at'];
const FOLDER = 'Site media';

function setPassword() {
  const NEW_EMAIL = 'you@example.com';
  const NEW_PASSWORD = 'change-this-password';
  if (NEW_PASSWORD.length < 8) throw new Error('Use at least 8 characters');
  const p = PropertiesService.getScriptProperties();
  p.setProperty('PW', hash_(NEW_PASSWORD));
  p.setProperty('EMAIL', String(NEW_EMAIL).trim().toLowerCase());
}
// Run once after pasting new code: grants permissions and creates the data sheet if needed.
function authorize() {
  ss_();
  DriveApp.getRootFolder();
}
function hash_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'blog-salt:' + s));
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function ss_() {
  let ss = SpreadsheetApp.getActive();
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('Site data');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function sheet_(name) {
  const ss = ss_();
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  if (name === 'posts' && s.getLastRow() === 0) s.appendRow(HEAD);
  return s;
}

/* ---------- Blog data ---------- */
function readPosts_(all) {
  const v = sheet_('posts').getDataRange().getValues().slice(1);
  return v.filter(r => r[0]).map(r => {
    const o = {};
    HEAD.forEach((k, i) => o[k] = r[i] instanceof Date ? r[i].toISOString() : String(r[i] === null ? '' : r[i]));
    return o;
  }).filter(p => all || new Date(p.published_at) <= new Date())
    .sort((a, b) => new Date(b.published_at) - new Date(a.published_at));
}
function readSettings_() {
  const t = sheet_('settings').getRange('A1').getValue();
  try { return t ? JSON.parse(t) : {}; } catch (e) { return {}; }
}

/* ---------- Portfolio data (one JSON document, stored in 40k-char chunks) ---------- */
function readPortfolio_() {
  const s = sheet_('portfolio'), n = s.getLastRow();
  if (!n) return {};
  const t = s.getRange(1, 1, n, 1).getValues().map(r => r[0]).join('');
  try { return t ? JSON.parse(t) : {}; } catch (e) { return {}; }
}
function writePortfolio_(o) {
  if (!o || typeof o !== 'object' || !o.site || typeof o.site !== 'object') throw new Error('Invalid portfolio data');
  const t = JSON.stringify(o);
  if (t.length > 400000) throw new Error('Portfolio data too large. Use file upload for images, not embedded data.');
  const s = sheet_('portfolio');
  s.clearContents();
  const parts = [];
  for (let i = 0; i < t.length; i += 40000) parts.push([t.slice(i, i + 40000)]);
  s.getRange(1, 1, parts.length, 1).setNumberFormat('@').setValues(parts);
}

/* ---------- Public read ---------- */
function doGet(e) {
  try {
    const only = e && e.parameter && e.parameter.only;
    if (only === 'portfolio') return out_({ portfolio: readPortfolio_() });
    return out_({ posts: readPosts_(), settings: readSettings_() });
  } catch (err) {
    return out_({ posts: [], settings: {}, portfolio: {}, error: String(err.message || err) });
  }
}

/* ---------- Authenticated actions ---------- */
function authed_(t) {
  return t && CacheService.getScriptCache().get('tok_' + t) === '1';
}
function doPost(e) {
  try {
    return handle_(e);
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}
function handle_(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'Bad request' }); }
  if (d.action === 'login') {
    const props = PropertiesService.getScriptProperties();
    const pw = props.getProperty('PW'), em = props.getProperty('EMAIL');
    const emailOk = d.email === undefined || !em || String(d.email).trim().toLowerCase() === em;
    if (!pw || !emailOk || hash_(d.password || '') !== pw) { Utilities.sleep(1000); return out_({ ok: false, error: 'Wrong email or password' }); }
    const t = Utilities.getUuid() + Utilities.getUuid();
    CacheService.getScriptCache().put('tok_' + t, '1', 21600);
    return out_({ ok: true, token: t });
  }
  if (!authed_(d.token)) return out_({ ok: false, error: 'Session expired. Sign in again.' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.action === 'list') return out_({ ok: true, posts: readPosts_(true), settings: readSettings_(), portfolio: readPortfolio_() });

    // Blog
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

    // Portfolio
    if (d.action === 'savePortfolio') {
      writePortfolio_(d.portfolio);
      return out_({ ok: true });
    }

    // Shared: image / PDF upload to Drive
    if (d.action === 'upload') {
      const it = DriveApp.getFoldersByName(FOLDER);
      const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
      const type = d.type || 'image/jpeg';
      const blob = Utilities.newBlob(Utilities.base64Decode(d.data), type, Date.now() + '-' + String(d.name || 'file').replace(/[^\w.]/g, '_'));
      const f = folder.createFile(blob);
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = type === 'application/pdf'
        ? 'https://drive.google.com/uc?export=download&id=' + f.getId()
        : 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1600';
      return out_({ ok: true, url: url });
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
https://script.google.com/macros/s/AKfycbw8AZmKEtGurjbcpqe0DPgAeBfL_8t5EVOtWZMo_-jjqLBpv9IFsRHVLtcIQ-wevI_j/exec