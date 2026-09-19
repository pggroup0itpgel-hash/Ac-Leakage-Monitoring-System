/**
 * AC Leakage - Apps Script OTP + Sheet Sync API
 * Deploy as Web App:
 * - Execute as: Me
 * - Who has access: Anyone
 */

const SHEET_ID = '1yecFSz_FrKQ0UFsavuwZjL_-AwvHTJbyx6_NMT9fN6M';
const OTP_TTL_MS = 10 * 60 * 1000;
const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000;
// IT admins that must always receive OTP (also auto-added to `admins` tab if missing).
const BOOTSTRAP_IT_ADMIN_EMAILS = ['verify.software2040@pgel.in'];
// OTP emails are sent FROM this address (requires Gmail "Send mail as" on script owner, or deploy as this user).
const DEFAULT_OTP_SENDER_EMAIL = 'verify.software2040@pgel.in';
const DEFAULT_OTP_SENDER_NAME = 'AC Leakage Monitoring';

const TAB = {
  activity: 'activity_logs',
  admins: 'admins',
  employees: 'employees',
  settings: 'settings',
  locations: 'locations',
  plants: 'plants',
  lines: 'lines',
  defects: 'defects',
  otp: 'otp',
  reports: 'defectreports_locationplantwise',
  productCatalog: 'product_catalog',
  uiSettings: 'ui_settings',
  setupSheet: 'setup_sheet',
  monthlySettings: 'monthly_report_settings',
  emailLogs: 'email_activity_logs'
};

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(payload.action || '').trim();
    let result;

    if (action === 'sendOtp') result = sendOtp_(payload);
    else if (action === 'verifyOtp') result = verifyOtp_(payload);
    else if (action === 'getConfig') result = getConfig_(payload);
    else if (action === 'syncSettings') result = syncSettings_(payload);
    else if (action === 'syncDefectReport') result = syncDefectReport_(payload);
    else if (action === 'logActivity') result = logActivity_(payload);
    else if (action === 'getActivityLogs') result = getActivityLogs_(payload);
    else if (action === 'getEmailLogs') result = getEmailLogs_(payload);
    else if (action === 'getDefectReports') result = getDefectReports_(payload);
    else if (action === 'getMonthlyReportSettings') result = getMonthlyReportSettings_(payload);
    else if (action === 'saveMonthlyReportSettings') result = saveMonthlyReportSettings_(payload);
    else if (action === 'previewMonthlyReport') result = previewMonthlyReport_(payload);
    else if (action === 'sendMonthlyReports') result = sendMonthlyPlantLeakageReports(payload);
    else if (action === 'setupMonthlyTrigger') result = setupMonthlyReportTrigger_();
    else if (action === 'applySetupSheet') result = applySetupSheet_(payload);
    else if (action === 'getStatus') result = getStatus_(payload);
    else throw new Error('Unsupported action: ' + action);

    return json_(Object.assign({ ok: true }, result || {}));
  } catch (err) {
    return json_({ ok: false, error: err.message || 'Unknown error' });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function generateOtp_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function randomToken_() {
  return Utilities.getUuid() + '-' + Utilities.getUuid();
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  const propId = String(props.getProperty('SHEET_ID') || '').trim();
  const id = (SHEET_ID && SHEET_ID !== 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE') ? SHEET_ID : propId;
  if (!id) {
    throw new Error('Google Sheet is not configured. Set SHEET_ID in Code.gs or in Apps Script Project Settings → Script properties (key: SHEET_ID).');
  }
  return SpreadsheetApp.openById(id);
}

function getStatus_() {
  const props = PropertiesService.getScriptProperties();
  const propId = String(props.getProperty('SHEET_ID') || '').trim();
  const configured = (!!SHEET_ID && SHEET_ID !== 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE') || !!propId;
  const tabs = Object.keys(TAB).map((k) => TAB[k]);
  const sender = getOtpSender_();
  let scriptOwner = '';
  try { scriptOwner = Session.getActiveUser().getEmail() || ''; } catch (_) {}
  return {
    serverTime: new Date().toISOString(),
    sheetConfigured: configured,
    requiredTabs: tabs,
    otpSenderEmail: sender.email,
    otpSenderName: sender.name,
    scriptOwnerEmail: scriptOwner
  };
}

function getUiSetting_(key) {
  const sh = ensureTab_(TAB.uiSettings, ['key', 'value', 'updatedAt']);
  const values = sh.getDataRange().getValues();
  const want = String(key || '').trim();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === want) return String(values[i][1] || '').trim();
  }
  return '';
}

function setUiSetting_(key, value) {
  const sh = ensureTab_(TAB.uiSettings, ['key', 'value', 'updatedAt']);
  const values = sh.getDataRange().getValues();
  const want = String(key || '').trim();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === want) {
      sh.getRange(i + 1, 2).setValue(value);
      sh.getRange(i + 1, 3).setValue(new Date());
      return;
    }
  }
  sh.appendRow([want, value, new Date()]);
}

function getOtpSender_() {
  const props = PropertiesService.getScriptProperties();
  const email = normalizeEmail_(
    getUiSetting_('otpSenderEmail') ||
    props.getProperty('OTP_SENDER_EMAIL') ||
    DEFAULT_OTP_SENDER_EMAIL
  );
  const name = String(
    getUiSetting_('otpSenderName') ||
    props.getProperty('OTP_SENDER_NAME') ||
    DEFAULT_OTP_SENDER_NAME
  ).trim() || DEFAULT_OTP_SENDER_NAME;
  return { email: email, name: name };
}

function ensureOtpSenderUiSettings_() {
  const legacyWrongSender = 'verify.software.2040@pgel.in';
  const current = normalizeEmail_(getUiSetting_('otpSenderEmail'));
  if (!current || current === normalizeEmail_(legacyWrongSender)) {
    setUiSetting_('otpSenderEmail', DEFAULT_OTP_SENDER_EMAIL);
  }
  if (!getUiSetting_('otpSenderName')) setUiSetting_('otpSenderName', DEFAULT_OTP_SENDER_NAME);
}

function sendOtpEmail_(toEmail, subject, htmlBody) {
  sendSystemEmail_(toEmail, '', subject, htmlBody);
}

function normalizeEmailList_(input) {
  if (!input) return [];
  if (Array.isArray(input)) {
    return input.map((e) => normalizeEmail_(e)).filter(Boolean);
  }
  return String(input)
    .split(/[,;\s]+/)
    .map((e) => normalizeEmail_(e))
    .filter(Boolean);
}

function sendSystemEmail_(toEmails, ccEmails, subject, htmlBody, plainBody, attachments) {
  const sender = getOtpSender_();
  const toList = normalizeEmailList_(toEmails);
  const ccList = normalizeEmailList_(ccEmails);

  if (!toList.length) {
    throw new Error('At least one recipient (To) email is required.');
  }

  const toStr = toList.join(',');
  const ccStr = ccList.join(',');
  const plain = plainBody || 'Please view this message in an email client that supports HTML.';

  const mailOpts = {
    htmlBody: htmlBody,
    name: sender.name,
    replyTo: sender.email,
    from: sender.email
  };
  if (ccStr) mailOpts.cc = ccStr;
  if (attachments && attachments.length) mailOpts.attachments = attachments;

  try {
    GmailApp.sendEmail(toStr, subject, plain, mailOpts);
    return;
  } catch (gmailErr) {
    const gmailMsg = gmailErr && gmailErr.message ? gmailErr.message : String(gmailErr);
    try {
      const basicMailOpts = {
        to: toStr,
        subject: subject,
        htmlBody: htmlBody,
        body: plain,
        name: sender.name,
        replyTo: sender.email,
        from: sender.email
      };
      if (ccStr) basicMailOpts.cc = ccStr;
      if (attachments && attachments.length) basicMailOpts.attachments = attachments;
      MailApp.sendEmail(basicMailOpts);
      return;
    } catch (mailErr) {
      const mailMsg = mailErr && mailErr.message ? mailErr.message : String(mailErr);
      const aliasIssue = (gmailMsg + ' ' + mailMsg).toLowerCase();
      if (aliasIssue.indexOf('from') >= 0 || aliasIssue.indexOf('alias') >= 0 || aliasIssue.indexOf('address') >= 0) {
        const fallbackOpts = {
          to: toStr,
          subject: subject,
          htmlBody: htmlBody,
          body: plain,
          name: sender.name,
          replyTo: sender.email
        };
        if (ccStr) fallbackOpts.cc = ccStr;
        if (attachments && attachments.length) fallbackOpts.attachments = attachments;
        MailApp.sendEmail(fallbackOpts);
        return;
      }
      throw new Error(
        'Failed to send email to ' + toStr + ' (CC: ' + (ccStr || 'none') + '). ' +
        (mailMsg || gmailMsg)
      );
    }
  }
}

function ensureTab_(name, headers) {
  const ss = getSheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function appendRow_(tabName, headers, row) {
  const sh = ensureTab_(tabName, headers);
  sh.appendRow(row);
}

function getRowsAsObjects_(sh) {
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  const headers = values[0].map((h) => String(h || '').trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;
      row[headers[j]] = values[i][j];
    }
    rows.push(row);
  }
  return rows;
}

function getBootstrapAdminUsers_() {
  const out = {};
  BOOTSTRAP_IT_ADMIN_EMAILS.forEach(function(email) {
    const norm = normalizeEmail_(email);
    if (!norm) return;
    out[norm] = {
      email: norm,
      role: 'it_admin',
      allowedPlants: ['*'],
      allowedLocations: ['*'],
      canSubmitDefects: true
    };
  });
  return out;
}

function ensureBootstrapAdmins_() {
  const adminsSh = ensureTab_(TAB.admins, ['email', 'role', 'allowedPlants', 'allowedLocations', 'canSubmitDefects', 'updatedAt']);
  const values = adminsSh.getDataRange().getValues();
  const existing = {};
  for (let i = 1; i < values.length; i++) {
    const e = normalizeEmail_(values[i][0]);
    if (e) existing[e] = true;
  }
  const legacyWrongAdmin = normalizeEmail_('verify.software.2040@pgel.in');
  BOOTSTRAP_IT_ADMIN_EMAILS.forEach(function(email) {
    const norm = normalizeEmail_(email);
    if (!norm || existing[norm]) return;
    adminsSh.appendRow([norm, 'it_admin', '*', '*', 'true', new Date()]);
    existing[norm] = true;
  });
  if (legacyWrongAdmin && existing[legacyWrongAdmin] && !existing[normalizeEmail_(DEFAULT_OTP_SENDER_EMAIL)]) {
    adminsSh.appendRow([normalizeEmail_(DEFAULT_OTP_SENDER_EMAIL), 'it_admin', '*', '*', 'true', new Date()]);
  }
}

function getUsersFromSheetsByRole_() {
  ensureBootstrapAdmins_();
  const admins = ensureTab_(TAB.admins, ['email', 'role', 'allowedPlants', 'allowedLocations', 'canSubmitDefects', 'updatedAt']);
  const employees = ensureTab_(TAB.employees, ['email', 'role', 'allowedPlants', 'allowedLocations', 'canSubmitDefects', 'updatedAt']);

  const adminUsers = getBootstrapAdminUsers_();
  const employeeUsers = {};
  const load = (sh, defaultRole, target) => {
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const email = normalizeEmail_(values[i][0]);
      if (!email) continue;
      target[email] = {
        email: email,
        role: values[i][1] || defaultRole,
        allowedPlants: parseCsv_(values[i][2]) || ['*'],
        allowedLocations: parseCsv_(values[i][3]) || ['*'],
        canSubmitDefects: String(values[i][4]).toLowerCase() === 'false' ? false : true
      };
    }
  };
  load(admins, 'it_admin', adminUsers);
  load(employees, 'employee', employeeUsers);
  return { adminUsers, employeeUsers };
}

function getUsersFromSheets_() {
  const byRole = getUsersFromSheetsByRole_();
  return Object.assign({}, byRole.adminUsers, byRole.employeeUsers);
}

function resolveUserProfileForSource_(email, source) {
  const byRole = getUsersFromSheetsByRole_();
  const s = String(source || '').trim().toLowerCase();
  if (s === 'dashboard') return byRole.adminUsers[email] || null;
  if (s === 'qr-report' || s === 'qr') return byRole.employeeUsers[email] || byRole.adminUsers[email] || null;
  return byRole.adminUsers[email] || byRole.employeeUsers[email] || null;
}

function parseCsv_(v) {
  const s = String(v || '').trim();
  if (!s) return null;
  return s.split(',').map(x => x.trim()).filter(Boolean);
}

function getConfig_() {
  const byRole = getUsersFromSheetsByRole_();
  const users = Object.assign({}, byRole.adminUsers, byRole.employeeUsers);
  const config = {
    users: {},
    plants: {},
    scannerCards: [],
    defectMaster: {}
  };
  Object.keys(users).forEach((email) => {
    const u = users[email];
    config.users[email.replace(/\./g, ',')] = {
      email: u.email,
      role: u.role || 'employee',
      allowedPlants: u.allowedPlants || ['*'],
      allowedLocations: u.allowedLocations || ['*'],
      canSubmitDefects: u.canSubmitDefects !== false,
      canManageScanner: ['it_admin', 'full_access'].includes(u.role)
    };
  });
  config.adminUsers = {};
  config.employeeUsers = {};
  Object.keys(byRole.adminUsers).forEach((email) => {
    config.adminUsers[email.replace(/\./g, ',')] = byRole.adminUsers[email];
  });
  Object.keys(byRole.employeeUsers).forEach((email) => {
    config.employeeUsers[email.replace(/\./g, ',')] = byRole.employeeUsers[email];
  });

  const plantsSh = ensureTab_(TAB.plants, ['plant', 'updatedAt']);
  const locationsSh = ensureTab_(TAB.locations, ['plant', 'location', 'updatedAt']);
  const linesSh = ensureTab_(TAB.lines, ['plant', 'line', 'updatedAt']);
  const defectsSh = ensureTab_(TAB.defects, ['defectKey', 'configJson', 'updatedAt']);
  const productCatalogSh = ensureTab_(TAB.productCatalog, ['location', 'productName', 'linesCsv', 'defectsCsv', 'stagesCsv', 'jointsJson', 'image1', 'image2', 'updatedAt']);
  const uiSettingsSh = ensureTab_(TAB.uiSettings, ['key', 'value', 'updatedAt']);
  const setupSheetSh = ensureTab_(TAB.setupSheet, ['location', 'productName', 'linesCsv', 'defectsCsv', 'stagesCsv', 'jointsJson', 'image1', 'image2', 'updatedAt']);

  const plants = plantsSh.getDataRange().getValues();
  for (let i = 1; i < plants.length; i++) {
    const plant = String(plants[i][0] || '').trim();
    if (!plant) continue;
    if (!config.plants[plant]) config.plants[plant] = { locations: [], lines: [] };
  }
  const locs = locationsSh.getDataRange().getValues();
  for (let i = 1; i < locs.length; i++) {
    const plant = String(locs[i][0] || '').trim();
    const location = String(locs[i][1] || '').trim();
    if (!plant || !location) continue;
    if (!config.plants[plant]) config.plants[plant] = { locations: [], lines: [] };
    if (config.plants[plant].locations.indexOf(location) < 0) config.plants[plant].locations.push(location);
  }
  const lines = linesSh.getDataRange().getValues();
  for (let i = 1; i < lines.length; i++) {
    const plant = String(lines[i][0] || '').trim();
    const line = String(lines[i][1] || '').trim();
    if (!plant || !line) continue;
    if (!config.plants[plant]) config.plants[plant] = { locations: [], lines: [] };
    if (config.plants[plant].lines.indexOf(line) < 0) config.plants[plant].lines.push(line);
  }
  const defects = defectsSh.getDataRange().getValues();
  for (let i = 1; i < defects.length; i++) {
    const key = String(defects[i][0] || '').trim();
    if (!key) continue;
    try {
      config.defectMaster[key] = JSON.parse(String(defects[i][1] || '{}'));
    } catch (_) {
      config.defectMaster[key] = {};
    }
  }
  if (!config.plants.PGTL) config.plants.PGTL = { locations: ['Pune'], lines: [] };

  config.defectsByLocation = {};
  const dbSh = ensureTab_('defects_by_location', ['location', 'defectsCsv', 'updatedAt']);
  const dbv = dbSh.getDataRange().getValues();
  for (let i = 1; i < dbv.length; i++) {
    const loc = String(dbv[i][0] || '').trim();
    const csv = String(dbv[i][1] || '').trim();
    if (!loc) continue;
    config.defectsByLocation[loc] = csv ? csv.split(',').map(function (s) { return s.trim(); }).filter(Boolean) : ['Assembly Line Defect', 'Headershop Defect', 'HE Shop Defect'];
  }

  config.locations = {};
  Object.keys(config.plants).forEach(function (plant) {
    var p = config.plants[plant] || {};
    var locs = p.locations || [];
    var lines = p.lines || [];
    locs.forEach(function (loc) {
      config.locations[loc] = config.locations[loc] || { plants: {} };
      config.locations[loc].plants[plant] = config.locations[loc].plants[plant] || { lines: [] };
      lines.forEach(function (line) {
        if (config.locations[loc].plants[plant].lines.indexOf(line) < 0) config.locations[loc].plants[plant].lines.push(line);
      });
    });
  });
  config.productCatalog = {};
  config.productCatalogByLocation = {};
  config.uiSecurity = { settingsPin: '1234', excelPin: '1234' };
  const uiRows = uiSettingsSh.getDataRange().getValues();
  for (let i = 1; i < uiRows.length; i++) {
    const key = String(uiRows[i][0] || '').trim();
    const value = String(uiRows[i][1] || '').trim();
    if (!key) continue;
    if (key === 'settingsPin') config.uiSecurity.settingsPin = value || '1234';
    if (key === 'excelPin') config.uiSecurity.excelPin = value || '1234';
    if (key === 'otpSenderEmail') config.otpSenderEmail = value || DEFAULT_OTP_SENDER_EMAIL;
    if (key === 'otpSenderName') config.otpSenderName = value || DEFAULT_OTP_SENDER_NAME;
  }
  const pcv = productCatalogSh.getDataRange().getValues();
  for (let i = 1; i < pcv.length; i++) {
    const location = String(pcv[i][0] || '').trim();
    const productName = String(pcv[i][1] || '').trim();
    if (!productName) continue;
    const payload = {
      lines: parseCsv_(pcv[i][2]) || [],
      defects: parseCsv_(pcv[i][3]) || [],
      detectionStages: parseCsv_(pcv[i][4]) || [],
      joints: [],
      images: {
        image1: String(pcv[i][6] || ''),
        image2: String(pcv[i][7] || '')
      }
    };
    try {
      payload.joints = JSON.parse(String(pcv[i][5] || '[]')) || [];
    } catch (_) {
      payload.joints = [];
    }
    if (location) {
      if (!config.productCatalogByLocation[location]) config.productCatalogByLocation[location] = {};
      config.productCatalogByLocation[location][productName] = payload;
    } else {
      config.productCatalog[productName] = payload;
    }
  }
  applySetupSheetOverrides_(config, getRowsAsObjects_(setupSheetSh));
  rebuildPlantsFromLocations_(config);
  return { config: config };
}

function rebuildPlantsFromLocations_(config) {
  config.plants = config.plants || {};
  Object.keys(config.locations || {}).forEach((loc) => {
    Object.keys((config.locations[loc] && config.locations[loc].plants) || {}).forEach((plant) => {
      config.plants[plant] = config.plants[plant] || { locations: [], lines: [] };
      if (config.plants[plant].locations.indexOf(loc) < 0) config.plants[plant].locations.push(loc);
      const lines = (config.locations[loc].plants[plant] && config.locations[loc].plants[plant].lines) || [];
      lines.forEach((ln) => {
        if (config.plants[plant].lines.indexOf(ln) < 0) config.plants[plant].lines.push(ln);
      });
    });
  });
}

function applySetupSheetOverrides_(config, setupRows) {
  (setupRows || []).forEach((raw) => {
    const section = String(raw.section || raw.entity || raw.type || '').trim().toLowerCase();
    const email = normalizeEmail_(raw.email || '');
    const role = String(raw.role || 'employee').trim();
    const allowedPlants = parseCsv_(raw.allowedPlants || raw.allowed_plants) || ['*'];
    const allowedLocations = parseCsv_(raw.allowedLocations || raw.allowed_locations) || ['*'];
    const canSubmitDefects = String(raw.canSubmitDefects || raw.can_submit_defects || 'true').toLowerCase() !== 'false';
    const location = String(raw.location || '').trim();
    const plant = String(raw.plant || '').trim();
    const line = String(raw.line || '').trim();
    const productName = String(raw.productName || raw.product || '').trim();
    const defectsCsv = String(raw.defectsCsv || raw.defects || '').trim();
    const linesCsv = String(raw.linesCsv || raw.lines || '').trim();
    const stagesCsv = String(raw.stagesCsv || raw.stages || '').trim();
    const jointsJson = String(raw.jointsJson || raw.joints || '').trim();
    const image1 = String(raw.image1 || '').trim();
    const image2 = String(raw.image2 || '').trim();
    const settingKey = String(raw.settingKey || raw.key || '').trim();
    const settingValue = String(raw.settingValue || raw.value || '').trim();

    const inferredSection = section
      || (email ? 'user' : '')
      || (settingKey ? 'ui' : '')
      || (productName ? 'product' : '')
      || (location && plant ? 'hierarchy' : '')
      || (location && defectsCsv ? 'defect' : '');

    if (inferredSection === 'user') {
      if (!email) return;
      config.users[email.replace(/\./g, ',')] = {
        email,
        role,
        allowedPlants,
        allowedLocations,
        canSubmitDefects,
        canManageScanner: ['it_admin', 'full_access'].includes(role)
      };
      return;
    }

    if (inferredSection === 'ui') {
      if (!settingKey) return;
      config.uiSecurity = config.uiSecurity || { settingsPin: '1234', excelPin: '1234' };
      if (settingKey === 'settingsPin' || settingKey === 'excelPin') {
        config.uiSecurity[settingKey] = settingValue || config.uiSecurity[settingKey];
      }
      return;
    }

    if (inferredSection === 'hierarchy') {
      if (!location || !plant) return;
      config.locations[location] = config.locations[location] || { plants: {} };
      config.locations[location].plants[plant] = config.locations[location].plants[plant] || { lines: [] };
      if (line && config.locations[location].plants[plant].lines.indexOf(line) < 0) {
        config.locations[location].plants[plant].lines.push(line);
      }
      return;
    }

    if (inferredSection === 'defect') {
      if (!location) return;
      config.defectsByLocation[location] = defectsCsv
        ? defectsCsv.split(',').map((s) => s.trim()).filter(Boolean)
        : ['Assembly Line Defect', 'Headershop Defect', 'HE Shop Defect'];
      return;
    }

    if (inferredSection === 'product') {
      if (!productName) return;
      const payload = {
        lines: parseCsv_(linesCsv) || [],
        defects: parseCsv_(defectsCsv) || [],
        detectionStages: parseCsv_(stagesCsv) || [],
        joints: [],
        images: { image1, image2 }
      };
      try { payload.joints = JSON.parse(jointsJson || '[]') || []; } catch (_) { payload.joints = []; }
      if (location) {
        config.productCatalogByLocation[location] = config.productCatalogByLocation[location] || {};
        config.productCatalogByLocation[location][productName] = payload;
      } else {
        config.productCatalog[productName] = payload;
      }
    }
  });
}

function sendOtp_(payload) {
  ensureBootstrapAdmins_();
  ensureOtpSenderUiSettings_();
  const email = normalizeEmail_(payload.email);
  const source = String(payload.source || '').trim().toLowerCase();
  if (!email) throw new Error('Email required');

  const profile = resolveUserProfileForSource_(email, source);
  if (!profile) {
    throw new Error('This email is not authorized. Please contact IT Admin.');
  }

  const otp = generateOtp_();
  const expiresAt = Date.now() + OTP_TTL_MS;
  const otpHash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, otp));

  appendRow_(TAB.otp,
    ['createdAt', 'email', 'otpHash', 'otpPlain', 'expiresAt', 'used', 'attempts'],
    [new Date(), email, otpHash, otp, new Date(expiresAt), 'false', 0]
  );

  const html = [
    '<div style="font-family:Segoe UI,Arial,sans-serif;background:#f3f4f6;padding:24px">',
    '  <div style="max-width:520px;margin:0 auto;background:#fff;border-radius:14px;overflow:hidden;border:1px solid #e5e7eb">',
    '    <div style="background:#0f172a;color:#fff;padding:16px 18px;font-weight:800;letter-spacing:0.2px;text-align:center">',
    '      AC Leakage Monitoring',
    '    </div>',
    '    <div style="padding:18px 18px 8px;color:#111827">',
    '      <div style="font-size:14px;color:#374151;margin-bottom:10px">Hello,</div>',
    '      <div style="font-size:14px;color:#374151;margin-bottom:14px">Your login verification code is:</div>',
    '      <div style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:12px;padding:18px;text-align:center">',
    '        <div style="font-size:38px;font-weight:900;letter-spacing:10px;color:#111827">' + otp + '</div>',
    '      </div>',
    '      <div style="font-size:12px;color:#6b7280;margin-top:12px">Valid for 10 minutes. Do not share with anyone.</div>',
    '    </div>',
    '    <div style="padding:12px 18px 18px;color:#6b7280;font-size:12px">',
    '      If you did not request this, please ignore this email.',
    '    </div>',
    '  </div>',
    '</div>'
  ].join('');
  sendOtpEmail_(email, 'AC Leakage Monitoring — Login Verification Code', html);

  logActivity_({ email: email, role: profile.role, activity: 'otp_sent', source: source || 'code.gs' });
  return { message: 'OTP sent' };
}

function verifyOtp_(payload) {
  ensureBootstrapAdmins_();
  const email = normalizeEmail_(payload.email);
  const source = String(payload.source || '').trim().toLowerCase();
  const otp = String(payload.otp || '').trim();
  if (!email || !otp) throw new Error('Email and OTP required');

  const profile = resolveUserProfileForSource_(email, source);
  if (!profile) {
    throw new Error('This email is not authorized. Please contact IT Admin.');
  }

  const sh = ensureTab_(TAB.otp, ['createdAt', 'email', 'otpHash', 'otpPlain', 'expiresAt', 'used', 'attempts']);
  const values = sh.getDataRange().getValues();
  let targetRow = -1;
  for (let i = values.length - 1; i >= 1; i--) {
    if (normalizeEmail_(values[i][1]) === email) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow < 0) throw new Error('OTP not found');

  const row = sh.getRange(targetRow, 1, 1, 7).getValues()[0];
  const expiresAt = new Date(row[4]).getTime();
  const used = String(row[5]).toLowerCase() === 'true';
  const attempts = Number(row[6] || 0);
  if (used) throw new Error('OTP already used');
  if (Date.now() > expiresAt) throw new Error('OTP expired');
  if (attempts >= 5) throw new Error('Too many attempts');

  const incomingHash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, otp));
  const dbHash = String(row[2] || '');
  if (incomingHash !== dbHash) {
    sh.getRange(targetRow, 7).setValue(attempts + 1);
    throw new Error('Invalid OTP');
  }

  sh.getRange(targetRow, 6).setValue('true');
  const token = randomToken_();
  const tokenExpiresAt = Date.now() + SESSION_TTL_MS;

  logActivity_({ email: email, role: profile.role, activity: 'otp_verified', source: source || 'code.gs' });
  return {
    token: token,
    tokenExpiresAt: tokenExpiresAt,
    profile: profile,
    user: { email: email }
  };
}

function logActivity_(payload) {
  appendRow_(TAB.activity,
    ['timestamp', 'email', 'role', 'activity', 'source', 'meta'],
    [
      new Date(),
      normalizeEmail_(payload.email),
      String(payload.role || ''),
      String(payload.activity || ''),
      String(payload.source || 'web'),
      JSON.stringify(payload.meta || {})
    ]
  );
  return { logged: true };
}

function syncDefectReport_(payload) {
  const report = payload.report || {};
  const scope = payload.scope || {};
  const qty = Number(report.quantity);
  const row = [
    new Date(report.timestamp || Date.now()),
    String(scope.plant || report.plant || ''),
    String(scope.location || report.location || ''),
    String(report.productionLine || ''),
    String(report.type || ''),
    String(report.defectType || ''),
    String(report.joint || ''),
    String(report.severity || ''),
    String(report.shift || ''),
    String(report.action || ''),
    String(report.operatorName || ''),
    String(report.reportedBy || ''),
    Number.isFinite(qty) && qty > 0 ? qty : 1,
    JSON.stringify(report)
  ];
  const headers = ['timestamp', 'plant', 'location', 'line', 'type', 'defectType', 'joint', 'severity', 'shift', 'action', 'operatorName', 'reportedBy', 'quantity', 'rawJson'];
  appendRow_(TAB.reports, headers, row);
  const location = String(scope.location || report.location || '').trim();
  if (location) {
    const tabName = getLocationReportTabName_(location);
    appendRow_(tabName, headers, row);
  }
  return { synced: true };
}

function getLocationReportTabName_(location) {
  const safe = String(location || 'unknown')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return 'reports_' + (safe || 'unknown');
}

function getActivityLogs_(payload) {
  const limit = Math.max(1, Math.min(500, Number(payload.limit || 100)));
  const sh = ensureTab_(TAB.activity, ['timestamp', 'email', 'role', 'activity', 'source', 'meta']);
  const values = sh.getDataRange().getValues();
  const rows = [];
  for (let i = values.length - 1; i >= 1 && rows.length < limit; i--) {
    const activity = String(values[i][3] || '');
    if (
      activity.indexOf('login') >= 0 ||
      activity.indexOf('otp') >= 0 ||
      activity.indexOf('submit') >= 0 ||
      activity.indexOf('settings') >= 0
    ) {
      rows.push({
        timestamp: values[i][0] ? new Date(values[i][0]).toISOString() : '',
        email: normalizeEmail_(values[i][1]),
        role: String(values[i][2] || ''),
        activity: activity,
        source: String(values[i][4] || ''),
        meta: String(values[i][5] || '')
      });
    }
  }
  return { logs: rows };
}

function getDefectReports_(payload) {
  const limit = Math.max(1, Math.min(2000, Number(payload.limit || 1000)));
  const sh = ensureTab_(TAB.reports, ['timestamp', 'plant', 'location', 'line', 'type', 'defectType', 'joint', 'severity', 'shift', 'action', 'operatorName', 'reportedBy', 'quantity', 'rawJson']);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = values.length - 1; i >= 1 && out.length < limit; i--) {
    let qty = Number(values[i][12]);
    if (!Number.isFinite(qty) || qty <= 0) {
      // Old rows: column 13 holds rawJson (pre-fix layout had no quantity column)
      try {
        const raw1 = JSON.parse(String(values[i][12] || '{}'));
        qty = Number(raw1.quantity);
      } catch (_) {}
    }
    if (!Number.isFinite(qty) || qty <= 0) {
      // New rows: column 14 holds rawJson
      try {
        const raw2 = JSON.parse(String(values[i][13] || '{}'));
        qty = Number(raw2.quantity);
      } catch (_) {}
    }
    if (!Number.isFinite(qty) || qty <= 0) qty = 1;
    out.push({
      timestamp: values[i][0] ? new Date(values[i][0]).toISOString() : '',
      plant: String(values[i][1] || ''),
      location: String(values[i][2] || ''),
      productionLine: String(values[i][3] || ''),
      type: String(values[i][4] || ''),
      defectType: String(values[i][5] || ''),
      joint: String(values[i][6] || ''),
      severity: String(values[i][7] || ''),
      shift: String(values[i][8] || ''),
      action: String(values[i][9] || ''),
      operatorName: String(values[i][10] || ''),
      reportedBy: String(values[i][11] || ''),
      quantity: qty
    });
  }
  return { reports: out };
}

function syncSettings_(payload) {
  const settings = payload.settings || {};
  const actorEmail = normalizeEmail_(payload.actorEmail || '');

  ensureTab_(TAB.settings, ['timestamp', 'actorEmail', 'rawJson'])
    .appendRow([new Date(), actorEmail, JSON.stringify(settings)]);

  const plants = settings.plants || {};
  const users = settings.users || {};
  const adminUsers = settings.adminUsers || {};
  const employeeUsers = settings.employeeUsers || {};
  const defectMaster = settings.defectMaster || {};

  const plantsSh = ensureTab_(TAB.plants, ['plant', 'updatedAt']);
  const locationsSh = ensureTab_(TAB.locations, ['plant', 'location', 'updatedAt']);
  const linesSh = ensureTab_(TAB.lines, ['plant', 'line', 'updatedAt']);
  const defectsSh = ensureTab_(TAB.defects, ['defectKey', 'configJson', 'updatedAt']);
  const productCatalogSh = ensureTab_(TAB.productCatalog, ['location', 'productName', 'linesCsv', 'defectsCsv', 'stagesCsv', 'jointsJson', 'image1', 'image2', 'updatedAt']);
  const uiSettingsSh = ensureTab_(TAB.uiSettings, ['key', 'value', 'updatedAt']);
  const adminsSh = ensureTab_(TAB.admins, ['email', 'role', 'allowedPlants', 'allowedLocations', 'canSubmitDefects', 'updatedAt']);
  const employeesSh = ensureTab_(TAB.employees, ['email', 'role', 'allowedPlants', 'allowedLocations', 'canSubmitDefects', 'updatedAt']);

  clearDataRows_(plantsSh);
  clearDataRows_(locationsSh);
  clearDataRows_(linesSh);
  clearDataRows_(defectsSh);
  clearDataRows_(productCatalogSh);
  clearDataRows_(uiSettingsSh);
  clearDataRows_(adminsSh);
  clearDataRows_(employeesSh);

  Object.keys(plants).forEach((plant) => {
    plantsSh.appendRow([plant, new Date()]);
    const locations = (plants[plant] && plants[plant].locations) || [];
    const lines = (plants[plant] && plants[plant].lines) || [];
    locations.forEach((location) => locationsSh.appendRow([plant, location, new Date()]));
    lines.forEach((line) => linesSh.appendRow([plant, line, new Date()]));
  });

  const writeUser = (u, fallbackKey, forceRole) => {
    const row = [
      normalizeEmail_(u.email || String(fallbackKey || '').replace(/,/g, '.')),
      forceRole || u.role || 'employee',
      (u.allowedPlants || ['*']).join(','),
      (u.allowedLocations || ['*']).join(','),
      String(u.canSubmitDefects !== false),
      new Date()
    ];
    if (String(row[1] || '').includes('admin') || row[1] === 'it_admin' || row[1] === 'full_access') adminsSh.appendRow(row);
    else employeesSh.appendRow(row);
  };

  Object.keys(adminUsers).forEach((key) => {
    writeUser(adminUsers[key] || {}, key, (adminUsers[key] && adminUsers[key].role) || 'admin');
  });
  Object.keys(employeeUsers).forEach((key) => {
    writeUser(employeeUsers[key] || {}, key, 'employee');
  });
  if (!Object.keys(adminUsers).length && !Object.keys(employeeUsers).length) Object.keys(users).forEach((key) => {
    const u = users[key] || {};
    writeUser(u, key, u.role || 'employee');
  });

  Object.keys(defectMaster).forEach((k) => {
    defectsSh.appendRow([k, JSON.stringify(defectMaster[k] || {}), new Date()]);
  });

  const defectsByLoc = settings.defectsByLocation || {};
  const dbLocSh = ensureTab_('defects_by_location', ['location', 'defectsCsv', 'updatedAt']);
  clearDataRows_(dbLocSh);
  Object.keys(defectsByLoc).forEach((loc) => {
    const arr = defectsByLoc[loc] || [];
    dbLocSh.appendRow([loc, arr.join(','), new Date()]);
  });
  const globalCatalog = settings.productCatalog || {};
  Object.keys(globalCatalog).forEach((productName) => {
    const p = globalCatalog[productName] || {};
    productCatalogSh.appendRow([
      '',
      productName,
      (p.lines || []).join(','),
      (p.defects || []).join(','),
      (p.detectionStages || []).join(','),
      JSON.stringify(p.joints || []),
      (p.images && p.images.image1) || '',
      (p.images && p.images.image2) || '',
      new Date()
    ]);
  });
  const locationCatalog = settings.productCatalogByLocation || {};
  Object.keys(locationCatalog).forEach((loc) => {
    const bucket = locationCatalog[loc] || {};
    Object.keys(bucket).forEach((productName) => {
      const p = bucket[productName] || {};
      productCatalogSh.appendRow([
        loc,
        productName,
        (p.lines || []).join(','),
        (p.defects || []).join(','),
        (p.detectionStages || []).join(','),
        JSON.stringify(p.joints || []),
        (p.images && p.images.image1) || '',
        (p.images && p.images.image2) || '',
        new Date()
      ]);
    });
  });
  const reportHeaders = ['timestamp', 'plant', 'location', 'line', 'type', 'defectType', 'joint', 'severity', 'shift', 'action', 'operatorName', 'reportedBy', 'quantity', 'rawJson'];
  Object.keys(settings.locations || {}).forEach((loc) => {
    ensureTab_(getLocationReportTabName_(loc), reportHeaders);
  });
  const uiSecurity = settings.uiSecurity || {};
  uiSettingsSh.appendRow(['settingsPin', String(uiSecurity.settingsPin || '1234'), new Date()]);
  uiSettingsSh.appendRow(['excelPin', String(uiSecurity.excelPin || '1234'), new Date()]);

  logActivity_({
    email: actorEmail,
    role: 'it_admin',
    activity: 'settings_synced',
    source: 'code.gs',
    meta: { plants: Object.keys(plants).length, users: Object.keys(users).length }
  });

  return { synced: true };
}

function clearDataRows_(sh) {
  const rows = sh.getLastRow();
  if (rows > 1) sh.getRange(2, 1, rows - 1, sh.getMaxColumns()).clearContent();
}

function applySetupSheet_(payload) {
  const actorEmail = normalizeEmail_(payload.actorEmail || 'setup_sheet@system');
  const cfg = getConfig_().config;
  const result = syncSettings_({ settings: cfg, actorEmail: actorEmail });
  logActivity_({
    email: actorEmail,
    role: 'it_admin',
    activity: 'setup_sheet_applied',
    source: 'code.gs'
  });
  return Object.assign({ applied: true }, result || {});
}

// Manual runner (visible in Apps Script function dropdown)
function applySetupSheet() {
  return applySetupSheet_({ actorEmail: 'manual_setup_sheet@script' });
}

// ==========================================
// MONTHLY PLANT-WISE AUTO REPORT SYSTEM
// ==========================================

function getMonthlyReportSettings_(payload) {
  const sh = ensureTab_(TAB.monthlySettings, ['key', 'value', 'updatedAt']);
  const values = sh.getDataRange().getValues();
  const settings = {
    enabled: true,
    defaultTo: DEFAULT_OTP_SENDER_EMAIL,
    defaultCc: '',
    locationRoutes: {},
    plantRoutes: {},
    sendDayOfMonth: 1,
    sendHour: 8
  };

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const val = String(values[i][1] || '').trim();
    if (!key) continue;

    if (key === 'enabled') settings.enabled = val.toLowerCase() !== 'false';
    else if (key === 'defaultTo') settings.defaultTo = val || DEFAULT_OTP_SENDER_EMAIL;
    else if (key === 'defaultCc') settings.defaultCc = val;
    else if (key === 'sendDayOfMonth') settings.sendDayOfMonth = Number(val) || 1;
    else if (key === 'sendHour') settings.sendHour = Number(val) || 8;
    else if (key.indexOf('loc_') === 0) {
      const locName = key.substring(4);
      try {
        settings.locationRoutes[locName] = JSON.parse(val);
      } catch (_) {
        settings.locationRoutes[locName] = { to: val, cc: '' };
      }
    }
    else if (key.indexOf('plant_') === 0) {
      const plantName = key.substring(6);
      try {
        settings.plantRoutes[plantName] = JSON.parse(val);
      } catch (_) {
        settings.plantRoutes[plantName] = { to: val, cc: '' };
      }
    }
  }
  return { settings: settings };
}

function saveMonthlyReportSettings_(payload) {
  const settings = payload.settings || {};
  const actorEmail = normalizeEmail_(payload.actorEmail || '');
  const sh = ensureTab_(TAB.monthlySettings, ['key', 'value', 'updatedAt']);
  clearDataRows_(sh);

  sh.appendRow(['enabled', String(settings.enabled !== false), new Date()]);
  sh.appendRow(['defaultTo', String(settings.defaultTo || DEFAULT_OTP_SENDER_EMAIL), new Date()]);
  sh.appendRow(['defaultCc', String(settings.defaultCc || ''), new Date()]);
  sh.appendRow(['sendDayOfMonth', String(settings.sendDayOfMonth || 1), new Date()]);
  sh.appendRow(['sendHour', String(settings.sendHour || 8), new Date()]);

  const locRoutes = settings.locationRoutes || {};
  Object.keys(locRoutes).forEach(function(loc) {
    const r = locRoutes[loc] || {};
    sh.appendRow(['loc_' + loc, JSON.stringify(r), new Date()]);
  });

  const routes = settings.plantRoutes || {};
  Object.keys(routes).forEach(function(plant) {
    const r = routes[plant] || {};
    sh.appendRow(['plant_' + plant, JSON.stringify(r), new Date()]);
  });

  logActivity_({
    email: actorEmail,
    role: 'it_admin',
    activity: 'monthly_report_settings_saved',
    source: 'code.gs',
    meta: {
      enabled: settings.enabled,
      locationsConfigured: Object.keys(locRoutes).length,
      plantsConfigured: Object.keys(routes).length
    }
  });

  return { saved: true, settings: settings };
}

function escapeHtml_(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function computePlantMetrics_(records) {
  records = records || [];
  let totalLeaks = 0;
  let criticalCount = 0;
  let majorCount = 0;
  let minorCount = 0;
  let reworkCount = 0;
  let scrapCount = 0;
  let acceptedCount = 0;

  const defectCounts = {};
  const lineCounts = {};
  const shiftCounts = {};
  const unitCounts = {};
  const jointCounts = {};

  records.forEach(function(r) {
    const qty = Number(r.quantity) > 0 ? Number(r.quantity) : 1;
    totalLeaks += qty;
    const sev = String(r.severity || '').toLowerCase();
    if (sev === 'critical') criticalCount += qty;
    else if (sev === 'major') majorCount += qty;
    else minorCount += qty;

    const act = String(r.action || '').toLowerCase();
    if (act.indexOf('rework') >= 0) reworkCount += qty;
    else if (act.indexOf('scrap') >= 0) scrapCount += qty;
    else if (act.indexOf('accept') >= 0) acceptedCount += qty;

    const def = String(r.defectType || 'Other').trim();
    defectCounts[def] = (defectCounts[def] || 0) + qty;

    const line = String(r.line || r.productionLine || 'Unknown').trim();
    lineCounts[line] = (lineCounts[line] || 0) + qty;

    const shift = String(r.shift || 'General').trim();
    shiftCounts[shift] = (shiftCounts[shift] || 0) + qty;

    const unit = String(r.type || r.unitType || r.unit || 'Other').trim();
    unitCounts[unit] = (unitCounts[unit] || 0) + qty;

    const joint = String(r.joint || r.jointNo || '').trim();
    if (joint) {
      jointCounts[joint] = (jointCounts[joint] || 0) + qty;
    }
  });

  const sortedDefects = Object.keys(defectCounts).map(function(k) {
    return { name: k, count: defectCounts[k], pct: totalLeaks > 0 ? Math.round((defectCounts[k] / totalLeaks) * 100) : 0 };
  }).sort(function(a, b) { return b.count - a.count; });

  const sortedUnits = Object.keys(unitCounts).map(function(k) {
    return { name: k, count: unitCounts[k], pct: totalLeaks > 0 ? Math.round((unitCounts[k] / totalLeaks) * 100) : 0 };
  }).sort(function(a, b) { return b.count - a.count; });

  const sortedLines = Object.keys(lineCounts).map(function(k) {
    return { name: k, count: lineCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; });

  const sortedShifts = Object.keys(shiftCounts).map(function(k) {
    return { name: k, count: shiftCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; });

  const sortedJoints = Object.keys(jointCounts).map(function(k) {
    return { name: k, count: jointCounts[k] };
  }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);

  return {
    totalLeaks: totalLeaks,
    criticalCount: criticalCount,
    majorCount: majorCount,
    minorCount: minorCount,
    reworkCount: reworkCount,
    scrapCount: scrapCount,
    acceptedCount: acceptedCount,
    sortedDefects: sortedDefects,
    sortedUnits: sortedUnits,
    sortedLines: sortedLines,
    sortedShifts: sortedShifts,
    sortedJoints: sortedJoints
  };
}

function buildSinglePlantSectionHtml_(grp, monthLabel, isMultiPlant) {
  const plantName = escapeHtml_(grp.plant || 'PGTL');
  const locationName = escapeHtml_(grp.location || 'Pune');
  const records = grp.records || [];
  const metrics = computePlantMetrics_(records);
  const totalLeaks = metrics.totalLeaks;

  if (totalLeaks === 0) {
    return [
      '<div style="background:#f0fdf4;border:1px solid #86efac;border-left:5px solid #22c55e;border-radius:10px;padding:22px 24px;margin-bottom:18px;">',
      '  <div style="display:inline-block;background:#22c55e;color:#ffffff;font-size:11px;font-weight:900;padding:3px 10px;border-radius:999px;letter-spacing:0.8px;text-transform:uppercase;margin-bottom:10px;">',
      '    ZERO DEFECT STATUS',
      '  </div>',
      '  <div style="font-size:16px;font-weight:800;color:#15803d;line-height:1.4;margin-bottom:8px;">',
      '    <span style="color:#16a34a;font-weight:900;margin-right:6px;">&#10003;</span> No leakage entries have been recorded in the AC Leakage Monitoring System for this month.',
      '  </div>',
      '  <div style="font-size:13px;color:#166534;line-height:1.6;">',
      '    During the reporting month of <b>' + escapeHtml_(monthLabel) + '</b>, all production lines and shifts operated completely leak-free with <b>0 defect entries logged</b> for <b>' + plantName + ' (' + locationName + ')</b>.',
      '  </div>',
      '  <div style="margin-top:14px;padding-top:12px;border-top:1px solid #bbf7d0;display:flex;gap:18px;font-size:12px;color:#15803d;">',
      '    <span><b>Plant:</b> ' + plantName + '</span>',
      '    <span><b>Location:</b> ' + locationName + '</span>',
      '    <span><b>Monthly Incidents:</b> 0 (Clean Operation)</span>',
      '  </div>',
      '</div>'
    ].join('\n');
  }

  // Active defects tables
  let defectRowsHtml = '';
  metrics.sortedDefects.forEach(function(d, idx) {
    const barColor = idx === 0 ? '#4f46e5' : idx === 1 ? '#06b6d4' : '#64748b';
    defectRowsHtml += [
      '<tr style="border-bottom:1px solid #f1f5f9;">',
      '  <td style="padding:10px 12px;color:#1e293b;font-weight:600;font-size:13px;">' + escapeHtml_(d.name) + '</td>',
      '  <td style="padding:10px 12px;text-align:center;font-weight:700;color:#0f172a;font-size:13px;">' + d.count + '</td>',
      '  <td style="padding:10px 12px;width:140px;">',
      '    <div style="display:flex;align-items:center;gap:8px;">',
      '      <div style="flex:1;background:#e2e8f0;border-radius:999px;height:8px;overflow:hidden;">',
      '        <div style="background:' + barColor + ';width:' + Math.min(100, d.pct) + '%;height:8px;border-radius:999px;"></div>',
      '      </div>',
      '      <span style="font-size:11px;font-weight:700;color:#475569;width:32px;text-align:right;">' + d.pct + '%</span>',
      '    </div>',
      '  </td>',
      '</tr>'
    ].join('');
  });

  let unitRowsHtml = '';
  metrics.sortedUnits.forEach(function(u, idx) {
    const barColor = idx === 0 ? '#10b981' : idx === 1 ? '#3b82f6' : '#8b5cf6';
    unitRowsHtml += [
      '<tr style="border-bottom:1px solid #f1f5f9;">',
      '  <td style="padding:9px 12px;color:#1e293b;font-weight:600;font-size:13px;">' + escapeHtml_(u.name) + '</td>',
      '  <td style="padding:9px 12px;text-align:center;font-weight:700;color:#0f172a;font-size:13px;">' + u.count + '</td>',
      '  <td style="padding:9px 12px;width:120px;">',
      '    <div style="display:flex;align-items:center;gap:6px;">',
      '      <div style="flex:1;background:#e2e8f0;border-radius:999px;height:7px;overflow:hidden;">',
      '        <div style="background:' + barColor + ';width:' + Math.min(100, u.pct) + '%;height:7px;border-radius:999px;"></div>',
      '      </div>',
      '      <span style="font-size:11px;font-weight:700;color:#475569;width:30px;text-align:right;">' + u.pct + '%</span>',
      '    </div>',
      '  </td>',
      '</tr>'
    ].join('');
  });

  let lineRowsHtml = '';
  metrics.sortedLines.forEach(function(l) {
    lineRowsHtml += [
      '<tr style="border-bottom:1px solid #f1f5f9;">',
      '  <td style="padding:9px 12px;color:#334155;font-weight:500;font-size:13px;">' + escapeHtml_(l.name) + '</td>',
      '  <td style="padding:9px 12px;text-align:right;font-weight:700;color:#0f172a;font-size:13px;">' + l.count + '</td>',
      '</tr>'
    ].join('');
  });

  let shiftRowsHtml = '';
  metrics.sortedShifts.forEach(function(s) {
    shiftRowsHtml += [
      '<tr style="border-bottom:1px solid #f1f5f9;">',
      '  <td style="padding:9px 12px;color:#334155;font-weight:500;font-size:13px;">' + escapeHtml_(s.name) + '</td>',
      '  <td style="padding:9px 12px;text-align:right;font-weight:700;color:#0f172a;font-size:13px;">' + s.count + '</td>',
      '</tr>'
    ].join('');
  });

  let jointRowsHtml = '';
  if (metrics.sortedJoints.length > 0) {
    metrics.sortedJoints.forEach(function(j) {
      jointRowsHtml += [
        '<tr style="border-bottom:1px solid #f1f5f9;">',
        '  <td style="padding:8px 12px;color:#334155;font-weight:600;font-size:12px;">' + escapeHtml_(j.name) + '</td>',
        '  <td style="padding:8px 12px;text-align:right;font-weight:700;color:#dc2626;font-size:12px;">' + j.count + '</td>',
        '</tr>'
      ].join('');
    });
  }

  return [
    '<!-- KPI Metric Cards Grid (6 cards) -->',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:20px;">',
    '  <tr>',
    '    <td width="33.33%" style="padding:4px;">',
    '      <div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:12px 8px;text-align:center;">',
    '        <div style="font-size:11px;font-weight:700;color:#1d4ed8;text-transform:uppercase;">TOTAL LEAKS</div>',
    '        <div style="font-size:24px;font-weight:900;color:#1e3a8a;margin-top:2px;">' + totalLeaks + '</div>',
    '      </div>',
    '    </td>',
    '    <td width="33.33%" style="padding:4px;">',
    '      <div style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:12px 8px;text-align:center;">',
    '        <div style="font-size:11px;font-weight:700;color:#b91c1c;text-transform:uppercase;">CRITICAL</div>',
    '        <div style="font-size:24px;font-weight:900;color:#991b1b;margin-top:2px;">' + metrics.criticalCount + '</div>',
    '      </div>',
    '    </td>',
    '    <td width="33.33%" style="padding:4px;">',
    '      <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;padding:12px 8px;text-align:center;">',
    '        <div style="font-size:11px;font-weight:700;color:#c2410c;text-transform:uppercase;">MAJOR</div>',
    '        <div style="font-size:24px;font-weight:900;color:#9a3412;margin-top:2px;">' + metrics.majorCount + '</div>',
    '      </div>',
    '    </td>',
    '  </tr>',
    '  <tr>',
    '    <td width="33.33%" style="padding:4px;">',
    '      <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:12px 8px;text-align:center;">',
    '        <div style="font-size:11px;font-weight:700;color:#15803d;text-transform:uppercase;">MINOR</div>',
    '        <div style="font-size:24px;font-weight:900;color:#166534;margin-top:2px;">' + metrics.minorCount + '</div>',
    '      </div>',
    '    </td>',
    '    <td width="33.33%" style="padding:4px;">',
    '      <div style="background:#fefce8;border:1px solid #fef08a;border-radius:10px;padding:12px 8px;text-align:center;">',
    '        <div style="font-size:11px;font-weight:700;color:#a16207;text-transform:uppercase;">REWORK</div>',
    '        <div style="font-size:24px;font-weight:900;color:#854d0e;margin-top:2px;">' + metrics.reworkCount + '</div>',
    '      </div>',
    '    </td>',
    '    <td width="33.33%" style="padding:4px;">',
    '      <div style="background:#fdf2f8;border:1px solid #fbcfe8;border-radius:10px;padding:12px 8px;text-align:center;">',
    '        <div style="font-size:11px;font-weight:700;color:#be185d;text-transform:uppercase;">SCRAP</div>',
    '        <div style="font-size:24px;font-weight:900;color:#9d174d;margin-top:2px;">' + metrics.scrapCount + '</div>',
    '      </div>',
    '    </td>',
    '  </tr>',
    '</table>',
    '',
    '<!-- Unit Type & Defect Breakdown -->',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:20px;">',
    '  <tr>',
    '    <td width="50%" valign="top" style="padding-right:8px;">',
    '      <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Unit Type Breakdown</div>',
    '      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '        <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '          <tr>',
    '            <th style="padding:8px 10px;text-align:left;">Unit Type</th>',
    '            <th style="padding:8px 6px;text-align:center;">Count</th>',
    '            <th style="padding:8px 10px;text-align:left;">Share</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody>' + (unitRowsHtml || '<tr><td colspan="3" style="padding:10px;text-align:center;color:#64748b;">No unit data.</td></tr>') + '</tbody>',
    '      </table>',
    '    </td>',
    '    <td width="50%" valign="top" style="padding-left:8px;">',
    '      <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Defect Categories</div>',
    '      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '        <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '          <tr>',
    '            <th style="padding:8px 10px;text-align:left;">Category</th>',
    '            <th style="padding:8px 6px;text-align:center;">Count</th>',
    '            <th style="padding:8px 10px;text-align:left;">Share</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody>' + (defectRowsHtml || '<tr><td colspan="3" style="padding:10px;text-align:center;color:#64748b;">No defect data.</td></tr>') + '</tbody>',
    '      </table>',
    '    </td>',
    '  </tr>',
    '</table>',
    '',
    '<!-- Line Breakdown & Shift Tables -->',
    '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:18px;">',
    '  <tr>',
    '    <td width="50%" valign="top" style="padding-right:8px;">',
    '      <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Line Breakdown</div>',
    '      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '        <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '          <tr>',
    '            <th style="padding:8px 12px;text-align:left;">Line</th>',
    '            <th style="padding:8px 12px;text-align:right;">Leaks</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody>' + (lineRowsHtml || '<tr><td colspan="2" style="padding:10px;text-align:center;color:#64748b;">No line records.</td></tr>') + '</tbody>',
    '      </table>',
    '    </td>',
    '    <td width="50%" valign="top" style="padding-left:8px;">',
    '      <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Shift Distribution</div>',
    '      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '        <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '          <tr>',
    '            <th style="padding:8px 12px;text-align:left;">Shift</th>',
    '            <th style="padding:8px 12px;text-align:right;">Leaks</th>',
    '          </tr>',
    '        </thead>',
    '        <tbody>' + (shiftRowsHtml || '<tr><td colspan="2" style="padding:10px;text-align:center;color:#64748b;">No shift records.</td></tr>') + '</tbody>',
    '      </table>',
    '    </td>',
    '  </tr>',
    '</table>',
    jointRowsHtml ? [
      '<div style="margin-bottom:18px;">',
      '  <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Top Leakage Joints</div>',
      '  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
      '    <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
      '      <tr>',
      '        <th style="padding:8px 12px;text-align:left;">Joint / Location</th>',
      '        <th style="padding:8px 12px;text-align:right;">Leak Frequency</th>',
      '      </tr>',
      '    </thead>',
      '    <tbody>' + jointRowsHtml + '</tbody>',
      '  </table>',
      '</div>'
    ].join('') : ''
  ].join('\n');
}

function buildMonthlyReportHtml_(plantOrGroups, locationName, monthLabel, records, totalLeaks) {
  let groups = [];
  if (Array.isArray(plantOrGroups)) {
    groups = plantOrGroups;
  } else {
    groups = [{
      plant: String(plantOrGroups || 'PGTL').trim(),
      location: String(locationName || 'Pune').trim(),
      records: records || []
    }];
  }

  // Calculate grand totals across all plants in this email
  let grandTotalLeaks = 0;
  let grandCritical = 0;
  let grandMajor = 0;
  let grandMinor = 0;
  let grandRework = 0;
  let grandScrap = 0;
  let grandAccepted = 0;

  const groupMetricsList = groups.map(function(grp) {
    const met = computePlantMetrics_(grp.records || []);
    grandTotalLeaks += met.totalLeaks;
    grandCritical += met.criticalCount;
    grandMajor += met.majorCount;
    grandMinor += met.minorCount;
    grandRework += met.reworkCount;
    grandScrap += met.scrapCount;
    grandAccepted += met.acceptedCount;
    return {
      grp: grp,
      metrics: met
    };
  });

  const isMultiPlant = groups.length > 1;
  const generatedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a');

  // Locations label
  const uniqueLocations = [];
  groups.forEach(function(g) {
    const loc = g.location || 'Pune';
    if (uniqueLocations.indexOf(loc) < 0) uniqueLocations.push(loc);
  });
  const locDisplay = uniqueLocations.join(', ');

  // Plants label
  const plantNames = groups.map(function(g) { return g.plant; });
  const plantDisplay = isMultiPlant ? plantNames.join(', ') : (plantNames[0] || 'PGTL');

  // Header Subtitle Badges
  const locBadge = '<span style="display:inline-block;background-color:#334155;border:1px solid #475569;border-radius:6px;padding:5px 12px;font-size:12px;font-weight:700;color:#ffffff;margin-right:8px;margin-bottom:6px;"><b style="color:#93c5fd;">LOCATION:</b> ' + escapeHtml_(locDisplay) + '</span>';
  const plantBadge = '<span style="display:inline-block;background-color:#334155;border:1px solid #475569;border-radius:6px;padding:5px 12px;font-size:12px;font-weight:700;color:#ffffff;margin-right:8px;margin-bottom:6px;"><b style="color:#93c5fd;">' + (isMultiPlant ? 'PLANTS (' + groups.length + '):' : 'PLANT:') + '</b> ' + escapeHtml_(plantDisplay) + '</span>';
  const periodBadge = '<span style="display:inline-block;background-color:#334155;border:1px solid #475569;border-radius:6px;padding:5px 12px;font-size:12px;font-weight:700;color:#ffffff;margin-bottom:6px;"><b style="color:#93c5fd;">PERIOD:</b> ' + escapeHtml_(monthLabel) + '</span>';

  // Multi-Plant Executive Overview Table
  let multiPlantSummaryTable = '';
  if (isMultiPlant) {
    const summaryRows = groupMetricsList.map(function(item) {
      const pName = escapeHtml_(item.grp.plant);
      const lName = escapeHtml_(item.grp.location);
      const m = item.metrics;
      const isClean = m.totalLeaks === 0;
      const statusBadge = isClean
        ? '<span style="display:inline-block;background:#dcfce7;color:#166534;font-weight:700;padding:3px 8px;border-radius:6px;font-size:11px;"><span style="color:#16a34a;font-weight:bold;margin-right:4px;">&#10003;</span>Zero Leakage</span>'
        : '<span style="display:inline-block;background:#fee2e2;color:#991b1b;font-weight:700;padding:3px 8px;border-radius:6px;font-size:11px;">' + m.totalLeaks + ' Leaks Logged</span>';

      return [
        '<tr style="border-bottom:1px solid #e2e8f0;">',
        '  <td style="padding:10px 12px;font-weight:700;color:#0f172a;">' + pName + '</td>',
        '  <td style="padding:10px 8px;color:#475569;">' + lName + '</td>',
        '  <td style="padding:10px 8px;text-align:center;font-weight:800;font-size:13px;color:' + (isClean ? '#166534' : '#1e3a8a') + ';">' + m.totalLeaks + '</td>',
        '  <td style="padding:10px 8px;text-align:center;font-weight:700;color:#991b1b;">' + m.criticalCount + '</td>',
        '  <td style="padding:10px 8px;text-align:center;font-weight:700;color:#9a3412;">' + m.majorCount + '</td>',
        '  <td style="padding:10px 8px;text-align:center;font-weight:700;color:#166534;">' + m.minorCount + '</td>',
        '  <td style="padding:10px 12px;">' + statusBadge + '</td>',
        '</tr>'
      ].join('');
    }).join('\n');

    multiPlantSummaryTable = [
      '<!-- Multi-Plant Executive Overview Table -->',
      '<div style="margin-bottom:26px;">',
      '  <div style="font-size:14px;font-weight:800;color:#0f172a;margin-bottom:10px;border-bottom:2px solid #3b82f6;padding-bottom:6px;display:flex;justify-content:space-between;align-items:center;">',
      '    <span>Executive Multi-Plant Performance Comparison</span>',
      '    <span style="font-size:12px;font-weight:700;color:#3b82f6;">' + groups.length + ' Plants Monitored</span>',
      '  </div>',
      '  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #cbd5e1;border-radius:8px;overflow:hidden;font-size:12px;width:100%;">',
      '    <thead style="background:#1e293b;color:#ffffff;text-transform:uppercase;font-size:11px;">',
      '      <tr>',
      '        <th style="padding:10px 12px;text-align:left;">Plant Name</th>',
      '        <th style="padding:10px 8px;text-align:left;">Location</th>',
      '        <th style="padding:10px 8px;text-align:center;">Total Leaks</th>',
      '        <th style="padding:10px 8px;text-align:center;">Critical</th>',
      '        <th style="padding:10px 8px;text-align:center;">Major</th>',
      '        <th style="padding:10px 8px;text-align:center;">Minor</th>',
      '        <th style="padding:10px 12px;text-align:left;">Monthly Status</th>',
      '      </tr>',
      '    </thead>',
      '    <tbody>',
      summaryRows,
      '      <tr style="background:#f1f5f9;font-weight:800;border-top:2px solid #94a3b8;">',
      '        <td style="padding:10px 12px;color:#0f172a;" colspan="2">TOTAL (ALL PLANTS CONSOLIDATED)</td>',
      '        <td style="padding:10px 8px;text-align:center;font-size:13px;color:#1e3a8a;">' + grandTotalLeaks + '</td>',
      '        <td style="padding:10px 8px;text-align:center;color:#991b1b;">' + grandCritical + '</td>',
      '        <td style="padding:10px 8px;text-align:center;color:#9a3412;">' + grandMajor + '</td>',
      '        <td style="padding:10px 8px;text-align:center;color:#166534;">' + grandMinor + '</td>',
      '        <td style="padding:10px 12px;color:#334155;font-size:11px;">' + (grandTotalLeaks === 0 ? 'All Plants 100% Leak-Free' : grandTotalLeaks + ' Total Leaks') + '</td>',
      '      </tr>',
      '    </tbody>',
      '  </table>',
      '</div>'
    ].join('\n');
  }

  // Render individual sections for each plant
  const plantSectionsHtml = groups.map(function(grp) {
    const pName = escapeHtml_(grp.plant);
    const lName = escapeHtml_(grp.location);
    const innerHtml = buildSinglePlantSectionHtml_(grp, monthLabel, isMultiPlant);

    if (isMultiPlant) {
      return [
        '<div style="margin-bottom:28px;border:1px solid #e2e8f0;border-radius:12px;padding:20px;background:#ffffff;box-shadow:0 2px 8px rgba(0,0,0,0.03);">',
        '  <div style="border-bottom:2px solid #3b82f6;padding-bottom:10px;margin-bottom:16px;">',
        '    <div style="font-size:11px;font-weight:900;color:#2563eb;text-transform:uppercase;letter-spacing:1px;">PLANT QUALITY PROFILE</div>',
        '    <h2 style="margin:2px 0 0 0;font-size:18px;font-weight:800;color:#0f172a;">' + pName + ' <span style="font-size:13px;font-weight:600;color:#64748b;">(' + lName + ')</span></h2>',
        '  </div>',
        innerHtml,
        '</div>'
      ].join('\n');
    }
    return innerHtml;
  }).join('\n');

  // Executive Header Callout Message
  let executiveCallout = '';
  if (isMultiPlant) {
    executiveCallout = [
      '<div style="background:#f8fafc;border-left:4px solid #3b82f6;padding:14px 18px;border-radius:0 8px 8px 0;margin-bottom:22px;">',
      '  <p style="margin:0;font-size:13px;color:#334155;line-height:1.6;">',
      '    Consolidated performance overview for <b>' + escapeHtml_(plantDisplay) + '</b> (' + escapeHtml_(locDisplay) + ') during <b>' + escapeHtml_(monthLabel) + '</b>. Total of <b>' + grandTotalLeaks + '</b> leakage defect(s) logged across all ' + groups.length + ' monitored plants.',
      '  </p>',
      '</div>'
    ].join('');
  } else {
    executiveCallout = [
      '<div style="background:#f8fafc;border-left:4px solid #3b82f6;padding:14px 18px;border-radius:0 8px 8px 0;margin-bottom:22px;">',
      '  <p style="margin:0;font-size:13px;color:#334155;line-height:1.6;">',
      '    Performance overview for <b>' + escapeHtml_(groups[0].plant) + ' (' + escapeHtml_(groups[0].location) + ')</b> during <b>' + escapeHtml_(monthLabel) + '</b>. Total of <b>' + grandTotalLeaks + '</b> leakage defect(s) logged across all production lines and shifts.',
      '  </p>',
      '</div>'
    ].join('');
  }

  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>',
    '<body style="margin:0;padding:0;background-color:#f8fafc;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif;">',
    '  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color:#f8fafc;padding:24px 0;">',
    '    <tr>',
    '      <td align="center">',
    '        <table role="presentation" width="100%" style="max-width:720px;background-color:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 8px 30px rgba(0,0,0,0.06);border:1px solid #e2e8f0;" cellspacing="0" cellpadding="0">',
    '          ',
    '          <!-- Header Banner -->',
    '          <tr>',
    '            <td bgcolor="#1e293b" style="background-color:#1e293b !important;background:#1e293b;padding:28px 24px;text-align:left;color:#ffffff;border-bottom:3px solid #3b82f6;">',
    '              <div style="font-size:12px;font-weight:900;letter-spacing:2px;text-transform:uppercase;color:#93c5fd;margin-bottom:6px;">PG GROUP</div>',
    '              <h1 style="margin:0 0 14px 0;font-size:22px;font-weight:800;color:#ffffff;line-height:1.3;">' + (isMultiPlant ? 'Monthly AC Leakage Quality Report — Consolidated' : 'Monthly AC Leakage Quality Report') + '</h1>',
    '              <div style="display:block;margin-top:10px;">',
    '                ' + locBadge,
    '                ' + plantBadge,
    '                ' + periodBadge,
    '              </div>',
    '            </td>',
    '          </tr>',
    '          ',
    '          <!-- Main Content Container -->',
    '          <tr>',
    '            <td style="padding:24px;">',
    '              ',
    executiveCallout,
    multiPlantSummaryTable,
    plantSectionsHtml,
    '              ',
    '              <!-- PROMINENT ATTACHMENT CALLOUT BANNER -->',
    '              <div style="margin-top:20px;padding:12px 18px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;color:#166534;font-size:12px;line-height:1.5;">',
    '                <div style="font-weight:800;font-size:12px;margin-bottom:3px;color:#15803d;letter-spacing:0.3px;">EXCEL DETAILED REPORT ATTACHED (.xls)</div>',
    '                <div>The complete defect dataset and executive summary workbook has been generated and attached to this email for your offline analysis and record keeping.</div>',
    '              </div>',
    '              ',
    '              <!-- PROMINENT AUTO-GENERATED DISCLAIMER -->',
    '              <div style="margin-top:14px;padding:14px 18px;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;color:#991b1b;font-size:12px;line-height:1.5;">',
    '                <div style="font-weight:800;font-size:12px;margin-bottom:4px;letter-spacing:0.5px;">[ AUTOMATED REPORT NOTICE ]</div>',
    '                <div>This is an <b>auto-generated report</b> produced by the <b>PG Group AC Leakage Monitoring System</b>. Please do not reply directly to this email. For any queries, discrepancies, or routing changes, please contact the Quality Team or IT Admin.</div>',
    '              </div>',
    '              ',
    '            </td>',
    '          </tr>',
    '          ',
    '          <!-- Footer -->',
    '          <tr>',
    '            <td style="background:#f8fafc;border-top:1px solid #e2e8f0;padding:16px 24px;text-align:center;font-size:11px;color:#64748b;">',
    '              Generated on ' + generatedDate + ' | AC Leakage Monitoring Portal | PG Group',
    '            </td>',
    '          </tr>',
    '          ',
    '        </table>',
    '      </td>',
    '    </tr>',
    '  </table>',
    '</body>',
    '</html>'
  ].join('\n');
}

function parseRecordDate_(ts) {
  if (!ts) return null;
  if (ts instanceof Date) return isNaN(ts.getTime()) ? null : ts;
  const s = String(ts).trim();
  if (!s) return null;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  const parts = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})(.*)$/);
  if (parts) {
    const d2 = new Date(Number(parts[3]), Number(parts[2]) - 1, Number(parts[1]));
    if (!isNaN(d2.getTime())) return d2;
  }
  return null;
}

function createMonthlyReportExcelAttachment_(plantOrGroups, locationName, monthLabel, records) {
  let groups = [];
  if (Array.isArray(plantOrGroups)) {
    groups = plantOrGroups;
  } else {
    groups = [{
      plant: String(plantOrGroups || 'PGTL').trim(),
      location: String(locationName || 'Pune').trim(),
      records: records || []
    }];
  }

  const isMulti = groups.length > 1;
  const cleanMonth = String(monthLabel || 'Monthly_Report').replace(/[^a-zA-Z0-9_-]/g, '_');
  const cleanTag = isMulti ? 'Consolidated_MultiPlant' : String(groups[0].plant || 'PGTL').replace(/[^a-zA-Z0-9_-]/g, '_');
  const fileName = 'PG_Group_AC_Leakage_Report_' + cleanTag + '_' + cleanMonth + '.xls';

  const escapeXml = function(val) {
    if (val == null) return '';
    return String(val)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  };

  const defectRowsXml = [];
  let grandTotalQty = 0;
  let grandCritical = 0;
  let grandMajor = 0;
  let grandMinor = 0;
  let grandRework = 0;
  let grandScrap = 0;
  let grandAccepted = 0;

  groups.forEach(function(grp) {
    const pName = grp.plant || 'PGTL';
    const lName = grp.location || 'Pune';
    const recs = grp.records || [];

    if (recs.length === 0) {
      defectRowsXml.push([
        '    <Row ss:Height="20">',
        '      <Cell ss:StyleID="DataCellCenter"><Data ss:Type="String">-</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(lName) + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(pName) + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell" ss:MergeAcross="9"><Data ss:Type="String">No leakage entries have been recorded in the AC Leakage Monitoring System for this month (Zero Defects Logged).</Data></Cell>',
        '    </Row>'
      ].join(''));
      return;
    }

    recs.forEach(function(r) {
      const q = Number(r.quantity) > 0 ? Number(r.quantity) : 1;
      grandTotalQty += q;
      const sev = String(r.severity || '').toLowerCase();
      let sevStyle = 'DataCellCenter';
      if (sev === 'critical') { grandCritical += q; sevStyle = 'CriticalBadge'; }
      else if (sev === 'major') { grandMajor += q; sevStyle = 'MajorBadge'; }
      else { grandMinor += q; sevStyle = 'MinorBadge'; }

      const act = String(r.action || '').toLowerCase();
      if (act.indexOf('rework') >= 0) grandRework += q;
      else if (act.indexOf('scrap') >= 0) grandScrap += q;
      else if (act.indexOf('accept') >= 0) grandAccepted += q;

      let formattedDate = '';
      const dt = parseRecordDate_(r.timestamp);
      if (dt) {
        try {
          formattedDate = Utilities.formatDate(dt, Session.getScriptTimeZone() || 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
        } catch (_) {
          formattedDate = String(r.timestamp);
        }
      } else {
        formattedDate = String(r.timestamp || '');
      }

      defectRowsXml.push([
        '    <Row ss:Height="19">',
        '      <Cell ss:StyleID="DataCellCenter"><Data ss:Type="String">' + escapeXml(formattedDate) + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.location || lName) + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.plant || pName) + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.type || '-') + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.productionLine || '-') + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.defectType || '-') + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.joint || '-') + '</Data></Cell>',
        '      <Cell ss:StyleID="' + sevStyle + '"><Data ss:Type="String">' + escapeXml(r.severity || '-') + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCellCenter"><Data ss:Type="String">' + escapeXml(r.shift || '-') + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.action || '-') + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + q + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.operatorName || '-') + '</Data></Cell>',
        '      <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(r.reportedBy || '-') + '</Data></Cell>',
        '    </Row>'
      ].join(''));
    });
  });

  const generatedTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');

  // Executive summary worksheet rows
  const executiveTableRowsXml = [];
  if (isMulti) {
    executiveTableRowsXml.push([
      '   <Row ss:Height="22">',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Plant Name</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Location</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Total Leaks</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Critical</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Major</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Minor</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Monthly Status</Data></Cell>',
      '   </Row>'
    ].join(''));

    groups.forEach(function(grp) {
      const met = computePlantMetrics_(grp.records || []);
      const statusText = met.totalLeaks === 0 ? 'Zero Leakage Logged' : met.totalLeaks + ' Leaks Logged';
      executiveTableRowsXml.push([
        '   <Row ss:Height="20">',
        '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(grp.plant) + '</Data></Cell>',
        '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(grp.location) + '</Data></Cell>',
        '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + met.totalLeaks + '</Data></Cell>',
        '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + met.criticalCount + '</Data></Cell>',
        '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + met.majorCount + '</Data></Cell>',
        '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + met.minorCount + '</Data></Cell>',
        '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(statusText) + '</Data></Cell>',
        '   </Row>'
      ].join(''));
    });

    executiveTableRowsXml.push([
      '   <Row ss:Height="22">',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">TOTAL (ALL PLANTS)</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">All Locations</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="Number">' + grandTotalQty + '</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="Number">' + grandCritical + '</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="Number">' + grandMajor + '</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="Number">' + grandMinor + '</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">' + (grandTotalQty === 0 ? 'All Plants Leak-Free' : 'Consolidated Total') + '</Data></Cell>',
      '   </Row>'
    ].join(''));
  } else {
    const singleGrp = groups[0] || {};
    executiveTableRowsXml.push([
      '   <Row ss:Height="22">',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Report Scope Parameter</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Value</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Plant Name</Data></Cell>',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(singleGrp.plant || 'PGTL') + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Location</Data></Cell>',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(singleGrp.location || 'Pune') + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Report Month / Year</Data></Cell>',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(monthLabel || '') + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Generated Timestamp</Data></Cell>',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapeXml(generatedTime) + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="12"><Cell ss:StyleID="DataCell"/><Cell ss:StyleID="DataCell"/></Row>',
      '   <Row ss:Height="22">',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Key Performance Indicator (KPI)</Data></Cell>',
      '    <Cell ss:StyleID="SectionHeader"><Data ss:Type="String">Total Count (Qty)</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Total Leakage Incidents</Data></Cell>',
      '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + grandTotalQty + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Critical Severity Leaks</Data></Cell>',
      '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + grandCritical + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Major Severity Leaks</Data></Cell>',
      '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + grandMajor + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Minor Severity Leaks</Data></Cell>',
      '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + grandMinor + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Rework Action Units</Data></Cell>',
      '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + grandRework + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Scrapped Units</Data></Cell>',
      '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + grandScrap + '</Data></Cell>',
      '   </Row>',
      '   <Row ss:Height="20">',
      '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">Accepted / Normal Units</Data></Cell>',
      '    <Cell ss:StyleID="DataCellNumber"><Data ss:Type="Number">' + grandAccepted + '</Data></Cell>',
      '   </Row>'
    ].join(''));
  }

  const xml = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<?mso-application progid="Excel.Sheet"?>',
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"',
    ' xmlns:o="urn:schemas-microsoft-com:office:office"',
    ' xmlns:x="urn:schemas-microsoft-com:office:excel"',
    ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"',
    ' xmlns:html="http://www.w3.org/TR/REC-html40">',
    ' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">',
    '  <Author>PG Group Quality Team</Author>',
    '  <LastAuthor>PG Group AC Leakage Monitoring System</LastAuthor>',
    '  <Created>' + new Date().toISOString() + '</Created>',
    '  <Company>PG GROUP</Company>',
    ' </DocumentProperties>',
    ' <Styles>',
    '  <Style ss:ID="Default" ss:Name="Normal">',
    '   <Alignment ss:Vertical="Center"/>',
    '   <Font ss:FontName="Segoe UI" x:Family="Swiss" ss:Size="10" ss:Color="#1E293B"/>',
    '  </Style>',
    '  <Style ss:ID="HeaderStyle">',
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>',
    '   <Borders>',
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0F172A"/>',
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0F172A"/>',
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0F172A"/>',
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0F172A"/>',
    '   </Borders>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#FFFFFF" ss:Bold="1"/>',
    '   <Interior ss:Color="#1E293B" ss:Pattern="Solid"/>',
    '  </Style>',
    '  <Style ss:ID="TitleStyle">',
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>',
    '   <Font ss:FontName="Segoe UI" ss:Size="13" ss:Color="#FFFFFF" ss:Bold="1"/>',
    '   <Interior ss:Color="#0F172A" ss:Pattern="Solid"/>',
    '  </Style>',
    '  <Style ss:ID="SubtitleStyle">',
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#FFFFFF"/>',
    '   <Interior ss:Color="#334155" ss:Pattern="Solid"/>',
    '  </Style>',
    '  <Style ss:ID="SectionHeader">',
    '   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#0F172A" ss:Bold="1"/>',
    '   <Interior ss:Color="#E2E8F0" ss:Pattern="Solid"/>',
    '  </Style>',
    '  <Style ss:ID="DataCell">',
    '   <Borders>',
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '   </Borders>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#1E293B"/>',
    '  </Style>',
    '  <Style ss:ID="DataCellCenter">',
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>',
    '   <Borders>',
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '   </Borders>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#1E293B"/>',
    '  </Style>',
    '  <Style ss:ID="DataCellNumber">',
    '   <Alignment ss:Horizontal="Right" ss:Vertical="Center"/>',
    '   <Borders>',
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '   </Borders>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#0F172A" ss:Bold="1"/>',
    '   <NumberFormat ss:Format="#,##0"/>',
    '  </Style>',
    '  <Style ss:ID="CriticalBadge">',
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>',
    '   <Borders>',
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '   </Borders>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#991B1B" ss:Bold="1"/>',
    '   <Interior ss:Color="#FEE2E2" ss:Pattern="Solid"/>',
    '  </Style>',
    '  <Style ss:ID="MajorBadge">',
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>',
    '   <Borders>',
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '   </Borders>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#9A3412" ss:Bold="1"/>',
    '   <Interior ss:Color="#FFEDD5" ss:Pattern="Solid"/>',
    '  </Style>',
    '  <Style ss:ID="MinorBadge">',
    '   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>',
    '   <Borders>',
    '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/>',
    '   </Borders>',
    '   <Font ss:FontName="Segoe UI" ss:Size="10" ss:Color="#166534" ss:Bold="1"/>',
    '   <Interior ss:Color="#DCFCE7" ss:Pattern="Solid"/>',
    '  </Style>',
    ' </Styles>',
    ' <Worksheet ss:Name="Defect Records">',
    '  <Table ss:DefaultRowHeight="19">',
    '   <Column ss:Width="135"/>',
    '   <Column ss:Width="95"/>',
    '   <Column ss:Width="85"/>',
    '   <Column ss:Width="105"/>',
    '   <Column ss:Width="115"/>',
    '   <Column ss:Width="135"/>',
    '   <Column ss:Width="115"/>',
    '   <Column ss:Width="85"/>',
    '   <Column ss:Width="65"/>',
    '   <Column ss:Width="105"/>',
    '   <Column ss:Width="65"/>',
    '   <Column ss:Width="115"/>',
    '   <Column ss:Width="115"/>',
    '   <Row ss:Height="24">',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Timestamp</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Location</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Plant</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Unit Type</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Production Line</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Defect Type</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Joint / Location</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Severity</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Shift</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Action Taken</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Quantity</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Operator Name</Data></Cell>',
    '    <Cell ss:StyleID="HeaderStyle"><Data ss:Type="String">Reported By</Data></Cell>',
    '   </Row>',
    defectRowsXml.join('\n'),
    '  </Table>',
    ' </Worksheet>',
    ' <Worksheet ss:Name="Executive Summary">',
    '  <Table ss:DefaultRowHeight="20">',
    '   <Column ss:Width="220"/>',
    '   <Column ss:Width="140"/>',
    '   <Column ss:Width="100"/>',
    '   <Column ss:Width="90"/>',
    '   <Column ss:Width="90"/>',
    '   <Column ss:Width="90"/>',
    '   <Column ss:Width="160"/>',
    '   <Row ss:Height="28">',
    '    <Cell ss:StyleID="TitleStyle" ss:MergeAcross="' + (isMulti ? 6 : 1) + '"><Data ss:Type="String">PG GROUP - AC LEAKAGE MONITORING SYSTEM</Data></Cell>',
    '   </Row>',
    '   <Row ss:Height="22">',
    '    <Cell ss:StyleID="SubtitleStyle" ss:MergeAcross="' + (isMulti ? 6 : 1) + '"><Data ss:Type="String">' + (isMulti ? 'Monthly Multi-Plant Executive Performance Summary' : 'Monthly Performance Executive Summary') + '</Data></Cell>',
    '   </Row>',
    '   <Row ss:Height="12"><Cell ss:StyleID="DataCell"/><Cell ss:StyleID="DataCell"/></Row>',
    executiveTableRowsXml.join('\n'),
    '  </Table>',
    ' </Worksheet>',
    '</Workbook>'
  ].join('\n');

  return Utilities.newBlob(xml, 'application/vnd.ms-excel', fileName);
}

function createMonthlyReportCsvBlob_(plantOrGroups, locationName, monthLabel, records) {
  let groups = [];
  if (Array.isArray(plantOrGroups)) {
    groups = plantOrGroups;
  } else {
    groups = [{
      plant: String(plantOrGroups || 'PGTL').trim(),
      location: String(locationName || 'Pune').trim(),
      records: records || []
    }];
  }

  const isMulti = groups.length > 1;
  const cleanMonth = String(monthLabel || 'Monthly_Report').replace(/[^a-zA-Z0-9_-]/g, '_');
  const cleanTag = isMulti ? 'Consolidated_MultiPlant' : String(groups[0].plant || 'PGTL').replace(/[^a-zA-Z0-9_-]/g, '_');
  const fileName = 'PG_Group_AC_Leakage_Report_' + cleanTag + '_' + cleanMonth + '.csv';

  const headers = [
    'Timestamp',
    'Location',
    'Plant',
    'Unit Type',
    'Production Line',
    'Defect Type',
    'Joint / Location',
    'Severity',
    'Shift',
    'Action Taken',
    'Quantity',
    'Operator Name',
    'Reported By'
  ];

  const escapeCsv = function(val) {
    const s = String(val == null ? '' : val).replace(/"/g, '""');
    return s.search(/("|,|\n|\r)/g) >= 0 ? '"' + s + '"' : s;
  };

  const csvRows = [headers.map(escapeCsv).join(',')];

  groups.forEach(function(grp) {
    const pName = grp.plant || 'PGTL';
    const lName = grp.location || 'Pune';
    const recs = grp.records || [];

    if (recs.length === 0) {
      csvRows.push([
        escapeCsv('N/A'),
        escapeCsv(lName),
        escapeCsv(pName),
        escapeCsv('-'),
        escapeCsv('-'),
        escapeCsv('No leakage entries recorded in AC Leakage Monitoring System for this month'),
        escapeCsv('-'),
        escapeCsv('Normal'),
        escapeCsv('-'),
        escapeCsv('-'),
        escapeCsv(0),
        escapeCsv('-'),
        escapeCsv('System')
      ].join(','));
      return;
    }

    recs.forEach(function(r) {
      let formattedDate = '';
      const dt = parseRecordDate_(r.timestamp);
      if (dt) {
        try {
          formattedDate = Utilities.formatDate(dt, Session.getScriptTimeZone() || 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
        } catch (_) {
          formattedDate = String(r.timestamp);
        }
      } else {
        formattedDate = String(r.timestamp || '');
      }

      csvRows.push([
        escapeCsv(formattedDate),
        escapeCsv(r.location || lName),
        escapeCsv(r.plant || pName),
        escapeCsv(r.type || '-'),
        escapeCsv(r.productionLine || '-'),
        escapeCsv(r.defectType || '-'),
        escapeCsv(r.joint || '-'),
        escapeCsv(r.severity || '-'),
        escapeCsv(r.shift || '-'),
        escapeCsv(r.action || '-'),
        escapeCsv(Number(r.quantity) > 0 ? Number(r.quantity) : 1),
        escapeCsv(r.operatorName || '-'),
        escapeCsv(r.reportedBy || '-')
      ].join(','));
    });
  });

  return Utilities.newBlob(csvRows.join('\r\n'), 'text/csv', fileName);
}

function sendMonthlyPlantLeakageReports(options) {
  options = options || {};
  const cfgRes = getMonthlyReportSettings_();
  const settings = cfgRes.settings || {};

  if (settings.enabled === false && !options.force && !options.testEmail) {
    return { ok: false, message: 'Monthly report auto-email is currently disabled in settings.' };
  }

  const now = new Date();
  let targetMonth = Number(options.month);
  let targetYear = Number(options.year);

  // Default: previous month
  if (!targetMonth || targetMonth < 1 || targetMonth > 12) {
    const prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    targetMonth = prevMonthDate.getMonth() + 1;
    targetYear = targetYear || prevMonthDate.getFullYear();
  } else {
    targetYear = targetYear || now.getFullYear();
  }

  const startDate = new Date(targetYear, targetMonth - 1, 1, 0, 0, 0, 0);
  const endDate = new Date(targetYear, targetMonth, 0, 23, 59, 59, 999);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const monthLabel = monthNames[targetMonth - 1] + ' ' + targetYear;

  // Retrieve reports from sheet
  const rawReports = getDefectReports_({ limit: 10000 }).reports || [];
  
  // Filter STRICTLY for target month & year
  const monthRecords = rawReports.filter(function(r) {
    const dt = parseRecordDate_(r.timestamp);
    if (!dt) return false;
    return dt >= startDate && dt <= endDate;
  });

  const targetLocationFilter = options.location ? String(options.location).trim() : '';
  const targetPlantFilter = options.plant ? String(options.plant).trim() : '';

  // Collect configured plants and locations
  const plantRoutes = settings.plantRoutes || {};
  const locationRoutes = settings.locationRoutes || {};
  const config = getConfig_().config || {};
  const knownLocations = config.locations || {};

  // Build target groups: Map "plant___location" -> { plant, location, records: [] }
  const groupsMap = {};

  // 1. Seed from catalog hierarchy
  Object.keys(knownLocations).forEach(function(locName) {
    const plantsInLoc = (knownLocations[locName] && knownLocations[locName].plants) || {};
    Object.keys(plantsInLoc).forEach(function(pName) {
      const gKey = pName.toLowerCase() + '___' + locName.toLowerCase();
      groupsMap[gKey] = { plant: pName, location: locName, records: [] };
    });
  });

  // 2. Seed from plant master
  const knownPlants = config.plants || {};
  Object.keys(knownPlants).forEach(function(pName) {
    const locs = (knownPlants[pName] && knownPlants[pName].locations) || ['Pune'];
    locs.forEach(function(lName) {
      const gKey = pName.toLowerCase() + '___' + lName.toLowerCase();
      if (!groupsMap[gKey]) {
        groupsMap[gKey] = { plant: pName, location: lName, records: [] };
      }
    });
  });

  // 3. Seed from plantRoutes
  Object.keys(plantRoutes).forEach(function(pName) {
    let loc = 'Pune';
    Object.keys(knownLocations).forEach(function(l) {
      if (knownLocations[l] && knownLocations[l].plants && knownLocations[l].plants[pName]) {
        loc = l;
      }
    });
    const gKey = pName.toLowerCase() + '___' + loc.toLowerCase();
    if (!groupsMap[gKey]) {
      groupsMap[gKey] = { plant: pName, location: loc, records: [] };
    }
  });

  // Seed default if empty
  if (Object.keys(groupsMap).length === 0) {
    groupsMap['pgtl___pune'] = { plant: 'PGTL', location: 'Pune', records: [] };
  }

  // Populate month records into matching groups
  monthRecords.forEach(function(r) {
    const p = String(r.plant || 'PGTL').trim();
    const loc = String(r.location || 'Pune').trim();
    const gKey = p.toLowerCase() + '___' + loc.toLowerCase();
    if (!groupsMap[gKey]) {
      groupsMap[gKey] = { plant: p, location: loc, records: [] };
    }
    groupsMap[gKey].records.push(r);
  });

  // If specific plant and location was passed in options, ensure it exists
  if (targetPlantFilter && targetPlantFilter !== '*' && targetLocationFilter && targetLocationFilter !== '*') {
    const specificKey = targetPlantFilter.toLowerCase() + '___' + targetLocationFilter.toLowerCase();
    if (!groupsMap[specificKey]) {
      groupsMap[specificKey] = { plant: targetPlantFilter, location: targetLocationFilter, records: [] };
    }
  }

  // Group plants by destination recipient so that recipients receiving reports for
  // multiple plants (e.g. global defaultTo, location-level routes, or All Plants test)
  // receive ONE consolidated, clearly formatted email instead of multiple separate emails.
  const recipientBatches = {};
  const emailResults = [];

  Object.keys(groupsMap).forEach(function(gKey) {
    const grp = groupsMap[gKey];
    const plant = grp.plant;
    const location = grp.location;

    // Apply Location filter if specified
    if (targetLocationFilter && targetLocationFilter !== '*' && location.toLowerCase() !== targetLocationFilter.toLowerCase()) {
      return;
    }
    // Apply Plant filter if specified
    if (targetPlantFilter && targetPlantFilter !== '*' && plant.toLowerCase() !== targetPlantFilter.toLowerCase()) {
      return;
    }

    // Hierarchical email resolution:
    // 1. Specific test override
    // 2. Specific Plant route (e.g. PGTL)
    // 3. Location-level route (e.g. Pune)
    // 4. Global default
    let toEmails = options.testEmail ||
      (plantRoutes[plant] && plantRoutes[plant].to) ||
      (locationRoutes[location] && locationRoutes[location].to) ||
      settings.defaultTo ||
      DEFAULT_OTP_SENDER_EMAIL;

    let ccEmails = (options.testEmail ? (options.testCc || '') : '') ||
      (plantRoutes[plant] && plantRoutes[plant].cc) ||
      (locationRoutes[location] && locationRoutes[location].cc) ||
      settings.defaultCc ||
      '';

    const toList = normalizeEmailList_(toEmails);
    const ccList = normalizeEmailList_(ccEmails);

    if (!toList.length) {
      emailResults.push({ plant: plant, location: location, status: 'skipped', reason: 'No recipient email configured' });
      return;
    }

    const batchKey = toList.slice().sort().join(',') + '___' + ccList.slice().sort().join(',');
    if (!recipientBatches[batchKey]) {
      recipientBatches[batchKey] = {
        toList: toList,
        ccList: ccList,
        groups: []
      };
    }
    recipientBatches[batchKey].groups.push(grp);
  });

  // Dispatch emails for each recipient batch
  Object.keys(recipientBatches).forEach(function(batchKey) {
    const batch = recipientBatches[batchKey];
    const toList = batch.toList;
    const ccList = batch.ccList;
    const grps = batch.groups;

    if (!grps.length) return;

    let totalLeaks = 0;
    grps.forEach(function(g) {
      (g.records || []).forEach(function(r) {
        totalLeaks += Number(r.quantity) > 0 ? Number(r.quantity) : 1;
      });
    });

    const isMulti = grps.length > 1;
    let subject = '';
    let plantAuditLabel = '';

    if (isMulti) {
      const plantNames = grps.map(function(g) { return g.plant; });
      subject = 'PG Group AC Leakage Performance Report — Consolidated (' + plantNames.join(' & ') + ') [' + monthLabel + ']';
      plantAuditLabel = 'Consolidated: ' + plantNames.join(', ');
    } else {
      subject = 'PG Group AC Leakage Performance Report — ' + grps[0].plant + ' (' + grps[0].location + ') [' + monthLabel + ']';
      plantAuditLabel = grps[0].plant + ' (' + grps[0].location + ')';
    }

    const htmlBody = buildMonthlyReportHtml_(grps, '', monthLabel);

    // Generate Excel attachment (.xls / .csv)
    let excelAttachment = null;
    try {
      excelAttachment = createMonthlyReportExcelAttachment_(grps, '', monthLabel);
    } catch (attErr) {
      Logger.log('Excel generation fallback: ' + attErr);
      excelAttachment = createMonthlyReportCsvBlob_(grps, '', monthLabel);
    }
    const attachments = excelAttachment ? [excelAttachment] : [];

    try {
      sendSystemEmail_(toList, ccList, subject, htmlBody, '', attachments);
      emailResults.push({
        plants: grps.map(function(g) { return g.plant + ' (' + g.location + ')'; }).join(', '),
        to: toList.join(', '),
        cc: ccList.join(', '),
        totalLeaks: totalLeaks,
        status: 'sent',
        attachment: excelAttachment ? excelAttachment.getName() : 'none'
      });
      logEmailActivity_({
        type: options.testEmail ? 'test_report' : 'monthly_report',
        plant: plantAuditLabel,
        month: monthLabel,
        to: toList.join(', '),
        cc: ccList.join(', '),
        subject: subject,
        totalLeaks: totalLeaks,
        status: 'SUCCESS',
        error: '',
        actor: options.actorEmail || 'system'
      });
    } catch (sendErr) {
      const errMsg = sendErr.message || String(sendErr);
      emailResults.push({
        plants: grps.map(function(g) { return g.plant + ' (' + g.location + ')'; }).join(', '),
        to: toList.join(', '),
        cc: ccList.join(', '),
        status: 'error',
        error: errMsg
      });
      logEmailActivity_({
        type: options.testEmail ? 'test_report' : 'monthly_report',
        plant: plantAuditLabel,
        month: monthLabel,
        to: toList.join(', '),
        cc: ccList.join(', '),
        subject: subject,
        totalLeaks: totalLeaks,
        status: 'ERROR',
        error: errMsg,
        actor: options.actorEmail || 'system'
      });
    }
  });

  return {
    ok: true,
    month: monthLabel,
    sentCount: emailResults.filter(function(x) { return x.status === 'sent'; }).length,
    results: emailResults
  };
}

function previewMonthlyReport_(payload) {
  payload = payload || {};
  const now = new Date();
  let targetMonth = Number(payload.month);
  let targetYear = Number(payload.year);

  if (!targetMonth || targetMonth < 1 || targetMonth > 12) {
    const prevMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    targetMonth = prevMonthDate.getMonth() + 1;
    targetYear = targetYear || prevMonthDate.getFullYear();
  } else {
    targetYear = targetYear || now.getFullYear();
  }

  const startDate = new Date(targetYear, targetMonth - 1, 1, 0, 0, 0, 0);
  const endDate = new Date(targetYear, targetMonth, 0, 23, 59, 59, 999);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const monthLabel = monthNames[targetMonth - 1] + ' ' + targetYear;

  const rawReports = getDefectReports_({ limit: 10000 }).reports || [];
  const locationName = String(payload.location || '').trim();
  const plantName = String(payload.plant || '*').trim();

  // Filter STRICTLY for target month & year
  const monthRecords = rawReports.filter(function(r) {
    const dt = parseRecordDate_(r.timestamp);
    if (!dt) return false;
    return dt >= startDate && dt <= endDate;
  });

  const config = getConfig_().config || {};
  const knownLocations = config.locations || {};
  const knownPlants = config.plants || {};
  const groupsMap = {};

  // Seed known plants and locations
  Object.keys(knownLocations).forEach(function(loc) {
    const plantsInLoc = (knownLocations[loc] && knownLocations[loc].plants) || {};
    Object.keys(plantsInLoc).forEach(function(p) {
      const gKey = p.toLowerCase() + '___' + loc.toLowerCase();
      groupsMap[gKey] = { plant: p, location: loc, records: [] };
    });
  });

  Object.keys(knownPlants).forEach(function(p) {
    const locs = (knownPlants[p] && knownPlants[p].locations) || ['Pune'];
    locs.forEach(function(loc) {
      const gKey = p.toLowerCase() + '___' + loc.toLowerCase();
      if (!groupsMap[gKey]) {
        groupsMap[gKey] = { plant: p, location: loc, records: [] };
      }
    });
  });

  if (Object.keys(groupsMap).length === 0) {
    groupsMap['pgtl___pune'] = { plant: 'PGTL', location: 'Pune', records: [] };
  }

  // Populate month records
  monthRecords.forEach(function(r) {
    const p = String(r.plant || 'PGTL').trim();
    const loc = String(r.location || 'Pune').trim();
    const gKey = p.toLowerCase() + '___' + loc.toLowerCase();
    if (!groupsMap[gKey]) {
      groupsMap[gKey] = { plant: p, location: loc, records: [] };
    }
    groupsMap[gKey].records.push(r);
  });

  // Filter groups according to payload
  const matchingGroups = [];
  Object.keys(groupsMap).forEach(function(gKey) {
    const grp = groupsMap[gKey];
    if (locationName && locationName !== '*' && grp.location.toLowerCase() !== locationName.toLowerCase()) return;
    if (plantName && plantName !== '*' && grp.plant.toLowerCase() !== plantName.toLowerCase()) return;
    matchingGroups.push(grp);
  });

  if (matchingGroups.length === 0) {
    matchingGroups.push({
      plant: plantName && plantName !== '*' ? plantName : 'PGTL',
      location: locationName && locationName !== '*' ? locationName : 'Pune',
      records: []
    });
  }

  let grandTotalLeaks = 0;
  matchingGroups.forEach(function(g) {
    (g.records || []).forEach(function(r) {
      grandTotalLeaks += Number(r.quantity) > 0 ? Number(r.quantity) : 1;
    });
  });

  const html = buildMonthlyReportHtml_(matchingGroups, '', monthLabel);

  return {
    ok: true,
    html: html,
    plant: plantName === '*' ? 'All Plants' : plantName,
    location: locationName === '*' ? 'All Locations' : (locationName || 'Pune'),
    month: monthLabel,
    totalLeaks: grandTotalLeaks,
    recordsCount: matchingGroups.reduce(function(acc, g) { return acc + (g.records || []).length; }, 0)
  };
}

function logEmailActivity_(entry) {
  const headers = ['timestamp', 'type', 'plant', 'month', 'to', 'cc', 'subject', 'totalLeaks', 'status', 'error', 'actor'];
  appendRow_(TAB.emailLogs, headers, [
    new Date(),
    String(entry.type || ''),
    String(entry.plant || ''),
    String(entry.month || ''),
    String(entry.to || ''),
    String(entry.cc || ''),
    String(entry.subject || ''),
    Number(entry.totalLeaks || 0),
    String(entry.status || 'SUCCESS'),
    String(entry.error || ''),
    String(entry.actor || '')
  ]);
}

function getEmailLogs_(payload) {
  const limit = Math.max(1, Math.min(200, Number((payload && payload.limit) || 50)));
  const sh = ensureTab_(TAB.emailLogs, ['timestamp', 'type', 'plant', 'month', 'to', 'cc', 'subject', 'totalLeaks', 'status', 'error', 'actor']);
  const values = sh.getDataRange().getValues();
  const logs = [];

  for (let i = values.length - 1; i >= 1 && logs.length < limit; i--) {
    logs.push({
      timestamp: values[i][0] ? new Date(values[i][0]).toISOString() : '',
      type: String(values[i][1] || ''),
      plant: String(values[i][2] || ''),
      month: String(values[i][3] || ''),
      to: String(values[i][4] || ''),
      cc: String(values[i][5] || ''),
      subject: String(values[i][6] || ''),
      totalLeaks: Number(values[i][7] || 0),
      status: String(values[i][8] || 'SUCCESS'),
      error: String(values[i][9] || ''),
      actor: String(values[i][10] || '')
    });
  }
  return { logs: logs };
}

function setupMonthlyReportTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trig) {
    if (trig.getHandlerFunction() === 'sendMonthlyPlantLeakageReports') {
      ScriptApp.deleteTrigger(trig);
    }
  });

  ScriptApp.newTrigger('sendMonthlyPlantLeakageReports')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  return { ok: true, message: 'Automated Monthly Report trigger scheduled for 1st of every month at 8:00 AM.' };
}

// Manual runner for Monthly Reports
function sendMonthlyReportsManually() {
  return sendMonthlyPlantLeakageReports({ force: true });
}