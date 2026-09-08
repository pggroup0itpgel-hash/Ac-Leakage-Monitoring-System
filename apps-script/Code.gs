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

function buildMonthlyReportHtml_(plantName, locationName, monthLabel, records, totalLeaks) {
  locationName = String(locationName || '').trim();
  plantName = String(plantName || '').trim() || 'All Plants';

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

  let defectRowsHtml = '';
  if (sortedDefects.length === 0) {
    defectRowsHtml = '<tr><td colspan="3" style="padding:12px;text-align:center;color:#64748b;">No defect data recorded.</td></tr>';
  } else {
    sortedDefects.forEach(function(d, idx) {
      const barColor = idx === 0 ? '#4f46e5' : idx === 1 ? '#06b6d4' : '#64748b';
      defectRowsHtml += [
        '<tr style="border-bottom:1px solid #f1f5f9;">',
        '  <td style="padding:10px 12px;color:#1e293b;font-weight:600;font-size:13px;">' + d.name + '</td>',
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
  }

  let unitRowsHtml = '';
  if (sortedUnits.length === 0) {
    unitRowsHtml = '<tr><td colspan="3" style="padding:12px;text-align:center;color:#64748b;">No unit data recorded.</td></tr>';
  } else {
    sortedUnits.forEach(function(u, idx) {
      const barColor = idx === 0 ? '#10b981' : idx === 1 ? '#3b82f6' : '#8b5cf6';
      unitRowsHtml += [
        '<tr style="border-bottom:1px solid #f1f5f9;">',
        '  <td style="padding:9px 12px;color:#1e293b;font-weight:600;font-size:13px;">' + u.name + '</td>',
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
  }

  let lineRowsHtml = '';
  if (sortedLines.length === 0) {
    lineRowsHtml = '<tr><td colspan="2" style="padding:12px;text-align:center;color:#64748b;">No line records.</td></tr>';
  } else {
    sortedLines.forEach(function(l) {
      lineRowsHtml += [
        '<tr style="border-bottom:1px solid #f1f5f9;">',
        '  <td style="padding:9px 12px;color:#334155;font-weight:500;font-size:13px;">' + l.name + '</td>',
        '  <td style="padding:9px 12px;text-align:right;font-weight:700;color:#0f172a;font-size:13px;">' + l.count + '</td>',
        '</tr>'
      ].join('');
    });
  }

  let shiftRowsHtml = '';
  if (sortedShifts.length === 0) {
    shiftRowsHtml = '<tr><td colspan="2" style="padding:12px;text-align:center;color:#64748b;">No shift records.</td></tr>';
  } else {
    sortedShifts.forEach(function(s) {
      shiftRowsHtml += [
        '<tr style="border-bottom:1px solid #f1f5f9;">',
        '  <td style="padding:9px 12px;color:#334155;font-weight:500;font-size:13px;">' + s.name + '</td>',
        '  <td style="padding:9px 12px;text-align:right;font-weight:700;color:#0f172a;font-size:13px;">' + s.count + '</td>',
        '</tr>'
      ].join('');
    });
  }

  let jointRowsHtml = '';
  if (sortedJoints.length > 0) {
    sortedJoints.forEach(function(j) {
      jointRowsHtml += [
        '<tr style="border-bottom:1px solid #f1f5f9;">',
        '  <td style="padding:8px 12px;color:#334155;font-weight:600;font-size:12px;">' + j.name + '</td>',
        '  <td style="padding:8px 12px;text-align:right;font-weight:700;color:#dc2626;font-size:12px;">' + j.count + '</td>',
        '</tr>'
      ].join('');
    });
  }

   const generatedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a');
  const locBadge = '<span style="display:inline-block;background-color:#334155;background:#334155;border:1px solid #475569;border-radius:6px;padding:5px 12px;font-size:12px;font-weight:700;color:#ffffff;margin-right:8px;margin-bottom:6px;"><b style="color:#93c5fd;">LOCATION:</b> ' + (locationName || 'Corporate') + '</span>';

  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>',
    '<body style="margin:0;padding:0;background-color:#f8fafc;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif;">',
    '  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background-color:#f8fafc;padding:24px 0;">',
    '    <tr>',
    '      <td align="center">',
    '        <table role="presentation" width="100%" style="max-width:680px;background-color:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 8px 30px rgba(0,0,0,0.06);border:1px solid #e2e8f0;" cellspacing="0" cellpadding="0">',
    '          ',
    '          <!-- Header Banner (Solid background for full Gmail & Outlook compatibility) -->',
    '          <tr>',
    '            <td bgcolor="#1e293b" style="background-color:#1e293b !important;background:#1e293b;padding:28px 24px;text-align:left;color:#ffffff;border-bottom:3px solid #3b82f6;">',
    '              <div style="font-size:12px;font-weight:900;letter-spacing:2px;text-transform:uppercase;color:#93c5fd;margin-bottom:6px;">PG GROUP</div>',
    '              <h1 style="margin:0 0 14px 0;font-size:22px;font-weight:800;color:#ffffff;line-height:1.3;">Monthly AC Leakage Quality Report</h1>',
    '              <div style="display:block;margin-top:10px;">',
    '                ' + locBadge,
    '                <span style="display:inline-block;background-color:#334155;background:#334155;border:1px solid #475569;border-radius:6px;padding:5px 12px;font-size:12px;font-weight:700;color:#ffffff;margin-right:8px;margin-bottom:6px;"><b style="color:#93c5fd;">PLANT:</b> ' + plantName + '</span>',
    '                <span style="display:inline-block;background-color:#334155;background:#334155;border:1px solid #475569;border-radius:6px;padding:5px 12px;font-size:12px;font-weight:700;color:#ffffff;margin-bottom:6px;"><b style="color:#93c5fd;">PERIOD:</b> ' + monthLabel + '</span>',
    '              </div>',
    '            </td>',
    '          </tr>',
    '          ',
    '          <!-- Main Content Container -->',
    '          <tr>',
    '            <td style="padding:24px;">',
    '              ',
    '              <!-- Executive Summary Callout -->',
    '              <div style="background:#f8fafc;border-left:4px solid #3b82f6;padding:14px 18px;border-radius:0 8px 8px 0;margin-bottom:22px;">',
    '                <p style="margin:0;font-size:13px;color:#334155;line-height:1.6;">',
    '                  Performance overview for <b>' + plantName + (locationName ? ' (' + locationName + ')' : '') + '</b> during <b>' + monthLabel + '</b>. Total of <b>' + totalLeaks + '</b> leakage defect(s) logged across all production lines and shifts.',
    '                </p>',
    '              </div>',
    '              ',
    '              <!-- KPI Metric Cards Grid (6 cards) -->',
    '              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:22px;">',
    '                <tr>',
    '                  <td width="33.33%" style="padding:4px;">',
    '                    <div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:12px 8px;text-align:center;">',
    '                      <div style="font-size:11px;font-weight:700;color:#1d4ed8;text-transform:uppercase;">TOTAL LEAKS</div>',
    '                      <div style="font-size:24px;font-weight:900;color:#1e3a8a;margin-top:2px;">' + totalLeaks + '</div>',
    '                    </div>',
    '                  </td>',
    '                  <td width="33.33%" style="padding:4px;">',
    '                    <div style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:12px 8px;text-align:center;">',
    '                      <div style="font-size:11px;font-weight:700;color:#b91c1c;text-transform:uppercase;">CRITICAL</div>',
    '                      <div style="font-size:24px;font-weight:900;color:#991b1b;margin-top:2px;">' + criticalCount + '</div>',
    '                    </div>',
    '                  </td>',
    '                  <td width="33.33%" style="padding:4px;">',
    '                    <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;padding:12px 8px;text-align:center;">',
    '                      <div style="font-size:11px;font-weight:700;color:#c2410c;text-transform:uppercase;">MAJOR</div>',
    '                      <div style="font-size:24px;font-weight:900;color:#9a3412;margin-top:2px;">' + majorCount + '</div>',
    '                    </div>',
    '                  </td>',
    '                </tr>',
    '                <tr>',
    '                  <td width="33.33%" style="padding:4px;">',
    '                    <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:12px 8px;text-align:center;">',
    '                      <div style="font-size:11px;font-weight:700;color:#15803d;text-transform:uppercase;">MINOR</div>',
    '                      <div style="font-size:24px;font-weight:900;color:#166534;margin-top:2px;">' + minorCount + '</div>',
    '                    </div>',
    '                  </td>',
    '                  <td width="33.33%" style="padding:4px;">',
    '                    <div style="background:#fefce8;border:1px solid #fef08a;border-radius:10px;padding:12px 8px;text-align:center;">',
    '                      <div style="font-size:11px;font-weight:700;color:#a16207;text-transform:uppercase;">REWORK</div>',
    '                      <div style="font-size:24px;font-weight:900;color:#854d0e;margin-top:2px;">' + reworkCount + '</div>',
    '                    </div>',
    '                  </td>',
    '                  <td width="33.33%" style="padding:4px;">',
    '                    <div style="background:#fdf2f8;border:1px solid #fbcfe8;border-radius:10px;padding:12px 8px;text-align:center;">',
    '                      <div style="font-size:11px;font-weight:700;color:#be185d;text-transform:uppercase;">SCRAP</div>',
    '                      <div style="font-size:24px;font-weight:900;color:#9d174d;margin-top:2px;">' + scrapCount + '</div>',
    '                    </div>',
    '                  </td>',
    '                </tr>',
    '              </table>',
    '              ',
    '              <!-- Unit Type & Defect Breakdown (Two Columns) -->',
    '              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:22px;">',
    '                <tr>',
    '                  <td width="50%" valign="top" style="padding-right:8px;">',
    '                    <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Unit Type Breakdown</div>',
    '                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '                      <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '                        <tr>',
    '                          <th style="padding:8px 10px;text-align:left;">Unit Type</th>',
    '                          <th style="padding:8px 6px;text-align:center;">Count</th>',
    '                          <th style="padding:8px 10px;text-align:left;">Share</th>',
    '                        </tr>',
    '                      </thead>',
    '                      <tbody>' + unitRowsHtml + '</tbody>',
    '                    </table>',
    '                  </td>',
    '                  <td width="50%" valign="top" style="padding-left:8px;">',
    '                    <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Defect Categories</div>',
    '                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '                      <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '                        <tr>',
    '                          <th style="padding:8px 10px;text-align:left;">Category</th>',
    '                          <th style="padding:8px 6px;text-align:center;">Count</th>',
    '                          <th style="padding:8px 10px;text-align:left;">Share</th>',
    '                        </tr>',
    '                      </thead>',
    '                      <tbody>' + defectRowsHtml + '</tbody>',
    '                    </table>',
    '                  </td>',
    '                </tr>',
    '              </table>',
    '              ',
    '              <!-- Line Breakdown & Shift Tables (Two Columns) -->',
    '              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin-bottom:20px;">',
    '                <tr>',
    '                  <td width="50%" valign="top" style="padding-right:8px;">',
    '                    <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Line Breakdown</div>',
    '                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '                      <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '                        <tr>',
    '                          <th style="padding:8px 12px;text-align:left;">Line</th>',
    '                          <th style="padding:8px 12px;text-align:right;">Leaks</th>',
    '                        </tr>',
    '                      </thead>',
    '                      <tbody>' + lineRowsHtml + '</tbody>',
    '                    </table>',
    '                  </td>',
    '                  <td width="50%" valign="top" style="padding-left:8px;">',
    '                    <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Shift Distribution</div>',
    '                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '                      <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '                        <tr>',
    '                          <th style="padding:8px 12px;text-align:left;">Shift</th>',
    '                          <th style="padding:8px 12px;text-align:right;">Leaks</th>',
    '                        </tr>',
    '                      </thead>',
    '                      <tbody>' + shiftRowsHtml + '</tbody>',
    '                    </table>',
    '                  </td>',
    '                </tr>',
    '              </table>',
    '              ',
    '              ' + (jointRowsHtml ? [
    '              <!-- Top Joint Leakage Points -->',
    '              <div style="margin-bottom:20px;">',
    '                <div style="font-size:13px;font-weight:800;color:#0f172a;margin-bottom:8px;border-bottom:2px solid #e2e8f0;padding-bottom:5px;">Top Leakage Joints</div>',
    '                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">',
    '                  <thead style="background:#f8fafc;color:#475569;font-size:11px;text-transform:uppercase;font-weight:700;">',
    '                    <tr>',
    '                      <th style="padding:8px 12px;text-align:left;">Joint / Location</th>',
    '                      <th style="padding:8px 12px;text-align:right;">Leak Frequency</th>',
    '                    </tr>',
    '                  </thead>',
    '                  <tbody>' + jointRowsHtml + '</tbody>',
    '                </table>',
    '              </div>'
    ].join('') : '') + '',
    '              ',
    '              <!-- PROMINENT ATTACHMENT CALLOUT BANNER -->',
    '              <div style="margin-top:20px;padding:12px 18px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;color:#166534;font-size:12px;line-height:1.5;">',
    '                <div style="font-weight:800;font-size:12px;margin-bottom:3px;color:#15803d;letter-spacing:0.3px;">📊 EXCEL DETAILED REPORT ATTACHED (.xlsx)</div>',
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
  ].join('');
}

function createMonthlyReportExcelAttachment_(plantName, locationName, monthLabel, records) {
  const cleanPlant = (plantName || 'PGTL').replace(/[^a-zA-Z0-9_-]/g, '_');
  const cleanMonth = (monthLabel || 'Monthly_Report').replace(/[^a-zA-Z0-9_-]/g, '_');
  const fileName = 'PG_Group_AC_Leakage_Report_' + cleanPlant + '_' + cleanMonth + '.xlsx';

  let tempSs = null;
  try {
    tempSs = SpreadsheetApp.create('Temp_Monthly_Leakage_Export_' + cleanPlant + '_' + cleanMonth);
    const ssId = tempSs.getId();

    // Sheet 1: Detailed Defect Records
    const dataSheet = tempSs.getSheets()[0];
    dataSheet.setName('Defect Records');

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

    const rows = [headers];
    let totalQty = 0;
    let criticalCount = 0;
    let majorCount = 0;
    let minorCount = 0;
    let reworkCount = 0;
    let scrapCount = 0;
    let acceptedCount = 0;

    records.forEach(function(r) {
      const q = Number(r.quantity) > 0 ? Number(r.quantity) : 1;
      totalQty += q;
      const sev = String(r.severity || '').toLowerCase();
      if (sev === 'critical') criticalCount += q;
      else if (sev === 'major') majorCount += q;
      else minorCount += q;

      const act = String(r.action || '').toLowerCase();
      if (act.indexOf('rework') >= 0) reworkCount += q;
      else if (act.indexOf('scrap') >= 0) scrapCount += q;
      else if (act.indexOf('accept') >= 0) acceptedCount += q;

      let formattedDate = '';
      if (r.timestamp) {
        try {
          formattedDate = Utilities.formatDate(new Date(r.timestamp), Session.getScriptTimeZone() || 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
        } catch (_) {
          formattedDate = String(r.timestamp);
        }
      }

      rows.push([
        formattedDate,
        String(r.location || locationName || 'Pune'),
        String(r.plant || plantName || 'PGTL'),
        String(r.type || '-'),
        String(r.productionLine || '-'),
        String(r.defectType || '-'),
        String(r.joint || '-'),
        String(r.severity || '-'),
        String(r.shift || '-'),
        String(r.action || '-'),
        q,
        String(r.operatorName || '-'),
        String(r.reportedBy || '-')
      ]);
    });

    if (rows.length === 1) {
      rows.push(['No defect records found for this period.', '', '', '', '', '', '', '', '', '', 0, '', '']);
    }

    const range = dataSheet.getRange(1, 1, rows.length, headers.length);
    range.setValues(rows);

    // Styling Header
    const headerRange = dataSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1e293b');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    // Freeze header row
    dataSheet.setFrozenRows(1);

    // Sheet 2: Executive Summary & KPIs
    const summarySheet = tempSs.insertSheet('Executive Summary');
    const summaryData = [
      ['PG GROUP - AC LEAKAGE MONITORING SYSTEM', ''],
      ['Monthly Performance Executive Summary', ''],
      ['', ''],
      ['Report Parameter', 'Value'],
      ['Plant Name', String(plantName || 'All Plants')],
      ['Location', String(locationName || 'Pune')],
      ['Month / Year', String(monthLabel || '')],
      ['Generated On', Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss')],
      ['', ''],
      ['Key Performance Indicator (KPI)', 'Count'],
      ['Total Leakage Incidents (Qty)', totalQty],
      ['Critical Severity Leaks', criticalCount],
      ['Major Severity Leaks', majorCount],
      ['Minor Severity Leaks', minorCount],
      ['Rework Action Items', reworkCount],
      ['Scrapped Units', scrapCount],
      ['Accepted / Normal', acceptedCount]
    ];

    summarySheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
    summarySheet.getRange('A1:B1').merge().setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
    summarySheet.getRange('A2:B2').merge().setBackground('#334155').setFontColor('#ffffff').setFontSize(11).setHorizontalAlignment('center');
    summarySheet.getRange('A4:B4').setBackground('#e2e8f0').setFontWeight('bold');
    summarySheet.getRange('A10:B10').setBackground('#e2e8f0').setFontWeight('bold');
    summarySheet.setColumnWidth(1, 260);
    summarySheet.setColumnWidth(2, 160);

    SpreadsheetApp.flush();

    // Export to Excel .xlsx via OAuth token
    const url = 'https://docs.google.com/feeds/download/spreadsheets/Export?key=' + ssId + '&exportFormat=xlsx';
    const params = {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    const res = UrlFetchApp.fetch(url, params);

    if (res.getResponseCode() === 200) {
      const blob = res.getBlob().setName(fileName);
      try {
        DriveApp.getFileById(ssId).setTrashed(true);
      } catch (_) {}
      return blob;
    }
  } catch (err) {
    Logger.log('createMonthlyReportExcelAttachment_ error: ' + err);
  } finally {
    if (tempSs) {
      try {
        DriveApp.getFileById(tempSs.getId()).setTrashed(true);
      } catch (_) {}
    }
  }

  // Fallback to CSV format if XLSX export cannot be generated
  return createMonthlyReportCsvBlob_(plantName, locationName, monthLabel, records);
}

function createMonthlyReportCsvBlob_(plantName, locationName, monthLabel, records) {
  const cleanPlant = (plantName || 'PGTL').replace(/[^a-zA-Z0-9_-]/g, '_');
  const cleanMonth = (monthLabel || 'Monthly_Report').replace(/[^a-zA-Z0-9_-]/g, '_');
  const fileName = 'PG_Group_AC_Leakage_Report_' + cleanPlant + '_' + cleanMonth + '.csv';

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

  records.forEach(function(r) {
    let formattedDate = '';
    if (r.timestamp) {
      try {
        formattedDate = Utilities.formatDate(new Date(r.timestamp), Session.getScriptTimeZone() || 'GMT+5:30', 'yyyy-MM-dd HH:mm:ss');
      } catch (_) {
        formattedDate = String(r.timestamp);
      }
    }
    csvRows.push([
      escapeCsv(formattedDate),
      escapeCsv(r.location || locationName || 'Pune'),
      escapeCsv(r.plant || plantName || 'PGTL'),
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

  const startDate = new Date(targetYear, targetMonth - 1, 1, 0, 0, 0);
  const endDate = new Date(targetYear, targetMonth, 0, 23, 59, 59, 999);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const monthLabel = monthNames[targetMonth - 1] + ' ' + targetYear;

  // Retrieve reports from sheet
  const rawReports = getDefectReports_({ limit: 5000 }).reports || [];
  const monthRecords = rawReports.filter(function(r) {
    if (!r.timestamp) return false;
    const dt = new Date(r.timestamp);
    return dt >= startDate && dt <= endDate;
  });

  // Group by plant and capture location
  const plantGroups = {};
  const plantLocations = {};

  monthRecords.forEach(function(r) {
    const p = String(r.plant || 'PGTL').trim();
    const loc = String(r.location || 'Pune').trim();
    if (!plantGroups[p]) plantGroups[p] = [];
    plantGroups[p].push(r);
    if (!plantLocations[p]) plantLocations[p] = loc;
  });

  // Ensure configured plants appear even if 0 leaks
  const plantRoutes = settings.plantRoutes || {};
  const locationRoutes = settings.locationRoutes || {};

  Object.keys(plantRoutes).forEach(function(p) {
    if (!plantGroups[p]) plantGroups[p] = [];
    if (!plantLocations[p]) plantLocations[p] = 'Pune';
  });
  if (Object.keys(plantGroups).length === 0) {
    plantGroups['PGTL'] = [];
    plantLocations['PGTL'] = 'Pune';
  }

  const targetLocationFilter = options.location ? String(options.location).trim() : '';
  const targetPlantFilter = options.plant ? String(options.plant).trim() : '';
  const emailResults = [];

  Object.keys(plantGroups).forEach(function(plant) {
    const location = plantLocations[plant] || 'Pune';

    if (targetLocationFilter && targetLocationFilter !== '*' && targetLocationFilter.toLowerCase() !== location.toLowerCase()) {
      return;
    }
    if (targetPlantFilter && targetPlantFilter !== '*' && targetPlantFilter.toLowerCase() !== plant.toLowerCase()) {
      return;
    }

    const recs = plantGroups[plant] || [];

    // Hierarchical email resolution:
    // 1. Specific test override
    // 2. Specific Plant route (e.g. PGTL)
    // 3. Location-level route (e.g. Pune - maps all plants in Pune)
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

    let totalLeaks = 0;
    recs.forEach(function(r) {
      const q = Number(r.quantity) > 0 ? Number(r.quantity) : 1;
      totalLeaks += q;
    });

    const htmlBody = buildMonthlyReportHtml_(plant, location, monthLabel, recs, totalLeaks);
    const subject = 'PG Group AC Leakage Performance Report — ' + plant + ' (' + location + ') [' + monthLabel + ']';

    // Generate Excel attachment for this plant report
    let excelAttachment = null;
    try {
      excelAttachment = createMonthlyReportExcelAttachment_(plant, location, monthLabel, recs);
    } catch (attErr) {
      Logger.log('Excel generation fallback: ' + attErr);
      excelAttachment = createMonthlyReportCsvBlob_(plant, location, monthLabel, recs);
    }
    const attachments = excelAttachment ? [excelAttachment] : [];

    try {
      sendSystemEmail_(toList, ccList, subject, htmlBody, '', attachments);
      emailResults.push({
        plant: plant,
        location: location,
        to: toList.join(', '),
        cc: ccList.join(', '),
        totalLeaks: totalLeaks,
        status: 'sent',
        attachment: excelAttachment ? excelAttachment.getName() : 'none'
      });
      logEmailActivity_({
        type: options.testEmail ? 'test_report' : 'monthly_report',
        plant: plant + ' (' + location + ')',
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
        plant: plant,
        location: location,
        to: toList.join(', '),
        cc: ccList.join(', '),
        status: 'error',
        error: errMsg
      });
      logEmailActivity_({
        type: options.testEmail ? 'test_report' : 'monthly_report',
        plant: plant + ' (' + location + ')',
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

  const startDate = new Date(targetYear, targetMonth - 1, 1, 0, 0, 0);
  const endDate = new Date(targetYear, targetMonth, 0, 23, 59, 59, 999);
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const monthLabel = monthNames[targetMonth - 1] + ' ' + targetYear;

  const rawReports = getDefectReports_({ limit: 5000 }).reports || [];
  const locationName = String(payload.location || '').trim();
  const plantName = String(payload.plant || 'PGTL').trim() || 'PGTL';

  const plantRecords = rawReports.filter(function(r) {
    if (!r.timestamp) return false;
    const dt = new Date(r.timestamp);
    const inDate = dt >= startDate && dt <= endDate;
    if (!inDate) return false;
    if (locationName && locationName !== '*' && String(r.location || '').toLowerCase() !== locationName.toLowerCase()) return false;
    if (plantName !== '*' && plantName && String(r.plant || '').toLowerCase() !== plantName.toLowerCase()) return false;
    return true;
  });

  let totalLeaks = 0;
  plantRecords.forEach(function(r) {
    const q = Number(r.quantity) > 0 ? Number(r.quantity) : 1;
    totalLeaks += q;
  });

  const html = buildMonthlyReportHtml_(
    plantName === '*' ? 'All Plants' : plantName,
    locationName === '*' ? 'All Locations' : (locationName || 'Pune'),
    monthLabel,
    plantRecords,
    totalLeaks
  );

  return {
    ok: true,
    html: html,
    plant: plantName,
    location: locationName || 'Pune',
    month: monthLabel,
    totalLeaks: totalLeaks,
    recordsCount: plantRecords.length
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