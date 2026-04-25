/**
 * AC Leakage - Apps Script OTP + Sheet Sync API
 * Deploy as Web App:
 * - Execute as: Me
 * - Who has access: Anyone
 */

const SHEET_ID = '1yecFSz_FrKQ0UFsavuwZjL_-AwvHTJbyx6_NMT9fN6M';
const OTP_TTL_MS = 10 * 60 * 1000;
const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000;
// No hardcoded admin emails. Authorization is controlled by Google Sheet tabs `admins` and `employees`.

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
  setupSheet: 'setup_sheet'
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
    else if (action === 'getDefectReports') result = getDefectReports_(payload);
    else if (action === 'applySetupSheet') result = applySetupSheet_(payload);
    else if (action === 'getStatus') result = getStatus_(payload);
    else throw new Error('Unsupported action');

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
  return {
    serverTime: new Date().toISOString(),
    sheetConfigured: configured,
    requiredTabs: tabs
  };
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

function getUsersFromSheetsByRole_() {
  const admins = ensureTab_(TAB.admins, ['email', 'role', 'allowedPlants', 'allowedLocations', 'canSubmitDefects', 'updatedAt']);
  const employees = ensureTab_(TAB.employees, ['email', 'role', 'allowedPlants', 'allowedLocations', 'canSubmitDefects', 'updatedAt']);

  const adminUsers = {};
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
  MailApp.sendEmail({
    to: email,
    subject: 'AC Leakage Monitoring — Login Verification Code',
    htmlBody: html,
    name: 'AC Leakage Monitoring'
  });

  logActivity_({ email: email, role: profile.role, activity: 'otp_sent', source: source || 'code.gs' });
  return { message: 'OTP sent' };
}

function verifyOtp_(payload) {
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
    JSON.stringify(report)
  ];
  const headers = ['timestamp', 'plant', 'location', 'line', 'type', 'defectType', 'joint', 'severity', 'shift', 'action', 'operatorName', 'reportedBy', 'rawJson'];
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
  const sh = ensureTab_(TAB.reports, ['timestamp', 'plant', 'location', 'line', 'type', 'defectType', 'joint', 'severity', 'shift', 'action', 'operatorName', 'reportedBy', 'rawJson']);
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = values.length - 1; i >= 1 && out.length < limit; i--) {
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
      reportedBy: String(values[i][11] || '')
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
  const reportHeaders = ['timestamp', 'plant', 'location', 'line', 'type', 'defectType', 'joint', 'severity', 'shift', 'action', 'operatorName', 'reportedBy', 'rawJson'];
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
