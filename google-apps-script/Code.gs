/**
 * Student Movement Tracker - Google Apps Script Backend API (Code.gs)
 *
 * Responsibilities:
 * - Header-mapped Google Sheets data layer across all 8 required sheets
 * - Role-based authentication & authorization (teacher, leading_teacher, principal, admin)
 * - Dedicated teacher token signing secret bound to TeacherID & TeacherName
 * - Salted iterated SHA-256 credential hashing & constant-time verification
 * - Atomic movement recording (OUT & RETURN) using LockService
 * - Authoritative timestamp preservation and 1899 date-corruption protection
 * - Comprehensive CRUD, bulk student & teacher import validation, and movement analytics
 */

var SMT_SHEETS = {
  STUDENTS: {
    name: 'Students',
    headers: ['StudentID', 'StudentName', 'ClassID', 'Status', 'CreatedAt', 'UpdatedAt']
  },
  CLASSES: {
    name: 'Classes',
    headers: ['ClassID', 'ClassName', 'Status', 'CreatedAt', 'UpdatedAt']
  },
  MOVEMENT_LOG: {
    name: 'Movement_Log',
    headers: [
      'MovementID',
      'StudentID',
      'StudentName',
      'ClassID',
      'ClassName',
      'TeacherID',
      'TeacherName',
      'Destination',
      'TimeOut',
      'TimeIn',
      'Status',
      'Notes'
    ]
  },
  SETTINGS: {
    name: 'Settings',
    headers: ['SettingKey', 'SettingValue']
  },
  TEACHERS: {
    name: 'Teachers',
    headers: ['TeacherID', 'TeacherName', 'Status', 'CreatedAt', 'UpdatedAt']
  },
  TEACHING_ASSIGNMENTS: {
    name: 'Teaching_Assignments',
    headers: ['AssignmentID', 'TeacherID', 'ClassID', 'Status', 'CreatedAt', 'UpdatedAt']
  },
  LEADING_TEACHERS: {
    name: 'Leading_Teachers',
    headers: ['LeadingTeacherID', 'LeadingTeacherName', 'Status', 'CreatedAt', 'UpdatedAt']
  },
  LEADING_TEACHER_ASSIGNMENTS: {
    name: 'Leading_Teacher_Assignments',
    headers: ['AssignmentID', 'LeadingTeacherID', 'ClassID', 'Period', 'Status', 'CreatedAt', 'UpdatedAt']
  }
};

var DEFAULT_PUBLIC_SETTINGS = {
  schoolName: 'Maldives Academic High School',
  schoolTimezone: 'Indian/Maldives',
  destinations: JSON.stringify([
    'Washroom',
    'Science Laboratory',
    'Library',
    'Main Office',
    'First Aid / Nurse',
    'Counselor Office',
    'Prayer Room'
  ]),
  warningMinutes: '10',
  criticalMinutes: '15',
  sessionTimeoutMinutes: '480',
  teacherSessionTimeoutMinutes: '240'
};

var SENSITIVE_SETTING_KEY_REGEX = /(passcode|password|pin|hash|secret|salt|token|credential|auth_|signing_)/i;

/**
 * HTTP GET handler for Google Apps Script Web App.
 */
function doGet(e) {
  return handleRequest_('GET', e);
}

/**
 * HTTP POST handler for Google Apps Script Web App.
 */
function doPost(e) {
  return handleRequest_('POST', e);
}

/**
 * Unified request dispatcher with consistent JSON envelope and safe error handling.
 */
function handleRequest_(method, e) {
  try {
    var payload = parseRequestPayload_(method, e);
    var action = String(payload.action || '').trim();
    if (!action) {
      return jsonResponse_(false, null, 'Missing required API action parameter.', 'MISSING_ACTION');
    }

    var ss = getSpreadsheet_();
    ensureWorkbookInitialized_(ss);
    migratePlaintextCredentialsIfAny_(ss);

    var result = dispatchAction_(ss, action, payload);
    if (result && result.success === false) {
      return createJsonOutput_(result);
    }
    return jsonResponse_(true, result.data !== undefined ? result.data : result, result.message || 'Operation completed successfully');
  } catch (err) {
    var safeMsg = (err && err.isPublicError && err.message) ? err.message : 'An unexpected server error occurred while processing the request.';
    var errCode = (err && err.code) ? err.code : 'INTERNAL_ERROR';
    return jsonResponse_(false, null, safeMsg, errCode);
  }
}

function appError_(message, code) {
  var err = new Error(message);
  err.isPublicError = true;
  err.code = code || 'BAD_REQUEST';
  return err;
}

function jsonResponse_(success, data, messageOrError, code) {
  var body = {
    success: Boolean(success)
  };
  if (success) {
    body.data = data !== null && data !== undefined ? data : {};
    body.message = messageOrError || 'Operation completed successfully';
  } else {
    body.error = messageOrError || 'Request failed';
    body.code = code || 'ERROR';
  }
  return createJsonOutput_(body);
}

function createJsonOutput_(obj) {
  var jsonString = JSON.stringify(obj);
  if (typeof ContentService !== 'undefined' && ContentService.createTextOutput) {
    return ContentService.createTextOutput(jsonString).setMimeType(ContentService.MimeType.JSON);
  }
  return {
    _rawJson: jsonString,
    getContent: function () { return jsonString; }
  };
}

function parseRequestPayload_(method, e) {
  var payload = {};
  if (!e) {
    return payload;
  }
  if (e.parameter && typeof e.parameter === 'object') {
    var keys = Object.keys(e.parameter);
    for (var i = 0; i < keys.length; i++) {
      payload[keys[i]] = e.parameter[keys[i]];
    }
  }
  if (e.postData && typeof e.postData.contents === 'string' && e.postData.contents.trim() !== '') {
    var raw = e.postData.contents.trim();
    try {
      var parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {
        var pKeys = Object.keys(parsed);
        for (var j = 0; j < pKeys.length; j++) {
          payload[pKeys[j]] = parsed[pKeys[j]];
        }
      } else {
        throw appError_('Request body must be a valid JSON object.', 'MALFORMED_JSON');
      }
    } catch (parseErr) {
      if (parseErr && parseErr.isPublicError) throw parseErr;
      throw appError_('Malformed JSON in request body.', 'MALFORMED_JSON');
    }
  }
  return payload;
}

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var configuredId = props ? String(props.getProperty('SPREADSHEET_ID') || '').trim() : '';
  if (configuredId) {
    var byId = SpreadsheetApp.openById(configuredId);
    if (!byId) {
      throw appError_('Configured Google Spreadsheet could not be opened. Check SPREADSHEET_ID in Script Properties.', 'SPREADSHEET_NOT_FOUND');
    }
    return byId;
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }
  throw appError_('No Google Spreadsheet is bound or configured in Script Properties (SPREADSHEET_ID).', 'SPREADSHEET_NOT_CONFIGURED');
}

function ensureWorkbookInitialized_(ss) {
  var sheetKeys = Object.keys(SMT_SHEETS);
  for (var i = 0; i < sheetKeys.length; i++) {
    var spec = SMT_SHEETS[sheetKeys[i]];
    var sheet = ss.getSheetByName(spec.name);
    if (!sheet) {
      sheet = ss.insertSheet(spec.name);
      sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    } else {
      var lastCol = sheet.getLastColumn();
      var lastRow = sheet.getLastRow();
      if (lastRow === 0 || lastCol === 0) {
        sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
      } else {
        var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
          return String(h || '').trim();
        });
        var missingHeaders = [];
        for (var hIdx = 0; hIdx < spec.headers.length; hIdx++) {
          var reqHeader = spec.headers[hIdx];
          var found = false;
          for (var eIdx = 0; eIdx < existingHeaders.length; eIdx++) {
            if (existingHeaders[eIdx].toLowerCase() === reqHeader.toLowerCase()) {
              found = true;
              break;
            }
          }
          if (!found) {
            missingHeaders.push(reqHeader);
          }
        }
        if (missingHeaders.length > 0) {
          sheet.getRange(1, lastCol + 1, 1, missingHeaders.length).setValues([missingHeaders]);
        }
      }
    }
  }

  var settingsMap = getRawSettingsMap_(ss);
  var defaultKeys = Object.keys(DEFAULT_PUBLIC_SETTINGS);
  var missingSettingRows = [];
  for (var k = 0; k < defaultKeys.length; k++) {
    var key = defaultKeys[k];
    if (settingsMap[key] === undefined) {
      missingSettingRows.push([key, DEFAULT_PUBLIC_SETTINGS[key]]);
    }
  }
  if (missingSettingRows.length > 0) {
    var settingsSheet = ss.getSheetByName(SMT_SHEETS.SETTINGS.name);
    var startRow = Math.max(settingsSheet.getLastRow() + 1, 2);
    settingsSheet.getRange(startRow, 1, missingSettingRows.length, 2).setValues(missingSettingRows);
  }
}

function readSheetRecords_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw appError_('Required sheet "' + sheetName + '" is missing.', 'SHEET_MISSING');
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { headers: [], headerMap: {}, records: [] };
  }
  var allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var rawHeaders = allValues[0];
  var headers = [];
  var headerMap = {};
  for (var c = 0; c < rawHeaders.length; c++) {
    var hName = String(rawHeaders[c] || '').trim();
    headers.push(hName);
    if (hName) {
      headerMap[hName.toLowerCase()] = { index: c, canonical: hName };
    }
  }
  var records = [];
  for (var r = 1; r < allValues.length; r++) {
    var row = allValues[r];
    var isBlank = true;
    var obj = { _rowNumber: r + 1 };
    for (var col = 0; col < headers.length; col++) {
      var colName = headers[col];
      if (!colName) continue;
      var val = row[col];
      if (val !== '' && val !== null && val !== undefined) {
        isBlank = false;
      }
      obj[colName] = val;
    }
    if (!isBlank) {
      records.push(obj);
    }
  }
  return {
    sheet: sheet,
    headers: headers,
    headerMap: headerMap,
    records: records
  };
}

function appendSheetRecord_(ss, sheetName, recordObj) {
  var table = readSheetRecords_(ss, sheetName);
  var rowArray = new Array(table.headers.length);
  for (var i = 0; i < table.headers.length; i++) {
    var h = table.headers[i];
    var matchedVal = recordObj[h];
    if (matchedVal === undefined) {
      var objKeys = Object.keys(recordObj);
      for (var k = 0; k < objKeys.length; k++) {
        if (objKeys[k].toLowerCase() === h.toLowerCase()) {
          matchedVal = recordObj[objKeys[k]];
          break;
        }
      }
    }
    rowArray[i] = (matchedVal !== undefined && matchedVal !== null) ? matchedVal : '';
  }
  var nextRow = Math.max(table.sheet.getLastRow() + 1, 2);
  table.sheet.getRange(nextRow, 1, 1, rowArray.length).setValues([rowArray]);
  return nextRow;
}

function updateSheetRecordByRow_(table, rowNumber, updatesObj) {
  var currentRow = table.sheet.getRange(rowNumber, 1, 1, table.headers.length).getValues()[0];
  var updateKeys = Object.keys(updatesObj);
  for (var i = 0; i < updateKeys.length; i++) {
    var key = updateKeys[i];
    var entry = table.headerMap[key.toLowerCase()];
    if (entry) {
      currentRow[entry.index] = updatesObj[key];
    }
  }
  table.sheet.getRange(rowNumber, 1, 1, table.headers.length).setValues([currentRow]);
}

function getRawSettingsMap_(ss) {
  var table = readSheetRecords_(ss, SMT_SHEETS.SETTINGS.name);
  var map = {};
  for (var i = 0; i < table.records.length; i++) {
    var rec = table.records[i];
    var k = String(rec.SettingKey || '').trim();
    if (k) {
      map[k] = rec.SettingValue !== undefined && rec.SettingValue !== null ? String(rec.SettingValue) : '';
    }
  }
  return map;
}

function upsertSettingsMap_(ss, updates) {
  var table = readSheetRecords_(ss, SMT_SHEETS.SETTINGS.name);
  var rowByKey = {};
  for (var i = 0; i < table.records.length; i++) {
    var rec = table.records[i];
    var k = String(rec.SettingKey || '').trim();
    if (k) {
      rowByKey[k.toLowerCase()] = rec._rowNumber;
    }
  }
  var keys = Object.keys(updates);
  var toAppend = [];
  for (var j = 0; j < keys.length; j++) {
    var key = keys[j];
    var val = String(updates[key] !== undefined && updates[key] !== null ? updates[key] : '');
    var existingRow = rowByKey[key.toLowerCase()];
    if (existingRow) {
      updateSheetRecordByRow_(table, existingRow, { SettingKey: key, SettingValue: val });
    } else {
      toAppend.push([key, val]);
    }
  }
  if (toAppend.length > 0) {
    var startRow = Math.max(table.sheet.getLastRow() + 1, 2);
    table.sheet.getRange(startRow, 1, toAppend.length, 2).setValues(toAppend);
  }
}

// ============================================================================
// CRYPTOGRAPHY, HASHING, CONSTANT-TIME COMPARISON & TOKEN SIGNING
// ============================================================================

function constantTimeEqual_(a, b) {
  var strA = String(a || '');
  var strB = String(b || '');
  var lenA = strA.length;
  var lenB = strB.length;
  var maxLen = Math.max(lenA, lenB);
  var diff = lenA ^ lenB;
  for (var i = 0; i < maxLen; i++) {
    var codeA = i < lenA ? strA.charCodeAt(i) : 0;
    var codeB = i < lenB ? strB.charCodeAt(i) : 0;
    diff |= (codeA ^ codeB);
  }
  return diff === 0 && lenA > 0;
}

function bytesToHex_(bytes) {
  var hex = [];
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i];
    if (b < 0) b += 256;
    var s = b.toString(16);
    hex.push(s.length === 1 ? '0' + s : s);
  }
  return hex.join('');
}

function sha256Hex_(input) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(input), Utilities.Charset.UTF_8);
  return bytesToHex_(bytes);
}

function hmacSha256Hex_(message, secret) {
  if (!secret || String(secret).trim() === '') {
    throw appError_('Token signing secret is missing.', 'CRYPTO_CONFIG_ERROR');
  }
  var bytes = Utilities.computeHmacSha256Signature(String(message), String(secret), Utilities.Charset.UTF_8);
  return bytesToHex_(bytes);
}

function generateRandomHex_(byteCount) {
  var uuid1 = Utilities.getUuid().replace(/-/g, '');
  var uuid2 = Utilities.getUuid().replace(/-/g, '');
  var combined = sha256Hex_(uuid1 + ':' + new Date().getTime() + ':' + uuid2 + ':' + Math.random());
  return combined.substring(0, (byteCount || 16) * 2);
}

function hashCredential_(plaintext, optionalSalt, optionalIterations) {
  var clean = String(plaintext || '');
  if (!clean) {
    throw appError_('Cannot hash an empty credential.', 'INVALID_CREDENTIAL');
  }
  var salt = optionalSalt || generateRandomHex_(16);
  var iterations = optionalIterations || 1200;
  var digest = sha256Hex_(salt + ':' + clean);
  for (var i = 1; i < iterations; i++) {
    digest = sha256Hex_(digest + ':' + salt + ':' + clean + ':' + i);
  }
  return 'sha256$' + iterations + '$' + salt + '$' + digest;
}

function verifyCredentialHash_(plaintext, storedHash) {
  if (!plaintext || !storedHash) return false;
  var strHash = String(storedHash).trim();
  var parts = strHash.split('$');
  if (parts.length !== 4 || parts[0] !== 'sha256') {
    return false;
  }
  var iterations = parseInt(parts[1], 10);
  var salt = parts[2];
  var expectedDigest = parts[3];
  if (!iterations || iterations < 1 || !salt || !expectedDigest) {
    return false;
  }
  var computed = hashCredential_(plaintext, salt, iterations);
  return constantTimeEqual_(computed, strHash);
}

function migratePlaintextCredentialsIfAny_(ss) {
  var rawMap = getRawSettingsMap_(ss);
  var credentialKeys = [
    { plain: 'adminPasscode', hashKey: 'AUTH_HASH_ADMIN' },
    { plain: 'principalPasscode', hashKey: 'AUTH_HASH_PRINCIPAL' },
    { plain: 'leadingTeacherPasscode', hashKey: 'AUTH_HASH_LEADING_TEACHER' },
    { plain: 'teacherPin', hashKey: 'AUTH_HASH_TEACHER_PIN' },
    { plain: 'AUTH_HASH_ADMIN', hashKey: 'AUTH_HASH_ADMIN' },
    { plain: 'AUTH_HASH_PRINCIPAL', hashKey: 'AUTH_HASH_PRINCIPAL' },
    { plain: 'AUTH_HASH_LEADING_TEACHER', hashKey: 'AUTH_HASH_LEADING_TEACHER' },
    { plain: 'AUTH_HASH_TEACHER_PIN', hashKey: 'AUTH_HASH_TEACHER_PIN' }
  ];
  var updates = {};
  var changed = false;
  for (var i = 0; i < credentialKeys.length; i++) {
    var item = credentialKeys[i];
    var val = rawMap[item.plain];
    if (val && String(val).trim() !== '' && String(val).indexOf('sha256$') !== 0) {
      updates[item.hashKey] = hashCredential_(String(val).trim());
      if (item.plain !== item.hashKey) {
        updates[item.plain] = '';
      }
      changed = true;
    }
  }
  if (changed) {
    upsertSettingsMap_(ss, updates);
  }
}

function getTeacherTokenSecret_(ss, rotate) {
  var props = PropertiesService.getScriptProperties();
  var secret = props ? props.getProperty('TEACHER_TOKEN_SECRET') : null;
  if (!secret || rotate) {
    secret = 'TCHSEC_' + generateRandomHex_(32);
    if (props) {
      props.setProperty('TEACHER_TOKEN_SECRET', secret);
    }
  }
  return secret;
}

function getRoleTokenSecret_(role, rotate) {
  var normRole = String(role || '').toUpperCase().replace(/[^A-Z_]/g, '');
  if (!normRole || normRole === 'TEACHER') {
    throw appError_('Invalid privileged role for role token secret.', 'INVALID_ROLE');
  }
  var propKey = 'ROLE_TOKEN_SECRET_' + normRole;
  var props = PropertiesService.getScriptProperties();
  var secret = props ? props.getProperty(propKey) : null;
  if (!secret || rotate) {
    secret = 'ROLESEC_' + normRole + '_' + generateRandomHex_(32);
    if (props) {
      props.setProperty(propKey, secret);
    }
  }
  return secret;
}

function base64UrlEncode_(str) {
  return Utilities.base64EncodeWebSafe(String(str), Utilities.Charset.UTF_8).replace(/=+$/, '');
}

function base64UrlDecode_(b64) {
  var str = String(b64 || '');
  while (str.length % 4 !== 0) {
    str += '=';
  }
  var bytes = Utilities.base64DecodeWebSafe(str, Utilities.Charset.UTF_8);
  return Utilities.newBlob(bytes).getDataAsString('UTF-8');
}

function issueTeacherToken_(ss, teacherId, teacherName, sessionMinutes) {
  var secret = getTeacherTokenSecret_(ss, false);
  var nowMs = new Date().getTime();
  var ttlMs = Math.max(parseInt(sessionMinutes || 240, 10), 5) * 60 * 1000;
  var claims = {
    typ: 'SMT_TEACHER',
    role: 'teacher',
    teacherId: String(teacherId).trim(),
    teacherName: String(teacherName).trim(),
    iat: nowMs,
    exp: nowMs + ttlMs,
    jti: generateRandomHex_(8)
  };
  var encodedClaims = base64UrlEncode_(JSON.stringify(claims));
  var sig = hmacSha256Hex_(encodedClaims, secret);
  return {
    token: encodedClaims + '.' + sig,
    expiresAt: new Date(claims.exp).toISOString(),
    teacherId: claims.teacherId,
    teacherName: claims.teacherName,
    role: 'teacher'
  };
}

function verifyTeacherTokenInternal_(ss, tokenString) {
  if (!tokenString || typeof tokenString !== 'string') {
    throw appError_('Teacher authentication token is required.', 'AUTH_REQUIRED');
  }
  var parts = String(tokenString).trim().split('.');
  if (parts.length !== 2 || !parts[0] || !parts[1]) {
    throw appError_('Malformed teacher authentication token.', 'INVALID_TOKEN');
  }
  var secret = getTeacherTokenSecret_(ss, false);
  var expectedSig = hmacSha256Hex_(parts[0], secret);
  if (!constantTimeEqual_(parts[1], expectedSig)) {
    throw appError_('Invalid or tampered teacher token signature.', 'INVALID_TOKEN');
  }
  var claims;
  try {
    claims = JSON.parse(base64UrlDecode_(parts[0]));
  } catch (e) {
    throw appError_('Corrupted teacher token payload.', 'INVALID_TOKEN');
  }
  if (!claims || claims.typ !== 'SMT_TEACHER' || claims.role !== 'teacher' || !claims.teacherId) {
    throw appError_('Token is not a valid teacher session token.', 'INVALID_TOKEN');
  }
  var nowMs = new Date().getTime();
  if (!claims.exp || nowMs > claims.exp) {
    throw appError_('Teacher session has expired. Please sign in again.', 'TOKEN_EXPIRED');
  }

  var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
  var matchedTeacher = null;
  for (var i = 0; i < teachersTable.records.length; i++) {
    var t = teachersTable.records[i];
    if (String(t.TeacherID).trim() === String(claims.teacherId).trim()) {
      matchedTeacher = t;
      break;
    }
  }
  if (!matchedTeacher || String(matchedTeacher.Status || '').trim().toLowerCase() !== 'active') {
    throw appError_('Teacher account is no longer active.', 'ACCOUNT_INACTIVE');
  }

  return {
    role: 'teacher',
    teacherId: String(matchedTeacher.TeacherID).trim(),
    teacherName: String(matchedTeacher.TeacherName).trim(),
    claims: claims
  };
}

function issueRoleToken_(ss, role, extraClaims, sessionMinutes) {
  var normRole = String(role || '').trim().toLowerCase();
  if (normRole !== 'admin' && normRole !== 'principal' && normRole !== 'leading_teacher') {
    throw appError_('Unsupported privileged role: ' + role, 'INVALID_ROLE');
  }
  var secret = getRoleTokenSecret_(normRole, false);
  var nowMs = new Date().getTime();
  var ttlMs = Math.max(parseInt(sessionMinutes || 480, 10), 5) * 60 * 1000;
  var claims = {
    typ: 'SMT_ROLE',
    role: normRole,
    iat: nowMs,
    exp: nowMs + ttlMs,
    jti: generateRandomHex_(8)
  };
  if (extraClaims && typeof extraClaims === 'object') {
    var keys = Object.keys(extraClaims);
    for (var i = 0; i < keys.length; i++) {
      claims[keys[i]] = extraClaims[keys[i]];
    }
  }
  var encodedClaims = base64UrlEncode_(JSON.stringify(claims));
  var sig = hmacSha256Hex_(encodedClaims, secret);
  return {
    token: encodedClaims + '.' + sig,
    expiresAt: new Date(claims.exp).toISOString(),
    role: normRole,
    leadingTeacherId: claims.leadingTeacherId || null,
    leadingTeacherName: claims.leadingTeacherName || null
  };
}

function requireAuth_(ss, payload, allowedRoles) {
  var tokenString = String(payload.token || payload.authToken || '').trim();
  if (!tokenString) {
    throw appError_('Authentication token is required for this operation.', 'AUTH_REQUIRED');
  }
  var parts = tokenString.split('.');
  if (parts.length !== 2 || !parts[0] || !parts[1]) {
    throw appError_('Malformed authentication token.', 'INVALID_TOKEN');
  }
  var claims;
  try {
    claims = JSON.parse(base64UrlDecode_(parts[0]));
  } catch (e) {
    throw appError_('Invalid authentication token encoding.', 'INVALID_TOKEN');
  }
  if (!claims || !claims.role) {
    throw appError_('Invalid token claims.', 'INVALID_TOKEN');
  }

  var role = String(claims.role).toLowerCase();
  if (allowedRoles && allowedRoles.indexOf(role) === -1) {
    throw appError_('Access denied: Role "' + role + '" is not authorized for this action.', 'FORBIDDEN');
  }

  if (role === 'teacher') {
    return verifyTeacherTokenInternal_(ss, tokenString);
  }

  if (claims.typ !== 'SMT_ROLE') {
    throw appError_('Invalid role token type.', 'INVALID_TOKEN');
  }
  var secret = getRoleTokenSecret_(role, false);
  var expectedSig = hmacSha256Hex_(parts[0], secret);
  if (!constantTimeEqual_(parts[1], expectedSig)) {
    throw appError_('Invalid or expired session signature. Please log in again.', 'INVALID_TOKEN');
  }
  var nowMs = new Date().getTime();
  if (!claims.exp || nowMs > claims.exp) {
    throw appError_('Session has expired. Please log in again.', 'TOKEN_EXPIRED');
  }

  if (role === 'leading_teacher') {
    if (!claims.leadingTeacherId) {
      throw appError_('Leading Teacher session is missing profile binding.', 'INVALID_TOKEN');
    }
    var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);
    var ltProfile = null;
    for (var i = 0; i < ltTable.records.length; i++) {
      if (String(ltTable.records[i].LeadingTeacherID).trim() === String(claims.leadingTeacherId).trim()) {
        ltProfile = ltTable.records[i];
        break;
      }
    }
    if (!ltProfile || String(ltProfile.Status || '').trim().toLowerCase() !== 'active') {
      throw appError_('Leading Teacher profile is inactive or unavailable.', 'ACCOUNT_INACTIVE');
    }
    var assignedClassIds = getActiveAssignedClassIdsForLeadingTeacher_(ss, claims.leadingTeacherId);
    return {
      role: 'leading_teacher',
      leadingTeacherId: String(ltProfile.LeadingTeacherID).trim(),
      leadingTeacherName: String(ltProfile.LeadingTeacherName).trim(),
      assignedClassIds: assignedClassIds,
      claims: claims
    };
  }

  return {
    role: role,
    claims: claims
  };
}

function getActiveAssignedClassIdsForLeadingTeacher_(ss, leadingTeacherId) {
  var ltaTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHER_ASSIGNMENTS.name);
  var classTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
  var activeClassMap = {};
  for (var c = 0; c < classTable.records.length; c++) {
    var cls = classTable.records[c];
    if (String(cls.Status || '').trim().toLowerCase() === 'active') {
      activeClassMap[String(cls.ClassID).trim()] = true;
    }
  }
  var assigned = [];
  var seen = {};
  for (var i = 0; i < ltaTable.records.length; i++) {
    var rec = ltaTable.records[i];
    var ltId = String(rec.LeadingTeacherID || '').trim();
    var cId = String(rec.ClassID || '').trim();
    var status = String(rec.Status || '').trim().toLowerCase();
    if (ltId === String(leadingTeacherId).trim() && status === 'active' && activeClassMap[cId] && !seen[cId]) {
      seen[cId] = true;
      assigned.push(cId);
    }
  }
  return assigned;
}

// ============================================================================
// AUTHORITATIVE TIMESTAMP & DURATION ENGINE
// ============================================================================

function nowIso_() {
  return new Date().toISOString();
}

function generateMovementId_(epochMs) {
  var ms = epochMs || new Date().getTime();
  var suffix = generateRandomHex_(3).toUpperCase();
  return 'MOV-' + ms + '-' + suffix;
}

function extractEpochFromMovementId_(movementId) {
  var m = /^MOV-(\d{12,14})-[A-Z0-9]+$/i.exec(String(movementId || '').trim());
  if (m && m[1]) {
    var parsed = parseInt(m[1], 10);
    if (!isNaN(parsed) && parsed > 1577836800000) {
      return parsed;
    }
  }
  return null;
}

function parseAuthoritativeTimestamp_(rawValue, movementId, fallbackReferenceMs) {
  var embeddedEpoch = extractEpochFromMovementId_(movementId);

  if (rawValue === null || rawValue === undefined || rawValue === '') {
    return embeddedEpoch ? new Date(embeddedEpoch) : null;
  }

  if (Object.prototype.toString.call(rawValue) === '[object Date]') {
    var timeMs = rawValue.getTime();
    if (isNaN(timeMs)) {
      return embeddedEpoch ? new Date(embeddedEpoch) : null;
    }
    var yr = rawValue.getUTCFullYear();
    if (yr <= 1905) {
      if (embeddedEpoch) {
        return new Date(embeddedEpoch);
      }
      var refDate = fallbackReferenceMs ? new Date(fallbackReferenceMs) : new Date();
      var corrected = new Date(refDate.getTime());
      corrected.setHours(rawValue.getHours(), rawValue.getMinutes(), rawValue.getSeconds(), 0);
      return corrected;
    }
    return rawValue;
  }

  var strVal = String(rawValue).trim();
  if (!strVal) {
    return embeddedEpoch ? new Date(embeddedEpoch) : null;
  }

  if (/^\d{12,14}$/.test(strVal)) {
    return new Date(parseInt(strVal, 10));
  }

  var timeOnlyMatch = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i.exec(strVal);
  if (timeOnlyMatch) {
    if (embeddedEpoch && !fallbackReferenceMs) {
      return new Date(embeddedEpoch);
    }
    var baseMs = fallbackReferenceMs || embeddedEpoch || new Date().getTime();
    var baseDate = new Date(baseMs);
    var hours = parseInt(timeOnlyMatch[1], 10);
    var mins = parseInt(timeOnlyMatch[2], 10);
    var secs = timeOnlyMatch[3] ? parseInt(timeOnlyMatch[3], 10) : 0;
    var ampm = timeOnlyMatch[4] ? timeOnlyMatch[4].toUpperCase() : null;
    if (ampm === 'PM' && hours < 12) hours += 12;
    if (ampm === 'AM' && hours === 12) hours = 0;
    baseDate.setHours(hours, mins, secs, 0);
    return baseDate;
  }

  var parsedDate = new Date(strVal);
  if (!isNaN(parsedDate.getTime())) {
    if (parsedDate.getUTCFullYear() <= 1905) {
      if (embeddedEpoch) return new Date(embeddedEpoch);
      var nowRef = new Date(fallbackReferenceMs || new Date().getTime());
      nowRef.setHours(parsedDate.getHours(), parsedDate.getMinutes(), parsedDate.getSeconds(), 0);
      return nowRef;
    }
    return parsedDate;
  }

  return embeddedEpoch ? new Date(embeddedEpoch) : null;
}

function computeMovementMetrics_(timeOutRaw, timeInRaw, movementId, warningMinutes, criticalMinutes, nowMsOverride) {
  var warnMins = parseFloat(warningMinutes || 10);
  var critMins = parseFloat(criticalMinutes || 15);
  var checkoutDate = parseAuthoritativeTimestamp_(timeOutRaw, movementId, null);
  if (!checkoutDate) {
    return {
      checkoutIso: '',
      returnIso: '',
      elapsedSeconds: 0,
      elapsedMinutes: 0,
      formattedDuration: '0m 00s',
      attentionLevel: 'Normal'
    };
  }
  var checkoutMs = checkoutDate.getTime();
  var returnDate = (timeInRaw !== null && timeInRaw !== undefined && String(timeInRaw).trim() !== '')
    ? parseAuthoritativeTimestamp_(timeInRaw, null, checkoutMs)
    : null;

  var endMs = returnDate ? returnDate.getTime() : (nowMsOverride || new Date().getTime());
  var elapsedSeconds = Math.max(0, Math.floor((endMs - checkoutMs) / 1000));
  var elapsedMinutes = elapsedSeconds / 60;
  var wholeMins = Math.floor(elapsedSeconds / 60);
  var remSecs = elapsedSeconds % 60;
  var formattedDuration = wholeMins + 'm ' + (remSecs < 10 ? '0' + remSecs : remSecs) + 's';

  var attentionLevel = 'Normal';
  if (elapsedMinutes >= critMins) {
    attentionLevel = 'Critical';
  } else if (elapsedMinutes >= warnMins) {
    attentionLevel = 'Warning';
  }

  return {
    checkoutIso: checkoutDate.toISOString(),
    returnIso: returnDate ? returnDate.toISOString() : '',
    elapsedSeconds: elapsedSeconds,
    elapsedMinutes: Math.round(elapsedMinutes * 100) / 100,
    formattedDuration: formattedDuration,
    attentionLevel: attentionLevel
  };
}

// ============================================================================
// ACTION DISPATCHER
// ============================================================================

function dispatchAction_(ss, action, payload) {
  switch (action) {
    case 'getSettings':
      return actionGetSettings_(ss);
    case 'getTeacherLoginOptions':
      return actionGetTeacherLoginOptions_(ss);
    case 'getLeadingTeacherLoginOptions':
      return actionGetLeadingTeacherLoginOptions_(ss);
    case 'getStudents':
      return actionGetStudents_(ss, payload);
    case 'getClasses':
      return actionGetClasses_(ss, payload);
    case 'getCurrentOut':
      return actionGetCurrentOut_(ss, payload);

    case 'verifyPasscode':
      return actionVerifyPasscode_(ss, payload);
    case 'verifyTeacherPasscode':
      return actionVerifyTeacherPasscode_(ss, payload);
    case 'verifyTeacherToken':
      return actionVerifyTeacherToken_(ss, payload);
    case 'changePasscode':
      return actionChangePasscode_(ss, payload);
    case 'getSecuritySettings':
      return actionGetSecuritySettings_(ss, payload);

    case 'getTeachers':
      return actionGetTeachers_(ss, payload);
    case 'addTeacher':
      return actionAddTeacher_(ss, payload);
    case 'updateTeacher':
      return actionUpdateTeacher_(ss, payload);
    case 'deactivateTeacher':
      return actionDeactivateTeacher_(ss, payload);
    case 'bulkAddTeachers':
      return actionBulkAddTeachers_(ss, payload);
    case 'getTeachingAssignments':
      return actionGetTeachingAssignments_(ss, payload);
    case 'addTeachingAssignment':
      return actionAddTeachingAssignment_(ss, payload);
    case 'updateTeachingAssignment':
      return actionUpdateTeachingAssignment_(ss, payload);
    case 'deactivateTeachingAssignment':
      return actionDeactivateTeachingAssignment_(ss, payload);

    case 'getLeadingTeachers':
      return actionGetLeadingTeachers_(ss, payload);
    case 'addLeadingTeacher':
      return actionAddLeadingTeacher_(ss, payload);
    case 'updateLeadingTeacher':
      return actionUpdateLeadingTeacher_(ss, payload);
    case 'deactivateLeadingTeacher':
      return actionDeactivateLeadingTeacher_(ss, payload);
    case 'getLeadingTeacherAssignments':
      return actionGetLeadingTeacherAssignments_(ss, payload);
    case 'addLeadingTeacherAssignment':
      return actionAddLeadingTeacherAssignment_(ss, payload);
    case 'updateLeadingTeacherAssignment':
      return actionUpdateLeadingTeacherAssignment_(ss, payload);
    case 'deactivateLeadingTeacherAssignment':
      return actionDeactivateLeadingTeacherAssignment_(ss, payload);

    case 'addStudent':
      return actionAddStudent_(ss, payload);
    case 'updateStudent':
      return actionUpdateStudent_(ss, payload);
    case 'deactivateStudent':
      return actionDeactivateStudent_(ss, payload);
    case 'bulkAddStudents':
      return actionBulkAddStudents_(ss, payload);

    case 'addClass':
      return actionAddClass_(ss, payload);
    case 'updateClass':
      return actionUpdateClass_(ss, payload);
    case 'deactivateClass':
      return actionDeactivateClass_(ss, payload);

    case 'recordOut':
      return actionRecordOut_(ss, payload);
    case 'recordReturn':
      return actionRecordReturn_(ss, payload);
    case 'getMovementHistory':
      return actionGetMovementHistory_(ss, payload);
    case 'getMovementReports':
      return actionGetMovementReports_(ss, payload);
    case 'getAdminOverview':
      return actionGetAdminOverview_(ss, payload);

    case 'updateSettings':
      return actionUpdateSettings_(ss, payload);
    case 'manageDestinations':
      return actionManageDestinations_(ss, payload);

    default:
      throw appError_('Unknown API action: ' + action, 'UNKNOWN_ACTION');
  }
}

// ============================================================================
// GENERAL & BOOTSTRAP ACTIONS
// ============================================================================

function actionGetSettings_(ss) {
  var raw = getRawSettingsMap_(ss);
  var safeSettings = {};
  var keys = Object.keys(raw);
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (SENSITIVE_SETTING_KEY_REGEX.test(k)) {
      continue;
    }
    safeSettings[k] = raw[k];
  }
  var destinations = [];
  try {
    destinations = JSON.parse(safeSettings.destinations || '[]');
    if (!Array.isArray(destinations)) destinations = [];
  } catch (e) {
    destinations = String(safeSettings.destinations || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  }
  return {
    data: {
      schoolName: safeSettings.schoolName || DEFAULT_PUBLIC_SETTINGS.schoolName,
      schoolTimezone: safeSettings.schoolTimezone || DEFAULT_PUBLIC_SETTINGS.schoolTimezone,
      destinations: destinations,
      warningMinutes: Number(safeSettings.warningMinutes || 10),
      criticalMinutes: Number(safeSettings.criticalMinutes || 15),
      sessionTimeoutMinutes: Number(safeSettings.sessionTimeoutMinutes || 480),
      teacherSessionTimeoutMinutes: Number(safeSettings.teacherSessionTimeoutMinutes || 240)
    },
    message: 'Settings retrieved successfully'
  };
}

function actionGetTeacherLoginOptions_(ss) {
  var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
  var activeTeachers = [];
  for (var i = 0; i < teachersTable.records.length; i++) {
    var t = teachersTable.records[i];
    if (String(t.Status || '').trim().toLowerCase() === 'active') {
      activeTeachers.push({
        TeacherID: String(t.TeacherID).trim(),
        TeacherName: String(t.TeacherName).trim()
      });
    }
  }
  activeTeachers.sort(function (a, b) {
    return a.TeacherName.localeCompare(b.TeacherName);
  });
  var settingsResult = actionGetSettings_(ss).data;
  return {
    data: {
      teachers: activeTeachers,
      schoolName: settingsResult.schoolName,
      schoolTimezone: settingsResult.schoolTimezone,
      destinations: settingsResult.destinations,
      warningMinutes: settingsResult.warningMinutes,
      criticalMinutes: settingsResult.criticalMinutes
    },
    message: 'Teacher login options retrieved'
  };
}

function actionGetLeadingTeacherLoginOptions_(ss) {
  var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);
  var activeLeadingTeachers = [];
  for (var i = 0; i < ltTable.records.length; i++) {
    var lt = ltTable.records[i];
    if (String(lt.Status || '').trim().toLowerCase() === 'active') {
      activeLeadingTeachers.push({
        LeadingTeacherID: String(lt.LeadingTeacherID).trim(),
        LeadingTeacherName: String(lt.LeadingTeacherName).trim()
      });
    }
  }
  activeLeadingTeachers.sort(function (a, b) {
    return a.LeadingTeacherName.localeCompare(b.LeadingTeacherName);
  });
  var settingsResult = actionGetSettings_(ss).data;
  return {
    data: {
      leadingTeachers: activeLeadingTeachers,
      schoolName: settingsResult.schoolName
    },
    message: 'Leading Teacher login options retrieved'
  };
}

function actionGetClasses_(ss, payload) {
  var auth = requireAuth_(ss, payload, ['teacher', 'leading_teacher', 'principal', 'admin']);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
  var studentsTable = readSheetRecords_(ss, SMT_SHEETS.STUDENTS.name);
  var includeInactive = Boolean(payload.includeInactive) && auth.role === 'admin';

  var enrollmentByClass = {};
  var activeEnrollmentByClass = {};
  for (var s = 0; s < studentsTable.records.length; s++) {
    var stu = studentsTable.records[s];
    var cid = String(stu.ClassID || '').trim();
    if (!cid) continue;
    enrollmentByClass[cid] = (enrollmentByClass[cid] || 0) + 1;
    if (String(stu.Status || '').trim().toLowerCase() === 'active') {
      activeEnrollmentByClass[cid] = (activeEnrollmentByClass[cid] || 0) + 1;
    }
  }

  var allowedSet = null;
  if (auth.role === 'leading_teacher') {
    allowedSet = {};
    for (var a = 0; a < auth.assignedClassIds.length; a++) {
      allowedSet[auth.assignedClassIds[a]] = true;
    }
  }

  var classes = [];
  for (var i = 0; i < classesTable.records.length; i++) {
    var c = classesTable.records[i];
    var classId = String(c.ClassID || '').trim();
    var status = String(c.Status || 'Active').trim();
    if (!includeInactive && status.toLowerCase() !== 'active') {
      continue;
    }
    if (allowedSet && !allowedSet[classId]) {
      continue;
    }
    classes.push({
      ClassID: classId,
      ClassName: String(c.ClassName || '').trim(),
      Status: status,
      CreatedAt: String(c.CreatedAt || ''),
      UpdatedAt: String(c.UpdatedAt || ''),
      enrolledCount: enrollmentByClass[classId] || 0,
      activeEnrolledCount: activeEnrollmentByClass[classId] || 0
    });
  }

  classes.sort(function (a, b) {
    return a.ClassName.localeCompare(b.ClassName);
  });

  return {
    data: { classes: classes },
    message: 'Classes retrieved successfully'
  };
}

function actionGetStudents_(ss, payload) {
  var auth = requireAuth_(ss, payload, ['teacher', 'leading_teacher', 'principal', 'admin']);
  var studentsTable = readSheetRecords_(ss, SMT_SHEETS.STUDENTS.name);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
  var movementsTable = readSheetRecords_(ss, SMT_SHEETS.MOVEMENT_LOG.name);

  var classMap = {};
  for (var c = 0; c < classesTable.records.length; c++) {
    var cls = classesTable.records[c];
    classMap[String(cls.ClassID).trim()] = String(cls.ClassName || '').trim();
  }

  var activeOutByStudent = {};
  for (var m = 0; m < movementsTable.records.length; m++) {
    var mov = movementsTable.records[m];
    if (String(mov.Status || '').trim().toUpperCase() === 'OUT') {
      activeOutByStudent[String(mov.StudentID).trim()] = String(mov.MovementID).trim();
    }
  }

  var filterClassId = String(payload.classId || '').trim();
  var search = String(payload.search || '').trim().toLowerCase();
  var includeInactive = Boolean(payload.includeInactive) && auth.role === 'admin';

  var allowedClassSet = null;
  if (auth.role === 'leading_teacher') {
    allowedClassSet = {};
    for (var i = 0; i < auth.assignedClassIds.length; i++) {
      allowedClassSet[auth.assignedClassIds[i]] = true;
    }
  }

  var list = [];
  for (var s = 0; s < studentsTable.records.length; s++) {
    var stu = studentsTable.records[s];
    var stuId = String(stu.StudentID || '').trim();
    var stuName = String(stu.StudentName || '').trim();
    var clsId = String(stu.ClassID || '').trim();
    var status = String(stu.Status || 'Active').trim();

    if (!includeInactive && status.toLowerCase() !== 'active') {
      continue;
    }
    if (allowedClassSet && !allowedClassSet[clsId]) {
      continue;
    }
    if (filterClassId && clsId !== filterClassId) {
      continue;
    }
    if (search) {
      if (stuName.toLowerCase().indexOf(search) === -1 && stuId.toLowerCase().indexOf(search) === -1) {
        continue;
      }
    }

    list.push({
      StudentID: stuId,
      StudentName: stuName,
      ClassID: clsId,
      ClassName: classMap[clsId] || clsId,
      Status: status,
      CreatedAt: String(stu.CreatedAt || ''),
      UpdatedAt: String(stu.UpdatedAt || ''),
      isCurrentlyOut: Boolean(activeOutByStudent[stuId]),
      activeMovementId: activeOutByStudent[stuId] || null
    });
  }

  list.sort(function (a, b) {
    return a.StudentName.localeCompare(b.StudentName);
  });

  return {
    data: { students: list },
    message: 'Students retrieved successfully'
  };
}

function actionGetCurrentOut_(ss, payload) {
  var auth = requireAuth_(ss, payload, ['teacher', 'leading_teacher', 'principal', 'admin']);
  var settings = actionGetSettings_(ss).data;
  var movementsTable = readSheetRecords_(ss, SMT_SHEETS.MOVEMENT_LOG.name);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);

  var classMap = {};
  for (var c = 0; c < classesTable.records.length; c++) {
    var cls = classesTable.records[c];
    classMap[String(cls.ClassID).trim()] = String(cls.ClassName || '').trim();
  }

  var allowedClassSet = null;
  if (auth.role === 'leading_teacher') {
    allowedClassSet = {};
    for (var i = 0; i < auth.assignedClassIds.length; i++) {
      allowedClassSet[auth.assignedClassIds[i]] = true;
    }
  }

  var filterClassId = String(payload.classId || '').trim();
  var filterDestination = String(payload.destination || '').trim();
  var filterAttention = String(payload.attentionLevel || '').trim().toLowerCase();
  var search = String(payload.search || '').trim().toLowerCase();
  var nowMs = new Date().getTime();

  var activeRecords = [];
  var affectedClasses = {};
  var occupiedDestinations = {};
  var over10Count = 0;
  var over15Count = 0;

  for (var r = 0; r < movementsTable.records.length; r++) {
    var row = movementsTable.records[r];
    var status = String(row.Status || '').trim().toUpperCase();
    if (status !== 'OUT') continue;

    var classId = String(row.ClassID || '').trim();
    if (allowedClassSet && !allowedClassSet[classId]) {
      continue;
    }

    var movId = String(row.MovementID || '').trim();
    var studentId = String(row.StudentID || '').trim();
    var studentName = String(row.StudentName || '').trim();
    var className = classMap[classId] || String(row.ClassName || classId).trim();
    var teacherId = String(row.TeacherID || '').trim();
    var teacherName = String(row.TeacherName || '').trim();
    var destination = String(row.Destination || '').trim();

    var metrics = computeMovementMetrics_(
      row.TimeOut,
      null,
      movId,
      settings.warningMinutes,
      settings.criticalMinutes,
      nowMs
    );

    affectedClasses[classId] = true;
    if (destination) occupiedDestinations[destination] = true;
    if (metrics.elapsedMinutes >= settings.warningMinutes) over10Count++;
    if (metrics.elapsedMinutes >= settings.criticalMinutes) over15Count++;

    if (filterClassId && classId !== filterClassId) continue;
    if (filterDestination && destination.toLowerCase() !== filterDestination.toLowerCase()) continue;
    if (filterAttention && metrics.attentionLevel.toLowerCase() !== filterAttention) continue;
    if (search) {
      if (studentName.toLowerCase().indexOf(search) === -1 && studentId.toLowerCase().indexOf(search) === -1) {
        continue;
      }
    }

    activeRecords.push({
      MovementID: movId,
      StudentID: studentId,
      StudentName: studentName,
      ClassID: classId,
      ClassName: className,
      TeacherID: teacherId,
      TeacherName: teacherName,
      Destination: destination,
      TimeOut: metrics.checkoutIso,
      Status: 'OUT',
      Notes: String(row.Notes || ''),
      elapsedSeconds: metrics.elapsedSeconds,
      elapsedMinutes: metrics.elapsedMinutes,
      formattedDuration: metrics.formattedDuration,
      attentionLevel: metrics.attentionLevel
    });
  }

  activeRecords.sort(function (a, b) {
    return b.elapsedSeconds - a.elapsedSeconds;
  });

  return {
    data: {
      records: activeRecords,
      stats: {
        totalOut: activeRecords.length,
        scopeTotalOut: Object.keys(affectedClasses).length > 0 ? activeRecords.length : 0,
        affectedClassesCount: Object.keys(affectedClasses).length,
        occupiedDestinationsCount: Object.keys(occupiedDestinations).length,
        warningOrAboveCount: over10Count,
        criticalCount: over15Count
      },
      serverTimestamp: new Date(nowMs).toISOString()
    },
    message: 'Current OUT records retrieved'
  };
}

// ============================================================================
// AUTHENTICATION & CREDENTIAL MANAGEMENT ACTIONS
// ============================================================================

function getStoredCredentialHash_(ss, roleKey) {
  var rawSettings = getRawSettingsMap_(ss);
  var props = PropertiesService.getScriptProperties();
  var keyMap = {
    admin: 'AUTH_HASH_ADMIN',
    principal: 'AUTH_HASH_PRINCIPAL',
    leading_teacher: 'AUTH_HASH_LEADING_TEACHER',
    teacher_pin: 'AUTH_HASH_TEACHER_PIN'
  };
  var settingKey = keyMap[roleKey];
  if (!settingKey) return '';
  var fromSheet = String(rawSettings[settingKey] || '').trim();
  if (fromSheet) return fromSheet;
  var fromProps = props ? String(props.getProperty(settingKey) || '').trim() : '';
  return fromProps;
}

function actionVerifyPasscode_(ss, payload) {
  var role = String(payload.role || '').trim().toLowerCase();
  var passcode = String(payload.passcode || '');
  if (role !== 'admin' && role !== 'principal' && role !== 'leading_teacher') {
    throw appError_('Invalid role specified for passcode verification.', 'INVALID_ROLE');
  }
  if (!passcode) {
    throw appError_('Passcode is required.', 'MISSING_CREDENTIALS');
  }

  var storedHash = getStoredCredentialHash_(ss, role);
  if (!storedHash) {
    throw appError_('Passcode for role "' + role + '" has not been provisioned yet. Run setupInitialCredentials in Apps Script or configure via Admin Security.', 'CREDENTIAL_NOT_CONFIGURED');
  }

  if (!verifyCredentialHash_(passcode, storedHash)) {
    throw appError_('Invalid passcode. Please check your credentials and try again.', 'INVALID_CREDENTIALS');
  }

  var settings = actionGetSettings_(ss).data;
  var extraClaims = {};

  if (role === 'leading_teacher') {
    var ltId = String(payload.leadingTeacherId || '').trim();
    if (!ltId) {
      throw appError_('Please select your Leading Teacher profile to sign in.', 'MISSING_LEADING_TEACHER_ID');
    }
    var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);
    var ltMatch = null;
    for (var i = 0; i < ltTable.records.length; i++) {
      var rec = ltTable.records[i];
      if (String(rec.LeadingTeacherID).trim() === ltId) {
        ltMatch = rec;
        break;
      }
    }
    if (!ltMatch || String(ltMatch.Status || '').trim().toLowerCase() !== 'active') {
      throw appError_('Selected Leading Teacher profile is not active.', 'ACCOUNT_INACTIVE');
    }
    extraClaims.leadingTeacherId = String(ltMatch.LeadingTeacherID).trim();
    extraClaims.leadingTeacherName = String(ltMatch.LeadingTeacherName).trim();
  }

  var sessionInfo = issueRoleToken_(ss, role, extraClaims, settings.sessionTimeoutMinutes);
  if (role === 'leading_teacher') {
    sessionInfo.assignedClassIds = getActiveAssignedClassIdsForLeadingTeacher_(ss, extraClaims.leadingTeacherId);
  }
  return {
    data: sessionInfo,
    message: 'Authenticated successfully'
  };
}

function actionVerifyTeacherPasscode_(ss, payload) {
  var teacherId = String(payload.teacherId || '').trim();
  var pin = String(payload.pin || payload.passcode || '');
  if (!teacherId) {
    throw appError_('Please select an active teacher profile.', 'MISSING_TEACHER_ID');
  }
  if (!pin) {
    throw appError_('Teacher App PIN is required.', 'MISSING_CREDENTIALS');
  }

  var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
  var matchedTeacher = null;
  for (var i = 0; i < teachersTable.records.length; i++) {
    var t = teachersTable.records[i];
    if (String(t.TeacherID).trim() === teacherId) {
      matchedTeacher = t;
      break;
    }
  }
  if (!matchedTeacher || String(matchedTeacher.Status || '').trim().toLowerCase() !== 'active') {
    throw appError_('Selected teacher profile was not found or is inactive.', 'INVALID_TEACHER');
  }

  var storedPinHash = getStoredCredentialHash_(ss, 'teacher_pin');
  if (!storedPinHash) {
    throw appError_('Teacher App PIN has not been configured by the Administrator yet.', 'CREDENTIAL_NOT_CONFIGURED');
  }

  if (!verifyCredentialHash_(pin, storedPinHash)) {
    throw appError_('Invalid Teacher App PIN.', 'INVALID_CREDENTIALS');
  }

  var settings = actionGetSettings_(ss).data;
  var tokenData = issueTeacherToken_(
    ss,
    matchedTeacher.TeacherID,
    matchedTeacher.TeacherName,
    settings.teacherSessionTimeoutMinutes
  );

  return {
    data: tokenData,
    message: 'Teacher authenticated successfully'
  };
}

function actionVerifyTeacherToken_(ss, payload) {
  var token = String(payload.token || payload.authToken || '').trim();
  var verified = verifyTeacherTokenInternal_(ss, token);
  return {
    data: {
      valid: true,
      role: 'teacher',
      teacherId: verified.teacherId,
      teacherName: verified.teacherName,
      expiresAt: new Date(verified.claims.exp).toISOString()
    },
    message: 'Teacher token is valid'
  };
}

function actionChangePasscode_(ss, payload) {
  var auth = requireAuth_(ss, payload, ['admin']);
  if (auth.role !== 'admin') {
    throw appError_('Only an Administrator can change passcodes or PINs.', 'FORBIDDEN');
  }

  var currentAdminPasscode = String(payload.currentAdminPasscode || '');
  if (!currentAdminPasscode) {
    throw appError_('Current Admin passcode is required to authorize credential changes.', 'ADMIN_REAUTH_REQUIRED');
  }
  var adminHash = getStoredCredentialHash_(ss, 'admin');
  if (!adminHash || !verifyCredentialHash_(currentAdminPasscode, adminHash)) {
    throw appError_('Current Admin passcode is incorrect.', 'INVALID_ADMIN_PASSCODE');
  }

  var targetRole = String(payload.targetRole || '').trim().toLowerCase();
  var newPasscode = String(payload.newPasscode || '');
  var validTargets = {
    admin: 'AUTH_HASH_ADMIN',
    principal: 'AUTH_HASH_PRINCIPAL',
    leading_teacher: 'AUTH_HASH_LEADING_TEACHER',
    teacher_pin: 'AUTH_HASH_TEACHER_PIN'
  };
  var settingKey = validTargets[targetRole];
  if (!settingKey) {
    throw appError_('Invalid target credential type. Choose admin, principal, leading_teacher, or teacher_pin.', 'INVALID_TARGET_ROLE');
  }

  if (targetRole === 'teacher_pin') {
    if (newPasscode.length < 4 || newPasscode.length > 32) {
      throw appError_('Teacher App PIN must be between 4 and 32 characters.', 'WEAK_PASSCODE');
    }
  } else {
    if (newPasscode.length < 6 || newPasscode.length > 64) {
      throw appError_('Role passcode must be between 6 and 64 characters.', 'WEAK_PASSCODE');
    }
  }

  var newHash = hashCredential_(newPasscode);
  var nowStr = nowIso_();
  var updates = {};
  updates[settingKey] = newHash;
  updates['CREDENTIAL_UPDATED_AT_' + targetRole.toUpperCase()] = nowStr;
  upsertSettingsMap_(ss, updates);

  var props = PropertiesService.getScriptProperties();
  if (props) {
    props.setProperty(settingKey, newHash);
  }

  var newAdminToken = null;
  if (targetRole === 'teacher_pin') {
    getTeacherTokenSecret_(ss, true);
  } else {
    getRoleTokenSecret_(targetRole, true);
    if (targetRole === 'admin') {
      var settings = actionGetSettings_(ss).data;
      newAdminToken = issueRoleToken_(ss, 'admin', {}, settings.sessionTimeoutMinutes);
    }
  }

  return {
    data: {
      targetRole: targetRole,
      updatedAt: nowStr,
      sessionsInvalidated: true,
      newAdminSession: newAdminToken
    },
    message: 'Credential updated and existing sessions for ' + targetRole + ' invalidated.'
  };
}

function actionGetSecuritySettings_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var raw = getRawSettingsMap_(ss);
  return {
    data: {
      credentialsConfigured: {
        admin: Boolean(getStoredCredentialHash_(ss, 'admin')),
        principal: Boolean(getStoredCredentialHash_(ss, 'principal')),
        leading_teacher: Boolean(getStoredCredentialHash_(ss, 'leading_teacher')),
        teacher_pin: Boolean(getStoredCredentialHash_(ss, 'teacher_pin'))
      },
      lastUpdated: {
        admin: raw.CREDENTIAL_UPDATED_AT_ADMIN || null,
        principal: raw.CREDENTIAL_UPDATED_AT_PRINCIPAL || null,
        leading_teacher: raw.CREDENTIAL_UPDATED_AT_LEADING_TEACHER || null,
        teacher_pin: raw.CREDENTIAL_UPDATED_AT_TEACHER_PIN || null
      },
      sessionTimeoutMinutes: Number(raw.sessionTimeoutMinutes || 480),
      teacherSessionTimeoutMinutes: Number(raw.teacherSessionTimeoutMinutes || 240)
    },
    message: 'Security settings metadata retrieved'
  };
}

// ============================================================================
// STUDENT MANAGEMENT ACTIONS
// ============================================================================

function actionAddStudent_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var studentId = String(payload.studentId || '').trim();
    var studentName = String(payload.studentName || '').trim();
    var classId = String(payload.classId || '').trim();
    var status = String(payload.status || 'Active').trim();

    if (!studentId) throw appError_('Student ID is required.', 'VALIDATION_ERROR');
    if (!studentName) throw appError_('Student Name is required.', 'VALIDATION_ERROR');
    if (!classId) throw appError_('Class selection is required.', 'VALIDATION_ERROR');
    if (status !== 'Active' && status !== 'Inactive') status = 'Active';

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    var classExists = false;
    for (var c = 0; c < classesTable.records.length; c++) {
      if (String(classesTable.records[c].ClassID).trim() === classId) {
        classExists = true;
        break;
      }
    }
    if (!classExists) {
      throw appError_('Referenced Class ID "' + classId + '" does not exist.', 'INVALID_CLASS');
    }

    var studentsTable = readSheetRecords_(ss, SMT_SHEETS.STUDENTS.name);
    for (var i = 0; i < studentsTable.records.length; i++) {
      if (String(studentsTable.records[i].StudentID).trim().toLowerCase() === studentId.toLowerCase()) {
        throw appError_('A student with ID "' + studentId + '" already exists.', 'DUPLICATE_STUDENT_ID');
      }
    }

    var now = nowIso_();
    var record = {
      StudentID: studentId,
      StudentName: studentName,
      ClassID: classId,
      Status: status,
      CreatedAt: now,
      UpdatedAt: now
    };
    appendSheetRecord_(ss, SMT_SHEETS.STUDENTS.name, record);
    return {
      data: { student: record },
      message: 'Student added successfully'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionUpdateStudent_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var studentId = String(payload.studentId || '').trim();
    if (!studentId) throw appError_('Student ID is required.', 'VALIDATION_ERROR');
    if (payload.newStudentId && String(payload.newStudentId).trim() !== studentId) {
      throw appError_('Changing an existing Student ID is not permitted to preserve historical movement integrity.', 'IMMUTABLE_STUDENT_ID');
    }

    var studentsTable = readSheetRecords_(ss, SMT_SHEETS.STUDENTS.name);
    var target = null;
    for (var i = 0; i < studentsTable.records.length; i++) {
      if (String(studentsTable.records[i].StudentID).trim() === studentId) {
        target = studentsTable.records[i];
        break;
      }
    }
    if (!target) {
      throw appError_('Student "' + studentId + '" was not found.', 'NOT_FOUND');
    }

    var newName = payload.studentName !== undefined ? String(payload.studentName).trim() : String(target.StudentName).trim();
    var newClassId = payload.classId !== undefined ? String(payload.classId).trim() : String(target.ClassID).trim();
    var newStatus = payload.status !== undefined ? String(payload.status).trim() : String(target.Status).trim();

    if (!newName) throw appError_('Student Name cannot be empty.', 'VALIDATION_ERROR');
    if (!newClassId) throw appError_('Class ID cannot be empty.', 'VALIDATION_ERROR');

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    var validClass = false;
    for (var c = 0; c < classesTable.records.length; c++) {
      if (String(classesTable.records[c].ClassID).trim() === newClassId) {
        validClass = true;
        break;
      }
    }
    if (!validClass) {
      throw appError_('Referenced Class ID "' + newClassId + '" does not exist.', 'INVALID_CLASS');
    }

    if (newStatus.toLowerCase() === 'inactive') {
      if (hasActiveOutMovementForStudent_(ss, studentId)) {
        throw appError_('Cannot deactivate student "' + newName + '" while they have an unresolved OUT movement.', 'UNRESOLVED_MOVEMENT');
      }
    }

    var now = nowIso_();
    var updates = {
      StudentName: newName,
      ClassID: newClassId,
      Status: newStatus,
      UpdatedAt: now
    };
    updateSheetRecordByRow_(studentsTable, target._rowNumber, updates);

    return {
      data: {
        student: {
          StudentID: studentId,
          StudentName: newName,
          ClassID: newClassId,
          Status: newStatus,
          CreatedAt: target.CreatedAt,
          UpdatedAt: now
        }
      },
      message: 'Student updated successfully'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionDeactivateStudent_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var studentId = String(payload.studentId || '').trim();
  var targetStatus = String(payload.status || 'Inactive').trim();
  return actionUpdateStudent_(ss, {
    token: payload.token,
    studentId: studentId,
    status: targetStatus
  });
}

function hasActiveOutMovementForStudent_(ss, studentId) {
  var movTable = readSheetRecords_(ss, SMT_SHEETS.MOVEMENT_LOG.name);
  for (var i = 0; i < movTable.records.length; i++) {
    var m = movTable.records[i];
    if (String(m.StudentID).trim() === String(studentId).trim() && String(m.Status || '').trim().toUpperCase() === 'OUT') {
      return true;
    }
  }
  return false;
}

function actionBulkAddStudents_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var rows = payload.students || payload.rows;
    if (!Array.isArray(rows) || rows.length === 0) {
      throw appError_('No student rows provided for bulk import.', 'EMPTY_IMPORT');
    }

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    var classById = {};
    var classByNameLower = {};
    for (var c = 0; c < classesTable.records.length; c++) {
      var cls = classesTable.records[c];
      var cid = String(cls.ClassID).trim();
      var cname = String(cls.ClassName).trim();
      classById[cid.toLowerCase()] = cid;
      classByNameLower[cname.toLowerCase()] = cid;
    }

    var studentsTable = readSheetRecords_(ss, SMT_SHEETS.STUDENTS.name);
    var existingIdsLower = {};
    for (var s = 0; s < studentsTable.records.length; s++) {
      var sid = String(studentsTable.records[s].StudentID).trim().toLowerCase();
      if (sid) existingIdsLower[sid] = true;
    }

    var seenInBatchLower = {};
    var added = [];
    var skipped = [];
    var invalid = [];
    var now = nowIso_();

    for (var i = 0; i < rows.length; i++) {
      var item = rows[i] || {};
      var rawId = String(item.studentId || item.StudentID || item.id || '').trim();
      var rawName = String(item.studentName || item.StudentName || item.name || '').trim();
      var rawClass = String(item.classId || item.ClassID || item.className || item.ClassName || item.class || '').trim();
      var rawStatus = String(item.status || item.Status || 'Active').trim();
      if (rawStatus !== 'Active' && rawStatus !== 'Inactive') rawStatus = 'Active';

      if (!rawId || !rawName || !rawClass) {
        invalid.push({
          rowIndex: i + 1,
          studentId: rawId,
          studentName: rawName,
          classRef: rawClass,
          reason: 'Missing required StudentID, StudentName, or Class reference.'
        });
        continue;
      }

      var resolvedClassId = classById[rawClass.toLowerCase()] || classByNameLower[rawClass.toLowerCase()] || null;
      if (!resolvedClassId) {
        invalid.push({
          rowIndex: i + 1,
          studentId: rawId,
          studentName: rawName,
          classRef: rawClass,
          reason: 'Class "' + rawClass + '" does not match any configured ClassID or ClassName.'
        });
        continue;
      }

      var idLower = rawId.toLowerCase();
      if (existingIdsLower[idLower]) {
        skipped.push({
          rowIndex: i + 1,
          studentId: rawId,
          studentName: rawName,
          classId: resolvedClassId,
          reason: 'StudentID already exists in Google Sheets.'
        });
        continue;
      }

      if (seenInBatchLower[idLower]) {
        invalid.push({
          rowIndex: i + 1,
          studentId: rawId,
          studentName: rawName,
          classRef: rawClass,
          reason: 'Duplicate StudentID within the same import batch.'
        });
        continue;
      }

      seenInBatchLower[idLower] = true;

      var newRecord = {
        StudentID: rawId,
        StudentName: rawName,
        ClassID: resolvedClassId,
        Status: rawStatus,
        CreatedAt: now,
        UpdatedAt: now
      };
      appendSheetRecord_(ss, SMT_SHEETS.STUDENTS.name, newRecord);
      added.push(newRecord);
    }

    return {
      data: {
        addedCount: added.length,
        skippedCount: skipped.length,
        invalidCount: invalid.length,
        added: added,
        skipped: skipped,
        invalid: invalid
      },
      message: 'Bulk import completed: ' + added.length + ' added, ' + skipped.length + ' skipped, ' + invalid.length + ' invalid.'
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// CLASS MANAGEMENT ACTIONS
// ============================================================================

function actionAddClass_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var className = String(payload.className || '').trim();
    var classId = String(payload.classId || '').trim();
    var status = String(payload.status || 'Active').trim();
    if (!className) throw appError_('Class Name is required.', 'VALIDATION_ERROR');
    if (status !== 'Active' && status !== 'Inactive') status = 'Active';

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    if (!classId) {
      classId = 'CLS-' + (classesTable.records.length + 101) + '-' + generateRandomHex_(2).toUpperCase();
    }

    for (var i = 0; i < classesTable.records.length; i++) {
      var existing = classesTable.records[i];
      if (String(existing.ClassID).trim().toLowerCase() === classId.toLowerCase()) {
        throw appError_('Class ID "' + classId + '" already exists.', 'DUPLICATE_CLASS_ID');
      }
      if (String(existing.ClassName).trim().toLowerCase() === className.toLowerCase()) {
        throw appError_('A class named "' + className + '" already exists (case-insensitive check).', 'DUPLICATE_CLASS_NAME');
      }
    }

    var now = nowIso_();
    var rec = {
      ClassID: classId,
      ClassName: className,
      Status: status,
      CreatedAt: now,
      UpdatedAt: now
    };
    appendSheetRecord_(ss, SMT_SHEETS.CLASSES.name, rec);
    return {
      data: { classRecord: rec },
      message: 'Class created successfully'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionUpdateClass_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var classId = String(payload.classId || '').trim();
    if (!classId) throw appError_('Class ID is required.', 'VALIDATION_ERROR');

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    var target = null;
    for (var i = 0; i < classesTable.records.length; i++) {
      if (String(classesTable.records[i].ClassID).trim() === classId) {
        target = classesTable.records[i];
        break;
      }
    }
    if (!target) throw appError_('Class "' + classId + '" not found.', 'NOT_FOUND');

    var newName = payload.className !== undefined ? String(payload.className).trim() : String(target.ClassName).trim();
    var newStatus = payload.status !== undefined ? String(payload.status).trim() : String(target.Status).trim();
    if (!newName) throw appError_('Class Name cannot be empty.', 'VALIDATION_ERROR');

    for (var j = 0; j < classesTable.records.length; j++) {
      var other = classesTable.records[j];
      if (String(other.ClassID).trim() !== classId && String(other.ClassName).trim().toLowerCase() === newName.toLowerCase()) {
        throw appError_('Another class with the name "' + newName + '" already exists.', 'DUPLICATE_CLASS_NAME');
      }
    }

    var studentsTable = readSheetRecords_(ss, SMT_SHEETS.STUDENTS.name);
    var enrolledStudentsCount = 0;
    for (var s = 0; s < studentsTable.records.length; s++) {
      if (String(studentsTable.records[s].ClassID).trim() === classId && String(studentsTable.records[s].Status || '').trim().toLowerCase() === 'active') {
        enrolledStudentsCount++;
      }
    }

    if (newStatus.toLowerCase() === 'inactive') {
      var movTable = readSheetRecords_(ss, SMT_SHEETS.MOVEMENT_LOG.name);
      for (var m = 0; m < movTable.records.length; m++) {
        if (String(movTable.records[m].ClassID).trim() === classId && String(movTable.records[m].Status || '').trim().toUpperCase() === 'OUT') {
          throw appError_('Cannot deactivate class "' + newName + '" while a student from this class is currently OUT.', 'UNRESOLVED_MOVEMENT');
        }
      }
      if (enrolledStudentsCount > 0 && !payload.confirmEnrolledWarning) {
        throw appError_('Class "' + newName + '" has ' + enrolledStudentsCount + ' active enrolled student(s). Pass confirmEnrolledWarning=true to confirm deactivation.', 'CLASS_HAS_ENROLLED_STUDENTS');
      }
    }

    var now = nowIso_();
    updateSheetRecordByRow_(classesTable, target._rowNumber, {
      ClassName: newName,
      Status: newStatus,
      UpdatedAt: now
    });

    return {
      data: {
        classRecord: {
          ClassID: classId,
          ClassName: newName,
          Status: newStatus,
          CreatedAt: target.CreatedAt,
          UpdatedAt: now,
          activeEnrolledCount: enrolledStudentsCount
        }
      },
      message: 'Class updated successfully'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionDeactivateClass_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  return actionUpdateClass_(ss, {
    token: payload.token,
    classId: payload.classId,
    status: payload.status || 'Inactive',
    confirmEnrolledWarning: payload.confirmEnrolledWarning !== undefined ? payload.confirmEnrolledWarning : true
  });
}

// ============================================================================
// TEACHER MANAGEMENT, BULK TEACHER IMPORT & TEACHING ASSIGNMENTS
// ============================================================================

function actionGetTeachers_(ss, payload) {
  requireAuth_(ss, payload, ['admin', 'principal']);
  var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
  var assignTable = readSheetRecords_(ss, SMT_SHEETS.TEACHING_ASSIGNMENTS.name);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);

  var classMap = {};
  for (var c = 0; c < classesTable.records.length; c++) {
    classMap[String(classesTable.records[c].ClassID).trim()] = String(classesTable.records[c].ClassName).trim();
  }

  var assignmentsByTeacher = {};
  for (var a = 0; a < assignTable.records.length; a++) {
    var asg = assignTable.records[a];
    if (String(asg.Status || '').trim().toLowerCase() === 'active') {
      var tid = String(asg.TeacherID).trim();
      var cid = String(asg.ClassID).trim();
      if (!assignmentsByTeacher[tid]) assignmentsByTeacher[tid] = [];
      assignmentsByTeacher[tid].push({
        AssignmentID: String(asg.AssignmentID).trim(),
        ClassID: cid,
        ClassName: classMap[cid] || cid
      });
    }
  }

  var teachers = [];
  for (var i = 0; i < teachersTable.records.length; i++) {
    var t = teachersTable.records[i];
    var tidKey = String(t.TeacherID).trim();
    teachers.push({
      TeacherID: tidKey,
      TeacherName: String(t.TeacherName || '').trim(),
      Status: String(t.Status || 'Active').trim(),
      CreatedAt: String(t.CreatedAt || ''),
      UpdatedAt: String(t.UpdatedAt || ''),
      assignedClasses: assignmentsByTeacher[tidKey] || []
    });
  }

  teachers.sort(function (a, b) {
    return a.TeacherName.localeCompare(b.TeacherName);
  });

  return {
    data: { teachers: teachers },
    message: 'Teachers retrieved successfully'
  };
}

function actionAddTeacher_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var teacherName = String(payload.teacherName || '').trim();
    var teacherId = String(payload.teacherId || '').trim();
    var status = String(payload.status || 'Active').trim();
    if (!teacherName) throw appError_('Teacher Name is required.', 'VALIDATION_ERROR');
    if (status !== 'Active' && status !== 'Inactive') status = 'Active';

    var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
    if (!teacherId) {
      teacherId = 'TCH-' + (teachersTable.records.length + 101) + '-' + generateRandomHex_(2).toUpperCase();
    }

    for (var i = 0; i < teachersTable.records.length; i++) {
      if (String(teachersTable.records[i].TeacherID).trim().toLowerCase() === teacherId.toLowerCase()) {
        throw appError_('Teacher ID "' + teacherId + '" already exists.', 'DUPLICATE_TEACHER_ID');
      }
    }

    var now = nowIso_();
    var rec = {
      TeacherID: teacherId,
      TeacherName: teacherName,
      Status: status,
      CreatedAt: now,
      UpdatedAt: now
    };
    appendSheetRecord_(ss, SMT_SHEETS.TEACHERS.name, rec);
    return {
      data: { teacher: rec },
      message: 'Teacher added successfully'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionUpdateTeacher_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var teacherId = String(payload.teacherId || '').trim();
    if (!teacherId) throw appError_('Teacher ID is required.', 'VALIDATION_ERROR');

    var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
    var target = null;
    for (var i = 0; i < teachersTable.records.length; i++) {
      if (String(teachersTable.records[i].TeacherID).trim() === teacherId) {
        target = teachersTable.records[i];
        break;
      }
    }
    if (!target) throw appError_('Teacher "' + teacherId + '" not found.', 'NOT_FOUND');

    var newName = payload.teacherName !== undefined ? String(payload.teacherName).trim() : String(target.TeacherName).trim();
    var newStatus = payload.status !== undefined ? String(payload.status).trim() : String(target.Status).trim();
    if (!newName) throw appError_('Teacher Name cannot be empty.', 'VALIDATION_ERROR');

    if (newStatus.toLowerCase() === 'inactive') {
      var movTable = readSheetRecords_(ss, SMT_SHEETS.MOVEMENT_LOG.name);
      for (var m = 0; m < movTable.records.length; m++) {
        var mov = movTable.records[m];
        if (String(mov.TeacherID).trim() === teacherId && String(mov.Status || '').trim().toUpperCase() === 'OUT') {
          throw appError_('Cannot deactivate teacher "' + newName + '" while they have an unresolved student OUT movement.', 'UNRESOLVED_MOVEMENT');
        }
      }
    }

    var now = nowIso_();
    updateSheetRecordByRow_(teachersTable, target._rowNumber, {
      TeacherName: newName,
      Status: newStatus,
      UpdatedAt: now
    });

    return {
      data: {
        teacher: {
          TeacherID: teacherId,
          TeacherName: newName,
          Status: newStatus,
          CreatedAt: target.CreatedAt,
          UpdatedAt: now
        }
      },
      message: 'Teacher updated successfully'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionDeactivateTeacher_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  return actionUpdateTeacher_(ss, {
    token: payload.token,
    teacherId: payload.teacherId,
    status: payload.status || 'Inactive'
  });
}

/**
 * Bulk teacher import (for 60+ teachers via CSV or Excel paste/upload).
 * Supports TeacherID (or auto-generated ID), TeacherName, Status, and optional assigned ClassIDs/ClassNames.
 * Never overwrites existing teachers and never touches Movement_Log.
 */
function actionBulkAddTeachers_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var rows = payload.teachers || payload.rows;
    if (!Array.isArray(rows) || rows.length === 0) {
      throw appError_('No teacher rows provided for bulk import.', 'EMPTY_IMPORT');
    }

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    var classById = {};
    var classByNameLower = {};
    for (var c = 0; c < classesTable.records.length; c++) {
      var cls = classesTable.records[c];
      var cid = String(cls.ClassID).trim();
      var cname = String(cls.ClassName).trim();
      if (cid) classById[cid.toLowerCase()] = cid;
      if (cname) classByNameLower[cname.toLowerCase()] = cid;
    }

    var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
    var existingIdsLower = {};
    for (var t = 0; t < teachersTable.records.length; t++) {
      var tid = String(teachersTable.records[t].TeacherID).trim().toLowerCase();
      if (tid) existingIdsLower[tid] = true;
    }

    var assignTable = readSheetRecords_(ss, SMT_SHEETS.TEACHING_ASSIGNMENTS.name);
    var nextAsgCounter = assignTable.records.length + 101;
    var nextTchCounter = teachersTable.records.length + 101;

    var seenInBatchLower = {};
    var added = [];
    var skipped = [];
    var invalid = [];
    var assignmentsCreatedCount = 0;
    var now = nowIso_();

    for (var i = 0; i < rows.length; i++) {
      var item = rows[i] || {};
      var rawId = String(item.teacherId || item.TeacherID || item.id || item.staffId || '').trim();
      var rawName = String(item.teacherName || item.TeacherName || item.name || '').trim();
      var rawStatus = String(item.status || item.Status || 'Active').trim();
      if (rawStatus.toLowerCase() === 'inactive') rawStatus = 'Inactive';
      else rawStatus = 'Active';

      if (!rawName) {
        invalid.push({
          rowIndex: i + 1,
          teacherId: rawId,
          teacherName: rawName,
          reason: 'Missing required TeacherName.'
        });
        continue;
      }

      if (!rawId) {
        do {
          rawId = 'TCH-' + (nextTchCounter++) + '-' + generateRandomHex_(2).toUpperCase();
        } while (existingIdsLower[rawId.toLowerCase()] || seenInBatchLower[rawId.toLowerCase()]);
      }

      var idLower = rawId.toLowerCase();
      if (existingIdsLower[idLower]) {
        skipped.push({
          rowIndex: i + 1,
          teacherId: rawId,
          teacherName: rawName,
          reason: 'TeacherID already exists in Google Sheets.'
        });
        continue;
      }

      if (seenInBatchLower[idLower]) {
        invalid.push({
          rowIndex: i + 1,
          teacherId: rawId,
          teacherName: rawName,
          reason: 'Duplicate TeacherID within the same import batch.'
        });
        continue;
      }

      // Resolve optional class assignments (array or semicolon/pipe/comma separated string)
      var rawClassList = [];
      if (Array.isArray(item.classIds)) {
        rawClassList = item.classIds;
      } else if (item.classes || item.Classes || item.classRef) {
        rawClassList = String(item.classes || item.Classes || item.classRef)
          .split(/[;|,]/)
          .map(function (s) { return s.trim(); })
          .filter(Boolean);
      }

      var resolvedClassIds = [];
      var unknownClass = null;
      var seenCls = {};
      for (var k = 0; k < rawClassList.length; k++) {
        var cref = String(rawClassList[k] || '').trim();
        if (!cref) continue;
        var matchedCid = classById[cref.toLowerCase()] || classByNameLower[cref.toLowerCase()] || null;
        if (!matchedCid) {
          unknownClass = cref;
          break;
        }
        if (!seenCls[matchedCid]) {
          seenCls[matchedCid] = true;
          resolvedClassIds.push(matchedCid);
        }
      }

      if (unknownClass) {
        invalid.push({
          rowIndex: i + 1,
          teacherId: rawId,
          teacherName: rawName,
          reason: 'Optional class "' + unknownClass + '" does not match any configured ClassID or ClassName.'
        });
        continue;
      }

      seenInBatchLower[idLower] = true;

      var newTeacher = {
        TeacherID: rawId,
        TeacherName: rawName,
        Status: rawStatus,
        CreatedAt: now,
        UpdatedAt: now
      };
      appendSheetRecord_(ss, SMT_SHEETS.TEACHERS.name, newTeacher);
      added.push(newTeacher);

      for (var aIdx = 0; aIdx < resolvedClassIds.length; aIdx++) {
        var newAsg = {
          AssignmentID: 'TASG-' + (nextAsgCounter++) + '-' + generateRandomHex_(2).toUpperCase(),
          TeacherID: rawId,
          ClassID: resolvedClassIds[aIdx],
          Status: 'Active',
          CreatedAt: now,
          UpdatedAt: now
        };
        appendSheetRecord_(ss, SMT_SHEETS.TEACHING_ASSIGNMENTS.name, newAsg);
        assignmentsCreatedCount++;
      }
    }

    return {
      data: {
        addedCount: added.length,
        skippedCount: skipped.length,
        invalidCount: invalid.length,
        assignmentsCreatedCount: assignmentsCreatedCount,
        added: added,
        skipped: skipped,
        invalid: invalid
      },
      message: 'Bulk teacher import completed: ' + added.length + ' added, ' + skipped.length + ' skipped, ' + invalid.length + ' invalid.'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionGetTeachingAssignments_(ss, payload) {
  requireAuth_(ss, payload, ['admin', 'principal', 'teacher']);
  var assignTable = readSheetRecords_(ss, SMT_SHEETS.TEACHING_ASSIGNMENTS.name);
  var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);

  var teacherMap = {};
  for (var t = 0; t < teachersTable.records.length; t++) {
    teacherMap[String(teachersTable.records[t].TeacherID).trim()] = String(teachersTable.records[t].TeacherName).trim();
  }
  var classMap = {};
  for (var c = 0; c < classesTable.records.length; c++) {
    classMap[String(classesTable.records[c].ClassID).trim()] = String(classesTable.records[c].ClassName).trim();
  }

  var list = [];
  for (var i = 0; i < assignTable.records.length; i++) {
    var rec = assignTable.records[i];
    var tid = String(rec.TeacherID).trim();
    var cid = String(rec.ClassID).trim();
    list.push({
      AssignmentID: String(rec.AssignmentID).trim(),
      TeacherID: tid,
      TeacherName: teacherMap[tid] || tid,
      ClassID: cid,
      ClassName: classMap[cid] || cid,
      Status: String(rec.Status || 'Active').trim(),
      CreatedAt: String(rec.CreatedAt || ''),
      UpdatedAt: String(rec.UpdatedAt || '')
    });
  }
  return {
    data: { assignments: list },
    message: 'Teaching assignments retrieved'
  };
}

function actionAddTeachingAssignment_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var teacherId = String(payload.teacherId || '').trim();
    var classId = String(payload.classId || '').trim();
    if (!teacherId || !classId) {
      throw appError_('Teacher ID and Class ID are required.', 'VALIDATION_ERROR');
    }

    var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
    var teacherActive = false;
    for (var t = 0; t < teachersTable.records.length; t++) {
      if (String(teachersTable.records[t].TeacherID).trim() === teacherId && String(teachersTable.records[t].Status || '').trim().toLowerCase() === 'active') {
        teacherActive = true;
        break;
      }
    }
    if (!teacherActive) throw appError_('Teacher must exist and be Active.', 'INVALID_TEACHER');

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    var classActive = false;
    for (var c = 0; c < classesTable.records.length; c++) {
      if (String(classesTable.records[c].ClassID).trim() === classId && String(classesTable.records[c].Status || '').trim().toLowerCase() === 'active') {
        classActive = true;
        break;
      }
    }
    if (!classActive) throw appError_('Class must exist and be Active.', 'INVALID_CLASS');

    var assignTable = readSheetRecords_(ss, SMT_SHEETS.TEACHING_ASSIGNMENTS.name);
    for (var i = 0; i < assignTable.records.length; i++) {
      var a = assignTable.records[i];
      if (
        String(a.TeacherID).trim() === teacherId &&
        String(a.ClassID).trim() === classId &&
        String(a.Status || '').trim().toLowerCase() === 'active'
      ) {
        throw appError_('An active teaching assignment already exists for this teacher and class.', 'DUPLICATE_ASSIGNMENT');
      }
    }

    var now = nowIso_();
    var rec = {
      AssignmentID: 'TASG-' + (assignTable.records.length + 101) + '-' + generateRandomHex_(2).toUpperCase(),
      TeacherID: teacherId,
      ClassID: classId,
      Status: 'Active',
      CreatedAt: now,
      UpdatedAt: now
    };
    appendSheetRecord_(ss, SMT_SHEETS.TEACHING_ASSIGNMENTS.name, rec);
    return {
      data: { assignment: rec },
      message: 'Teaching assignment created'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionUpdateTeachingAssignment_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var assignmentId = String(payload.assignmentId || '').trim();
    var newStatus = String(payload.status || 'Inactive').trim();
    if (!assignmentId) throw appError_('Assignment ID is required.', 'VALIDATION_ERROR');

    var assignTable = readSheetRecords_(ss, SMT_SHEETS.TEACHING_ASSIGNMENTS.name);
    var target = null;
    for (var i = 0; i < assignTable.records.length; i++) {
      if (String(assignTable.records[i].AssignmentID).trim() === assignmentId) {
        target = assignTable.records[i];
        break;
      }
    }
    if (!target) throw appError_('Teaching assignment not found.', 'NOT_FOUND');

    var targetTeacherId = payload.teacherId ? String(payload.teacherId).trim() : String(target.TeacherID).trim();
    var targetClassId = payload.classId ? String(payload.classId).trim() : String(target.ClassID).trim();

    if (newStatus.toLowerCase() === 'active') {
      for (var j = 0; j < assignTable.records.length; j++) {
        var other = assignTable.records[j];
        if (
          String(other.AssignmentID).trim() !== assignmentId &&
          String(other.TeacherID).trim() === targetTeacherId &&
          String(other.ClassID).trim() === targetClassId &&
          String(other.Status || '').trim().toLowerCase() === 'active'
        ) {
          throw appError_('Duplicate active teaching assignment for this teacher and class.', 'DUPLICATE_ASSIGNMENT');
        }
      }
    }

    var now = nowIso_();
    updateSheetRecordByRow_(assignTable, target._rowNumber, {
      TeacherID: targetTeacherId,
      ClassID: targetClassId,
      Status: newStatus,
      UpdatedAt: now
    });
    return {
      data: {
        assignment: {
          AssignmentID: assignmentId,
          TeacherID: targetTeacherId,
          ClassID: targetClassId,
          Status: newStatus,
          CreatedAt: target.CreatedAt,
          UpdatedAt: now
        }
      },
      message: 'Teaching assignment updated'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionDeactivateTeachingAssignment_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  return actionUpdateTeachingAssignment_(ss, {
    token: payload.token,
    assignmentId: payload.assignmentId,
    status: payload.status || 'Inactive'
  });
}

// ============================================================================
// LEADING TEACHER MANAGEMENT & CLASS ASSIGNMENTS
// ============================================================================

function actionGetLeadingTeachers_(ss, payload) {
  var auth = requireAuth_(ss, payload, ['admin', 'principal', 'leading_teacher']);
  var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);
  var ltaTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHER_ASSIGNMENTS.name);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);

  var classMap = {};
  for (var c = 0; c < classesTable.records.length; c++) {
    classMap[String(classesTable.records[c].ClassID).trim()] = String(classesTable.records[c].ClassName).trim();
  }

  var assignmentsByLt = {};
  for (var a = 0; a < ltaTable.records.length; a++) {
    var asg = ltaTable.records[a];
    var ltId = String(asg.LeadingTeacherID).trim();
    var cid = String(asg.ClassID).trim();
    if (!assignmentsByLt[ltId]) assignmentsByLt[ltId] = [];
    assignmentsByLt[ltId].push({
      AssignmentID: String(asg.AssignmentID).trim(),
      ClassID: cid,
      ClassName: classMap[cid] || cid,
      Period: String(asg.Period || 'Full Year').trim(),
      Status: String(asg.Status || 'Active').trim()
    });
  }

  var list = [];
  for (var i = 0; i < ltTable.records.length; i++) {
    var lt = ltTable.records[i];
    var idKey = String(lt.LeadingTeacherID).trim();
    if (auth.role === 'leading_teacher' && idKey !== auth.leadingTeacherId) {
      continue;
    }
    list.push({
      LeadingTeacherID: idKey,
      LeadingTeacherName: String(lt.LeadingTeacherName || '').trim(),
      Status: String(lt.Status || 'Active').trim(),
      CreatedAt: String(lt.CreatedAt || ''),
      UpdatedAt: String(lt.UpdatedAt || ''),
      assignments: assignmentsByLt[idKey] || []
    });
  }

  return {
    data: { leadingTeachers: list },
    message: 'Leading Teachers retrieved'
  };
}

function actionAddLeadingTeacher_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var name = String(payload.leadingTeacherName || '').trim();
    var ltId = String(payload.leadingTeacherId || '').trim();
    var status = String(payload.status || 'Active').trim();
    if (!name) throw appError_('Leading Teacher Name is required.', 'VALIDATION_ERROR');
    if (status !== 'Active' && status !== 'Inactive') status = 'Active';

    var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);
    if (!ltId) {
      ltId = 'LT-' + (ltTable.records.length + 101) + '-' + generateRandomHex_(2).toUpperCase();
    }
    for (var i = 0; i < ltTable.records.length; i++) {
      if (String(ltTable.records[i].LeadingTeacherID).trim().toLowerCase() === ltId.toLowerCase()) {
        throw appError_('Leading Teacher ID "' + ltId + '" already exists.', 'DUPLICATE_ID');
      }
    }

    var now = nowIso_();
    var rec = {
      LeadingTeacherID: ltId,
      LeadingTeacherName: name,
      Status: status,
      CreatedAt: now,
      UpdatedAt: now
    };
    appendSheetRecord_(ss, SMT_SHEETS.LEADING_TEACHERS.name, rec);
    return {
      data: { leadingTeacher: rec },
      message: 'Leading Teacher added successfully'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionUpdateLeadingTeacher_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ltId = String(payload.leadingTeacherId || '').trim();
    if (!ltId) throw appError_('Leading Teacher ID is required.', 'VALIDATION_ERROR');

    var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);
    var target = null;
    for (var i = 0; i < ltTable.records.length; i++) {
      if (String(ltTable.records[i].LeadingTeacherID).trim() === ltId) {
        target = ltTable.records[i];
        break;
      }
    }
    if (!target) throw appError_('Leading Teacher not found.', 'NOT_FOUND');

    var newName = payload.leadingTeacherName !== undefined ? String(payload.leadingTeacherName).trim() : String(target.LeadingTeacherName).trim();
    var newStatus = payload.status !== undefined ? String(payload.status).trim() : String(target.Status).trim();
    if (!newName) throw appError_('Leading Teacher Name cannot be empty.', 'VALIDATION_ERROR');

    var now = nowIso_();
    updateSheetRecordByRow_(ltTable, target._rowNumber, {
      LeadingTeacherName: newName,
      Status: newStatus,
      UpdatedAt: now
    });

    return {
      data: {
        leadingTeacher: {
          LeadingTeacherID: ltId,
          LeadingTeacherName: newName,
          Status: newStatus,
          CreatedAt: target.CreatedAt,
          UpdatedAt: now
        }
      },
      message: 'Leading Teacher updated successfully'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionDeactivateLeadingTeacher_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  return actionUpdateLeadingTeacher_(ss, {
    token: payload.token,
    leadingTeacherId: payload.leadingTeacherId,
    status: payload.status || 'Inactive'
  });
}

function actionGetLeadingTeacherAssignments_(ss, payload) {
  var auth = requireAuth_(ss, payload, ['admin', 'principal', 'leading_teacher']);
  var ltaTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHER_ASSIGNMENTS.name);
  var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);

  var ltMap = {};
  for (var l = 0; l < ltTable.records.length; l++) {
    ltMap[String(ltTable.records[l].LeadingTeacherID).trim()] = String(ltTable.records[l].LeadingTeacherName).trim();
  }
  var classMap = {};
  for (var c = 0; c < classesTable.records.length; c++) {
    classMap[String(classesTable.records[c].ClassID).trim()] = String(classesTable.records[c].ClassName).trim();
  }

  var list = [];
  for (var i = 0; i < ltaTable.records.length; i++) {
    var rec = ltaTable.records[i];
    var ltId = String(rec.LeadingTeacherID).trim();
    if (auth.role === 'leading_teacher' && ltId !== auth.leadingTeacherId) {
      continue;
    }
    var cid = String(rec.ClassID).trim();
    list.push({
      AssignmentID: String(rec.AssignmentID).trim(),
      LeadingTeacherID: ltId,
      LeadingTeacherName: ltMap[ltId] || ltId,
      ClassID: cid,
      ClassName: classMap[cid] || cid,
      Period: String(rec.Period || 'Full Year').trim(),
      Status: String(rec.Status || 'Active').trim(),
      CreatedAt: String(rec.CreatedAt || ''),
      UpdatedAt: String(rec.UpdatedAt || '')
    });
  }
  return {
    data: { assignments: list },
    message: 'Leading Teacher assignments retrieved'
  };
}

function actionAddLeadingTeacherAssignment_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ltId = String(payload.leadingTeacherId || '').trim();
    var classId = String(payload.classId || '').trim();
    var period = String(payload.period || 'Term 1').trim();
    if (!ltId || !classId || !period) {
      throw appError_('Leading Teacher ID, Class ID, and Period are required.', 'VALIDATION_ERROR');
    }

    var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);
    var ltActive = false;
    for (var l = 0; l < ltTable.records.length; l++) {
      if (String(ltTable.records[l].LeadingTeacherID).trim() === ltId && String(ltTable.records[l].Status || '').trim().toLowerCase() === 'active') {
        ltActive = true;
        break;
      }
    }
    if (!ltActive) throw appError_('Referenced Leading Teacher must exist and be Active.', 'INVALID_LEADING_TEACHER');

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    var classActive = false;
    for (var c = 0; c < classesTable.records.length; c++) {
      if (String(classesTable.records[c].ClassID).trim() === classId && String(classesTable.records[c].Status || '').trim().toLowerCase() === 'active') {
        classActive = true;
        break;
      }
    }
    if (!classActive) throw appError_('Referenced Class must exist and be Active.', 'INVALID_CLASS');

    var ltaTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHER_ASSIGNMENTS.name);
    for (var i = 0; i < ltaTable.records.length; i++) {
      var existing = ltaTable.records[i];
      if (
        String(existing.ClassID).trim() === classId &&
        String(existing.Period || '').trim().toLowerCase() === period.toLowerCase() &&
        String(existing.Status || '').trim().toLowerCase() === 'active'
      ) {
        throw appError_('An active Leading Teacher assignment already exists for class "' + classId + '" in period "' + period + '".', 'DUPLICATE_ASSIGNMENT');
      }
    }

    var now = nowIso_();
    var rec = {
      AssignmentID: 'LTASG-' + (ltaTable.records.length + 101) + '-' + generateRandomHex_(2).toUpperCase(),
      LeadingTeacherID: ltId,
      ClassID: classId,
      Period: period,
      Status: 'Active',
      CreatedAt: now,
      UpdatedAt: now
    };
    appendSheetRecord_(ss, SMT_SHEETS.LEADING_TEACHER_ASSIGNMENTS.name, rec);
    return {
      data: { assignment: rec },
      message: 'Leading Teacher class assignment created'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionUpdateLeadingTeacherAssignment_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var assignmentId = String(payload.assignmentId || '').trim();
    if (!assignmentId) throw appError_('Assignment ID is required.', 'VALIDATION_ERROR');

    var ltaTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHER_ASSIGNMENTS.name);
    var target = null;
    for (var i = 0; i < ltaTable.records.length; i++) {
      if (String(ltaTable.records[i].AssignmentID).trim() === assignmentId) {
        target = ltaTable.records[i];
        break;
      }
    }
    if (!target) throw appError_('Leading Teacher assignment not found.', 'NOT_FOUND');

    var newLtId = payload.leadingTeacherId ? String(payload.leadingTeacherId).trim() : String(target.LeadingTeacherID).trim();
    var newClassId = payload.classId ? String(payload.classId).trim() : String(target.ClassID).trim();
    var newPeriod = payload.period ? String(payload.period).trim() : String(target.Period).trim();
    var newStatus = payload.status ? String(payload.status).trim() : String(target.Status).trim();

    if (newStatus.toLowerCase() === 'active') {
      for (var j = 0; j < ltaTable.records.length; j++) {
        var other = ltaTable.records[j];
        if (
          String(other.AssignmentID).trim() !== assignmentId &&
          String(other.ClassID).trim() === newClassId &&
          String(other.Period || '').trim().toLowerCase() === newPeriod.toLowerCase() &&
          String(other.Status || '').trim().toLowerCase() === 'active'
        ) {
          throw appError_('Conflict: Another active assignment exists for class "' + newClassId + '" and period "' + newPeriod + '".', 'DUPLICATE_ASSIGNMENT');
        }
      }
    }

    var now = nowIso_();
    updateSheetRecordByRow_(ltaTable, target._rowNumber, {
      LeadingTeacherID: newLtId,
      ClassID: newClassId,
      Period: newPeriod,
      Status: newStatus,
      UpdatedAt: now
    });

    return {
      data: {
        assignment: {
          AssignmentID: assignmentId,
          LeadingTeacherID: newLtId,
          ClassID: newClassId,
          Period: newPeriod,
          Status: newStatus,
          CreatedAt: target.CreatedAt,
          UpdatedAt: now
        }
      },
      message: 'Leading Teacher assignment updated'
    };
  } finally {
    lock.releaseLock();
  }
}

function actionDeactivateLeadingTeacherAssignment_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  return actionUpdateLeadingTeacherAssignment_(ss, {
    token: payload.token,
    assignmentId: payload.assignmentId,
    status: payload.status || 'Inactive'
  });
}

// ============================================================================
// ATOMIC MOVEMENT OPERATIONS (RECORD OUT & RECORD RETURN)
// ============================================================================

function actionRecordOut_(ss, payload) {
  var auth = requireAuth_(ss, payload, ['teacher']);
  var lock = LockService.getScriptLock();
  lock.waitLock(12000);
  try {
    var studentId = String(payload.studentId || '').trim();
    var classId = String(payload.classId || '').trim();
    var destination = String(payload.destination || '').trim();
    var notes = String(payload.notes || '').trim();

    if (!studentId) throw appError_('Student ID is required.', 'VALIDATION_ERROR');
    if (!destination) throw appError_('Destination is required.', 'VALIDATION_ERROR');

    if (payload.teacherId && String(payload.teacherId).trim() !== auth.teacherId) {
      throw appError_('Authenticated teacher identity mismatch. Please sign in with the selected teacher account.', 'IDENTITY_MISMATCH');
    }

    var settings = actionGetSettings_(ss).data;
    var validDest = false;
    for (var d = 0; d < settings.destinations.length; d++) {
      if (settings.destinations[d].toLowerCase() === destination.toLowerCase()) {
        destination = settings.destinations[d];
        validDest = true;
        break;
      }
    }
    if (!validDest) {
      throw appError_('Destination "' + destination + '" is not in the configured destination list.', 'INVALID_DESTINATION');
    }

    var studentsTable = readSheetRecords_(ss, SMT_SHEETS.STUDENTS.name);
    var studentRec = null;
    for (var s = 0; s < studentsTable.records.length; s++) {
      if (String(studentsTable.records[s].StudentID).trim() === studentId) {
        studentRec = studentsTable.records[s];
        break;
      }
    }
    if (!studentRec || String(studentRec.Status || '').trim().toLowerCase() !== 'active') {
      throw appError_('Student "' + studentId + '" was not found or is inactive.', 'INVALID_STUDENT');
    }

    var effectiveClassId = classId || String(studentRec.ClassID).trim();
    if (String(studentRec.ClassID).trim() !== effectiveClassId) {
      throw appError_('Student "' + studentRec.StudentName + '" does not belong to class "' + effectiveClassId + '".', 'CLASS_MISMATCH');
    }

    var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
    var classRec = null;
    for (var c = 0; c < classesTable.records.length; c++) {
      if (String(classesTable.records[c].ClassID).trim() === effectiveClassId) {
        classRec = classesTable.records[c];
        break;
      }
    }
    if (!classRec || String(classRec.Status || '').trim().toLowerCase() !== 'active') {
      throw appError_('Class "' + effectiveClassId + '" is not active.', 'INVALID_CLASS');
    }

    var movTable = readSheetRecords_(ss, SMT_SHEETS.MOVEMENT_LOG.name);
    for (var m = 0; m < movTable.records.length; m++) {
      var existingMov = movTable.records[m];
      if (
        String(existingMov.StudentID).trim() === studentId &&
        String(existingMov.Status || '').trim().toUpperCase() === 'OUT'
      ) {
        throw appError_(
          'Student "' + studentRec.StudentName + '" is already marked OUT (Movement ID: ' + existingMov.MovementID + ').',
          'DUPLICATE_ACTIVE_MOVEMENT'
        );
      }
    }

    var nowMs = new Date().getTime();
    var checkoutIso = new Date(nowMs).toISOString();
    var movementId = generateMovementId_(nowMs);

    var movementRecord = {
      MovementID: movementId,
      StudentID: String(studentRec.StudentID).trim(),
      StudentName: String(studentRec.StudentName).trim(),
      ClassID: effectiveClassId,
      ClassName: String(classRec.ClassName).trim(),
      TeacherID: auth.teacherId,
      TeacherName: auth.teacherName,
      Destination: destination,
      TimeOut: checkoutIso,
      TimeIn: '',
      Status: 'OUT',
      Notes: notes
    };

    appendSheetRecord_(ss, SMT_SHEETS.MOVEMENT_LOG.name, movementRecord);

    return {
      data: {
        movement: {
          MovementID: movementRecord.MovementID,
          StudentID: movementRecord.StudentID,
          StudentName: movementRecord.StudentName,
          ClassID: movementRecord.ClassID,
          ClassName: movementRecord.ClassName,
          TeacherID: movementRecord.TeacherID,
          TeacherName: movementRecord.TeacherName,
          Destination: movementRecord.Destination,
          TimeOut: movementRecord.TimeOut,
          TimeIn: '',
          Status: 'OUT',
          Notes: movementRecord.Notes,
          elapsedSeconds: 0,
          elapsedMinutes: 0,
          formattedDuration: '0m 00s',
          attentionLevel: 'Normal'
        }
      },
      message: movementRecord.StudentName + ' marked OUT to ' + destination
    };
  } finally {
    lock.releaseLock();
  }
}

function actionRecordReturn_(ss, payload) {
  requireAuth_(ss, payload, ['teacher', 'admin']);
  var lock = LockService.getScriptLock();
  lock.waitLock(12000);
  try {
    var movementId = String(payload.movementId || '').trim();
    var studentId = String(payload.studentId || '').trim();

    if (!movementId && !studentId) {
      throw appError_('Movement ID (or Student ID) is required to record a return.', 'VALIDATION_ERROR');
    }

    var movTable = readSheetRecords_(ss, SMT_SHEETS.MOVEMENT_LOG.name);
    var target = null;

    if (movementId) {
      for (var i = 0; i < movTable.records.length; i++) {
        var r = movTable.records[i];
        if (String(r.MovementID).trim() === movementId) {
          target = r;
          break;
        }
      }
      if (!target) {
        throw appError_('Movement record "' + movementId + '" was not found.', 'MOVEMENT_NOT_FOUND');
      }
      if (String(target.Status || '').trim().toUpperCase() !== 'OUT') {
        throw appError_('Movement "' + movementId + '" has already been returned.', 'ALREADY_RETURNED');
      }
      if (studentId && String(target.StudentID).trim() !== studentId) {
        throw appError_(
          'Movement "' + movementId + '" belongs to student "' + target.StudentID + '", not "' + studentId + '". Operation rejected.',
          'WRONG_STUDENT_MOVEMENT'
        );
      }
    } else {
      var matches = [];
      for (var j = 0; j < movTable.records.length; j++) {
        var rec = movTable.records[j];
        if (String(rec.StudentID).trim() === studentId && String(rec.Status || '').trim().toUpperCase() === 'OUT') {
          matches.push(rec);
        }
      }
      if (matches.length === 0) {
        throw appError_('No active OUT movement found for student "' + studentId + '".', 'ACTIVE_MOVEMENT_NOT_FOUND');
      }
      target = matches[matches.length - 1];
    }

    var nowMs = new Date().getTime();
    var returnIso = new Date(nowMs).toISOString();
    var settings = actionGetSettings_(ss).data;
    var metrics = computeMovementMetrics_(
      target.TimeOut,
      returnIso,
      String(target.MovementID).trim(),
      settings.warningMinutes,
      settings.criticalMinutes,
      nowMs
    );

    var updatedNotes = payload.notes !== undefined
      ? String(payload.notes).trim()
      : String(target.Notes || '');

    updateSheetRecordByRow_(movTable, target._rowNumber, {
      TimeIn: returnIso,
      Status: 'RETURNED',
      Notes: updatedNotes
    });

    return {
      data: {
        movement: {
          MovementID: String(target.MovementID).trim(),
          StudentID: String(target.StudentID).trim(),
          StudentName: String(target.StudentName).trim(),
          ClassID: String(target.ClassID).trim(),
          ClassName: String(target.ClassName).trim(),
          TeacherID: String(target.TeacherID).trim(),
          TeacherName: String(target.TeacherName).trim(),
          Destination: String(target.Destination).trim(),
          TimeOut: metrics.checkoutIso,
          TimeIn: returnIso,
          Status: 'RETURNED',
          Notes: updatedNotes,
          elapsedSeconds: metrics.elapsedSeconds,
          elapsedMinutes: metrics.elapsedMinutes,
          formattedDuration: metrics.formattedDuration,
          attentionLevel: metrics.attentionLevel
        }
      },
      message: target.StudentName + ' marked RETURNED (' + metrics.formattedDuration + ')'
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// MOVEMENT HISTORY, ANALYTICS & REPORTS
// ============================================================================

function resolveDateFilterBounds_(payload) {
  var rangeType = String(payload.dateRange || payload.period || 'all').trim().toLowerCase();
  var now = new Date();
  var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
  var endOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);

  if (rangeType === 'today') {
    return { startMs: startOfToday.getTime(), endMs: endOfToday.getTime() };
  }
  if (rangeType === 'yesterday') {
    var yStart = new Date(startOfToday.getTime() - 86400000);
    var yEnd = new Date(endOfToday.getTime() - 86400000);
    return { startMs: yStart.getTime(), endMs: yEnd.getTime() };
  }
  if (rangeType === 'this_week') {
    var dayOfWeek = startOfToday.getDay();
    var weekStart = new Date(startOfToday.getTime() - dayOfWeek * 86400000);
    return { startMs: weekStart.getTime(), endMs: endOfToday.getTime() };
  }
  if (rangeType === 'this_month') {
    var monthStart = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
    return { startMs: monthStart.getTime(), endMs: endOfToday.getTime() };
  }
  if (rangeType === 'custom' || payload.startDate || payload.endDate) {
    var sMs = null;
    var eMs = null;
    if (payload.startDate) {
      var sd = new Date(String(payload.startDate).trim() + 'T00:00:00');
      if (!isNaN(sd.getTime())) sMs = sd.getTime();
    }
    if (payload.endDate) {
      var ed = new Date(String(payload.endDate).trim() + 'T23:59:59.999');
      if (!isNaN(ed.getTime())) eMs = ed.getTime();
    }
    return { startMs: sMs, endMs: eMs };
  }
  return { startMs: null, endMs: null };
}

function actionGetMovementHistory_(ss, payload) {
  var auth = requireAuth_(ss, payload, ['admin', 'principal', 'leading_teacher']);
  var settings = actionGetSettings_(ss).data;
  var movTable = readSheetRecords_(ss, SMT_SHEETS.MOVEMENT_LOG.name);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);

  var classMap = {};
  for (var c = 0; c < classesTable.records.length; c++) {
    classMap[String(classesTable.records[c].ClassID).trim()] = String(classesTable.records[c].ClassName).trim();
  }

  var allowedClassSet = null;
  if (auth.role === 'leading_teacher') {
    allowedClassSet = {};
    for (var i = 0; i < auth.assignedClassIds.length; i++) {
      allowedClassSet[auth.assignedClassIds[i]] = true;
    }
  }

  var bounds = resolveDateFilterBounds_(payload);
  var filterClassId = String(payload.classId || '').trim();
  var filterDestination = String(payload.destination || '').trim();
  var filterStatus = String(payload.status || '').trim().toUpperCase();
  var filterStudentId = String(payload.studentId || '').trim();
  var search = String(payload.search || '').trim().toLowerCase();
  var nowMs = new Date().getTime();

  var records = [];
  for (var r = 0; r < movTable.records.length; r++) {
    var row = movTable.records[r];
    var classId = String(row.ClassID || '').trim();
    if (allowedClassSet && !allowedClassSet[classId]) continue;
    if (filterClassId && classId !== filterClassId) continue;

    var status = String(row.Status || '').trim().toUpperCase();
    if (filterStatus && status !== filterStatus) continue;

    var destination = String(row.Destination || '').trim();
    if (filterDestination && destination.toLowerCase() !== filterDestination.toLowerCase()) continue;

    var studentId = String(row.StudentID || '').trim();
    var studentName = String(row.StudentName || '').trim();
    if (filterStudentId && studentId.toLowerCase() !== filterStudentId.toLowerCase()) continue;
    if (search) {
      if (
        studentName.toLowerCase().indexOf(search) === -1 &&
        studentId.toLowerCase().indexOf(search) === -1 &&
        String(row.MovementID || '').toLowerCase().indexOf(search) === -1
      ) {
        continue;
      }
    }

    var movId = String(row.MovementID || '').trim();
    var metrics = computeMovementMetrics_(
      row.TimeOut,
      status === 'RETURNED' ? row.TimeIn : null,
      movId,
      settings.warningMinutes,
      settings.criticalMinutes,
      nowMs
    );

    var checkoutMs = metrics.checkoutIso ? new Date(metrics.checkoutIso).getTime() : null;
    if (bounds.startMs !== null && (!checkoutMs || checkoutMs < bounds.startMs)) continue;
    if (bounds.endMs !== null && (!checkoutMs || checkoutMs > bounds.endMs)) continue;

    records.push({
      MovementID: movId,
      StudentID: studentId,
      StudentName: studentName,
      ClassID: classId,
      ClassName: classMap[classId] || String(row.ClassName || classId).trim(),
      TeacherID: String(row.TeacherID || '').trim(),
      TeacherName: String(row.TeacherName || '').trim(),
      Destination: destination,
      TimeOut: metrics.checkoutIso,
      TimeIn: metrics.returnIso,
      Status: status,
      Notes: String(row.Notes || ''),
      elapsedSeconds: metrics.elapsedSeconds,
      elapsedMinutes: metrics.elapsedMinutes,
      formattedDuration: metrics.formattedDuration,
      attentionLevel: metrics.attentionLevel
    });
  }

  records.sort(function (a, b) {
    return String(b.TimeOut).localeCompare(String(a.TimeOut));
  });

  return {
    data: { records: records },
    message: 'Movement history retrieved'
  };
}

function actionGetMovementReports_(ss, payload) {
  var historyResult = actionGetMovementHistory_(ss, payload);
  var records = historyResult.data.records;

  var totalMovements = records.length;
  var currentOutCount = 0;
  var returnedCount = 0;
  var criticalMovementsCount = 0;
  var warningMovementsCount = 0;
  var sumSeconds = 0;
  var maxSeconds = 0;

  var byDestinationMap = {};
  var byClassMap = {};
  var byStudentMap = {};

  for (var i = 0; i < records.length; i++) {
    var rec = records[i];
    if (rec.Status === 'OUT') {
      currentOutCount++;
    } else if (rec.Status === 'RETURNED') {
      returnedCount++;
    }
    if (rec.attentionLevel === 'Critical') {
      criticalMovementsCount++;
    } else if (rec.attentionLevel === 'Warning') {
      warningMovementsCount++;
    }

    sumSeconds += rec.elapsedSeconds;
    if (rec.elapsedSeconds > maxSeconds) {
      maxSeconds = rec.elapsedSeconds;
    }

    var dest = rec.Destination || 'Unspecified';
    byDestinationMap[dest] = (byDestinationMap[dest] || 0) + 1;

    var clsLabel = rec.ClassName || rec.ClassID || 'Unspecified';
    byClassMap[clsLabel] = (byClassMap[clsLabel] || 0) + 1;

    var sid = rec.StudentID;
    if (!byStudentMap[sid]) {
      byStudentMap[sid] = {
        StudentID: sid,
        StudentName: rec.StudentName,
        ClassID: rec.ClassID,
        ClassName: rec.ClassName,
        movementCount: 0,
        totalDurationMinutes: 0,
        criticalCount: 0
      };
    }
    byStudentMap[sid].movementCount++;
    byStudentMap[sid].totalDurationMinutes = Math.round((byStudentMap[sid].totalDurationMinutes + rec.elapsedMinutes) * 100) / 100;
    if (rec.attentionLevel === 'Critical') {
      byStudentMap[sid].criticalCount++;
    }
  }

  var avgMinutes = totalMovements > 0 ? Math.round(((sumSeconds / totalMovements) / 60) * 100) / 100 : 0;
  var maxMinutes = Math.round((maxSeconds / 60) * 100) / 100;

  var byDestination = Object.keys(byDestinationMap).map(function (k) {
    return { destination: k, count: byDestinationMap[k] };
  }).sort(function (a, b) { return b.count - a.count; });

  var byClass = Object.keys(byClassMap).map(function (k) {
    return { className: k, count: byClassMap[k] };
  }).sort(function (a, b) { return b.count - a.count; });

  var repeatedStudents = Object.keys(byStudentMap).map(function (k) {
    return byStudentMap[k];
  }).filter(function (s) {
    return s.movementCount > 1;
  }).sort(function (a, b) {
    return b.movementCount - a.movementCount;
  });

  return {
    data: {
      summary: {
        totalMovements: totalMovements,
        currentOutCount: currentOutCount,
        returnedCount: returnedCount,
        averageDurationMinutes: avgMinutes,
        maxDurationMinutes: maxMinutes,
        warningMovementsCount: warningMovementsCount,
        criticalMovementsCount: criticalMovementsCount
      },
      byDestination: byDestination,
      byClass: byClass,
      repeatedStudents: repeatedStudents,
      records: records
    },
    message: 'Movement analytics report generated'
  };
}

function actionGetAdminOverview_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var studentsTable = readSheetRecords_(ss, SMT_SHEETS.STUDENTS.name);
  var classesTable = readSheetRecords_(ss, SMT_SHEETS.CLASSES.name);
  var teachersTable = readSheetRecords_(ss, SMT_SHEETS.TEACHERS.name);
  var ltTable = readSheetRecords_(ss, SMT_SHEETS.LEADING_TEACHERS.name);

  var activeStudents = 0;
  for (var s = 0; s < studentsTable.records.length; s++) {
    if (String(studentsTable.records[s].Status || '').trim().toLowerCase() === 'active') activeStudents++;
  }
  var activeClasses = 0;
  for (var c = 0; c < classesTable.records.length; c++) {
    if (String(classesTable.records[c].Status || '').trim().toLowerCase() === 'active') activeClasses++;
  }
  var activeTeachers = 0;
  for (var t = 0; t < teachersTable.records.length; t++) {
    if (String(teachersTable.records[t].Status || '').trim().toLowerCase() === 'active') activeTeachers++;
  }
  var activeLt = 0;
  for (var l = 0; l < ltTable.records.length; l++) {
    if (String(ltTable.records[l].Status || '').trim().toLowerCase() === 'active') activeLt++;
  }

  var currentOutData = actionGetCurrentOut_(ss, payload).data;
  var historyData = actionGetMovementHistory_(ss, { token: payload.token }).data;

  return {
    data: {
      totalStudents: studentsTable.records.length,
      activeStudents: activeStudents,
      totalClasses: classesTable.records.length,
      activeClasses: activeClasses,
      activeTeachers: activeTeachers,
      activeLeadingTeachers: activeLt,
      currentOutCount: currentOutData.stats.totalOut,
      criticalMovementsCount: currentOutData.stats.criticalCount,
      recentActivity: historyData.records.slice(0, 15)
    },
    message: 'Admin overview retrieved'
  };
}

// ============================================================================
// SETTINGS & DESTINATIONS MANAGEMENT
// ============================================================================

function actionUpdateSettings_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var updates = {};
  if (payload.schoolName !== undefined) {
    var name = String(payload.schoolName).trim();
    if (!name) throw appError_('School Name cannot be empty.', 'VALIDATION_ERROR');
    updates.schoolName = name;
  }
  if (payload.schoolTimezone !== undefined) {
    var tz = String(payload.schoolTimezone).trim();
    if (!tz) throw appError_('School Timezone cannot be empty.', 'VALIDATION_ERROR');
    updates.schoolTimezone = tz;
  }
  if (payload.warningMinutes !== undefined) {
    var wm = Number(payload.warningMinutes);
    if (isNaN(wm) || wm < 1 || wm > 120) throw appError_('Warning threshold must be between 1 and 120 minutes.', 'VALIDATION_ERROR');
    updates.warningMinutes = String(wm);
  }
  if (payload.criticalMinutes !== undefined) {
    var cm = Number(payload.criticalMinutes);
    var wmCheck = updates.warningMinutes ? Number(updates.warningMinutes) : Number(getRawSettingsMap_(ss).warningMinutes || 10);
    if (isNaN(cm) || cm <= wmCheck || cm > 240) {
      throw appError_('Critical threshold must be greater than warning threshold (' + wmCheck + 'm).', 'VALIDATION_ERROR');
    }
    updates.criticalMinutes = String(cm);
  }
  if (payload.sessionTimeoutMinutes !== undefined) {
    var stm = Number(payload.sessionTimeoutMinutes);
    if (isNaN(stm) || stm < 5 || stm > 1440) throw appError_('Session timeout must be between 5 and 1440 minutes.', 'VALIDATION_ERROR');
    updates.sessionTimeoutMinutes = String(stm);
  }
  if (payload.teacherSessionTimeoutMinutes !== undefined) {
    var tstm = Number(payload.teacherSessionTimeoutMinutes);
    if (isNaN(tstm) || tstm < 5 || tstm > 1440) throw appError_('Teacher session timeout must be between 5 and 1440 minutes.', 'VALIDATION_ERROR');
    updates.teacherSessionTimeoutMinutes = String(tstm);
  }

  upsertSettingsMap_(ss, updates);
  return actionGetSettings_(ss);
}

function actionManageDestinations_(ss, payload) {
  requireAuth_(ss, payload, ['admin']);
  var settings = actionGetSettings_(ss).data;
  var current = settings.destinations.slice();
  var subAction = String(payload.operation || 'set').trim().toLowerCase();

  if (subAction === 'add') {
    var toAdd = String(payload.destination || '').trim();
    if (!toAdd) throw appError_('Destination name is required.', 'VALIDATION_ERROR');
    for (var i = 0; i < current.length; i++) {
      if (current[i].toLowerCase() === toAdd.toLowerCase()) {
        throw appError_('Destination "' + toAdd + '" already exists.', 'DUPLICATE_DESTINATION');
      }
    }
    current.push(toAdd);
  } else if (subAction === 'remove') {
    var toRemove = String(payload.destination || '').trim();
    current = current.filter(function (d) {
      return d.toLowerCase() !== toRemove.toLowerCase();
    });
    if (current.length === 0) {
      throw appError_('At least one destination must remain configured.', 'VALIDATION_ERROR');
    }
  } else if (subAction === 'set' && Array.isArray(payload.destinations)) {
    var cleaned = [];
    var seen = {};
    for (var j = 0; j < payload.destinations.length; j++) {
      var dName = String(payload.destinations[j] || '').trim();
      if (dName && !seen[dName.toLowerCase()]) {
        seen[dName.toLowerCase()] = true;
        cleaned.push(dName);
      }
    }
    if (cleaned.length === 0) {
      throw appError_('At least one valid destination is required.', 'VALIDATION_ERROR');
    }
    current = cleaned;
  } else {
    throw appError_('Invalid destination management operation.', 'VALIDATION_ERROR');
  }

  upsertSettingsMap_(ss, { destinations: JSON.stringify(current) });
  return {
    data: { destinations: current },
    message: 'Destinations updated successfully'
  };
}

function setupInitialCredentials(config) {
  if (!config || typeof config !== 'object') {
    throw new Error('Provide an object with adminPasscode, principalPasscode, leadingTeacherPasscode, and teacherPin.');
  }
  var ss = getSpreadsheet_();
  ensureWorkbookInitialized_(ss);
  var updates = {};
  var nowStr = nowIso_();
  if (config.adminPasscode) {
    updates.AUTH_HASH_ADMIN = hashCredential_(config.adminPasscode);
    updates.CREDENTIAL_UPDATED_AT_ADMIN = nowStr;
  }
  if (config.principalPasscode) {
    updates.AUTH_HASH_PRINCIPAL = hashCredential_(config.principalPasscode);
    updates.CREDENTIAL_UPDATED_AT_PRINCIPAL = nowStr;
  }
  if (config.leadingTeacherPasscode) {
    updates.AUTH_HASH_LEADING_TEACHER = hashCredential_(config.leadingTeacherPasscode);
    updates.CREDENTIAL_UPDATED_AT_LEADING_TEACHER = nowStr;
  }
  if (config.teacherPin) {
    updates.AUTH_HASH_TEACHER_PIN = hashCredential_(config.teacherPin);
    updates.CREDENTIAL_UPDATED_AT_TEACHER_PIN = nowStr;
  }
  upsertSettingsMap_(ss, updates);
  getTeacherTokenSecret_(ss, false);
  getRoleTokenSecret_('admin', false);
  getRoleTokenSecret_('principal', false);
  getRoleTokenSecret_('leading_teacher', false);
  return { success: true, updatedKeys: Object.keys(updates) };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    SMT_SHEETS: SMT_SHEETS,
    DEFAULT_PUBLIC_SETTINGS: DEFAULT_PUBLIC_SETTINGS,
    doGet: doGet,
    doPost: doPost,
    setupInitialCredentials: setupInitialCredentials,
    hashCredential_: hashCredential_,
    verifyCredentialHash_: verifyCredentialHash_,
    constantTimeEqual_: constantTimeEqual_,
    parseAuthoritativeTimestamp_: parseAuthoritativeTimestamp_,
    computeMovementMetrics_: computeMovementMetrics_,
    generateMovementId_: generateMovementId_,
    extractEpochFromMovementId_: extractEpochFromMovementId_
  };
}
