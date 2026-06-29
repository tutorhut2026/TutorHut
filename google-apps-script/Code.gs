/* ================================================================
   TutorHut – Google Apps Script Backend
   ─────────────────────────────────────
   HOW TO DEPLOY:
   1. Open your existing Apps Script project at script.google.com
      (or create a new one linked to your Google Sheet)
   2. Replace ALL existing code in Code.gs with this file
   3. Click "Deploy" → "Manage deployments" → edit the existing
      deployment → set version to "New version" → Deploy
      (The URL stays the same — no need to update firebase.js)
   4. First time: click "Deploy" → "New deployment" → Web app
      Execute as: Me | Who has access: Anyone → Deploy
      Copy the /exec URL into firebase.js as SHEETS_URL

   FIREBASE AUTH DELETION (optional – for full account removal):
   1. Go to Google Cloud Console → IAM → Service Accounts
   2. Create a service account with "Firebase Admin" role
   3. Download the JSON key
   4. In Apps Script: Project Settings → Script Properties
   5. Add property: FIREBASE_SERVICE_ACCOUNT = <paste full JSON>
   Without this step, Auth accounts are left intact but the tutor
   is blocked from logging in by the login page check.

   GOOGLE DRIVE SETUP (required for tutor folder creation):
   1. Create a folder named "TutorHut Documents" in Google Drive
   2. Open the folder and copy the Folder ID from the URL:
        https://drive.google.com/drive/folders/{FOLDER_ID_HERE}
   3. In Apps Script: Project Settings → Script Properties
   4. Add property: TUTORHUT_DOCUMENTS_ROOT_FOLDER_ID = <Folder ID>
   5. Run driveTestSetup() to verify. Run setupDriveColumns() to
      insert the two new columns into the live Applications sheet.

   VERIFICATION FRAMEWORK SETUP:
   1. Run setupVerificationColumns() to insert 6 status columns
      after driveFolderUrl in the live Applications sheet.
      Prerequisite: setupDriveColumns() must have been run first.
   2. Run migrateVerificationStatuses() to backfill 'Not Submitted'
      for any existing rows that have empty verification columns.

   DOCUMENT SERVICE:
   The Documents sheet is auto-created the first time
   createDocumentRecord() is called. No manual setup required.
   ================================================================ */


// ── Utilities ────────────────────────────────────────────────────

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    var defaultHeaders = {
      Applications: ['id','uid','tutorId','driveFolderId','driveFolderUrl',
                     'identityStatus','rightToWorkStatus','qualificationStatus',
                     'dbsStatus','referencesStatus','safeguardingStatus',
                     'name','email','phone','bio','qualification','field',
                     'experience','subjects','availability','verified','status','submittedAt'],
      Requests:     ['id','studentUid','tutorId','subject','level','message',
                     'availability','status','submittedAt'],
      Documents:    ['documentId','applicationId','tutorId','firebaseUid','category',
                     'driveFileId','driveFolderId','driveUrl',
                     'originalFilename','storedFilename','mimeType','fileSize',
                     'version','status','uploadedAt','uploadedBy',
                     'reviewStatus','reviewedBy','reviewedAt','reviewNotes']
    };
    if (defaultHeaders[name]) sh.appendRow(defaultHeaders[name]);
  }
  return sh;
}

function sheetToObjects(sh) {
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var hdrs = vals[0].map(function(h) { return String(h).trim(); });
  return vals.slice(1)
    .filter(function(row) { return row.some(function(c) { return c !== ''; }); })
    .map(function(row) {
      var obj = {};
      hdrs.forEach(function(h, i) {
        if (h) obj[h] = row[i] !== undefined ? String(row[i]) : '';
      });
      return obj;
    });
}

function findRowIndex(sh, id) {
  var vals = sh.getDataRange().getValues();
  var hdrs = vals[0].map(function(h) { return String(h).trim(); });
  var idCol = hdrs.indexOf('id');
  if (idCol < 0) return -1;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]).trim() === String(id).trim()) return i + 1; // 1-based
  }
  return -1;
}

function setCellValue(sh, rowIdx, fieldName, value) {
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = -1;
  for (var i = 0; i < hdrs.length; i++) {
    if (String(hdrs[i]).trim() === fieldName) { col = i; break; }
  }
  if (col >= 0) sh.getRange(rowIdx, col + 1).setValue(value);
}

/**
 * Finds the row where the 'id' column equals rowId, then writes every
 * key-value pair in updates to its corresponding column.
 * One sheet read, one row scan, N setValue calls, one flush.
 *
 * @param  {Sheet}   sh      Apps Script Sheet object
 * @param  {string}  rowId   Value to match in the 'id' column
 * @param  {Object}  updates { columnName: newValue, … }
 * @return {boolean}         true = row found and updated, false = not found
 */
function updateRowFields(sh, rowId, updates) {
  var vals  = sh.getDataRange().getValues();
  if (vals.length < 2) return false;

  var hdrs  = vals[0].map(function(h) { return String(h).trim(); });
  var idCol = hdrs.indexOf('id');
  if (idCol < 0) return false;

  var rowIdx = -1;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]).trim() === String(rowId).trim()) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx < 0) return false;

  var changed = false;
  Object.keys(updates).forEach(function(field) {
    var col = hdrs.indexOf(field);
    if (col >= 0) {
      sh.getRange(rowIdx, col + 1).setValue(updates[field]);
      changed = true;
    } else {
      Logger.log('[updateRowFields] WARNING: column "' + field +
                 '" not found in sheet "' + sh.getName() + '" — skipped');
    }
  });

  if (changed) SpreadsheetApp.flush();
  return true;
}


// ── DELETE Firebase Auth user (requires FIREBASE_SERVICE_ACCOUNT prop) ──

function b64url(data) {
  return Utilities.base64EncodeWebSafe(data).replace(/=+$/, '');
}

function deleteFirebaseAuthUser(uid) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('FIREBASE_SERVICE_ACCOUNT');
    if (!raw) return null;
    var sa = JSON.parse(raw);

    var now     = Math.floor(Date.now() / 1000);
    var header  = b64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
    var payload = b64url(JSON.stringify({
      iss:   sa.client_email,
      sub:   sa.client_email,
      aud:   'https://oauth2.googleapis.com/token',
      scope: 'https://www.googleapis.com/auth/cloud-platform',
      iat:   now,
      exp:   now + 3600
    }));
    var toSign = header + '.' + payload;
    var sig    = b64url(Utilities.computeRsaSha256Signature(toSign, sa.private_key));
    var jwt    = toSign + '.' + sig;

    var tokenRes = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: 'grant_type=' + encodeURIComponent('urn:ietf:params:oauth:grant-type:jwt-bearer')
             + '&assertion=' + encodeURIComponent(jwt),
      muteHttpExceptions: true
    });
    var tokenData = JSON.parse(tokenRes.getContentText());
    if (!tokenData.access_token) {
      Logger.log('Firebase token error: ' + tokenRes.getContentText());
      return false;
    }

    var delRes = UrlFetchApp.fetch(
      'https://identitytoolkit.googleapis.com/v1/accounts:delete',
      {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + tokenData.access_token },
        payload: JSON.stringify({ localId: uid }),
        muteHttpExceptions: true
      }
    );
    var code = delRes.getResponseCode();
    Logger.log('Firebase Auth delete: ' + code + ' ' + delRes.getContentText());
    return (code === 200 || code === 404);
  } catch (err) {
    Logger.log('deleteFirebaseAuthUser error: ' + err);
    return false;
  }
}


// ── Tutor ID ─────────────────────────────────────────────────────

function isValidTutorId(tutorId) {
  if (!tutorId || typeof tutorId !== 'string') return false;
  return /^TH\d{9}$/.test(tutorId);
}

// Standalone ID generation for testing and manual use only.
// For atomic generate-and-write, call _nextTutorId() inside a held lock (see doPost).
function generateTutorId() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return _nextTutorId();
  } finally {
    lock.releaseLock();
  }
}

// Must always be called from inside a held LockService lock when the result
// will be written back to the sheet.
function _nextTutorId() {
  var year   = new Date().getFullYear();
  var prefix = 'TH' + year;
  var sh     = getSheet('Applications');
  var rows   = sheetToObjects(sh);
  var maxSeq = 0;

  rows.forEach(function(row) {
    var tid = String(row.tutorId || '').trim();
    if (isValidTutorId(tid) && tid.substring(0, 6) === prefix) {
      var seq = parseInt(tid.substring(6), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });

  var nextSeq = maxSeq + 1;
  if (nextSeq > 99999) {
    throw new Error('[TutorId] Sequence exhausted for ' + year +
                    '. Maximum 99,999 tutors per calendar year.');
  }

  var seqStr = String(nextSeq);
  while (seqStr.length < 5) seqStr = '0' + seqStr;
  return prefix + seqStr;
}


// ── Verification Framework ────────────────────────────────────────
//
// VERIFICATION_SCHEMA is the single source of truth for all verification
// metadata. Every verification-related feature reads from this one object:
//
//   Tutor Upload Portal  → driveFolder, allowedExtensions, multipleFiles
//   Admin Portal         → label, required
//   Tutor Dashboard      → label, required
//   Reminder System      → requiresExpiry, expiryReminderDays
//   Document Service     → driveFolder, multipleFiles
//   Sheet layer          → column (read/write column name)
//
// The driveFolder values correspond to the subfolder names that
// DriveHelpers.gs creates inside each tutor's root Drive folder.
//
// State machine (VERIFICATION_TRANSITIONS):
//   Not Submitted → Uploaded → Pending Review → Approved
//                                             → Rejected → Re-upload Required → Uploaded
//   Approved is terminal (maps to itself so re-saves are idempotent).

var VERIFICATION_SCHEMA = {
  identity: {
    column:             'identityStatus',
    label:              'Identity',
    description:        'Passport, national ID, or driving licence',
    driveFolder:        'Identity',
    required:           true,
    multipleFiles:      false,
    allowedExtensions:  ['pdf', 'jpg', 'jpeg', 'png'],
    requiresExpiry:     false,
    expiryReminderDays: null
  },
  rightToWork: {
    column:             'rightToWorkStatus',
    label:              'Right To Work',
    description:        'Visa, work permit, or Biometric Residence Permit',
    driveFolder:        'Right To Work',
    required:           true,
    multipleFiles:      true,
    allowedExtensions:  ['pdf', 'jpg', 'jpeg', 'png'],
    requiresExpiry:     true,
    expiryReminderDays: 30
  },
  qualification: {
    column:             'qualificationStatus',
    label:              'Qualifications',
    description:        'Degree certificates and relevant qualifications',
    driveFolder:        'Qualifications',
    required:           true,
    multipleFiles:      true,
    allowedExtensions:  ['pdf', 'jpg', 'jpeg', 'png'],
    requiresExpiry:     false,
    expiryReminderDays: null
  },
  dbs: {
    column:             'dbsStatus',
    label:              'DBS',
    description:        'Enhanced DBS disclosure certificate',
    driveFolder:        'DBS',
    required:           true,
    multipleFiles:      false,
    allowedExtensions:  ['pdf', 'jpg', 'jpeg', 'png'],
    requiresExpiry:     true,
    expiryReminderDays: 60
  },
  references: {
    column:             'referencesStatus',
    label:              'References',
    description:        'Written professional or academic references',
    driveFolder:        'References',
    required:           true,
    multipleFiles:      true,
    allowedExtensions:  ['pdf', 'jpg', 'jpeg', 'png'],
    requiresExpiry:     false,
    expiryReminderDays: null
  },
  safeguarding: {
    column:             'safeguardingStatus',
    label:              'Safeguarding',
    description:        'Safeguarding training certificate',
    driveFolder:        'Safeguarding',
    required:           true,
    multipleFiles:      false,
    allowedExtensions:  ['pdf', 'jpg', 'jpeg', 'png'],
    requiresExpiry:     true,
    expiryReminderDays: 90
  }
};

var VERIFICATION_STATUSES = [
  'Not Submitted',
  'Uploaded',
  'Pending Review',
  'Approved',
  'Rejected',
  'Re-upload Required'
];

var VERIFICATION_STATUS_DEFAULT = 'Not Submitted';

var VERIFICATION_TRANSITIONS = {
  'Not Submitted':      ['Uploaded'],
  'Uploaded':           ['Pending Review'],
  'Pending Review':     ['Approved', 'Rejected'],
  'Approved':           ['Approved'],
  'Rejected':           ['Re-upload Required'],
  'Re-upload Required': ['Uploaded']
};

function isValidVerificationStatus(status) {
  return VERIFICATION_STATUSES.indexOf(status) !== -1;
}

function isValidVerificationCategory(category) {
  return Object.prototype.hasOwnProperty.call(VERIFICATION_SCHEMA, category);
}

// Returns true if transitioning from → to is permitted by the state machine.
// Returns true if from === to (no-op write, idempotent).
function canTransitionVerificationStatus(from, to) {
  if (from === to) return true;
  var allowed = VERIFICATION_TRANSITIONS[from];
  if (!allowed) return false;
  return allowed.indexOf(to) !== -1;
}

function _verificationColumnName(category) {
  return VERIFICATION_SCHEMA[category].column;
}

// Returns an ordered array of all verification column names, matching sheet order.
function _verificationColumns() {
  return Object.keys(VERIFICATION_SCHEMA).map(function(k) {
    return VERIFICATION_SCHEMA[k].column;
  });
}

/**
 * Returns the current verification status for one category on one application.
 *
 * @param  {string} appId    The application 'id' field value
 * @param  {string} category One of the VERIFICATION_SCHEMA keys
 * @return {string|null}     Current status string, or null if not found
 */
function getVerificationStatus(appId, category) {
  if (!isValidVerificationCategory(category)) {
    Logger.log('[Verification] getVerificationStatus: invalid category "' + category + '"');
    return null;
  }
  var col  = _verificationColumnName(category);
  var rows = sheetToObjects(getSheet('Applications'));
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id).trim() === String(appId).trim()) {
      return String(rows[i][col] || VERIFICATION_STATUS_DEFAULT);
    }
  }
  Logger.log('[Verification] getVerificationStatus: appId "' + appId + '" not found');
  return null;
}

/**
 * Transitions a verification category to a new status, enforcing the state machine.
 * Writes via updateRowFields() and logs the transition.
 *
 * @param  {string} appId     The application 'id' field value
 * @param  {string} category  One of the VERIFICATION_SCHEMA keys
 * @param  {string} newStatus One of the VERIFICATION_STATUSES values
 * @param  {string} changedBy Identifier of the actor (e.g. admin email or 'system')
 * @return {Object}           { ok, error }
 */
function setVerificationStatus(appId, category, newStatus, changedBy) {
  var ts = new Date().toISOString();

  if (!isValidVerificationCategory(category)) {
    var catErr = 'invalid category "' + category + '"';
    Logger.log('[Verification][' + ts + '] setVerificationStatus ERROR ' + catErr);
    return { ok: false, error: catErr };
  }

  if (!isValidVerificationStatus(newStatus)) {
    var statusErr = 'invalid status "' + newStatus + '"';
    Logger.log('[Verification][' + ts + '] setVerificationStatus ERROR ' + statusErr);
    return { ok: false, error: statusErr };
  }

  var currentStatus = getVerificationStatus(appId, category);
  if (currentStatus === null) {
    var notFoundErr = 'application "' + appId + '" not found';
    Logger.log('[Verification][' + ts + '] setVerificationStatus ERROR ' + notFoundErr);
    return { ok: false, error: notFoundErr };
  }

  if (!currentStatus || !isValidVerificationStatus(currentStatus)) {
    currentStatus = VERIFICATION_STATUS_DEFAULT;
  }

  if (!canTransitionVerificationStatus(currentStatus, newStatus)) {
    var transErr = 'transition not allowed: "' + currentStatus + '" → "' + newStatus +
                   '" for category "' + category + '"';
    Logger.log('[Verification][' + ts + '] setVerificationStatus ERROR ' + transErr);
    return { ok: false, error: transErr };
  }

  var update = {};
  update[_verificationColumnName(category)] = newStatus;

  var written = updateRowFields(getSheet('Applications'), appId, update);
  if (!written) {
    var writeErr = 'updateRowFields returned false for appId "' + appId + '"';
    Logger.log('[Verification][' + ts + '] setVerificationStatus ERROR ' + writeErr);
    return { ok: false, error: writeErr };
  }

  Logger.log('[Verification][' + ts + '] STATUS_CHANGE' +
             ' appId='    + appId +
             ' category=' + category +
             ' from="'    + currentStatus + '"' +
             ' to="'      + newStatus + '"' +
             ' changedBy=' + (changedBy || 'unknown'));

  return { ok: true };
}

/**
 * Resets a verification category to 'Not Submitted', bypassing the state machine.
 * Use only for admin corrections.
 *
 * @param  {string} appId    The application 'id' field value
 * @param  {string} category One of the VERIFICATION_SCHEMA keys
 * @param  {string} resetBy  Identifier of the actor performing the reset
 * @return {Object}          { ok, error }
 */
function resetVerificationStatus(appId, category, resetBy) {
  var ts = new Date().toISOString();

  if (!isValidVerificationCategory(category)) {
    var catErr = 'invalid category "' + category + '"';
    Logger.log('[Verification][' + ts + '] resetVerificationStatus ERROR ' + catErr);
    return { ok: false, error: catErr };
  }

  var currentStatus = getVerificationStatus(appId, category);
  if (currentStatus === null) {
    var notFoundErr = 'application "' + appId + '" not found';
    Logger.log('[Verification][' + ts + '] resetVerificationStatus ERROR ' + notFoundErr);
    return { ok: false, error: notFoundErr };
  }

  var update = {};
  update[_verificationColumnName(category)] = VERIFICATION_STATUS_DEFAULT;

  var written = updateRowFields(getSheet('Applications'), appId, update);
  if (!written) {
    var writeErr = 'updateRowFields returned false for appId "' + appId + '"';
    Logger.log('[Verification][' + ts + '] resetVerificationStatus ERROR ' + writeErr);
    return { ok: false, error: writeErr };
  }

  Logger.log('[Verification][' + ts + '] STATUS_RESET' +
             ' appId='     + appId +
             ' category='  + category +
             ' previous="' + currentStatus + '"' +
             ' resetTo="'  + VERIFICATION_STATUS_DEFAULT + '"' +
             ' resetBy='   + (resetBy || 'unknown'));

  return { ok: true };
}


// ── Document Service ──────────────────────────────────────────────
//
// Single source of truth for all tutor document metadata.
// No other layer reads from or writes to the Documents sheet directly.
//
// Document record lifecycle:
//   Active → Archived  (via archiveDocument or replaceDocument)
//   Active → Deleted   (via deleteDocumentRecord — soft delete, Drive file untouched)
//   Archived and Deleted records are never restored or overwritten.
//
// Versioning (multipleFiles = false categories only):
//   replaceDocument() archives the existing Active record, moves the Drive file
//   to the category's _archive subfolder (best effort), then creates a new record
//   with version = previous max version + 1. Version history is never lost.
//
// Drive file handling:
//   createDocumentRecord() — receives already-uploaded Drive metadata from the caller.
//                            The caller is responsible for the actual file upload.
//   archiveDocument()      — metadata only. Caller manages Drive if needed.
//   replaceDocument()      — archives old metadata + best-effort Drive file move.
//   deleteDocumentRecord() — metadata only. Drive file is left in place.
//
// Application relationship:
//   Each document stores applicationId permanently. createDocumentRecord() resolves
//   it once — from params.applicationId if provided, otherwise by scanning the
//   Applications sheet for the tutorId match. setVerificationStatus() is then called
//   with the stored applicationId, eliminating repeated cross-sheet lookups.
//
// Review fields (reserved — not yet implemented):
//   reviewStatus, reviewedBy, reviewedAt, reviewNotes are stored on every record
//   as empty strings. They will be populated by the Admin Review Portal (Prompt 11).
//
// Verification integration:
//   createDocumentRecord() calls setVerificationStatus('Uploaded') after writing
//   the record. This is best-effort — a blocked transition (e.g. already Approved)
//   logs a warning but does not prevent the record from being created.
//
// Document ID format: DOC + 4-digit year + 5-digit zero-padded sequence = 12 chars.
// Examples: DOC202600001, DOC202699999.

var DOCUMENT_STATUS_ACTIVE   = 'Active';
var DOCUMENT_STATUS_ARCHIVED = 'Archived';
var DOCUMENT_STATUS_DELETED  = 'Deleted';
var DOCUMENT_STATUSES = [DOCUMENT_STATUS_ACTIVE, DOCUMENT_STATUS_ARCHIVED, DOCUMENT_STATUS_DELETED];

function isValidDocumentId(docId) {
  if (!docId || typeof docId !== 'string') return false;
  return /^DOC\d{9}$/.test(docId);
}

// Standalone ID generation for testing and manual use only.
// For atomic generate-and-write, use _nextDocumentIdFromRows() inside a held lock.
function generateDocumentId() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return _nextDocumentIdFromRows(sheetToObjects(getSheet('Documents')));
  } finally {
    lock.releaseLock();
  }
}

// Internal: pure ID generation from pre-read rows.
// Avoids a second sheet read when called from inside createDocumentRecord's lock.
// Must always be called from inside a held LockService lock when the result
// will be written back to the sheet.
function _nextDocumentIdFromRows(rows) {
  var year   = new Date().getFullYear();
  var prefix = 'DOC' + year; // 7 chars
  var maxSeq = 0;

  rows.forEach(function(row) {
    var did = String(row.documentId || '').trim();
    if (isValidDocumentId(did) && did.substring(0, 7) === prefix) {
      var seq = parseInt(did.substring(7), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });

  var nextSeq = maxSeq + 1;
  if (nextSeq > 99999) {
    throw new Error('[DocumentId] Sequence exhausted for ' + year +
                    '. Maximum 99,999 documents per calendar year.');
  }

  var seqStr = String(nextSeq);
  while (seqStr.length < 5) seqStr = '0' + seqStr;
  return prefix + seqStr;
}

// Internal: updates a Documents sheet row by its documentId column.
// Mirrors updateRowFields() but uses 'documentId' as the key instead of 'id'.
function _updateDocumentRow(documentId, updates) {
  var sh   = getSheet('Documents');
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return false;

  var hdrs  = vals[0].map(function(h) { return String(h).trim(); });
  var idCol = hdrs.indexOf('documentId');
  if (idCol < 0) return false;

  var rowIdx = -1;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]).trim() === String(documentId).trim()) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx < 0) return false;

  var changed = false;
  Object.keys(updates).forEach(function(field) {
    var col = hdrs.indexOf(field);
    if (col >= 0) {
      sh.getRange(rowIdx, col + 1).setValue(updates[field]);
      changed = true;
    } else {
      Logger.log('[Document] _updateDocumentRow WARNING: column "' + field + '" not found — skipped');
    }
  });

  if (changed) SpreadsheetApp.flush();
  return true;
}

/**
 * Creates a new document record in the Documents sheet and transitions the
 * category's verification status to 'Uploaded' (best effort).
 *
 * The caller is responsible for uploading the file to Drive first and must
 * pass the resulting driveFileId, driveFolderId, and driveUrl in params.
 *
 * @param  {Object} params
 *   Required: tutorId, firebaseUid, category, driveFileId, driveFolderId,
 *             driveUrl, originalFilename, storedFilename, mimeType, fileSize,
 *             uploadedBy
 *   Optional: applicationId (Applications 'id' — derived from tutorId if omitted;
 *                            stored permanently on the record)
 *             appId         (alias for applicationId — accepted for backward compatibility)
 *             uploadedAt    (ISO timestamp — defaults to current time)
 * @return {Object} { ok, documentId, version, error }
 */
function createDocumentRecord(params) {
  var ts = new Date().toISOString();

  // ── Validation ─────────────────────────────────────────────────────────────
  if (!params || typeof params !== 'object') {
    return { ok: false, error: 'params must be an object' };
  }
  if (!isValidTutorId(params.tutorId)) {
    return { ok: false, error: 'invalid tutorId "' + params.tutorId + '"' };
  }
  if (!isValidVerificationCategory(params.category)) {
    return { ok: false, error: 'invalid category "' + params.category + '"' };
  }
  var required = ['firebaseUid','driveFileId','driveFolderId','driveUrl',
                  'originalFilename','storedFilename','mimeType','uploadedBy'];
  for (var ri = 0; ri < required.length; ri++) {
    var f = required[ri];
    if (!params[f] || String(params[f]).trim() === '') {
      return { ok: false, error: 'missing required param "' + f + '"' };
    }
  }

  // ── Resolve applicationId (before lock — reads Applications sheet) ──────────
  // Caller may pass applicationId directly (preferred) or appId (alias). If neither
  // is provided, the Applications sheet is scanned once for the tutorId match and
  // the result is stored permanently on the record. Phase 2 uses this value directly
  // — no further cross-sheet lookup is ever needed.
  var applicationId = params.applicationId || params.appId || null;
  if (!applicationId) {
    var appRows = sheetToObjects(getSheet('Applications'));
    for (var ai = 0; ai < appRows.length; ai++) {
      if (String(appRows[ai].tutorId || '').trim() === String(params.tutorId).trim()) {
        applicationId = appRows[ai].id || null;
        break;
      }
    }
  }
  if (!applicationId) {
    Logger.log('[Document][' + ts + '] WARN applicationId not resolved for tutorId="' +
               params.tutorId + '" — record will be created without applicationId');
  }

  // ── Phase 1: Atomic (lock covers ID generation + version + row write) ──────
  // One Documents sheet read serves both ID generation and version computation.
  var documentId = null;
  var version    = null;
  var lockErr    = null;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (waitErr) {
    Logger.log('[Document][' + ts + '] ERROR createDocumentRecord lock timeout: ' + waitErr.message);
    return { ok: false, error: 'could not acquire lock: ' + waitErr.message };
  }

  try {
    var docSh   = getSheet('Documents');
    var allDocs = sheetToObjects(docSh);

    documentId = _nextDocumentIdFromRows(allDocs);

    // Compute version: max version for this tutorId + category (includes Archived, excludes Deleted).
    var maxVersion = 0;
    allDocs.forEach(function(doc) {
      if (String(doc.tutorId  || '').trim() === String(params.tutorId  || '').trim() &&
          String(doc.category || '').trim() === String(params.category  || '').trim() &&
          String(doc.status   || '').trim() !== DOCUMENT_STATUS_DELETED) {
        var v = parseInt(String(doc.version || '0'), 10);
        if (!isNaN(v) && v > maxVersion) maxVersion = v;
      }
    });
    version = maxVersion + 1;

    var record = {
      documentId:       documentId,
      applicationId:    applicationId ? String(applicationId) : '',
      tutorId:          String(params.tutorId),
      firebaseUid:      String(params.firebaseUid),
      category:         String(params.category),
      driveFileId:      String(params.driveFileId),
      driveFolderId:    String(params.driveFolderId),
      driveUrl:         String(params.driveUrl),
      originalFilename: String(params.originalFilename),
      storedFilename:   String(params.storedFilename),
      mimeType:         String(params.mimeType),
      fileSize:         params.fileSize !== undefined && params.fileSize !== null
                          ? String(params.fileSize) : '',
      version:          String(version),
      status:           DOCUMENT_STATUS_ACTIVE,
      uploadedAt:       params.uploadedAt || ts,
      uploadedBy:       String(params.uploadedBy)
    };

    _appendToSheet(docSh, record);

    Logger.log('[Document][' + ts + '] CREATED' +
               ' documentId=' + documentId +
               ' tutorId='    + params.tutorId +
               ' category='   + params.category +
               ' version='    + version +
               ' uploadedBy=' + params.uploadedBy +
               ' status='     + DOCUMENT_STATUS_ACTIVE);
  } catch (err) {
    lockErr = err.message;
  } finally {
    lock.releaseLock();
  }

  if (lockErr) {
    Logger.log('[Document][' + ts + '] ERROR createDocumentRecord: ' + lockErr);
    return { ok: false, error: lockErr };
  }

  // ── Phase 2: Verification status transition (outside lock, best effort) ─────
  // applicationId was resolved before Phase 1 — no cross-sheet lookup needed here.
  // A blocked transition (e.g. status already Approved) is logged but does not
  // fail the document record — the record is already committed to the sheet.
  try {
    if (applicationId) {
      var vsResult = setVerificationStatus(applicationId, params.category, 'Uploaded',
                                           params.uploadedBy || 'system');
      if (!vsResult.ok) {
        Logger.log('[Document][' + ts + '] WARN verification status not updated for' +
                   ' documentId=' + documentId + ': ' + vsResult.error);
      }
    } else {
      Logger.log('[Document][' + ts + '] WARN applicationId is null for tutorId="' +
                 params.tutorId + '" — verification status not updated');
    }
  } catch (vsErr) {
    Logger.log('[Document][' + ts + '] WARN verification status error: ' + vsErr.message);
  }

  return { ok: true, documentId: documentId, version: version };
}

/**
 * Returns one document record by documentId, or null if not found.
 *
 * @param  {string} documentId
 * @return {Object|null}
 */
function getDocument(documentId) {
  if (!isValidDocumentId(documentId)) return null;
  var rows = sheetToObjects(getSheet('Documents'));
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].documentId || '').trim() === documentId) {
      return rows[i];
    }
  }
  return null;
}

/**
 * Returns all non-Deleted documents for a tutor, in sheet order.
 *
 * @param  {string} tutorId
 * @return {Object[]}
 */
function getTutorDocuments(tutorId) {
  if (!isValidTutorId(tutorId)) {
    Logger.log('[Document] getTutorDocuments: invalid tutorId "' + tutorId + '"');
    return [];
  }
  return sheetToObjects(getSheet('Documents')).filter(function(doc) {
    return String(doc.tutorId || '').trim() === tutorId &&
           String(doc.status  || '').trim() !== DOCUMENT_STATUS_DELETED;
  });
}

/**
 * Returns all non-Deleted documents for a tutor in one verification category.
 *
 * @param  {string} tutorId
 * @param  {string} category  One of the VERIFICATION_SCHEMA keys
 * @return {Object[]}
 */
function getCategoryDocuments(tutorId, category) {
  if (!isValidTutorId(tutorId)) {
    Logger.log('[Document] getCategoryDocuments: invalid tutorId "' + tutorId + '"');
    return [];
  }
  if (!isValidVerificationCategory(category)) {
    Logger.log('[Document] getCategoryDocuments: invalid category "' + category + '"');
    return [];
  }
  return sheetToObjects(getSheet('Documents')).filter(function(doc) {
    return String(doc.tutorId  || '').trim() === tutorId &&
           String(doc.category || '').trim() === category &&
           String(doc.status   || '').trim() !== DOCUMENT_STATUS_DELETED;
  });
}

/**
 * Marks a document record as Archived. Metadata only — does not move Drive files.
 * Idempotent: returns { ok: true } if already Archived.
 * For a full replacement workflow (archive + Drive move + new record), use replaceDocument().
 *
 * @param  {string} documentId
 * @param  {string} archivedBy
 * @return {Object} { ok, error }
 */
function archiveDocument(documentId, archivedBy) {
  var ts = new Date().toISOString();

  if (!isValidDocumentId(documentId)) {
    return { ok: false, error: 'invalid documentId "' + documentId + '"' };
  }

  var doc = getDocument(documentId);
  if (!doc) {
    return { ok: false, error: 'document "' + documentId + '" not found' };
  }
  if (doc.status === DOCUMENT_STATUS_ARCHIVED) {
    Logger.log('[Document][' + ts + '] archiveDocument: "' + documentId + '" already Archived — no-op');
    return { ok: true };
  }
  if (doc.status === DOCUMENT_STATUS_DELETED) {
    return { ok: false, error: 'document "' + documentId + '" is Deleted and cannot be archived' };
  }

  var written = _updateDocumentRow(documentId, { status: DOCUMENT_STATUS_ARCHIVED });
  if (!written) {
    return { ok: false, error: '_updateDocumentRow returned false for "' + documentId + '"' };
  }

  Logger.log('[Document][' + ts + '] ARCHIVED' +
             ' documentId=' + documentId +
             ' tutorId='    + (doc.tutorId   || '') +
             ' category='   + (doc.category  || '') +
             ' version='    + (doc.version   || '') +
             ' archivedBy=' + (archivedBy || 'unknown'));

  return { ok: true };
}

/**
 * Replaces a document for a multipleFiles=false category.
 *
 * Steps:
 *   1. Archives the existing Active document record (metadata).
 *   2. Attempts to move the old Drive file to the category's _archive subfolder
 *      (best effort — a Drive failure logs a warning but does not abort the replace).
 *   3. Creates a new document record from newParams with version = oldVersion + 1.
 *      The new record triggers setVerificationStatus('Uploaded').
 *
 * The caller must upload the new file to Drive first and pass the resulting
 * Drive metadata (driveFileId, driveFolderId, driveUrl) in newParams.
 *
 * @param  {string} oldDocumentId  documentId of the Active document to replace
 * @param  {Object} newParams      Same params as createDocumentRecord()
 * @param  {string} replacedBy     Actor identifier
 * @return {Object} { ok, documentId, version, error }
 */
function replaceDocument(oldDocumentId, newParams, replacedBy) {
  var ts = new Date().toISOString();

  if (!isValidDocumentId(oldDocumentId)) {
    return { ok: false, error: 'invalid oldDocumentId "' + oldDocumentId + '"' };
  }

  var oldDoc = getDocument(oldDocumentId);
  if (!oldDoc) {
    return { ok: false, error: 'document "' + oldDocumentId + '" not found' };
  }
  if (oldDoc.status !== DOCUMENT_STATUS_ACTIVE) {
    return { ok: false, error: 'document "' + oldDocumentId + '" is not Active' +
             ' (current status: "' + oldDoc.status + '"). Only Active documents can be replaced.' };
  }

  var category = oldDoc.category;
  if (!isValidVerificationCategory(category)) {
    return { ok: false, error: 'invalid category "' + category + '" on document record' };
  }
  if (VERIFICATION_SCHEMA[category].multipleFiles) {
    return { ok: false, error: 'replaceDocument() is only valid for multipleFiles=false categories. ' +
             '"' + category + '" (label: "' + VERIFICATION_SCHEMA[category].label + '") allows ' +
             'multiple files — use archiveDocument() and createDocumentRecord() separately.' };
  }

  // ── Step 1: Archive old document record ───────────────────────────────────
  var archResult = archiveDocument(oldDocumentId, replacedBy || 'system');
  if (!archResult.ok) {
    return { ok: false, error: 'failed to archive old document: ' + archResult.error };
  }

  // ── Step 2: Move old Drive file to _archive subfolder (best effort) ────────
  // driveGetSubfolder() throws if the subfolder doesn't exist, so we use
  // driveCheckFolderExists() to check for _archive safely first.
  if (oldDoc.driveFileId) {
    try {
      var tutorFolderResult = driveGetTutorFolder(oldDoc.tutorId);
      if (tutorFolderResult.found) {
        var categoryFolderName = VERIFICATION_SCHEMA[category].driveFolder;
        var categoryFolder     = driveGetSubfolder(tutorFolderResult.folder, categoryFolderName);
        var archiveCheck       = driveCheckFolderExists(categoryFolder, '_archive');

        if (archiveCheck.exists) {
          DriveApp.getFileById(oldDoc.driveFileId).moveTo(archiveCheck.folder);
          Logger.log('[Document][' + ts + '] DRIVE_MOVED old file to _archive' +
                     ' documentId=' + oldDocumentId +
                     ' fileId='     + oldDoc.driveFileId +
                     ' tutorId='    + oldDoc.tutorId +
                     ' category='   + category);
        } else {
          Logger.log('[Document][' + ts + '] WARN _archive folder not found for' +
                     ' category="' + category + '" tutorId=' + oldDoc.tutorId +
                     ' — record archived, Drive file stays in category folder');
        }
      } else {
        Logger.log('[Document][' + ts + '] WARN tutor Drive folder not found for' +
                   ' tutorId=' + oldDoc.tutorId + ' — Drive file not moved');
      }
    } catch (driveErr) {
      Logger.log('[Document][' + ts + '] WARN Drive file move failed' +
                 ' documentId=' + oldDocumentId +
                 ' error=' + driveErr.message +
                 ' — record archived, Drive file not moved');
    }
  }

  // ── Step 3: Create new document record ────────────────────────────────────
  // createDocumentRecord() computes version = max(existing) + 1 and calls
  // setVerificationStatus('Uploaded') automatically.
  var createResult = createDocumentRecord(newParams);
  if (!createResult.ok) {
    Logger.log('[Document][' + ts + '] ERROR replaceDocument: old record archived but' +
               ' new record creation failed: ' + createResult.error);
    return { ok: false, error: 'old document archived but new document creation failed: ' +
             createResult.error };
  }

  Logger.log('[Document][' + ts + '] REPLACED' +
             ' old='        + oldDocumentId +
             ' new='        + createResult.documentId +
             ' version='    + createResult.version +
             ' tutorId='    + (oldDoc.tutorId  || '') +
             ' category='   + category +
             ' replacedBy=' + (replacedBy || 'unknown'));

  return { ok: true, documentId: createResult.documentId, version: createResult.version };
}

/**
 * Soft-deletes a document record (status → 'Deleted').
 * The Drive file is NOT touched. Idempotent: returns { ok: true } if already Deleted.
 *
 * @param  {string} documentId
 * @param  {string} deletedBy
 * @return {Object} { ok, error }
 */
function deleteDocumentRecord(documentId, deletedBy) {
  var ts = new Date().toISOString();

  if (!isValidDocumentId(documentId)) {
    return { ok: false, error: 'invalid documentId "' + documentId + '"' };
  }

  var doc = getDocument(documentId);
  if (!doc) {
    return { ok: false, error: 'document "' + documentId + '" not found' };
  }
  if (doc.status === DOCUMENT_STATUS_DELETED) {
    Logger.log('[Document][' + ts + '] deleteDocumentRecord: "' + documentId + '" already Deleted — no-op');
    return { ok: true };
  }

  var written = _updateDocumentRow(documentId, { status: DOCUMENT_STATUS_DELETED });
  if (!written) {
    return { ok: false, error: '_updateDocumentRow returned false for "' + documentId + '"' };
  }

  Logger.log('[Document][' + ts + '] DELETED (soft)' +
             ' documentId=' + documentId +
             ' tutorId='    + (doc.tutorId  || '') +
             ' category='   + (doc.category || '') +
             ' version='    + (doc.version  || '') +
             ' deletedBy='  + (deletedBy || 'unknown') +
             ' Drive file untouched');

  return { ok: true };
}


// ── Admin Tools ───────────────────────────────────────────────────
// Run these functions MANUALLY from the Apps Script editor.
// They are NOT exposed via doGet or doPost.

// Run ONCE to insert the tutorId column immediately after uid in the live
// Applications sheet. No-op if tutorId already exists.
function setupTutorIdColumn() {
  var sh   = getSheet('Applications');
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
               .map(function(h) { return String(h).trim(); });

  if (hdrs.indexOf('tutorId') !== -1) {
    Logger.log('[TutorId] setupTutorIdColumn: column already exists at position ' +
               (hdrs.indexOf('tutorId') + 1) + '. Nothing to do.');
    return { done: false, reason: 'column already exists' };
  }

  var uidPos = hdrs.indexOf('uid');
  if (uidPos === -1) {
    Logger.log('[TutorId] setupTutorIdColumn: uid column not found. Aborting.');
    return { done: false, reason: 'uid column not found' };
  }

  sh.insertColumnAfter(uidPos + 1);
  sh.getRange(1, uidPos + 2).setValue('tutorId');
  SpreadsheetApp.flush();

  Logger.log('[TutorId] setupTutorIdColumn: inserted tutorId at column ' + (uidPos + 2));
  return { done: true, columnPosition: uidPos + 2 };
}

// Run ONCE to assign Tutor IDs to any existing rows submitted before this
// feature was deployed. Safe to run multiple times — never overwrites a valid ID.
function migrateTutorIds() {
  var startTs = new Date().toISOString();
  Logger.log('[TutorId][' + startTs + '] === Migration starting ===');

  var sh   = getSheet('Applications');
  var data = sh.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log('[TutorId] No data rows found. Nothing to migrate.');
    return { migrated: 0, skipped: 0 };
  }

  var hdrs       = data[0].map(function(h) { return String(h).trim(); });
  var tutorIdCol = hdrs.indexOf('tutorId');
  var uidCol     = hdrs.indexOf('uid');

  if (tutorIdCol === -1) {
    Logger.log('[TutorId] migrateTutorIds: tutorId column not found. ' +
               'Run setupTutorIdColumn() first, then retry.');
    return { migrated: 0, skipped: 0, error: 'tutorId column missing' };
  }

  var migrated = 0;
  var skipped  = 0;

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var isBlank = !row.some(function(c) { return c !== ''; });
    if (isBlank) continue;

    var existing = String(row[tutorIdCol] || '').trim();
    if (existing && isValidTutorId(existing)) {
      skipped++;
      continue;
    }

    var newId = _nextTutorId();
    sh.getRange(i + 1, tutorIdCol + 1).setValue(newId);
    SpreadsheetApp.flush();

    var uid   = uidCol >= 0 ? String(row[uidCol]) : 'unknown';
    var rowTs = new Date().toISOString();
    Logger.log('[TutorId][' + rowTs + '] MIGRATED tutorId=' + newId + ' uid=' + uid);
    migrated++;
  }

  var endTs = new Date().toISOString();
  Logger.log('[TutorId][' + endTs + '] === Migration complete: migrated=' +
             migrated + ' skipped=' + skipped + ' ===');
  return { migrated: migrated, skipped: skipped };
}

// Run ONCE to insert driveFolderId and driveFolderUrl columns immediately after
// tutorId. Requires: tutorId column must already exist.
// Safe to run multiple times — no-op if both columns already exist.
// Insertion order (reverse so anchor tutorId never shifts):
//   Step 1: insert driveFolderUrl at tutorIdPos+1
//   Step 2: insert driveFolderId at tutorIdPos+1  →  driveFolderUrl shifts right
//   Final:  uid | tutorId | driveFolderId | driveFolderUrl | …
function setupDriveColumns() {
  var sh   = getSheet('Applications');
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
               .map(function(h) { return String(h).trim(); });

  var tutorIdPos   = hdrs.indexOf('tutorId');
  var hasFolderId  = hdrs.indexOf('driveFolderId')  !== -1;
  var hasFolderUrl = hdrs.indexOf('driveFolderUrl') !== -1;

  if (tutorIdPos === -1) {
    Logger.log('[Drive] setupDriveColumns: tutorId column not found. ' +
               'Run setupTutorIdColumn() first.');
    return { done: false, reason: 'tutorId column not found' };
  }

  if (hasFolderId && hasFolderUrl) {
    Logger.log('[Drive] setupDriveColumns: both Drive columns already exist. Nothing to do.');
    return { done: false, reason: 'columns already exist' };
  }

  if (!hasFolderUrl) {
    sh.insertColumnAfter(tutorIdPos + 1);
    sh.getRange(1, tutorIdPos + 2).setValue('driveFolderUrl');
    SpreadsheetApp.flush();
    Logger.log('[Drive] setupDriveColumns: inserted driveFolderUrl at column ' + (tutorIdPos + 2));
    hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
              .map(function(h) { return String(h).trim(); });
    tutorIdPos = hdrs.indexOf('tutorId');
  }

  if (!hasFolderId) {
    sh.insertColumnAfter(tutorIdPos + 1);
    sh.getRange(1, tutorIdPos + 2).setValue('driveFolderId');
    SpreadsheetApp.flush();
    Logger.log('[Drive] setupDriveColumns: inserted driveFolderId at column ' + (tutorIdPos + 2));
  }

  Logger.log('[Drive] setupDriveColumns: complete. ' +
             'Final order: uid | tutorId | driveFolderId | driveFolderUrl | name | …');
  return { done: true };
}

// Run ONCE (or after any Drive outage) to create Drive folders for tutor rows
// that have a valid tutorId but an empty driveFolderId.
// Safe to run multiple times — never overwrites an existing driveFolderId.
// Note: uses direct row-index writes (not updateRowFields) to avoid N full
//       sheet re-reads in a loop.
function migrateTutorFolders() {
  var startTs = new Date().toISOString();
  Logger.log('[Drive][' + startTs + '] === migrateTutorFolders starting ===');

  var sh   = getSheet('Applications');
  var data = sh.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log('[Drive] No data rows found. Nothing to migrate.');
    return { migrated: 0, skipped: 0, failed: 0 };
  }

  var hdrs              = data[0].map(function(h) { return String(h).trim(); });
  var tutorIdCol        = hdrs.indexOf('tutorId');
  var driveFolderIdCol  = hdrs.indexOf('driveFolderId');
  var driveFolderUrlCol = hdrs.indexOf('driveFolderUrl');

  if (tutorIdCol === -1) {
    Logger.log('[Drive] migrateTutorFolders: tutorId column not found. Aborting.');
    return { migrated: 0, skipped: 0, failed: 0, error: 'tutorId column missing' };
  }
  if (driveFolderIdCol === -1 || driveFolderUrlCol === -1) {
    Logger.log('[Drive] migrateTutorFolders: Drive columns not found. ' +
               'Run setupDriveColumns() first.');
    return { migrated: 0, skipped: 0, failed: 0, error: 'Drive columns missing' };
  }

  var migrated = 0;
  var skipped  = 0;
  var failed   = 0;

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var isBlank = !row.some(function(c) { return c !== ''; });
    if (isBlank) continue;

    var tutorId          = String(row[tutorIdCol]       || '').trim();
    var existingFolderId = String(row[driveFolderIdCol] || '').trim();

    if (existingFolderId) { skipped++; continue; }

    if (!isValidTutorId(tutorId)) {
      Logger.log('[Drive] migrateTutorFolders: row ' + (i + 1) +
                 ' has no valid tutorId ("' + tutorId + '") — skipping.');
      skipped++;
      continue;
    }

    var ts = new Date().toISOString();
    try {
      var result = driveGetOrCreateTutorFolder(tutorId);
      if (result.success) {
        sh.getRange(i + 1, driveFolderIdCol  + 1).setValue(result.folderId);
        sh.getRange(i + 1, driveFolderUrlCol + 1).setValue(result.folderUrl);
        SpreadsheetApp.flush();
        Logger.log('[Drive][' + ts + '] ' +
                   (result.alreadyExisted ? 'REUSED' : 'CREATED') +
                   ' tutorId=' + tutorId + ' folderId=' + result.folderId);
        migrated++;
      } else {
        Logger.log('[Drive][' + ts + '] FAILED tutorId=' + tutorId + ' error=' + result.error);
        failed++;
      }
    } catch (e) {
      Logger.log('[Drive][' + ts + '] FAILED tutorId=' + tutorId + ' error=' + e.message);
      failed++;
    }
  }

  var endTs = new Date().toISOString();
  Logger.log('[Drive][' + endTs + '] === migrateTutorFolders complete:' +
             ' migrated=' + migrated + ' skipped=' + skipped + ' failed=' + failed + ' ===');
  return { migrated: migrated, skipped: skipped, failed: failed };
}

// Run ONCE to insert 6 verification status columns immediately after driveFolderUrl.
// Prerequisite: setupDriveColumns() must have been run first.
// Safe to run multiple times — skips any column that already exists.
// Insertion uses reverse-order trick: insert last column first, each at urlPos+1,
// so driveFolderUrl stays left of all insertions and its position never shifts.
function setupVerificationColumns() {
  var sh   = getSheet('Applications');
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
               .map(function(h) { return String(h).trim(); });

  var urlPos = hdrs.indexOf('driveFolderUrl');
  if (urlPos === -1) {
    Logger.log('[Verification] setupVerificationColumns: driveFolderUrl column not found. ' +
               'Run setupDriveColumns() first.');
    return { done: false, reason: 'driveFolderUrl column not found' };
  }

  var allPresent = _verificationColumns().every(function(col) {
    return hdrs.indexOf(col) !== -1;
  });
  if (allPresent) {
    Logger.log('[Verification] setupVerificationColumns: all 6 columns already exist. Nothing to do.');
    return { done: false, reason: 'all verification columns already exist' };
  }

  var insertOrder = Object.keys(VERIFICATION_SCHEMA).reverse().map(function(k) {
    return VERIFICATION_SCHEMA[k].column;
  });

  var inserted = [];
  insertOrder.forEach(function(col) {
    hdrs   = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                .map(function(h) { return String(h).trim(); });
    urlPos = hdrs.indexOf('driveFolderUrl');

    if (hdrs.indexOf(col) !== -1) {
      Logger.log('[Verification] setupVerificationColumns: "' + col + '" already exists — skipped');
      return;
    }

    sh.insertColumnAfter(urlPos + 1);
    sh.getRange(1, urlPos + 2).setValue(col);
    SpreadsheetApp.flush();
    Logger.log('[Verification] setupVerificationColumns: inserted "' + col +
               '" at column ' + (urlPos + 2));
    inserted.push(col);
  });

  Logger.log('[Verification] setupVerificationColumns: complete. Inserted: [' +
             inserted.join(', ') + ']');
  return { done: true, inserted: inserted };
}

// Run ONCE (or after setupVerificationColumns) to backfill 'Not Submitted' into
// any existing rows where verification columns are empty.
// Safe to run multiple times — never overwrites a non-empty value.
// Note: uses direct row-index writes (not updateRowFields) to avoid N full
//       sheet re-reads in a loop.
function migrateVerificationStatuses() {
  var startTs = new Date().toISOString();
  Logger.log('[Verification][' + startTs + '] === migrateVerificationStatuses starting ===');

  var sh   = getSheet('Applications');
  var data = sh.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log('[Verification] No data rows found. Nothing to migrate.');
    return { migrated: 0, skipped: 0 };
  }

  var hdrs    = data[0].map(function(h) { return String(h).trim(); });
  var verCols = _verificationColumns();

  var colIndexMap = {};
  var missingCols = [];
  verCols.forEach(function(col) {
    var idx = hdrs.indexOf(col);
    if (idx === -1) { missingCols.push(col); } else { colIndexMap[col] = idx; }
  });

  if (missingCols.length > 0) {
    Logger.log('[Verification] migrateVerificationStatuses: missing columns [' +
               missingCols.join(', ') + ']. Run setupVerificationColumns() first.');
    return { migrated: 0, skipped: 0, error: 'columns missing: ' + missingCols.join(', ') };
  }

  var migrated = 0;
  var skipped  = 0;

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var isBlank = !row.some(function(c) { return c !== ''; });
    if (isBlank) continue;

    var rowMigrated = false;
    verCols.forEach(function(col) {
      var colIdx  = colIndexMap[col];
      var current = String(row[colIdx] || '').trim();
      if (!current) {
        sh.getRange(i + 1, colIdx + 1).setValue(VERIFICATION_STATUS_DEFAULT);
        rowMigrated = true;
      }
    });

    if (rowMigrated) {
      SpreadsheetApp.flush();
      Logger.log('[Verification][' + new Date().toISOString() + '] MIGRATED row ' + (i + 1));
      migrated++;
    } else {
      skipped++;
    }
  }

  var endTs = new Date().toISOString();
  Logger.log('[Verification][' + endTs + '] === migrateVerificationStatuses complete:' +
             ' migrated=' + migrated + ' skipped=' + skipped + ' ===');
  return { migrated: migrated, skipped: skipped };
}


// ── GET Handler ──────────────────────────────────────────────────

function doGet(e) {
  try {
    var p = e.parameter || {};

    switch (p.action) {

      case 'get_applications':
        return jsonOut(sheetToObjects(getSheet('Applications')));

      case 'get_approved_tutors':
        return jsonOut(
          sheetToObjects(getSheet('Applications'))
            .filter(function(r) { return r.status === 'approved'; })
        );

      case 'get_tutor_by_uid': {
        var all   = sheetToObjects(getSheet('Applications'));
        var found = all.filter(function(r) { return r.uid === String(p.uid); });
        return jsonOut(found);
      }

      case 'update_status': {
        var sh  = getSheet(p.sheet || 'Applications');
        var ok  = updateRowFields(sh, p.id, { status: p.status });
        return jsonOut({ ok: ok });
      }

      case 'update_field': {
        var sh2     = getSheet(p.sheet || 'Applications');
        var update2 = {};
        update2[p.field] = p.value;
        var ok2 = updateRowFields(sh2, p.id, update2);
        return jsonOut({ ok: ok2 });
      }

      case 'delete_row': {
        var sh3  = getSheet(p.sheet || 'Applications');
        var row3 = findRowIndex(sh3, p.id);
        if (row3 < 0) {
          return jsonOut({ ok: false, rowDeleted: false, error: 'Row not found', id: p.id });
        }
        sh3.deleteRow(row3);
        SpreadsheetApp.flush();
        var authResult = null;
        if (p.uid) authResult = deleteFirebaseAuthUser(p.uid);
        return jsonOut({ ok: true, rowDeleted: true, authDeleted: authResult === true });
      }

      case 'get_requests':
        return jsonOut(sheetToObjects(getSheet('Requests')));

      case 'get_student_requests':
        return jsonOut(
          sheetToObjects(getSheet('Requests'))
            .filter(function(r) { return r.studentUid === String(p.uid); })
        );

      case 'get_upload_categories': {
        // Returns VERIFICATION_SCHEMA fields required by the upload portal UI,
        // plus the tutor's current status and per-category upload metadata.
        // This is the single authoritative source for client-side schema rendering —
        // the frontend maintains no copy of VERIFICATION_SCHEMA.
        var uid      = String(p.uid || '').trim();
        var statuses = {};
        var docMeta  = {};
        if (uid) {
          // ── Application statuses ──────────────────────────────────────────────
          var ucRows = sheetToObjects(getSheet('Applications'));
          for (var ui = 0; ui < ucRows.length; ui++) {
            if (String(ucRows[ui].uid || '').trim() === uid) {
              Object.keys(VERIFICATION_SCHEMA).forEach(function(k) {
                statuses[k] = ucRows[ui][VERIFICATION_SCHEMA[k].column] ||
                              VERIFICATION_STATUS_DEFAULT;
              });
              break;
            }
          }
          // ── Per-category upload metadata from Documents sheet ─────────────────
          // getSheet auto-creates the sheet if absent; sheetToObjects returns []
          // for an empty sheet, so this is always safe for brand-new accounts.
          var docRows          = sheetToObjects(getSheet('Documents'));
          var activeByCategory = {};
          docRows.forEach(function(row) {
            if (String(row.firebaseUid || '').trim() === uid &&
                String(row.status      || '').trim() === DOCUMENT_STATUS_ACTIVE) {
              var catKey = String(row.category || '').trim();
              if (!activeByCategory[catKey]) activeByCategory[catKey] = [];
              activeByCategory[catKey].push(row);
            }
          });
          Object.keys(VERIFICATION_SCHEMA).forEach(function(k) {
            var docs      = activeByCategory[k] || [];
            var latestAt  = null;
            var latestVer = null;
            docs.forEach(function(doc) {
              // ISO 8601 strings sort lexicographically — no Date parsing needed.
              var at = String(doc.uploadedAt || '').trim();
              if (at && (!latestAt || at > latestAt)) latestAt = at;
              var v = parseInt(String(doc.version || '0'), 10);
              if (!isNaN(v) && v > 0 && (latestVer === null || v > latestVer)) latestVer = v;
            });
            docMeta[k] = {
              activeDocuments: docs.length,
              latestUploadAt:  latestAt,
              latestVersion:   latestVer
            };
          });
        }
        return jsonOut(Object.keys(VERIFICATION_SCHEMA).map(function(k) {
          var s    = VERIFICATION_SCHEMA[k];
          var meta = docMeta[k] || { activeDocuments: 0, latestUploadAt: null, latestVersion: null };
          return {
            key:               k,
            label:             s.label,
            description:       s.description,
            allowedExtensions: s.allowedExtensions,
            multipleFiles:     s.multipleFiles,
            required:          s.required,
            status:            statuses[k] || VERIFICATION_STATUS_DEFAULT,
            activeDocuments:   meta.activeDocuments,
            latestUploadAt:    meta.latestUploadAt,
            latestVersion:     meta.latestVersion
          };
        }));
      }

      default:
        return jsonOut([]);
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return jsonOut({ error: String(err) });
  }
}


// ── POST Handler ─────────────────────────────────────────────────

/**
 * Handles upload_document POST actions.
 *
 * Resolves tutor identity server-side from firebaseUid — never trusts any
 * tutorId or applicationId supplied by the client. Validates the file, uploads
 * it to the correct Drive subfolder, then calls createDocumentRecord() or
 * replaceDocument() depending on the category's multipleFiles setting.
 *
 * Called by doPost() only. Not exposed via doGet() or any HTTP endpoint directly.
 *
 * @param  {Object} data  Parsed POST body (type field already removed by doPost)
 *   Required: firebaseUid, category, base64Content, originalFilename, mimeType
 *   Optional: fileSize (bytes — used for server-side size validation + recording)
 * @return {Object} { ok, documentId, version, category, tutorId } | { ok: false, error }
 */
function _handleDocumentUpload(data) {
  var ts = new Date().toISOString();

  // ── 1. Extract and coerce inputs ────────────────────────────────────────────
  var firebaseUid      = String(data.firebaseUid      || '').trim();
  var category         = String(data.category         || '').trim();
  var base64Content    = String(data.base64Content    || '').trim();
  var originalFilename = String(data.originalFilename || '').trim();
  var mimeType         = String(data.mimeType         || 'application/octet-stream').trim();
  var fileSize         = parseInt(String(data.fileSize || '0'), 10) || 0;

  if (!firebaseUid)      return { ok: false, error: 'firebaseUid is required' };
  if (!base64Content)    return { ok: false, error: 'file content is required' };
  if (!originalFilename) return { ok: false, error: 'filename is required' };

  // ── 2. Validate category ────────────────────────────────────────────────────
  if (!isValidVerificationCategory(category)) {
    return { ok: false, error: 'invalid category "' + category + '"' };
  }
  var schema = VERIFICATION_SCHEMA[category];

  // ── 3. Validate file extension (server-side enforcement) ────────────────────
  var ext = originalFilename.toLowerCase().split('.').pop();
  if (schema.allowedExtensions.indexOf(ext) === -1) {
    return { ok: false, error: 'file type ".' + ext + '" is not allowed for ' +
             schema.label + '. Accepted: ' + schema.allowedExtensions.join(', ') };
  }

  // ── 4. Validate file size: 10 MB hard cap ───────────────────────────────────
  var MAX_BYTES = 10 * 1024 * 1024;
  if (fileSize > MAX_BYTES) {
    return { ok: false, error: 'file is too large (' +
             (Math.round(fileSize / 1024 / 1024 * 10) / 10) + ' MB). Maximum is 10 MB.' };
  }

  // ── 5. Resolve tutor identity server-side (never trust client) ──────────────
  // Scan Applications sheet for the row whose uid matches the authenticated UID.
  var appRows  = sheetToObjects(getSheet('Applications'));
  var tutorRow = null;
  for (var i = 0; i < appRows.length; i++) {
    if (String(appRows[i].uid || '').trim() === firebaseUid) {
      tutorRow = appRows[i];
      break;
    }
  }
  if (!tutorRow) {
    Logger.log('[Upload][' + ts + '] ERROR uid not found in Applications: ' + firebaseUid);
    return { ok: false, error: 'tutor record not found for this account' };
  }

  var tutorId       = String(tutorRow.tutorId || '').trim();
  var applicationId = String(tutorRow.id      || '').trim();

  if (!isValidTutorId(tutorId)) {
    Logger.log('[Upload][' + ts + '] ERROR invalid tutorId "' + tutorId + '" uid=' + firebaseUid);
    return { ok: false, error: 'your tutor ID is not yet assigned — contact support' };
  }
  if (!applicationId) {
    Logger.log('[Upload][' + ts + '] ERROR applicationId missing for uid=' + firebaseUid);
    return { ok: false, error: 'application record is incomplete — contact support' };
  }

  // ── 6. Get or create tutor Drive folder (ensures complete subfolder structure) ─
  var driveResult = null;
  try {
    driveResult = driveGetOrCreateTutorFolder(tutorId);
  } catch (driveSetupErr) {
    Logger.log('[Upload][' + ts + '] ERROR driveGetOrCreateTutorFolder: ' + driveSetupErr.message);
    return { ok: false, error: 'could not access your document storage folder' };
  }
  if (!driveResult.success) {
    Logger.log('[Upload][' + ts + '] ERROR Drive folder failed: ' + driveResult.error);
    return { ok: false, error: 'could not access your document storage folder' };
  }

  // Backfill driveFolderId into Applications if the row was missing it.
  if (!String(tutorRow.driveFolderId || '').trim()) {
    try {
      updateRowFields(getSheet('Applications'), applicationId, {
        driveFolderId:  driveResult.folderId,
        driveFolderUrl: driveResult.folderUrl
      });
    } catch (fwErr) {
      Logger.log('[Upload][' + ts + '] WARN could not backfill driveFolderId: ' + fwErr.message);
    }
  }

  // ── 7. Navigate to category subfolder ────────────────────────────────────────
  // driveGetSubfolder throws if the subfolder doesn't exist.
  // driveGetOrCreateTutorFolder above guarantees the structure is in place.
  var categoryFolder = null;
  try {
    var tutorDriveFolder = DriveApp.getFolderById(driveResult.folderId);
    categoryFolder       = driveGetSubfolder(tutorDriveFolder, schema.driveFolder);
  } catch (navErr) {
    Logger.log('[Upload][' + ts + '] ERROR navigating to subfolder "' +
               schema.driveFolder + '": ' + navErr.message);
    return { ok: false, error: 'document storage subfolder not accessible — contact support' };
  }

  // ── 8. Decode base64 and upload file to Drive ────────────────────────────────
  var storedFilename = category + '_' + ts.replace(/[:.]/g, '-') + '.' + ext;
  var driveFileId    = null;
  var driveUrl       = null;
  var driveFolderId  = categoryFolder.getId();

  try {
    var fileBytes = Utilities.base64Decode(base64Content);
    var blob      = Utilities.newBlob(fileBytes, mimeType, storedFilename);
    var driveFile = categoryFolder.createFile(blob);
    driveFileId   = driveFile.getId();
    driveUrl      = driveFile.getUrl();
  } catch (uploadErr) {
    Logger.log('[Upload][' + ts + '] ERROR Drive file upload: ' + uploadErr.message);
    return { ok: false, error: 'file upload to storage failed — please try again' };
  }

  Logger.log('[Upload][' + ts + '] DRIVE_UPLOADED' +
             ' tutorId='    + tutorId +
             ' category='   + category +
             ' fileId='     + driveFileId +
             ' stored='     + storedFilename +
             ' original='   + originalFilename +
             ' size='       + fileSize);

  // ── 9. Create or replace Document record ─────────────────────────────────────
  var docParams = {
    tutorId:          tutorId,
    applicationId:    applicationId,
    firebaseUid:      firebaseUid,
    category:         category,
    driveFileId:      driveFileId,
    driveFolderId:    driveFolderId,
    driveUrl:         driveUrl,
    originalFilename: originalFilename,
    storedFilename:   storedFilename,
    mimeType:         mimeType,
    fileSize:         String(fileSize),
    uploadedBy:       firebaseUid,
    uploadedAt:       ts
  };

  // ── 9. Create or replace Document record ─────────────────────────────────────
  // If record creation fails after Drive upload succeeds, roll back by trashing
  // the Drive file so Drive and the Documents sheet remain in sync.
  var createResult;
  try {
    if (!schema.multipleFiles) {
      // Single-document category: find any existing Active record and replace it.
      // replaceDocument() archives the old record, moves the old Drive file to
      // _archive (best effort), then calls createDocumentRecord() for the new one.
      var existing = getCategoryDocuments(tutorId, category).filter(function(d) {
        return String(d.status || '').trim() === DOCUMENT_STATUS_ACTIVE;
      });
      if (existing.length > 0) {
        createResult = replaceDocument(existing[0].documentId, docParams, firebaseUid);
      } else {
        createResult = createDocumentRecord(docParams);
      }
    } else {
      // Multi-document category: always add a new record (never replaces).
      createResult = createDocumentRecord(docParams);
    }
  } catch (recordErr) {
    Logger.log('[Upload][' + ts + '] RECORD_ERROR category=' + category +
               ' error=' + recordErr.message + ' — rolling back driveFileId=' + driveFileId);
    _rollbackDriveFile(driveFileId, ts);
    return { ok: false, error: 'Upload failed: your file was received but the record could not be saved. The file has been removed — please try again.' };
  }

  if (!createResult || !createResult.ok) {
    Logger.log('[Upload][' + ts + '] RECORD_FAILED category=' + category +
               ' error=' + (createResult && createResult.error) +
               ' — rolling back driveFileId=' + driveFileId);
    _rollbackDriveFile(driveFileId, ts);
    return { ok: false, error: (createResult && createResult.error) ||
             'Upload failed: record could not be saved. The file has been removed — please try again.' };
  }

  Logger.log('[Upload][' + ts + '] SUCCESS' +
             ' tutorId='    + tutorId +
             ' documentId=' + createResult.documentId +
             ' version='    + createResult.version +
             ' category='   + category +
             ' original='   + originalFilename +
             ' size='       + fileSize +
             ' driveFileId=' + driveFileId);

  return {
    ok:         true,
    documentId: createResult.documentId,
    version:    createResult.version,
    category:   category,
    tutorId:    tutorId
  };
}

function _rollbackDriveFile(fileId, ts) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    Logger.log('[Upload][' + ts + '] ROLLBACK_OK driveFileId=' + fileId);
  } catch (e) {
    Logger.log('[Upload][' + ts + '] ROLLBACK_FAILED driveFileId=' + fileId + ' error=' + e.message);
  }
}

function _appendToSheet(sh, data) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(Object.keys(data));
    sh.appendRow(Object.keys(data).map(function(k) {
      return data[k] !== undefined ? String(data[k]) : '';
    }));
  } else {
    var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                 .map(function(h) { return String(h).trim(); });
    Object.keys(data).forEach(function(k) {
      if (k && !hdrs.includes(k)) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(k);
        hdrs.push(k);
      }
    });
    sh.appendRow(hdrs.map(function(h) {
      return data[h] !== undefined ? String(data[h]) : '';
    }));
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var type = data.type;
    delete data.type;

    if (type === 'tutor_application') {

      // ── Phase 1: Atomic ───────────────────────────────────────────────────
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        if (!data.tutorId) {
          var genTs = new Date().toISOString();
          try {
            data.tutorId = _nextTutorId();
            Logger.log('[TutorId][' + genTs + '] GENERATED tutorId=' + data.tutorId +
                       ' uid=' + (data.uid || 'unknown') + ' action=generated');
          } catch (genErr) {
            Logger.log('[TutorId][' + genTs + '] WARN _nextTutorId failed: ' +
                       genErr.message + ' — application saved without tutorId');
          }
        }

        _verificationColumns().forEach(function(col) {
          if (!data[col]) data[col] = VERIFICATION_STATUS_DEFAULT;
        });

        _appendToSheet(getSheet('Applications'), data);
      } finally {
        lock.releaseLock();
      }

      // ── Phase 2: Drive folder ─────────────────────────────────────────────
      if (data.tutorId && isValidTutorId(data.tutorId)) {
        var driveTs = new Date().toISOString();
        try {
          var driveResult = driveGetOrCreateTutorFolder(data.tutorId);
          if (driveResult.success) {
            updateRowFields(getSheet('Applications'), data.id, {
              driveFolderId:  driveResult.folderId,
              driveFolderUrl: driveResult.folderUrl
            });
            Logger.log('[Drive][' + driveTs + '] ' +
                       (driveResult.alreadyExisted ? 'REUSED' : 'CREATED') +
                       ' folder tutorId=' + data.tutorId +
                       ' uid=' + (data.uid || 'unknown') +
                       ' folderId=' + driveResult.folderId);
          } else {
            Logger.log('[Drive][' + driveTs + '] WARN folder creation failed' +
                       ' tutorId=' + data.tutorId + ' error=' + driveResult.error +
                       ' — application saved. Run migrateTutorFolders() to repair.');
          }
        } catch (driveErr) {
          Logger.log('[Drive][' + driveTs + '] WARN unexpected error' +
                     ' tutorId=' + data.tutorId + ' error=' + driveErr.message +
                     ' — application saved. Run migrateTutorFolders() to repair.');
        }
      }

      return ContentService.createTextOutput('ok');
    }

    if (type === 'upload_document') {
      return jsonOut(_handleDocumentUpload(data));
    }

    var sheetName = type === 'student_request' ? 'Requests' : null;
    if (!sheetName) return ContentService.createTextOutput('unknown type');
    _appendToSheet(getSheet(sheetName), data);
    return ContentService.createTextOutput('ok');

  } catch (err) {
    Logger.log('doPost error: ' + err);
    return ContentService.createTextOutput('error: ' + err);
  }
}


// ── Development Reset Utility ─────────────────────────────────────
//
// DEVELOPMENT ONLY.
// These functions are executable ONLY from the Apps Script editor.
// They are NOT referenced anywhere in doGet() or doPost() and cannot
// be triggered from the website or any HTTP request.
//
// HOW TO USE (select from the function dropdown → Run):
//   developmentResetDryRun()    — preview only, zero deletions
//   developmentResetConfirmed() — performs the full reset
//   developmentReset()          — aborts immediately (safe default)
//   developmentReset(true)      — performs the full reset (programmatic)

function developmentReset(confirmed) {
  if (confirmed !== true) {
    var abortTs = new Date().toISOString();
    Logger.log('[DevReset][' + abortTs + '] ABORTED — confirmation not supplied. ' +
               'Run developmentResetConfirmed() or call developmentReset(true).');
    return { aborted: true, reason: 'confirmed !== true' };
  }
  return _performReset(false);
}

function developmentResetConfirmed() {
  return developmentReset(true);
}

function developmentResetDryRun() {
  Logger.log('[DevReset][' + new Date().toISOString() + '] DRY RUN — no data will be deleted.');
  return _performReset(true);
}

function _performReset(isDryRun) {
  var ts   = new Date().toISOString();
  var mode = isDryRun ? '[DRY RUN] ' : '';

  Logger.log('[DevReset][' + ts + '] ' +
             '=== Development Reset ' + (isDryRun ? 'DRY RUN' : 'STARTING') + ' ===');

  var summary = {
    isDryRun:            isDryRun,
    applicationsDeleted: 0,
    requestsDeleted:     0,
    documentsDeleted:    0,
    driveFoldersDeleted: 0,
    driveStatus:         'not attempted',
    firebaseStatus:      'skipped — delete manually via Firebase Console',
    startedAt:           ts,
    completedAt:         null
  };

  // ── 1. Applications sheet ─────────────────────────────────────────────────
  try {
    var appSh       = getSheet('Applications');
    var appLastRow  = appSh.getLastRow();
    var appDataRows = appLastRow > 1 ? appLastRow - 1 : 0;
    if (!isDryRun && appDataRows > 0) {
      appSh.deleteRows(2, appDataRows);
      SpreadsheetApp.flush();
    }
    summary.applicationsDeleted = appDataRows;
    Logger.log('[DevReset][' + new Date().toISOString() + '] ' + mode +
               (isDryRun ? 'WOULD DELETE' : 'DELETED') + ' ' + appDataRows +
               ' row(s) from Applications — header preserved');
  } catch (e) {
    Logger.log('[DevReset] ERROR clearing Applications: ' + e.message);
    summary.applicationsError = e.message;
  }

  // ── 2. Requests sheet ─────────────────────────────────────────────────────
  try {
    var reqSh       = getSheet('Requests');
    var reqLastRow  = reqSh.getLastRow();
    var reqDataRows = reqLastRow > 1 ? reqLastRow - 1 : 0;
    if (!isDryRun && reqDataRows > 0) {
      reqSh.deleteRows(2, reqDataRows);
      SpreadsheetApp.flush();
    }
    summary.requestsDeleted = reqDataRows;
    Logger.log('[DevReset][' + new Date().toISOString() + '] ' + mode +
               (isDryRun ? 'WOULD DELETE' : 'DELETED') + ' ' + reqDataRows +
               ' row(s) from Requests — header preserved');
  } catch (e) {
    Logger.log('[DevReset] ERROR clearing Requests: ' + e.message);
    summary.requestsError = e.message;
  }

  // ── 3. Documents sheet ────────────────────────────────────────────────────
  try {
    var docSh       = getSheet('Documents');
    var docLastRow  = docSh.getLastRow();
    var docDataRows = docLastRow > 1 ? docLastRow - 1 : 0;
    if (!isDryRun && docDataRows > 0) {
      docSh.deleteRows(2, docDataRows);
      SpreadsheetApp.flush();
    }
    summary.documentsDeleted = docDataRows;
    Logger.log('[DevReset][' + new Date().toISOString() + '] ' + mode +
               (isDryRun ? 'WOULD DELETE' : 'DELETED') + ' ' + docDataRows +
               ' row(s) from Documents — header preserved');
  } catch (e) {
    Logger.log('[DevReset] ERROR clearing Documents: ' + e.message);
    summary.documentsError = e.message;
  }

  // ── 4. Google Drive tutor folders ─────────────────────────────────────────
  try {
    var rootFolder = driveGetRootFolder();
    var folderIter = rootFolder.getFolders();
    var driveCount = 0;

    while (folderIter.hasNext()) {
      var folder     = folderIter.next();
      var folderName = folder.getName();

      if (!isValidTutorId(folderName)) {
        Logger.log('[DevReset][' + new Date().toISOString() + '] ' +
                   'SKIPPED non-tutor folder: "' + folderName + '"');
        continue;
      }

      if (!isDryRun) {
        var folderId = folder.getId();
        folder.setTrashed(true);
        Logger.log('[DevReset][' + new Date().toISOString() + '] ' +
                   'TRASHED folder: ' + folderName + ' (id=' + folderId + ')');
      } else {
        Logger.log('[DevReset][' + new Date().toISOString() + '] ' +
                   '[DRY RUN] WOULD TRASH folder: ' + folderName +
                   ' (id=' + folder.getId() + ')');
      }
      driveCount++;
    }

    summary.driveFoldersDeleted = driveCount;
    summary.driveStatus         = 'ok';
    Logger.log('[DevReset][' + new Date().toISOString() + '] ' + mode +
               (isDryRun ? 'WOULD TRASH' : 'TRASHED') + ' ' + driveCount +
               ' tutor folder(s) — root "TutorHut Documents" folder preserved');
  } catch (e) {
    Logger.log('[DevReset] ERROR accessing Google Drive: ' + e.message);
    summary.driveStatus = 'error: ' + e.message;
  }

  // ── 5. Firebase Authentication ────────────────────────────────────────────
  Logger.log('[DevReset][' + new Date().toISOString() + '] ' +
             'SKIPPED Firebase Authentication — delete test users manually: ' +
             'Firebase Console → Authentication → Users → select → Delete.');

  summary.completedAt = new Date().toISOString();
  Logger.log('[DevReset][' + summary.completedAt + '] ' +
             '=== ' + (isDryRun ? 'Dry run' : 'Reset') + ' complete ===');
  Logger.log('[DevReset] Summary: applications=' + summary.applicationsDeleted +
             ' requests='  + summary.requestsDeleted +
             ' documents=' + summary.documentsDeleted +
             ' drive='     + summary.driveFoldersDeleted +
             ' driveStatus=' + summary.driveStatus);

  return summary;
}
