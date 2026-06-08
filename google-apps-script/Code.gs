/* ================================================================
   TutorHut – Google Apps Script Backend
   ─────────────────────────────────────
   HOW TO DEPLOY:
   1. Open your existing Apps Script project at script.google.com
      (or create a new one linked to your Google Sheet)
   2. Replace ALL existing code with this file
   3. Click "Deploy" → "Manage deployments" → edit the existing
      deployment → set version to "New version" → Deploy
      (The URL stays the same — no need to update firebase.js)
   4. First time: click "Deploy" → "New deployment" → Web app
      Execute as: Me | Who has access: Anyone → Deploy
      Copy the /exec URL into firebase.js as SHEETS_URL

   FIREBASE AUTH DELETION (optional – for full account removal):
   1. Go to Google Cloud Console →
    IAM → Service Accounts
   2. Create a service account with "Firebase Admin" role
   3. Download the JSON key
   4. In Apps Script: Project Settings → Script Properties
   5. Add property: FIREBASE_SERVICE_ACCOUNT = <paste full JSON>
   Without this step, Auth accounts are left intact but the tutor
   is blocked from logging in by the login page check.
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
      Applications: ['id','uid','name','email','phone','bio','qualification','field',
                     'experience','subjects','availability','verified','status','submittedAt'],
      Requests:     ['id','studentUid','tutorId','subject','level','message',
                     'availability','status','submittedAt']
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

// ── DELETE Firebase Auth user (requires FIREBASE_SERVICE_ACCOUNT prop) ──

function b64url(data) {
  // base64url with no padding — required for valid JWTs
  return Utilities.base64EncodeWebSafe(data).replace(/=+$/, '');
}

function deleteFirebaseAuthUser(uid) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('FIREBASE_SERVICE_ACCOUNT');
    if (!raw) return null; // not configured — caller treats this as "skip"
    var sa = JSON.parse(raw);

    // Build a signed JWT to exchange for an OAuth2 access token
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

    // Exchange JWT for a Google OAuth2 access token
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

    // Delete the user via the Firebase Auth REST API (admin endpoint)
    // Uses POST + {localId} body — same as what the Admin Node SDK calls internally
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
    // 200 = deleted, 404 = already gone (treat as success so delete is idempotent)
    return (code === 200 || code === 404);
  } catch (err) {
    Logger.log('deleteFirebaseAuthUser error: ' + err);
    return false;
  }
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
        var all = sheetToObjects(getSheet('Applications'));
        var found = all.filter(function(r) { return r.uid === String(p.uid); });
        return jsonOut(found); // [] or [{...}]
      }

      case 'update_status': {
        var sh = getSheet(p.sheet || 'Applications');
        var row = findRowIndex(sh, p.id);
        if (row > 0) setCellValue(sh, row, 'status', p.status);
        return jsonOut({ ok: row > 0 });
      }

      case 'update_field': {
        var sh2 = getSheet(p.sheet || 'Applications');
        var row2 = findRowIndex(sh2, p.id);
        if (row2 > 0) setCellValue(sh2, row2, p.field, p.value);
        return jsonOut({ ok: row2 > 0 });
      }

      case 'delete_row': {
        var sh3 = getSheet(p.sheet || 'Applications');
        var row3 = findRowIndex(sh3, p.id);
        if (row3 < 0) return jsonOut({ ok: false, rowDeleted: false, error: 'Row not found', id: p.id });

        // Delete the Sheets row first — this is the primary action.
        sh3.deleteRow(row3);
        SpreadsheetApp.flush(); // force the write to commit immediately

        // Then attempt Firebase Auth deletion (best-effort secondary action).
        // null  = service account not configured → skip (tutor can't log in but can re-register)
        // true  = Auth account fully deleted → can re-register with same email
        // false = Auth deletion failed → admin sees a warning
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

      default:
        return jsonOut([]);
    }
  } catch (err) {
    Logger.log('doGet error: ' + err);
    return jsonOut({ error: String(err) });
  }
}

// ── POST Handler ─────────────────────────────────────────────────

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var type = data.type;
    delete data.type;

    var sheetName = type === 'tutor_application' ? 'Applications'
                  : type === 'student_request'   ? 'Requests'
                  : null;
    if (!sheetName) return ContentService.createTextOutput('unknown type');

    var sh = getSheet(sheetName);

    if (sh.getLastRow() === 0) {
      // Empty sheet — create headers from incoming data keys
      sh.appendRow(Object.keys(data));
      sh.appendRow(Object.keys(data).map(function(k) {
        return data[k] !== undefined ? String(data[k]) : '';
      }));
    } else {
      var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                   .map(function(h) { return String(h).trim(); });

      // Add any new columns that don't exist yet
      Object.keys(data).forEach(function(k) {
        if (k && !hdrs.includes(k)) {
          sh.getRange(1, sh.getLastColumn() + 1).setValue(k);
          hdrs.push(k);
        }
      });

      var row = hdrs.map(function(h) {
        return data[h] !== undefined ? String(data[h]) : '';
      });
      sh.appendRow(row);
    }

    return ContentService.createTextOutput('ok');
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return ContentService.createTextOutput('error: ' + err);
  }
}
