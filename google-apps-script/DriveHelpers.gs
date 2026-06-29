/* ============================================================================
   DriveHelpers.gs — Google Drive Helper Layer
   TutorHut Document Management System
   ============================================================================

   PURPOSE
   -------
   Reusable helper functions for managing the TutorHut Google Drive folder
   structure. This module handles folder creation, lookup, and validation.
   It does NOT handle file uploads — those are built on top of these helpers
   in future modules.

   REQUIRED SETUP (one-time, performed by admin before first use)
   -------------------------------------------------------------
   1. Sign in to the Google account that OWNS the Apps Script deployment.
   2. Create a folder in Google Drive named "TutorHut Documents".
   3. Open the folder — copy the Folder ID from the browser URL:
        https://drive.google.com/drive/folders/{FOLDER_ID_IS_HERE}
   4. In the Apps Script editor:
        Project Settings → Script Properties → Add property:
          Key:   TUTORHUT_DOCUMENTS_ROOT_FOLDER_ID
          Value: {paste the Folder ID}
   5. Run driveTestSetup() from the Apps Script editor to verify.

   TUTOR FOLDER NAMING CONVENTION
   --------------------------------
   Tutor folders use the TutorHut public ID — NOT the Firebase UID.
   Format: TH{4-digit year}{5-digit sequence}
   Examples:
     TH202600001   (first tutor, 2026)
     TH202600002
     TH202700001   (first tutor registered in 2027)

   DRIVE STRUCTURE CREATED PER TUTOR
   -----------------------------------
   TutorHut Documents/
   └── TH202600001/
       ├── Identity/
       │   └── _archive/
       ├── Right To Work/
       │   └── _archive/
       ├── Qualifications/
       ├── DBS/
       │   └── _archive/
       ├── References/
       ├── Safeguarding/
       │   └── _archive/
       └── Additional/

   RELATIONSHIP TO GOOGLE SHEETS
   ------------------------------
   When a tutor folder is created, the caller is responsible for storing
   the returned folderId in the Tutors sheet (tutorDriveFolderId column).
   This module does not write to any sheet — it is Drive-only.

   FUTURE USE
   ----------
   Upload modules (passport, DBS, qualifications, etc.) will call:
     driveGetOrCreateTutorFolder(tutorId)   → get or create root folder
     driveGetSubfolder(tutorFolder, name)   → navigate to category subfolder
   These two functions are the public API for future upload handlers.
   ============================================================================ */


// ── Configuration ─────────────────────────────────────────────────────────────

// Script Property key. Value is set by admin — never hardcoded here.
var DRIVE_ROOT_FOLDER_PROP = 'TUTORHUT_DOCUMENTS_ROOT_FOLDER_ID';

// Valid tutor ID pattern: TH + 4-digit year + 5-digit zero-padded sequence.
// Total: 11 characters. Examples: TH202600001, TH203099999.
var TUTOR_ID_PATTERN = /^TH\d{9}$/;

// Standard subfolder structure created inside every tutor root folder.
// archive: true  → an "_archive" child folder is created inside this subfolder.
// archive: false → flat subfolder with no children (additive or single-use).
var TUTOR_SUBFOLDER_STRUCTURE = [
  { name: 'Identity',      archive: true  },
  { name: 'Right To Work', archive: true  },
  { name: 'Qualifications',archive: false },
  { name: 'DBS',           archive: true  },
  { name: 'References',    archive: false },
  { name: 'Safeguarding',  archive: true  },
  { name: 'Additional',    archive: false }
];


// ── Public API ────────────────────────────────────────────────────────────────
// These are the functions that future upload modules should call.

/**
 * Primary entry point.
 * Returns the tutor's root Drive folder, creating the complete subfolder
 * structure if it does not already exist.
 *
 * @param  {string} tutorId  TutorHut public ID (e.g. "TH202600001")
 * @return {object} { success, tutorId, folderId, folderUrl, alreadyExisted }
 *                  On failure: { success: false, tutorId, error }
 */
function driveGetOrCreateTutorFolder(tutorId) {
  return _driveCreateTutorFolder(tutorId);
}

/**
 * Returns a named subfolder within a tutor's root Drive folder object.
 * Used by upload modules to navigate to the correct category before
 * uploading a file.
 *
 * @param  {Folder} tutorFolder   Drive Folder object (from driveGetOrCreateTutorFolder)
 * @param  {string} subfolderName One of the TUTOR_SUBFOLDER_STRUCTURE names
 * @return {Folder}               Drive Folder object for the requested subfolder
 * @throws {Error}                If the subfolder does not exist
 */
function driveGetSubfolder(tutorFolder, subfolderName) {
  var check = _driveFolderExists(tutorFolder, subfolderName);
  if (!check.exists) {
    throw new Error(
      '[DriveHelpers] Subfolder "' + subfolderName + '" not found inside ' +
      '"' + tutorFolder.getName() + '". ' +
      'Call driveGetOrCreateTutorFolder() before accessing any subfolder.'
    );
  }
  return check.folder;
}


// ── Core Folder Functions ─────────────────────────────────────────────────────

/**
 * Reads TUTORHUT_DOCUMENTS_ROOT_FOLDER_ID from Script Properties and
 * returns the corresponding Drive Folder object.
 *
 * @return {Folder}  Drive Folder object for the TutorHut Documents root
 * @throws {Error}   If the Script Property is missing or the folder is inaccessible
 */
function driveGetRootFolder() {
  var props    = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(DRIVE_ROOT_FOLDER_PROP);

  if (!folderId || folderId.trim() === '') {
    throw new Error(
      '[DriveHelpers] Script Property "' + DRIVE_ROOT_FOLDER_PROP + '" is not set. ' +
      'Admin must: (1) create a "TutorHut Documents" folder in Google Drive, ' +
      '(2) copy its Folder ID, (3) add it in Apps Script → Project Settings → ' +
      'Script Properties. See the setup instructions at the top of DriveHelpers.gs.'
    );
  }

  return driveGetFolderById(folderId.trim());
}

/**
 * Returns the Drive Folder object for a given Folder ID.
 * Throws a descriptive error if the ID is invalid or inaccessible.
 *
 * @param  {string} folderId  Google Drive Folder ID
 * @return {Folder}           Drive Folder object
 * @throws {Error}
 */
function driveGetFolderById(folderId) {
  var result = driveValidateFolderId(folderId);
  if (!result.valid) {
    throw new Error(result.error);
  }
  return result.folder;
}

/**
 * Validates a Drive Folder ID without throwing.
 * Safe to call for pre-flight checks.
 *
 * @param  {string} folderId
 * @return {{ valid: boolean, folder: Folder|null, error: string|undefined }}
 */
function driveValidateFolderId(folderId) {
  if (!folderId || typeof folderId !== 'string' || folderId.trim() === '') {
    return { valid: false, folder: null, error: '[DriveHelpers] Folder ID is empty or not a string.' };
  }

  try {
    var folder = DriveApp.getFolderById(folderId.trim());
    return { valid: true, folder: folder };
  } catch (e) {
    return {
      valid: false,
      folder: null,
      error:
        '[DriveHelpers] Folder ID "' + folderId + '" is not accessible. ' +
        'Check that the ID is correct, the folder has not been deleted, and ' +
        'this Apps Script deployment runs as the Google account that owns the folder. ' +
        'Original error: ' + e.message
    };
  }
}

/**
 * Looks up an existing tutor folder by tutorId within the root folder.
 * Does not create anything.
 *
 * @param  {string} tutorId  TutorHut public ID (e.g. "TH202600001")
 * @return {{ found: boolean, folder: Folder|null }}
 * @throws {Error}  If tutorId format is invalid or the root folder is unavailable
 */
function driveGetTutorFolder(tutorId) {
  _assertValidTutorId(tutorId);
  var rootFolder = driveGetRootFolder();
  return _driveFolderExists(rootFolder, tutorId);
}

/**
 * Checks whether a named subfolder exists within a parent Drive folder.
 * Returns the folder object if found. Does not throw.
 *
 * @param  {Folder} parentFolder  Drive Folder object
 * @param  {string} folderName    Name of the subfolder to look for
 * @return {{ exists: boolean, folder: Folder|null }}
 * @throws {Error}  If Drive denies access to the parent folder's contents
 */
function driveCheckFolderExists(parentFolder, folderName) {
  return _driveFolderExists(parentFolder, folderName);
}


// ── Validation ────────────────────────────────────────────────────────────────

/**
 * Returns true if tutorId matches the TH{YYYY}{NNNNN} format.
 *
 * @param  {string} tutorId
 * @return {boolean}
 */
function driveIsValidTutorId(tutorId) {
  if (!tutorId || typeof tutorId !== 'string') return false;
  return TUTOR_ID_PATTERN.test(tutorId);
}


// ── Setup Verification ────────────────────────────────────────────────────────

/**
 * Run this function manually from the Apps Script editor to verify that
 * the Drive setup is correct BEFORE any tutor folders are created.
 *
 * HOW TO RUN:
 *   In the Apps Script editor → select "driveTestSetup" from the function
 *   dropdown at the top → click Run → check the Execution Log below.
 *
 * @return {{ pass: boolean, failedStep?: number, reason?: string }}
 */
function driveTestSetup() {
  var ts = new Date().toISOString();
  Logger.log('[DriveHelpers][' + ts + '] === Drive Setup Verification ===');

  // Step 1: Script Property must exist and be non-empty.
  var props    = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(DRIVE_ROOT_FOLDER_PROP);

  if (!folderId || folderId.trim() === '') {
    Logger.log(
      '[DriveHelpers] FAIL Step 1/4: Script Property "' +
      DRIVE_ROOT_FOLDER_PROP + '" is not set. See setup instructions at the ' +
      'top of DriveHelpers.gs.'
    );
    return { pass: false, failedStep: 1, reason: 'Script Property missing' };
  }
  Logger.log('[DriveHelpers] PASS Step 1/4: Script Property is set.');

  // Step 2: Folder ID must resolve to an accessible Drive folder.
  var validation = driveValidateFolderId(folderId.trim());
  if (!validation.valid) {
    Logger.log('[DriveHelpers] FAIL Step 2/4: ' + validation.error);
    return { pass: false, failedStep: 2, reason: validation.error };
  }
  Logger.log(
    '[DriveHelpers] PASS Step 2/4: Root folder accessible — "' +
    validation.folder.getName() + '"'
  );

  // Step 3: tutorId format validation must accept valid IDs and reject invalid ones.
  var validIds   = ['TH202600001', 'TH202600002', 'TH202799999', 'TH203000001'];
  var invalidIds = ['th202600001', 'TH20261', 'TH2026000001', 'TH202600001X',
                    'firebase_uid_abc', '12345', '', null];
  var validationOk = true;

  validIds.forEach(function(id) {
    if (!driveIsValidTutorId(id)) {
      Logger.log('[DriveHelpers] FAIL Step 3/4: Valid ID "' + id + '" was rejected.');
      validationOk = false;
    }
  });
  invalidIds.forEach(function(id) {
    if (driveIsValidTutorId(id)) {
      Logger.log('[DriveHelpers] FAIL Step 3/4: Invalid ID "' + id + '" was accepted.');
      validationOk = false;
    }
  });

  if (!validationOk) {
    return { pass: false, failedStep: 3, reason: 'tutorId validation logic error' };
  }
  Logger.log('[DriveHelpers] PASS Step 3/4: tutorId format validation is correct.');

  // Step 4: Subfolder structure definition must be complete and contain all required names.
  var required = ['Identity', 'Right To Work', 'Qualifications',
                  'DBS', 'References', 'Safeguarding', 'Additional'];
  var defined  = TUTOR_SUBFOLDER_STRUCTURE.map(function(s) { return s.name; });
  var missing  = required.filter(function(name) { return defined.indexOf(name) === -1; });

  if (missing.length > 0) {
    Logger.log(
      '[DriveHelpers] FAIL Step 4/4: TUTOR_SUBFOLDER_STRUCTURE is missing: ' +
      missing.join(', ')
    );
    return {
      pass: false,
      failedStep: 4,
      reason: 'Missing subfolders: ' + missing.join(', ')
    };
  }
  Logger.log('[DriveHelpers] PASS Step 4/4: Subfolder structure definition is complete.');

  Logger.log(
    '[DriveHelpers][' + ts + '] === All 4 checks passed. ' +
    'Drive helpers are ready. Run driveTestCreateFolder() to test folder creation. ==='
  );
  return { pass: true };
}

/**
 * Creates a TEST tutor folder (TH000000000) to verify end-to-end folder
 * creation works. Run manually from the Apps Script editor.
 * Delete the test folder from Drive afterwards.
 *
 * @return {object} Result from driveGetOrCreateTutorFolder
 */
function driveTestCreateFolder() {
  var TEST_ID = 'TH000000000';
  Logger.log('[DriveHelpers] Running folder creation test with tutorId=' + TEST_ID);
  var result = driveGetOrCreateTutorFolder(TEST_ID);
  Logger.log('[DriveHelpers] Test result: ' + JSON.stringify(result));
  if (result.success) {
    Logger.log(
      '[DriveHelpers] SUCCESS. Folder created at: ' + result.folderUrl + '\n' +
      'Please delete the test folder "' + TEST_ID + '" from Google Drive after verification.'
    );
  } else {
    Logger.log('[DriveHelpers] FAILED: ' + result.error);
  }
  return result;
}


// ── Internal Helpers (prefixed with _ — not part of the public API) ───────────

/**
 * Internal: Creates a tutor folder and full subfolder structure.
 * Returns existing folder untouched if it already exists.
 *
 * @param  {string} tutorId
 * @return {{ success, tutorId, folderId, folderUrl, alreadyExisted } | { success: false, tutorId, error }}
 */
function _driveCreateTutorFolder(tutorId) {
  var ts = new Date().toISOString();
  var tag = '[DriveHelpers][' + ts + '] tutorId=' + tutorId;

  if (!driveIsValidTutorId(tutorId)) {
    var validationErr =
      'Invalid tutorId "' + tutorId + '". ' +
      'Required format: TH{4-digit year}{5-digit sequence}. ' +
      'Examples: TH202600001, TH202600002.';
    Logger.log(tag + ' FAIL: ' + validationErr);
    return { success: false, tutorId: tutorId, error: validationErr };
  }

  try {
    var rootFolder = driveGetRootFolder();

    // Return existing folder without touching its structure.
    var existsCheck = _driveFolderExists(rootFolder, tutorId);
    if (existsCheck.exists) {
      var ef = existsCheck.folder;
      Logger.log(tag + ' EXISTS folderId=' + ef.getId());
      return {
        success:       true,
        tutorId:       tutorId,
        folderId:      ef.getId(),
        folderUrl:     ef.getUrl(),
        alreadyExisted: true
      };
    }

    // Create the tutor root folder.
    var tutorFolder = rootFolder.createFolder(tutorId);
    Logger.log(tag + ' CREATED root folderId=' + tutorFolder.getId());

    // Create all standard subfolders defined in TUTOR_SUBFOLDER_STRUCTURE.
    TUTOR_SUBFOLDER_STRUCTURE.forEach(function(def) {
      var sub = tutorFolder.createFolder(def.name);
      Logger.log(tag + ' CREATED ' + tutorId + '/' + def.name + '/');

      if (def.archive) {
        sub.createFolder('_archive');
        Logger.log(tag + ' CREATED ' + tutorId + '/' + def.name + '/_archive/');
      }
    });

    var folderId  = tutorFolder.getId();
    var folderUrl = tutorFolder.getUrl();

    Logger.log(tag + ' SUCCESS folderId=' + folderId);

    return {
      success:       true,
      tutorId:       tutorId,
      folderId:      folderId,
      folderUrl:     folderUrl,
      alreadyExisted: false
    };

  } catch (e) {
    Logger.log(tag + ' FAIL error=' + e.message);
    return { success: false, tutorId: tutorId, error: e.message };
  }
}

/**
 * Internal: Checks whether a named subfolder exists within parentFolder.
 *
 * @param  {Folder} parentFolder
 * @param  {string} folderName
 * @return {{ exists: boolean, folder: Folder|null }}
 * @throws {Error}  On Drive permission error
 */
function _driveFolderExists(parentFolder, folderName) {
  try {
    var iter = parentFolder.getFoldersByName(folderName);
    if (iter.hasNext()) {
      return { exists: true, folder: iter.next() };
    }
    return { exists: false, folder: null };
  } catch (e) {
    throw new Error(
      '[DriveHelpers] Drive permission error reading contents of folder "' +
      parentFolder.getName() + '": ' + e.message
    );
  }
}

/**
 * Internal: Throws if tutorId does not pass format validation.
 * Used in functions that must throw rather than return an error object.
 *
 * @param  {string} tutorId
 * @throws {Error}
 */
function _assertValidTutorId(tutorId) {
  if (!driveIsValidTutorId(tutorId)) {
    throw new Error(
      '[DriveHelpers] Invalid tutorId "' + tutorId + '". ' +
      'Required format: TH{4-digit year}{5-digit sequence}. Example: TH202600001.'
    );
  }
}
