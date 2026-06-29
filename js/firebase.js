import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import {
  getAuth,
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  onAuthStateChanged,
  updateProfile,
  signOut
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";

const firebaseConfig = {
  apiKey: "AIzaSyBvb6XXN8PQ48xpiz-5uq_2UMpmcIalFzo",
  authDomain: "tutorhut-1db0e.firebaseapp.com",
  projectId: "tutorhut-1db0e",
  storageBucket: "tutorhut-1db0e.firebasestorage.app",
  messagingSenderId: "752543618866",
  appId: "1:752543618866:web:22b87f412177c2ad0e292b"
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);

export {
  createUserWithEmailAndPassword, signInWithEmailAndPassword,
  onAuthStateChanged, updateProfile, signOut
};

export const SHEETS_URL  = "https://script.google.com/macros/s/AKfycbwSrtU4mu1NzhW4Hl0jhAykSvDJB_XSA5FBCHPifZ62UqoyB3zQiCBFKXwye1BdGTNJhw/exec";
export const ADMIN_EMAIL = "admin@tutorhut.com";

function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 9);
}

async function sheetsGet(params) {
  try {
    const url  = SHEETS_URL + "?" + new URLSearchParams(params).toString();
    const res  = await fetch(url);
    const data = await res.json();
    return Array.isArray(data) ? data : [];
  } catch (e) { console.warn("Sheets read failed:", e); return []; }
}

async function sheetsPost(body) {
  try {
    await fetch(SHEETS_URL, { method: "POST", mode: "no-cors", body: JSON.stringify(body) });
  } catch (e) { console.warn("Sheets write failed:", e); }
}

export function formatDate(val) {
  if (!val) return "—";
  const d = new Date(val);
  return isNaN(d) ? "—" : d.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
}

// Kept for compatibility — data now flows through the typed save functions
export async function saveToSheets() {}

/* ── Tutor Applications ── */
export async function saveTutorApplication(uid, data) {
  const id = genId();
  await sheetsPost({ type: "tutor_application", id, uid, ...data });
  return id;
}

export async function getTutorApplications() {
  return sheetsGet({ action: "get_applications" });
}

export async function getApprovedTutors() {
  return sheetsGet({ action: "get_approved_tutors" });
}

export async function updateApplicationStatus(id, status) {
  return sheetsGet({ action: "update_status", sheet: "Applications", id, status });
}

export async function updateTutorField(id, field, value) {
  return sheetsGet({ action: "update_field", sheet: "Applications", id, field, value });
}

export async function deleteTutorApplication(id, uid) {
  // The Apps Script deletes Firebase Auth first, then the Sheets row.
  // If Auth deletion fails the row is kept and { ok: false } is returned
  // so the admin can see the error and the tutor stays manageable.
  try {
    const url  = SHEETS_URL + "?" + new URLSearchParams({ action: "delete_row", sheet: "Applications", id, uid: uid || "" });
    const res  = await fetch(url);
    const data = await res.json();
    return data; // { ok, authDeleted?, error? }
  } catch (e) {
    console.warn("deleteTutorApplication failed:", e);
    return { ok: false, error: "Network error" };
  }
}

export async function getTutorByUid(uid) {
  // sheetsGet always returns an array; extract the first match or null.
  const results = await sheetsGet({ action: "get_tutor_by_uid", uid });
  return Array.isArray(results) && results.length > 0 ? results[0] : null;
}

/* ── Student Requests ── */
export async function saveStudentRequest(uid, data) {
  const id = genId();
  await sheetsPost({ type: "student_request", id, studentUid: uid, ...data });
  return id;
}

export async function getStudentRequests(uid) {
  return sheetsGet({ action: "get_student_requests", uid });
}

export async function getAllRequests() {
  return sheetsGet({ action: "get_requests" });
}

export async function updateRequestStatus(id, status) {
  return sheetsGet({ action: "update_status", sheet: "Requests", id, status });
}

/* ── Document Uploads ── */

/**
 * Returns the verification schema for every upload category, including the
 * tutor's current status for each category.
 *
 * This is the single source of truth for upload portal rendering — the
 * frontend maintains no local copy of the schema.
 *
 * @param  {string} uid  Firebase UID of the authenticated tutor
 * @return {Array}  [{ key, label, description, allowedExtensions,
 *                     multipleFiles, required, status }, ...]
 */
export async function getUploadCategories() {
  const uid = auth.currentUser?.uid;
  if (!uid) return [];
  return sheetsGet({ action: "get_upload_categories", uid });
}

/**
 * Uploads a verification document for the authenticated tutor.
 *
 * The caller passes a File object — Base64 encoding and POST transport are
 * internal implementation details of this function. The rest of the frontend
 * is transport-agnostic.
 *
 * @param  {string} uid       Firebase UID of the authenticated tutor
 * @param  {string} category  Verification category key (e.g. "identity")
 * @param  {File}   file      Browser File object selected by the tutor
 * @return {Object} { ok, documentId, version, category, tutorId } | { ok: false, error }
 */
export async function uploadDocument(category, file) {
  const uid = auth.currentUser?.uid;
  if (!uid) return { ok: false, error: "Not authenticated." };
  try {
    const base64Content = await _fileToBase64(file);
    // Content-Type: text/plain avoids CORS pre-flight while still sending JSON.
    // Apps Script receives the body via e.postData.contents regardless of content type.
    const res = await fetch(SHEETS_URL, {
      method:  "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify({
        type:             "upload_document",
        firebaseUid:      uid,
        category,
        base64Content,
        originalFilename: file.name,
        mimeType:         file.type || "application/octet-stream",
        fileSize:         file.size
      })
    });
    if (!res.ok) throw new Error("HTTP " + res.status);
    return await res.json();
  } catch (e) {
    console.warn("uploadDocument failed:", e);
    return { ok: false, error: "Upload failed — please check your connection and try again." };
  }
}

// Private: encodes a File as a base64 string (strips the data URI prefix).
// This is an implementation detail of uploadDocument() — not part of the public API.
function _fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = (e) => resolve(e.target.result.split(",")[1]);
    reader.onerror = () => reject(new Error("FileReader error"));
    reader.readAsDataURL(file);
  });
}

export function authErrorMessage(code) {
  const map = {
    "auth/email-already-in-use": "This email is already registered. Try signing in instead.",
    "auth/weak-password":        "Password must be at least 6 characters.",
    "auth/invalid-email":        "Please enter a valid email address.",
    "auth/user-not-found":       "No account found with this email.",
    "auth/wrong-password":       "Incorrect password. Please try again.",
    "auth/invalid-credential":   "Invalid email or password.",
    "auth/too-many-requests":    "Too many failed attempts. Please try again later."
  };
  return map[code] || "Something went wrong. Please try again.";
}
