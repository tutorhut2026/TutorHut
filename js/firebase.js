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

export const SHEETS_URL  = "https://script.google.com/macros/s/AKfycbwan5IHb7NYCeqjvS34PdewngOHBZpcSu5XfqzqTfRnOinrKU-PwPX91tWAoo-wLNnz/exec";
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
  // Step 1: mark as removed so the login check blocks them immediately,
  // even if the row deletion below fails for any reason.
  await updateApplicationStatus(id, "removed");
  // Step 2: delete the row from Sheets (also triggers Firebase Auth deletion
  // in the Apps Script if FIREBASE_SERVICE_ACCOUNT is configured there).
  return sheetsGet({ action: "delete_row", sheet: "Applications", id, uid: uid || "" });
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
