import { auth, onAuthStateChanged, signOut, ADMIN_EMAIL } from "./firebase.js";

onAuthStateChanged(auth, (user) => {
  const area = document.getElementById("nav-auth-area");
  if (!area || !user) return;

  const dn      = user.displayName || "";
  const isAdmin = user.email === ADMIN_EMAIL;
  const isTutor = dn.startsWith("tutor:");
  const name    = dn.replace(/^(tutor|student):/, "") || user.email.split("@")[0];
  const initial = name.charAt(0).toUpperCase();

  let dashUrl, dashLabel;
  if (isAdmin)       { dashUrl = "admin-dashboard.html"; dashLabel = "Admin Dashboard"; }
  else if (isTutor)  { dashUrl = "tutor-dashboard.html";  dashLabel = "My Dashboard"; }
  else               { dashUrl = "student-dashboard.html"; dashLabel = "My Dashboard"; }

  area.innerHTML = `
    <div class="dropdown">
      <button class="btn btn-outline-primary btn-sm dropdown-toggle d-flex align-items-center gap-2"
              type="button" data-bs-toggle="dropdown" aria-expanded="false">
        <span style="width:26px;height:26px;border-radius:50%;background:var(--primary);
                     color:white;display:inline-flex;align-items:center;justify-content:center;
                     font-size:.75rem;font-weight:700;flex-shrink:0">${initial}</span>
        <span class="d-none d-lg-inline">${name}</span>
      </button>
      <ul class="dropdown-menu dropdown-menu-end shadow-sm">
        <li><span class="dropdown-item-text small text-muted pb-1">${user.email}</span></li>
        <li><hr class="dropdown-divider my-1"></li>
        <li><a class="dropdown-item" href="${dashUrl}"><i class="bi bi-grid me-2"></i>${dashLabel}</a></li>
        <li><a class="dropdown-item text-danger" href="#" id="nav-signout-btn">
          <i class="bi bi-box-arrow-right me-2"></i>Sign Out
        </a></li>
      </ul>
    </div>
  `;

  document.getElementById("nav-signout-btn")?.addEventListener("click", async (e) => {
    e.preventDefault();
    await signOut(auth);
    window.location.href = "index.html";
  });
});
