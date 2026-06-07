import { auth, onAuthStateChanged, signOut, ADMIN_EMAIL } from "./firebase.js";

const NAV_LINKS = [
  { label: "Home",          href: "index.html" },
  { label: "Browse Tutors", href: "browse-tutors.html" },
  { label: "About",         href: "about.html" },
  { label: "Contact",       href: "contact.html" },
];

function currentPage() {
  return window.location.pathname.split("/").pop() || "index.html";
}

function buildNav(user) {
  const nav = document.getElementById("site-navbar");
  if (!nav) return;

  const cur   = currentPage();
  const links = NAV_LINKS.map(p =>
    `<li class="nav-item"><a class="nav-link${cur === p.href ? " active" : ""}" href="${p.href}">${p.label}</a></li>`
  ).join("") +
  `<li class="nav-item"><a class="nav-link text-muted" href="tutor-login.html" style="font-size:.85rem">Tutor Area</a></li>`;

  let authHtml;
  if (user) {
    const dn      = user.displayName || "";
    const isAdmin = user.email === ADMIN_EMAIL;
    const isTutor = dn.startsWith("tutor:");
    const name    = dn.replace(/^(tutor|student):/, "") || user.email.split("@")[0];
    const initial = name.charAt(0).toUpperCase();
    const dashUrl = isAdmin ? "admin-dashboard.html" : isTutor ? "tutor-dashboard.html" : "student-dashboard.html";
    const dashLbl = isAdmin ? "Admin Dashboard" : "My Dashboard";

    authHtml = `
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
          <li><a class="dropdown-item" href="${dashUrl}"><i class="bi bi-grid me-2"></i>${dashLbl}</a></li>
          <li><a class="dropdown-item text-danger" href="#" id="nav-signout-btn">
            <i class="bi bi-box-arrow-right me-2"></i>Sign Out
          </a></li>
        </ul>
      </div>`;
  } else {
    authHtml = `
      <a href="student-register.html" class="btn btn-outline-primary btn-sm"><i class="bi bi-mortarboard me-1"></i>Student</a>
      <a href="tutor-register.html" class="btn btn-primary btn-sm"><i class="bi bi-person-workspace me-1"></i>Tutor</a>`;
  }

  nav.innerHTML = `
    <div class="container">
      <a class="navbar-brand" href="index.html">Tutor<span>Hut</span></a>
      <button class="navbar-toggler border-0" type="button" data-bs-toggle="collapse" data-bs-target="#mainNav">
        <span class="navbar-toggler-icon"></span>
      </button>
      <div class="collapse navbar-collapse" id="mainNav">
        <ul class="navbar-nav mx-auto gap-1">${links}</ul>
        <div class="d-flex align-items-center gap-2 mt-2 mt-lg-0">${authHtml}</div>
      </div>
    </div>`;

  document.getElementById("nav-signout-btn")?.addEventListener("click", async (e) => {
    e.preventDefault();
    await signOut(auth);
    window.location.href = "index.html";
  });
}

onAuthStateChanged(auth, buildNav);
