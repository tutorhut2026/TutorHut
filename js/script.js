/* =============================================
   TutorHut – script.js
   ============================================= */

document.addEventListener('DOMContentLoaded', function () {

  /* ── Toggle day/time chips ── */
  document.querySelectorAll('.day-btn, .time-btn').forEach(btn => {
    btn.addEventListener('click', function () {
      this.classList.toggle('active');
    });
  });

  /* ── Subject pill checkboxes ── */
  document.querySelectorAll('.subject-check-label').forEach(label => {
    label.addEventListener('click', function () {
      const inp = this.querySelector('input[type="checkbox"]');
      if (inp) {
        inp.checked = !inp.checked;
        this.classList.toggle('selected', inp.checked);
      }
    });
  });

  /* ── Browse Tutors: filter ── */
  const filterForm = document.getElementById('tutor-filter-form');
  if (filterForm) {
    filterForm.addEventListener('submit', function (e) {
      e.preventDefault();
      const subject  = document.getElementById('filter-subject')?.value.toLowerCase();
      const day      = document.getElementById('filter-day')?.value.toLowerCase();
      const timeSlot = document.getElementById('filter-time')?.value.toLowerCase();

      document.querySelectorAll('.tutor-card-item').forEach(card => {
        const subjects   = (card.dataset.subjects || '').toLowerCase();
        const days       = (card.dataset.days || '').toLowerCase();
        const times      = (card.dataset.times || '').toLowerCase();

        const matchSub  = !subject  || subjects.includes(subject);
        const matchDay  = !day      || days.includes(day);
        const matchTime = !timeSlot || times.includes(timeSlot);

        card.closest('.col').style.display = (matchSub && matchDay && matchTime) ? '' : 'none';
      });
    });

    const resetBtn = document.getElementById('filter-reset');
    if (resetBtn) {
      resetBtn.addEventListener('click', function () {
        filterForm.reset();
        document.querySelectorAll('.tutor-card-item').forEach(card => {
          card.closest('.col').style.display = '';
        });
      });
    }
  }

  /* ── Admin: approve / reject buttons (demo) ── */
  document.querySelectorAll('.btn-approve').forEach(btn => {
    btn.addEventListener('click', function () {
      const row = this.closest('tr');
      const badge = row?.querySelector('.status-badge');
      if (badge) {
        badge.textContent = 'Approved';
        badge.className = 'badge badge-status-approved status-badge';
      }
      this.disabled = true;
      const rejectBtn = row?.querySelector('.btn-reject');
      if (rejectBtn) rejectBtn.disabled = true;
      showToast('Tutor approved successfully!', 'success');
    });
  });

  document.querySelectorAll('.btn-reject').forEach(btn => {
    btn.addEventListener('click', function () {
      const row = this.closest('tr');
      const badge = row?.querySelector('.status-badge');
      if (badge) {
        badge.textContent = 'Rejected';
        badge.className = 'badge badge-status-rejected status-badge';
      }
      this.disabled = true;
      const approveBtn = row?.querySelector('.btn-approve');
      if (approveBtn) approveBtn.disabled = true;
      showToast('Application rejected.', 'danger');
    });
  });

  /* ── Toast notification ── */
  function showToast(message, type = 'success') {
    const toastContainer = getOrCreateToastContainer();
    const id = 'toast-' + Date.now();
    const bgClass = type === 'success' ? 'bg-success' : type === 'danger' ? 'bg-danger' : 'bg-primary';
    const html = `
      <div id="${id}" class="toast align-items-center text-white ${bgClass} border-0" role="alert">
        <div class="d-flex">
          <div class="toast-body fw-500">${message}</div>
          <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
        </div>
      </div>`;
    toastContainer.insertAdjacentHTML('beforeend', html);
    const toastEl = document.getElementById(id);
    const toast   = new bootstrap.Toast(toastEl, { delay: 3000 });
    toast.show();
    toastEl.addEventListener('hidden.bs.toast', () => toastEl.remove());
  }

  function getOrCreateToastContainer() {
    let c = document.getElementById('toast-container');
    if (!c) {
      c = document.createElement('div');
      c.id = 'toast-container';
      c.className = 'toast-container position-fixed bottom-0 end-0 p-3';
      c.style.zIndex = 9999;
      document.body.appendChild(c);
    }
    return c;
  }

  /* ── Tuition request form: auto-fill tutor name ── */
  const tutorParam = new URLSearchParams(window.location.search).get('tutor');
  const tutorInput = document.getElementById('request-tutor');
  if (tutorInput && tutorParam) tutorInput.value = decodeURIComponent(tutorParam);

  /* ── Generic form submission feedback ── */
  document.querySelectorAll('form.th-form').forEach(form => {
    if (form.dataset.firebase) return; // handled by firebase.js
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      if (!this.checkValidity()) { this.reportValidity(); return; }
      const successAlert = this.querySelector('.form-success');
      if (successAlert) {
        successAlert.classList.remove('d-none');
        this.querySelectorAll('input, select, textarea').forEach(el => el.disabled = true);
        this.querySelector('[type="submit"]').disabled = true;
      }
    });
  });

  /* ── Smooth active nav link highlight ── */
  const path = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.navbar-nav .nav-link').forEach(link => {
    if (link.getAttribute('href') === path) link.classList.add('active');
  });

  /* ── Availability summary in registration ── */
  updateAvailabilitySummary();
  document.querySelectorAll('.day-btn, .time-btn').forEach(btn => {
    btn.addEventListener('click', updateAvailabilitySummary);
  });

  function updateAvailabilitySummary() {
    const summaryEl = document.getElementById('availability-summary');
    if (!summaryEl) return;
    const days  = [...document.querySelectorAll('.day-btn.active')].map(b => b.textContent.trim());
    const times = [...document.querySelectorAll('.time-btn.active')].map(b => b.textContent.trim());
    if (!days.length && !times.length) {
      summaryEl.textContent = 'No availability selected yet.';
    } else {
      summaryEl.textContent = `Days: ${days.join(', ') || 'None'} | Times: ${times.join(', ') || 'None'}`;
    }
  }

  /* ── Tooltips init ── */
  const tooltipEls = document.querySelectorAll('[data-bs-toggle="tooltip"]');
  tooltipEls.forEach(el => new bootstrap.Tooltip(el));

});
