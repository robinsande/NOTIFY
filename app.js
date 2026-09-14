const API_BASE = 'https://backend-notify-3pvi.onrender.com/api';
const APP_STATE_KEY = 'notify-app-state';
const NAV_STATE_KEY = 'notify-nav-hidden';
let currentUser = null;
let navLinksHidden = localStorage.getItem(NAV_STATE_KEY) === 'true';

function saveAppState(viewKey, authenticated = true) {
  try {
    localStorage.setItem(APP_STATE_KEY, JSON.stringify({ view: viewKey, authenticated, user: currentUser }));
  } catch (error) {
    console.warn('Unable to save app state', error);
  }
}

function restoreAppState() {
  try {
    const state = JSON.parse(localStorage.getItem(APP_STATE_KEY) || '{}');
    if (state.user) {
      currentUser = state.user;
    }
    return state;
  } catch (error) {
    return {};
  }
}

function clearAppState() {
  try {
    localStorage.removeItem(APP_STATE_KEY);
  } catch (error) {
    console.warn('Unable to clear app state', error);
  }
}

function toggleNavLinks(forceState) {
  navLinksHidden = typeof forceState === 'boolean' ? forceState : !navLinksHidden;
  try {
    localStorage.setItem(NAV_STATE_KEY, String(navLinksHidden));
  } catch (error) {
    console.warn('Unable to save navigation preference', error);
  }

  const shell = document.querySelector('.app-shell');
  if (shell) {
    shell.classList.toggle('nav-links-hidden', navLinksHidden);
  }

  const toggle = document.querySelector('[data-nav-links-toggle]');
  if (toggle) {
    toggle.setAttribute('aria-pressed', String(navLinksHidden));
    toggle.textContent = navLinksHidden ? 'Show nav' : 'Hide nav';
  }
}

function logoutUser() {
  currentUser = null;
  clearAppState();
  if (window.location) {
    window.location.hash = '';
  }
  renderLogin();
}

function navigateToSavedView(viewKey) {
  switch (viewKey) {
    case 'dashboard':
      return renderDashboard();
    case 'events':
      return renderEvents();
    case 'calendar':
      return renderCalendar();
    case 'tasks':
      return renderTasks();
    case 'attendees':
      return renderAttendees();
    case 'reminders':
      return renderNotifications();
    case 'documents':
      return renderDocuments();
    case 'reports':
      return renderReports();
    case 'departments':
      return renderDepartments();
    case 'eventTypes':
      return renderEventTypes();
    case 'admin':
      return renderAdminPanel();
    default:
      return renderLogin();
  }
}

function navigateToRoute(routeKey) {
  const normalized = String(routeKey || '').trim();
  const routeMap = {
    dashboard: renderDashboard,
    events: renderEvents,
    calendar: renderCalendar,
    tasks: renderTasks,
    attendees: renderAttendees,
    reminders: renderNotifications,
    documents: renderDocuments,
    reports: renderReports,
    departments: renderDepartments,
    eventTypes: renderEventTypes,
    admin: renderAdminPanel
  };

  const target = routeMap[normalized] || renderDashboard;
  const nextHash = normalized === 'dashboard' ? '' : `#/` + normalized;
  if (window.location.hash !== nextHash) {
    window.location.hash = nextHash;
  }
  target();
}

function handleRouteFromHash() {
  const hashRoute = window.location.hash.replace(/^#\/?/, '').trim();
  if (!hashRoute) {
    return;
  }

  const routes = ['dashboard','events','calendar','tasks','attendees','reminders','documents','reports','departments','eventTypes','admin'];
  if (routes.includes(hashRoute)) {
    const routeMap = {
      dashboard: renderDashboard,
      events: renderEvents,
      calendar: renderCalendar,
      tasks: renderTasks,
      attendees: renderAttendees,
      reminders: renderNotifications,
      documents: renderDocuments,
      reports: renderReports,
      departments: renderDepartments,
      eventTypes: renderEventTypes,
      admin: renderAdminPanel
    };
    routeMap[hashRoute]();
  }
}

const navigationItems = [
  { key: 'dashboard', icon: '📊', label: 'Dashboard', action: 'renderDashboard()' },
  { key: 'events', icon: '🗓️', label: 'Events', action: 'renderEvents()' },
  { key: 'calendar', icon: '📅', label: 'Calendar', action: 'renderCalendar()' },
  { key: 'tasks', icon: '✅', label: 'Tasks', action: 'renderTasks()' },
  { key: 'attendees', icon: '👥', label: 'Attendees', action: 'renderAttendees()' },
  { key: 'reminders', icon: '🔔', label: 'Reminders', action: 'renderNotifications()' },
  { key: 'documents', icon: '📄', label: 'Documents', action: 'renderDocuments()' },
  { key: 'reports', icon: '📈', label: 'Reports', action: 'renderReports()' },
  { key: 'departments', icon: '🏢', label: 'Departments', action: 'renderDepartments()' },
  { key: 'eventTypes', icon: '🏷️', label: 'Event Types', action: 'renderEventTypes()' }
];

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.error || 'Request failed');
  }

  const payload = await response.json();
  if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
    payload.ok = response.ok;
  }
  return payload;
}

function adminRequestOptions(options = {}) {
  return {
    ...options,
    headers: {
      ...(options.headers || {}),
      'X-Notify-User-Id': currentUser?.id || ''
    }
  };
}

function normalizeHeader(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function normalizeCellValue(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'number') {
    return String(value);
  }
  return String(value).trim();
}

function getMappedValue(row, aliases) {
  for (const [key, rawValue] of Object.entries(row)) {
    if (aliases.includes(normalizeHeader(key))) {
      return normalizeCellValue(rawValue);
    }
  }
  return '';
}

function splitAttendeeValues(value) {
  if (value === null || value === undefined || value === '') {
    return [];
  }

  return String(value)
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split(/[\n;|,]/)
    .map(item => item.trim())
    .filter(Boolean);
}

function parseSpreadsheetFile(file) {
  return new Promise((resolve, reject) => {
    const isCsv = file.name.toLowerCase().endsWith('.csv');
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = event.target.result;
        const workbook = XLSX.read(isCsv ? data : data, { type: isCsv ? 'string' : 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        resolve(rows);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('Unable to read the selected file.'));
    if (isCsv) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

async function importSpreadsheetRows({ type, file, onImported }) {
  if (!file) {
    return { imported: 0, skipped: 0 };
  }

  const rows = await parseSpreadsheetFile(file);
  const items = [];

  rows.forEach((row) => {
    const cleaned = Object.fromEntries(Object.entries(row).map(([key, value]) => [key, normalizeCellValue(value)]));

    if (type === 'attendee') {
      const name = getMappedValue(cleaned, ['name','fullname','fullName','attendeename','participantname','attendeenames']) || '';
      const email = getMappedValue(cleaned, ['email','emailaddress','attendeeemail','participantemail']) || '';
      const status = getMappedValue(cleaned, ['status','attendancestatus','attendeestatus']) || 'Confirmed';
      const attendeeNames = splitAttendeeValues(getMappedValue(cleaned, ['attendeenames','guestnames','participants','names']));
      const attendeeEmails = splitAttendeeValues(getMappedValue(cleaned, ['attendeeemails','guestemails','participantemails','emails']));

      if (name || email || attendeeNames.length || attendeeEmails.length) {
        if (name && email) {
          items.push({ name, email, status });
          return;
        }

        if (attendeeNames.length || attendeeEmails.length) {
          const maxEntries = Math.max(attendeeNames.length, attendeeEmails.length, 0);
          for (let index = 0; index < maxEntries; index++) {
            const entryName = attendeeNames[index] || '';
            const entryEmail = attendeeEmails[index] || '';
            if (entryName || entryEmail) {
              items.push({ name: entryName, email: entryEmail, status });
            }
          }
          return;
        }

        if (name || email) {
          items.push({ name, email, status });
        }
      }
    }

    if (type === 'event') {
      const name = getMappedValue(cleaned, ['eventname','name','title']) || '';
      const typeValue = getMappedValue(cleaned, ['type','eventtype']) || 'Other';
      const startDate = getMappedValue(cleaned, ['startdate','date','eventstartdate','start']) || '';
      const endDate = getMappedValue(cleaned, ['enddate','eventenddate','end']) || '';
      const venue = getMappedValue(cleaned, ['venue','location','place']) || '';
      const coordinator = getMappedValue(cleaned, ['coordinator','contactperson','organizer','eventcoordinator']) || '';
      const expectedAttendees = getMappedValue(cleaned, ['expectedattendees','attendees','expected']) || '';
      const department = getMappedValue(cleaned, ['department','division','unit']) || '';
      const status = getMappedValue(cleaned, ['status','eventstatus']) || 'Planning';

      const attendeeName = getMappedValue(cleaned, ['attendeename','guestname','participantname','nameofattendee']) || '';
      const attendeeEmail = getMappedValue(cleaned, ['attendeeemail','guestemail','participantemail','emailaddress']) || '';
      const rawNameList = getMappedValue(cleaned, ['attendeenames','guestnames','participants','names','attendeename','guestname']) || '';
      const rawEmailList = getMappedValue(cleaned, ['attendeeemails','guestemails','participantemails','emails','attendeeemail','guestemail']) || '';

      const parsedAttendees = [];
      const directNames = splitAttendeeValues(rawNameList);
      const directEmails = splitAttendeeValues(rawEmailList);

      if (attendeeName && attendeeEmail) {
        parsedAttendees.push({ name: attendeeName, email: attendeeEmail, status: 'Confirmed' });
      }

      if (directNames.length || directEmails.length) {
        const size = Math.max(directNames.length, directEmails.length, 0);
        for (let index = 0; index < size; index++) {
          const nameValue = directNames[index] || '';
          const emailValue = directEmails[index] || '';
          if (nameValue || emailValue) {
            parsedAttendees.push({ name: nameValue, email: emailValue, status: 'Confirmed' });
          }
        }
      }

      if (name) {
        items.push({
          name,
          type: typeValue,
          date: startDate,
          startDate,
          endDate,
          venue,
          coordinator,
          expectedAttendees,
          department,
          status,
          attendees: parsedAttendees
        });
      }
    }
  });

  for (const item of items) {
    if (type === 'attendee') {
      await fetchJson(`${API_BASE}/attendees`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(item) });
    } else {
      await fetchJson(`${API_BASE}/events`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(item) });
    }
  }

  return { imported: items.length, skipped: rows.length - items.length };
}

function formatDate(value) {
  if (!value) return '—';
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleDateString('en', { month: 'short', day: 'numeric', year: 'numeric' });
}

function statusBadge(status) {
  const classes = {
    Planning: 'badge-planning',
    Confirmed: 'badge-confirmed',
    Completed: 'badge-completed',
    Cancelled: 'badge-cancelled',
    Pending: 'badge-pending',
    'In Progress': 'badge-progress',
    Overdue: 'badge-overdue',
    Scheduled: 'badge-planning',
    Sent: 'badge-confirmed'
  };
  const cls = classes[status] || 'badge-neutral';
  return `<span class="status-badge ${cls}">${status}</span>`;
}

let calendarViewDate = new Date();
let calendarSelectedDate = new Date();

function getCalendarDateKey(date) {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, '0');
  const day = `${date.getDate()}`.padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getEventDateKey(event) {
  const rawValue = event.startDate || event.date || '';
  const parsedDate = new Date(rawValue);
  return Number.isNaN(parsedDate.getTime()) ? null : getCalendarDateKey(parsedDate);
}

function getCalendarDays(viewDate) {
  const year = viewDate.getFullYear();
  const monthIndex = viewDate.getMonth();
  const firstDay = new Date(year, monthIndex, 1);
  const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
  const leadingDays = (firstDay.getDay() + 6) % 7;
  const totalCells = Math.ceil((leadingDays + daysInMonth) / 7) * 7;

  return Array.from({ length: totalCells }, (_, index) => {
    const dayOffset = index - leadingDays + 1;
    const date = new Date(year, monthIndex, dayOffset);
    return {
      date,
      isCurrentMonth: date.getMonth() === monthIndex
    };
  });
}

function shiftCalendarMonth(offset) {
  calendarViewDate = new Date(calendarViewDate.getFullYear(), calendarViewDate.getMonth() + offset, 1);
  renderCalendar();
}

function selectCalendarDate(dateKey) {
  calendarSelectedDate = new Date(`${dateKey}T00:00:00`);
  renderCalendar();
}

function getTaskStatusSummary(tasks = []) {
  const summary = { Pending: 0, 'In Progress': 0, Completed: 0, Overdue: 0 };
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  tasks.forEach((task) => {
    const dueDate = new Date(task.dueDate);
    const isOverdue = !['Completed', 'Cancelled'].includes(task.status) && dueDate < today;
    if (isOverdue) {
      summary.Overdue += 1;
    } else if (summary[task.status] !== undefined) {
      summary[task.status] += 1;
    } else {
      summary.Pending += 1;
    }
  });

  return summary;
}

function getDepartmentSummary(events = []) {
  const summary = {};
  events.forEach((event) => {
    summary[event.department] = (summary[event.department] || 0) + 1;
  });
  return summary;
}

let sidebarOpen = window.innerWidth > 767;

function isSmallScreen() {
  return window.matchMedia('(max-width: 767px)').matches;
}

function toggleSidebarDropdown(forceState) {
  const nextState = typeof forceState === 'boolean' ? forceState : !sidebarOpen;
  sidebarOpen = nextState;

  const sidebar = document.querySelector('.sidebar');
  if (!sidebar) {
    return;
  }

  if (isSmallScreen()) {
    sidebar.classList.toggle('is-open', sidebarOpen);
    sidebar.classList.toggle('is-collapsed', !sidebarOpen);
  } else {
    sidebar.classList.remove('is-open');
    sidebar.classList.remove('is-collapsed');
  }

  const toggleButton = document.querySelector('.sidebar-toggle');
  if (toggleButton) {
    toggleButton.setAttribute('aria-expanded', String(sidebarOpen));
    toggleButton.textContent = sidebarOpen ? '✕' : '☰';
  }
}

window.addEventListener('resize', () => {
  if (window.innerWidth > 767) {
    sidebarOpen = true;
    toggleSidebarDropdown(true);
    return;
  }

  sidebarOpen = false;
  toggleSidebarDropdown(false);
});

function renderShell(activeKey, title, subtitle, content, actionsHtml = '') {
  const navHtml = navigationItems.map(item => `
    <li class="nav-item">
      <a class="nav-link ${item.key === activeKey ? 'active' : ''}" href="#" onclick="event.preventDefault(); navigateToRoute('${item.key}')">
        <span class="nav-icon">${item.icon}</span>
        <span>${item.label}</span>
      </a>
    </li>
  `).join('');

  const adminPanelHtml = currentUser && currentUser.role === 'admin' ? `
    <li class="nav-item">
      <a class="nav-link ${activeKey === 'admin' ? 'active' : ''}" href="#" onclick="event.preventDefault(); navigateToRoute('admin')">
        <span class="nav-icon">⚙️</span>
        <span>Admin Panel</span>
      </a>
    </li>
  ` : '';

  const userRoleLabel = currentUser ? (currentUser.role === 'admin' ? 'Administrator' : 'Viewer') : 'Administrator';

  return `
    <div class="app-shell ${navLinksHidden ? 'nav-links-hidden' : ''}">
      <aside class="sidebar ${sidebarOpen ? 'is-open' : ''}">
        <div>
          <div class="brand mb-4">
            <img class="brand-mark" src="notify-logo.svg" alt="NOTIFY logo" />
            <div>
              <h3>NOTIFY</h3>
              <p>Operations Control</p>
            </div>
            <button class="sidebar-toggle" type="button" aria-label="Toggle sidebar" aria-expanded="${String(sidebarOpen)}" onclick="toggleSidebarDropdown()">${sidebarOpen ? '✕' : '☰'}</button>
          </div>
          <div class="sidebar-nav-header">
            <span>Navigation</span>
            <button class="btn btn-sm btn-outline-light" type="button" data-nav-links-toggle aria-pressed="${String(navLinksHidden)}" onclick="toggleNavLinks()">${navLinksHidden ? 'Show nav' : 'Hide nav'}</button>
          </div>
          <ul class="nav flex-column sidebar-nav">${navHtml}${adminPanelHtml}</ul>
        </div>
        <div class="sidebar-footer">
          <div class="small text-white-50">${userRoleLabel}</div>
          <button class="btn btn-outline-light btn-sm mt-2" onclick="logoutUser()">Logout</button>
        </div>
      </aside>
      <main class="main-panel">
        <header class="topbar">
          <div>
            <p class="eyebrow">Enterprise administration</p>
            <h2>${title}</h2>
            <p class="text-muted mb-0">${subtitle}</p>
          </div>
          <div class="topbar-actions">
            <div class="notification-pill">🔔 Notifications</div>
            ${actionsHtml}
          </div>
        </header>
        <div class="content-area">${content}</div>
      </main>
    </div>
  `;
}

function renderLogin() {
  clearAppState();
  if (window.location) {
    window.location.hash = '';
  }
  document.getElementById('app').innerHTML = `
    <div class="login-wrap">
      <div class="card login-card p-4 p-lg-5">
        <div class="text-center mb-4">
          <img class="brand-mark brand-mark-large mx-auto" src="notify-logo.svg" alt="NOTIFY logo" />
          <h3 class="mt-3 mb-1">Welcome back</h3>
          <p class="text-muted">Access the NOTIFY operations dashboard securely</p>
        </div>

        <form id="auth-form" class="auth-form">
          <div class="mb-3 auth-field" id="full-name-field" style="display: none;">
            <label class="form-label">Full name</label>
            <input class="form-control" name="fullName" />
          </div>
          <div class="mb-3">
            <label class="form-label">Email address</label>
            <input class="form-control" name="email" placeholder="example@gmail.com" />
            <div class="form-text text-muted" id="otp-help-text" style="display: none;">Enter your email and we’ll send you a one-time code.</div>
          </div>
          <div class="mb-3" id="password-field">
            <label class="form-label">Password</label>
            <input type="password" class="form-control" name="password" />
          </div>
          <div class="mb-3" id="confirm-password-field" style="display: none;">
            <label class="form-label">Confirm password</label>
            <input type="password" class="form-control" name="confirmPassword" />
          </div>
          <div class="mb-3" id="otp-field" style="display: none;">
            <label class="form-label">OTP code</label>
            <input class="form-control" name="otpCode" placeholder="Enter the 6-digit code" />
          </div>
          <div class="form-check mb-3" id="forgot-password-check">
            <input class="form-check-input" type="checkbox" id="forgot-password" />
            <label class="form-check-label" for="forgot-password">Forgot password? Send OTP</label>
          </div>
          <div class="d-grid gap-2 mt-4">
            <button class="btn btn-primary w-100" type="submit" id="auth-submit-btn">Sign In</button>
          </div>
        </form>

        <div class="text-center mt-4 pt-3 border-top" id="auth-footer">
          <p class="text-muted mb-2" id="auth-footer-text">New here?</p>
          <button class="btn btn-outline-secondary btn-sm" type="button" id="create-account-link">Create an account</button>
          <button class="btn btn-link btn-sm mt-2" type="button" id="back-to-signin-link" style="display: none;">Back to sign in</button>
        </div>
      </div>
    </div>
  `;

  const form = document.getElementById('auth-form');
  const fullNameField = document.getElementById('full-name-field');
  const passwordField = document.getElementById('password-field');
  const confirmPasswordField = document.getElementById('confirm-password-field');
  const otpField = document.getElementById('otp-field');
  const forgotPasswordCheck = document.getElementById('forgot-password');
  const submitButton = document.getElementById('auth-submit-btn');
  const otpHelpText = document.getElementById('otp-help-text');
  const createAccountLink = document.getElementById('create-account-link');
  const backToSigninLink = document.getElementById('back-to-signin-link');
  const authFooterText = document.getElementById('auth-footer-text');
  let activeView = 'signin';
  let resetStage = 'request';

  function clearAuthFormFields() {
    form.reset();
    form.elements.email.value = '';
    form.elements.password.value = '';
    form.elements.fullName.value = '';
    form.elements.confirmPassword.value = '';
    form.elements.otpCode.value = '';
    forgotPasswordCheck.checked = false;
  }

  function setView(view) {
    activeView = view;
    const isRegister = view === 'register';
    const isReset = view === 'reset';
    const showPasswordField = !isReset || resetStage === 'new-password';
    fullNameField.style.display = isRegister ? 'block' : 'none';
    passwordField.style.display = showPasswordField ? 'block' : 'none';
    confirmPasswordField.style.display = isRegister ? 'block' : 'none';
    otpField.style.display = isReset && resetStage === 'verify' ? 'block' : 'none';
    otpHelpText.style.display = isReset ? 'block' : 'none';
    submitButton.textContent = isRegister ? 'Create Account' : isReset ? (resetStage === 'new-password' ? 'Set New Password' : resetStage === 'verify' ? 'Verify OTP' : 'Send OTP') : 'Sign In';

    if (isReset) {
      forgotPasswordCheck.checked = true;
    } else {
      forgotPasswordCheck.checked = false;
    }

    const showBackLink = view !== 'signin';
    backToSigninLink.style.display = showBackLink ? 'inline-block' : 'none';
    authFooterText.textContent = view === 'register' ? 'Already have an account?' : view === 'reset' ? 'Remembered your password?' : 'New here?';
    createAccountLink.style.display = view === 'register' ? 'none' : 'inline-block';
  }

  createAccountLink.addEventListener('click', () => {
    resetStage = 'request';
    clearAuthFormFields();
    setView('register');
  });

  backToSigninLink.addEventListener('click', () => {
    resetStage = 'request';
    clearAuthFormFields();
    setView('signin');
  });

  forgotPasswordCheck.addEventListener('change', () => {
    if (forgotPasswordCheck.checked) {
      activeView = 'reset';
      resetStage = 'request';
      clearAuthFormFields();
      setView('reset');
    } else {
      resetStage = 'request';
      clearAuthFormFields();
      setView('signin');
    }
  });
  setView('signin');

  function showAuthMessage(messageText, type = 'danger') {
    const existing = form.querySelector('.auth-feedback');
    if (existing) {
      existing.remove();
    }

    const message = document.createElement('div');
    message.className = `alert alert-${type} mt-3 auth-feedback`;
    message.textContent = messageText;
    form.appendChild(message);
  }

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const email = form.elements.email.value;
    const password = form.elements.password.value;
    const fullName = form.elements.fullName.value;
    const confirmPassword = form.elements.confirmPassword.value;
    const otpCode = form.elements.otpCode.value;

    if (activeView === 'register') {
      if (password !== confirmPassword) {
        showAuthMessage('Passwords do not match.', 'danger');
        return;
      }

      try {
        await fetchJson(`${API_BASE}/auth/register`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email, password, fullName })
        });

        clearAuthFormFields();
        resetStage = 'request';
        setView('signin');
        showAuthMessage('Account created successfully. Please sign in.', 'success');
      } catch (error) {
        showAuthMessage(error.message || 'Unable to create account. Please try a different email.', 'danger');
      }
      return;
    }

    if (activeView === 'reset') {
      if (resetStage === 'request') {
        try {
          await fetchJson(`${API_BASE}/auth/password-reset`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email })
          });
          resetStage = 'verify';
          setView('reset');
          showAuthMessage('OTP sent to your email. Enter the code to continue.', 'success');
        } catch (error) {
          showAuthMessage(error.message || 'Unable to send OTP.', 'danger');
        }
        return;
      }

      if (resetStage === 'verify') {
        if (!otpCode) {
          showAuthMessage('Please enter the OTP code.', 'danger');
          return;
        }

        resetStage = 'new-password';
        setView('reset');
        showAuthMessage('OTP verified. Enter your new password below.', 'info');
        return;
      }

      if (!password) {
        showAuthMessage('Please enter a new password.', 'danger');
        return;
      }

      try {
        await fetchJson(`${API_BASE}/auth/password-reset`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email, newPassword: password, otpCode })
        });
        showAuthMessage('Password updated successfully. You can now sign in.', 'success');
      } catch (error) {
        showAuthMessage(error.message || 'Unable to reset password.', 'danger');
      }
      return;
    }

    try {
      const response = await fetchJson(`${API_BASE}/auth/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password })
      });

      if (response && response.ok && response.user) {
        currentUser = response.user;
        saveAppState('dashboard', true);
        
        // Force password change on first login for viewers
        if (response.user.isFirstLogin) {
          renderFirstLoginPasswordChange(response.user.email);
          return;
        }
        
        renderDashboard();
        return;
      }

      throw new Error('Invalid email or password');
    } catch (error) {
      showAuthMessage(error.message || 'Invalid email or password.', 'danger');
    }
  });
}

function renderFirstLoginPasswordChange(userEmail) {
  document.getElementById('app').innerHTML = `
    <div class="login-wrap">
      <div class="card login-card p-4 p-lg-5">
        <div class="text-center mb-4">
          <img class="brand-mark brand-mark-large mx-auto" src="notify-logo.svg" alt="NOTIFY logo" />
          <h3 class="mt-3 mb-1">Set Your Password</h3>
          <p class="text-muted">This is your first login. Please set a new password.</p>
        </div>

        <form id="first-login-form" class="auth-form">
          <div class="mb-3">
            <label class="form-label">Email</label>
            <input class="form-control" type="email" value="${userEmail}" disabled />
          </div>
          <div class="mb-3">
            <label class="form-label">New Password</label>
            <input type="password" class="form-control" name="password" placeholder="Enter new password" />
          </div>
          <div class="mb-3">
            <label class="form-label">Confirm Password</label>
            <input type="password" class="form-control" name="confirmPassword" placeholder="Confirm password" />
          </div>
          <div class="d-grid gap-2 mt-4">
            <button class="btn btn-primary w-100" type="submit">Set Password & Login</button>
          </div>
        </form>
      </div>
    </div>
  `;

  const form = document.getElementById('first-login-form');
  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const password = form.elements.password.value;
    const confirmPassword = form.elements.confirmPassword.value;

    if (password !== confirmPassword) {
      alert('Passwords do not match');
      return;
    }

    if (!password) {
      alert('Password is required');
      return;
    }

    try {
      await fetchJson(`${API_BASE}/auth/complete-first-login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: userEmail, newPassword: password })
      });
      renderDashboard();
    } catch (error) {
      alert(error.message || 'Unable to set password');
    }
  });
}

async function renderAdminPanel() {
  if (!currentUser || currentUser.role !== 'admin') {
    renderDashboard();
    return;
  }

  saveAppState('admin', true);

  const users = await fetchJson(`${API_BASE}/admin/users`, adminRequestOptions()).catch(() => []);

  const content = `
    <div class="row mb-4">
      <div class="col">
        <h3 class="mb-3">User Management</h3>
        <div class="card">
          <div class="card-body">
            <h5 class="card-title">Create New Viewer User</h5>
            <form id="create-user-form" class="row g-3">
              <div class="col-md-6">
                <label class="form-label">Full Name</label>
                <input type="text" class="form-control" name="fullName" placeholder="User full name" />
              </div>
              <div class="col-md-6">
                <label class="form-label">Email</label>
                <input type="email" class="form-control" name="email" placeholder="user@example.com" />
              </div>
              <div class="col-12">
                <button type="submit" class="btn btn-primary">Create User</button>
                <div id="create-user-feedback" class="mt-2"></div>
              </div>
            </form>
          </div>
        </div>
      </div>
    </div>

    <div class="row">
      <div class="col">
        <h5 class="mb-3">Active Users</h5>
        <div class="table-responsive">
          <table class="table table-sm">
            <thead>
              <tr>
                <th>Name</th>
                <th>Email</th>
                <th>Role</th>
                <th>First Login</th>
                <th>Created</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              ${users.map(user => `
                <tr>
                  <td>${user.fullName}</td>
                  <td>${user.email}</td>
                  <td>
                    <select class="form-select form-select-sm" data-role-user-id="${user.id}" ${String(user.id) === String(currentUser.id) ? 'disabled' : ''}>
                      <option value="viewer" ${user.role === 'viewer' ? 'selected' : ''}>Viewer</option>
                      <option value="admin" ${user.role === 'admin' ? 'selected' : ''}>Admin</option>
                    </select>
                  </td>
                  <td>${user.isFirstLogin ? '✓ Pending' : '—'}</td>
                  <td>${new Date(user.createdAt).toLocaleDateString()}</td>
                  <td>
                    <button class="btn btn-sm btn-outline-primary" type="button" data-save-role-user-id="${user.id}" ${String(user.id) === String(currentUser.id) ? 'disabled' : ''}>Save role</button>
                    ${String(user.id) !== String(currentUser.id) ? `<button class="btn btn-sm btn-outline-danger" type="button" data-delete-user-id="${user.id}">Delete</button>` : '<span class="text-muted small">Current account</span>'}
                  </td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  document.getElementById('app').innerHTML = renderShell('admin', 'User Management', 'Manage viewer users and permissions', content);

  const form = document.getElementById('create-user-form');
  const feedback = document.getElementById('create-user-feedback');

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    const fullName = form.elements.fullName.value;
    const email = form.elements.email.value;

    if (!email) {
      feedback.innerHTML = '<div class="alert alert-danger">Email is required</div>';
      return;
    }

    try {
      const response = await fetchJson(`${API_BASE}/admin/users`, adminRequestOptions({
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, fullName })
      }));

      const generatedPassword = response.user && response.user.password ? response.user.password : 'Check the user email';
      const emailStatus = response.emailSent
        ? 'The temporary password was emailed to the user.'
        : `The temporary password was not emailed. ${response.emailError || 'Configure Brevo in the backend deployment settings.'}`;
      feedback.innerHTML = `<div class="alert alert-success">User created successfully.<br>Temporary password: <strong>${generatedPassword}</strong><br><small>${emailStatus}</small></div>`;
      form.reset();
      setTimeout(() => renderAdminPanel(), 30 * 60 * 1000);
    } catch (error) {
      feedback.innerHTML = `<div class="alert alert-danger">${error.message}</div>`;
    }
  });

  document.querySelectorAll('[data-delete-user-id]').forEach((button) => {
    button.addEventListener('click', async () => {
      const userId = button.getAttribute('data-delete-user-id');
      if (!userId) return;

      if (!window.confirm('Are you sure you want to delete this user?')) {
        return;
      }

      try {
        await fetchJson(`${API_BASE}/admin/users/${userId}`, adminRequestOptions({ method: 'DELETE' }));
        renderAdminPanel();
      } catch (error) {
        window.alert(error.message || 'Unable to delete the user.');
      }
    });
  });

  document.querySelectorAll('[data-save-role-user-id]').forEach((button) => {
    button.addEventListener('click', async () => {
      const userId = button.getAttribute('data-save-role-user-id');
      const roleSelect = document.querySelector(`[data-role-user-id="${userId}"]`);
      if (!userId || !roleSelect) return;

      try {
        await fetchJson(`${API_BASE}/admin/users/${userId}/role`, adminRequestOptions({
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ role: roleSelect.value })
        }));
        window.alert('User role updated successfully.');
        renderAdminPanel();
      } catch (error) {
        window.alert(error.message || 'Unable to update the user role.');
      }
    });
  });
}

async function renderDashboard() {
  saveAppState('dashboard', true);

  const summary = await fetchJson(`${API_BASE}/dashboard/summary`).catch(() => ({
    totalEvents: 0,
    upcomingEvents: 0,
    outstandingTasks: 0,
    overdueTasks: 0,
    tasksDueSoon: 0,
    notifications: [],
    events: [],
    tasks: []
  }));

  const events = summary.events || [];
  const tasks = summary.tasks || [];
  const alerts = (summary.notifications || []).slice(0, 4);
  const currentEvents = events.slice(0, 4);
  const recentActivity = tasks.slice(0, 4).map(task => ({
    label: task.title || 'Task update',
    meta: `${task.assignee || 'Unassigned'} • ${task.status || 'Pending'}`,
    time: task.dueDate ? formatDate(task.dueDate) : 'No due date'
  }));

  const eventMetricMessage = (summary.upcomingEvents || 0) > 0
    ? `${summary.upcomingEvents} upcoming event${summary.upcomingEvents === 1 ? '' : 's'}`
    : 'No upcoming events';
  const tasksMetricMessage = (summary.tasksDueSoon || 0) > 0
    ? `${summary.tasksDueSoon} task${summary.tasksDueSoon === 1 ? '' : 's'} due soon`
    : 'No tasks due soon';
  const alertsMetricMessage = alerts.length > 0
    ? `${alerts.length} notification${alerts.length === 1 ? '' : 's'} active`
    : 'No active alerts';
  const overdueMetricMessage = (summary.overdueTasks || 0) > 0
    ? `${summary.overdueTasks} item${summary.overdueTasks === 1 ? '' : 's'} need follow-up`
    : 'Everything is on track';

  const content = `
    <section class="dashboard-hero card">
      <div class="d-flex flex-column flex-md-row justify-content-between align-items-start gap-3">
        <div>
          <h3 class="mb-2">NOTIFY admin portal</h3>
          <p class="text-muted mb-0">Access the core operations modules from a simplified control surface.</p>
        </div>
        <button class="btn btn-primary btn-lg dashboard-create-btn" onclick="renderEventForm()">Create Event</button>
      </div>

      <div class="dashboard-metric-grid">
        <div class="metric-card" data-tone="events">
          <div class="metric-card-row">
            <div class="metric-icon metric-icon-events">E</div>
            <div class="metric-label">Events</div>
          </div>
          <div class="metric-value">${summary.totalEvents || 0}</div>
          <div class="metric-note">${eventMetricMessage}</div>
        </div>
        <div class="metric-card" data-tone="tasks">
          <div class="metric-card-row">
            <div class="metric-icon metric-icon-tasks">T</div>
            <div class="metric-label">Tasks</div>
          </div>
          <div class="metric-value">${summary.outstandingTasks || 0}</div>
          <div class="metric-note">${tasksMetricMessage}</div>
        </div>
        <div class="metric-card" data-tone="alerts">
          <div class="metric-card-row">
            <div class="metric-icon metric-icon-alerts">A</div>
            <div class="metric-label">Alerts</div>
          </div>
          <div class="metric-value">${alerts.length}</div>
          <div class="metric-note">${alertsMetricMessage}</div>
        </div>
        <div class="metric-card" data-tone="overdue">
          <div class="metric-card-row">
            <div class="metric-icon metric-icon-overdue">O</div>
            <div class="metric-label">Overdue</div>
          </div>
          <div class="metric-value">${summary.overdueTasks || 0}</div>
          <div class="metric-note">${overdueMetricMessage}</div>
        </div>
      </div>
    </section>

    <section class="dashboard-module-grid">
      <div class="module-card module-card-blue" onclick="navigateToRoute('events')">
        <div class="module-icon module-icon-blue">E</div>
        <div>
          <h5>Events</h5>
          <p>Create, edit, confirm, complete or cancel events.</p>
        </div>
      </div>
      <div class="module-card module-card-purple" onclick="navigateToRoute('tasks')">
        <div class="module-icon module-icon-purple">T</div>
        <div>
          <h5>Tasks</h5>
          <p>Assign tasks, due dates, statuses and progress updates.</p>
        </div>
      </div>
      <div class="module-card module-card-orange" onclick="navigateToRoute('reminders')">
        <div class="module-icon module-icon-orange">N</div>
        <div>
          <h5>Notifications</h5>
          <p>Automated reminders, alerts and in-app notification center.</p>
        </div>
      </div>
      <div class="module-card module-card-teal" onclick="navigateToRoute('calendar')">
        <div class="module-icon module-icon-teal">C</div>
        <div>
          <h5>Calendar</h5>
          <p>Visual view of events, deadlines and overdue work.</p>
        </div>
      </div>
      <div class="module-card module-card-green" onclick="navigateToRoute('attendees')">
        <div class="module-icon module-icon-green">A</div>
        <div>
          <h5>Attendees</h5>
          <p>Invite people and track RSVP status.</p>
        </div>
      </div>
      <div class="module-card module-card-blue" onclick="navigateToRoute('documents')">
        <div class="module-icon module-icon-blue">D</div>
        <div>
          <h5>Documents</h5>
          <p>Store agendas, quotations, programs and reports.</p>
        </div>
      </div>
      <div class="module-card module-card-purple" onclick="navigateToRoute('reports')">
        <div class="module-icon module-icon-purple">R</div>
        <div>
          <h5>Reports</h5>
          <p>Event, task, staff and notification analytics.</p>
        </div>
      </div>
      <div class="module-card module-card-dark" onclick="renderAuditLogs()">
        <div class="module-icon module-icon-dark">L</div>
        <div>
          <h5>Audit Logs</h5>
          <p>Record important administrator actions.</p>
        </div>
      </div>
      <div class="module-card module-card-teal" onclick="navigateToRoute('eventTypes')">
        <div class="module-icon module-icon-teal">ET</div>
        <div>
          <h5>Event Types</h5>
          <p>Manage and route the classifications used in planning and reporting.</p>
        </div>
      </div>
      <div class="module-card module-card-purple" onclick="navigateToRoute('departments')">
        <div class="module-icon module-icon-purple">D</div>
        <div>
          <h5>Departments</h5>
          <p>Track ownership and structure across all operational units.</p>
        </div>
      </div>
    </section>

    <section class="row g-4 mt-1">
      <div class="col-lg-7">
        <div class="card p-4">
          <div class="d-flex justify-content-between align-items-center mb-3">
            <h4 class="mb-0">Current events</h4>
            <button class="btn btn-outline-secondary btn-sm" onclick="navigateToRoute('events')">View all</button>
          </div>
          <div class="table-responsive">
            <table class="table align-middle mb-0">
              <thead>
                <tr><th>Event</th><th>Type</th><th>Date</th><th>Status</th></tr>
              </thead>
              <tbody>
                ${currentEvents.length ? currentEvents.map(item => `
                  <tr>
                    <td>${item.name || 'Untitled event'}</td>
                    <td>${item.type || 'Other'}</td>
                    <td>${formatDate(item.startDate || item.date)}</td>
                    <td>${statusBadge(item.status || 'Planning')}</td>
                  </tr>
                `).join('') : `<tr><td colspan="4" class="text-muted">No events yet. Create one to populate this dashboard.</td></tr>`}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <div class="col-lg-5">
        <div class="card p-4">
          <h4 class="mb-3">Recent activity</h4>
          <div class="list-group list-group-flush">
            ${recentActivity.length ? recentActivity.map(item => `
              <div class="list-group-item px-0">
                <div class="fw-semibold">${item.label}</div>
                <div class="text-muted small">${item.meta}</div>
                <div class="text-muted small">${item.time}</div>
              </div>
            `).join('') : `<div class="text-muted small">No task activity yet.</div>`}
          </div>
        </div>
      </div>
    </section>

    <section class="row g-4 mt-1">
      <div class="col-lg-8">
        <div class="card p-4">
          <h4 class="mb-3">System alerts</h4>
          <ul class="list-group list-group-flush">
            ${alerts.length ? alerts.map(item => `
              <li class="list-group-item px-0">
                <div class="d-flex justify-content-between align-items-start gap-3">
                  <div>
                    <div class="fw-semibold">${item.message || 'System notification'}</div>
                    <div class="text-muted small">${item.channel || 'email'} • ${item.status || 'sent'}</div>
                  </div>
                  ${statusBadge(item.status || 'Sent')}
                </div>
              </li>
            `).join('') : `<li class="list-group-item px-0 text-muted">No active alerts right now.</li>`}
          </ul>
        </div>
      </div>

      <div class="col-lg-4">
        <div class="card p-4">
          <h4 class="mb-3">Quick actions</h4>
          <div class="d-grid gap-2">
            <button class="btn btn-primary" onclick="renderEventForm()">Create Event</button>
            <button class="btn btn-outline-secondary" onclick="renderTaskForm()">Create Task</button>
            <button class="btn btn-outline-secondary" onclick="navigateToRoute('reports')">Open Reports</button>
          </div>
        </div>
      </div>
    </section>
  `;

  document.getElementById('app').innerHTML = renderShell('dashboard', 'Dashboard', 'Access the core operations modules from a simplified control surface.', content);
}

async function renderEvents() {
  saveAppState('events', true);
  const events = await fetchJson(`${API_BASE}/events`);
  const isAdmin = currentUser && currentUser.role === 'admin';
  const content = `
    <div class="card p-4">
      <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <div>
          <h4 class="mb-1">Event Management</h4>
          <p class="text-muted mb-0">Create and review organizational events from planning through completion.</p>
        </div>
        <div class="d-flex gap-2">
          ${isAdmin ? `<button class="btn btn-outline-secondary" onclick="triggerImport('event')">Import Excel</button>` : ''}
          ${isAdmin ? `<button class="btn btn-primary" onclick="renderEventForm()">Create Event</button>` : ''}
        </div>
      </div>
      ${isAdmin ? `<div class="alert alert-secondary mb-3">
        <strong>Import format:</strong> use columns such as Name, Type, Start Date, End Date, Venue, Coordinator, Expected Attendees, Department, Status, and Attendee Name / Attendee Email (or Attendee Names / Attendee Emails) for the guest roster.
      </div>` : ''}
      ${events.length === 0 ? `<div class="alert alert-info mb-3">No events yet. Create the first one to start building your timeline.</div>` : ''}
      <div class="table-responsive">
        <table class="table align-middle">
          <thead><tr><th>Event</th><th>Type</th><th>Start Date</th><th>End Date</th><th>Department</th><th>Status</th>${isAdmin ? '<th>Actions</th>' : ''}</tr></thead>
          <tbody>${events.map(event => `
            <tr>
              <td>${event.name}</td>
              <td>${event.type}</td>
              <td>${formatDate(event.startDate || event.date)}</td>
              <td>${formatDate(event.endDate)}</td>
              <td>${event.department}</td>
              <td>${statusBadge(event.status)}</td>
              ${isAdmin ? `<td class="d-flex gap-2">
                <button class="btn btn-outline-secondary btn-sm" onclick="renderEventForm('${event.id}')">Edit</button>
                <button class="btn btn-outline-danger btn-sm" onclick="deleteEvent('${event.id}')">Delete</button>
              </td>` : ''}
            </tr>
          `).join('')}</tbody>
        </table>
      </div>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('events', 'Events', 'Plan, coordinate and monitor all organizational events.', content, '<button class="btn btn-outline-secondary btn-sm" onclick="renderCalendar()">Open Calendar</button>');
}

async function renderCalendar() {
  saveAppState('calendar', true);
  const events = await fetchJson(`${API_BASE}/events`);
  const calendarDays = getCalendarDays(calendarViewDate);
  const selectedKey = getCalendarDateKey(calendarSelectedDate);
  const selectedEvents = events.filter(event => getEventDateKey(event) === selectedKey);

  const content = `
    <div class="row g-4">
      <div class="col-lg-8">
        <div class="card p-4">
          <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
            <div>
              <h4 class="mb-1">Calendar Overview</h4>
              <p class="text-muted mb-0">Browse by month and see the events you add in the calendar.</p>
            </div>
            <div class="btn-group">
              <button class="btn btn-outline-secondary btn-sm" type="button" onclick="shiftCalendarMonth(-1)">← Prev</button>
              <button class="btn btn-outline-secondary btn-sm" type="button" onclick="shiftCalendarMonth(1)">Next →</button>
            </div>
          </div>
          <div class="calendar-card">
            <div class="calendar-header">${calendarViewDate.toLocaleDateString('en', { month: 'long', year: 'numeric' })}</div>
            <div class="calendar-grid">
              ${['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].map(day => `<div class="calendar-day-label">${day}</div>`).join('')}
              ${calendarDays.map((cell) => {
                const dayKey = getCalendarDateKey(cell.date);
                const dayEvents = events.filter(event => getEventDateKey(event) === dayKey);
                const isSelected = dayKey === selectedKey;
                const isToday = dayKey === getCalendarDateKey(new Date());
                return `
                  <button class="calendar-cell ${cell.isCurrentMonth ? '' : 'calendar-cell-muted'} ${isSelected ? 'selected' : ''} ${dayEvents.length ? 'has-event' : ''} ${isToday ? 'today' : ''}" type="button" onclick="selectCalendarDate('${dayKey}')">
                    <span class="calendar-day-number">${cell.date.getDate()}</span>
                    ${dayEvents.length ? `<span class="calendar-event-pill">${Math.min(dayEvents.length, 2)} event${dayEvents.length > 1 ? 's' : ''}</span>` : ''}
                  </button>
                `;
              }).join('')}
            </div>
          </div>
        </div>
      </div>
      <div class="col-lg-4">
        <div class="card p-4">
          <h5 class="mb-3">Selected day</h5>
          <div class="text-muted small mb-3">${calendarSelectedDate.toLocaleDateString('en', { month: 'long', day: 'numeric', year: 'numeric' })}</div>
          ${selectedEvents.length ? `
            <ul class="list-group list-group-flush">
              ${selectedEvents.map(event => `
                <li class="list-group-item px-0">
                  <div class="fw-semibold">${event.name}</div>
                  <div class="text-muted small">${event.type} • ${event.status}</div>
                </li>
              `).join('')}
            </ul>
          ` : '<div class="alert alert-light border-0 mb-0">No events for this day yet.</div>'}
        </div>
      </div>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('calendar', 'Calendar', 'Track significant dates, task deadlines and major milestones.', content);
}

async function renderTasks() {
  saveAppState('tasks', true);
  const isAdmin = currentUser && currentUser.role === 'admin';
  const tasks = await fetchJson(`${API_BASE}/tasks`);
  const content = `
    <div class="card p-4">
      <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <div>
          <h4 class="mb-1">${isAdmin ? 'Task Management' : 'Task Overview'}</h4>
          <p class="text-muted mb-0">${isAdmin ? 'Assign, monitor and update work for every event.' : 'View current assignments and deadlines without editing permissions.'}</p>
        </div>
        ${isAdmin ? '<button class="btn btn-primary" onclick="renderTaskForm()">Create Task</button>' : ''}
      </div>
      ${!isAdmin ? '<div class="alert alert-info mb-3">View-only access: you can review tasks but cannot create, edit, or delete them.</div>' : ''}
      <div class="table-responsive">
        <table class="table align-middle">
          <thead><tr><th>Task</th><th>Event</th><th>Assignee</th><th>Due Date</th><th>Status</th>${isAdmin ? '<th>Actions</th>' : ''}</tr></thead>
          <tbody>${tasks.map(task => `
            <tr>
              <td>${task.title}</td>
              <td>${task.eventId}</td>
              <td>${task.assignee}</td>
              <td>${formatDate(task.dueDate)}</td>
              <td>${statusBadge(task.status)}</td>
              ${isAdmin ? `<td class="d-flex gap-2">
                <button class="btn btn-outline-secondary btn-sm" onclick="renderTaskForm('${task.id}')">Edit</button>
                <button class="btn btn-outline-danger btn-sm" onclick="deleteTask('${task.id}')">Delete</button>
              </td>` : ''}
            </tr>
          `).join('')}</tbody>
        </table>
      </div>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('tasks', 'Tasks', 'Keep every assignment visible, accountable and on schedule.', content);
}

async function renderAttendees() {
  saveAppState('attendees', true);
  const isAdmin = currentUser && currentUser.role === 'admin';

  let attendees = [];
  try {
    attendees = await fetchJson(`${API_BASE}/attendees`);
  } catch (error) {
    attendees = [];
  }

  const content = `
    <div class="card p-4">
      <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <div>
          <h4 class="mb-1">${isAdmin ? 'Attendee Management' : 'Attendee Directory'}</h4>
          <p class="text-muted mb-0">${isAdmin ? 'Add and update attendee details as registrations change.' : 'View attendee details without editing permissions.'}</p>
        </div>
        ${isAdmin ? `
          <div class="d-flex gap-2">
            <button class="btn btn-outline-secondary" onclick="triggerImport('attendee')">Import Excel</button>
            <button class="btn btn-primary" onclick="renderAttendeeForm()">Add Attendee</button>
          </div>
        ` : ''}
      </div>
      ${isAdmin ? `<div class="alert alert-secondary mb-3"><strong>Import format:</strong> use columns like Name, Email, Status.</div>` : `<div class="alert alert-info mb-3">View-only access: you can review attendee records but cannot create, edit, or delete them.</div>`}
      ${attendees.length === 0 ? `<div class="alert alert-info mb-3">No attendees yet.</div>` : ''}
      <table class="table align-middle">
        <thead><tr><th>Name</th><th>Email</th><th>Event</th><th>Status</th>${isAdmin ? '<th>Actions</th>' : ''}</tr></thead>
        <tbody>${attendees.map(attendee => `
          <tr>
            <td>${attendee.name}</td>
            <td>${attendee.email}</td>
            <td>${attendee.eventName || '—'}</td>
            <td>${statusBadge(attendee.status)}</td>
            ${isAdmin ? `<td class="d-flex gap-2">
              <button class="btn btn-outline-secondary btn-sm" onclick="renderAttendeeForm('${attendee.id}')">Edit</button>
              <button class="btn btn-outline-danger btn-sm" onclick="deleteAttendee('${attendee.id}')">Delete</button>
            </td>` : ''}
          </tr>
        `).join('')}</tbody>
      </table>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('attendees', 'Attendees', 'Track RSVPs, attendance and communication status.', content);
}

async function renderNotifications() {
  saveAppState('reminders', true);
  const notifications = await fetchJson(`${API_BASE}/notifications`);
  const content = `
    <div class="card p-4">
      <h4 class="mb-3">Notification Center</h4>
      <ul class="list-group list-group-flush">
        ${notifications.map(item => `
          <li class="list-group-item px-0">
            <div class="d-flex justify-content-between align-items-start gap-3">
              <div>
                <div class="fw-semibold">${item.message}</div>
                <div class="text-muted small">${item.channel} • ${item.status}</div>
              </div>
              <div class="d-flex align-items-center gap-2">
                ${statusBadge(item.status)}
                <button class="btn btn-outline-danger btn-sm" onclick="deleteNotification('${item.id}')">Delete</button>
              </div>
            </div>
          </li>
        `).join('')}
      </ul>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('notifications', 'Notifications', 'Keep every reminder and alert visible.', content);
}

async function renderDocuments() {
  saveAppState('documents', true);
  const isAdmin = currentUser && currentUser.role === 'admin';
  const documents = await fetchJson(`${API_BASE}/documents`).catch(() => []);
  const content = `
    <div class="card p-4">
      <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <div>
          <h4 class="mb-1">Event Documents</h4>
          <p class="text-muted mb-0">Upload supporting files and export event records in spreadsheet format.</p>
        </div>
        ${isAdmin ? `
          <div class="d-flex gap-2">
            <button class="btn btn-outline-secondary" onclick="handleDocumentUpload()">Upload PDF / Excel</button>
            <button class="btn btn-primary" onclick="exportReportsToExcel()">Export Excel</button>
          </div>
        ` : ''}
      </div>
      ${!isAdmin ? `<div class="alert alert-info mb-3">View-only access: you can review document files but cannot upload or delete them.</div>` : ''}
      <div class="row g-3">
        ${documents.map(doc => `
          <div class="col-md-6">
            <div class="document-card p-3">
              <div class="d-flex justify-content-between align-items-start gap-2">
                <div>
                  <div class="fw-semibold">${doc.name}</div>
                  <div class="text-muted small">${doc.type || 'Document'} • Uploaded ${formatDate(doc.uploadedAt)}</div>
                </div>
                <div class="d-flex gap-2">
                  ${doc.dataUrl ? `<button class="btn btn-outline-secondary btn-sm" onclick="downloadDocument('${doc.id}')">Download</button>` : ''}
                  ${isAdmin ? `<button class="btn btn-outline-danger btn-sm" onclick="deleteDocument('${doc.id}')">Delete</button>` : ''}
                </div>
              </div>
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('documents', 'Documents', 'Store supporting materials for each event.', content);
}

async function handleDocumentUpload() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.pdf,.xlsx,.xls,.csv,.doc,.docx';
  input.onchange = async () => {
    const file = input.files[0];
    if (!file) return;

    try {
      const fileData = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('Unable to read document file.'));
        reader.readAsDataURL(file);
      });

      const payload = {
        name: file.name,
        type: file.type || 'application/octet-stream',
        uploadedAt: new Date().toISOString(),
        size: file.size,
        dataUrl: fileData
      };

      await fetchJson(`${API_BASE}/documents`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      renderDocuments();
    } catch (error) {
      window.alert('Unable to upload the selected file.');
    }
  };
  input.click();
}

function downloadDocument(documentId) {
  const doc = window.__notifyDocuments ? window.__notifyDocuments.find(item => item.id === documentId) : null;
  if (!doc || !doc.dataUrl) {
    window.alert('This document is not available for download.');
    return;
  }
  const link = document.createElement('a');
  link.href = doc.dataUrl;
  link.download = doc.name || 'document';
  document.body.appendChild(link);
  link.click();
  link.remove();
}

window.__notifyDocuments = [];

async function refreshDocumentCache() {
  try {
    window.__notifyDocuments = await fetchJson(`${API_BASE}/documents`);
  } catch (error) {
    window.__notifyDocuments = [];
  }
}

(async function initializeDocumentCache() {
  await refreshDocumentCache();
})();

async function downloadWorkbook(filename, workbook) {
  try {
    if (typeof XLSX.writeFile === 'function') {
      XLSX.writeFile(workbook, filename);
      return;
    }

    const wopts = { bookType: 'xlsx', type: 'array' };
    const wbout = XLSX.write(workbook, wopts);
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  } catch (error) {
    console.error('Excel download failed', error);
    window.alert('The Excel report could not be downloaded. Please try again.');
  }
}

function buildReportsWorkbook(events = [], attendees = [], tasks = []) {
  const workbook = XLSX.utils.book_new();
  const rows = [
    ['NOTIFY Report'],
    [],
    ['Generated On', new Date().toISOString().slice(0, 10)],
    [],
    ['Events'],
    ['Name', 'Type', 'Date', 'Department', 'Status'],
    ...events.map(event => [event.name || '', event.type || '', event.startDate || event.date || '', event.department || '', event.status || '']),
    [],
    ['Attendees'],
    ['Name', 'Email', 'Role', 'Status'],
    ...attendees.map(attendee => [
      attendee.name || '',
      attendee.email || '',
      attendee.role || attendee.directorRole || '',
      attendee.status || ''
    ]),
    [],
    ['Tasks'],
    ['Title', 'Assignee', 'Due Date', 'Status', 'Event'],
    ...tasks.map(task => [task.title || '', task.assignee || '', task.dueDate || '', task.status || '', task.eventId || ''])
  ];

  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols'] = [
    { wch: 28 },
    { wch: 24 },
    { wch: 16 },
    { wch: 16 },
    { wch: 20 }
  ];

  XLSX.utils.book_append_sheet(workbook, sheet, 'NOTIFY Report');
  return workbook;
}

async function exportReportsToExcel() {
  try {
    const [events, attendees, tasks] = await Promise.all([
      fetchJson(`${API_BASE}/events`),
      fetchJson(`${API_BASE}/attendees`),
      fetchJson(`${API_BASE}/tasks`)
    ]);

    const workbook = buildReportsWorkbook(events, attendees, tasks);
    const filename = `notify-report-${new Date().toISOString().slice(0, 10)}.xlsx`;
    await downloadWorkbook(filename, workbook);
  } catch (error) {
    console.error('Report export failed', error);
    window.alert('Unable to generate the Excel report right now. Please check the connection and try again.');
  }
}

async function renderReports() {
  saveAppState('reports', true);
  const data = await fetchJson(`${API_BASE}/dashboard/summary`);
  const content = `
    <div class="d-flex justify-content-between align-items-center mb-4">
      <div>
        <h4 class="mb-1">Enterprise Reports</h4>
        <p class="text-muted mb-0">Download event, attendee and task details in Excel format for executive review.</p>
      </div>
      <button class="btn btn-primary" onclick="exportReportsToExcel()">Download Excel Report</button>
    </div>
    <div class="row g-4">
      <div class="col-md-6">
        <div class="card p-4">
          <h5>Event Report</h5>
          <p class="text-muted mb-0">${data.totalEvents} total events with ${data.planningEvents} planning, ${data.confirmedEvents} confirmed and ${data.completedEvents} completed.</p>
        </div>
      </div>
      <div class="col-md-6">
        <div class="card p-4">
          <h5>Task Report</h5>
          <p class="text-muted mb-0">${data.outstandingTasks} outstanding tasks and ${data.overdueTasks} overdue follow-ups.</p>
        </div>
      </div>
      <div class="col-md-6">
        <div class="card p-4">
          <h5>Notification Report</h5>
          <p class="text-muted mb-0">${(data.notifications || []).length} recent system notifications available for review.</p>
        </div>
      </div>
      <div class="col-md-6">
        <div class="card p-4">
          <h5>Department Summary</h5>
          <p class="text-muted mb-0">${[...new Set((data.events || []).map(event => event.department))].join(', ')}</p>
        </div>
      </div>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('reports', 'Reports', 'Analyse performance, activity and operational health.', content);
}

function renderDepartments() {
  saveAppState('departments', true);
  const content = `
    <div class="card p-4">
      <h4 class="mb-3">Departments</h4>
      <div class="row g-3">
        ${['Administration','IT','Finance','Human Resources','Procurement','Operations','Executive'].map(item => `
          <div class="col-md-6 col-lg-4">
            <div class="document-card p-3">
              <div class="fw-semibold">${item}</div>
              <div class="text-muted small">Managed centrally from the administrator console.</div>
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('departments', 'Departments', 'Maintain the organizational structure behind every event.', content);
}

function renderEventTypes() {
  saveAppState('eventTypes', true);
  const eventTypes = ['Meeting','Training','Workshop','Conference','Seminar','Staff Event','Community Event','Launch','Other'];
  const content = `
    <div class="card p-4">
      <h4 class="mb-3">Event Types</h4>
      <div class="row g-3">
        ${eventTypes.map(item => `
          <div class="col-md-6 col-lg-4">
            <div class="document-card p-3 event-type-card" onclick="renderEventTypeDetail('${item}')" style="cursor:pointer; transition: transform 0.2s ease, box-shadow 0.2s ease;">
              <div class="fw-semibold">${item}</div>
              <div class="text-muted small">Ready for planning and reporting workflows.</div>
            </div>
          </div>
        `).join('')}
      </div>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('eventTypes', 'Event Types', 'Classify events for reporting, filtering and planning.', content);
}

async function renderEventTypeDetail(typeName) {
  const events = await fetchJson(`${API_BASE}/events`);
  const relatedEvents = events.filter(event => String(event.type || '').toLowerCase() === String(typeName || '').toLowerCase());

  const activities = relatedEvents.length
    ? relatedEvents.map(event => `${event.name || 'Untitled event'} • ${formatDate(event.startDate || event.date)} • ${event.status || 'Planning'}`)
    : [`No events created for ${typeName} yet.`];

  const content = `
    <div class="card p-4">
      <div class="d-flex justify-content-between align-items-center mb-3 gap-3 flex-wrap">
        <div>
          <p class="eyebrow">Event Type</p>
          <h4 class="mb-0">${typeName}</h4>
        </div>
        <div class="d-flex gap-2 flex-wrap">
          <button class="btn btn-primary" onclick="renderEventForm(null, '${typeName}')">Create event for this type</button>
          <button class="btn btn-outline-secondary" onclick="renderEventTypes()">Back to Event Types</button>
        </div>
      </div>

      <div class="row g-3">
        <div class="col-lg-7">
          <div class="p-3 border rounded-4 bg-light">
            <h5 class="mb-3">Activities</h5>
            <ul class="list-group list-group-flush">
              ${activities.map(item => `
                <li class="list-group-item px-0">
                  <div class="d-flex align-items-center gap-2">
                    <span class="badge rounded-pill text-bg-primary">•</span>
                    <span>${item}</span>
                  </div>
                </li>
              `).join('')}
            </ul>
          </div>
        </div>

        <div class="col-lg-5">
          <div class="p-3 border rounded-4 bg-white">
            <h5 class="mb-3">Events in this category</h5>
            ${relatedEvents.length ? `
              <div class="list-group list-group-flush">
                ${relatedEvents.map(event => `
                  <div class="list-group-item px-0">
                    <div class="fw-semibold">${event.name || 'Untitled event'}</div>
                    <div class="text-muted small">${formatDate(event.startDate || event.date)} • ${event.status || 'Planning'}</div>
                  </div>
                `).join('')}
              </div>
            ` : '<div class="text-muted">No events are currently assigned to this type.</div>'}
          </div>
        </div>
      </div>
    </div>
  `;

  document.getElementById('app').innerHTML = renderShell('eventTypes', typeName, 'Event activities and planning items for this category.', content);
}

async function renderAuditLogs() {
  const logs = await fetchJson(`${API_BASE}/audit-logs`);
  const content = `
    <div class="card p-4">
      <h4 class="mb-3">Audit Trail</h4>
      <table class="table align-middle">
        <thead><tr><th>Action</th><th>Performed By</th><th>Date</th></tr></thead>
        <tbody>${logs.map(log => `
          <tr>
            <td>${log.action}</td>
            <td>${log.performedBy}</td>
            <td>${formatDate(log.createdAt)}</td>
          </tr>
        `).join('')}</tbody>
      </table>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('audit', 'Audit Logs', 'Review the administrator history for every key action.', content);
}

function renderSettings() {
  const content = `
    <div class="card p-4">
      <h4 class="mb-3">System Settings</h4>
      <div class="row g-3">
        <div class="col-md-6">
          <label class="form-label">Organization Name</label>
          <input class="form-control" value="CARE NOTIFY" />
        </div>
        <div class="col-md-6">
          <label class="form-label">Notification Email</label>
          <input class="form-control" value="notify@care.org" />
        </div>
        <div class="col-md-6">
          <label class="form-label">Default Reminder Intervals</label>
          <input class="form-control" value="7 days, 3 days, 1 day" />
        </div>
        <div class="col-md-6">
          <label class="form-label">Time Zone</label>
          <input class="form-control" value="Africa/Nairobi" />
        </div>
      </div>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('settings', 'Settings', 'Configure the operating environment for the platform.', content);
}

async function renderEventForm(eventId = null, defaultType = '') {
  const isAdmin = currentUser && currentUser.role === 'admin';
  if (!isAdmin) {
    const content = `
      <div class="card p-4">
        <div class="alert alert-warning mb-0">
          <strong>Access Restricted</strong> — Only administrators can create or edit events.
        </div>
      </div>
    `;
    document.getElementById('app').innerHTML = renderShell('events', 'Events', 'Plan, coordinate and monitor all organizational events.', content);
    return;
  }
  
  saveAppState('events', true);
  let initialData = {};
  if (eventId) {
    const events = await fetchJson(`${API_BASE}/events`);
    initialData = events.find(item => item.id === eventId) || {};
  }

  const selectedType = initialData.type || defaultType || 'Conference';
  const departmentOptions = ['Administration','IT','Finance','Human Resources','Procurement','Operations','Executive'];

  const attendeeRows = (initialData.attendees && initialData.attendees.length ? initialData.attendees : [{ name: '', email: '', status: 'Confirmed' }]).map((attendee, index) => `
    <div class="row g-2 attendee-row align-items-end" data-row-index="${index}">
      <div class="col-md-3">
        <label class="form-label">Name</label>
        <input class="form-control" name="attendeeName[]" value="${attendee.name || ''}" />
      </div>
      <div class="col-md-3">
        <label class="form-label">Email</label>
        <input type="email" class="form-control" name="attendeeEmail[]" value="${attendee.email || ''}" />
      </div>
      <div class="col-md-3">
        <label class="form-label">Director role</label>
        <select class="form-select" name="attendeeRole[]">
          <option value="" ${!(attendee.directorRole || attendee.role) ? 'selected' : ''}>-- None --</option>
          <option value="CD" ${String(attendee.directorRole || attendee.role || '') === 'CD' ? 'selected' : ''}>Country Director (CD)</option>
          <option value="RD" ${String(attendee.directorRole || attendee.role || '') === 'RD' ? 'selected' : ''}>Regional Director (RD)</option>
          <option value="Both" ${String(attendee.directorRole || attendee.role || '') === 'Both' ? 'selected' : ''}>Both</option>
        </select>
      </div>
      <div class="col-md-2">
        <label class="form-label">Status</label>
        <select class="form-select" name="attendeeStatus[]">
          ${['Confirmed','Pending','Cancelled'].map(option => `<option value="${option}" ${String(attendee.status || 'Confirmed') === option ? 'selected' : ''}>${option}</option>`).join('')}
        </select>
      </div>
      <div class="col-md-1 d-flex justify-content-end">
        <button type="button" class="btn btn-outline-danger btn-sm remove-attendee-row" data-row-index="${index}">Remove</button>
      </div>
    </div>
  `).join('');

  function appendAttendeeRow({ name = '', email = '', role = '', status = 'Confirmed' } = {}) {
    const row = document.createElement('div');
    row.className = 'row g-2 attendee-row align-items-end';
    row.innerHTML = `
      <div class="col-md-3">
        <label class="form-label">Name</label>
        <input class="form-control" name="attendeeName[]" value="${name}" />
      </div>
      <div class="col-md-3">
        <label class="form-label">Email</label>
        <input type="email" class="form-control" name="attendeeEmail[]" value="${email}" />
      </div>
      <div class="col-md-3">
        <label class="form-label">Director role</label>
        <select class="form-select" name="attendeeRole[]">
          <option value="" ${!role ? 'selected' : ''}>-- None --</option>
          <option value="CD" ${role === 'CD' ? 'selected' : ''}>Country Director (CD)</option>
          <option value="RD" ${role === 'RD' ? 'selected' : ''}>Regional Director (RD)</option>
          <option value="Both" ${role === 'Both' ? 'selected' : ''}>Both</option>
        </select>
      </div>
      <div class="col-md-2">
        <label class="form-label">Status</label>
        <select class="form-select" name="attendeeStatus[]">
          ${['Confirmed','Pending','Cancelled'].map(option => `<option value="${option}" ${status === option ? 'selected' : ''}>${option}</option>`).join('')}
        </select>
      </div>
      <div class="col-md-1 d-flex justify-content-end">
        <button type="button" class="btn btn-outline-danger btn-sm remove-attendee-row">Remove</button>
      </div>
    `;
    row.querySelector('.remove-attendee-row').addEventListener('click', () => row.remove());
    attendeeList.appendChild(row);
  }

  const content = `
    <div class="card p-4">
      <div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
        <div>
          <h4 class="mb-1">${eventId ? 'Edit Event' : 'Create Event'}</h4>
          <p class="text-muted mb-0">Keep the event brief and easy to manage.</p>
        </div>
      </div>
      <form id="event-form">
        <div class="row g-3">
          <div class="col-md-6">
            <label class="form-label">Event Name</label>
            <input class="form-control" name="name" value="${initialData.name || ''}" required />
          </div>
          <div class="col-md-6">
            <label class="form-label">Event Type</label>
            <select class="form-select" name="type" required>
              ${['Conference','Workshop','Training','Meeting','Seminar','Stakeholder Engagement','Donor Visit','Team Building','Community Event','Webinar','Other'].map(option => `<option value="${option}" ${selectedType === option ? 'selected' : ''}>${option}</option>`).join('')}
            </select>
          </div>
          <div class="col-md-4">
            <label class="form-label">Start Date</label>
            <input type="date" class="form-control" name="startDate" value="${initialData.startDate || initialData.date || ''}" required />
          </div>
          <div class="col-md-4">
            <label class="form-label">End Date</label>
            <input type="date" class="form-control" name="endDate" value="${initialData.endDate || ''}" />
          </div>
          <div class="col-md-4">
            <label class="form-label">Status</label>
            <select class="form-select" name="status" required>
              ${['Planning','Confirmed','Completed','Cancelled'].map(option => `<option value="${option}" ${initialData.status === option ? 'selected' : ''}>${option}</option>`).join('')}
            </select>
          </div>
          <div class="col-md-6">
            <label class="form-label">Venue</label>
            <input class="form-control" name="venue" value="${initialData.venue || ''}" />
          </div>
          <div class="col-md-6">
            <label class="form-label">Coordinator</label>
            <input class="form-control" name="coordinator" value="${initialData.coordinator || ''}" />
          </div>
          <div class="col-md-6">
            <label class="form-label">Expected Attendees</label>
            <input type="number" min="1" class="form-control" name="expectedAttendees" value="${initialData.expectedAttendees || ''}" />
          </div>
          <div class="col-md-6">
            <label class="form-label">Department</label>
            <select class="form-select" name="department">
              ${departmentOptions.map(option => `<option value="${option}" ${String(initialData.department || '') === option ? 'selected' : ''}>${option}</option>`).join('')}
            </select>
          </div>
          <div class="col-12">
            <div class="d-flex justify-content-between align-items-center mb-2 flex-wrap gap-2">
              <h5 class="mb-0">Attendee List</h5>
              <div class="d-flex gap-2 flex-wrap">
                <button type="button" class="btn btn-outline-primary btn-sm" id="add-attendee-row">Add attendee</button>
                <button type="button" class="btn btn-outline-secondary btn-sm" id="import-attendee-excel">Import Excel</button>
              </div>
            </div>
            <div class="small text-muted mb-2">
              Use columns like Name, Email, Director role, Status, or Attendee Name / Attendee Email in your Excel sheet.
            </div>
            <div id="event-attendee-list" class="d-grid gap-3">
              ${attendeeRows}
            </div>
          </div>
        </div>
        <button class="btn btn-primary mt-3" type="submit">${eventId ? 'Update Event' : 'Save Event'}</button>
      </form>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('events', eventId ? 'Edit Event' : 'Create Event', 'Add a new event and trigger the reminder workflow instantly.', content);

  const addAttendeeButton = document.getElementById('add-attendee-row');
  const attendeeList = document.getElementById('event-attendee-list');

  addAttendeeButton?.addEventListener('click', () => {
    appendAttendeeRow();
  });

  document.getElementById('import-attendee-excel')?.addEventListener('click', () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls,.csv';
    input.onchange = async () => {
      const file = input.files[0];
      if (!file) return;

      try {
        const rows = await parseSpreadsheetFile(file);
        const imported = [];

        rows.forEach((row) => {
          const cleaned = Object.fromEntries(Object.entries(row).map(([key, value]) => [key, normalizeCellValue(value)]));
          const name = getMappedValue(cleaned, ['name','fullname','fullName','attendeename','guestname','participantname','nameofattendee']) || '';
          const email = getMappedValue(cleaned, ['email','emailaddress','attendeeemail','guestemail','participantemail']) || '';
          const role = getMappedValue(cleaned, ['directorrole','role','attendeerole','directorshiprole']) || '';
          const status = getMappedValue(cleaned, ['status','attendancestatus','attendeestatus']) || 'Confirmed';
          const attendeeNames = splitAttendeeValues(getMappedValue(cleaned, ['attendeenames','guestnames','participants','names']));
          const attendeeEmails = splitAttendeeValues(getMappedValue(cleaned, ['attendeeemails','guestemails','participantemails','emails']));

          if (name || email || attendeeNames.length || attendeeEmails.length) {
            if (name && email) {
              imported.push({ name, email, role, status });
              return;
            }

            if (attendeeNames.length || attendeeEmails.length) {
              const maxEntries = Math.max(attendeeNames.length, attendeeEmails.length, 0);
              for (let index = 0; index < maxEntries; index++) {
                const entryName = attendeeNames[index] || '';
                const entryEmail = attendeeEmails[index] || '';
                if (entryName || entryEmail) {
                  imported.push({ name: entryName, email: entryEmail, role, status });
                }
              }
              return;
            }

            if (name || email) {
              imported.push({ name, email, role, status });
            }
          }
        });

        if (!imported.length) {
          window.alert('No attendee rows were found in the selected file.');
          return;
        }

        attendeeList.innerHTML = '';
        imported.forEach((entry) => appendAttendeeRow(entry));
        window.alert(`${imported.length} attendee${imported.length === 1 ? '' : 's'} imported.`);
      } catch (error) {
        window.alert('Unable to import attendees from the selected file.');
      }
    };
    input.click();
  });

  attendeeList?.querySelectorAll('.remove-attendee-row').forEach((button) => {
    button.addEventListener('click', (event) => {
      const row = event.target.closest('.attendee-row');
      if (row) {
        row.remove();
      }
    });
  });

  document.getElementById('event-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const attendeeNames = formData.getAll('attendeeName[]');
    const attendeeEmails = formData.getAll('attendeeEmail[]');
    const attendeeRoles = formData.getAll('attendeeRole[]');
    const attendeeStatuses = formData.getAll('attendeeStatus[]');
    const attendees = attendeeNames.map((name, index) => {
      const normalizedName = String(name || '').trim();
      const normalizedEmail = String(attendeeEmails[index] || '').trim();
      const normalizedRole = String(attendeeRoles[index] || '').trim();
      const normalizedStatus = String(attendeeStatuses[index] || 'Confirmed').trim();
      if (!normalizedName && !normalizedEmail) {
        return null;
      }
      return {
        name: normalizedName,
        email: normalizedEmail,
        role: normalizedRole || '',
        status: normalizedStatus || 'Confirmed'
      };
    }).filter(Boolean);

    const payload = {
      name: formData.get('name'),
      type: formData.get('type'),
      date: formData.get('startDate'),
      startDate: formData.get('startDate'),
      endDate: formData.get('endDate'),
      venue: formData.get('venue'),
      coordinator: formData.get('coordinator'),
      expectedAttendees: formData.get('expectedAttendees'),
      department: formData.get('department'),
      status: formData.get('status'),
      attendees
    };

    if (eventId) {
      await fetchJson(`${API_BASE}/events/${eventId}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    } else {
      await fetchJson(`${API_BASE}/events`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    }
    if (defaultType) {
      renderEventTypeDetail(defaultType);
      return;
    }
    renderEvents();
  });
}

function renderTaskForm(taskId = null) {
  const isAdmin = currentUser && currentUser.role === 'admin';
  if (!isAdmin) {
    const content = `
      <div class="card p-4">
        <div class="alert alert-warning mb-0">
          <strong>Access Restricted</strong> — Only administrators can create or edit tasks.
        </div>
      </div>
    `;
    document.getElementById('app').innerHTML = renderShell('tasks', 'Tasks', 'Keep every assignment visible, accountable and on schedule.', content);
    return;
  }

  const content = `
    <div class="card p-4">
      <h4 class="mb-3">${taskId ? 'Edit Task' : 'Create Task'}</h4>
      <form id="task-form">
        <div class="row g-3">
          <div class="col-md-6"><label class="form-label">Title</label><input class="form-control" name="title" /></div>
          <div class="col-md-6"><label class="form-label">Assignee</label><input class="form-control" name="assignee" /></div>
          <div class="col-md-6"><label class="form-label">Due Date</label><input type="date" class="form-control" name="dueDate" /></div>
          <div class="col-md-6"><label class="form-label">Status</label><select class="form-select" name="status"><option>Pending</option><option>In Progress</option><option>Completed</option><option>Cancelled</option></select></div>
          <div class="col-12"><label class="form-label">Event ID</label><input class="form-control" name="eventId" /></div>
        </div>
        <button class="btn btn-primary mt-3" type="submit">${taskId ? 'Update Task' : 'Save Task'}</button>
      </form>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('tasks', taskId ? 'Edit Task' : 'Create Task', 'Assign work to staff and set the reminder cadence.', content);

  document.getElementById('task-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const payload = Object.fromEntries(formData.entries());
    if (taskId) {
      await fetchJson(`${API_BASE}/tasks/${taskId}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    } else {
      await fetchJson(`${API_BASE}/tasks`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    }
    renderTasks();
  });
}

async function triggerImport(type) {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xls,.csv';
  input.onchange = async () => {
    const file = input.files[0];
    if (!file) return;
    try {
      const result = await importSpreadsheetRows({ type, file });
      const target = type === 'attendee' ? renderAttendees : renderEvents;
      await target();
      window.alert(`${result.imported} ${type === 'attendee' ? 'attendee' : 'event'}${result.imported === 1 ? '' : 's'} imported${result.skipped ? `, ${result.skipped} skipped` : ''}.`);
    } catch (error) {
      window.alert('Unable to import the selected file.');
    }
  };
  input.click();
}

function renderAttendeeForm(attendeeId = null) {
  const isAdmin = currentUser && currentUser.role === 'admin';
  if (!isAdmin) {
    const content = `
      <div class="card p-4">
        <div class="alert alert-warning mb-0">
          <strong>Access Restricted</strong> — Only administrators can manage attendees.
        </div>
      </div>
    `;
    document.getElementById('app').innerHTML = renderShell('attendees', 'Attendees', 'Track RSVPs, attendance and communication status.', content);
    return;
  }
  
  saveAppState('attendees', true);
  const content = `
    <div class="card p-4">
      <h4 class="mb-3">${attendeeId ? 'Edit Attendee' : 'Add Attendee'}</h4>
      <form id="attendee-form">
        <div class="row g-3">
          <div class="col-md-4"><label class="form-label">Name</label><input class="form-control" name="name" /></div>
          <div class="col-md-4"><label class="form-label">Email</label><input type="email" class="form-control" name="email" /></div>
          <div class="col-md-4"><label class="form-label">Status</label><select class="form-select" name="status"><option>Confirmed</option><option>Pending</option><option>Cancelled</option></select></div>
        </div>
        <button class="btn btn-primary mt-3" type="submit">${attendeeId ? 'Update Attendee' : 'Save Attendee'}</button>
      </form>
    </div>
  `;
  document.getElementById('app').innerHTML = renderShell('attendees', attendeeId ? 'Edit Attendee' : 'Add Attendee', 'Add and update attendee details as registrations change.', content);

  document.getElementById('attendee-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const formData = new FormData(event.target);
    const payload = Object.fromEntries(formData.entries());
    if (attendeeId) {
      await fetchJson(`${API_BASE}/attendees/${attendeeId}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    } else {
      await fetchJson(`${API_BASE}/attendees`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    }
    renderAttendees();
  });
}

async function deleteNotification(notificationId) {
  if (!window.confirm('Delete this notification? This action cannot be undone.')) {
    return;
  }
  await fetchJson(`${API_BASE}/notifications/${notificationId}`, { method: 'DELETE' });
  renderNotifications();
}

async function deleteDocument(documentId) {
  if (!window.confirm('Delete this document? This action cannot be undone.')) {
    return;
  }
  await fetchJson(`${API_BASE}/documents/${documentId}`, { method: 'DELETE' });
  renderDocuments();
}

async function deleteEvent(eventId) {
  const isAdmin = currentUser && currentUser.role === 'admin';
  if (!isAdmin) {
    alert('Only administrators can delete events.');
    return;
  }
  if (!window.confirm('Delete this event? This action cannot be undone.')) {
    return;
  }
  await fetchJson(`${API_BASE}/events/${eventId}`, { method: 'DELETE' });
  renderEvents();
}

async function deleteTask(taskId) {
  const isAdmin = currentUser && currentUser.role === 'admin';
  if (!isAdmin) {
    alert('Only administrators can delete tasks.');
    return;
  }
  if (!window.confirm('Delete this task? This action cannot be undone.')) {
    return;
  }
  await fetchJson(`${API_BASE}/tasks/${taskId}`, { method: 'DELETE' });
  renderTasks();
}

async function deleteAttendee(attendeeId) {
  const isAdmin = currentUser && currentUser.role === 'admin';
  if (!isAdmin) {
    alert('Only administrators can delete attendees.');
    return;
  }
  if (!window.confirm('Delete this attendee? This action cannot be undone.')) {
    return;
  }
  await fetchJson(`${API_BASE}/attendees/${attendeeId}`, { method: 'DELETE' });
  renderAttendees();
}

window.addEventListener('DOMContentLoaded', () => {
  const hashRoute = window.location.hash.replace(/^#\/?/, '').trim();
  const savedState = restoreAppState();
  const validRoutes = ['dashboard','events','calendar','tasks','attendees','reminders','documents','reports','departments','eventTypes','admin'];
  const hasSavedUser = savedState.authenticated && savedState.user && savedState.view;

  try {
    const startupRender = hashRoute && validRoutes.includes(hashRoute)
      ? navigateToRoute(hashRoute)
      : hasSavedUser
        ? navigateToSavedView(savedState.view)
        : renderLogin();

    Promise.resolve(startupRender).catch((error) => {
      console.error('Unable to load NOTIFY', error);
      clearAppState();
      renderLogin();
    });
  } catch (error) {
    console.error('Unable to load NOTIFY', error);
    clearAppState();
    renderLogin();
  }
});

window.addEventListener('hashchange', () => {
  handleRouteFromHash();
});
