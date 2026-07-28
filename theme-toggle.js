const THEME_STORAGE_KEY = 'qrtimeclock-theme';
const LIGHT_THEME = 'light';
const DARK_THEME = 'dark';

const lightThemeCss = `
html[data-theme="light"] {
  color-scheme: light;
  --bg: #eef4fb;
  --bg-soft: #f7faff;
  --card: rgba(255, 255, 255, 0.9);
  --line: rgba(15, 23, 42, 0.14);
  --text: #102033;
  --muted: #53657c;
  --accent: #1677d2;
  --accent-2: #39a9e8;
  --good: #148a4b;
  --warn: #a86500;
  --danger: #c93d3d;
  --shadow: 0 18px 45px rgba(31, 64, 104, 0.14);
}

html[data-theme="light"] body {
  background: radial-gradient(circle at top, #d8ebff 0%, #eef4fb 45%, #e7eef7 100%);
}

html[data-theme="light"] .orb-1 { background: #75c8ff; opacity: 0.22; }
html[data-theme="light"] .orb-2 { background: #8caeff; opacity: 0.18; }
html[data-theme="light"] .card { backdrop-filter: blur(14px); }
html[data-theme="light"] input,
html[data-theme="light"] select {
  background: rgba(255, 255, 255, 0.92);
  border-color: rgba(15, 23, 42, 0.16);
  color: var(--text);
}
html[data-theme="light"] input::placeholder { color: rgba(42, 58, 78, 0.52); }
html[data-theme="light"] .secondary-btn {
  background: rgba(22, 119, 210, 0.1);
  border-color: rgba(22, 119, 210, 0.28);
}
html[data-theme="light"] .ghost-btn,
html[data-theme="light"] .button-link {
  border-color: rgba(15, 23, 42, 0.18);
}
html[data-theme="light"] .danger-btn {
  background: rgba(201, 61, 61, 0.1);
  color: #a92d2d;
  border-color: rgba(201, 61, 61, 0.25);
}
html[data-theme="light"] .session-chip,
html[data-theme="light"] .tabs,
html[data-theme="light"] .stat-card,
html[data-theme="light"] code {
  background: rgba(255, 255, 255, 0.65);
}
html[data-theme="light"] .pill,
html[data-theme="light"] #appVersion {
  color: #145b95;
}
html[data-theme="light"] .proprietary-notice {
  background: rgba(255, 202, 87, 0.2);
  color: #704a00;
}
html[data-theme="light"] table,
html[data-theme="light"] th,
html[data-theme="light"] td {
  border-color: rgba(15, 23, 42, 0.12);
}
html[data-theme="light"] th { color: #31445a; }
html[data-theme="light"] .toast { box-shadow: 0 14px 35px rgba(31, 64, 104, 0.2); }

.theme-toggle-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  min-width: 118px;
  white-space: nowrap;
}

.theme-toggle-icon {
  font-size: 1rem;
  line-height: 1;
}

@media (max-width: 700px) {
  .topbar { align-items: flex-start; flex-wrap: wrap; }
  .theme-toggle-btn { min-width: 0; }
}
`;

function getSavedTheme() {
  try {
    return localStorage.getItem(THEME_STORAGE_KEY) === LIGHT_THEME ? LIGHT_THEME : DARK_THEME;
  } catch {
    return DARK_THEME;
  }
}

function saveTheme(theme) {
  try {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  } catch {
    // Theme still works for the current page if storage is unavailable.
  }
}

function applyTheme(theme) {
  const normalizedTheme = theme === LIGHT_THEME ? LIGHT_THEME : DARK_THEME;
  document.documentElement.dataset.theme = normalizedTheme;
  document.documentElement.style.colorScheme = normalizedTheme;

  const button = document.getElementById('themeToggleBtn');
  if (button) {
    const nextTheme = normalizedTheme === DARK_THEME ? LIGHT_THEME : DARK_THEME;
    button.setAttribute('aria-label', `Switch to ${nextTheme} mode`);
    button.setAttribute('title', `Switch to ${nextTheme} mode`);
    button.innerHTML = normalizedTheme === DARK_THEME
      ? '<span class="theme-toggle-icon" aria-hidden="true">☀️</span><span>Light mode</span>'
      : '<span class="theme-toggle-icon" aria-hidden="true">🌙</span><span>Dark mode</span>';
  }
}

function installThemeStyles() {
  if (document.getElementById('qrTimeclockThemeStyles')) return;
  const style = document.createElement('style');
  style.id = 'qrTimeclockThemeStyles';
  style.textContent = lightThemeCss;
  document.head.appendChild(style);
}

function installThemeToggle() {
  if (document.getElementById('themeToggleBtn')) return;

  const button = document.createElement('button');
  button.id = 'themeToggleBtn';
  button.className = 'ghost-btn theme-toggle-btn';
  button.type = 'button';
  button.addEventListener('click', () => {
    const currentTheme = document.documentElement.dataset.theme === LIGHT_THEME ? LIGHT_THEME : DARK_THEME;
    const nextTheme = currentTheme === DARK_THEME ? LIGHT_THEME : DARK_THEME;
    applyTheme(nextTheme);
    saveTheme(nextTheme);
  });

  const topbar = document.querySelector('.topbar');
  const sessionChip = document.getElementById('sessionChip');
  if (sessionChip) {
    sessionChip.parentNode.insertBefore(button, sessionChip);
  } else if (topbar) {
    topbar.appendChild(button);
  } else {
    document.body.prepend(button);
  }

  applyTheme(document.documentElement.dataset.theme || getSavedTheme());
}

// Keep the existing dark appearance as the default and apply a saved light preference early.
applyTheme(getSavedTheme());

function initializeThemeControls() {
  installThemeStyles();
  installThemeToggle();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeThemeControls, { once: true });
} else {
  initializeThemeControls();
}
