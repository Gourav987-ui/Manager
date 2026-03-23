const API = '/api';
const sheetsBody = document.getElementById('sheetsBody');
const searchInput = document.getElementById('search');
const fileInput = document.getElementById('fileInput');
const refreshBtn = document.getElementById('refresh');
const emptyState = document.getElementById('emptyState');
const userAvatar = document.getElementById('userAvatar');
const leftTab = document.getElementById('leftTab');
const resizeHandle = document.getElementById('resizeHandle');

const DEFAULT_LEFT_TAB_WIDTH = 16.67;

let allSheets = [];

function setLeftTabWidth(pct) {
  const p = Math.min(25, Math.max(10, pct));
  document.documentElement.style.setProperty('--left-tab-width', p + '%');
}

(function initSettings() {
  const STORAGE_THEME = 'tsm_theme';
  const settingsBtn = document.getElementById('settingsBtn');
  const settingsPanel = document.getElementById('settingsPanel');
  const themeToggle = document.getElementById('themeToggle');
  const themeIcon = document.getElementById('themeIcon');

  const savedTheme = localStorage.getItem(STORAGE_THEME);
  if (savedTheme === 'light') document.body.classList.add('theme-light');

  settingsBtn?.addEventListener('click', () => {
    settingsPanel.hidden = !settingsPanel.hidden;
  });

  themeToggle?.addEventListener('click', () => {
    const isLight = document.body.classList.toggle('theme-light');
    localStorage.setItem(STORAGE_THEME, isLight ? 'light' : 'dark');
    if (themeIcon) themeIcon.textContent = isLight ? '☀' : '☾';
  });

  if (themeIcon) themeIcon.textContent = document.body.classList.contains('theme-light') ? '☀' : '☾';
})();

(function initResize() {
  setLeftTabWidth(DEFAULT_LEFT_TAB_WIDTH);
  if (!resizeHandle || !leftTab) return;
  let startX, startWidth;
  resizeHandle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    resizeHandle.classList.add('active');
    startX = e.clientX;
    startWidth = (leftTab.offsetWidth / window.innerWidth) * 100;
    const onMove = (e) => {
      const dx = ((e.clientX - startX) / window.innerWidth) * 100;
      setLeftTabWidth(startWidth + dx);
    };
    const onUp = () => {
      resizeHandle.classList.remove('active');
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
})();

(function initAvatarDropdown() {
  const avatarDropdown = document.getElementById('avatarDropdown');
  const avatarBackdrop = document.getElementById('avatarBackdrop');
  const avatarBtn = document.getElementById('userAvatar');
  const logoutBtn = document.getElementById('logout');
  if (!avatarDropdown || !avatarBackdrop || !avatarBtn) return;

  function close() {
    avatarDropdown.hidden = true;
    avatarBackdrop.hidden = true;
    avatarDropdown.style.top = '';
    avatarDropdown.style.right = '';
    avatarDropdown.style.transform = '';
    avatarBtn.setAttribute('aria-expanded', 'false');
  }

  function open() {
    const rect = avatarBtn.getBoundingClientRect();
    avatarBackdrop.hidden = false;
    avatarDropdown.hidden = false;
    avatarDropdown.style.top = (rect.bottom + 8) + 'px';
    avatarDropdown.style.right = (window.innerWidth - rect.right) + 'px';
    avatarBtn.setAttribute('aria-expanded', 'true');
  }

  avatarBtn.addEventListener('click', function (e) {
    e.preventDefault();
    e.stopPropagation();
    avatarDropdown.hidden ? open() : close();
  });

  logoutBtn?.addEventListener('click', async function () {
    close();
    await fetch(`${API}/logout`, { method: 'POST', credentials: 'include' });
    window.location.href = '/login';
  });

  avatarBackdrop.addEventListener('click', close);

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') close();
  });
})();

function checkAuth(res) {
  if (res.status === 401) {
    window.location.href = '/login';
    throw new Error('Unauthorized');
  }
  return res;
}

async function fetchSheets() {
  const res = await fetch(`${API}/sheets`, { credentials: 'include' }).then(checkAuth);
  if (!res.ok) throw new Error('Failed to load sheets');
  allSheets = await res.json();
  renderSheets(allSheets);
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDate(iso) {
  const d = new Date(iso);
  return d.toLocaleDateString(undefined, {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

function filterSheets(sheets, query) {
  if (!query.trim()) return sheets;
  const q = query.toLowerCase();
  return sheets.filter(
    (s) =>
      (s.ticketKey && s.ticketKey.toLowerCase().includes(q)) ||
      s.filename.toLowerCase().includes(q)
  );
}

function renderSheets(sheets) {
  const filtered = filterSheets(sheets, searchInput.value.trim());
  if (filtered.length === 0) {
    sheetsBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="6">${allSheets.length === 0 ? 'No test sheets yet. Upload one to get started.' : 'No matches for your filter.'}</td>
      </tr>
    `;
  } else {
    sheetsBody.innerHTML = filtered
      .map(
        (s) => `
      <tr class="selectable-row" data-filename="${escapeHtml(s.filename)}">
        <td class="td-checkbox"><input type="checkbox" class="row-select" data-filename="${escapeHtml(s.filename)}"></td>
        <td>
          ${s.ticketKey
            ? `<span class="ticket-badge"><a href="https://orion-advisor.atlassian.net/browse/${s.ticketKey}" target="_blank" rel="noopener">${s.ticketKey}</a></span>`
            : '<span class="no-ticket">—</span>'}
        </td>
        <td><span class="filename">${escapeHtml(s.filename)}</span></td>
        <td><span class="modified">${formatDate(s.modified)}</span></td>
        <td><span class="size">${formatSize(s.size)}</span></td>
        <td>
          <div class="actions">
            <a class="btn btn-open" href="${API}/sheets/open?file=${encodeURIComponent(s.filename)}" target="_blank">Open</a>
          </div>
        </td>
      </tr>
    `
      )
      .join('');
  }

  emptyState.hidden = allSheets.length > 0;
  updateBulkActions();
  bindSelectListeners();
}

function escapeHtml(s) {
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

function getSelectedFilenames() {
  return Array.from(document.querySelectorAll('.row-select:checked')).map((cb) => cb.dataset.filename);
}

function updateBulkActions() {
  const bulkActions = document.getElementById('bulkActions');
  const selectAll = document.getElementById('selectAll');
  const selected = getSelectedFilenames();
  const total = document.querySelectorAll('.row-select').length;
  if (bulkActions) bulkActions.hidden = selected.length === 0;
  if (selectAll) {
    selectAll.checked = total > 0 && selected.length === total;
    selectAll.indeterminate = selected.length > 0 && selected.length < total;
  }
}

function bindSelectListeners() {
  const selectAll = document.getElementById('selectAll');
  selectAll?.addEventListener('change', () => {
    document.querySelectorAll('.row-select').forEach((cb) => (cb.checked = selectAll.checked));
    updateBulkActions();
  });
  sheetsBody.querySelectorAll('.row-select').forEach((cb) => {
    cb.addEventListener('change', updateBulkActions);
  });
  sheetsBody.querySelectorAll('.selectable-row').forEach((row) => {
    row.addEventListener('click', (e) => {
      if (!e.ctrlKey && !e.metaKey) return;
      const cb = row.querySelector('.row-select');
      if (!cb) return;
      if (e.target.closest('a') || e.target.closest('input[type="checkbox"]')) return;
      e.preventDefault();
      cb.checked = !cb.checked;
      updateBulkActions();
    });
  });
  document.getElementById('bulkDownload')?.addEventListener('click', bulkDownload);
  document.getElementById('bulkDelete')?.addEventListener('click', bulkDelete);
}

async function bulkDownload() {
  const files = getSelectedFilenames();
  if (files.length === 0) return;
  if (files.length === 1) {
    window.location.href = `${API}/sheets/${encodeURIComponent(files[0])}/download`;
    return;
  }
  const res = await fetch(`${API}/sheets/download-batch`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ files }),
    credentials: 'include',
  }).then(checkAuth);
  if (!res.ok) {
    alert('Download failed');
    return;
  }
  const blob = await res.blob();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'test-sheets.zip';
  a.click();
  URL.revokeObjectURL(a.href);
}

async function bulkDelete() {
  const files = getSelectedFilenames();
  if (files.length === 0) return;
  if (!confirm(`Delete ${files.length} file(s)?`)) return;
  for (const f of files) {
    const res = await fetch(`${API}/sheets/${encodeURIComponent(f)}`, {
      method: 'DELETE',
      credentials: 'include',
    }).then(checkAuth);
    if (!res.ok) console.error('Delete failed:', f);
  }
  await fetchSheets();
}

async function deleteSheet(filename) {
  if (!confirm(`Delete "${filename}"?`)) return;
  const res = await fetch(`${API}/sheets/${encodeURIComponent(filename)}`, {
    method: 'DELETE',
    credentials: 'include',
  }).then(checkAuth);
  if (!res.ok) throw new Error('Delete failed');
  await fetchSheets();
}

async function uploadFile(file) {
  const form = new FormData();
  form.append('file', file);
  const res = await fetch(`${API}/sheets/upload`, {
    method: 'POST',
    body: form,
    credentials: 'include',
  }).then(checkAuth);
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || 'Upload failed');
  }
  await fetchSheets();
}

document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
    if (e.target.matches('input, textarea, [contenteditable]')) return;
    e.preventDefault();
    const selectAll = document.getElementById('selectAll');
    const checkboxes = document.querySelectorAll('.row-select');
    if (checkboxes.length === 0) return;
    checkboxes.forEach((cb) => (cb.checked = true));
    if (selectAll) selectAll.checked = true;
    updateBulkActions();
  }
});

searchInput.addEventListener('input', () => renderSheets(allSheets));

refreshBtn.addEventListener('click', () => fetchSheets().catch(console.error));

fileInput.addEventListener('change', async (e) => {
  const files = Array.from(e.target.files || []);
  if (files.length === 0) return;
  const errors = [];
  for (const file of files) {
    try {
      await uploadFile(file);
    } catch (err) {
      errors.push(`${file.name}: ${err.message}`);
    }
  }
  if (errors.length > 0) alert(errors.join('\n'));
  fileInput.value = '';
});

async function init() {
  let me = { email: 'guest@local' };
  try {
    const meRes = await fetch(`${API}/me`, { credentials: 'include' });
    if (meRes.status === 401) {
      window.location.href = '/login';
      return;
    }
    if (meRes.ok) me = await meRes.json().catch(() => me);
  } catch {
    /* API unreachable (e.g. static deploy) – continue as guest */
  }

  const userEmailEl = document.getElementById('userEmail');
  if (userEmailEl) userEmailEl.textContent = me.email || '—';
  const avatarInitialsEl = document.getElementById('avatarInitials');
  if (userAvatar && me.email) {
    userAvatar.title = me.email;
    const local = me.email.split('@')[0] || '';
    const parts = local.split(/[._\s-]+/).filter(Boolean);
    const initials = parts.slice(0, 2).map((p) => p[0]).join('').toUpperCase() || '';
    const avatarColors = ['#E91E63', '#9C27B0', '#2196F3', '#00BCD4', '#009688', '#4CAF50', '#8BC34A', '#FF9800', '#FF5722', '#F44336', '#673AB7', '#3F51B5', '#03A9F4', '#00ACC1', '#7CB342', '#FFC107'];
    const bg = avatarColors[Math.floor(Math.random() * avatarColors.length)];
    userAvatar.style.background = bg;
    if (avatarInitialsEl && initials) {
      avatarInitialsEl.textContent = initials;
      avatarInitialsEl.style.color = '#fff';
      userAvatar.classList.add('has-initials');
    }
  }

  fetchSheets().catch((err) => {
    sheetsBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="6">Error: ${escapeHtml(err.message)}. Make sure the server is running and config.testSheetsPath exists.</td>
      </tr>
    `;
  });
}

init();
