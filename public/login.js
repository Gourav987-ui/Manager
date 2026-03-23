const STORAGE_KEY = 'tsm_saved_passwords';

const form = document.getElementById('loginForm');
const emailInput = document.getElementById('email');
const passwordInput = document.getElementById('password');
const passwordToggle = document.getElementById('passwordToggle');
const rememberCheckbox = document.getElementById('remember');
const loginError = document.getElementById('loginError');
const loginBtn = document.getElementById('loginBtn');

function getSavedPasswords() {
  try {
    const data = localStorage.getItem(STORAGE_KEY);
    return data ? JSON.parse(data) : {};
  } catch {
    return {};
  }
}

function savePassword(email, password) {
  const saved = getSavedPasswords();
  saved[email.toLowerCase()] = password;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(saved));
}

emailInput?.addEventListener('input', () => {
  const email = emailInput.value.trim().toLowerCase();
  if (!email) return;
  const saved = getSavedPasswords();
  if (saved[email]) {
    passwordInput.value = saved[email];
  }
});

emailInput?.addEventListener('blur', () => {
  const email = emailInput.value.trim().toLowerCase();
  if (!email) return;
  const saved = getSavedPasswords();
  if (saved[email]) {
    passwordInput.value = saved[email];
  }
});

passwordToggle?.addEventListener('click', () => {
  const isHidden = passwordInput.type === 'password';
  passwordInput.type = isHidden ? 'text' : 'password';
  const iconEye = passwordToggle.querySelector('.icon-eye');
  const iconEyeOff = passwordToggle.querySelector('.icon-eye-off');
  const nowVisible = passwordInput.type === 'text';
  iconEye.style.display = nowVisible ? 'none' : 'block';
  iconEyeOff.style.display = nowVisible ? 'block' : 'none';
  passwordToggle.title = nowVisible ? 'Hide password' : 'Show password';
});

fetch('/api/me', { credentials: 'include' }).then((r) => {
  if (r.ok) window.location.href = '/';
});

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  loginError.hidden = true;
  loginBtn.disabled = true;

  try {
    const res = await fetch('/api/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        email: emailInput.value.trim().toLowerCase(),
        password: passwordInput.value,
      }),
      credentials: 'include',
    });

    const data = await res.json().catch(() => ({}));

    if (!res.ok) {
      loginError.textContent = data.error || 'Login failed';
      loginError.hidden = false;
      return;
    }

    if (rememberCheckbox?.checked) {
      savePassword(emailInput.value.trim().toLowerCase(), passwordInput.value);
    }

    window.location.href = '/';
  } catch (err) {
    loginError.textContent = 'Connection error. Please try again.';
    loginError.hidden = false;
  } finally {
    loginBtn.disabled = false;
  }
});
