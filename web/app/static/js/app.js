/* ======== Theme Management ======== */
const THEMES = ['classic-blue', 'ocean-teal', 'apple-clean'];

function getTheme() {
  return localStorage.getItem('hs-theme') || 'classic-blue';
}

function setTheme(theme) {
  if (!THEMES.includes(theme)) {
    theme = 'classic-blue';
    localStorage.removeItem('hs-theme');
  }
  localStorage.setItem('hs-theme', theme);
  document.body.setAttribute('data-theme', theme);
  document.querySelectorAll('.theme-switcher button').forEach(btn => {
    btn.classList.toggle('active-theme', btn.dataset.theme === theme);
  });
}

document.addEventListener('DOMContentLoaded', () => {
  setTheme(getTheme());
});


/* ======== Mobile Menu ======== */
function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  if (sidebar) sidebar.classList.toggle('open');
}

function closeSidebar() {
  const sidebar = document.getElementById('sidebar');
  if (sidebar) sidebar.classList.remove('open');
}

function toggleTopNav() {
  const nav = document.getElementById('topbar-nav');
  if (nav) nav.classList.toggle('open');
}


/* ======== Toast Notifications ======== */
function showToast(message, type = 'info', duration = 3000) {
  let container = document.querySelector('.toast-container');
  if (!container) {
    container = document.createElement('div');
    container.className = 'toast-container';
    document.body.appendChild(container);
  }
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(100%)';
    toast.style.transition = '0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}


/* ======== Modal Management ======== */
function openModal(id) {
  const overlay = document.getElementById(id);
  if (overlay) overlay.classList.add('show');
}

function closeModal(id) {
  const overlay = document.getElementById(id);
  if (overlay) overlay.classList.remove('show');
}

document.addEventListener('click', (e) => {
  if (e.target.classList.contains('modal-overlay')) {
    e.target.classList.remove('show');
  }
});


/* ======== Dropzone ======== */
function initDropzone(el, onFiles) {
  if (!el) return;
  let dragCounter = 0;

  el.addEventListener('dragenter', (e) => {
    e.preventDefault();
    dragCounter++;
    el.classList.add('dragover');
  });
  el.addEventListener('dragover', (e) => {
    e.preventDefault();
  });
  el.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dragCounter--;
    if (dragCounter <= 0) {
      dragCounter = 0;
      el.classList.remove('dragover');
    }
  });
  el.addEventListener('drop', (e) => {
    e.preventDefault();
    dragCounter = 0;
    el.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) onFiles(e.dataTransfer.files);
  });
  el.addEventListener('click', () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx';
    input.multiple = true;
    input.onchange = () => { if (input.files.length > 0) onFiles(input.files); };
    input.click();
  });
}


/* ======== API Helpers ======== */
async function apiFetch(url, options = {}) {
  if (typeof BASE_URL !== 'undefined' && BASE_URL && !url.startsWith('http')) {
    url = BASE_URL + url;
  }

  const headers = { ...options.headers };
  if (options.body && typeof options.body === 'object' && !(options.body instanceof FormData)) {
    options.body = JSON.stringify(options.body);
    if (!headers['Content-Type']) headers['Content-Type'] = 'application/json';
  }

  let resp;
  try {
    resp = await fetch(url, { ...options, headers });
  } catch (e) {
    showToast('网络连接失败', 'danger');
    throw e;
  }

  let data;
  try {
    data = await resp.json();
  } catch (e) {
    if (!resp.ok) {
      showToast(`请求失败 (${resp.status})`, 'danger');
      throw new Error(`HTTP ${resp.status}`);
    }
    return {};
  }

  if (!resp.ok) {
    showToast(data.error || `请求失败 (${resp.status})`, 'danger');
    throw new Error(data.error || `HTTP ${resp.status}`);
  }
  return data;
}


/* ======== Debounce ======== */
function debounce(fn, delay = 300) {
  let timer;
  return function (...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}


/* ======== Format Helpers ======== */
function formatDate(isoStr) {
  if (!isoStr) return '-';
  const d = new Date(isoStr);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function formatMoney(val, currency = '') {
  if (val == null || val === '') return '-';
  const n = parseFloat(val);
  if (isNaN(n)) return val;
  const s = n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return currency ? `${currency} ${s}` : s;
}
