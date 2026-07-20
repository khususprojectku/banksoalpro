// app.js - Main Application Logic & SPA Router for BankSoalPro
import { db } from './db.js';
import { exportToWord, exportToExcel, exportToPDF } from './exporter.js';
import { parseExcelOrCSV, downloadExcelTemplate } from './importer.js';

// Init DB
db.init();

// Global State
let activeUser = null;
let currentChartMapel = null;
let currentChartKesulitan = null;
let currentChartGuru = null;
let cbtAudio = null; // Global CBT background music player

// DOM Selectors
const appMount = document.getElementById('app');
const globalModal = document.getElementById('global-modal');
const modalContent = document.getElementById('modal-content-container');
const toastContainer = document.getElementById('toast-container');

// Core Router
function handleRouting() {
  activeUser = JSON.parse(sessionStorage.getItem('active_user') || 'null');

  if (!activeUser) {
    // Cleanup CBT resources if logging out
    if (cbtAudio) {
      cbtAudio.pause();
      cbtAudio = null;
    }
    const canvas = document.getElementById('cbt-animation-canvas');
    if (canvas && canvas.cleanup) {
      canvas.cleanup();
      canvas.remove();
    }
    renderLogin();
    return;
  }

  // Parse Hash
  const hash = window.location.hash || '#dashboard';
  const page = hash.split('?')[0];
  const queryParams = parseQueryParams(hash);

  // Cleanup audio & animation when routing away from CBT
  if (page !== '#cbt') {
    if (cbtAudio) {
      cbtAudio.pause();
      cbtAudio = null;
    }
    const canvas = document.getElementById('cbt-animation-canvas');
    if (canvas && canvas.cleanup) {
      canvas.cleanup();
      canvas.remove();
    }
  }

  if (page === '#cbt') {
    renderCBTScreen(queryParams.id);
  } else if (page === '#cbt-result') {
    renderCBTResultScreen(queryParams.id);
  } else {
    renderAppShell(page, queryParams);
  }
}

// Query parameters parser helper
function parseQueryParams(hash) {
  const params = {};
  const splitHash = hash.split('?');
  if (splitHash.length > 1) {
    const searchParams = new URLSearchParams(splitHash[1]);
    for (const [key, value] of searchParams.entries()) {
      params[key] = value;
    }
  }
  return params;
}

// Router Event Listeners
window.addEventListener('hashchange', handleRouting);
window.addEventListener('load', handleRouting);

// Toast Helper
function showToast(message, type = 'success') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;

  let icon = 'info';
  if (type === 'success') icon = 'check-circle';
  if (type === 'error') icon = 'alert-triangle';

  toast.innerHTML = `
    <i data-lucide="${icon}"></i>
    <span>${message}</span>
  `;

  toastContainer.appendChild(toast);
  lucide.createIcons();

  // Slide out and remove after 3s
  setTimeout(() => {
    toast.style.animation = 'slideIn 0.3s reverse forwards';
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

// Modal Helpers
function openModal(htmlContent, sizeClass = '') {
  modalContent.innerHTML = htmlContent;

  // Clear any existing sizes and apply new
  modalContent.className = 'modal-content';
  if (sizeClass) {
    modalContent.classList.add(sizeClass);
  }

  globalModal.classList.add('active');
  lucide.createIcons();

  // Bind close buttons
  const closeBtns = modalContent.querySelectorAll('[data-close-modal]');
  closeBtns.forEach(btn => {
    btn.addEventListener('click', closeModal);
  });
}

function closeModal() {
  globalModal.classList.remove('active');
}

// Reusable custom visual editor logic (WYSIWYG mockup)
function setupVisualEditor(textareaId, previewId) {
  const editor = document.getElementById(textareaId);
  const preview = document.getElementById(previewId);
  if (!editor) return;

  const updatePreview = () => {
    if (preview) {
      preview.innerHTML = editor.value;
      if (window.MathJax && window.MathJax.typesetPromise) {
        window.MathJax.typesetPromise([preview]).catch(err => console.error(err));
      }
    }
  };

  editor.addEventListener('input', updatePreview);

  // Bind toolbar buttons
  const container = editor.closest('.rich-editor-mock');
  if (container) {
    // Add Equation Editor toggle button if not exists
    const toolbar = container.querySelector('.editor-toolbar');
    if (toolbar && !toolbar.querySelector('#btn-eq-editor-toggle')) {
      const latexDisplayBtn = toolbar.querySelector('[data-editor-cmd="latex-display"]');
      
      const toggleBtn = document.createElement('button');
      toggleBtn.type = 'button';
      toggleBtn.className = 'toolbar-btn';
      toggleBtn.id = 'btn-eq-editor-toggle';
      toggleBtn.title = 'Equation Editor (Word Style)';
      toggleBtn.innerHTML = '<i data-lucide="square-pi" style="width:14px; height:14px;"></i>';
      
      if (latexDisplayBtn) {
        latexDisplayBtn.after(toggleBtn);
        const divider = document.createElement('span');
        divider.className = 'toolbar-divider';
        latexDisplayBtn.after(divider);
      } else {
        toolbar.appendChild(toggleBtn);
      }
      lucide.createIcons();
    }

    // Add Equation Editor Panel if not exists
    let eqPanel = container.querySelector('.eq-editor-panel');
    if (!eqPanel) {
      eqPanel = document.createElement('div');
      eqPanel.className = 'eq-editor-panel hidden';
      eqPanel.innerHTML = `
        <div class="eq-tabs">
          <button type="button" class="eq-tab-btn active" data-eq-tab="eq-pecahan">Pecahan</button>
          <button type="button" class="eq-tab-btn" data-eq-tab="eq-pangkat">Pangkat</button>
          <button type="button" class="eq-tab-btn" data-eq-tab="eq-akar">Akar & Log</button>
          <button type="button" class="eq-tab-btn" data-eq-tab="eq-kalkulus">Kalkulus</button>
          <button type="button" class="eq-tab-btn" data-eq-tab="eq-matriks">Matriks & Kurung</button>
          <button type="button" class="eq-tab-btn" data-eq-tab="eq-simbol">Simbol</button>
          <button type="button" class="eq-tab-btn" data-eq-tab="eq-yunani">Yunani</button>
        </div>
        <div class="eq-tab-contents" style="background:#fff; border:1px solid var(--neutral-400); border-radius:var(--radius-sm); padding:10px;">
          <!-- Pecahan -->
          <div class="eq-tab-pane active" id="eq-pecahan">
            <button type="button" class="eq-btn" data-latex="\\frac{x}{y}">$$\\frac{x}{y}$$</button>
            <button type="button" class="eq-btn" data-latex="\\frac{dy}{dx}">$$\\frac{dy}{dx}$$</button>
            <button type="button" class="eq-btn" data-latex="\\frac{\\partial y}{\\partial x}">$$\\frac{\\partial y}{\\partial x}$$</button>
            <button type="button" class="eq-btn" data-latex="x/y">$$x/y$$</button>
          </div>
          <!-- Pangkat -->
          <div class="eq-tab-pane" id="eq-pangkat">
            <button type="button" class="eq-btn" data-latex="x^{y}">$$x^{y}$$</button>
            <button type="button" class="eq-btn" data-latex="x_{y}">$$x_{y}$$</button>
            <button type="button" class="eq-btn" data-latex="x_{y}^{z}">$$x_{y}^{z}$$</button>
            <button type="button" class="eq-btn" data-latex="^{y}_{x}X">$$^{y}_{x}X$$</button>
          </div>
          <!-- Akar -->
          <div class="eq-tab-pane" id="eq-akar">
            <button type="button" class="eq-btn" data-latex="\\sqrt{x}">$$\\sqrt{x}$$</button>
            <button type="button" class="eq-btn" data-latex="\\sqrt[n]{x}">$$\\sqrt[n]{x}$$</button>
            <button type="button" class="eq-btn" data-latex="\\sqrt{x^2+y^2}">$$\\sqrt{x^2+y^2}$$</button>
            <button type="button" class="eq-btn" data-latex="\\log(x)">$$\\log(x)$$</button>
            <button type="button" class="eq-btn" data-latex="\\log_{b}(a)">$$\\log_{b}(a)$$</button>
            <button type="button" class="eq-btn" data-latex="\\ln(x)">$$\\ln(x)$$</button>
          </div>
          <!-- Kalkulus -->
          <div class="eq-tab-pane" id="eq-kalkulus">
            <button type="button" class="eq-btn" data-latex="\\sum_{i=1}^{n}">$$\\sum_{i=1}^{n}$$</button>
            <button type="button" class="eq-btn" data-latex="\\prod_{i=1}^{n}">$$\\prod_{i=1}^{n}$$</button>
            <button type="button" class="eq-btn" data-latex="\\int x \\, dx">$$\\int x \\, dx$$</button>
            <button type="button" class="eq-btn" data-latex="\\int_{a}^{b} x \\, dx">$$\\int_{a}^{b} x \\, dx$$</button>
            <button type="button" class="eq-btn" data-latex="\\iint x \\, dx \\, dy">$$\\iint x \\, dx \\, dy$$</button>
            <button type="button" class="eq-btn" data-latex="\\lim_{x \\to \\infty}">$$\\lim_{x \\to \\infty}$$</button>
            <button type="button" class="eq-btn" data-latex="\\lim_{x \\to 0}">$$\\lim_{x \\to 0}$$</button>
          </div>
          <!-- Matriks & Kurung -->
          <div class="eq-tab-pane" id="eq-matriks">
            <button type="button" class="eq-btn" data-latex="\\left( x \\right)">$$\\left( x \\right)$$</button>
            <button type="button" class="eq-btn" data-latex="\\left[ x \\right]">$$\\left[ x \\right]$$</button>
            <button type="button" class="eq-btn" data-latex="\\left\\{ x \\right\\}">$$\\left\\{ x \\right\\}$$</button>
            <button type="button" class="eq-btn" data-latex="\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}">$$\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}$$</button>
            <button type="button" class="eq-btn" data-latex="\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}">$$\\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix}$$</button>
            <button type="button" class="eq-btn" data-latex="\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}">$$\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}$$</button>
          </div>
          <!-- Simbol -->
          <div class="eq-tab-pane" id="eq-simbol">
            <button type="button" class="eq-btn" data-latex="\\pm">$$\\pm$$</button>
            <button type="button" class="eq-btn" data-latex="\\times">$$\\times$$</button>
            <button type="button" class="eq-btn" data-latex="\\div">$$\\div$$</button>
            <button type="button" class="eq-btn" data-latex="\\neq">$$\\neq$$</button>
            <button type="button" class="eq-btn" data-latex="\\approx">$$\\approx$$</button>
            <button type="button" class="eq-btn" data-latex="\\le">$$\\le$$</button>
            <button type="button" class="eq-btn" data-latex="\\ge">$$\\ge$$</button>
            <button type="button" class="eq-btn" data-latex="\\infty">$$\\infty$$</button>
            <button type="button" class="eq-btn" data-latex="\\to">$$\\to$$</button>
            <button type="button" class="eq-btn" data-latex="\\implies">$$\\implies$$</button>
            <button type="button" class="eq-btn" data-latex="\\iff">$$\\iff$$</button>
            <button type="button" class="eq-btn" data-latex="\\in">$$\\in$$</button>
            <button type="button" class="eq-btn" data-latex="\\notin">$$\\notin$$</button>
            <button type="button" class="eq-btn" data-latex="\\subset">$$\\subset$$</button>
            <button type="button" class="eq-btn" data-latex="\\cap">$$\\cap$$</button>
            <button type="button" class="eq-btn" data-latex="\\cup">$$\\cup$$</button>
            <button type="button" class="eq-btn" data-latex="\\forall">$$\\forall$$</button>
            <button type="button" class="eq-btn" data-latex="\\exists">$$\\exists$$</button>
            <button type="button" class="eq-btn" data-latex="\\therefore">$$\\therefore$$</button>
            <button type="button" class="eq-btn" data-latex="\\because">$$\\because$$</button>
          </div>
          <!-- Yunani -->
          <div class="eq-tab-pane" id="eq-yunani">
            <button type="button" class="eq-btn" data-latex="\\alpha">$$\\alpha$$</button>
            <button type="button" class="eq-btn" data-latex="\\beta">$$\\beta$$</button>
            <button type="button" class="eq-btn" data-latex="\\gamma">$$\\gamma$$</button>
            <button type="button" class="eq-btn" data-latex="\\delta">$$\\delta$$</button>
            <button type="button" class="eq-btn" data-latex="\\theta">$$\\theta$$</button>
            <button type="button" class="eq-btn" data-latex="\\pi">$$\\pi$$</button>
            <button type="button" class="eq-btn" data-latex="\\sigma">$$\\sigma$$</button>
            <button type="button" class="eq-btn" data-latex="\\omega">$$\\omega$$</button>
            <button type="button" class="eq-btn" data-latex="\\mu">$$\\mu$$</button>
            <button type="button" class="eq-btn" data-latex="\\lambda">$$\\lambda$$</button>
            <button type="button" class="eq-btn" data-latex="\\phi">$$\\phi$$</button>
            <button type="button" class="eq-btn" data-latex="\\Delta">$$\\Delta$$</button>
            <button type="button" class="eq-btn" data-latex="\\Sigma">$$\\Sigma$$</button>
            <button type="button" class="eq-btn" data-latex="\\Omega">$$\\Omega$$</button>
            <button type="button" class="eq-btn" data-latex="\\Theta">$$\\Theta$$</button>
          </div>
        </div>
      `;
      editor.before(eqPanel);

      // Tab switching handlers
      const tabBtns = eqPanel.querySelectorAll('.eq-tab-btn');
      const tabPanes = eqPanel.querySelectorAll('.eq-tab-pane');
      tabBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
          e.preventDefault();
          tabBtns.forEach(b => b.classList.remove('active'));
          tabPanes.forEach(p => p.classList.remove('active'));
          
          btn.classList.add('active');
          const paneId = btn.dataset.eqTab;
          eqPanel.querySelector(`#${paneId}`).classList.add('active');
        });
      });

      // Equation button click handlers
      const eqBtns = eqPanel.querySelectorAll('.eq-btn');
      eqBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
          e.preventDefault();
          const latex = btn.dataset.latex;
          
          const start = editor.selectionStart;
          const end = editor.selectionEnd;
          const text = editor.value;

          // Simple check for math block
          const beforeText = text.substring(0, start);
          const dollarCount = (beforeText.match(/\$/g) || []).length;
          const inMath = (dollarCount % 2 === 1);

          let replacement = latex;
          if (!inMath) {
            replacement = `$${latex}$`;
          }

          editor.value = text.substring(0, start) + replacement + text.substring(end);
          editor.focus();
          const newCursorPos = start + replacement.length;
          editor.setSelectionRange(newCursorPos, newCursorPos);

          // Dispatch input event to update previews
          editor.dispatchEvent(new Event('input', { bubbles: true }));
        });
      });
    }

    // Toggle panel event handler
    const toggleBtn = toolbar.querySelector('#btn-eq-editor-toggle');
    if (toggleBtn) {
      toggleBtn.addEventListener('click', (e) => {
        e.preventDefault();
        const isHidden = eqPanel.classList.toggle('hidden');
        if (!isHidden) {
          if (window.MathJax && window.MathJax.typesetPromise) {
            window.MathJax.typesetPromise([eqPanel]).catch(err => console.error(err));
          }
        }
      });
    }

    // Bind original toolbar buttons
    const buttons = container.querySelectorAll('[data-editor-cmd]');
    buttons.forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        const cmd = btn.dataset.editorCmd;
        insertTextAtCursor(editor, cmd);
        updatePreview();
      });
    });
  }

  updatePreview();
}

function insertTextAtCursor(textarea, cmd) {
  const start = textarea.selectionStart;
  const end = textarea.selectionEnd;
  const text = textarea.value;
  const selectedText = text.substring(start, end);

  let replacement = '';
  switch (cmd) {
    case 'bold':
      replacement = `<b>${selectedText || 'Teks Tebal'}</b>`;
      break;
    case 'italic':
      replacement = `<i>${selectedText || 'Teks Miring'}</i>`;
      break;
    case 'latex-inline':
      replacement = `$${selectedText || 'x^2 + y^2 = z^2'}$`;
      break;
    case 'latex-display':
      replacement = `$$${selectedText || '\\sum_{i=1}^n i = \\frac{n(n+1)}{2}$'}$`;
      break;
    case 'table':
      replacement = `
<table border="1" style="border-collapse: collapse; width: 100%; text-align: center;">
  <thead>
    <tr style="background-color: #f2f2f2;">
      <th>Kolom 1</th>
      <th>Kolom 2</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Baris 1 Sel 1</td>
      <td>Baris 1 Sel 2</td>
    </tr>
  </tbody>
</table>\n`;
      break;
    case 'image':
      const url = prompt('Masukkan URL Gambar:', 'https://images.unsplash.com/photo-1434030216411-0b793f4b4173?w=400');
      if (url) {
        replacement = `<img src="${url}" alt="Gambar Soal" style="max-width: 100%; height: auto; margin: 10px 0; border-radius: 8px;" />`;
      }
      break;
    default:
      return;
  }

  textarea.value = text.substring(0, start) + replacement + text.substring(end);
  textarea.focus();
  textarea.selectionStart = start + replacement.length;
  textarea.selectionEnd = start + replacement.length;
}

// ----------------------------------------------------
// 1. RENDER LOGIN
// ----------------------------------------------------
function renderLogin() {
  appMount.innerHTML = `
    <div class="auth-page">
      <div class="auth-card">
        <div class="auth-header">
          <div class="auth-logo">
            <div class="logo-icon">BSP</div>
            <div class="logo-text">BankSoalPro</div>
          </div>
          <h2 class="auth-title">Selamat Datang</h2>
          <p class="auth-desc">Masuk untuk mengelola bank soal sekolah Anda</p>
        </div>
        
        <form id="login-form">
          <div class="form-group">
            <label class="form-label" for="login-email">Alamat Email</label>
            <input type="email" id="login-email" class="form-input" placeholder="contoh: abukhoirmkom@sman4kisaran.sch.id" required>
          </div>
          <div class="form-group">
            <label class="form-label" for="login-password">
              <span>Kata Sandi</span>
              <a href="#" style="color: var(--primary); text-decoration: none; font-size: 11px;">Lupa Sandi?</a>
            </label>
            <input type="password" id="login-password" class="form-input" placeholder="••••••••" required>
          </div>
          
          <button type="submit" class="btn btn-primary" style="width: 100%; margin-top: 10px;">
            <span>Masuk ke Akun</span>
            <i data-lucide="arrow-right"></i>
          </button>
        </form>
        
        <div class="demo-accounts">
          <div class="demo-title">Akun Uji Coba Quick Login</div>
          <div class="demo-grid">
            <button class="btn-demo" data-email="admin@banksoal.pro">Super Admin</button>
            <button class="btn-demo" data-email="admin.sekolah@sman1.sch.id">Admin Sekolah</button>
            <button class="btn-demo" data-email="budi.guru@sman1.sch.id">Guru Informatika</button>
            <button class="btn-demo" data-email="siti.guru@sman1.sch.id">Guru Matematika</button>
            <button class="btn-demo" data-email="edi.reviewer@sman1.sch.id">Reviewer Soal</button>
            <button class="btn-demo" data-email="andi.siswa@sman1.sch.id">Siswa Demo</button>
          </div>
          <div style="text-align: center; margin-top: 16px; border-top: 1px dashed var(--neutral-400); padding-top: 12px;">
            <a href="#" id="btn-reset-login" style="color: var(--danger-text); text-decoration: none; font-size: 12px; font-weight: 600; display: inline-flex; align-items: center; gap: 6px;">
              <i data-lucide="refresh-cw" style="width:12px; height:12px;"></i>
              <span>Reset Database & Sinkronkan Paket Matematika Baru</span>
            </a>
          </div>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  // Handle Form Submission
  const form = document.getElementById('login-form');
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const email = document.getElementById('login-email').value;
    const password = document.getElementById('login-password').value;
    attemptLogin(email, password);
  });

  // Handle Demo login clicks
  const demoBtns = document.querySelectorAll('.btn-demo');
  demoBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const email = btn.dataset.email;
      const user = db.get('users').find(u => u.email === email);
      if (user) {
        document.getElementById('login-email').value = user.email;
        document.getElementById('login-password').value = user.password;
        attemptLogin(user.email, user.password);
      }
    });
  });

  // Handle Reset button click
  const resetLoginBtn = document.getElementById('btn-reset-login');
  if (resetLoginBtn) {
    resetLoginBtn.addEventListener('click', (e) => {
      e.preventDefault();
      if (confirm('PERINGATAN: Semua data Anda akan disetel kembali ke contoh bawaan. Paket matematika 50 soal dan data hasil ujian siswa akan diinisialisasi. Lanjutkan?')) {
        db.reset();
        showToast('Database disetel ulang ke setelan pabrik!', 'success');
        handleRouting();
      }
    });
  }
}

function attemptLogin(email, password) {
  const user = db.get('users').find(u => u.email === email && u.password === password);

  if (user) {
    if (user.status !== 'AKTIF') {
      showToast('Akun Anda dinonaktifkan. Hubungi admin.', 'error');
      return;
    }

    sessionStorage.setItem('active_user', JSON.stringify(user));
    db.log(user.id, user.name, 'LOGIN', 'Berhasil login ke dalam sistem.');
    showToast(`Selamat datang kembali, ${user.name}!`, 'success');
    window.location.hash = '#dashboard';
  } else {
    showToast('Email atau password salah.', 'error');
  }
}

// ----------------------------------------------------
// 2. RENDER APP SHELL (Sidebar & Header Layout)
// ----------------------------------------------------
function renderAppShell(activePage, queryParams) {
  const school = db.get('schools')[0] || { name: 'BankSoalPro', npsn: '-' };

  const menuConfig = [
    { page: '#dashboard', label: 'Dashboard', icon: 'layout-dashboard', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER', 'SISWA'] },
    { page: '#profil-sekolah', label: 'Profil Sekolah', icon: 'school', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH'] },
    { page: '#guru', label: 'Manajemen Guru', icon: 'users', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH'] },
    { page: '#struktur', label: 'Struktur Kurikulum', icon: 'git-branch', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER'] },
    { page: '#soal', label: 'Bank Soal', icon: 'file-text', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER'] },
    { page: '#impor', label: 'Import Soal', icon: 'file-up', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU'] },
    { page: '#paket', label: 'Paket Soal', icon: 'package', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER'] },
    { page: '#rekap-ujian', label: 'Hasil & Rekap Ujian', icon: 'award', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'SISWA'] },
    { page: '#logs', label: 'Audit Log & Backup', icon: 'shield-alert', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH'] },
  ];

  let sidebarMenuHtml = '';
  menuConfig.forEach(item => {
    if (item.roles.includes(activeUser.role)) {
      const isActive = activePage === item.page ? 'active' : '';
      sidebarMenuHtml += `
        <a class="menu-item ${isActive}" href="${item.page}">
          <i data-lucide="${item.icon}"></i>
          <span>${item.label}</span>
        </a>
      `;
    }
  });

  const initials = activeUser.name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();

  appMount.innerHTML = `
    <div class="app-container">
      <div class="sidebar-overlay" id="sidebar-overlay"></div>
      
      <!-- Sidebar -->
      <aside class="sidebar">
        <div class="sidebar-logo">
          <div class="logo-icon">BSP</div>
          <div class="logo-text">BankSoalPro</div>
        </div>
        
        <nav class="sidebar-menu">
          <div class="menu-section-title">Menu Utama</div>
          ${sidebarMenuHtml}
        </nav>
        
        <div class="sidebar-user">
          <div class="user-avatar">${initials}</div>
          <div class="user-info">
            <div class="user-name" title="${activeUser.name}">${activeUser.name}</div>
            <div class="user-role">${activeUser.role.replace('_', ' ')}</div>
          </div>
          <button class="logout-btn" id="logout-btn-action" title="Keluar">
            <i data-lucide="log-out"></i>
          </button>
        </div>
      </aside>
      
      <!-- Main Content -->
      <div class="main-wrapper">
        <header class="header">
          <div style="display:flex; align-items:center; gap:12px;">
            <button id="sidebar-toggle-btn" class="btn" style="display:none; padding:8px; background:none; border:1px solid var(--neutral-400); color:var(--neutral-800); border-radius:var(--radius-sm); cursor:pointer;">
              <i data-lucide="menu" style="width:20px; height:20px;"></i>
            </button>
            <div class="header-title-area">
              <h1 class="header-title" id="page-display-title">Dashboard</h1>
              <span class="header-subtitle" id="page-display-subtitle">Statistik & analisis data sistem</span>
            </div>
          </div>
          
          <div class="header-actions">
            <div class="school-badge">
              <i data-lucide="school" style="width:16px; height:16px;"></i>
              <span>${school.name} (${school.academicYear})</span>
            </div>
          </div>
        </header>
        
        <main class="content-body" id="shell-content-mount"></main>
      </div>
    </div>
  `;

  // Bind logout button
  document.getElementById('logout-btn-action').addEventListener('click', () => {
    db.log(activeUser.id, activeUser.name, 'LOGOUT', 'User keluar dari aplikasi.');
    sessionStorage.removeItem('active_user');
    window.location.hash = '';
    showToast('Berhasil keluar akun.', 'info');
    handleRouting();
  });

  lucide.createIcons();

  // Bind sidebar toggle on mobile
  const toggleBtn = document.getElementById('sidebar-toggle-btn');
  const sidebar = document.querySelector('.sidebar');
  const overlay = document.getElementById('sidebar-overlay');

  if (toggleBtn && sidebar && overlay) {
    const toggleSidebar = () => {
      sidebar.classList.toggle('open');
      overlay.classList.toggle('open');
    };

    toggleBtn.addEventListener('click', toggleSidebar);
    overlay.addEventListener('click', toggleSidebar);

    // Close sidebar when clicking links
    sidebar.querySelectorAll('.menu-item').forEach(item => {
      item.addEventListener('click', () => {
        sidebar.classList.remove('open');
        overlay.classList.remove('open');
      });
    });
  }

  // Render Sub page contents
  const contentMount = document.getElementById('shell-content-mount');
  const titleDisplay = document.getElementById('page-display-title');
  const subtitleDisplay = document.getElementById('page-display-subtitle');

  switch (activePage) {
    case '#dashboard':
      titleDisplay.innerText = 'Dashboard Utama';
      subtitleDisplay.innerText = 'Pemantauan statistik bank soal terpusat';
      renderDashboard(contentMount);
      break;
    case '#profil-sekolah':
      titleDisplay.innerText = 'Profil Sekolah';
      subtitleDisplay.innerText = 'Kelola informasi identitas & tahun ajaran sekolah';
      renderProfilSekolah(contentMount);
      break;
    case '#guru':
      titleDisplay.innerText = 'Manajemen Pendidik';
      subtitleDisplay.innerText = 'Kelola akun guru mata pelajaran dan hak akses';
      renderGuru(contentMount);
      break;
    case '#struktur':
      titleDisplay.innerText = 'Struktur Kurikulum';
      subtitleDisplay.innerText = 'Kelola Mata Pelajaran, Tingkat Kelas, Bab, dan Topik Soal';
      renderKurikulum(contentMount);
      break;
    case '#soal':
      if (queryParams.action === 'tambah' || queryParams.action === 'edit') {
        titleDisplay.innerText = queryParams.action === 'tambah' ? 'Buat Soal Baru' : 'Ubah Soal';
        subtitleDisplay.innerText = 'Form penyusunan soal ujian standard nasional';
        renderFormSoal(contentMount, queryParams.id);
      } else {
        titleDisplay.innerText = 'Bank Penyimpanan Soal';
        subtitleDisplay.innerText = 'Cari, saring, edit, dan verifikasi kualitas soal';
        renderBankSoal(contentMount);
      }
      break;
    case '#impor':
      titleDisplay.innerText = 'Import Soal Excel';
      subtitleDisplay.innerText = 'Unggah lembar kerja excel untuk memasukkan soal secara massal';
      renderImportSoal(contentMount);
      break;
    case '#paket':
      titleDisplay.innerText = 'Manajemen Paket Soal';
      subtitleDisplay.innerText = 'Penyusunan paket ujian otomatis/acak dan ekspor berkas';
      renderPaketSoal(contentMount);
      break;
    case '#rekap-ujian':
      titleDisplay.innerText = 'Hasil & Rekap Ujian';
      subtitleDisplay.innerText = activeUser.role === 'SISWA'
        ? 'Pemantauan riwayat nilai dan analisis hasil ujian mandiri Anda'
        : 'Pemantauan rekapitulasi nilai dan riwayat pengerjaan ujian siswa';
      renderRekapUjian(contentMount);
      break;
    case '#logs':
      titleDisplay.innerText = 'Audit Keamanan & Cadangan';
      subtitleDisplay.innerText = 'Riwayat log aktivitas administrator dan backup/restore basis data';
      renderLogs(contentMount);
      break;
    default:
      window.location.hash = '#dashboard';
      break;
  }
}

// ----------------------------------------------------
// 3. PAGE: DASHBOARD
// ----------------------------------------------------
function renderDashboard(mount) {
  if (activeUser.role === 'SISWA') {
    renderStudentDashboard(mount);
    return;
  }

  const users = db.get('users').filter(u => u.role === 'GURU');
  const subjects = db.get('subjects');
  const questions = db.get('questions');
  const packages = db.get('packages');
  const logs = db.get('logs');

  mount.innerHTML = `
    <!-- Metrics Rows -->
    <div class="metrics-grid">
      <div class="metric-card">
        <div class="metric-info">
          <span class="metric-label">Jumlah Guru</span>
          <span class="metric-value">${users.length}</span>
        </div>
        <div class="metric-icon-box metric-blue">
          <i data-lucide="users"></i>
        </div>
      </div>
      
      <div class="metric-card">
        <div class="metric-info">
          <span class="metric-label">Mata Pelajaran</span>
          <span class="metric-value">${subjects.length}</span>
        </div>
        <div class="metric-icon-box metric-purple">
          <i data-lucide="book-open"></i>
        </div>
      </div>
      
      <div class="metric-card">
        <div class="metric-info">
          <span class="metric-label">Total Soal</span>
          <span class="metric-value">${questions.length}</span>
        </div>
        <div class="metric-icon-box metric-indigo">
          <i data-lucide="file-text"></i>
        </div>
      </div>
      
      <div class="metric-card">
        <div class="metric-info">
          <span class="metric-label">Paket Ujian</span>
          <span class="metric-value">${packages.length}</span>
        </div>
        <div class="metric-icon-box metric-orange">
          <i data-lucide="package"></i>
        </div>
      </div>
    </div>

    <!-- Charts Row -->
    <div class="charts-grid">
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Proporsi Soal per Mata Pelajaran</h3>
          <i data-lucide="bar-chart-3" style="color: var(--neutral-600);"></i>
        </div>
        <div class="card-body">
          <div class="chart-container">
            <canvas id="chart-mapel"></canvas>
          </div>
        </div>
      </div>
      
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Tingkat Kesulitan</h3>
          <i data-lucide="pie-chart" style="color: var(--neutral-600);"></i>
        </div>
        <div class="card-body">
          <div class="chart-container">
            <canvas id="chart-kesulitan"></canvas>
          </div>
        </div>
      </div>
    </div>

    <!-- Lists Row -->
    <div class="charts-grid">
      <!-- Latest Questions -->
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Soal Terbaru Dibuat</h3>
          <a href="#soal" class="btn btn-secondary" style="padding: 6px 12px; font-size: 12px;">Buka Bank Soal</a>
        </div>
        <div class="card-body" style="padding:0;">
          <div class="table-wrapper">
            <table class="table">
              <thead>
                <tr>
                  <th>Kode</th>
                  <th>Mata Pelajaran</th>
                  <th>Tipe</th>
                  <th>Tingkat</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                ${questions.slice(0, 5).map(q => {
    const sub = subjects.find(s => s.id === q.subjectId);
    const statusClass = q.status === 'APPROVED' ? 'success' : q.status === 'REVIEW' ? 'warning' : 'neutral';
    return `
                    <tr>
                      <td style="font-weight: 700; color: var(--primary);">${q.code}</td>
                      <td>${sub ? sub.name : 'Umum'}</td>
                      <td><span class="badge badge-primary">${q.type}</span></td>
                      <td><span class="badge badge-neutral">${q.difficulty}</span></td>
                      <td><span class="badge badge-${statusClass}">${q.status}</span></td>
                    </tr>
                  `;
  }).join('') || '<tr><td colspan="5" style="text-align:center;">Belum ada data soal.</td></tr>'}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Logs -->
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Log Aktivitas Guru</h3>
          <i data-lucide="activity" style="color: var(--neutral-600);"></i>
        </div>
        <div class="card-body" style="max-height: 310px; overflow-y: auto; padding: 12px;">
          <div style="display:flex; flex-direction:column; gap:12px;">
            ${logs.slice(0, 8).map(l => {
    const time = new Date(l.timestamp).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
    return `
                <div style="display:flex; flex-direction:column; border-bottom: 1px solid var(--neutral-400); padding-bottom: 8px;">
                  <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-weight:700; font-size:13px; color: var(--neutral-900);">${l.userName}</span>
                    <span style="font-size:11px; color: var(--neutral-600);">${time}</span>
                  </div>
                  <span style="font-size:12px; color: var(--neutral-700); margin-top:2px;">${l.details}</span>
                </div>
              `;
  }).join('') || '<p style="text-align:center; color: var(--neutral-600); font-size:13px;">Belum ada aktivitas.</p>'}
          </div>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  // Initialize charts after browser paint
  setTimeout(() => {
    // 1. Chart Mapel (Bar Chart)
    const mapelCounts = {};
    subjects.forEach(s => { mapelCounts[s.name] = 0; });
    questions.forEach(q => {
      const s = subjects.find(sub => sub.id === q.subjectId);
      if (s) {
        mapelCounts[s.name] = (mapelCounts[s.name] || 0) + 1;
      }
    });

    const ctxMapel = document.getElementById('chart-mapel').getContext('2d');
    if (currentChartMapel) currentChartMapel.destroy();
    currentChartMapel = new Chart(ctxMapel, {
      type: 'bar',
      data: {
        labels: Object.keys(mapelCounts),
        datasets: [{
          label: 'Jumlah Soal',
          data: Object.values(mapelCounts),
          backgroundColor: 'rgba(79, 70, 229, 0.75)',
          borderColor: 'rgb(79, 70, 229)',
          borderWidth: 1.5,
          borderRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { stepSize: 1 } },
          x: { grid: { display: false } }
        }
      }
    });

    // 2. Chart Kesulitan (Pie Chart)
    const diffCounts = { 'MUDAH': 0, 'SEDANG': 0, 'SULIT': 0 };
    questions.forEach(q => {
      if (diffCounts[q.difficulty] !== undefined) {
        diffCounts[q.difficulty]++;
      }
    });

    const ctxKesulitan = document.getElementById('chart-kesulitan').getContext('2d');
    if (currentChartKesulitan) currentChartKesulitan.destroy();
    currentChartKesulitan = new Chart(ctxKesulitan, {
      type: 'doughnut',
      data: {
        labels: Object.keys(diffCounts),
        datasets: [{
          data: Object.values(diffCounts),
          backgroundColor: [
            'rgba(34, 197, 94, 0.75)',  // Mudah - Green
            'rgba(245, 158, 11, 0.75)', // Sedang - Orange/Yellow
            'rgba(239, 68, 68, 0.75)'   // Sulit - Red
          ],
          borderColor: ['#fff', '#fff', '#fff'],
          borderWidth: 2
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom', labels: { boxWidth: 12, font: { family: 'Plus Jakarta Sans' } } }
        },
        cutout: '60%'
      }
    });
  }, 100);
}

// ----------------------------------------------------
// 4. PAGE: PROFIL SEKOLAH
// ----------------------------------------------------
function renderProfilSekolah(mount) {
  const school = db.get('schools')[0] || {};

  mount.innerHTML = `
    <div class="card" style="max-width:700px;">
      <div class="card-header">
        <h3 class="card-title">Identitas Sekolah</h3>
      </div>
      <div class="card-body">
        <form id="school-form">
          <div class="form-group">
            <label class="form-label" for="sch-name">Nama Sekolah</label>
            <input type="text" id="sch-name" class="form-input" value="${school.name || ''}" required>
          </div>
          
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
            <div class="form-group">
              <label class="form-label" for="sch-npsn">NPSN</label>
              <input type="text" id="sch-npsn" class="form-input" value="${school.npsn || ''}" required>
            </div>
            <div class="form-group">
              <label class="form-label" for="sch-year">Tahun Ajaran</label>
              <input type="text" id="sch-year" class="form-input" value="${school.academicYear || ''}" placeholder="2025/2026" required>
            </div>
          </div>
          
          <div class="form-group">
            <label class="form-label" for="sch-logo">URL Logo Sekolah</label>
            <input type="text" id="sch-logo" class="form-input" value="${school.logo || ''}">
          </div>

          <div class="form-group">
            <label class="form-label" for="sch-address">Alamat Lengkap</label>
            <input type="text" id="sch-address" class="form-input" value="${school.address || ''}">
          </div>

          <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
            <div class="form-group">
              <label class="form-label" for="sch-email">Alamat Email</label>
              <input type="email" id="sch-email" class="form-input" value="${school.email || ''}">
            </div>
            <div class="form-group">
              <label class="form-label" for="sch-phone">Nomor Telepon</label>
              <input type="text" id="sch-phone" class="form-input" value="${school.phone || ''}">
            </div>
          </div>

          <button type="submit" class="btn btn-primary" style="margin-top:12px;">
            <i data-lucide="save"></i>
            <span>Simpan Perubahan</span>
          </button>
        </form>
      </div>
    </div>
  `;

  lucide.createIcons();

  document.getElementById('school-form').addEventListener('submit', (e) => {
    e.preventDefault();
    const updated = {
      name: document.getElementById('sch-name').value,
      npsn: document.getElementById('sch-npsn').value,
      academicYear: document.getElementById('sch-year').value,
      logo: document.getElementById('sch-logo').value,
      address: document.getElementById('sch-address').value,
      email: document.getElementById('sch-email').value,
      phone: document.getElementById('sch-phone').value
    };

    db.update('schools', school.id, updated);
    showToast('Profil sekolah berhasil diperbarui.', 'success');
    handleRouting(); // Reload app header, dll.
  });
}

// ----------------------------------------------------
// 5. PAGE: MANAJEMEN GURU
// ----------------------------------------------------
function renderGuru(mount) {
  const teachers = db.get('users').filter(u => u.role === 'GURU');
  const subjects = db.get('subjects');

  mount.innerHTML = `
    <div class="action-row">
      <div></div>
      <button class="btn btn-primary" id="btn-tambah-guru">
        <i data-lucide="plus"></i>
        <span>Tambah Guru</span>
      </button>
    </div>

    <div class="card">
      <div class="card-body" style="padding:0;">
        <div class="table-wrapper">
          <table class="table">
            <thead>
              <tr>
                <th>Nama Lengkap</th>
                <th>NIP</th>
                <th>Email</th>
                <th>Mengajar Mapel</th>
                <th>Status</th>
                <th style="text-align:right;">Aksi</th>
              </tr>
            </thead>
            <tbody>
              ${teachers.map(t => {
    const subNames = (t.subjectIds || []).map(sid => {
      const s = subjects.find(sub => sub.id === sid);
      return s ? s.name : '';
    }).filter(Boolean).join(', ') || 'Belum diatur';

    const statusClass = t.status === 'AKTIF' ? 'success' : 'danger';

    return `
                  <tr>
                    <td>
                      <div style="font-weight:700; color: var(--neutral-900);">${t.name}</div>
                    </td>
                    <td><code>${t.nip || '-'}</code></td>
                    <td>${t.email}</td>
                    <td>${subNames}</td>
                    <td><span class="badge badge-${statusClass}">${t.status}</span></td>
                    <td style="text-align:right;">
                      <div style="display:inline-flex; gap:8px;">
                        <button class="btn btn-secondary btn-icon btn-edit-guru" data-id="${t.id}" title="Edit">
                          <i data-lucide="edit" style="width:16px; height:16px;"></i>
                        </button>
                        <button class="btn btn-danger btn-icon btn-delete-guru" data-id="${t.id}" title="Hapus">
                          <i data-lucide="trash-2" style="width:16px; height:16px;"></i>
                        </button>
                      </div>
                    </td>
                  </tr>
                `;
  }).join('') || '<tr><td colspan="6" style="text-align:center;">Belum ada data guru.</td></tr>'}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  // Add click listeners
  document.getElementById('btn-tambah-guru').addEventListener('click', () => openGuruModal());

  const editBtns = mount.querySelectorAll('.btn-edit-guru');
  editBtns.forEach(btn => {
    btn.addEventListener('click', () => openGuruModal(btn.dataset.id));
  });

  const deleteBtns = mount.querySelectorAll('.btn-delete-guru');
  deleteBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.id;
      const t = db.get('users').find(u => u.id === id);
      if (confirm(`Apakah Anda yakin ingin menghapus akun guru "${t.name}"?`)) {
        db.delete('users', id);
        showToast('Akun guru berhasil dihapus.', 'success');
        renderGuru(mount);
      }
    });
  });
}

function openGuruModal(teacherId = null) {
  const isEdit = !!teacherId;
  const teacher = isEdit ? db.get('users').find(u => u.id === teacherId) : {};
  const subjects = db.get('subjects');

  const title = isEdit ? 'Edit Data Guru' : 'Tambah Guru Baru';

  const modalHtml = `
    <div class="modal-header">
      <h3 class="modal-title">${title}</h3>
      <button class="modal-close" data-close-modal><i data-lucide="x"></i></button>
    </div>
    <form id="guru-modal-form">
      <div class="modal-body">
        <div class="form-group">
          <label class="form-label" for="g-name">Nama Lengkap & Gelar</label>
          <input type="text" id="g-name" class="form-input" value="${teacher.name || ''}" required>
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
          <div class="form-group">
            <label class="form-label" for="g-nip">NIP</label>
            <input type="text" id="g-nip" class="form-input" value="${teacher.nip || ''}">
          </div>
          <div class="form-group">
            <label class="form-label" for="g-status">Status Akun</label>
            <select id="g-status" class="form-input">
              <option value="AKTIF" ${teacher.status === 'AKTIF' ? 'selected' : ''}>Aktif</option>
              <option value="NONAKTIF" ${teacher.status === 'NONAKTIF' ? 'selected' : ''}>Non-Aktif</option>
            </select>
          </div>
        </div>
        
        <div class="form-group">
          <label class="form-label" for="g-email">Alamat Email (Sebagai Username)</label>
          <input type="email" id="g-email" class="form-input" value="${teacher.email || ''}" required>
        </div>
        
        <div class="form-group">
          <label class="form-label" for="g-password">Kata Sandi</label>
          <input type="password" id="g-password" class="form-input" placeholder="${isEdit ? 'Kosongkan jika tidak ingin diubah' : 'Min 6 karakter'}" ${isEdit ? '' : 'required'}>
        </div>

        <div class="form-group">
          <label class="form-label">Mata Pelajaran yang Diajar</label>
          <div style="display:grid; grid-template-columns:repeat(2, 1fr); gap:8px; margin-top:4px;">
            ${subjects.map(s => {
    const isChecked = isEdit && (teacher.subjectIds || []).includes(s.id) ? 'checked' : '';
    return `
                <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;">
                  <input type="checkbox" class="g-subject-chk" value="${s.id}" ${isChecked}>
                  <span>${s.name}</span>
                </label>
              `;
  }).join('')}
          </div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
        <button type="submit" class="btn btn-primary">Simpan Data</button>
      </div>
    </form>
  `;

  openModal(modalHtml);

  document.getElementById('guru-modal-form').addEventListener('submit', (e) => {
    e.preventDefault();

    const subjectIds = [];
    document.querySelectorAll('.g-subject-chk:checked').forEach(chk => {
      subjectIds.push(chk.value);
    });

    const email = document.getElementById('g-email').value;
    // Check if email already registered for another user
    const existing = db.get('users').find(u => u.email === email && u.id !== teacherId);
    if (existing) {
      showToast('Alamat email sudah digunakan oleh akun lain.', 'error');
      return;
    }

    const name = document.getElementById('g-name').value;
    const nip = document.getElementById('g-nip').value;
    const status = document.getElementById('g-status').value;
    const passwordInput = document.getElementById('g-password').value;

    const data = {
      name,
      nip,
      email,
      status,
      subjectIds,
      role: 'GURU',
      schoolId: 'sch-1'
    };

    if (passwordInput) {
      data.password = passwordInput;
    }

    if (isEdit) {
      db.update('users', teacherId, data);
      showToast('Profil guru berhasil diperbarui.', 'success');
    } else {
      if (!passwordInput) {
        showToast('Password wajib diisi untuk guru baru.', 'error');
        return;
      }
      db.insert('users', data);
      showToast('Guru baru berhasil ditambahkan.', 'success');
    }

    closeModal();
    renderGuru(document.getElementById('shell-content-mount'));
  });
}

// ----------------------------------------------------
// 6. PAGE: STRUKTUR KURIKULUM (Mapel, Kelas, Bab, Topik)
// ----------------------------------------------------
function renderKurikulum(mount) {
  mount.innerHTML = `
    <div class="nested-structure-grid">
      <!-- 1. Column Mapel & Kelas -->
      <div style="display:flex; flex-direction:column; gap:24px;">
        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Mata Pelajaran (Subject)</h3>
            <button class="btn btn-primary" style="padding:6px 12px; font-size:12px;" id="btn-add-subject">
              <i data-lucide="plus" style="width:14px; height:14px;"></i> Tambah
            </button>
          </div>
          <div class="card-body" style="padding:0;">
            <div class="structure-list-box" id="subject-list-mount"></div>
          </div>
        </div>
        
        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Tingkat Kelas</h3>
            <button class="btn btn-primary" style="padding:6px 12px; font-size:12px;" id="btn-add-class">
              <i data-lucide="plus" style="width:14px; height:14px;"></i> Tambah
            </button>
          </div>
          <div class="card-body" style="padding:0;">
            <div class="structure-list-box" id="class-list-mount"></div>
          </div>
        </div>
      </div>
      
      <!-- 2. Column Bab & Topik (Dinamis bergantung pada seleksi Mapel & Kelas) -->
      <div style="display:flex; flex-direction:column; gap:24px;">
        <div class="card">
          <div class="card-header">
            <h3 class="card-title" id="bab-card-title">Bab (Chapters)</h3>
            <button class="btn btn-primary" style="padding:6px 12px; font-size:12px; display:none;" id="btn-add-chapter">
              <i data-lucide="plus" style="width:14px; height:14px;"></i> Tambah Bab
            </button>
          </div>
          <div class="card-body" style="padding:0;">
            <div style="padding:16px 20px; font-size:13px; color:var(--neutral-600); border-bottom:1px solid var(--neutral-400);" id="bab-selection-info">
              Silakan pilih Mata Pelajaran dan Tingkat Kelas di sebelah kiri terlebih dahulu.
            </div>
            <div class="structure-list-box" id="chapter-list-mount"></div>
          </div>
        </div>

        <div class="card">
          <div class="card-header">
            <h3 class="card-title" id="topik-card-title">Topik Sub-Bab</h3>
            <button class="btn btn-primary" style="padding:6px 12px; font-size:12px; display:none;" id="btn-add-topic">
              <i data-lucide="plus" style="width:14px; height:14px;"></i> Tambah Topik
            </button>
          </div>
          <div class="card-body" style="padding:0;">
            <div style="padding:16px 20px; font-size:13px; color:var(--neutral-600); border-bottom:1px solid var(--neutral-400);" id="topic-selection-info">
              Silakan pilih Bab di atas terlebih dahulu.
            </div>
            <div class="structure-list-box" id="topic-list-mount"></div>
          </div>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  let selectedSubId = null;
  let selectedClsId = null;
  let selectedChId = null;

  const reloadSubjects = () => {
    const list = document.getElementById('subject-list-mount');
    const items = db.get('subjects');
    list.innerHTML = items.map(s => `
      <div class="structure-item ${selectedSubId === s.id ? 'selected' : ''}" data-sub-id="${s.id}">
        <span style="font-weight:600;">${s.name}</span>
        <div style="display:inline-flex; gap:4px;" class="item-actions">
          <button class="btn btn-secondary btn-icon edit-sub" data-id="${s.id}" style="padding:4px;"><i data-lucide="edit" style="width:12px; height:12px;"></i></button>
          <button class="btn btn-danger btn-icon del-sub" data-id="${s.id}" style="padding:4px;"><i data-lucide="trash-2" style="width:12px; height:12px;"></i></button>
        </div>
      </div>
    `).join('') || '<div style="padding:20px; text-align:center; font-size:13px;">Belum ada Mata Pelajaran.</div>';

    lucide.createIcons();
    // Bind click
    list.querySelectorAll('.structure-item').forEach(el => {
      el.addEventListener('click', (e) => {
        if (e.target.closest('.item-actions')) return;
        selectedSubId = el.dataset.subId;
        reloadSubjects();
        reloadChapters();
      });
    });

    list.querySelectorAll('.edit-sub').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        const current = db.get('subjects').find(s => s.id === id);
        const name = prompt('Ubah Nama Mata Pelajaran:', current.name);
        if (name) {
          db.update('subjects', id, { name });
          reloadSubjects();
        }
      });
    });

    list.querySelectorAll('.del-sub').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        if (confirm('Hapus mata pelajaran ini? Semua Bab terkait akan terpengaruh.')) {
          db.delete('subjects', id);
          if (selectedSubId === id) selectedSubId = null;
          reloadSubjects();
          reloadChapters();
        }
      });
    });
  };

  const reloadClasses = () => {
    const list = document.getElementById('class-list-mount');
    const items = db.get('classes');
    list.innerHTML = items.map(c => `
      <div class="structure-item ${selectedClsId === c.id ? 'selected' : ''}" data-cls-id="${c.id}">
        <span style="font-weight:600;">Kelas ${c.name}</span>
        <div style="display:inline-flex; gap:4px;" class="item-actions">
          <button class="btn btn-secondary btn-icon edit-cls" data-id="${c.id}" style="padding:4px;"><i data-lucide="edit" style="width:12px; height:12px;"></i></button>
          <button class="btn btn-danger btn-icon del-cls" data-id="${c.id}" style="padding:4px;"><i data-lucide="trash-2" style="width:12px; height:12px;"></i></button>
        </div>
      </div>
    `).join('') || '<div style="padding:20px; text-align:center; font-size:13px;">Belum ada Tingkat Kelas.</div>';

    lucide.createIcons();

    list.querySelectorAll('.structure-item').forEach(el => {
      el.addEventListener('click', (e) => {
        if (e.target.closest('.item-actions')) return;
        selectedClsId = el.dataset.clsId;
        reloadClasses();
        reloadChapters();
      });
    });

    list.querySelectorAll('.edit-cls').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        const current = db.get('classes').find(c => c.id === id);
        const name = prompt('Ubah Nama Kelas:', current.name);
        if (name) {
          db.update('classes', id, { name });
          reloadClasses();
        }
      });
    });

    list.querySelectorAll('.del-cls').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        if (confirm('Hapus kelas ini?')) {
          db.delete('classes', id);
          if (selectedClsId === id) selectedClsId = null;
          reloadClasses();
          reloadChapters();
        }
      });
    });
  };

  const reloadChapters = () => {
    const list = document.getElementById('chapter-list-mount');
    const info = document.getElementById('bab-selection-info');
    const addBtn = document.getElementById('btn-add-chapter');

    if (!selectedSubId || !selectedClsId) {
      list.innerHTML = '';
      info.style.display = 'block';
      addBtn.style.display = 'none';
      reloadTopics();
      return;
    }

    const sub = db.get('subjects').find(s => s.id === selectedSubId);
    const cls = db.get('classes').find(c => c.id === selectedClsId);

    info.style.display = 'none';
    addBtn.style.display = 'inline-flex';
    document.getElementById('bab-card-title').innerText = `Bab - ${sub.name} (Kelas ${cls.name})`;

    const items = db.get('chapters').filter(c => c.subjectId === selectedSubId && c.classId === selectedClsId);
    list.innerHTML = items.map(c => `
      <div class="structure-item ${selectedChId === c.id ? 'selected' : ''}" data-ch-id="${c.id}">
        <span>${c.name}</span>
        <div style="display:inline-flex; gap:4px;" class="item-actions">
          <button class="btn btn-secondary btn-icon edit-ch" data-id="${c.id}" style="padding:4px;"><i data-lucide="edit" style="width:12px; height:12px;"></i></button>
          <button class="btn btn-danger btn-icon del-ch" data-id="${c.id}" style="padding:4px;"><i data-lucide="trash-2" style="width:12px; height:12px;"></i></button>
        </div>
      </div>
    `).join('') || '<div style="padding:20px; text-align:center; font-size:13px; color:var(--neutral-600);">Belum ada Bab untuk kombinasi ini.</div>';

    lucide.createIcons();

    list.querySelectorAll('.structure-item').forEach(el => {
      el.addEventListener('click', (e) => {
        if (e.target.closest('.item-actions')) return;
        selectedChId = el.dataset.chId;
        reloadChapters();
        reloadTopics();
      });
    });

    list.querySelectorAll('.edit-ch').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        const current = db.get('chapters').find(c => c.id === id);
        const name = prompt('Ubah Nama Bab:', current.name);
        if (name) {
          db.update('chapters', id, { name });
          reloadChapters();
        }
      });
    });

    list.querySelectorAll('.del-ch').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        if (confirm('Hapus Bab ini? Semua topik di dalamnya akan hilang.')) {
          db.delete('chapters', id);
          if (selectedChId === id) selectedChId = null;
          reloadChapters();
          reloadTopics();
        }
      });
    });
  };

  const reloadTopics = () => {
    const list = document.getElementById('topic-list-mount');
    const info = document.getElementById('topic-selection-info');
    const addBtn = document.getElementById('btn-add-topic');

    if (!selectedChId) {
      list.innerHTML = '';
      info.style.display = 'block';
      addBtn.style.display = 'none';
      return;
    }

    const ch = db.get('chapters').find(c => c.id === selectedChId);
    info.style.display = 'none';
    addBtn.style.display = 'inline-flex';
    document.getElementById('topik-card-title').innerText = `Topik Sub-Bab - ${ch.name}`;

    const items = db.get('topics').filter(t => t.chapterId === selectedChId);
    list.innerHTML = items.map(t => `
      <div class="structure-item" data-tp-id="${t.id}">
        <span>${t.name}</span>
        <div style="display:inline-flex; gap:4px;" class="item-actions">
          <button class="btn btn-secondary btn-icon edit-tp" data-id="${t.id}" style="padding:4px;"><i data-lucide="edit" style="width:12px; height:12px;"></i></button>
          <button class="btn btn-danger btn-icon del-tp" data-id="${t.id}" style="padding:4px;"><i data-lucide="trash-2" style="width:12px; height:12px;"></i></button>
        </div>
      </div>
    `).join('') || '<div style="padding:20px; text-align:center; font-size:13px; color:var(--neutral-600);">Belum ada Topik untuk Bab ini.</div>';

    lucide.createIcons();

    list.querySelectorAll('.edit-tp').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        const current = db.get('topics').find(t => t.id === id);
        const name = prompt('Ubah Nama Topik:', current.name);
        if (name) {
          db.update('topics', id, { name });
          reloadTopics();
        }
      });
    });

    list.querySelectorAll('.del-tp').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        if (confirm('Hapus Topik ini?')) {
          db.delete('topics', id);
          reloadTopics();
        }
      });
    });
  };

  // Bind Add Buttons
  document.getElementById('btn-add-subject').addEventListener('click', () => {
    const name = prompt('Masukkan Nama Mata Pelajaran Baru:');
    if (name) {
      db.insert('subjects', { name, schoolId: 'sch-1' });
      reloadSubjects();
    }
  });

  document.getElementById('btn-add-class').addEventListener('click', () => {
    const name = prompt('Masukkan Tingkat Kelas Baru (misal: X, XI, XII):');
    if (name) {
      db.insert('classes', { name, schoolId: 'sch-1' });
      reloadClasses();
    }
  });

  document.getElementById('btn-add-chapter').addEventListener('click', () => {
    if (!selectedSubId || !selectedClsId) return;
    const name = prompt('Masukkan Nama Bab Baru:');
    if (name) {
      db.insert('chapters', { name, subjectId: selectedSubId, classId: selectedClsId });
      reloadChapters();
    }
  });

  document.getElementById('btn-add-topic').addEventListener('click', () => {
    if (!selectedChId) return;
    const name = prompt('Masukkan Nama Topik Sub-Bab Baru:');
    if (name) {
      db.insert('topics', { name, chapterId: selectedChId });
      reloadTopics();
    }
  });

  // Initial Load
  reloadSubjects();
  reloadClasses();
}

// ----------------------------------------------------
// 7. PAGE: BANK SOAL LIST
// ----------------------------------------------------
function renderBankSoal(mount) {
  const subjects = db.get('subjects');
  const classes = db.get('classes');
  const chapters = db.get('chapters');
  const questions = db.get('questions');
  const bookmarks = db.get('bookmarks');

  // Sidebar filter state
  let filterKeyword = '';
  let filterSubject = '';
  let filterDifficulty = '';
  let filterType = '';
  let filterStatus = '';

  mount.innerHTML = `
    <!-- Top actions row -->
    <div class="action-row">
      <div class="search-filter-box">
        <input type="text" id="soal-search" class="form-input" style="flex:1;" placeholder="Cari kode soal, pertanyaan, tag...">
        <select id="filter-mapel" class="form-input" style="width:160px;">
          <option value="">Semua Mapel</option>
          ${subjects.map(s => `<option value="${s.id}">${s.name}</option>`).join('')}
        </select>
        <select id="filter-tingkat" class="form-input" style="width:110px;">
          <option value="">Semua Tipe</option>
          <option value="PG">Pilihan Ganda</option>
          <option value="PG_KOMPLEKS">PG Kompleks</option>
          <option value="BENAR_SALAH">Benar Salah</option>
          <option value="MENJODOHKAN">Menjodohkan</option>
          <option value="ISIAN_SINGKAT">Isian Singkat</option>
          <option value="ESSAY">Esai</option>
          <option value="NUMERIK">Numerik</option>
        </select>
        <select id="filter-status" class="form-input" style="width:120px;">
          <option value="">Semua Status</option>
          <option value="DRAFT">DRAFT</option>
          <option value="REVIEW">REVIEW</option>
          <option value="APPROVED">APPROVED</option>
          <option value="REJECTED">REJECTED</option>
        </select>
      </div>
      
      <div style="display:inline-flex; gap:12px;">
        <a class="btn btn-primary" href="#soal?action=tambah">
          <i data-lucide="plus"></i>
          <span>Tulis Soal Baru</span>
        </a>
      </div>
    </div>

    <!-- Questions list wrapper -->
    <div id="questions-list-mount"></div>
  `;

  lucide.createIcons();

  const listMount = document.getElementById('questions-list-mount');

  const reloadList = () => {
    let list = [...questions];

    // Apply filters
    if (filterKeyword) {
      const kw = filterKeyword.toLowerCase();
      list = list.filter(q =>
        (q.code || '').toLowerCase().includes(kw) ||
        (q.questionText || '').toLowerCase().includes(kw) ||
        (q.tag || '').toLowerCase().includes(kw)
      );
    }

    if (filterSubject) {
      list = list.filter(q => q.subjectId === filterSubject);
    }

    if (filterType) {
      list = list.filter(q => q.type === filterType);
    }

    if (filterStatus) {
      list = list.filter(q => q.status === filterStatus);
    }

    listMount.innerHTML = list.map(q => {
      const mapel = subjects.find(s => s.id === q.subjectId)?.name || 'Umum';
      const kls = classes.find(c => c.id === q.classId)?.name || 'X';
      const bab = chapters.find(c => c.id === q.chapterId)?.name || 'Umum';

      const isBookmarked = bookmarks.some(b => b.userId === activeUser.id && b.questionId === q.id);
      const starIcon = isBookmarked ? 'star-off' : 'star';
      const starClass = isBookmarked ? 'color: #ecc94b; fill: #ecc94b;' : '';

      const statusClass = q.status === 'APPROVED' ? 'success' : q.status === 'REVIEW' ? 'warning' : q.status === 'REJECTED' ? 'danger' : 'neutral';

      // Options html for MC
      let choicesHtml = '';
      if ((q.type === 'PG' || q.type === 'PG_KOMPLEKS') && q.choices) {
        choicesHtml = `
          <div class="choices-preview-list">
            ${Object.entries(q.choices).map(([key, text]) => {
          const isCorrect = Array.isArray(q.correctAnswer) ? q.correctAnswer.includes(key) : q.correctAnswer === key;
          const correctClass = isCorrect ? 'correct' : '';
          return `<div class="choice-preview-item ${correctClass}"><strong>${key}.</strong> <span>${text}</span></div>`;
        }).join('')}
          </div>
        `;
      } else if (q.type === 'BENAR_SALAH') {
        choicesHtml = `
          <div class="choices-preview-list">
            <div class="choice-preview-item ${q.correctAnswer === 'BENAR' ? 'correct' : ''}">BENAR</div>
            <div class="choice-preview-item ${q.correctAnswer === 'SALAH' ? 'correct' : ''}">SALAH</div>
          </div>
        `;
      } else if (q.type === 'MENJODOHKAN' && q.choices) {
        choicesHtml = `
          <div style="font-size:12px; margin-top:8px;">
            <strong>Menjodohkan:</strong>
            ${(q.choices.premises || []).map((p, i) => `<div style="padding:2px 0;">${p} &harr; ${q.choices.responses?.[i] || ''}</div>`).join('')}
          </div>
        `;
      } else {
        choicesHtml = `<div style="font-size:12px; margin-top:8px; color: var(--neutral-700);"><strong>Jawaban Benar:</strong> ${q.correctAnswer}</div>`;
      }

      // Review flow action buttons
      let reviewActions = '';
      if (activeUser.role === 'REVIEWER' && q.status === 'REVIEW') {
        reviewActions = `
          <button class="btn btn-primary approve-soal" style="padding:6px 12px; font-size:12px;" data-id="${q.id}">
            <i data-lucide="check" style="width:14px; height:14px;"></i> Setujui
          </button>
          <button class="btn btn-danger reject-soal" style="padding:6px 12px; font-size:12px;" data-id="${q.id}">
            <i data-lucide="x" style="width:14px; height:14px;"></i> Tolak
          </button>
        `;
      } else if (activeUser.role === 'GURU' && q.status === 'DRAFT') {
        reviewActions = `
          <button class="btn btn-primary submit-review" style="padding:6px 12px; font-size:12px;" data-id="${q.id}">
            <i data-lucide="send" style="width:14px; height:14px;"></i> Ajukan Review
          </button>
        `;
      }

      return `
        <div class="question-item-card">
          <div class="question-card-header">
            <div class="q-meta-badges">
              <span style="font-weight:800; color:var(--primary); font-size:13px;">${q.code}</span>
              <span class="badge badge-primary">${mapel}</span>
              <span class="badge badge-neutral">Kelas ${kls}</span>
              <span class="badge badge-neutral">${q.type}</span>
              <span class="badge badge-neutral">${q.difficulty}</span>
              <span class="badge badge-${statusClass}">${q.status}</span>
            </div>
            
            <div style="display:inline-flex; gap:8px;">
              <button class="btn btn-secondary btn-icon toggle-fav" data-id="${q.id}" style="padding:6px;" title="Bookmark">
                <i data-lucide="${starIcon}" style="width:16px; height:16px; ${starClass}"></i>
              </button>
              <a class="btn btn-secondary btn-icon" style="padding:6px;" href="#soal?action=edit&id=${q.id}" title="Edit">
                <i data-lucide="edit" style="width:16px; height:16px;"></i>
              </a>
              <button class="btn btn-danger btn-icon delete-soal" data-id="${q.id}" style="padding:6px;" title="Hapus">
                <i data-lucide="trash-2" style="width:16px; height:16px;"></i>
              </button>
            </div>
          </div>
          
          <div class="q-body-preview">${q.questionText}</div>
          
          ${choicesHtml}

          <!-- Footer/Aktivitas -->
          <div style="display:flex; justify-content:space-between; align-items:center; border-top:1px solid var(--neutral-400); padding-top:12px; margin-top:8px; font-size:11px; color: var(--neutral-600);">
            <div>Dibuat oleh: <strong>${q.creatorName || 'Guru'}</strong> | Bab: ${bab}</div>
            <div style="display:inline-flex; gap:8px;">
              ${reviewActions}
            </div>
          </div>
        </div>
      `;
    }).join('') || '<div class="card" style="padding:40px; text-align:center; color: var(--neutral-600);">Tidak ada soal yang cocok dengan filter pencarian.</div>';

    lucide.createIcons();

    // LaTeX rendering trigger
    if (window.MathJax && window.MathJax.typesetPromise) {
      window.MathJax.typesetPromise([listMount]).catch(err => console.error(err));
    }

    // Bind item action events
    listMount.querySelectorAll('.toggle-fav').forEach(btn => {
      btn.addEventListener('click', () => {
        const qId = btn.dataset.id;
        let bms = db.get('bookmarks');
        const idx = bms.findIndex(b => b.userId === activeUser.id && b.questionId === qId);
        if (idx !== -1) {
          bms.splice(idx, 1);
          db.save('bookmarks', bms);
          showToast('Bookmark soal dihapus.', 'info');
        } else {
          bms.push({ id: `bm-${Date.now()}`, userId: activeUser.id, questionId: qId });
          db.save('bookmarks', bms);
          showToast('Soal disimpan ke favorit.', 'success');
        }
        reloadList();
      });
    });

    listMount.querySelectorAll('.delete-soal').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        if (confirm('Apakah Anda yakin ingin menghapus soal ini?')) {
          db.delete('questions', id);
          showToast('Soal berhasil dihapus.', 'success');
          reloadList();
        }
      });
    });

    // Review flow handlers
    listMount.querySelectorAll('.submit-review').forEach(btn => {
      btn.addEventListener('click', () => {
        db.update('questions', btn.dataset.id, { status: 'REVIEW' });
        showToast('Soal diajukan untuk direview oleh Reviewer.', 'info');
        reloadList();
      });
    });

    listMount.querySelectorAll('.approve-soal').forEach(btn => {
      btn.addEventListener('click', () => {
        db.update('questions', btn.dataset.id, { status: 'APPROVED', reviewerId: activeUser.id, reviewerName: activeUser.name });
        showToast('Kualitas soal disetujui.', 'success');
        reloadList();
      });
    });

    listMount.querySelectorAll('.reject-soal').forEach(btn => {
      btn.addEventListener('click', () => {
        const feedback = prompt('Masukkan alasan penolakan/revisi:', 'Perlu revisi teks');
        if (feedback !== null) {
          db.update('questions', btn.dataset.id, { status: 'REJECTED', discussion: (db.get('questions').find(q => q.id === btn.dataset.id).discussion || '') + `<p style="color:red">Catatan Reviewer: ${feedback}</p>` });
          showToast('Soal dikembalikan untuk direvisi.', 'error');
          reloadList();
        }
      });
    });
  };

  // Bind filtering inputs
  document.getElementById('soal-search').addEventListener('input', (e) => {
    filterKeyword = e.target.value;
    reloadList();
  });
  document.getElementById('filter-mapel').addEventListener('change', (e) => {
    filterSubject = e.target.value;
    reloadList();
  });
  document.getElementById('filter-tingkat').addEventListener('change', (e) => {
    filterType = e.target.value;
    reloadList();
  });
  document.getElementById('filter-status').addEventListener('change', (e) => {
    filterStatus = e.target.value;
    reloadList();
  });

  // Run initial list render
  reloadList();
}

// ----------------------------------------------------
// 8. PAGE: CREATE/EDIT QUESTION FORM
// ----------------------------------------------------
function renderFormSoal(mount, questionId = null) {
  const isEdit = !!questionId;
  const q = isEdit ? db.get('questions').find(soal => soal.id === questionId) : {};

  const subjects = db.get('subjects');
  const classes = db.get('classes');

  mount.innerHTML = `
    <div style="display:grid; grid-template-columns: 2fr 1.2fr; gap:28px;">
      <!-- Input Column -->
      <form id="soal-form" style="display:flex; flex-direction:column; gap:20px;">
        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Meta Kurikulum & Kriteria</h3>
          </div>
          <div class="card-body" style="display:flex; flex-direction:column; gap:16px;">
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
              <div class="form-group">
                <label class="form-label" for="q-sub">Mata Pelajaran</label>
                <select id="q-sub" class="form-input" required>
                  <option value="">Pilih Mapel</option>
                  ${subjects.map(s => `<option value="${s.id}" ${q.subjectId === s.id ? 'selected' : ''}>${s.name}</option>`).join('')}
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="q-cls">Tingkat Kelas</label>
                <select id="q-cls" class="form-input" required>
                  <option value="">Pilih Kelas</option>
                  ${classes.map(c => `<option value="${c.id}" ${q.classId === c.id ? 'selected' : ''}>Kelas ${c.name}</option>`).join('')}
                </select>
              </div>
            </div>

            <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
              <div class="form-group">
                <label class="form-label" for="q-ch">Bab (Pilih Mapel & Kelas dahulu)</label>
                <select id="q-ch" class="form-input" required>
                  <option value="">Pilih Bab</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="q-tp">Topik Sub-Bab</label>
                <select id="q-tp" class="form-input">
                  <option value="">Pilih Topik</option>
                </select>
              </div>
            </div>

            <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:16px;">
              <div class="form-group">
                <label class="form-label" for="q-code">Kode Soal (Unik)</label>
                <input type="text" id="q-code" class="form-input" placeholder="INF-X-001" value="${q.code || ''}" required>
              </div>
              <div class="form-group">
                <label class="form-label" for="q-type">Tipe Soal</label>
                <select id="q-type" class="form-input" required>
                  <option value="PG" ${q.type === 'PG' ? 'selected' : ''}>Pilihan Ganda</option>
                  <option value="PG_KOMPLEKS" ${q.type === 'PG_KOMPLEKS' ? 'selected' : ''}>Pilihan Ganda Kompleks</option>
                  <option value="BENAR_SALAH" ${q.type === 'BENAR_SALAH' ? 'selected' : ''}>Benar/Salah</option>
                  <option value="MENJODOHKAN" ${q.type === 'MENJODOHKAN' ? 'selected' : ''}>Menjodohkan</option>
                  <option value="ISIAN_SINGKAT" ${q.type === 'ISIAN_SINGKAT' ? 'selected' : ''}>Isian Singkat</option>
                  <option value="ESSAY" ${q.type === 'ESSAY' ? 'selected' : ''}>Esai</option>
                  <option value="NUMERIK" ${q.type === 'NUMERIK' ? 'selected' : ''}>Numerik</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="q-diff">Tingkat Kesulitan</label>
                <select id="q-diff" class="form-input" required>
                  <option value="MUDAH" ${q.difficulty === 'MUDAH' ? 'selected' : ''}>Mudah</option>
                  <option value="SEDANG" ${q.difficulty === 'SEDANG' ? 'selected' : ''}>Sedang</option>
                  <option value="SULIT" ${q.difficulty === 'SULIT' ? 'selected' : ''}>Sulit</option>
                </select>
              </div>
            </div>
            
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
              <div class="form-group">
                <label class="form-label" for="q-comp">Kompetensi Inti/Dasar</label>
                <input type="text" id="q-comp" class="form-input" placeholder="contoh: Menganalisis Logika" value="${q.competence || ''}">
              </div>
              <div class="form-group">
                <label class="form-label" for="q-goal">Tujuan Pembelajaran</label>
                <input type="text" id="q-goal" class="form-input" placeholder="contoh: Siswa dapat melukis algoritma..." value="${q.learningGoal || ''}">
              </div>
            </div>
          </div>
        </div>

        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Teks Pertanyaan (Visual Editor)</h3>
          </div>
          <div class="card-body">
            <div class="rich-editor-mock">
              <div class="editor-toolbar">
                <button class="toolbar-btn" data-editor-cmd="bold" title="Tebal"><i data-lucide="bold" style="width:14px; height:14px;"></i></button>
                <button class="toolbar-btn" data-editor-cmd="italic" title="Miring"><i data-lucide="italic" style="width:14px; height:14px;"></i></button>
                <span class="toolbar-divider"></span>
                <button class="toolbar-btn" data-editor-cmd="latex-inline" title="LaTeX Inline ($)"><i data-lucide="sigma" style="width:14px; height:14px;"></i></button>
                <button class="toolbar-btn" data-editor-cmd="latex-display" title="LaTeX Display ($$)"><i data-lucide="superscript" style="width:14px; height:14px;"></i></button>
                <span class="toolbar-divider"></span>
                <button class="toolbar-btn" data-editor-cmd="table" title="Tabel"><i data-lucide="table" style="width:14px; height:14px;"></i></button>
                <button class="toolbar-btn" data-editor-cmd="image" title="Gambar Link"><i data-lucide="image" style="width:14px; height:14px;"></i></button>
              </div>
              <textarea id="q-text" class="editor-textarea" required>${q.questionText || ''}</textarea>
            </div>
          </div>
        </div>

        <!-- Dynamic Answer choices block -->
        <div class="card" id="answers-card-mount">
          <div class="card-header">
            <h3 class="card-title">Konfigurasi Pilihan Jawaban & Kunci</h3>
          </div>
          <div class="card-body" id="answers-inputs-mount"></div>
        </div>

        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Pembahasan & Referensi Soal</h3>
          </div>
          <div class="card-body" style="display:flex; flex-direction:column; gap:16px;">
            <div class="form-group">
              <label class="form-label" for="q-disc">Teks Pembahasan (Mendukung LaTeX)</label>
              <textarea id="q-disc" class="form-input" style="height:100px;">${q.discussion || ''}</textarea>
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
              <div class="form-group">
                <label class="form-label" for="q-ref">Referensi Sumber Soal</label>
                <input type="text" id="q-ref" class="form-input" placeholder="contoh: Buku Kemdikbud Hlm 10" value="${q.references || ''}">
              </div>
              <div class="form-group">
                <label class="form-label" for="q-tag">Tag / Label (Pisahkan dengan koma)</label>
                <input type="text" id="q-tag" class="form-input" placeholder="contoh: aljabar, ujian, mudah" value="${q.tag || ''}">
              </div>
            </div>
          </div>
        </div>

        <!-- Action Row -->
        <div style="display:inline-flex; gap:12px; margin-bottom:40px;">
          <button type="submit" class="btn btn-primary">
            <i data-lucide="save"></i>
            <span>Simpan ke Bank Soal</span>
          </button>
          <a class="btn btn-secondary" href="#soal">Batal</a>
        </div>
      </form>

      <!-- Real-Time Rendering Preview Column -->
      <div style="position: sticky; top: var(--header-height); height: calc(100vh - 100px); overflow-y: auto;">
        <div class="card" style="border-color: var(--primary);">
          <div class="card-header" style="background-color: var(--primary-light);">
            <h3 class="card-title" style="color:var(--primary); font-weight:800;">Real-Time Preview</h3>
            <span class="badge badge-primary">Standard Siswa</span>
          </div>
          <div class="card-body" style="background:#fff;">
            <div id="preview-question-text" style="font-size: 15px; margin-bottom: 20px; line-height: 1.5;"></div>
            <div id="preview-choices-list"></div>
            <div id="preview-discussion-box" style="margin-top:20px; padding-top:15px; border-top:1px dashed var(--neutral-400); display:none;"></div>
          </div>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  // Elements mapping
  const form = document.getElementById('soal-form');
  const subSelect = document.getElementById('q-sub');
  const clsSelect = document.getElementById('q-cls');
  const chSelect = document.getElementById('q-ch');
  const tpSelect = document.getElementById('q-tp');
  const typeSelect = document.getElementById('q-type');

  // Load interactive dynamic selects for Chapters & Topics
  const updateChaptersAndTopics = (defaultChId = null, defaultTpId = null) => {
    const sId = subSelect.value;
    const cId = clsSelect.value;

    chSelect.innerHTML = '<option value="">Pilih Bab</option>';
    tpSelect.innerHTML = '<option value="">Pilih Topik</option>';

    if (sId && cId) {
      const chapters = db.get('chapters').filter(c => c.subjectId === sId && c.classId === cId);
      chapters.forEach(ch => {
        const isSelected = ch.id === defaultChId ? 'selected' : '';
        chSelect.innerHTML += `<option value="${ch.id}" ${isSelected}>${ch.name}</option>`;
      });

      if (defaultChId) {
        const topics = db.get('topics').filter(t => t.chapterId === defaultChId);
        topics.forEach(tp => {
          const isSelected = tp.id === defaultTpId ? 'selected' : '';
          tpSelect.innerHTML += `<option value="${tp.id}" ${isSelected}>${tp.name}</option>`;
        });
      }
    }
  };

  subSelect.addEventListener('change', () => updateChaptersAndTopics());
  clsSelect.addEventListener('change', () => updateChaptersAndTopics());
  chSelect.addEventListener('change', () => {
    const chId = chSelect.value;
    tpSelect.innerHTML = '<option value="">Pilih Topik</option>';
    if (chId) {
      const topics = db.get('topics').filter(t => t.chapterId === chId);
      topics.forEach(tp => {
        tpSelect.innerHTML += `<option value="${tp.id}">${tp.name}</option>`;
      });
    }
  });

  // Pre-load chapters if edit
  if (isEdit) {
    updateChaptersAndTopics(q.chapterId, q.topicId);
  }

  // Answer Choices UI controller based on Question Type selected
  const updateChoicesForm = () => {
    const type = typeSelect.value;
    const mount = document.getElementById('answers-inputs-mount');

    let html = '';

    if (type === 'PG') {
      html = `
        <div class="form-group">
          <label class="form-label">Tentukan Pilihan & Kunci Jawaban Benar (Radio)</label>
          <div style="display:flex; flex-direction:column; gap:12px; margin-top:8px;">
            ${['A', 'B', 'C', 'D', 'E'].map(l => {
        const text = q.choices?.[l] || '';
        const isChecked = q.correctAnswer === l ? 'checked' : '';
        return `
                <div style="display:flex; align-items:center; gap:12px;">
                  <input type="radio" name="pg-correct" value="${l}" ${isChecked} required>
                  <strong style="width:20px;">${l}.</strong>
                  <input type="text" class="form-input pg-choice-input" data-choice-key="${l}" value="${text}" placeholder="Teks opsi pilihan ${l}" required style="flex:1;">
                </div>
              `;
      }).join('')}
          </div>
        </div>
      `;
    } else if (type === 'PG_KOMPLEKS') {
      html = `
        <div class="form-group">
          <label class="form-label">Tentukan Pilihan & Centang Semua Kunci Jawaban Benar (Checkbox)</label>
          <div style="display:flex; flex-direction:column; gap:12px; margin-top:8px;">
            ${['A', 'B', 'C', 'D', 'E'].map(l => {
        const text = q.choices?.[l] || '';
        const isChecked = Array.isArray(q.correctAnswer) && q.correctAnswer.includes(l) ? 'checked' : '';
        return `
                <div style="display:flex; align-items:center; gap:12px;">
                  <input type="checkbox" class="pg-complex-chk" value="${l}" ${isChecked}>
                  <strong style="width:20px;">${l}.</strong>
                  <input type="text" class="form-input pg-complex-choice-input" data-choice-key="${l}" value="${text}" placeholder="Teks opsi pilihan ${l}" style="flex:1;">
                </div>
              `;
      }).join('')}
          </div>
        </div>
      `;
    } else if (type === 'BENAR_SALAH') {
      const isBenar = q.correctAnswer === 'BENAR';
      html = `
        <div class="form-group">
          <label class="form-label">Pilih Jawaban Benar</label>
          <div style="display:flex; gap:24px; margin-top:8px;">
            <label style="display:flex; align-items:center; gap:8px; cursor:pointer;">
              <input type="radio" name="bs-correct" value="BENAR" ${isBenar ? 'checked' : ''} required>
              <span>BENAR</span>
            </label>
            <label style="display:flex; align-items:center; gap:8px; cursor:pointer;">
              <input type="radio" name="bs-correct" value="SALAH" ${!isBenar ? 'checked' : ''} required>
              <span>SALAH</span>
            </label>
          </div>
        </div>
      `;
    } else if (type === 'MENJODOHKAN') {
      // Premises and responses fields list
      let rows = '';
      const premises = q.choices?.premises || ['', ''];
      const responses = q.choices?.responses || ['', ''];
      const maxRows = Math.max(premises.length, responses.length, 2);

      for (let i = 0; i < maxRows; i++) {
        rows += `
          <div class="matching-pair-row">
            <input type="text" class="form-input match-premise" placeholder="Pernyataan Kiri" value="${premises[i] || ''}" required>
            <input type="text" class="form-input match-response" placeholder="Jawaban Kanan" value="${responses[i] || ''}" required>
            <button type="button" class="btn btn-danger btn-icon remove-match-row" style="padding:8px;"><i data-lucide="trash-2" style="width:14px; height:14px;"></i></button>
          </div>
        `;
      }

      html = `
        <div class="form-group">
          <label class="form-label">Daftar Jodoh Pernyataan</label>
          <div style="display:flex; flex-direction:column; gap:4px; margin-top:8px;" id="matching-rows-mount">
            ${rows}
          </div>
          <button type="button" class="btn btn-secondary" style="padding:6px 12px; font-size:12px; margin-top:8px;" id="btn-add-match-row">
            <i data-lucide="plus" style="width:12px; height:12px;"></i> Tambah Baris
          </button>
        </div>
      `;
    } else if (type === 'ISIAN_SINGKAT') {
      html = `
        <div class="form-group">
          <label class="form-label" for="q-ans-txt">Teks Kunci Jawaban Singkat</label>
          <input type="text" id="q-ans-txt" class="form-input" placeholder="contoh: HTTPS atau 5" value="${q.correctAnswer || ''}" required>
        </div>
      `;
    } else if (type === 'ESSAY' || type === 'URAIAN') {
      html = `
        <div class="form-group">
          <label class="form-label" for="q-ans-essay">Pedoman Penskoran / Kunci Jawaban Esai</label>
          <textarea id="q-ans-essay" class="form-input" style="height:100px;" placeholder="Tuliskan indikator jawaban atau paragraf jawaban yang benar" required>${q.correctAnswer || ''}</textarea>
        </div>
      `;
    } else if (type === 'NUMERIK') {
      html = `
        <div class="form-group">
          <label class="form-label" for="q-ans-num">Kunci Jawaban Angka Eksak</label>
          <input type="number" step="any" id="q-ans-num" class="form-input" placeholder="contoh: 2.5 atau -10" value="${q.correctAnswer || ''}" required>
        </div>
      `;
    }

    mount.innerHTML = html;
    lucide.createIcons();

    // Bind matching row buttons if type MENJODOHKAN
    if (type === 'MENJODOHKAN') {
      const rowsMount = document.getElementById('matching-rows-mount');
      const addRowBtn = document.getElementById('btn-add-match-row');

      addRowBtn.addEventListener('click', () => {
        const div = document.createElement('div');
        div.className = 'matching-pair-row';
        div.innerHTML = `
          <input type="text" class="form-input match-premise" placeholder="Pernyataan Kiri" required>
          <input type="text" class="form-input match-response" placeholder="Jawaban Kanan" required>
          <button type="button" class="btn btn-danger btn-icon remove-match-row" style="padding:8px;"><i data-lucide="trash-2" style="width:14px; height:14px;"></i></button>
        `;
        rowsMount.appendChild(div);
        lucide.createIcons();
        div.querySelector('.remove-match-row').addEventListener('click', () => div.remove());
        updateLivePreview();
      });

      rowsMount.querySelectorAll('.remove-match-row').forEach(btn => {
        btn.addEventListener('click', () => {
          btn.closest('.matching-pair-row').remove();
          updateLivePreview();
        });
      });

      rowsMount.querySelectorAll('input').forEach(inp => {
        inp.addEventListener('input', updateLivePreview);
      });
    }

    // Bind change listeners to re-render preview
    mount.querySelectorAll('input, select, textarea').forEach(inp => {
      inp.addEventListener('input', updateLivePreview);
    });

    updateLivePreview();
  };

  typeSelect.addEventListener('change', updateChoicesForm);

  // Live Preview Builder
  const updateLivePreview = () => {
    const qText = document.getElementById('q-text').value;
    const type = typeSelect.value;

    const textPreview = document.getElementById('preview-question-text');
    const choicesPreview = document.getElementById('preview-choices-list');
    const discPreview = document.getElementById('preview-discussion-box');

    textPreview.innerHTML = qText || '<span style="color:#aaa;">Ketik teks pertanyaan untuk melihat pratinjau...</span>';
    choicesPreview.innerHTML = '';

    // Preview Choices list
    if (type === 'PG') {
      const optionInputs = document.querySelectorAll('.pg-choice-input');
      const correctVal = document.querySelector('input[name="pg-correct"]:checked')?.value;

      optionInputs.forEach(inp => {
        const key = inp.dataset.choiceKey;
        const val = inp.value;
        if (val) {
          const isCorrect = key === correctVal;
          const style = isCorrect ? 'border-color:var(--success); background:var(--success-light); font-weight:600;' : '';
          choicesPreview.innerHTML += `
            <div style="padding:8px 12px; border:1px solid var(--neutral-400); border-radius:var(--radius-sm); margin-bottom:6px; font-size:13px; display:flex; gap:8px; ${style}">
              <span>${key}.</span> <span>${val}</span>
            </div>
          `;
        }
      });
    } else if (type === 'PG_KOMPLEKS') {
      const optionInputs = document.querySelectorAll('.pg-complex-choice-input');
      const correctVals = Array.from(document.querySelectorAll('.pg-complex-chk:checked')).map(chk => chk.value);

      optionInputs.forEach(inp => {
        const key = inp.dataset.choiceKey;
        const val = inp.value;
        if (val) {
          const isCorrect = correctVals.includes(key);
          const style = isCorrect ? 'border-color:var(--success); background:var(--success-light); font-weight:600;' : '';
          choicesPreview.innerHTML += `
            <div style="padding:8px 12px; border:1px solid var(--neutral-400); border-radius:var(--radius-sm); margin-bottom:6px; font-size:13px; display:flex; gap:8px; ${style}">
              <span>[ ${isCorrect ? 'x' : ' '} ] ${key}.</span> <span>${val}</span>
            </div>
          `;
        }
      });
    } else if (type === 'BENAR_SALAH') {
      const bsVal = document.querySelector('input[name="bs-correct"]:checked')?.value;
      choicesPreview.innerHTML = `
        <div style="display:flex; gap:12px;">
          <div style="padding:8px 16px; border:1px solid var(--neutral-400); border-radius:var(--radius-sm); font-size:13px; ${bsVal === 'BENAR' ? 'border-color:var(--success); background:var(--success-light); font-weight:600;' : ''}">BENAR</div>
          <div style="padding:8px 16px; border:1px solid var(--neutral-400); border-radius:var(--radius-sm); font-size:13px; ${bsVal === 'SALAH' ? 'border-color:var(--success); background:var(--success-light); font-weight:600;' : ''}">SALAH</div>
        </div>
      `;
    } else if (type === 'MENJODOHKAN') {
      const premises = Array.from(document.querySelectorAll('.match-premise')).map(i => i.value).filter(Boolean);
      const responses = Array.from(document.querySelectorAll('.match-response')).map(i => i.value).filter(Boolean);
      if (premises.length > 0) {
        let tableRows = '';
        const maxLen = Math.max(premises.length, responses.length);
        for (let i = 0; i < maxLen; i++) {
          tableRows += `
            <tr>
              <td style="border:1px solid var(--neutral-400); padding:6px 10px;">${premises[i] || ''}</td>
              <td style="border:1px solid var(--neutral-400); padding:6px 10px;">${responses[i] || ''}</td>
            </tr>
          `;
        }
        choicesPreview.innerHTML = `
          <table style="width:100%; border-collapse:collapse; font-size:13px;">
            <thead>
              <tr style="background:#f2f2f2;">
                <th style="border:1px solid var(--neutral-400); padding:6px 10px; text-align:left;">Pernyataan</th>
                <th style="border:1px solid var(--neutral-400); padding:6px 10px; text-align:left;">Jawaban</th>
              </tr>
            </thead>
            <tbody>${tableRows}</tbody>
          </table>
        `;
      }
    } else if (type === 'ISIAN_SINGKAT') {
      const ansVal = document.getElementById('q-ans-txt')?.value;
      choicesPreview.innerHTML = `<div style="font-size:13px;"><strong>Kunci Isian:</strong> <span style="background:var(--success-light); padding:2px 8px; border-radius:4px; color:var(--success-text); border:1px solid var(--success);">${ansVal || '...'}</span></div>`;
    } else if (type === 'NUMERIK') {
      const ansVal = document.getElementById('q-ans-num')?.value;
      choicesPreview.innerHTML = `<div style="font-size:13px;"><strong>Kunci Angka:</strong> <span style="background:var(--success-light); padding:2px 8px; border-radius:4px; color:var(--success-text); border:1px solid var(--success);">${ansVal || '...'}</span></div>`;
    } else if (type === 'ESSAY' || type === 'URAIAN') {
      const ansVal = document.getElementById('q-ans-essay')?.value;
      choicesPreview.innerHTML = `<div style="font-size:13px; font-style:italic; border-left:3px solid var(--neutral-600); padding-left:10px;"><strong>Pedoman Jawaban:</strong> <br>${ansVal || '...'}</div>`;
    }

    const qDisc = document.getElementById('q-disc').value;
    if (qDisc) {
      discPreview.style.display = 'block';
      discPreview.innerHTML = `<strong>Pembahasan Soal:</strong> <br> ${qDisc}`;
    } else {
      discPreview.style.display = 'none';
    }

    // Typeset MathJax
    if (window.MathJax && window.MathJax.typesetPromise) {
      window.MathJax.typesetPromise([textPreview, choicesPreview, discPreview]).catch(err => console.error(err));
    }
  };

  // Run initial tools binder
  setupVisualEditor('q-text', 'preview-question-text');
  document.getElementById('q-disc').addEventListener('input', updateLivePreview);

  // Render choices inputs initially
  updateChoicesForm();

  // Handle Form submit
  form.addEventListener('submit', (e) => {
    e.preventDefault();

    const type = typeSelect.value;
    let choices = {};
    let correctAnswer = '';

    if (type === 'PG') {
      correctAnswer = document.querySelector('input[name="pg-correct"]:checked').value;
      document.querySelectorAll('.pg-choice-input').forEach(inp => {
        choices[inp.dataset.choiceKey] = inp.value;
      });
    } else if (type === 'PG_KOMPLEKS') {
      correctAnswer = Array.from(document.querySelectorAll('.pg-complex-chk:checked')).map(chk => chk.value);
      if (correctAnswer.length === 0) {
        showToast('Pilih minimal 1 kunci jawaban benar untuk PG Kompleks.', 'error');
        return;
      }
      document.querySelectorAll('.pg-complex-choice-input').forEach(inp => {
        choices[inp.dataset.choiceKey] = inp.value;
      });
    } else if (type === 'BENAR_SALAH') {
      correctAnswer = document.querySelector('input[name="bs-correct"]:checked').value;
      choices = { A: 'BENAR', B: 'SALAH' };
    } else if (type === 'MENJODOHKAN') {
      const premises = Array.from(document.querySelectorAll('.match-premise')).map(i => i.value).filter(Boolean);
      const responses = Array.from(document.querySelectorAll('.match-response')).map(i => i.value).filter(Boolean);
      if (premises.length === 0) {
        showToast('Tambahkan minimal 1 baris pernyataan untuk jodoh.', 'error');
        return;
      }
      choices = { premises, responses };

      // Correct answers list maps premises directly to responses
      correctAnswer = premises.map((p, i) => ({ premise: p, response: responses[i] || '' }));
    } else if (type === 'ISIAN_SINGKAT') {
      correctAnswer = document.getElementById('q-ans-txt').value;
    } else if (type === 'NUMERIK') {
      correctAnswer = document.getElementById('q-ans-num').value;
    } else if (type === 'ESSAY' || type === 'URAIAN') {
      correctAnswer = document.getElementById('q-ans-essay').value;
    }

    const data = {
      code: document.getElementById('q-code').value,
      type,
      difficulty: document.getElementById('q-diff').value,
      questionText: document.getElementById('q-text').value,
      choices,
      correctAnswer,
      discussion: document.getElementById('q-disc').value,
      references: document.getElementById('q-ref').value,
      tag: document.getElementById('q-tag').value,
      competence: document.getElementById('q-comp').value,
      learningGoal: document.getElementById('q-goal').value,
      subjectId: subSelect.value,
      classId: clsSelect.value,
      chapterId: chSelect.value,
      topicId: tpSelect.value,
      schoolId: 'sch-1',
      status: 'DRAFT', // Reverts to Draft on change/creation
      creatorId: activeUser.id,
      creatorName: activeUser.name,
      createdAt: new Date().toISOString()
    };

    if (isEdit) {
      db.update('questions', q.id, data);
      showToast('Soal berhasil diperbarui.', 'success');
    } else {
      // Check for code uniqueness
      const existing = db.get('questions').find(soal => soal.code === data.code);
      if (existing) {
        showToast(`Kode Soal '${data.code}' sudah digunakan. Buat kode lain.`, 'error');
        return;
      }
      db.insert('questions', data);
      showToast('Soal baru berhasil ditambahkan.', 'success');
    }

    window.location.hash = '#soal';
  });
}

// ----------------------------------------------------
// 9. PAGE: IMPORT SOAL (EXCEL)
// ----------------------------------------------------
function renderImportSoal(mount) {
  mount.innerHTML = `
    <div style="display:grid; grid-template-columns: 1.2fr 2fr; gap:28px;">
      <!-- Upload area -->
      <div style="display:flex; flex-direction:column; gap:20px;">
        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Unggah Berkas Excel / CSV</h3>
          </div>
          <div class="card-body" style="display:flex; flex-direction:column; gap:20px;">
            <p style="font-size:13px; color:var(--neutral-600); line-height:1.5;">
              Unduh template Excel standar di bawah ini, isi data pertanyaan sesuai format, kemudian unggah kembali untuk diproses ke dalam database virtual.
            </p>
            
            <button class="btn btn-secondary" style="width:100%;" id="btn-download-tmpl">
              <i data-lucide="download"></i>
              <span>Unduh Template Excel</span>
            </button>
            
            <div class="upload-dropzone" id="excel-dropzone">
              <i data-lucide="file-spreadsheet" class="upload-icon" style="width:48px; height:48px;"></i>
              <h4 style="font-size:14px; font-weight:700; margin-bottom:4px;">Klik untuk memilih file</h4>
              <p style="font-size:11px; color:var(--neutral-600);">Format file .xlsx, .xls atau .csv</p>
              <input type="file" id="excel-file-input" style="display:none;" accept=".xlsx, .xls, .csv">
            </div>
            
            <div id="upload-status-info" style="font-size:13px; text-align:center; display:none;"></div>
          </div>
        </div>

        <!-- Import details selector -->
        <div class="card" id="import-mapping-card" style="display:none;">
          <div class="card-header">
            <h3 class="card-title">Target Kurikulum Impor</h3>
          </div>
          <div class="card-body" style="display:flex; flex-direction:column; gap:16px;">
            <div class="form-group">
              <label class="form-label" for="imp-sub">Petakan ke Mata Pelajaran</label>
              <select id="imp-sub" class="form-input" required>
                <option value="">Pilih Mapel</option>
                ${db.get('subjects').map(s => `<option value="${s.id}">${s.name}</option>`).join('')}
              </select>
            </div>
            <div class="form-group">
              <label class="form-label" for="imp-cls">Petakan ke Tingkat Kelas</label>
              <select id="imp-cls" class="form-input" required>
                <option value="">Pilih Kelas</option>
                ${db.get('classes').map(c => `<option value="${c.id}">Kelas ${c.name}</option>`).join('')}
              </select>
            </div>
            
            <button class="btn btn-primary" id="btn-submit-import-data" style="width:100%;" disabled>
              <i data-lucide="check-square"></i>
              <span>Impor Soal Valid ke Bank Soal</span>
            </button>
          </div>
        </div>
      </div>

      <!-- Preview parsed questions table -->
      <div class="card" style="min-height:300px;">
        <div class="card-header">
          <h3 class="card-title">Preview Data Soal Yang Dibaca</h3>
          <span class="badge badge-neutral" id="import-stats-badge">0 baris terbaca</span>
        </div>
        <div class="card-body" style="padding:0; overflow-y:auto; max-height: 60vh;">
          <div id="import-error-warnings" style="display:none; padding:16px 20px; background:var(--danger-light); color:var(--danger-text); border-bottom:1px solid var(--neutral-400); font-size:12px;"></div>
          <div class="table-wrapper">
            <table class="table" id="import-preview-table" style="font-size:12px;">
              <thead>
                <tr>
                  <th>Baris</th>
                  <th>Kode</th>
                  <th>Tipe</th>
                  <th>Tingkat</th>
                  <th>Pertanyaan</th>
                </tr>
              </thead>
              <tbody id="import-preview-tbody">
                <tr>
                  <td colspan="5" style="text-align:center; padding:40px; color:var(--neutral-600);">Belum ada data file yang diunggah.</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  // Elements mapping
  const dropzone = document.getElementById('excel-dropzone');
  const fileInput = document.getElementById('excel-file-input');
  const downloadTmplBtn = document.getElementById('btn-download-tmpl');
  const statusInfo = document.getElementById('upload-status-info');
  const mappingCard = document.getElementById('import-mapping-card');
  const submitImportBtn = document.getElementById('btn-submit-import-data');
  const errorWarnings = document.getElementById('import-error-warnings');
  const previewTbody = document.getElementById('import-preview-tbody');
  const statsBadge = document.getElementById('import-stats-badge');

  let importedQuestionsList = [];

  // Template download click
  downloadTmplBtn.addEventListener('click', downloadExcelTemplate);

  // File selection triggers
  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      handleImportFile(file);
    }
  });

  const handleImportFile = (file) => {
    statusInfo.style.display = 'block';
    statusInfo.innerHTML = `<span style="color:var(--neutral-700);">Membaca file "${file.name}"...</span>`;

    parseExcelOrCSV(file)
      .then(res => {
        const { questions, errors } = res;
        importedQuestionsList = questions;

        statusInfo.innerHTML = `<span style="color:var(--success-text); font-weight:700;">Membaca selesai!</span>`;
        statsBadge.innerText = `${questions.length} Soal Valid Terbaca`;

        // Render warnings/errors
        if (errors.length > 0) {
          errorWarnings.style.display = 'block';
          errorWarnings.innerHTML = `
            <strong>Ditemukan kesalahan/peringatan pada file:</strong>
            <ul style="margin-top:6px; padding-left:16px;">
              ${errors.map(e => `<li>Baris ${e.row} [Kode: ${e.code}]: ${e.messages.join(', ')}</li>`).join('')}
            </ul>
          `;
          showToast('Terdapat kesalahan pengisian file Excel.', 'error');
        } else {
          errorWarnings.style.display = 'none';
        }

        // Render preview table body
        if (questions.length > 0) {
          previewTbody.innerHTML = questions.map((q, i) => `
            <tr>
              <td>${i + 1}</td>
              <td style="font-weight:700; color:var(--primary);">${q.code}</td>
              <td><span class="badge badge-primary">${q.type}</span></td>
              <td><span class="badge badge-neutral">${q.difficulty}</span></td>
              <td style="max-width:200px; text-overflow:ellipsis; white-space:nowrap; overflow:hidden;">${q.questionText}</td>
            </tr>
          `).join('');

          mappingCard.style.display = 'block';
          submitImportBtn.removeAttribute('disabled');
        } else {
          previewTbody.innerHTML = '<tr><td colspan="5" style="text-align:center; padding:40px;">Tidak ada baris soal yang valid terdeteksi.</td></tr>';
          submitImportBtn.setAttribute('disabled', 'true');
        }
      })
      .catch(err => {
        statusInfo.innerHTML = `<span style="color:var(--danger-text); font-weight:700;">Gagal membaca!</span>`;
        previewTbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:40px; color:var(--danger-text);">${err.message}</td></tr>`;
        showToast(err.message, 'error');
      });
  };

  // Bind change inputs for mapping card
  const validateMapping = () => {
    const subVal = document.getElementById('imp-sub').value;
    const clsVal = document.getElementById('imp-cls').value;
    if (subVal && clsVal && importedQuestionsList.length > 0) {
      submitImportBtn.removeAttribute('disabled');
    } else {
      submitImportBtn.setAttribute('disabled', 'true');
    }
  };

  document.getElementById('imp-sub').addEventListener('change', validateMapping);
  document.getElementById('imp-cls').addEventListener('change', validateMapping);

  // Submitting imported questions data to LocalStorage
  submitImportBtn.addEventListener('click', () => {
    const sId = document.getElementById('imp-sub').value;
    const cId = document.getElementById('imp-cls').value;

    if (!sId || !cId) {
      showToast('Pilih Mata Pelajaran dan Tingkat Kelas terlebih dahulu.', 'error');
      return;
    }

    // Insert each question
    importedQuestionsList.forEach(q => {
      // Find matching default Bab/Topic or set empty
      const defaultBab = db.get('chapters').find(ch => ch.subjectId === sId && ch.classId === cId);
      const defaultTopic = defaultBab ? db.get('topics').find(tp => tp.chapterId === defaultBab.id) : null;

      q.subjectId = sId;
      q.classId = cId;
      q.chapterId = defaultBab ? defaultBab.id : '';
      q.topicId = defaultTopic ? defaultTopic.id : '';
      q.schoolId = 'sch-1';
      q.creatorId = activeUser.id;
      q.creatorName = activeUser.name;
      q.createdAt = new Date().toISOString();

      db.insert('questions', q);
    });

    showToast(`Berhasil mengimpor ${importedQuestionsList.length} soal ke bank soal.`, 'success');
    window.location.hash = '#soal';
  });
}

// ----------------------------------------------------
// 10. PAGE: PAKET SOAL
// ----------------------------------------------------
function renderPaketSoal(mount) {
  const packages = db.get('packages');
  const subjects = db.get('subjects');
  const classes = db.get('classes');

  mount.innerHTML = `
    <div class="action-row">
      <div></div>
      <button class="btn btn-primary" id="btn-tambah-paket">
        <i data-lucide="plus"></i>
        <span>Penyusunan Paket Baru</span>
      </button>
    </div>

    <div class="card">
      <div class="card-body" style="padding:0;">
        <div class="table-wrapper">
          <table class="table">
            <thead>
              <tr>
                <th>Nama Paket Soal</th>
                <th>Jenis</th>
                <th>Mata Pelajaran</th>
                <th>Kelas</th>
                <th>Jumlah Soal</th>
                <th style="text-align:right;">Aksi Berkas</th>
              </tr>
            </thead>
            <tbody>
              ${packages.map(p => {
    const sub = subjects.find(s => s.id === p.subjectId)?.name || '-';
    const cls = classes.find(c => c.id === p.classId)?.name || '-';
    return `
                  <tr>
                    <td>
                      <div style="font-weight:700; color:var(--neutral-900);">${p.name}</div>
                      <div style="font-size:11px; color:var(--neutral-600); margin-top:2px;">Disusun oleh: ${p.creatorName || 'Guru'}</div>
                    </td>
                    <td><span class="badge badge-primary">${p.type}</span></td>
                    <td>${sub}</td>
                    <td>Kelas ${cls}</td>
                    <td><strong style="color:var(--primary); font-size:15px;">${(p.questionIds || []).length}</strong> butir</td>
                    <td style="text-align:right;">
                      <div style="display:inline-flex; gap:8px;">
                        <button class="btn btn-secondary btn-icon btn-preview-pkg" data-id="${p.id}" title="Pratinjau Soal">
                          <i data-lucide="eye" style="width:16px; height:16px;"></i>
                        </button>
                        <button class="btn btn-secondary btn-icon btn-export-doc" data-id="${p.id}" title="Ekspor ke Word (MS Word)">
                          <i data-lucide="file-text" style="width:16px; height:16px; color:#2b6cb0;"></i>
                        </button>
                        <button class="btn btn-secondary btn-icon btn-export-xls" data-id="${p.id}" title="Ekspor ke Excel">
                          <i data-lucide="file-spreadsheet" style="width:16px; height:16px; color:#2f855a;"></i>
                        </button>
                        <button class="btn btn-secondary btn-icon btn-export-pdf" data-id="${p.id}" title="Cetak / Simpan PDF">
                          <i data-lucide="printer" style="width:16px; height:16px; color:#c53030;"></i>
                        </button>
                        <button class="btn btn-danger btn-icon btn-delete-pkg" data-id="${p.id}" title="Hapus Paket">
                          <i data-lucide="trash-2" style="width:16px; height:16px;"></i>
                        </button>
                      </div>
                    </td>
                  </tr>
                `;
  }).join('') || '<tr><td colspan="6" style="text-align:center;">Belum ada paket soal dibuat.</td></tr>'}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  // Add click listener
  document.getElementById('btn-tambah-paket').addEventListener('click', openCreatePackageModal);

  // File action handlers
  mount.querySelectorAll('.btn-preview-pkg').forEach(btn => {
    btn.addEventListener('click', () => previewPackageQuestions(btn.dataset.id));
  });

  mount.querySelectorAll('.btn-export-doc').forEach(btn => {
    btn.addEventListener('click', () => triggerExport(btn.dataset.id, 'doc'));
  });

  mount.querySelectorAll('.btn-export-xls').forEach(btn => {
    btn.addEventListener('click', () => triggerExport(btn.dataset.id, 'xls'));
  });

  mount.querySelectorAll('.btn-export-pdf').forEach(btn => {
    btn.addEventListener('click', () => triggerExport(btn.dataset.id, 'pdf'));
  });

  mount.querySelectorAll('.btn-delete-pkg').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.id;
      const p = db.get('packages').find(pkg => pkg.id === id);
      if (confirm(`Apakah Anda yakin ingin menghapus paket "${p.name}"?`)) {
        db.delete('packages', id);
        showToast('Paket soal berhasil dihapus.', 'success');
        renderPaketSoal(mount);
      }
    });
  });
}

function triggerExport(pkgId, type) {
  const pkg = db.get('packages').find(p => p.id === pkgId);
  const allQ = db.get('questions');
  const pkgQ = (pkg.questionIds || []).map(id => allQ.find(q => q.id === id)).filter(Boolean);
  const school = db.get('schools')[0] || { name: 'BankSoalPro', npsn: '-' };

  // Set maps details to package
  const sub = db.get('subjects').find(s => s.id === pkg.subjectId);
  const cls = db.get('classes').find(c => c.id === pkg.classId);
  pkg.subjectName = sub ? sub.name : '';
  pkg.className = cls ? cls.name : '';

  if (pkgQ.length === 0) {
    showToast('Paket soal kosong. Silakan isi paket terlebih dahulu.', 'error');
    return;
  }

  if (type === 'xls') {
    exportToExcel(pkg, pkgQ);
    showToast('Mengekspor berkas Excel...', 'success');
  } else {
    // Word / PDF export options dialog
    const html = `
      <div class="modal-header">
        <h3 class="modal-title">Opsi Ekspor Dokumen</h3>
        <button class="modal-close" data-close-modal><i data-lucide="x"></i></button>
      </div>
      <div class="modal-body" style="display:flex; flex-direction:column; gap:16px;">
        <p style="font-size:13px; color:var(--neutral-600);">Tentukan detail lembar ujian yang ingin Anda unduh.</p>
        <label style="display:flex; align-items:center; gap:10px; cursor:pointer; font-size:14px;">
          <input type="checkbox" id="chk-inc-ans" style="width:16px; height:16px;">
          <span>Sertakan Lembar Kunci Jawaban</span>
        </label>
        <label style="display:flex; align-items:center; gap:10px; cursor:pointer; font-size:14px;">
          <input type="checkbox" id="chk-inc-disc" style="width:16px; height:16px;">
          <span>Sertakan Pembahasan Soal</span>
        </label>
      </div>
      <div class="modal-footer">
        <button class="btn btn-secondary" data-close-modal>Batal</button>
        <button class="btn btn-primary" id="btn-confirm-export">Mulai Unduh</button>
      </div>
    `;

    openModal(html);

    document.getElementById('btn-confirm-export').addEventListener('click', () => {
      const includeAnswer = document.getElementById('chk-inc-ans').checked;
      const includeDiscussion = document.getElementById('chk-inc-disc').checked;

      closeModal();

      if (type === 'doc') {
        exportToWord(pkg, pkgQ, school, { includeAnswer, includeDiscussion });
        showToast('Mengekspor berkas Word...', 'success');
      } else if (type === 'pdf') {
        exportToPDF(pkg, pkgQ, school, { includeAnswer, includeDiscussion });
        showToast('Membuka jendela cetak PDF...', 'success');
      }
    });
  }
}

// Preview Package Questions list
function previewPackageQuestions(pkgId) {
  const pkg = db.get('packages').find(p => p.id === pkgId);
  const allQ = db.get('questions');
  const pkgQ = (pkg.questionIds || []).map(id => allQ.find(q => q.id === id)).filter(Boolean);

  const html = `
    <div class="modal-header">
      <h3 class="modal-title">Pratinjau Soal - ${pkg.name}</h3>
      <button class="modal-close" data-close-modal><i data-lucide="x"></i></button>
    </div>
    <div class="modal-body" style="display:flex; flex-direction:column; gap:20px; max-height:60vh; overflow-y:auto;" id="preview-pkg-list-mount">
      ${pkgQ.map((q, idx) => {
    let optionsStr = '';
    if (q.choices && (q.type === 'PG' || q.type === 'PG_KOMPLEKS')) {
      optionsStr = `
            <div class="choices-preview-list" style="margin-top:8px;">
              ${Object.entries(q.choices).map(([k, v]) => `<div><strong>${k}.</strong> ${v}</div>`).join('')}
            </div>
          `;
    }
    return `
          <div style="border-bottom:1px solid var(--neutral-400); padding-bottom:16px;">
            <div style="font-weight:700; color:var(--primary); font-size:12px; margin-bottom:6px;">No ${idx + 1} | ${q.code} [${q.type}]</div>
            <div>${q.questionText}</div>
            ${optionsStr}
          </div>
        `;
  }).join('') || '<p style="text-align:center;">Paket ini tidak memiliki soal.</p>'}
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" data-close-modal>Tutup</button>
    </div>
  `;

  openModal(html, 'modal-large');

  // Trigger math typesetting
  setTimeout(() => {
    const mount = document.getElementById('preview-pkg-list-mount');
    if (window.MathJax && window.MathJax.typesetPromise && mount) {
      window.MathJax.typesetPromise([mount]).catch(err => console.error(err));
    }
  }, 100);
}

// Modal Form Creator for Packages (Manual & Randomizer Generator)
function openCreatePackageModal() {
  const subjects = db.get('subjects');
  const classes = db.get('classes');

  const html = `
    <div class="modal-header">
      <h3 class="modal-title">Penyusunan Paket Soal</h3>
      <button class="modal-close" data-close-modal><i data-lucide="x"></i></button>
    </div>
    <form id="pkg-create-form">
      <div class="modal-body" style="display:flex; flex-direction:column; gap:16px; max-height:65vh; overflow-y:auto;">
        <div class="form-group">
          <label class="form-label" for="p-name">Nama Paket Ujian</label>
          <input type="text" id="p-name" class="form-input" placeholder="contoh: UTS Ganjil Matematika Kelas X" required>
        </div>
        
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
          <div class="form-group">
            <label class="form-label" for="p-type">Jenis Ujian</label>
            <select id="p-type" class="form-input" required>
              <option value="UTS">UTS (Ulangan Tengah Semester)</option>
              <option value="UAS">UAS (Ulangan Akhir Semester)</option>
              <option value="PAS">PAS (Penilaian Akhir Semester)</option>
              <option value="PAT">PAT (Penilaian Akhir Tahun)</option>
              <option value="TRY_OUT">Try Out</option>
              <option value="LATIHAN">Latihan Mandiri</option>
            </select>
          </div>
          
          <div class="form-group">
            <label class="form-label" for="p-mode">Metode Pemilihan Soal</label>
            <select id="p-mode" class="form-input" required>
              <option value="MANUAL">Pilih Manual dari Bank Soal</option>
              <option value="ACAK">Acak Otomatis (Generator)</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
          <div class="form-group">
            <label class="form-label" for="p-sub">Mata Pelajaran</label>
            <select id="p-sub" class="form-input" required>
              <option value="">Pilih Mapel</option>
              ${subjects.map(s => `<option value="${s.id}">${s.name}</option>`).join('')}
            </select>
          </div>
          <div class="form-group">
            <label class="form-label" for="p-cls">Tingkat Kelas</label>
            <select id="p-cls" class="form-input" required>
              <option value="">Pilih Kelas</option>
              ${classes.map(c => `<option value="${c.id}">Kelas ${c.name}</option>`).join('')}
            </select>
          </div>
        </div>

        <!-- Section 1: Manual List Select (Visible in MANUAL mode) -->
        <div id="pkg-manual-selection-block" style="border-top:1px solid var(--neutral-400); padding-top:16px;">
          <h4 style="font-size:13px; font-weight:700; margin-bottom:8px;">Daftar Soal Terverifikasi (APPROVED)</h4>
          <div style="padding:10px; font-size:12px; background:var(--neutral-100); color:var(--neutral-600); margin-bottom:8px; border-radius:4px;">
            Pilih Mata Pelajaran & Kelas di atas terlebih dahulu untuk menampilkan soal.
          </div>
          <div style="display:flex; flex-direction:column; gap:8px; max-height:200px; overflow-y:auto;" id="manual-questions-checklist-mount"></div>
        </div>

        <!-- Section 2: Randomizer Input (Visible in ACAK mode) -->
        <div id="pkg-random-selection-block" style="border-top:1px solid var(--neutral-400); padding-top:16px; display:none;">
          <h4 style="font-size:13px; font-weight:700; margin-bottom:12px;">Kriteria Pengacakan Soal</h4>
          
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
            <div class="form-group">
              <label class="form-label" for="p-num-q">Jumlah Butir Soal</label>
              <input type="number" min="1" max="100" id="p-num-q" class="form-input" value="10">
            </div>
            
            <div class="form-group">
              <label class="form-label" for="p-random-diff">Proporsi Kesulitan</label>
              <select id="p-random-diff" class="form-input">
                <option value="SEMUA">Bebas (Campuran Acak)</option>
                <option value="MUDAH">Khusus MUDAH</option>
                <option value="SEDANG">Khusus SEDANG</option>
                <option value="SULIT">Khusus SULIT</option>
              </select>
            </div>
          </div>
        <!-- Audio & Animation Settings -->
        <div style="border-top:1px dashed var(--neutral-400); padding-top:16px; margin-top:8px;">
          <h4 style="font-size:13px; font-weight:700; margin-bottom:12px;">Fitur Interaktif CBT Ujian</h4>
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
            <label style="display:flex; align-items:center; gap:8px; font-size:13px; font-weight:600; cursor:pointer;">
              <input type="checkbox" id="p-enable-audio" style="width:16px; height:16px;">
              <span>Aktifkan Musik Latar (quizmusic.mp3)</span>
            </label>
            <label style="display:flex; align-items:center; gap:8px; font-size:13px; font-weight:600; cursor:pointer;">
              <input type="checkbox" id="p-enable-anim" style="width:16px; height:16px;">
              <span>Aktifkan Animasi Latar Belakang (Partikel)</span>
            </label>
          </div>
        </div>

      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
        <button type="submit" class="btn btn-primary" id="btn-save-pkg">Simpan Paket Soal</button>
      </div>
    </form>
  `;

  openModal(html, 'modal-large');

  // Interactive blocks display
  const modeSelect = document.getElementById('p-mode');
  const subSelect = document.getElementById('p-sub');
  const clsSelect = document.getElementById('p-cls');
  const manualBlock = document.getElementById('pkg-manual-selection-block');
  const randomBlock = document.getElementById('pkg-random-selection-block');
  const checkMount = document.getElementById('manual-questions-checklist-mount');

  modeSelect.addEventListener('change', () => {
    if (modeSelect.value === 'MANUAL') {
      manualBlock.style.display = 'block';
      randomBlock.style.display = 'none';
    } else {
      manualBlock.style.display = 'none';
      randomBlock.style.display = 'block';
    }
  });

  const reloadChecklist = () => {
    const sId = subSelect.value;
    const cId = clsSelect.value;

    if (!sId || !cId) {
      checkMount.innerHTML = '<p style="font-size:12px; color:var(--neutral-600); text-align:center;">Pilih Mapel & Kelas dahulu.</p>';
      return;
    }

    // Only APPROVED questions can be added to package
    const list = db.get('questions').filter(q => q.subjectId === sId && q.classId === cId && q.status === 'APPROVED');
    checkMount.innerHTML = list.map(q => `
      <label style="display:flex; align-items:flex-start; gap:10px; padding:8px; border:1px solid var(--neutral-400); border-radius:4px; font-size:12px; cursor:pointer;">
        <input type="checkbox" class="pkg-q-check" value="${q.id}" style="margin-top:2px;">
        <div style="flex:1;">
          <div style="font-weight:700; color:var(--primary);">${q.code} | <span style="color:var(--neutral-700);">${q.type} (${q.difficulty})</span></div>
          <div style="margin-top:4px; text-overflow:ellipsis; overflow:hidden; white-space:nowrap; max-width:600px;">${q.questionText}</div>
        </div>
      </label>
    `).join('') || '<p style="font-size:12px; color:var(--danger-text); text-align:center; padding:10px;">Tidak ditemukan Soal berstatus APPROVED untuk kriteria ini.</p>';
  };

  subSelect.addEventListener('change', reloadChecklist);
  clsSelect.addEventListener('change', reloadChecklist);

  // Form submit handler
  document.getElementById('pkg-create-form').addEventListener('submit', (e) => {
    e.preventDefault();

    const name = document.getElementById('p-name').value;
    const type = document.getElementById('p-type').value;
    const mode = modeSelect.value;
    const subjectId = subSelect.value;
    const classId = clsSelect.value;

    let questionIds = [];

    if (mode === 'MANUAL') {
      questionIds = Array.from(document.querySelectorAll('.pkg-q-check:checked')).map(chk => chk.value);
      if (questionIds.length === 0) {
        showToast('Pilih minimal 1 butir soal untuk menyusun paket.', 'error');
        return;
      }
    } else {
      // Automatic randomizer logic
      const numQ = parseInt(document.getElementById('p-num-q').value) || 10;
      const diffProp = document.getElementById('p-random-diff').value;

      let sourceList = db.get('questions').filter(q => q.subjectId === subjectId && q.classId === classId && q.status === 'APPROVED');

      if (diffProp !== 'SEMUA') {
        sourceList = sourceList.filter(q => q.difficulty === diffProp);
      }

      if (sourceList.length === 0) {
        showToast('Tidak ada soal APPROVED yang cocok untuk pengacakan otomatis.', 'error');
        return;
      }

      // Shuffle array
      const shuffled = [...sourceList].sort(() => 0.5 - Math.random());
      questionIds = shuffled.slice(0, numQ).map(q => q.id);

      showToast(`Generator mengacak ${questionIds.length} soal yang cocok.`, 'success');
    }

    const enableAudio = document.getElementById('p-enable-audio').checked;
    const enableAnim = document.getElementById('p-enable-anim').checked;

    db.insert('packages', {
      name,
      type,
      schoolId: 'sch-1',
      subjectId,
      classId,
      creatorId: activeUser.id,
      creatorName: activeUser.name,
      questionIds,
      enableAudio,
      enableAnim,
      createdAt: new Date().toISOString()
    });

    showToast('Paket soal berhasil dibuat.', 'success');
    closeModal();
    renderPaketSoal(document.getElementById('shell-content-mount'));
  });
}

// ----------------------------------------------------
// 11. PAGE: AUDIT LOGS & BACKUP/RESTORE DATABASE
// ----------------------------------------------------
function renderLogs(mount) {
  const logs = db.get('logs');

  mount.innerHTML = `
    <div style="display:grid; grid-template-columns: 2fr 1fr; gap:28px;">
      <!-- Logs table card -->
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Riwayat Audit Log Pengguna</h3>
        </div>
        <div class="card-body" style="padding:0; max-height:70vh; overflow-y:auto;">
          <div class="table-wrapper">
            <table class="table" style="font-size:12px;">
              <thead>
                <tr>
                  <th>Waktu</th>
                  <th>Nama Pengguna</th>
                  <th>Aksi</th>
                  <th>Detail Aksi</th>
                </tr>
              </thead>
              <tbody>
                ${logs.map(l => {
    const date = new Date(l.timestamp).toLocaleDateString('id-ID');
    const time = new Date(l.timestamp).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    return `
                    <tr>
                      <td style="white-space:nowrap; color:var(--neutral-600);">${date} - ${time}</td>
                      <td style="font-weight:700;">${l.userName}</td>
                      <td><span class="badge badge-neutral">${l.action}</span></td>
                      <td style="color:var(--neutral-700);">${l.details}</td>
                    </tr>
                  `;
  }).join('') || '<tr><td colspan="4" style="text-align:center;">Belum ada riwayat audit log.</td></tr>'}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Backup Restore action card -->
      <div style="display:flex; flex-direction:column; gap:20px;">
        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Backup Database JSON</h3>
          </div>
          <div class="card-body" style="display:flex; flex-direction:column; gap:16px;">
            <p style="font-size:13px; color:var(--neutral-600); line-height:1.5;">
              Unduh salinan cadangan seluruh isi database virtual (pengguna, sekolah, soal, kurikulum, dan log) dalam satu file berkas JSON.
            </p>
            <button class="btn btn-primary" id="btn-backup-db" style="width:100%;">
              <i data-lucide="shield-check"></i>
              <span>Unduh Backup JSON</span>
            </button>
          </div>
        </div>
        
        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Restore Database</h3>
          </div>
          <div class="card-body" style="display:flex; flex-direction:column; gap:16px;">
            <p style="font-size:13px; color:var(--neutral-600); line-height:1.5;">
              Pilih file cadangan JSON hasil backup untuk memulihkan seluruh basis data sistem.
            </p>
            <div class="upload-dropzone" id="restore-dropzone" style="padding:20px;">
              <i data-lucide="upload-cloud" class="upload-icon" style="width:36px; height:36px;"></i>
              <h4 style="font-size:13px; font-weight:700;">Pilih file backup .json</h4>
              <input type="file" id="restore-file-input" style="display:none;" accept=".json">
            </div>
            
            <button class="btn btn-danger" id="btn-reset-db" style="width:100%; padding:8px 12px; font-size:12px;">
              <i data-lucide="refresh-cw"></i>
              <span>Reset Database ke Setelan Pabrik</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  const backupBtn = document.getElementById('btn-backup-db');
  const restoreZone = document.getElementById('restore-dropzone');
  const restoreInput = document.getElementById('restore-file-input');
  const resetBtn = document.getElementById('btn-reset-db');

  // 1. Trigger Backup JSON Download
  backupBtn.addEventListener('click', () => {
    const data = db.getAll();
    const jsonStr = JSON.stringify(data, null, 2);

    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `backup_banksoalpro_${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    db.log(activeUser.id, activeUser.name, 'BACKUP', 'Melakukan download backup database JSON.');
    showToast('File backup berhasil diunduh.', 'success');
  });

  // 2. Trigger Restore JSON Upload
  restoreZone.addEventListener('click', () => restoreInput.click());
  restoreInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = function (evt) {
        try {
          const parsedData = JSON.parse(evt.target.result);

          // Basic verification
          if (parsedData.users && parsedData.questions && parsedData.schools) {
            db.saveAll(parsedData);
            showToast('Basis data dipulihkan! Hubungkan kembali sesi login.', 'success');

            // Clear current user and force login
            sessionStorage.removeItem('active_user');
            window.location.hash = '';
            handleRouting();
          } else {
            showToast('Format berkas JSON backup tidak valid.', 'error');
          }
        } catch (err) {
          showToast('Gagal memproses JSON: Format berkas rusak.', 'error');
        }
      };
      reader.readAsText(file);
    }
  });

  // 3. Trigger Reset
  resetBtn.addEventListener('click', () => {
    if (confirm('PERINGATAN: Semua data guru, kurikulum, dan bank soal yang Anda tambahkan akan dihapus secara permanen dan disetel kembali ke contoh bawaan. Lanjutkan?')) {
      db.reset();
      showToast('Database disetel ulang. Memuat kembali...', 'info');
      setTimeout(() => {
        window.location.hash = '';
        handleRouting();
      }, 1000);
    }
  });
}

// ----------------------------------------------------
// 12. SISWA: DASHBOARD KHUSUS SISWA
// ----------------------------------------------------
function renderStudentDashboard(mount) {
  const packages = db.get('packages');
  const subjects = db.get('subjects');
  const classes = db.get('classes');
  const results = db.get('results').filter(r => r.userId === activeUser.id);

  // Calculate average score
  const avgScore = results.length > 0
    ? Math.round(results.reduce((acc, r) => acc + r.score, 0) / results.length)
    : 0;

  mount.innerHTML = `
    <!-- Student Header Welcome Card -->
    <div style="background: linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%); color: white; padding: 32px; border-radius: var(--radius-lg); margin-bottom: 32px; box-shadow: var(--shadow-md); display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 20px;">
      <div>
        <h2 style="font-size: 26px; font-weight: 800; margin-bottom: 8px;">Selamat Belajar, ${activeUser.name}!</h2>
        <p style="font-size: 14px; opacity: 0.9; max-width: 500px;">
          Gunakan simulator ini untuk melatih pemahaman Anda terhadap bank soal sekolah. Klik "Mulai Ujian" untuk memulai pengerjaan.
        </p>
      </div>
      <div style="display:flex; gap:16px;">
        <div style="background: rgba(255,255,255,0.15); backdrop-filter: blur(8px); padding: 16px 24px; border-radius: var(--radius-md); text-align: center; border: 1px solid rgba(255,255,255,0.2);">
          <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; opacity: 0.8; letter-spacing: 0.5px;">Ujian Selesai</div>
          <div style="font-size: 24px; font-weight: 800; margin-top: 4px;">${results.length}</div>
        </div>
        <div style="background: rgba(255,255,255,0.15); backdrop-filter: blur(8px); padding: 16px 24px; border-radius: var(--radius-md); text-align: center; border: 1px solid rgba(255,255,255,0.2);">
          <div style="font-size: 11px; font-weight: 700; text-transform: uppercase; opacity: 0.8; letter-spacing: 0.5px;">Rata-Rata Nilai</div>
          <div style="font-size: 24px; font-weight: 800; margin-top: 4px; color: ${avgScore >= 75 ? '#68d391' : '#f6ad55'};">${avgScore}</div>
        </div>
      </div>
    </div>

    <div class="student-dashboard-grid">
      <!-- Active Packages List -->
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Daftar Paket Ujian Yang Tersedia</h3>
          <i data-lucide="award" style="color:var(--primary);"></i>
        </div>
        <div class="card-body" style="padding:0;">
          <div class="table-wrapper">
            <table class="table">
              <thead>
                <tr>
                  <th>Nama Paket</th>
                  <th>Mata Pelajaran</th>
                  <th>Jumlah Soal</th>
                  <th style="text-align:right;">Aksi</th>
                </tr>
              </thead>
              <tbody>
                ${packages.map(p => {
    const sub = subjects.find(s => s.id === p.subjectId)?.name || '-';
    const cls = classes.find(c => c.id === p.classId)?.name || '-';
    const qCount = (p.questionIds || []).length;

    return `
                    <tr>
                      <td>
                        <div style="font-weight: 700; color: var(--neutral-900);">${p.name}</div>
                        <div style="font-size: 11px; color: var(--neutral-600); margin-top:2px;">Kelas ${cls} | Penguji: ${p.creatorName}</div>
                      </td>
                      <td><span class="badge badge-primary">${sub}</span></td>
                      <td><strong>${qCount}</strong> butir</td>
                      <td style="text-align:right;">
                        <a href="#cbt?id=${p.id}" class="btn btn-primary" style="padding: 8px 16px; font-size: 13px;">
                          <i data-lucide="play-circle" style="width:16px; height:16px;"></i>
                          <span>Mulai Ujian</span>
                        </a>
                      </td>
                    </tr>
                  `;
  }).join('') || '<tr><td colspan="4" style="text-align:center; padding:40px;">Belum ada paket ujian yang dirilis oleh guru.</td></tr>'}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Practice History -->
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Riwayat Latihan Ujian</h3>
          <i data-lucide="history" style="color:var(--neutral-600);"></i>
        </div>
        <div class="card-body" style="padding:0; max-height: 400px; overflow-y:auto;">
          <div class="table-wrapper">
            <table class="table" style="font-size:13px;">
              <thead>
                <tr>
                  <th>Nama Paket</th>
                  <th>Nilai</th>
                  <th style="text-align:right;">Aksi</th>
                </tr>
              </thead>
              <tbody>
                ${results.map(r => {
    const scoreColor = r.score >= 75 ? 'var(--success-text)' : 'var(--danger-text)';
    return `
                    <tr>
                      <td>
                        <div style="font-weight: 600; text-overflow:ellipsis; overflow:hidden; white-space:nowrap; max-width:180px;" title="${r.packageName}">${r.packageName}</div>
                        <div style="font-size:10px; color:var(--neutral-600); margin-top:2px;">${new Date(r.timestamp).toLocaleDateString('id-ID')}</div>
                      </td>
                      <td>
                        <span style="font-weight: 800; font-size: 15px; color: ${scoreColor};">${r.score}</span>
                      </td>
                      <td style="text-align:right;">
                        <a href="#cbt-result?id=${r.id}" class="btn btn-secondary" style="padding: 4px 8px; font-size: 11px;">
                          <span>Review</span>
                        </a>
                      </td>
                    </tr>
                  `;
  }).join('') || '<tr><td colspan="3" style="text-align:center; padding:30px; color:var(--neutral-600);">Belum ada riwayat pengerjaan.</td></tr>'}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();
}

// ----------------------------------------------------
// 12.A. REKAP & RIWAYAT HASIL UJIAN (SISWA & GURU)
// ----------------------------------------------------
function renderRekapUjian(mount) {
  const subjects = db.get('subjects');
  const classes = db.get('classes');
  
  if (activeUser.role === 'SISWA') {
    const results = db.get('results').filter(r => r.userId === activeUser.id);
    
    // Summary calculations
    const totalExams = results.length;
    const avgScore = totalExams > 0 ? Math.round(results.reduce((acc, r) => acc + r.score, 0) / totalExams) : 0;
    const highestScore = totalExams > 0 ? Math.max(...results.map(r => r.score)) : 0;
    const passCount = results.filter(r => r.score >= 75).length;
    const passRate = totalExams > 0 ? Math.round((passCount / totalExams) * 100) : 0;

    mount.innerHTML = `
      <!-- Summary metrics cards -->
      <div class="metrics-grid" style="margin-bottom: 24px;">
        <div class="metric-card">
          <div class="metric-info">
            <span class="metric-label">Ujian Diikuti</span>
            <span class="metric-value">${totalExams}</span>
          </div>
          <div class="metric-icon-box metric-blue"><i data-lucide="clipboard-list"></i></div>
        </div>
        <div class="metric-card">
          <div class="metric-info">
            <span class="metric-label">Rerata Nilai</span>
            <span class="metric-value" style="color: ${avgScore >= 75 ? 'var(--success-text)' : 'var(--warning-text)'};">${avgScore}</span>
          </div>
          <div class="metric-icon-box metric-purple"><i data-lucide="award"></i></div>
        </div>
        <div class="metric-card">
          <div class="metric-info">
            <span class="metric-label">Nilai Tertinggi</span>
            <span class="metric-value">${highestScore}</span>
          </div>
          <div class="metric-icon-box metric-indigo"><i data-lucide="trending-up"></i></div>
        </div>
        <div class="metric-card">
          <div class="metric-info">
            <span class="metric-label">Persentase Lulus KKM</span>
            <span class="metric-value">${passRate}%</span>
          </div>
          <div class="metric-icon-box metric-orange"><i data-lucide="percent"></i></div>
        </div>
      </div>

      <div class="charts-grid" style="margin-bottom: 24px;">
        <!-- Line Chart showing score progression -->
        <div class="card">
          <div class="card-header">
            <h3 class="card-title">Grafik Rekap Perkembangan Nilai</h3>
            <i data-lucide="line-chart" style="color: var(--neutral-600);"></i>
          </div>
          <div class="card-body">
            <div class="chart-container" style="position: relative; height:250px; width:100%;">
              <canvas id="chart-rekap-siswa"></canvas>
            </div>
          </div>
        </div>
      </div>

      <!-- History Table -->
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Riwayat Lengkap Ujian Mandiri</h3>
          <i data-lucide="history" style="color: var(--neutral-600);"></i>
        </div>
        <div class="card-body" style="padding:0;">
          <div class="table-wrapper">
            <table class="table">
              <thead>
                <tr>
                  <th>Nama Paket Soal</th>
                  <th>Mata Pelajaran</th>
                  <th>Tanggal Pengerjaan</th>
                  <th>Nilai Ujian</th>
                  <th>Status</th>
                  <th style="text-align:right;">Tinjauan</th>
                </tr>
              </thead>
              <tbody>
                ${results.map(r => {
                  const isPassed = r.score >= 75;
                  const scoreColor = isPassed ? 'var(--success-text)' : 'var(--danger-text)';
                  const badgeClass = isPassed ? 'badge-success' : 'badge-danger';
                  return `
                    <tr>
                      <td><strong style="color:var(--neutral-900);">${r.packageName}</strong></td>
                      <td><span class="badge badge-primary">${r.subjectName}</span></td>
                      <td>${new Date(r.timestamp).toLocaleString('id-ID')}</td>
                      <td><span style="font-weight: 800; font-size:16px; color:${scoreColor};">${r.score}</span></td>
                      <td><span class="badge ${badgeClass}">${isPassed ? 'TUNTAS' : 'REMEDIAL'}</span></td>
                      <td style="text-align:right;">
                        <a href="#cbt-result?id=${r.id}" class="btn btn-secondary" style="padding:6px 12px; font-size:12px; display:inline-flex; align-items:center; gap:4px;">
                          <i data-lucide="eye" style="width:12px; height:12px;"></i>
                          <span>Lihat Review</span>
                        </a>
                      </td>
                    </tr>
                  `;
                }).join('') || '<tr><td colspan="6" style="text-align:center; padding:30px;">Belum ada riwayat pengerjaan ujian.</td></tr>'}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;

    lucide.createIcons();

    // Render line chart
    if (totalExams > 0) {
      setTimeout(() => {
        const canvas = document.getElementById('chart-rekap-siswa');
        if (!canvas) return;
        
        const sorted = [...results].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
        const labels = sorted.map(r => new Date(r.timestamp).toLocaleDateString('id-ID'));
        const scores = sorted.map(r => r.score);

        new Chart(canvas, {
          type: 'line',
          data: {
            labels: labels,
            datasets: [{
              label: 'Nilai Ujian',
              data: scores,
              borderColor: 'rgb(79, 70, 229)',
              backgroundColor: 'rgba(79, 70, 229, 0.1)',
              borderWidth: 3,
              tension: 0.3,
              fill: true,
              pointBackgroundColor: 'rgb(79, 70, 229)',
              pointRadius: 6,
              pointHoverRadius: 8
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
              y: {
                min: 0,
                max: 100,
                ticks: {
                  stepSize: 20
                }
              }
            },
            plugins: {
              legend: {
                display: false
              }
            }
          }
        });
      }, 50);
    }
  } else {
    // For GURU, ADMIN_SEKOLAH, SUPER_ADMIN
    const allResults = db.get('results');
    
    const renderTable = (filtered) => {
      const tbody = document.getElementById('rekap-tbody-mount');
      if (!tbody) return;

      tbody.innerHTML = filtered.map(r => {
        const isPassed = r.score >= 75;
        const scoreColor = isPassed ? 'var(--success-text)' : 'var(--danger-text)';
        const badgeClass = isPassed ? 'badge-success' : 'badge-danger';
        return `
          <tr>
            <td>
              <div style="font-weight:700; color:var(--neutral-900);">${r.userName}</div>
              <div style="font-size:10px; color:var(--neutral-600); margin-top:2px;">Kelas ${r.className || '-'}</div>
            </td>
            <td><span class="badge badge-primary">${r.subjectName}</span></td>
            <td>
              <div style="font-weight:600; color:var(--neutral-800);">${r.packageName}</div>
              <div style="font-size:10px; color:var(--neutral-600); margin-top:2px;">${new Date(r.timestamp).toLocaleString('id-ID')}</div>
            </td>
            <td><span style="font-weight: 800; font-size:15px; color:${scoreColor};">${r.score}</span></td>
            <td><span class="badge ${badgeClass}">${isPassed ? 'TUNTAS' : 'REMEDIAL'}</span></td>
            <td style="text-align:right;">
              <a href="#cbt-result?id=${r.id}" class="btn btn-secondary" style="padding:6px 12px; font-size:12px; display:inline-flex; align-items:center; gap:4px;">
                <i data-lucide="eye" style="width:12px; height:12px;"></i>
                <span>Review</span>
              </a>
            </td>
          </tr>
        `;
      }).join('') || '<tr><td colspan="6" style="text-align:center; padding:30px;">Tidak ditemukan data rekapitulasi nilai.</td></tr>';
      
      lucide.createIcons();
    };

    const updateStats = (filtered) => {
      const total = filtered.length;
      const avg = total > 0 ? Math.round(filtered.reduce((acc, r) => acc + r.score, 0) / total) : 0;
      const highest = total > 0 ? Math.max(...filtered.map(r => r.score)) : 0;
      const passCount = filtered.filter(r => r.score >= 75).length;
      const passRate = total > 0 ? Math.round((passCount / total) * 100) : 0;

      document.getElementById('stat-total-exams').innerText = total;
      document.getElementById('stat-avg-score').innerText = avg;
      document.getElementById('stat-avg-score').style.color = avg >= 75 ? 'var(--success-text)' : 'var(--warning-text)';
      document.getElementById('stat-highest-score').innerText = highest;
      document.getElementById('stat-pass-rate').innerText = `${passRate}%`;
    };

    mount.innerHTML = `
      <!-- Stats Summary -->
      <div class="metrics-grid" style="margin-bottom: 24px;">
        <div class="metric-card">
          <div class="metric-info">
            <span class="metric-label">Ujian Dikerjakan</span>
            <span class="metric-value" id="stat-total-exams">0</span>
          </div>
          <div class="metric-icon-box metric-blue"><i data-lucide="clipboard-list"></i></div>
        </div>
        <div class="metric-card">
          <div class="metric-info">
            <span class="metric-label">Rerata Nilai</span>
            <span class="metric-value" id="stat-avg-score">0</span>
          </div>
          <div class="metric-icon-box metric-purple"><i data-lucide="award"></i></div>
        </div>
        <div class="metric-card">
          <div class="metric-info">
            <span class="metric-label">Nilai Tertinggi</span>
            <span class="metric-value" id="stat-highest-score">0</span>
          </div>
          <div class="metric-icon-box metric-indigo"><i data-lucide="trending-up"></i></div>
        </div>
        <div class="metric-card">
          <div class="metric-info">
            <span class="metric-label">Tingkat Ketuntasan</span>
            <span class="metric-value" id="stat-pass-rate">0%</span>
          </div>
          <div class="metric-icon-box metric-orange"><i data-lucide="check-circle"></i></div>
        </div>
      </div>

      <!-- Filters & Search block -->
      <div class="card" style="margin-bottom: 24px;">
        <div class="card-body" style="display:flex; gap:16px; flex-wrap:wrap; align-items:center;">
          <div style="flex:1; min-width:200px;">
            <label class="form-label" style="margin-bottom:6px;">Cari Siswa atau Paket</label>
            <input type="text" id="filter-search" class="form-input" placeholder="Ketik nama siswa atau nama paket...">
          </div>
          <div style="width:150px;">
            <label class="form-label" style="margin-bottom:6px;">Tingkat Kelas</label>
            <select id="filter-class" class="form-input">
              <option value="">Semua Kelas</option>
              ${classes.map(c => `<option value="${c.name}">Kelas ${c.name}</option>`).join('')}
            </select>
          </div>
          <div style="width:180px;">
            <label class="form-label" style="margin-bottom:6px;">Mata Pelajaran</label>
            <select id="filter-subject" class="form-input">
              <option value="">Semua Mapel</option>
              ${subjects.map(s => `<option value="${s.name}">${s.name}</option>`).join('')}
            </select>
          </div>
          <div style="align-self:flex-end;">
            <button class="btn btn-primary" id="btn-export-rekap-excel" style="background-color:var(--success); border-color:var(--success); font-weight:700; height:42px; display:flex; align-items:center; gap:8px;">
              <i data-lucide="file-spreadsheet"></i>
              <span>Ekspor Excel</span>
            </button>
          </div>
        </div>
      </div>

      <!-- Table results -->
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Rekapitulasi Nilai Seluruh Siswa</h3>
          <i data-lucide="users" style="color:var(--neutral-600);"></i>
        </div>
        <div class="card-body" style="padding:0;">
          <div class="table-wrapper">
            <table class="table">
              <thead>
                <tr>
                  <th>Nama Siswa</th>
                  <th>Mata Pelajaran</th>
                  <th>Paket Ujian & Waktu</th>
                  <th>Nilai</th>
                  <th>Status</th>
                  <th style="text-align:right;">Aksi</th>
                </tr>
              </thead>
              <tbody id="rekap-tbody-mount">
                <!-- Injected dynamically -->
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;

    lucide.createIcons();

    const searchInput = document.getElementById('filter-search');
    const classSelect = document.getElementById('filter-class');
    const subjectSelect = document.getElementById('filter-subject');
    const exportBtn = document.getElementById('btn-export-rekap-excel');

    const runFilters = () => {
      const searchVal = searchInput.value.toLowerCase();
      const classVal = classSelect.value;
      const subjectVal = subjectSelect.value;

      let filtered = allResults;

      // Filter by search
      if (searchVal) {
        filtered = filtered.filter(r => 
          r.userName.toLowerCase().includes(searchVal) || 
          r.packageName.toLowerCase().includes(searchVal)
        );
      }
      
      // Filter by Class
      if (classVal) {
        filtered = filtered.filter(r => r.className === classVal);
      }

      // Filter by Subject
      if (subjectVal) {
        filtered = filtered.filter(r => r.subjectName === subjectVal);
      }

      renderTable(filtered);
      updateStats(filtered);
      return filtered;
    };

    // Bind listeners
    searchInput.addEventListener('input', runFilters);
    classSelect.addEventListener('change', runFilters);
    subjectSelect.addEventListener('change', runFilters);

    // Initial render
    const initialList = runFilters();

    // Export to Excel handler
    exportBtn.addEventListener('click', (e) => {
      e.preventDefault();
      const currentFiltered = runFilters();
      
      if (currentFiltered.length === 0) {
        showToast('Tidak ada data untuk diekspor.', 'error');
        return;
      }

      const excelData = currentFiltered.map((r, idx) => ({
        'No': idx + 1,
        'Nama Siswa': r.userName,
        'Kelas': `Kelas ${r.className || ''}`,
        'Mata Pelajaran': r.subjectName,
        'Nama Paket Ujian': r.packageName,
        'Nilai Ujian': r.score,
        'Status Kelulusan': r.score >= 75 ? 'TUNTAS' : 'REMEDIAL',
        'Tanggal Pengerjaan': new Date(r.timestamp).toLocaleString('id-ID')
      }));

      const ws = XLSX.utils.json_to_sheet(excelData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Rekap Nilai Siswa");
      
      // Set Column widths
      const wscols = [
        {wch: 5},
        {wch: 25},
        {wch: 12},
        {wch: 20},
        {wch: 35},
        {wch: 12},
        {wch: 15},
        {wch: 25}
      ];
      ws['!cols'] = wscols;

      XLSX.writeFile(wb, `Rekap_Nilai_Ujian_${new Date().toISOString().slice(0,10)}.xlsx`);
      showToast('Berhasil mengunduh rekap hasil ujian Excel!', 'success');
    });
  }
}

// ----------------------------------------------------
// 13. SISWA: INTERACTIVE CBT EXAM SCREEN (FULLSCREEN MOCK)
// ----------------------------------------------------
let cbtInterval = null; // Global interval tracker to prevent leaks

function renderCBTScreen(packageId) {
  const pkg = db.get('packages').find(p => p.id === packageId);
  const subjects = db.get('subjects');
  const classes = db.get('classes');

  if (!pkg) {
    showToast('Paket ujian tidak ditemukan!', 'error');
    window.location.hash = '#dashboard';
    return;
  }

  const allQuestions = db.get('questions');
  const pkgQuestions = (pkg.questionIds || []).map(id => allQuestions.find(q => q.id === id)).filter(Boolean);

  if (pkgQuestions.length === 0) {
    showToast('Paket ujian ini tidak memiliki soal!', 'error');
    window.location.hash = '#dashboard';
    return;
  }

  // Clear any existing intervals
  if (cbtInterval) {
    clearInterval(cbtInterval);
  }

  // Local CBT state
  let currentIdx = 0;
  const answers = {}; // Maps questionId -> answer
  const doubtful = {}; // Maps questionId -> boolean (true if marked as doubtful)
  let timeLeft = pkgQuestions.length * 2 * 60; // 2 minutes per question in seconds

  // Initialize answers and doubtful map
  pkgQuestions.forEach(q => {
    doubtful[q.id] = false;
    if (q.type === 'PG_KOMPLEKS') {
      answers[q.id] = [];
    } else {
      answers[q.id] = '';
    }
  });

  const subjectName = subjects.find(s => s.id === pkg.subjectId)?.name || 'Umum';
  const className = classes.find(c => c.id === pkg.classId)?.name || 'X';

  // Build full-screen CBT workspace layout inside #app
  appMount.innerHTML = `
    <!-- CBT Isolation style -->
    <style>
      .cbt-app-shell {
        display: flex;
        flex-direction: column;
        height: 100vh;
        background-color: #f8fafc;
        font-family: 'Plus Jakarta Sans', sans-serif;
      }
      .cbt-header {
        height: 70px;
        background-color: white;
        border-bottom: 1px solid var(--neutral-400);
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 40px;
        box-shadow: var(--shadow-sm);
        z-index: 10;
      }
      .cbt-body {
        flex: 1;
        display: flex;
        overflow: hidden;
      }
      .cbt-content {
        flex: 1;
        padding: 40px;
        overflow-y: auto;
        display: flex;
        flex-direction: column;
        gap: 24px;
      }
      .cbt-sidebar {
        width: 320px;
        background-color: white;
        border-left: 1px solid var(--neutral-400);
        padding: 30px;
        display: flex;
        flex-direction: column;
        gap: 24px;
        overflow-y: auto;
      }
      .cbt-question-card {
        background: white;
        border: 1px solid var(--neutral-400);
        border-radius: var(--radius-lg);
        padding: 32px;
        box-shadow: var(--shadow-sm);
        min-height: 350px;
        display: flex;
        flex-direction: column;
        gap: 20px;
      }
      .cbt-options-list {
        display: flex;
        flex-direction: column;
        gap: 12px;
      }
      .cbt-option-item {
        display: flex;
        align-items: flex-start;
        gap: 12px;
        padding: 16px 20px;
        border: 1px solid var(--neutral-400);
        border-radius: var(--radius-md);
        cursor: pointer;
        transition: all var(--transition-fast);
        background-color: #f8fafc;
        font-size: 14px;
      }
      .cbt-option-item:hover {
        border-color: var(--primary);
        background-color: var(--primary-light);
      }
      .cbt-option-item.selected {
        border-color: var(--primary);
        background-color: var(--primary-light);
        font-weight: 600;
        box-shadow: 0 0 0 1px var(--primary);
      }
      .cbt-num-grid {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 8px;
      }
      .cbt-num-btn {
        aspect-ratio: 1;
        border-radius: var(--radius-sm);
        border: 1px solid var(--neutral-400);
        background-color: white;
        font-weight: 700;
        font-size: 13px;
        color: var(--neutral-700);
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: all var(--transition-fast);
      }
      .cbt-num-btn.active {
        border-color: var(--primary);
        background-color: var(--primary);
        color: white;
        box-shadow: var(--shadow-glow);
      }
      .cbt-num-btn.answered {
        background-color: #e6fffa;
        border-color: #319795;
        color: #234e52;
      }
      .cbt-num-btn.viewed {
        border-color: var(--warning);
        color: var(--warning-text);
      }
      .cbt-num-btn.doubtful {
        background-color: #feebc8 !important;
        border-color: #dd6b20 !important;
        color: #dd6b20 !important;
      }
      .cbt-num-btn.doubtful.active {
        background-color: #f6ad55 !important;
        border-color: var(--primary) !important;
        color: white !important;
      }
      .timer-display {
        font-size: 18px;
        font-weight: 800;
        font-family: monospace;
        color: var(--neutral-900);
        display: flex;
        align-items: center;
        gap: 8px;
        background-color: #fffaf0;
        border: 1px solid #feebc8;
        padding: 6px 16px;
        border-radius: var(--radius-sm);
      }
      .timer-danger {
        background-color: #fff5f5;
        border-color: #fed7d7;
        color: var(--danger-text);
        animation: pulse 1s infinite alternate;
      }
      @keyframes pulse {
        from { opacity: 1; }
        to { opacity: 0.6; }
      }
      
      /* CBT Mobile Responsive Styling */
      @media (max-width: 768px) {
        .cbt-header {
          padding: 0 16px;
          height: auto;
          min-height: 70px;
          flex-direction: column;
          align-items: stretch;
          justify-content: center;
          gap: 12px;
          padding-top: 12px;
          padding-bottom: 12px;
        }
        .cbt-header div[style*="display:flex; flex-direction:column;"] {
          text-align: center;
          align-items: center;
        }
        .timer-display {
          justify-content: center;
          font-size: 16px;
        }
        .cbt-body {
          flex-direction: column-reverse;
          overflow-y: auto;
        }
        .cbt-content {
          padding: 20px 16px;
          flex: none;
          overflow-y: visible;
        }
        .cbt-sidebar {
          width: 100%;
          border-left: none;
          border-top: 1px solid var(--neutral-400);
          padding: 24px 16px;
          flex: none;
          overflow-y: visible;
        }
        .cbt-question-card {
          padding: 20px 16px;
          min-height: auto;
        }
        .cbt-num-grid {
          grid-template-columns: repeat(auto-fill, minmax(42px, 1fr));
        }
        .cbt-option-item {
          padding: 12px 16px;
          font-size: 13px;
        }
      }
    </style>

    <div class="cbt-app-shell">
      <!-- Header -->
      <header class="cbt-header">
        <div style="display:flex; flex-direction:column;">
          <span style="font-size:12px; font-weight:700; color:var(--primary); text-transform:uppercase; letter-spacing:0.5px;">CBT SIMULATOR</span>
          <span style="font-size:16px; font-weight:800; color:var(--neutral-900); max-width:400px; text-overflow:ellipsis; white-space:nowrap; overflow:hidden;" title="${pkg.name}">${pkg.name}</span>
        </div>
        
        <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap; justify-content:flex-end;">
          ${pkg.enableAudio ? `
            <button class="btn btn-secondary" id="btn-mute-audio" style="width:38px; height:38px; padding:0; display:flex; align-items:center; justify-content:center; border-radius:50%; cursor:pointer;" title="Bisukan / Bunyikan Musik">
              <i data-lucide="volume-2" style="width:16px; height:16px;"></i>
            </button>
          ` : ''}

          <button class="btn btn-secondary" id="btn-toggle-sidebar" style="font-weight:600; display:flex; align-items:center; gap:8px; padding: 8px 14px; font-size: 13px;">
            <i data-lucide="layout-sidebar"></i>
            <span id="text-toggle-sidebar">Sembunyikan Nomor</span>
          </button>
          
          <div class="timer-display" id="cbt-timer-mount" style="padding: 6px 12px; font-size: 15px;">
            <i data-lucide="clock" style="width:16px; height:16px;"></i>
            <span id="cbt-timer-clock">00:00:00</span>
          </div>
          
          <button class="btn btn-primary" id="btn-submit-cbt" style="background-color:var(--success); border-color:var(--success); font-weight:700; padding: 8px 14px; font-size: 13px;">
            <i data-lucide="check-square"></i>
            <span>Selesai Ujian</span>
          </button>
        </div>
      </header>
      
      <!-- Body workspace -->
      <div class="cbt-body">
        <!-- Question Workspace -->
        <div class="cbt-content">
          <!-- Question text card -->
          <div class="cbt-question-card">
            <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid var(--neutral-400); padding-bottom:12px; font-size:13px; color:var(--neutral-600);">
              <span id="cbt-question-number-title" style="font-weight:700; color:var(--neutral-900);">Soal Nomor 1</span>
              <span>Mapel: ${subjectName} | Kelas: ${className}</span>
            </div>
            
            <div id="cbt-question-text-mount" style="font-size:16px; line-height:1.6; color:var(--neutral-900);"></div>
            
            <div id="cbt-options-mount" style="margin-top:10px;"></div>
          </div>
          
          <!-- Bottom Navigation Row -->
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
            <button class="btn btn-secondary" id="btn-cbt-prev" disabled>
              <i data-lucide="chevron-left"></i>
              <span>Soal Sebelumnya</span>
            </button>
            
            <button class="btn" id="btn-cbt-doubt" style="background-color:#feebc8; color:#dd6b20; border:1px solid #dd6b20; font-weight:700; display:flex; align-items:center; gap:8px;">
              <i data-lucide="help-circle" style="width:16px; height:16px;"></i>
              <span>Ragu-Ragu</span>
            </button>

            <button class="btn btn-primary" id="btn-cbt-next">
              <span>Soal Berikutnya</span>
              <i data-lucide="chevron-right"></i>
            </button>
          </div>
        </div>
        
        <!-- Sidebar Navigation Grid -->
        <aside class="cbt-sidebar">
          <div>
            <h3 style="font-size:14px; font-weight:700; margin-bottom:12px; color:var(--neutral-900);">Navigasi Soal</h3>
            <div class="cbt-num-grid" id="cbt-number-grid-mount"></div>
          </div>
          
          <div style="margin-top:auto; padding-top:20px; border-top:1px solid var(--neutral-400); font-size:12px; color:var(--neutral-600); display:flex; flex-direction:column; gap:8px;">
            <div style="display:flex; align-items:center; gap:8px;"><div style="width:16px; height:16px; background:#e6fffa; border:1px solid #319795; border-radius:3px;"></div><span>Sudah dijawab</span></div>
            <div style="display:flex; align-items:center; gap:8px;"><div style="width:16px; height:16px; background:#fff; border:1px solid var(--neutral-400); border-radius:3px;"></div><span>Belum dijawab</span></div>
            <div style="display:flex; align-items:center; gap:8px;"><div style="width:16px; height:16px; background:#feebc8; border:1px solid #dd6b20; border-radius:3px;"></div><span>Ragu-Ragu</span></div>
            <div style="display:flex; align-items:center; gap:8px;"><div style="width:16px; height:16px; background:var(--primary); border:1px solid var(--primary); border-radius:3px;"></div><span>Soal Aktif</span></div>
          </div>
        </aside>
      </div>
    </div>
  `;

  lucide.createIcons();

  // Background Music Logic
  if (pkg.enableAudio) {
    if (cbtAudio) {
      cbtAudio.pause();
    }
    cbtAudio = new Audio('quizmusic.mp3');
    cbtAudio.loop = true;
    cbtAudio.volume = 0.4;

    const startAudio = () => {
      if (cbtAudio) {
        cbtAudio.play().catch(err => console.log('Autoplay blocked. Will try again on click.'));
      }
      document.removeEventListener('click', startAudio);
      document.removeEventListener('keydown', startAudio);
    };

    document.addEventListener('click', startAudio);
    document.addEventListener('keydown', startAudio);

    // Try playing immediately
    cbtAudio.play().catch(() => {});

    // Mute event listener
    const muteBtn = document.getElementById('btn-mute-audio');
    if (muteBtn) {
      muteBtn.addEventListener('click', (e) => {
        e.preventDefault();
        if (cbtAudio) {
          cbtAudio.muted = !cbtAudio.muted;
          if (cbtAudio.muted) {
            muteBtn.innerHTML = '<i data-lucide="volume-x" style="width:16px; height:16px;"></i>';
            showToast('Musik dibisukan', 'info');
          } else {
            muteBtn.innerHTML = '<i data-lucide="volume-2" style="width:16px; height:16px;"></i>';
            showToast('Musik dinyalakan', 'info');
          }
          lucide.createIcons();
        }
      });
    }
  }

  // Background Particle Animation Canvas Injection
  if (pkg.enableAnim) {
    const bodyContainer = document.querySelector('.cbt-body');
    if (bodyContainer) {
      bodyContainer.style.position = 'relative';
      const canvas = document.createElement('canvas');
      canvas.id = 'cbt-animation-canvas';
      canvas.style.position = 'absolute';
      canvas.style.top = '0';
      canvas.style.left = '0';
      canvas.style.width = '100%';
      canvas.style.height = '100%';
      canvas.style.pointerEvents = 'none';
      canvas.style.zIndex = '0';
      canvas.style.opacity = '0.08';
      bodyContainer.appendChild(canvas);
      
      startCBTAnimation(canvas);
    }
  }

  // DOM selections inside CBT
  const timerClock = document.getElementById('cbt-timer-clock');
  const timerBox = document.getElementById('cbt-timer-mount');
  const submitBtn = document.getElementById('btn-submit-cbt');
  const qNumTitle = document.getElementById('cbt-question-number-title');
  const qTextMount = document.getElementById('cbt-question-text-mount');
  const optMount = document.getElementById('cbt-options-mount');
  const prevBtn = document.getElementById('btn-cbt-prev');
  const nextBtn = document.getElementById('btn-cbt-next');
  const numGridMount = document.getElementById('cbt-number-grid-mount');

  const doubtBtn = document.getElementById('btn-cbt-doubt');

  // Toggle Sidebar (Show/Hide Question Numbers Grid)
  const toggleSidebarBtn = document.getElementById('btn-toggle-sidebar');
  const cbtSidebar = document.querySelector('.cbt-sidebar');
  const toggleText = document.getElementById('text-toggle-sidebar');
  let sidebarCollapsed = false;

  toggleSidebarBtn.addEventListener('click', () => {
    sidebarCollapsed = !sidebarCollapsed;
    if (sidebarCollapsed) {
      cbtSidebar.classList.add('collapsed');
      toggleText.innerText = 'Tampilkan Nomor';
      toggleSidebarBtn.innerHTML = `
        <i data-lucide="layout-sidebar-off" style="width:16px; height:16px;"></i>
        <span id="text-toggle-sidebar">Tampilkan Nomor</span>
      `;
    } else {
      cbtSidebar.classList.remove('collapsed');
      toggleText.innerText = 'Sembunyikan Nomor';
      toggleSidebarBtn.innerHTML = `
        <i data-lucide="layout-sidebar" style="width:16px; height:16px;"></i>
        <span id="text-toggle-sidebar">Sembunyikan Nomor</span>
      `;
    }
    lucide.createIcons();
  });

  // Renders the Grid of question numbers
  const renderNumberGrid = () => {
    numGridMount.innerHTML = pkgQuestions.map((q, idx) => {
      let statusClass = '';
      if (idx === currentIdx) {
        statusClass = 'active';
      } else {
        const ans = answers[q.id];
        const isAnswered = Array.isArray(ans) ? ans.length > 0 : ans !== '';
        if (isAnswered) statusClass = 'answered';
      }

      if (doubtful[q.id]) {
        statusClass += ' doubtful';
      }

      return `
        <button class="cbt-num-btn ${statusClass}" data-index="${idx}">${idx + 1}</button>
      `;
    }).join('');

    // Bind click events on numbers
    numGridMount.querySelectorAll('.cbt-num-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        saveCurrentAnswer();
        currentIdx = parseInt(btn.dataset.index);
        renderQuestion();
      });
    });
  };

  // Saves current selected/typed answer to memory
  const saveCurrentAnswer = () => {
    const q = pkgQuestions[currentIdx];
    if (q.type === 'PG') {
      const selected = document.querySelector('input[name="cbt-choice"]:checked');
      answers[q.id] = selected ? selected.value : '';
    } else if (q.type === 'PG_KOMPLEKS') {
      const checked = Array.from(document.querySelectorAll('.cbt-complex-chk:checked')).map(chk => chk.value);
      answers[q.id] = checked;
    } else if (q.type === 'BENAR_SALAH') {
      const selected = document.querySelector('input[name="cbt-bs"]:checked');
      answers[q.id] = selected ? selected.value : '';
    } else if (q.type === 'ISIAN_SINGKAT' || q.type === 'NUMERIK') {
      const input = document.getElementById('cbt-input-ans');
      answers[q.id] = input ? input.value.trim() : '';
    } else if (q.type === 'ESSAY' || q.type === 'URAIAN') {
      const area = document.getElementById('cbt-area-ans');
      answers[q.id] = area ? area.value.trim() : '';
    }
  };

  // Renders active question content
  const renderQuestion = () => {
    const q = pkgQuestions[currentIdx];
    qNumTitle.innerText = `Soal Nomor ${currentIdx + 1} dari ${pkgQuestions.length}`;
    qTextMount.innerHTML = q.questionText;

    // Build specific answer fields based on question type
    let optHtml = '';
    const storedAns = answers[q.id];

    if (q.type === 'PG' && q.choices) {
      optHtml = `<div class="cbt-options-list">`;
      Object.entries(q.choices).forEach(([key, text]) => {
        const isSelected = storedAns === key ? 'selected' : '';
        optHtml += `
          <label class="cbt-option-item ${isSelected}" data-choice-key="${key}">
            <input type="radio" name="cbt-choice" value="${key}" style="display:none;" ${storedAns === key ? 'checked' : ''}>
            <strong style="min-width:20px;">${key}.</strong>
            <span>${text}</span>
          </label>
        `;
      });
      optHtml += `</div>`;
    } else if (q.type === 'PG_KOMPLEKS' && q.choices) {
      optHtml = `
        <div style="font-size:12px; color:var(--neutral-600); margin-bottom:12px; font-style:italic;">(Pilihlah beberapa jawaban yang benar)</div>
        <div class="cbt-options-list">
      `;
      Object.entries(q.choices).forEach(([key, text]) => {
        const isSelected = Array.isArray(storedAns) && storedAns.includes(key) ? 'selected' : '';
        optHtml += `
          <label class="cbt-option-item ${isSelected}" data-complex-key="${key}">
            <input type="checkbox" class="cbt-complex-chk" value="${key}" style="display:none;" ${Array.isArray(storedAns) && storedAns.includes(key) ? 'checked' : ''}>
            <strong style="min-width:20px;">[  ] ${key}.</strong>
            <span>${text}</span>
          </label>
        `;
      });
      optHtml += `</div>`;
    } else if (q.type === 'BENAR_SALAH') {
      const isBSelected = storedAns === 'BENAR' ? 'selected' : '';
      const isSSelected = storedAns === 'SALAH' ? 'selected' : '';
      optHtml = `
        <div class="cbt-options-list" style="flex-direction:row; gap:20px;">
          <label class="cbt-option-item ${isBSelected}" data-bs-val="BENAR" style="flex:1; justify-content:center;">
            <input type="radio" name="cbt-bs" value="BENAR" style="display:none;" ${storedAns === 'BENAR' ? 'checked' : ''}>
            <span>BENAR</span>
          </label>
          <label class="cbt-option-item ${isSSelected}" data-bs-val="SALAH" style="flex:1; justify-content:center;">
            <input type="radio" name="cbt-bs" value="SALAH" style="display:none;" ${storedAns === 'SALAH' ? 'checked' : ''}>
            <span>SALAH</span>
          </label>
        </div>
      `;
    } else if (q.type === 'ISIAN_SINGKAT') {
      optHtml = `
        <div class="form-group">
          <label class="form-label" for="cbt-input-ans">Ketik Jawaban Singkat Anda:</label>
          <input type="text" id="cbt-input-ans" class="form-input" value="${storedAns || ''}" placeholder="Ketik jawaban di sini..." style="width:100%; max-width:400px;">
        </div>
      `;
    } else if (q.type === 'NUMERIK') {
      optHtml = `
        <div class="form-group">
          <label class="form-label" for="cbt-input-ans">Ketik Jawaban Angka:</label>
          <input type="number" step="any" id="cbt-input-ans" class="form-input" value="${storedAns || ''}" placeholder="Ketik angka saja..." style="width:100%; max-width:400px;">
        </div>
      `;
    } else if (q.type === 'ESSAY' || q.type === 'URAIAN') {
      optHtml = `
        <div class="form-group">
          <label class="form-label" for="cbt-area-ans">Tuliskan Jawaban Uraian Anda secara Lengkap:</label>
          <textarea id="cbt-area-ans" class="form-input" style="height:180px; width:100%; resize:vertical;" placeholder="Tuliskan penjelasan langkah-langkah di sini...">${storedAns || ''}</textarea>
        </div>
      `;
    }

    optMount.innerHTML = optHtml;

    // Bind custom selection item behaviors (visual selections instead of raw radios/checks)
    if (q.type === 'PG') {
      const items = optMount.querySelectorAll('.cbt-option-item');
      items.forEach(item => {
        item.addEventListener('click', () => {
          items.forEach(i => i.classList.remove('selected'));
          item.classList.add('selected');
          const radio = item.querySelector('input');
          if (radio) radio.checked = true;
          saveCurrentAnswer();
          renderNumberGrid();
        });
      });
    } else if (q.type === 'PG_KOMPLEKS') {
      const items = optMount.querySelectorAll('.cbt-option-item');
      items.forEach(item => {
        item.addEventListener('click', (e) => {
          e.preventDefault(); // Prevent double trigger
          item.classList.toggle('selected');
          const chk = item.querySelector('input');
          if (chk) chk.checked = !chk.checked;
          saveCurrentAnswer();
          renderNumberGrid();
        });
      });
    } else if (q.type === 'BENAR_SALAH') {
      const items = optMount.querySelectorAll('.cbt-option-item');
      items.forEach(item => {
        item.addEventListener('click', () => {
          items.forEach(i => i.classList.remove('selected'));
          item.classList.add('selected');
          const radio = item.querySelector('input');
          if (radio) radio.checked = true;
          saveCurrentAnswer();
          renderNumberGrid();
        });
      });
    } else {
      const input = document.getElementById('cbt-input-ans') || document.getElementById('cbt-area-ans');
      if (input) {
        input.addEventListener('input', () => {
          saveCurrentAnswer();
          renderNumberGrid();
        });
      }
    }

    // Configure Nav buttons
    prevBtn.disabled = currentIdx === 0;
    nextBtn.innerHTML = currentIdx === pkgQuestions.length - 1
      ? `<span>Selesai</span> <i data-lucide="check"></i>`
      : `<span>Soal Berikutnya</span> <i data-lucide="chevron-right"></i>`;
    lucide.createIcons();

    // Update doubt button styling
    const isDoubtful = doubtful[q.id];
    if (isDoubtful) {
      doubtBtn.style.backgroundColor = '#dd6b20';
      doubtBtn.style.color = 'white';
      doubtBtn.innerHTML = `<i data-lucide="help-circle" style="width:16px; height:16px;"></i> <span>Ragu-Ragu (Aktif)</span>`;
    } else {
      doubtBtn.style.backgroundColor = '#feebc8';
      doubtBtn.style.color = '#dd6b20';
      doubtBtn.innerHTML = `<i data-lucide="help-circle" style="width:16px; height:16px;"></i> <span>Ragu-Ragu</span>`;
    }

    renderNumberGrid();

    // LaTeX typeset
    if (window.MathJax && window.MathJax.typesetPromise) {
      window.MathJax.typesetPromise([qTextMount, optMount]).catch(err => console.error(err));
    }
  };

  // Nav button listeners
  prevBtn.addEventListener('click', () => {
    saveCurrentAnswer();
    if (currentIdx > 0) {
      currentIdx--;
      renderQuestion();
    }
  });

  nextBtn.addEventListener('click', () => {
    saveCurrentAnswer();
    if (currentIdx < pkgQuestions.length - 1) {
      currentIdx++;
      renderQuestion();
    } else {
      // Trigger submission confirmation
      submitCBTPaper();
    }
  });

  // Ragu-Ragu button listener
  doubtBtn.addEventListener('click', () => {
    const q = pkgQuestions[currentIdx];
    doubtful[q.id] = !doubtful[q.id];

    // Toggle active styling
    const isDoubtful = doubtful[q.id];
    if (isDoubtful) {
      doubtBtn.style.backgroundColor = '#dd6b20';
      doubtBtn.style.color = 'white';
      doubtBtn.innerHTML = `<i data-lucide="help-circle" style="width:16px; height:16px;"></i> <span>Ragu-Ragu (Aktif)</span>`;
    } else {
      doubtBtn.style.backgroundColor = '#feebc8';
      doubtBtn.style.color = '#dd6b20';
      doubtBtn.innerHTML = `<i data-lucide="help-circle" style="width:16px; height:16px;"></i> <span>Ragu-Ragu</span>`;
    }
    lucide.createIcons();
    renderNumberGrid();
  });

  // Submission handler
  const submitCBTPaper = () => {
    saveCurrentAnswer();

    // Count unanswered
    let unanswered = 0;
    pkgQuestions.forEach(q => {
      const ans = answers[q.id];
      const isAnswered = Array.isArray(ans) ? ans.length > 0 : ans !== '';
      if (!isAnswered) unanswered++;
    });

    const msg = unanswered > 0
      ? `Terdapat ${unanswered} soal yang belum dijawab. Apakah Anda yakin ingin menyelesaikan ujian?`
      : 'Apakah Anda yakin ingin menyelesaikan ujian dan mengirim lembar jawaban Anda?';

    if (confirm(msg)) {
      evaluateAndSaveCBTResults();
    }
  };

  submitBtn.addEventListener('click', submitCBTPaper);

  // CBT scoring evaluator
  const evaluateAndSaveCBTResults = () => {
    // Clear timer loop
    if (cbtInterval) {
      clearInterval(cbtInterval);
    }

    let correctCount = 0;
    let incorrectCount = 0;
    let unansweredCount = 0;

    pkgQuestions.forEach(q => {
      const studentAns = answers[q.id];
      const correctAns = q.correctAnswer;

      const isAnsEmpty = Array.isArray(studentAns) ? studentAns.length === 0 : studentAns === '';
      if (isAnsEmpty) {
        unansweredCount++;
        return;
      }

      // Perform question scoring evaluation by type
      if (q.type === 'PG' || q.type === 'BENAR_SALAH' || q.type === 'ISIAN_SINGKAT' || q.type === 'NUMERIK') {
        const isCorrect = String(studentAns).trim().toUpperCase() === String(correctAns).trim().toUpperCase();
        if (isCorrect) correctCount++;
        else incorrectCount++;
      } else if (q.type === 'PG_KOMPLEKS') {
        // Compare arrays
        const sortedStudent = [...studentAns].sort();
        const sortedCorrect = [...correctAns].sort();
        const isCorrect = sortedStudent.length === sortedCorrect.length && sortedStudent.every((val, idx) => val === sortedCorrect[idx]);
        if (isCorrect) correctCount++;
        else incorrectCount++;
      } else if (q.type === 'MENJODOHKAN') {
        // Compare items
        let allPairsMatch = true;
        (correctAns || []).forEach(pair => {
          const match = (studentAns || []).find(p => p.premise === pair.premise);
          if (!match || match.response !== pair.response) {
            allPairsMatch = false;
          }
        });
        if (allPairsMatch) correctCount++;
        else incorrectCount++;
      } else {
        // For essays/uraian, count as correct if anything was written (self-practice helper)
        correctCount++;
      }
    });

    const totalQ = pkgQuestions.length;
    const rawScore = (correctCount / totalQ) * 100;
    const finalScore = Math.round(rawScore);

    const result = {
      id: `res-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      userId: activeUser.id,
      userName: activeUser.name,
      packageId: pkg.id,
      packageName: pkg.name,
      subjectName: subjectName,
      className: className,
      score: finalScore,
      totalQuestions: totalQ,
      correctCount,
      incorrectCount,
      unansweredCount,
      answers: answers,
      timestamp: new Date().toISOString()
    };

    db.insert('results', result);
    db.log(activeUser.id, activeUser.name, 'SELESAI_UJIAN', `Menyelesaikan ujian "${pkg.name}" dengan nilai ${finalScore}`);

    showToast('Jawaban dikirim! Skor berhasil dikalkulasi.', 'success');
    window.location.hash = `#cbt-result?id=${result.id}`;
  };

  // Timer loop initialization
  const updateTimerClock = () => {
    if (timeLeft <= 0) {
      clearInterval(cbtInterval);
      showToast('Waktu Ujian Habis! Jawaban Anda otomatis terkirim.', 'error');
      evaluateAndSaveCBTResults();
      return;
    }

    timeLeft--;

    // Format hours, minutes, seconds
    const hrs = Math.floor(timeLeft / 3600);
    const mins = Math.floor((timeLeft % 3600) / 60);
    const secs = timeLeft % 60;

    const formatted = [
      String(hrs).padStart(2, '0'),
      String(mins).padStart(2, '0'),
      String(secs).padStart(2, '0')
    ].join(':');

    timerClock.innerText = formatted;

    // Visual warning when time is low (< 5 minutes)
    if (timeLeft < 300) {
      timerBox.classList.add('timer-danger');
    }
  };

  // Initial draw
  updateTimerClock();
  renderQuestion();

  // Run timer interval loop
  cbtInterval = setInterval(updateTimerClock, 1000);
}

// ----------------------------------------------------
// 14. SISWA: CBT RESULT SCORE & ANSWER KEY REVIEW PAGE
// ----------------------------------------------------
function renderCBTResultScreen(resultId) {
  // Clear cbt interval in case redirect
  if (cbtInterval) {
    clearInterval(cbtInterval);
  }

  const results = db.get('results');
  const res = results.find(r => r.id === resultId);
  const pkg = db.get('packages').find(p => p.id === res?.packageId);
  const allQ = db.get('questions');
  const subjects = db.get('subjects');

  if (!res) {
    showToast('Hasil ujian tidak ditemukan!', 'error');
    window.location.hash = '#dashboard';
    return;
  }

  const pkgQuestions = (pkg?.questionIds || []).map(id => allQ.find(q => q.id === id)).filter(Boolean);

  const isPassed = res.score >= 75;
  const scoreColor = isPassed ? 'var(--success-text)' : 'var(--danger-text)';
  const badgeClass = isPassed ? 'badge-success' : 'badge-danger';
  const scoreLabel = isPassed ? 'LULUS KKM' : 'REMIDI';

  appMount.innerHTML = `
    <!-- Distraction-free result layout inside #app -->
    <div style="background-color: #f8fafc; min-height:100vh; padding: 40px 20px; font-family: 'Plus Jakarta Sans', sans-serif;">
      <div style="max-width: 800px; margin: 0 auto; display:flex; flex-direction:column; gap:28px;">
        
        <!-- Score summary card -->
        <div class="card" style="border-color: ${isPassed ? 'var(--success)' : 'var(--danger)'};">
          <div class="card-body" style="padding: 40px; text-align:center; display:flex; flex-direction:column; align-items:center; gap:20px;">
            <span style="font-size:12px; font-weight:700; color:var(--neutral-600); text-transform:uppercase; letter-spacing:1px;">Hasil Evaluasi Ujian</span>
            <h2 style="font-size:20px; font-weight:800; color:var(--neutral-900);">${res.packageName}</h2>
            
            <div style="width: 140px; height: 140px; border-radius: 50%; border: 6px solid ${isPassed ? '#e6fffa' : '#fff5f5'}; background: white; display: flex; align-items: center; justify-content: center; box-shadow: var(--shadow-md); margin: 10px 0;">
              <span style="font-size: 48px; font-weight: 900; color: ${scoreColor};">${res.score}</span>
            </div>
            
            <span class="badge ${badgeClass}" style="padding:6px 16px; font-size:12px;">${scoreLabel}</span>
            
            <!-- Details metrics row -->
            <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap:16px; width:100%; max-width:500px; margin-top:10px; border-top:1px solid var(--neutral-400); padding-top:20px;">
              <div>
                <div style="font-size:11px; color:var(--neutral-600); font-weight:600;">Total Soal</div>
                <div style="font-size:18px; font-weight:800; margin-top:4px; color:var(--neutral-900);">${res.totalQuestions}</div>
              </div>
              <div>
                <div style="font-size:11px; color:var(--neutral-600); font-weight:600;">Benar</div>
                <div style="font-size:18px; font-weight:800; margin-top:4px; color:var(--success-text);">${res.correctCount}</div>
              </div>
              <div>
                <div style="font-size:11px; color:var(--neutral-600); font-weight:600;">Salah</div>
                <div style="font-size:18px; font-weight:800; margin-top:4px; color:var(--danger-text);">${res.incorrectCount}</div>
              </div>
              <div>
                <div style="font-size:11px; color:var(--neutral-600); font-weight:600;">Kosong</div>
                <div style="font-size:18px; font-weight:800; margin-top:4px; color:var(--neutral-700);">${res.unansweredCount}</div>
              </div>
            </div>
            
            <a href="#dashboard" class="btn btn-primary" style="margin-top:15px; padding: 12px 30px;">
              <i data-lucide="layout-dashboard"></i>
              <span>Kembali ke Dashboard</span>
            </a>
          </div>
        </div>

        <!-- Answers Review Section -->
        <div>
          <h3 style="font-size:18px; font-weight:800; margin-bottom:16px; color:var(--neutral-900);">Tinjauan & Pembahasan Soal</h3>
          <div style="display:flex; flex-direction:column; gap:20px;" id="cbt-review-list-mount">
            ${pkgQuestions.map((q, idx) => {
    const studAns = res.answers[q.id];
    const isCorrect = db.get('questions').find(soal => soal.id === q.id)?.type === 'ESSAY'
      ? true // Self practice essays counted correct for display
      : (q.type === 'PG_KOMPLEKS'
        ? [...(studAns || [])].sort().join() === [...(q.correctAnswer || [])].sort().join()
        : String(studAns).trim().toUpperCase() === String(q.correctAnswer).trim().toUpperCase());

    const isAnswerEmpty = Array.isArray(studAns) ? studAns.length === 0 : studAns === '';

    let statusLabelHtml = '';
    let cardBorderColor = '';
    if (isAnswerEmpty) {
      statusLabelHtml = `<span class="badge badge-neutral" style="margin-bottom:8px;">KOSONG (TIDAK DIJAWAB)</span>`;
      cardBorderColor = 'border-left: 5px solid var(--neutral-600);';
    } else if (isCorrect) {
      statusLabelHtml = `<span class="badge badge-success" style="margin-bottom:8px;"><i data-lucide="check" style="width:12px; height:12px; display:inline; margin-right:4px;"></i> BENAR</span>`;
      cardBorderColor = 'border-left: 5px solid var(--success-text);';
    } else {
      statusLabelHtml = `<span class="badge badge-danger" style="margin-bottom:8px;"><i data-lucide="x" style="width:12px; height:12px; display:inline; margin-right:4px;"></i> SALAH</span>`;
      cardBorderColor = 'border-left: 5px solid var(--danger);';
    }

    // Format Answers strings
    let formattedStudentAns = studAns;
    if (Array.isArray(studAns)) formattedStudentAns = studAns.join(', ');
    if (isAnswerEmpty) formattedStudentAns = '<em>Tidak ada jawaban</em>';

    let formattedCorrectAns = q.correctAnswer;
    if (Array.isArray(q.correctAnswer)) formattedCorrectAns = q.correctAnswer.join(', ');

    let optionsStr = '';
    if (q.choices && (q.type === 'PG' || q.type === 'PG_KOMPLEKS')) {
      optionsStr = `
                  <div class="choices-preview-list" style="margin-top:12px; margin-bottom:12px;">
                    ${Object.entries(q.choices).map(([k, v]) => `
                      <div class="choice-preview-item ${k === q.correctAnswer || (Array.isArray(q.correctAnswer) && q.correctAnswer.includes(k)) ? 'correct' : ''}">
                        <strong>${k}.</strong> ${v}
                      </div>
                    `).join('')}
                  </div>
                `;
    }

    return `
                <div class="card" style="${cardBorderColor}">
                  <div class="card-body" style="padding:24px;">
                    ${statusLabelHtml}
                    <div style="font-weight:700; color:var(--neutral-600); font-size:11px; margin-bottom:8px;">Soal Nomor ${idx + 1} | ${q.code}</div>
                    
                    <div style="font-size:14px; color:var(--neutral-900); line-height:1.5;">${q.questionText}</div>
                    
                    ${optionsStr}
                    
                    <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:12px; font-size:13px; border-top:1px solid var(--neutral-400); padding-top:12px;">
                      <div>
                        <span style="color:var(--neutral-600); font-size:11px; font-weight:600; display:block;">Jawaban Anda:</span>
                        <strong style="color:${isCorrect && !isAnswerEmpty ? 'var(--success-text)' : 'var(--neutral-900)'};">${formattedStudentAns}</strong>
                      </div>
                      <div>
                        <span style="color:var(--neutral-600); font-size:11px; font-weight:600; display:block;">Kunci Jawaban:</span>
                        <strong style="color:var(--success-text);">${formattedCorrectAns}</strong>
                      </div>
                    </div>

                    <!-- Pembahasan -->
                    ${q.discussion ? `
                      <div style="background-color: var(--neutral-100); border-radius: var(--radius-sm); padding: 12px 16px; margin-top:16px; font-size:13px; color: var(--neutral-800); border-left:3px solid var(--neutral-600);">
                        <strong style="display:block; margin-bottom:4px; font-size:11px; color:var(--neutral-700);">Pembahasan:</strong>
                        <div>${q.discussion}</div>
                      </div>
                    ` : ''}
                  </div>
                </div>
              `;
  }).join('')}
          </div>
        </div>

      </div>
    </div>
  `;

  lucide.createIcons();

  // Trigger math typesetting on results list
  setTimeout(() => {
    const listMount = document.getElementById('cbt-review-list-mount');
    if (window.MathJax && window.MathJax.typesetPromise && listMount) {
      window.MathJax.typesetPromise([listMount]).catch(err => console.error(err));
    }
  }, 100);
}

// Canvas-based background float particle animator
function startCBTAnimation(canvas) {
  const ctx = canvas.getContext('2d');
  let animationFrameId;

  const resize = () => {
    if (canvas && canvas.parentElement) {
      canvas.width = canvas.parentElement.clientWidth;
      canvas.height = canvas.parentElement.clientHeight;
    }
  };
  resize();
  window.addEventListener('resize', resize);

  const particles = [];
  const particleCount = 20;

  for (let i = 0; i < particleCount; i++) {
    particles.push({
      x: Math.random() * (canvas.width || 800),
      y: Math.random() * (canvas.height || 600),
      r: Math.random() * 30 + 15,
      vx: (Math.random() - 0.5) * 0.3,
      vy: (Math.random() - 0.5) * 0.3,
      color: `hsl(${Math.random() * 360}, 65%, 80%)`
    });
  }

  const animate = () => {
    if (!canvas || !canvas.parentElement) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    particles.forEach(p => {
      p.x += p.vx;
      p.y += p.vy;

      // Wrap-around boundary handling
      if (p.x < -p.r) p.x = canvas.width + p.r;
      if (p.x > canvas.width + p.r) p.x = -p.r;
      if (p.y < -p.r) p.y = canvas.height + p.r;
      if (p.y > canvas.height + p.r) p.y = -p.r;

      ctx.beginPath();
      ctx.arc(p.x, p.y, p.r, 0, Math.PI * 2);
      ctx.fillStyle = p.color;
      ctx.fill();
    });
    animationFrameId = requestAnimationFrame(animate);
  };
  animate();

  canvas.cleanup = () => {
    cancelAnimationFrame(animationFrameId);
    window.removeEventListener('resize', resize);
  };
}

