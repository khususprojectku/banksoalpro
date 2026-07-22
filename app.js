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
let currentCaptchaCode = ''; // CAPTCHA code untuk halaman login

// DOM Selectors
const appMount = document.getElementById('app');
const globalModal = document.getElementById('global-modal');
const modalContent = document.getElementById('modal-content-container');
const toastContainer = document.getElementById('toast-container');

// Core Router
function handleRouting() {
  activeUser = JSON.parse(sessionStorage.getItem('active_user') || 'null');

  // Parse Hash
  const hash = window.location.hash || '#home';
  const page = hash.split('?')[0];
  const queryParams = parseQueryParams(hash);

  if (!activeUser || page === '#home' || page === '#login') {
    if (cbtAudio) {
      cbtAudio.pause();
      cbtAudio = null;
    }
    const canvas = document.getElementById('cbt-animation-canvas');
    if (canvas && canvas.cleanup) {
      canvas.cleanup();
      canvas.remove();
    }
    if (!activeUser || page === '#home' || page === '#login') {
      renderLandingPage();
      return;
    }
  }

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
  } else if (page === '#android-exam') {
    renderAndroidExamAppScreen(queryParams.id);
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

// ─── Global Error Handler: Tampilkan error langsung di layar ───
window.onerror = function(msg, src, line, col, err) {
  var appEl = document.getElementById('app');
  if (appEl) {
    appEl.innerHTML = '<div style="padding:40px;background:#1e293b;min-height:100vh;font-family:monospace;">' +
      '<h2 style="color:#f87171;margin:0 0 16px">⛔ JavaScript Error</h2>' +
      '<p style="color:#fbbf24;font-size:14px;margin:0 0 8px"><b>Pesan:</b> ' + msg + '</p>' +
      '<p style="color:#94a3b8;font-size:12px;margin:0 0 4px"><b>File:</b> ' + src + '</p>' +
      '<p style="color:#94a3b8;font-size:12px;margin:0 0 16px"><b>Baris:</b> ' + line + ', Kolom: ' + col + '</p>' +
      '<p style="color:#64748b;font-size:11px">Buka DevTools (F12) → Console tab untuk detail lengkap.</p>' +
      '</div>';
  }
  return false;
};

// Router Event Listeners
window.addEventListener('hashchange', handleRouting);

// Jalankan routing: DOM sudah siap karena script ada di bawah <body>
try {
  handleRouting();
} catch(e) {
  var appEl2 = document.getElementById('app');
  if (appEl2) {
    appEl2.innerHTML = '<div style="padding:40px;background:#1e293b;min-height:100vh;font-family:monospace;">' +
      '<h2 style="color:#f87171;margin:0 0 16px">⛔ Error saat Routing</h2>' +
      '<p style="color:#fbbf24;font-size:14px;margin:0 0 8px"><b>Pesan:</b> ' + e.message + '</p>' +
      '<pre style="color:#94a3b8;font-size:11px;overflow:auto;background:#0f172a;padding:12px;border-radius:8px">' + e.stack + '</pre>' +
      '</div>';
  }
}

// Safe Lucide icon renderer helper
function refreshIcons() {
  if (window.lucide && typeof window.lucide.createIcons === 'function') {
    try {
      window.lucide.createIcons();
    } catch (e) {
      console.warn('Lucide icons load warning:', e);
    }
  }
}

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
  refreshIcons();

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
  refreshIcons();

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
// 1. RENDER HOMEPAGE LANDING PAGE & DIRECT LOGIN WITH CAPTCHA
// ----------------------------------------------------

function renderLandingPage() {
  currentCaptchaCode = db.generateCaptcha();

  appMount.innerHTML = `
    <div class="landing-page">
      <!-- Navbar -->
      <nav class="landing-nav">
        <div style="display:flex; align-items:center; gap:12px;">
          <div class="logo-icon">BSP</div>
          <div style="font-weight:900; font-size:20px; color:white; letter-spacing:-0.5px;">BankSoalPro</div>
        </div>

        <div class="landing-nav-links">
          <a href="#features" class="landing-nav-link">Fitur Unggulan</a>
          <a href="#pricing" class="landing-nav-link">Harga & Paket</a>
          <a href="#testimonials" class="landing-nav-link">Testimoni</a>
          <a href="#faq" class="landing-nav-link">FAQ</a>
          ${activeUser ? `
            <a href="#dashboard" class="btn btn-primary" style="padding:8px 20px;">
              <i data-lucide="layout-dashboard"></i> Buka Dashboard
            </a>
          ` : `
            <a href="#login-card-hero" class="btn btn-primary" style="padding:8px 20px;">
              <i data-lucide="log-in"></i> Masuk Sekarang
            </a>
          `}
        </div>
      </nav>

      <!-- Hero Section -->
      <header class="landing-hero">
        <div>
          <div class="hero-badge">
            <i data-lucide="sparkles" style="width:14px; height:14px;"></i>
            <span>PLATFORM CBT & BANK SOAL AI NO. 1 DI INDONESIA</span>
          </div>

          <h1 class="hero-title">
            Kelola Bank Soal & Ujian CBT Sekolah <span class="hero-title-highlight"> Lebih Cepat, Aman, & Otomatis</span>
          </h1>

          <p class="hero-desc">
            BankSoalPro adalah solusi modern untuk Guru, Sekolah, dan Yayasan dalam mengolah bank soal, ekspor format MS Word/Excel, penyusunan kisi-kisi, hingga pelaksanaan ujian CBT Android Anti-Cheat berstandar nasional.
          </p>

          <div style="display:flex; gap:16px; align-items:center; flex-wrap:wrap;">
            <a href="#pricing" class="btn btn-primary" style="padding:14px 28px; font-size:15px; font-weight:800; border-radius:12px;">
              <span>Mulai Berlangganan Sekarang</span>
              <i data-lucide="arrow-right"></i>
            </a>
            <a href="#features" class="btn btn-secondary" style="padding:14px 24px; font-size:15px; font-weight:700; color:white; border-color:rgba(255,255,255,0.2); border-radius:12px;">
              <span>Pelajari Fitur</span>
            </a>
          </div>

          <!-- Trust Badges -->
          <div style="display:flex; gap:32px; margin-top:40px; border-top:1px solid rgba(255,255,255,0.1); padding-top:24px;">
            <div>
              <div style="font-size:24px; font-weight:900; color:white;">500+</div>
              <div style="font-size:12px; color:#94a3b8;">Sekolah Terdaftar</div>
            </div>
            <div>
              <div style="font-size:24px; font-weight:900; color:#38bdf8;">50.000+</div>
              <div style="font-size:12px; color:#94a3b8;">Butir Soal Terarsip</div>
            </div>
            <div>
              <div style="font-size:24px; font-weight:900; color:#818cf8;">99.9%</div>
              <div style="font-size:12px; color:#94a3b8;">Keamanan Anti-Cheat</div>
            </div>
          </div>
        </div>

        <!-- Login & Student Exam Entry -->
        <div class="landing-login-card" id="login-card-hero" style="padding:0; overflow:hidden;">
          <div style="display:flex; border-bottom:1px solid rgba(255,255,255,0.1); background:rgba(15,23,42,0.8);">
            <button id="tab-login" class="tab-btn active" style="flex:1; padding:16px; background:transparent; border:none; color:white; font-weight:700; cursor:pointer; border-bottom:2px solid var(--primary);">Login Sistem</button>
            <button id="tab-exam" class="tab-btn" style="flex:1; padding:16px; background:transparent; border:none; color:#cbd5e1; font-weight:700; cursor:pointer; border-bottom:2px solid transparent;">Ujian Siswa</button>
          </div>

          <!-- PANE: LOGIN SYSTEM -->
          <div id="pane-login" style="padding:24px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
              <div>
                <h3 style="font-size:20px; font-weight:800; color:white; margin:0;">Direct Login Akun</h3>
                <p style="font-size:12px; color:#94a3b8; margin-top:4px;">Masuk langsung untuk mencoba platform</p>
              </div>
              <span class="cyber-status-pill" style="background:rgba(56,189,248,0.15); color:#38bdf8; border:1px solid rgba(56,189,248,0.3);">
                <i data-lucide="shield-check" style="width:12px; height:12px;"></i> CAPTCHA
              </span>
            </div>

            <form id="landing-login-form" style="display:flex; flex-direction:column; gap:16px;">
              <div class="form-group">
                <label class="form-label" for="l-email" style="color:#cbd5e1;">Alamat Email</label>
                <input type="email" id="l-email" class="form-input" style="background:rgba(30,41,59,0.8); border-color:rgba(255,255,255,0.15); color:white;" placeholder="email@sekolah.sch.id" required>
              </div>
              <div class="form-group">
                <label class="form-label" for="l-pass" style="color:#cbd5e1;">Kata Sandi</label>
                <input type="password" id="l-pass" class="form-input" style="background:rgba(30,41,59,0.8); border-color:rgba(255,255,255,0.15); color:white;" placeholder="••••••••" required>
              </div>
              <div class="form-group">
                <label class="form-label" style="color:#cbd5e1; display:flex; justify-content:space-between;">
                  <span>Verifikasi CAPTCHA</span>
                </label>
                <div class="captcha-box">
                  <div class="captcha-display" id="captcha-code-mount">${currentCaptchaCode}</div>
                  <button type="button" class="btn-refresh-captcha" id="btn-refresh-captcha" title="Acak Ulang Kode CAPTCHA">
                    <i data-lucide="refresh-cw" style="width:18px; height:18px;"></i>
                  </button>
                </div>
                <input type="text" id="l-captcha-input" class="form-input" style="background:rgba(30,41,59,0.8); border-color:rgba(99,102,241,0.4); color:white; font-weight:800; letter-spacing:2px; text-transform:uppercase; margin-top:8px;" placeholder="Ketik Kode CAPTCHA" required>
              </div>
              <button type="submit" class="btn btn-primary" style="width:100%; padding:12px; font-weight:800; margin-top:6px;">
                <i data-lucide="shield-check"></i> Verifikasi & Masuk
              </button>
            </form>

            <div style="margin-top:20px; border-top:1px dashed rgba(255,255,255,0.1); padding-top:16px;">
              <div style="font-size:11px; font-weight:700; color:#94a3b8; margin-bottom:10px; text-transform:uppercase;">Quick Login Demo:</div>
              <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
                <button class="btn-demo quick-landing-demo" data-email="admin@banksoal.pro" style="background:rgba(255,255,255,0.05); color:#cbd5e1; border:1px solid rgba(255,255,255,0.1); font-size:11px; padding:6px;">Super Admin</button>
                <button class="btn-demo quick-landing-demo" data-email="admin.sekolah@sman1.sch.id" style="background:rgba(255,255,255,0.05); color:#cbd5e1; border:1px solid rgba(255,255,255,0.1); font-size:11px; padding:6px;">Admin Sekolah</button>
                <button class="btn-demo quick-landing-demo" data-email="budi.guru@sman1.sch.id" style="background:rgba(255,255,255,0.05); color:#cbd5e1; border:1px solid rgba(255,255,255,0.1); font-size:11px; padding:6px;">Guru Info</button>
                <button class="btn-demo quick-landing-demo" data-email="edi.reviewer@sman1.sch.id" style="background:rgba(255,255,255,0.05); color:#cbd5e1; border:1px solid rgba(255,255,255,0.1); font-size:11px; padding:6px;">Reviewer</button>
              </div>
            </div>
          </div>

          <!-- PANE: STUDENT EXAM CBT -->
          <div id="pane-exam" style="padding:24px; display:none;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
              <div>
                <h3 style="font-size:20px; font-weight:800; color:white; margin:0;">Portal Ujian Siswa</h3>
                <p style="font-size:12px; color:#94a3b8; margin-top:4px;">Masukkan Kode Ujian & Token</p>
              </div>
              <span class="cyber-status-pill" style="background:rgba(56,189,248,0.15); color:#38bdf8; border:1px solid rgba(56,189,248,0.3);">
                <i data-lucide="scan" style="width:12px; height:12px;"></i> KIOSK
              </span>
            </div>

            <form id="student-exam-form" style="display:flex; flex-direction:column; gap:16px;">
              <div class="form-group">
                <label class="form-label" style="color:#cbd5e1;">1. Kode Paket Ujian</label>
                <input type="text" id="s-exam-code" class="form-input" style="background:rgba(30,41,59,0.8); border-color:rgba(255,255,255,0.15); color:white; text-transform:uppercase; font-weight:bold; letter-spacing:1px;" placeholder="EXAM-UAS-1234" required>
              </div>
              
              <div class="form-group" id="token-group" style="display:none;">
                <label class="form-label" style="color:#cbd5e1;">2. Token Akses Ujian (Dari Guru)</label>
                <input type="text" id="s-exam-token" class="form-input" style="background:rgba(30,41,59,0.8); border-color:var(--primary); color:white; text-transform:uppercase; font-weight:bold; letter-spacing:1px;" placeholder="TOKEN-123">
              </div>

              <div id="student-login-group" style="display:none; flex-direction:column; gap:16px; border-top:1px dashed rgba(255,255,255,0.2); padding-top:16px; margin-top:8px;">
                <label class="form-label" style="color:#cbd5e1;">3. Kredensial Siswa</label>
                <input type="email" id="s-email" class="form-input" style="background:rgba(30,41,59,0.8); border-color:rgba(255,255,255,0.15); color:white;" placeholder="Email Siswa (cth: andi.siswa@...)">
                <input type="password" id="s-pass" class="form-input" style="background:rgba(30,41,59,0.8); border-color:rgba(255,255,255,0.15); color:white;" placeholder="Password">
              </div>

              <button type="submit" id="btn-student-flow" class="btn btn-primary" style="width:100%; padding:12px; font-weight:800; margin-top:6px;">
                Cek Kode Ujian <i data-lucide="arrow-right"></i>
              </button>
            </form>
            
            <div style="margin-top:20px; border-top:1px dashed rgba(255,255,255,0.1); padding-top:16px;">
               <div style="font-size:11px; font-weight:700; color:#94a3b8; margin-bottom:10px; text-transform:uppercase;">Hint Demo Ujian Siswa:</div>
               <p style="font-size:11px; color:#cbd5e1; line-height:1.5;">
                 Kode: <strong>EXAM-INF-8921</strong> &rarr; Token: <strong>INF001</strong><br>
                 Email: <strong>andi.siswa@sman1.sch.id</strong> / Pass: <strong>siswa123</strong>
               </p>
            </div>
          </div>
        </div>
      </header>

      <!-- Features Section -->
      <section class="features-section" id="features">
        <div class="section-header">
          <div class="section-subtitle">FITUR UNGGULAN APLIKASI</div>
          <h2 class="section-title">Mengapa BankSoalPro Menjadi Pilihan Utama Sekolah?</h2>
        </div>

        <div class="features-grid">
          <div class="feature-card">
            <div class="feature-icon"><i data-lucide="building-2"></i></div>
            <h3 style="font-size:18px; font-weight:800; color:white;">Multi-Tenant Sekolah</h3>
            <p style="font-size:14px; color:#94a3b8; margin-top:8px; line-height:1.6;">
              Mendukung pengelolaan banyak sekolah sekaligus dalam 1 platform pusat dengan pemisah hak akses yang ketat.
            </p>
          </div>

          <div class="feature-card">
            <div class="feature-icon"><i data-lucide="sparkles"></i></div>
            <h3 style="font-size:18px; font-weight:800; color:white;">Generator Paket Soal Otomatis</h3>
            <p style="font-size:14px; color:#94a3b8; margin-top:8px; line-height:1.6;">
              Pengacakan butir soal dan pilihan jawaban secara instan berdasarkan proporsi kesulitan (Mudah, Sedang, Sulit).
            </p>
          </div>

          <div class="feature-card">
            <div class="feature-icon"><i data-lucide="smartphone"></i></div>
            <h3 style="font-size:18px; font-weight:800; color:white;">Android Exam App Kiosk Mode</h3>
            <p style="font-size:14px; color:#94a3b8; margin-top:8px; line-height:1.6;">
              Aplikasi khusus ujian di perangkat Android siswa yang mengunci layar (Lockdown), memblokir pindah tab, dan mendukung scan QR Code.
            </p>
          </div>

          <div class="feature-card">
            <div class="feature-icon"><i data-lucide="shield-check"></i></div>
            <h3 style="font-size:18px; font-weight:800; color:white;">Cyber Security & Anti-Cheat</h3>
            <p style="font-size:14px; color:#94a3b8; margin-top:8px; line-height:1.6;">
              Perlindungan dari hacker: Sanitizer Anti-XSS, Brute-force rate limiter, Watermark melayang Nama/IP, dan CAPTCHA verifikasi.
            </p>
          </div>

          <div class="feature-card">
            <div class="feature-icon"><i data-lucide="square-pi"></i></div>
            <h3 style="font-size:18px; font-weight:800; color:white;">Editor WYSIWYG & Rumus MathJax</h3>
            <p style="font-size:14px; color:#94a3b8; margin-top:8px; line-height:1.6;">
              Dukungan lengkap pembuatan rumus matematika, fisika, kimia, gambar, dan audio.
            </p>
          </div>

          <div class="feature-card">
            <div class="feature-icon"><i data-lucide="file-text"></i></div>
            <h3 style="font-size:18px; font-weight:800; color:white;">Ekspor 1-Klik MS Word, Excel & PDF</h3>
            <p style="font-size:14px; color:#94a3b8; margin-top:8px; line-height:1.6;">
              Ekspor naskah soal lengkap dengan kop sekolah, nomor urut otomatis, kunci jawaban, dan lembar pembahasan soal.
            </p>
          </div>
        </div>
      </section>

      <!-- Pricing Plans Section -->
      <section class="pricing-section" id="pricing">
        <div class="section-header">
          <div class="section-subtitle">PAKET HARGA BERLANGGANAN</div>
          <h2 class="section-title">Pilih Paket Langganan Sesuai Kebutuhan Sekolah</h2>
          <p style="color:#94a3b8; margin-top:8px;">Tanpa biaya tersembunyi. Dapatkan jaminan update fitur dan dukungan teknis 24/7.</p>
        </div>

        <div class="pricing-toggle-wrap">
          <span style="font-size:14px; font-weight:700; color:#cbd5e1;">Skema Tagihan:</span>
          <button class="btn btn-primary" id="btn-price-monthly" style="font-size:12px; padding:6px 16px;">Bulanan</button>
          <button class="btn btn-secondary" id="btn-price-yearly" style="font-size:12px; padding:6px 16px; border-color:rgba(255,255,255,0.2);">
            Tahunan <span class="badge badge-success" style="margin-left:4px;">HEMAT 20%</span>
          </button>
        </div>

        <div class="pricing-grid">
          <!-- Basic Plan -->
          <div class="pricing-card">
            <h3 style="font-size:20px; font-weight:800; color:white;">Paket Sekolah Dasar</h3>
            <p style="font-size:13px; color:#94a3b8; margin-top:4px;">Ideal untuk 1 Sekolah Negeri/Swasta Mandiri</p>
            <div style="margin-top:20px;">
              <span class="price-val" id="price-basic">Rp 299rb</span>
              <span style="font-size:13px; color:#94a3b8;" id="period-basic">/ bulan</span>
            </div>
            <ul class="pricing-features">
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> 1 Lisensi Sekolah</li>
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Hingga 50 Akun Guru</li>
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Bank Soal Unlimited</li>
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Ekspor Word & PDF</li>
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Simulasi Ujian CBT Standard</li>
            </ul>
            <button class="btn btn-secondary btn-order-plan" data-plan="Sekolah Dasar" style="width:100%; margin-top:auto; font-weight:800;">Pilih Paket Ini</button>
          </div>

          <!-- Professional Plan (Popular) -->
          <div class="pricing-card popular">
            <div class="pricing-badge-popular">PALING POPULER</div>
            <h3 style="font-size:20px; font-weight:800; color:white;">Paket Professional MGMP</h3>
            <p style="font-size:13px; color:#c7d2fe; margin-top:4px;">Untuk Sekolah Unggulan & Komunitas MGMP</p>
            <div style="margin-top:20px;">
              <span class="price-val" id="price-pro">Rp 599rb</span>
              <span style="font-size:13px; color:#c7d2fe;" id="period-pro">/ bulan</span>
            </div>
            <ul class="pricing-features" style="color:#e0e7ff;">
              <li><i data-lucide="check" style="color:#34d399; width:16px; height:16px;"></i> 3 Lisensi Sekolah (Multi-Tenant)</li>
              <li><i data-lucide="check" style="color:#34d399; width:16px; height:16px;"></i> Akun Guru & Reviewer Unlimited</li>
              <li><i data-lucide="check" style="color:#34d399; width:16px; height:16px;"></i> App Android Exam Kiosk Mode</li>
              <li><i data-lucide="check" style="color:#34d399; width:16px; height:16px;"></i> Anti-Cheat & Floating Watermark</li>
              <li><i data-lucide="check" style="color:#34d399; width:16px; height:16px;"></i> Analisis Butir Soal ($P$ & $D$)</li>
              <li><i data-lucide="check" style="color:#34d399; width:16px; height:16px;"></i> Dukungan Prioritas WhatsApp</li>
            </ul>
            <button class="btn btn-primary btn-order-plan" data-plan="Professional MGMP" style="width:100%; margin-top:auto; font-weight:800; background:#6366f1;">Berlangganan Sekarang</button>
          </div>

          <!-- Enterprise Plan -->
          <div class="pricing-card">
            <h3 style="font-size:20px; font-weight:800; color:white;">Paket Enterprise Yayasan</h3>
            <p style="font-size:13px; color:#94a3b8; margin-top:4px;">Untuk Jaringan Sekolah Yayasan / Dinas</p>
            <div style="margin-top:20px;">
              <span class="price-val" id="price-enterprise">Rp 1.299rb</span>
              <span style="font-size:13px; color:#94a3b8;" id="period-enterprise">/ bulan</span>
            </div>
            <ul class="pricing-features">
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Unlimited Lisensi Sekolah</li>
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Full Cyber Defense Dashboard</li>
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Custom Domain & Logo Yayasan</li>
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Dedicated Private Server (SLA 99.9%)</li>
              <li><i data-lucide="check" style="color:#10b981; width:16px; height:16px;"></i> Pelatihan On-Site & Integrasi DAPODIK</li>
            </ul>
            <button class="btn btn-secondary btn-order-plan" data-plan="Enterprise Yayasan" style="width:100%; margin-top:auto; font-weight:800;">Hubungi Tim Sales</button>
          </div>
        </div>
      </section>

      <!-- Testimonials Section -->
      <section style="padding:80px 40px; max-width:1280px; margin:0 auto;" id="testimonials">
        <div class="section-header">
          <div class="section-subtitle">TESTIMONIAL PENGGUNA</div>
          <h2 class="section-title">Dipercaya Oleh Ratusan Kepala Sekolah & Guru</h2>
        </div>

        <div class="testimonials-grid">
          <div class="testimonial-card">
            <div style="display:flex; gap:4px; color:#f59e0b;">
              <i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i>
            </div>
            <p style="font-size:14px; color:#cbd5e1; line-height:1.6;">
              "Sejak menggunakan BankSoalPro, proses penyusunan soal PAS dan PAT di sekolah kami menjadi 5 kali lebih cepat. Fitur ekspor ke Word lengkap dengan kop dan kunci jawaban sangat membantu guru."
            </p>
            <div style="display:flex; align-items:center; gap:12px; margin-top:auto;">
              <div style="width:40px; height:40px; border-radius:50%; background:#4f46e5; display:flex; align-items:center; justify-content:center; font-weight:800;">AK</div>
              <div>
                <div style="font-weight:700; color:white; font-size:14px;">Drs. Abu Khoir, M.Kom</div>
                <div style="font-size:11px; color:#94a3b8;">Kepala Sekolah SMA Negeri 4 Kisaran</div>
              </div>
            </div>
          </div>

          <div class="testimonial-card">
            <div style="display:flex; gap:4px; color:#f59e0b;">
              <i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i>
            </div>
            <p style="font-size:14px; color:#cbd5e1; line-height:1.6;">
              "Fitur Android Exam Kiosk dan Watermark Anti-Cheatnya luar biasa tangguh! Siswa tidak bisa lagi melakukan kecurangan atau berpindah tab saat ujian CBT berlangsung."
            </p>
            <div style="display:flex; align-items:center; gap:12px; margin-top:auto;">
              <div style="width:40px; height:40px; border-radius:50%; background:#10b981; display:flex; align-items:center; justify-content:center; font-weight:800;">BS</div>
              <div>
                <div style="font-weight:700; color:white; font-size:14px;">Budi Santoso, S.Pd</div>
                <div style="font-size:11px; color:#94a3b8;">Ketua MGMP Informatika Sumatera Utara</div>
              </div>
            </div>
          </div>

          <div class="testimonial-card">
            <div style="display:flex; gap:4px; color:#f59e0b;">
              <i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i><i data-lucide="star"></i>
            </div>
            <p style="font-size:14px; color:#cbd5e1; line-height:1.6;">
              "Input rumus matematika dengan MathJax LaTeX dan fitur Analisis Butir Soal (Daya Beda & Tingkat Kesulitan) memberikan wawasan ilmiah yang sangat presisi untuk evaluasi belajar."
            </p>
            <div style="display:flex; align-items:center; gap:12px; margin-top:auto;">
              <div style="width:40px; height:40px; border-radius:50%; background:#818cf8; display:flex; align-items:center; justify-content:center; font-weight:800;">SR</div>
              <div>
                <div style="font-weight:700; color:white; font-size:14px;">Siti Rahma, M.Pd</div>
                <div style="font-size:11px; color:#94a3b8;">Guru Matematika SMAN 1 Medan</div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <!-- FAQ Section -->
      <section style="padding:80px 40px; background:rgba(15,23,42,0.4);" id="faq">
        <div class="section-header">
          <div class="section-subtitle">PERTANYAAN POPULER</div>
          <h2 class="section-title">Pertanyaan Sering Diajukan (FAQ)</h2>
        </div>

        <div class="faq-accordion">
          <div class="faq-item open">
            <div class="faq-question">
              <span>Apakah BankSoalPro aman dari peretasan dan kejahatan siber?</span>
              <i data-lucide="chevron-down"></i>
            </div>
            <div class="faq-answer">
              Ya, BankSoalPro dilengkapi dengan fitur Cyber Security Defense terpadu: Proteksi Sanitizer Anti-XSS pada editor, Brute-Force Rate Limiter (kunci otomatis jika 5 kali gagal login), enkripsi session, dan verifikasi CAPTCHA visual pada setiap halaman login.
            </div>
          </div>

          <div class="faq-item">
            <div class="faq-question">
              <span>Apakah bisa diunggah dan dijalankan di hosting biasa seperti Hostinger?</span>
              <i data-lucide="chevron-down"></i>
            </div>
            <div class="faq-answer">
              Sangat bisa! BankSoalPro dirancang ringan berbasis HTML5, Vanilla CSS, dan ES Modules Javascript. Anda cukup mengunggah berkas ke Shared Hosting Hostinger tanpa perlu server Node.js khusus.
            </div>
          </div>

          <div class="faq-item">
            <div class="faq-question">
              <span>Bagaimana cara mengimpor soal dalam jumlah banyak?</span>
              <i data-lucide="chevron-down"></i>
            </div>
            <div class="faq-answer">
              Anda dapat mengunduh Template Excel (.xlsx) atau CSV yang telah disediakan di menu Import Soal, mengisinya, dan mengunggahnya. Sistem akan secara otomatis memvalidasi dan menyimpan soal ke bank data.
            </div>
          </div>

          <div class="faq-item">
            <div class="faq-question">
              <span>Apakah aplikasi Android ujian dapat digunakan tanpa internet?</span>
              <i data-lucide="chevron-down"></i>
            </div>
            <div class="faq-answer">
              Aplikasi Android Kiosk Mode mendukung ujian jaringan lokal (LAN/Wi-Fi lokal sekolah) maupun jaringan internet publik dengan efisiensi kuota data yang sangat hemat.
            </div>
          </div>
        </div>
      </section>

      <!-- CTA Subscription Footer Banner -->
      <footer style="padding:80px 40px; text-align:center; background:linear-gradient(180deg, #0b0f19 0%, #1e1b4b 100%); border-top:1px solid rgba(255,255,255,0.1);">
        <h2 style="font-size:36px; font-weight:900; color:white;">Siap Modernisasi Bank Soal Sekolah Anda?</h2>
        <p style="color:#94a3b8; margin-top:12px; max-width:600px; margin-left:auto; margin-right:auto;">
          Bergabunglah sekarang dengan ratusan sekolah lainnya. Dapatkan akses instant ke seluruh fitur premium BankSoalPro.
        </p>

        <div style="margin-top:32px;">
          <a href="#login-card-hero" class="btn btn-primary" style="padding:14px 36px; font-size:16px; font-weight:800; border-radius:12px;">
            <i data-lucide="sparkles"></i>
            <span>Daftar / Login Berlangganan Sekarang</span>
          </a>
        </div>

        <div style="margin-top:60px; font-size:12px; color:#64748b; border-top:1px solid rgba(255,255,255,0.05); padding-top:24px;">
          &copy; 2026 BankSoalPro Inc. Hak Cipta Dilindungi Undang-Undang. Aplikasi Web Bank Soal & CBT Online Indonesia.
        </div>
      </footer>
    </div>
  `;

  refreshIcons();

  // CAPTCHA Refresh Button handler
  document.getElementById('btn-refresh-captcha').addEventListener('click', () => {
    currentCaptchaCode = db.generateCaptcha();
    document.getElementById('captcha-code-mount').innerText = currentCaptchaCode;
  });

  // Login Form Submission handler
  const form = document.getElementById('landing-login-form');
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    const email = document.getElementById('l-email').value.trim();
    const password = document.getElementById('l-pass').value;
    const userCaptcha = document.getElementById('l-captcha-input').value.trim().toUpperCase();

    // Verify CAPTCHA
    if (userCaptcha !== currentCaptchaCode.toUpperCase()) {
      showToast('KODE CAPTCHA SALAH! Silakan masukkan kode visual yang sesuai.', 'error');
      db.log('GUEST', email, 'SECURITY_ALERT', `Gagal login: Kode CAPTCHA salah (${userCaptcha})`);
      // Refresh CAPTCHA code for security
      currentCaptchaCode = db.generateCaptcha();
      document.getElementById('captcha-code-mount').innerText = currentCaptchaCode;
      document.getElementById('l-captcha-input').value = '';
      return;
    }

    attemptLogin(email, password);
  });

  // Quick Demo Login clicks handler
  document.querySelectorAll('.quick-landing-demo').forEach(btn => {
    btn.addEventListener('click', () => {
      const email = btn.dataset.email;
      const user = db.get('users').find(u => u.email === email);
      if (user) {
        document.getElementById('l-email').value = user.email;
        document.getElementById('l-pass').value = user.password;
        document.getElementById('l-captcha-input').value = currentCaptchaCode;
        showToast(`Mengisi kredensial & CAPTCHA untuk ${user.name}...`, 'info');
      }
    });
  });

  // Handle Login Tabs
  const tabLogin = document.getElementById('tab-login');
  const tabExam = document.getElementById('tab-exam');
  const paneLogin = document.getElementById('pane-login');
  const paneExam = document.getElementById('pane-exam');

  if (tabLogin && tabExam) {
    tabLogin.addEventListener('click', () => {
      tabLogin.className = 'tab-btn active';
      tabLogin.style.borderBottom = '2px solid var(--primary)';
      tabLogin.style.color = 'white';
      tabExam.className = 'tab-btn';
      tabExam.style.borderBottom = '2px solid transparent';
      tabExam.style.color = '#cbd5e1';
      paneLogin.style.display = 'block';
      paneExam.style.display = 'none';
    });

    tabExam.addEventListener('click', () => {
      tabExam.className = 'tab-btn active';
      tabExam.style.borderBottom = '2px solid var(--primary)';
      tabExam.style.color = 'white';
      tabLogin.className = 'tab-btn';
      tabLogin.style.borderBottom = '2px solid transparent';
      tabLogin.style.color = '#cbd5e1';
      paneLogin.style.display = 'none';
      paneExam.style.display = 'block';
    });
  }

  // Handle Student Exam Entry Flow
  // State vars: 0=enter code, 1=enter token, 2=enter login
  let targetPkg = null;
  let examStep = 0; // step tracker
  const sForm = document.getElementById('student-exam-form');
  if (sForm) {
    sForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const code = document.getElementById('s-exam-code').value.trim().toUpperCase();

      // --- STEP 0: Look up exam code ---
      if (examStep === 0) {
        targetPkg = db.get('packages').find(p => (p.examCode || '').toUpperCase() === code);
        if (!targetPkg) {
          showToast('Kode Paket Ujian tidak ditemukan! Pastikan kode sudah benar.', 'error');
          return;
        }

        // If exam has a token, advance to token step
        if (targetPkg.examToken) {
          examStep = 1;
          const tokenGroup = document.getElementById('token-group');
          tokenGroup.style.display = 'block';
          document.getElementById('s-exam-token').required = true;
          document.getElementById('btn-student-flow').textContent = 'Verifikasi Token →';
          showToast(`Paket "${targetPkg.name}" ditemukan. Masukkan Token dari Guru.`, 'info');
        } else {
          // No token needed, jump straight to login step
          examStep = 2;
          const loginGroup = document.getElementById('student-login-group');
          loginGroup.style.display = 'flex';
          document.getElementById('s-exam-code').disabled = true;
          document.getElementById('s-email').required = true;
          document.getElementById('s-pass').required = true;
          document.getElementById('btn-student-flow').innerHTML = '<i data-lucide="log-in"></i> Verifikasi & Mulai Ujian CBT';
          lucide.createIcons();
          showToast(`Paket "${targetPkg.name}" valid. Silakan login akun Siswa Anda.`, 'success');
        }
        return;
      }

      // --- STEP 1: Validate token ---
      if (examStep === 1) {
        const tokenInput = document.getElementById('s-exam-token').value.trim().toUpperCase();
        const correctToken = (targetPkg.examToken || '').toUpperCase();
        if (tokenInput !== correctToken) {
          showToast('Token Ujian SALAH! Minta token yang benar dari Guru pengawas.', 'error');
          return;
        }
        // Token correct, advance to login
        examStep = 2;
        const loginGroup = document.getElementById('student-login-group');
        loginGroup.style.display = 'flex';
        document.getElementById('s-exam-code').disabled = true;
        document.getElementById('s-exam-token').disabled = true;
        document.getElementById('s-email').required = true;
        document.getElementById('s-pass').required = true;
        document.getElementById('btn-student-flow').innerHTML = '<i data-lucide="log-in"></i> Verifikasi & Mulai Ujian CBT';
        lucide.createIcons();
        showToast('Token valid! Silakan login akun Siswa Anda.', 'success');
        return;
      }

      // --- STEP 2: Student login and redirect to exam ---
      if (examStep === 2) {
        const email = document.getElementById('s-email').value.trim();
        const pass = document.getElementById('s-pass').value;
        const user = db.get('users').find(u => u.email === email && u.password === pass);

        if (!user) {
          showToast('Email atau password Siswa salah!', 'error');
          return;
        }
        if (user.role !== 'SISWA') {
          showToast('Hanya akun Siswa yang dapat memulai ujian melalui portal ini!', 'error');
          return;
        }

        attemptLogin(email, pass, `#cbt?id=${targetPkg.id}`);
      }
    });
  }

  // Pricing Toggle Monthly / Yearly handler
  const btnMonthly = document.getElementById('btn-price-monthly');
  const btnYearly = document.getElementById('btn-price-yearly');

  btnMonthly.addEventListener('click', () => {
    btnMonthly.className = 'btn btn-primary';
    btnYearly.className = 'btn btn-secondary';
    document.getElementById('price-basic').innerText = 'Rp 299rb';
    document.getElementById('period-basic').innerText = '/ bulan';
    document.getElementById('price-pro').innerText = 'Rp 599rb';
    document.getElementById('period-pro').innerText = '/ bulan';
    document.getElementById('price-enterprise').innerText = 'Rp 1.299rb';
    document.getElementById('period-enterprise').innerText = '/ bulan';
  });

  btnYearly.addEventListener('click', () => {
    btnYearly.className = 'btn btn-primary';
    btnMonthly.className = 'btn btn-secondary';
    document.getElementById('price-basic').innerText = 'Rp 2.870rb';
    document.getElementById('period-basic').innerText = '/ tahun';
    document.getElementById('price-pro').innerText = 'Rp 5.750rb';
    document.getElementById('period-pro').innerText = '/ tahun';
    document.getElementById('price-enterprise').innerText = 'Rp 12.470rb';
    document.getElementById('period-enterprise').innerText = '/ tahun';
  });

  // Order Plan Modal buttons handler
  document.querySelectorAll('.btn-order-plan').forEach(btn => {
    btn.addEventListener('click', () => {
      const planName = btn.dataset.plan;
      const html = `
        <div class="modal-header">
          <h3 class="modal-title">Formulir Berlangganan - ${planName}</h3>
          <button class="modal-close" data-close-modal><i data-lucide="x"></i></button>
        </div>
        <form id="form-subscribe-plan">
          <div class="modal-body" style="display:flex; flex-direction:column; gap:16px;">
            <p style="font-size:13px; color:var(--neutral-600);">Isi data instansi Anda untuk mengaktifkan lisensi paket <strong>${planName}</strong>.</p>
            <div class="form-group">
              <label class="form-label" for="sub-school-name">Nama Sekolah / Lembaga</label>
              <input type="text" id="sub-school-name" class="form-input" placeholder="contoh: SMAN 1 Medan" required>
            </div>
            <div class="form-group">
              <label class="form-label" for="sub-pic-name">Nama Penanggung Jawab / Kepala Sekolah</label>
              <input type="text" id="sub-pic-name" class="form-input" placeholder="contoh: Budi Santoso, M.Pd" required>
            </div>
            <div class="form-group">
              <label class="form-label" for="sub-pic-phone">Nomor WhatsApp Aktif</label>
              <input type="tel" id="sub-pic-phone" class="form-input" placeholder="081234567890" required>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-close-modal>Batal</button>
            <button type="submit" class="btn btn-primary">Kirim Permohonan Langganan</button>
          </div>
        </form>
      `;

      openModal(html);

      document.getElementById('form-subscribe-plan').addEventListener('submit', (e) => {
        e.preventDefault();
        const schoolName = document.getElementById('sub-school-name').value;
        showToast(`Permohonan langganan paket "${planName}" untuk ${schoolName} berhasil dikirim! Tim kami akan menghubungi Anda via WhatsApp.`, 'success');
        closeModal();
      });
    });
  });

  // FAQ Accordion click handlers
  document.querySelectorAll('.faq-question').forEach(q => {
    q.addEventListener('click', () => {
      const item = q.parentElement;
      item.classList.toggle('open');
    });
  });
}

function attemptLogin(email, password, targetRedirect = '#dashboard') {
  const user = db.get('users').find(u => u.email === email && u.password === password);

  if (user) {
    if (user.status !== 'AKTIF') {
      showToast('Akun Anda dinonaktifkan. Hubungi admin.', 'error');
      return;
    }

    sessionStorage.setItem('active_user', JSON.stringify(user));
    db.log(user.id, user.name, 'LOGIN', 'Berhasil login ke dalam sistem.');
    showToast(`Selamat datang kembali, ${user.name}!`, 'success');
    window.location.hash = targetRedirect;
  } else {
    showToast('Email atau password salah.', 'error');
  }
}

// ----------------------------------------------------
// 2. RENDER APP SHELL (Sidebar & Header Layout)
// ----------------------------------------------------
function renderAppShell(activePage, queryParams) {
  const activeSchoolId = sessionStorage.getItem('active_school_id') || activeUser.schoolId || 'sch-1';
  const allSchools = db.get('schools');
  const currentSchool = allSchools.find(s => s.id === activeSchoolId) || allSchools[0] || { name: 'BankSoalPro', academicYear: '-' };

  const menuConfig = [
    { page: '#dashboard', label: 'Dashboard', icon: 'layout-dashboard', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER', 'SISWA'] },
    { page: '#sekolah', label: 'Kelola Sekolah', icon: 'building-2', roles: ['SUPER_ADMIN'] },
    { page: '#profil-sekolah', label: 'Profil Sekolah', icon: 'school', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH'] },
    { page: '#guru', label: 'Manajemen Guru', icon: 'users', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH'] },
    { page: '#kelola-user', label: 'Kelola Pengguna', icon: 'user-cog', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH'] },
    { page: '#kelola-siswa', label: 'Kelola Siswa', icon: 'graduation-cap', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH'] },
    { page: '#privileges', label: 'Hak Akses & Privilege', icon: 'key-round', roles: ['SUPER_ADMIN'] },
    { page: '#struktur', label: 'Struktur Kurikulum', icon: 'git-branch', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER'] },
    { page: '#soal', label: 'Bank Soal', icon: 'file-text', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER'] },
    { page: '#impor', label: 'Import Soal', icon: 'file-up', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU'] },
    { page: '#paket', label: 'Paket Soal', icon: 'package', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER'] },
    { page: '#android-exam', label: 'App Ujian Android', icon: 'smartphone', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'SISWA'] },
    { page: '#rekap-ujian', label: 'Hasil & Rekap Ujian', icon: 'award', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'SISWA'] },
    { page: '#logs', label: 'Audit Log & Keamanan', icon: 'shield-alert', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH'] },
    { page: '#profil-saya', label: 'Profil Saya', icon: 'circle-user-round', roles: ['SUPER_ADMIN', 'ADMIN_SEKOLAH', 'GURU', 'REVIEWER', 'SISWA'] },
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
  const notifications = db.getNotifications(activeUser.id);
  const unreadCount = notifications.filter(n => !n.isRead).length;

  let schoolHeaderHtml = `
    <div class="school-badge">
      <i data-lucide="school" style="width:16px; height:16px;"></i>
      <span>${currentSchool.name} (${currentSchool.academicYear})</span>
    </div>
  `;

  if (activeUser.role === 'SUPER_ADMIN') {
    schoolHeaderHtml = `
      <div style="display:flex; align-items:center; gap:8px;">
        <span style="font-size:12px; font-weight:700; color:var(--neutral-700);">Konteks Sekolah:</span>
        <select id="super-admin-school-select" class="school-switcher-select">
          ${allSchools.map(s => `<option value="${s.id}" ${s.id === currentSchool.id ? 'selected' : ''}>${s.name}</option>`).join('')}
        </select>
      </div>
    `;
  }

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
          
          <div class="header-actions" style="position:relative; display:flex; align-items:center; gap:16px;">
            ${schoolHeaderHtml}
            
            <button class="btn btn-secondary btn-icon" id="btn-notif-toggle" style="position:relative;" title="Notifikasi System">
              <i data-lucide="bell" style="width:18px; height:18px;"></i>
              ${unreadCount > 0 ? `<span style="position:absolute; top:-2px; right:-2px; background:var(--danger-text); color:white; font-size:10px; font-weight:800; padding:2px 5px; border-radius:10px;">${unreadCount}</span>` : ''}
            </button>

            <!-- Notifications Popover -->
            <div class="notif-popover" id="notif-popover">
              <div class="notif-header">
                <span>Pemberitahuan System</span>
                <span style="font-size:11px; color:var(--primary);">${unreadCount} Baru</span>
              </div>
              <div style="max-height:260px; overflow-y:auto;">
                ${notifications.map(n => `
                  <div class="notif-item ${n.isRead ? '' : 'unread'}">
                    <strong style="color:var(--neutral-900);">${n.title}</strong>
                    <span style="color:var(--neutral-700);">${n.message}</span>
                    <span style="font-size:10px; color:var(--neutral-500); margin-top:2px;">${new Date(n.timestamp).toLocaleString('id-ID')}</span>
                  </div>
                `).join('') || '<div style="padding:16px; text-align:center; color:var(--neutral-600);">Tidak ada notifikasi.</div>'}
              </div>
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

  // Bind Super Admin School Switcher
  const schoolSelect = document.getElementById('super-admin-school-select');
  if (schoolSelect) {
    schoolSelect.addEventListener('change', (e) => {
      sessionStorage.setItem('active_school_id', e.target.value);
      showToast(`Konteks sekolah diubah: ${e.target.options[e.target.selectedIndex].text}`, 'success');
      handleRouting();
    });
  }

  // Bind Notif Toggle Popover
  const notifBtn = document.getElementById('btn-notif-toggle');
  const notifPopover = document.getElementById('notif-popover');
  if (notifBtn && notifPopover) {
    notifBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      notifPopover.classList.toggle('open');
    });
    document.addEventListener('click', () => notifPopover.classList.remove('open'));
  }

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
    case '#sekolah':
      titleDisplay.innerText = 'Kelola Sekolah Multi-Tenant';
      subtitleDisplay.innerText = 'Super Admin: Tambah, ubah, dan navigasi data sekolah terdaftar';
      renderKelolaSekolah(contentMount);
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
    case '#privileges':
      titleDisplay.innerText = 'Hak Akses & Privilege Peran';
      subtitleDisplay.innerText = 'Super Admin: Atur privilege fitur untuk setiap jenis akun';
      renderPrivileges(contentMount);
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
      subtitleDisplay.innerText = 'Penyusunan paket ujian otomatis/acak, QR Code, dan ekspor berkas';
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
      titleDisplay.innerText = 'Audit Log & Pertahanan Keamanan';
      subtitleDisplay.innerText = 'Riwayat log aktivitas, pertahanan cyber crime, dan backup/restore basis data';
      renderLogs(contentMount);
      break;
    case '#profil-saya':
      titleDisplay.innerText = 'Profil Akun Saya';
      subtitleDisplay.innerText = 'Lihat informasi akun, sekolah, dan mata pelajaran yang diampu';
      renderProfilSaya(contentMount);
      break;
    case '#kelola-user':
      titleDisplay.innerText = 'Kelola Semua Pengguna';
      subtitleDisplay.innerText = 'CRUD akun admin, guru, reviewer, dan siswa';
      renderKelolaUser(contentMount);
      break;
    case '#kelola-siswa':
      titleDisplay.innerText = 'Kelola Data Siswa';
      subtitleDisplay.innerText = 'CRUD data siswa, impor dari Excel/CSV, dan ekspor data';
      renderKelolaSiswa(contentMount);
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

  let filterScope = 'SEMUA';
  let filterKeyword = '';
  let filterSubject = '';
  let filterType = '';
  let filterStatus = '';

  mount.innerHTML = `
    <!-- Scope filter bar -->
    <div style="display:flex; gap:8px; margin-bottom:16px; border-bottom:1px solid var(--neutral-400); padding-bottom:10px;">
      <button class="btn btn-primary scope-tab" data-scope="SEMUA">Semua Soal</button>
      <button class="btn btn-secondary scope-tab" data-scope="SAYA">Soal Karya Saya</button>
      <button class="btn btn-secondary scope-tab" data-scope="GURU_LAIN">Soal Guru Lain (Sama Mapel/Kelas)</button>
      <button class="btn btn-secondary scope-tab" data-scope="FAVORIT">Soal Favorit Saya</button>
    </div>

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
  const scopeTabs = mount.querySelectorAll('.scope-tab');

  scopeTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      scopeTabs.forEach(t => t.className = 'btn btn-secondary scope-tab');
      tab.className = 'btn btn-primary scope-tab';
      filterScope = tab.dataset.scope;
      reloadList();
    });
  });

  const reloadList = () => {
    let list = db.get('questions');

    // Apply Subject Filter for Teachers
    if (activeUser.role === 'GURU' && activeUser.subjectIds && activeUser.subjectIds.length > 0) {
      list = list.filter(q => activeUser.subjectIds.includes(q.subjectId));
    }

    // Apply Scope Filter
    if (filterScope === 'SAYA') {
      list = list.filter(q => q.creatorId === activeUser.id);
    } else if (filterScope === 'GURU_LAIN') {
      list = list.filter(q => q.creatorId !== activeUser.id);
    } else if (filterScope === 'FAVORIT') {
      list = list.filter(q => db.isBookmarked(activeUser.id, q.id));
    }

    // Apply filters
    if (filterKeyword) {
      const kw = filterKeyword.toLowerCase();
      list = list.filter(q =>
        (q.code || '').toLowerCase().includes(kw) ||
        (q.questionText || '').toLowerCase().includes(kw) ||
        (q.tag || '').toLowerCase().includes(kw)
      );
    }

    if (filterSubject) list = list.filter(q => q.subjectId === filterSubject);
    if (filterType) list = list.filter(q => q.type === filterType);
    if (filterStatus) list = list.filter(q => q.status === filterStatus);

    listMount.innerHTML = list.map(q => {
      const mapel = subjects.find(s => s.id === q.subjectId)?.name || 'Umum';
      const kls = classes.find(c => c.id === q.classId)?.name || 'X';
      const bab = chapters.find(c => c.id === q.chapterId)?.name || 'Umum';

      const isBookmarked = db.isBookmarked(activeUser.id, q.id);
      const starIcon = isBookmarked ? 'star-off' : 'star';
      const starClass = isBookmarked ? 'color: #ecc94b; fill: #ecc94b;' : '';

      const isMyQuestion = q.creatorId === activeUser.id || activeUser.role === 'SUPER_ADMIN' || activeUser.role === 'ADMIN_SEKOLAH';
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

      // Action buttons according to ownership
      let editButtons = '';
      if (isMyQuestion) {
        editButtons = `
          <a class="btn btn-secondary btn-icon" style="padding:6px;" href="#soal?action=edit&id=${q.id}" title="Edit Soal Saya">
            <i data-lucide="edit" style="width:16px; height:16px;"></i>
          </a>
          <button class="btn btn-danger btn-icon delete-soal" data-id="${q.id}" style="padding:6px;" title="Hapus Soal Saya">
            <i data-lucide="trash-2" style="width:16px; height:16px;"></i>
          </button>
        `;
      } else {
        editButtons = `
          <button class="btn btn-primary adopt-soal" style="padding:6px 12px; font-size:11px;" data-id="${q.id}" title="Ambil / Salin Soal ke Bank Saya">
            <i data-lucide="copy" style="width:14px; height:14px;"></i> Salin ke Bank Saya
          </button>
        `;
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
      } else if (activeUser.role === 'GURU' && q.status === 'DRAFT' && isMyQuestion) {
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
              <button class="btn btn-secondary btn-icon btn-item-analysis" data-id="${q.id}" style="padding:6px;" title="Analisis Butir Soal">
                <i data-lucide="bar-chart-2" style="width:16px; height:16px; color:var(--primary);"></i>
              </button>
              ${editButtons}
            </div>
          </div>
          
          <div class="q-body-preview">${q.questionText}</div>
          
          ${choicesHtml}

          <!-- Footer/Aktivitas -->
          <div style="display:flex; justify-content:space-between; align-items:center; border-top:1px solid var(--neutral-400); padding-top:12px; margin-top:8px; font-size:11px; color: var(--neutral-600);">
            <div>Dibuat oleh: <strong>${q.creatorName || 'Guru'}</strong> ${q.adoptedFrom ? `(Disalin dari ${q.adoptedFrom})` : ''} | Bab: ${bab}</div>
            <div style="display:inline-flex; gap:8px;">
              ${reviewActions}
            </div>
          </div>
        </div>
      `;
    }).join('') || '<div class="card" style="padding:40px; text-align:center; color: var(--neutral-600);">Tidak ada soal yang cocok dengan kriteria pencarian.</div>';

    lucide.createIcons();

    if (window.MathJax && window.MathJax.typesetPromise) {
      window.MathJax.typesetPromise([listMount]).catch(err => console.error(err));
    }

    // Event Bindings
    listMount.querySelectorAll('.toggle-fav').forEach(btn => {
      btn.addEventListener('click', () => {
        const isB = db.toggleBookmark(activeUser.id, btn.dataset.id);
        showToast(isB ? 'Soal ditambahkan ke favorit!' : 'Bookmark dihapus.', isB ? 'success' : 'info');
        reloadList();
      });
    });

    listMount.querySelectorAll('.btn-item-analysis').forEach(btn => {
      btn.addEventListener('click', () => openItemAnalysisModal(btn.dataset.id));
    });

    listMount.querySelectorAll('.adopt-soal').forEach(btn => {
      btn.addEventListener('click', () => {
        const cloned = db.adoptQuestion(btn.dataset.id, activeUser);
        if (cloned) {
          showToast(`Berhasil menyalin soal ${cloned.code} ke Bank Soal Anda!`, 'success');
          reloadList();
        }
      });
    });

    listMount.querySelectorAll('.delete-soal').forEach(btn => {
      btn.addEventListener('click', () => {
        if (confirm('Apakah Anda yakin ingin menghapus soal milik Anda ini?')) {
          db.delete('questions', btn.dataset.id);
          showToast('Soal berhasil dihapus.', 'success');
          reloadList();
        }
      });
    });

    listMount.querySelectorAll('.submit-review').forEach(btn => {
      btn.addEventListener('click', () => {
        db.update('questions', btn.dataset.id, { status: 'REVIEW' });
        showToast('Soal diajukan ke Reviewer.', 'success');
        reloadList();
      });
    });

    listMount.querySelectorAll('.approve-soal').forEach(btn => {
      btn.addEventListener('click', () => {
        db.update('questions', btn.dataset.id, { status: 'APPROVED', reviewerId: activeUser.id, reviewerName: activeUser.name });
        showToast('Soal berhasil disetujui (APPROVED).', 'success');
        reloadList();
      });
    });

    listMount.querySelectorAll('.reject-soal').forEach(btn => {
      btn.addEventListener('click', () => {
        db.update('questions', btn.dataset.id, { status: 'REJECTED', reviewerId: activeUser.id, reviewerName: activeUser.name });
        showToast('Soal ditolak (REJECTED).', 'error');
        reloadList();
      });
    });
  };

  // Bind Filter Input listeners
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

  reloadList();
}


// ----------------------------------------------------
// 8. PAGE: CREATE/EDIT QUESTION FORM
// ----------------------------------------------------
function renderFormSoal(mount, questionId = null) {
  const isEdit = !!questionId;
  const q = isEdit ? db.get('questions').find(soal => soal.id === questionId) : {};

  // Authorization check for Edit
  if (isEdit && activeUser.role === 'GURU' && q.creatorId !== activeUser.id) {
    showToast('Akses Ditolak: Anda hanya dapat mengedit soal buatan Anda sendiri.', 'error');
    window.location.hash = '#soal';
    return;
  }

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
                <label class="form-label" for="q-code">Kode Soal (Unik)
                  ${!isEdit ? '<button type="button" id="btn-auto-code" style="float:right;font-size:10px;color:var(--primary);background:none;border:none;cursor:pointer;text-decoration:underline;">🔄 Auto</button>' : ''}
                </label>
                <input type="text" id="q-code" class="form-input" placeholder="INF-X-001" value="${q.code || ''}" required style="font-family:monospace;letter-spacing:1px;">
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
                <button class="toolbar-btn" id="btn-undo" title="Undo (Ctrl+Z)"><i data-lucide="undo-2" style="width:14px; height:14px;"></i></button>
                <button class="toolbar-btn" id="btn-redo" title="Redo (Ctrl+Y)"><i data-lucide="redo-2" style="width:14px; height:14px;"></i></button>
                <span class="toolbar-divider"></span>
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
  
  // Auto-generate Kode Soal
  const btnAutoCode = document.getElementById('btn-auto-code');
  if (btnAutoCode) {
    btnAutoCode.addEventListener('click', () => {
      const sId = subSelect.value;
      if (!sId) {
        showToast('Pilih Mata Pelajaran terlebih dahulu untuk auto-generate kode.', 'error');
        return;
      }
      document.getElementById('q-code').value = db.generateQuestionCode(sId);
      showToast('Kode soal berhasil dibuat otomatis.', 'success');
    });
  }

  // Undo / Redo logic for visual editor
  const editorEl = document.getElementById('q-text');
  let historyStack = [editorEl.value];
  let historyIdx = 0;

  editorEl.addEventListener('input', () => {
    historyStack = historyStack.slice(0, historyIdx + 1);
    historyStack.push(editorEl.value);
    historyIdx++;
    updateLivePreview();
  });

  document.getElementById('btn-undo')?.addEventListener('click', (e) => {
    e.preventDefault();
    if (historyIdx > 0) {
      historyIdx--;
      editorEl.value = historyStack[historyIdx];
      updateLivePreview();
    }
  });

  document.getElementById('btn-redo')?.addEventListener('click', (e) => {
    e.preventDefault();
    if (historyIdx < historyStack.length - 1) {
      historyIdx++;
      editorEl.value = historyStack[historyIdx];
      updateLivePreview();
    }
  });

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
  const canExport = db.can(activeUser.role, 'canExport');

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
                <th>Jenis & Durasi</th>
                <th>Mata Pelajaran</th>
                <th>Kelas</th>
                <th>Kode & QR Ujian</th>
                <th style="text-align:right;">Aksi & Izin Siswa</th>
              </tr>
            </thead>
            <tbody>
              ${packages.map(p => {
    const sub = subjects.find(s => s.id === p.subjectId)?.name || '-';
    const cls = classes.find(c => c.id === p.classId)?.name || '-';
    const examCode = p.examCode || 'EXAM-INF-8921';
    return `
                  <tr>
                    <td>
                      <div class="pkg-name-clickable" data-id="${p.id}" style="font-weight:700; color:var(--primary); cursor:pointer; text-decoration:underline;" title="Klik untuk mengedit detail paket & token">
                        ${p.name}
                      </div>
                      <div style="font-size:11px; color:var(--neutral-600); margin-top:2px;">
                        Disusun oleh: ${p.creatorName || 'Guru'} | Soal: ${(p.questionIds || []).length} butir
                      </div>
                    </td>
                    <td>
                      <span class="badge badge-primary">${p.type}</span>
                      <div style="font-size:11px; font-weight:700; color:var(--neutral-700); margin-top:4px;">
                        <i data-lucide="clock" style="width:12px; height:12px; display:inline-block; vertical-align:middle;"></i> ${p.duration || 60} Menit
                      </div>
                    </td>
                    <td>${sub}</td>
                    <td>Kelas ${cls}</td>
                    <td>
                      <button class="btn btn-secondary btn-show-qr" data-id="${p.id}" style="padding:4px 10px; font-size:11px; font-weight:800; color:var(--primary);">
                        <i data-lucide="qr-code" style="width:14px; height:14px;"></i> ${examCode}
                      </button>
                    </td>
                    <td style="text-align:right;">
                      <div style="display:inline-flex; gap:6px; flex-wrap:wrap; justify-content:flex-end;">
                        <button class="btn btn-secondary btn-icon btn-preview-pkg" data-id="${p.id}" title="Pratinjau Soal">
                          <i data-lucide="eye" style="width:15px; height:15px;"></i>
                        </button>
                        
                        <button class="btn btn-secondary btn-manage-student-perm" data-id="${p.id}" style="padding:6px 10px; font-size:11px;" title="Kelola Persetujuan Izin Siswa">
                          <i data-lucide="user-check" style="width:14px; height:14px;"></i> Izin Siswa
                        </button>

                        ${canExport ? `
                          <button class="btn btn-secondary btn-icon btn-export-doc" data-id="${p.id}" title="Ekspor Word">
                            <i data-lucide="file-text" style="width:15px; height:15px; color:#2b6cb0;"></i>
                          </button>
                          <button class="btn btn-secondary btn-icon btn-export-xls" data-id="${p.id}" title="Ekspor Excel">
                            <i data-lucide="file-spreadsheet" style="width:15px; height:15px; color:#2f855a;"></i>
                          </button>
                          <button class="btn btn-secondary btn-icon btn-export-pdf" data-id="${p.id}" title="Cetak PDF">
                            <i data-lucide="printer" style="width:15px; height:15px; color:#c53030;"></i>
                          </button>
                        ` : ''}

                        <button class="btn btn-danger btn-icon btn-delete-pkg" data-id="${p.id}" title="Hapus Paket">
                          <i data-lucide="trash-2" style="width:15px; height:15px;"></i>
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

  document.getElementById('btn-tambah-paket').addEventListener('click', openCreatePackageModal);

  mount.querySelectorAll('.pkg-name-clickable').forEach(el => {
    el.addEventListener('click', () => openPackageDetailModal(el.dataset.id));
  });

  mount.querySelectorAll('.btn-show-qr').forEach(btn => {
    btn.addEventListener('click', () => openQRCodeModal(btn.dataset.id));
  });

  mount.querySelectorAll('.btn-manage-student-perm').forEach(btn => {
    btn.addEventListener('click', () => openStudentPermissionModal(btn.dataset.id));
  });

  mount.querySelectorAll('.btn-preview-pkg').forEach(btn => {
    btn.addEventListener('click', () => previewPackageQuestions(btn.dataset.id));
  });

  if (canExport) {
    mount.querySelectorAll('.btn-export-doc').forEach(btn => {
      btn.addEventListener('click', () => triggerExport(btn.dataset.id, 'doc'));
    });

    mount.querySelectorAll('.btn-export-xls').forEach(btn => {
      btn.addEventListener('click', () => triggerExport(btn.dataset.id, 'xls'));
    });

    mount.querySelectorAll('.btn-export-pdf').forEach(btn => {
      btn.addEventListener('click', () => triggerExport(btn.dataset.id, 'pdf'));
    });
  }

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

function openPackageDetailModal(pkgId) {
  const p = db.get('packages').find(pkg => pkg.id === pkgId);
  if (!p) return;
  const subjects = db.get('subjects');
  const classes = db.get('classes');
  const sub = subjects.find(s => s.id === p.subjectId)?.name || '-';
  const cls = classes.find(c => c.id === p.classId)?.name || '-';

  modalContent.innerHTML = `
    <div style="padding:24px; max-width:500px; width:100%;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h3 style="font-size:18px; font-weight:800; color:var(--neutral-900);">Detail Paket Soal</h3>
        <button class="btn btn-icon btn-secondary btn-close-modal"><i data-lucide="x"></i></button>
      </div>

      <div style="background:var(--neutral-100); padding:16px; border-radius:8px; margin-bottom:16px;">
        <div style="margin-bottom:8px;"><strong>Nama Paket:</strong> ${p.name}</div>
        <div style="margin-bottom:8px;"><strong>Mata Pelajaran:</strong> ${sub}</div>
        <div style="margin-bottom:8px;"><strong>Kelas:</strong> ${cls}</div>
        <div style="margin-bottom:8px;"><strong>Kode Paket Ujian:</strong> <span style="font-family:monospace; font-weight:bold; color:var(--primary);">${p.examCode || '-'}</span></div>
      </div>

      <div class="form-group">
        <label class="form-label">Token Ujian (Akses Siswa)</label>
        <input type="text" id="edit-pkg-token" class="form-input" value="${p.examToken || ''}" style="text-transform:uppercase; font-weight:bold; letter-spacing:1px;" placeholder="Contoh: TOKEN123">
        <p style="font-size:11px; color:var(--neutral-500); margin-top:4px;">Token ini harus diberikan kepada siswa agar dapat memulai ujian. Kosongkan jika tidak perlu token.</p>
      </div>

      <div style="display:flex; gap:12px; margin-top:24px; justify-content:flex-end;">
        <button class="btn btn-secondary btn-close-modal">Batal</button>
        <button class="btn btn-primary" id="btn-save-pkg-detail">Simpan Perubahan</button>
      </div>
    </div>
  `;

  globalModal.style.display = 'flex';
  lucide.createIcons();

  modalContent.querySelectorAll('.btn-close-modal').forEach(btn => {
    btn.onclick = () => { globalModal.style.display = 'none'; };
  });

  document.getElementById('btn-save-pkg-detail').onclick = () => {
    const newToken = document.getElementById('edit-pkg-token').value.trim().toUpperCase();
    p.examToken = newToken;
    db.update('packages', p.id, p);
    showToast('Token ujian berhasil diperbarui.', 'success');
    globalModal.style.display = 'none';
    handleRouting();
  };
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
        
        <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:16px;">
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
            <label class="form-label" for="p-duration">Durasi Ujian (Menit)</label>
            <input type="number" min="5" max="300" id="p-duration" class="form-input" value="60" required>
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

        <!-- Shuffle Settings -->
        <div style="border-top:1px dashed var(--neutral-400); padding-top:12px;">
          <h4 style="font-size:13px; font-weight:700; margin-bottom:8px;">Pengaturan Pengacakan Ujian</h4>
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
            <label style="display:flex; align-items:center; gap:8px; font-size:13px; font-weight:600; cursor:pointer;">
              <input type="checkbox" id="p-rand-q" checked style="width:16px; height:16px;">
              <span>Acak Urutan Soal Ujian</span>
            </label>
            <label style="display:flex; align-items:center; gap:8px; font-size:13px; font-weight:600; cursor:pointer;">
              <input type="checkbox" id="p-rand-c" checked style="width:16px; height:16px;">
              <span>Acak Urutan Pilihan Jawaban (A, B, C, D)</span>
            </label>
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

        <!-- Token Ujian -->
        <div style="border-top:1px dashed var(--neutral-400); padding-top:16px; margin-top:8px;">
          <h4 style="font-size:13px; font-weight:700; margin-bottom:4px; display:flex; align-items:center; gap:8px;">
            <i data-lucide="lock-keyhole" style="width:15px;height:15px;color:var(--primary)"></i>
            Token Akses Ujian (Wajib Diisi Siswa Sebelum Mulai)
          </h4>
          <p style="font-size:11px; color:var(--neutral-600); margin-bottom:12px;">
            Siswa wajib memasukkan kode token ini setelah memasukkan kode paket ujian agar dapat mulai mengerjakan.
          </p>
          <div style="display:flex; gap:10px; align-items:flex-end;">
            <div class="form-group" style="flex:1; margin:0;">
              <label class="form-label" for="p-token">Kode Token Ujian</label>
              <input type="text" id="p-token" class="form-input" placeholder="contoh: INF001" maxlength="12"
                style="text-transform:uppercase; font-family:monospace; font-weight:700; letter-spacing:2px; font-size:16px;">
            </div>
            <button type="button" id="btn-gen-token" class="btn btn-secondary" style="padding:10px 14px; white-space:nowrap;">
              <i data-lucide="shuffle" style="width:14px;height:14px;"></i> Auto Generate
            </button>
          </div>
          <p style="font-size:11px; color:var(--neutral-500); margin-top:6px;">
            <i data-lucide="info" style="width:11px;height:11px;display:inline;"></i>
            Kosongkan jika ujian tidak memerlukan token (siswa bisa langsung masuk).
          </p>
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

  // Auto generate token button
  document.getElementById('btn-gen-token').addEventListener('click', () => {
    const token = db.generateExamToken();
    document.getElementById('p-token').value = token;
    showToast(`Token ujian di-generate: ${token}`, 'info');
  });

  // Form submit handler
  document.getElementById('pkg-create-form').addEventListener('submit', (e) => {
    e.preventDefault();

    const name = document.getElementById('p-name').value;
    const type = document.getElementById('p-type').value;
    const duration = parseInt(document.getElementById('p-duration').value) || 60;
    const mode = modeSelect.value;
    const subjectId = subSelect.value;
    const classId = clsSelect.value;
    const randomizeQuestions = document.getElementById('p-rand-q').checked;
    const randomizeChoices = document.getElementById('p-rand-c').checked;
    const examToken = document.getElementById('p-token').value.trim().toUpperCase();

    let questionIds = [];

    if (mode === 'MANUAL') {
      questionIds = Array.from(document.querySelectorAll('.pkg-q-check:checked')).map(chk => chk.value);
      if (questionIds.length === 0) {
        showToast('Pilih minimal 1 butir soal untuk menyusun paket.', 'error');
        return;
      }
    } else {
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

      const shuffled = [...sourceList].sort(() => 0.5 - Math.random());
      questionIds = shuffled.slice(0, numQ).map(q => q.id);
      showToast(`Generator mengacak ${questionIds.length} soal yang cocok.`, 'success');
    }

    const enableAudio = document.getElementById('p-enable-audio').checked;
    const enableAnim = document.getElementById('p-enable-anim').checked;

    const examCode = `EXAM-${type}-${Math.floor(1000 + Math.random() * 9000)}`;

    db.insert('packages', {
      name,
      type,
      duration,
      examCode,
      examToken,
      qrData: JSON.stringify({ examCode, duration, packageName: name }),
      randomizeQuestions,
      randomizeChoices,
      schoolId: activeUser.schoolId || 'sch-1',
      subjectId,
      classId,
      creatorId: activeUser.id,
      creatorName: activeUser.name,
      questionIds,
      enableAudio,
      enableAnim,
      createdAt: new Date().toISOString()
    });

    showToast('Paket soal berhasil dibuat dengan Kode Ujian & QR Code.', 'success');
    closeModal();
    renderPaketSoal(document.getElementById('shell-content-mount'));
  });
}

// ----------------------------------------------------
// 11. PAGE: AUDIT LOGS & BACKUP/RESTORE DATABASE
// ----------------------------------------------------
function renderLogs(mount) {
  const logs = db.get('logs') || [];

  mount.innerHTML = `
    <!-- Cyber Security Defense Header Dashboard -->
    <div class="cyber-dashboard">
      <div class="cyber-card">
        <div style="display:flex; justify-content:space-between; align-items:center;">
          <h4 style="font-size:14px; font-weight:800; color:#38bdf8;">Proteksi Anti-XSS Sanitizer</h4>
          <span class="cyber-status-pill"><i data-lucide="shield-check" style="width:12px; height:12px;"></i> AKTIF</span>
        </div>
        <p style="font-size:12px; color:#94a3b8; margin-top:8px; line-height:1.4;">
          Pembersihan otomatis payload skrip jahat (&lt;script&gt;, onerror, javascript:) pada WYSIWYG editor.
        </p>
      </div>

      <div class="cyber-card">
        <div style="display:flex; justify-content:space-between; align-items:center;">
          <h4 style="font-size:14px; font-weight:800; color:#38bdf8;">Brute-Force Rate Limiter</h4>
          <span class="cyber-status-pill"><i data-lucide="lock" style="width:12px; height:12px;"></i> AKTIF</span>
        </div>
        <p style="font-size:12px; color:#94a3b8; margin-top:8px; line-height:1.4;">
          Pemblokiran otomatis 15 menit apabila terjadi 5 kali kegagalan kata sandi beruntun.
        </p>
      </div>

      <div class="cyber-card">
        <div style="display:flex; justify-content:space-between; align-items:center;">
          <h4 style="font-size:14px; font-weight:800; color:#38bdf8;">Lockdown & Anti-Cheat CBT</h4>
          <span class="cyber-status-pill"><i data-lucide="eye" style="width:12px; height:12px;"></i> MONITORING</span>
        </div>
        <p style="font-size:12px; color:#94a3b8; margin-top:8px; line-height:1.4;">
          Floating Watermark nama/IP siswa, pemblokiran shortcut & deteksi otomatis tab blur.
        </p>
      </div>
    </div>

    <div style="display:grid; grid-template-columns: 2fr 1fr; gap:28px;">
      <!-- Logs table card -->
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Riwayat Audit Log & Aktivitas Pengguna</h3>
        </div>
        <div class="card-body" style="padding:0; max-height:70vh; overflow-y:auto;">
          <div class="table-wrapper">
            <table class="table" style="font-size:12px;">
              <thead>
                <tr>
                  <th>Waktu</th>
                  <th>Nama Pengguna</th>
                  <th>Aksi</th>
                  <th>Detail Aktivitas & IP</th>
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
                      <td style="color:var(--neutral-700);">${l.details} <span style="font-size:10px; color:var(--neutral-500);">(IP: 192.168.1.104)</span></td>
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

  const backupBtn = mount.querySelector('#btn-backup-db');
  const restoreZone = mount.querySelector('#restore-dropzone');
  const restoreInput = mount.querySelector('#restore-file-input');
  const resetBtn = mount.querySelector('#btn-reset-db');

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
                        <a href="#home" class="btn btn-primary" style="padding: 8px 16px; font-size: 13px;" title="Gunakan Portal Ujian Siswa untuk memasukkan Kode &amp; Token">
                          <i data-lucide="qr-code" style="width:16px; height:16px;"></i>
                          <span>Ikuti Ujian</span>
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
    <!-- Anti-Cheat Interstitial Overlay -->
    <div id="cbt-anti-cheat-overlay" style="position:fixed; top:0; left:0; width:100%; height:100%; background:var(--neutral-900); z-index:9999; display:flex; flex-direction:column; align-items:center; justify-content:center; color:white; font-family:'Plus Jakarta Sans', sans-serif;">
      <i data-lucide="shield-alert" style="width:64px; height:64px; color:var(--danger); margin-bottom:24px;"></i>
      <h2 style="font-size:28px; font-weight:800; margin-bottom:12px; text-align:center;">PERSIAPAN UJIAN CBT KIOSK</h2>
      <p style="font-size:16px; color:var(--neutral-400); max-width:600px; text-align:center; line-height:1.6; margin-bottom:32px;">
        Ujian ini dilindungi oleh Sistem Keamanan Anti-Cheat. Segala bentuk kecurangan (membuka tab baru, split screen, alt-tab, copy-paste) akan <strong>langsung membatalkan ujian Anda secara permanen</strong> pada pelanggaran pertama (1-Strike Rule).
      </p>
      
      <div style="display:flex; flex-direction:column; gap:16px; width:100%; max-width:400px; margin-bottom:32px; background:rgba(255,255,255,0.05); padding:20px; border-radius:12px; border:1px solid rgba(255,255,255,0.1);">
        <div style="display:flex; align-items:center; gap:12px;">
          <i data-lucide="maximize" style="color:var(--primary); width:20px;"></i>
          <span>Layar Penuh (Fullscreen) Wajib Aktif</span>
        </div>
        <div style="display:flex; align-items:center; gap:12px;">
          <i data-lucide="monitor-off" style="color:var(--primary); width:20px;"></i>
          <span>Anti Split-Screen & Kehilangan Fokus Tab</span>
        </div>
        <div style="display:flex; align-items:center; gap:12px;">
          <i data-lucide="headphones" style="color:var(--primary); width:20px;"></i>
          <span>Deteksi Otomatis Earphone/Bluetooth Audio</span>
        </div>
        <div style="display:flex; align-items:center; gap:12px;">
          <i data-lucide="copy-x" style="color:var(--primary); width:20px;"></i>
          <span>Pencegahan Copy, Paste, & Klik Kanan</span>
        </div>
      </div>

      <button id="btn-start-cbt-lockdown" class="btn btn-primary" style="font-size:18px; padding:16px 32px; font-weight:800; background:var(--danger); border-color:var(--danger); transition: all 0.2s;">
        <i data-lucide="lock"></i> SAYA MENGERTI, MULAI & KUNCI UJIAN
      </button>
    </div>

    <!-- CBT Isolation style -->
    <style>
      .cbt-app-shell {
        display: none; /* Disembunyikan dulu sampai tombol Mulai diklik */
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
      <!-- Anti-Cheat Watermark -->
      <div class="cbt-anti-cheat-watermark">
        ${activeUser.name} | ${activeUser.nip || 'SISWA'} | IP: 192.168.1.104
      </div>

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

  // Evaluasi dan Simpan (termasuk pembersihan keamanan)
  const acHandlers = window._cbtAcHandlers || {};
  const cleanupAntiCheat = () => {
    if (acHandlers.visibility) document.removeEventListener('visibilitychange', acHandlers.visibility);
    if (acHandlers.blur) window.removeEventListener('blur', acHandlers.blur);
    if (acHandlers.resize) window.removeEventListener('resize', acHandlers.resize);
    if (acHandlers.prevent) {
      document.removeEventListener('contextmenu', acHandlers.prevent);
      document.removeEventListener('copy', acHandlers.prevent);
      document.removeEventListener('cut', acHandlers.prevent);
      document.removeEventListener('paste', acHandlers.prevent);
      document.removeEventListener('dragstart', acHandlers.prevent);
      document.removeEventListener('drop', acHandlers.prevent);
    }
    if (acHandlers.keydown) document.removeEventListener('keydown', acHandlers.keydown);
    if (acHandlers.fullscreen) document.removeEventListener('fullscreenchange', acHandlers.fullscreen);
    
    if (document.fullscreenElement) {
      document.exitFullscreen().catch(err => {});
    }
  };

  const evaluateAndSaveCBTResults = () => {
    // Clear timer loop
    if (cbtInterval) clearInterval(cbtInterval);
    cleanupAntiCheat();

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

  // Initial draw setup (but timer is delayed until user clicks Start)
  updateTimerClock();
  renderQuestion();

  // ----------------------------------------------------
  // ANTI-CHEAT & KIOSK MODE LOGIC (1-Strike Rule)
  // ----------------------------------------------------
  const antiCheatOverlay = document.getElementById('cbt-anti-cheat-overlay');
  const cbtShell = document.querySelector('.cbt-app-shell');
  const startLockdownBtn = document.getElementById('btn-start-cbt-lockdown');

  window._cbtAcHandlers = {};

  const terminateCBTForCheating = (reason) => {
    if (cbtInterval) clearInterval(cbtInterval);
    cleanupAntiCheat();

    showToast(`PELANGGARAN! Ujian dihentikan: ${reason}`, 'error');
    db.log(activeUser.id, activeUser.name, 'SECURITY_ALERT', `Kecurangan CBT (${pkg.name}): ${reason}`);

    // Hitung score nyata, tapi kita akan simpan score 0 dengan status PELANGGARAN
    let correctCount = 0;
    pkgQuestions.forEach(q => {
      const studentAns = answers[q.id];
      const correctAns = q.correctAnswer;
      if (q.type === 'PG' || q.type === 'BENAR_SALAH' || q.type === 'ISIAN_SINGKAT' || q.type === 'NUMERIK') {
        if (String(studentAns).trim().toUpperCase() === String(correctAns).trim().toUpperCase()) correctCount++;
      } else if (q.type === 'PG_KOMPLEKS') {
        const sortedStudent = [...(studentAns||[])].sort();
        const sortedCorrect = [...(correctAns||[])].sort();
        if (sortedStudent.length === sortedCorrect.length && sortedStudent.every((val, idx) => val === sortedCorrect[idx])) correctCount++;
      }
    });

    const totalQ = pkgQuestions.length;
    const realScore = Math.round((correctCount / totalQ) * 100);

    const result = {
      id: `res-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      userId: activeUser.id,
      userName: activeUser.name,
      packageId: pkg.id,
      packageName: pkg.name,
      subjectName: subjectName,
      className: className,
      score: 0, // Nilai 0 karena pelanggaran berat 1-Strike
      realScore: realScore,
      status: 'PELANGGARAN', 
      violationReason: reason,
      totalQuestions: totalQ,
      answers: answers,
      timestamp: new Date().toISOString()
    };

    db.insert('results', result);
    alert(`UJIAN DIHENTIKAN PAKSA!\nAlasan: ${reason}\n\nNilai Anda digugurkan menjadi 0. Silakan melapor ke pengawas.`);
    
    // Redirect & Reload to flush everything
    window.location.hash = `#dashboard`;
    window.location.reload();
  };

  const bindAntiCheatListeners = () => {
    // 1. Loss Focus (Tab Switch / Alt-Tab)
    window._cbtAcHandlers.visibility = () => {
      if (document.visibilityState === 'hidden') terminateCBTForCheating('Berpindah Tab / Aplikasi (Visibility Hidden)');
    };
    window._cbtAcHandlers.blur = () => {
      terminateCBTForCheating('Kehilangan Fokus Layar (Alt-Tab / Klik di luar / Notifikasi Muncul)');
    };
    document.addEventListener('visibilitychange', window._cbtAcHandlers.visibility);
    window.addEventListener('blur', window._cbtAcHandlers.blur);

    // 2. Split Screen (Resize)
    window._cbtAcHandlers.resize = () => {
      if (window.innerWidth < window.screen.width * 0.85 || window.innerHeight < window.screen.height * 0.85) {
        terminateCBTForCheating('Membuka Split Screen / Mengubah Ukuran Jendela');
      }
    };
    window.addEventListener('resize', window._cbtAcHandlers.resize);

    // 3. Copy Paste & Context Menu
    window._cbtAcHandlers.prevent = (e) => {
      e.preventDefault();
      return false;
    };
    document.addEventListener('contextmenu', window._cbtAcHandlers.prevent);
    document.addEventListener('copy', window._cbtAcHandlers.prevent);
    document.addEventListener('cut', window._cbtAcHandlers.prevent);
    document.addEventListener('paste', window._cbtAcHandlers.prevent);
    document.addEventListener('dragstart', window._cbtAcHandlers.prevent);
    document.addEventListener('drop', window._cbtAcHandlers.prevent);

    // 4. Keyboard Shortcuts (F12, Ctrl+C, PrintScreen)
    window._cbtAcHandlers.keydown = (e) => {
      // PrintScreen usually fires keyup but sometimes keydown
      if (e.key === 'PrintScreen' || e.code === 'PrintScreen') {
        navigator.clipboard.writeText(''); // Clear clipboard immediately
        terminateCBTForCheating('Mencoba melakukan Screenshot (PrintScreen)');
        e.preventDefault();
      }
      // F12 or Ctrl+Shift+I
      if (e.key === 'F12' || (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'i')) {
        terminateCBTForCheating('Mencoba membuka Developer Tools (F12 / Inspect)');
        e.preventDefault();
      }
      // Ctrl+C, Ctrl+V, Ctrl+X, Ctrl+P
      if (e.ctrlKey && ['c','v','x','p'].includes(e.key.toLowerCase())) {
        e.preventDefault();
      }
    };
    document.addEventListener('keydown', window._cbtAcHandlers.keydown);

    // 5. Exit Fullscreen checking
    window._cbtAcHandlers.fullscreen = () => {
      if (!document.fullscreenElement) {
        terminateCBTForCheating('Keluar dari Mode Layar Penuh (Esc)');
      }
    };
    document.addEventListener('fullscreenchange', window._cbtAcHandlers.fullscreen);
  };

  const startExamFlow = async () => {
    startLockdownBtn.innerHTML = '<i data-lucide="loader" class="spinner"></i> Memeriksa Keamanan...';
    lucide.createIcons();
    startLockdownBtn.disabled = true;

    // 1. Check Audio Devices (Bluetooth/Earphone)
    try {
      // Minta izin mikrofon sesaat untuk mendapatkan label media devices yang valid
      if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
         const stream = await navigator.mediaDevices.getUserMedia({ audio: true }).catch(() => null);
         const devices = await navigator.mediaDevices.enumerateDevices();
         const hasBluetoothOrHeadset = devices.some(d => 
           (d.kind === 'audiooutput' || d.kind === 'audioinput') && 
           (d.label.toLowerCase().includes('bluetooth') || d.label.toLowerCase().includes('headset') || d.label.toLowerCase().includes('airpods') || d.label.toLowerCase().includes('buds'))
         );
         
         if (stream) {
           stream.getTracks().forEach(track => track.stop()); // hentikan stream segera
         }

         if (hasBluetoothOrHeadset) {
           alert("PERINGATAN ANTI-CHEAT:\nPerangkat Bluetooth / Headset terdeteksi! Harap lepaskan atau matikan perangkat audio eksternal Anda sebelum memulai ujian.");
           startLockdownBtn.innerHTML = '<i data-lucide="lock"></i> SAYA MENGERTI, MULAI & KUNCI UJIAN';
           startLockdownBtn.disabled = false;
           lucide.createIcons();
           return; 
         }
      }
    } catch (err) {
      console.warn('Media devices check skipped or failed.', err);
    }

    // 2. Request Fullscreen
    try {
      if (document.documentElement.requestFullscreen) {
        await document.documentElement.requestFullscreen();
      }
    } catch (err) {
      alert("Browser Anda menolak Mode Layar Penuh. Pastikan Anda memberikan izin akses layar penuh.");
      startLockdownBtn.innerHTML = '<i data-lucide="lock"></i> SAYA MENGERTI, MULAI & KUNCI UJIAN';
      startLockdownBtn.disabled = false;
      lucide.createIcons();
      return;
    }

    // 3. Start Exam Setup
    antiCheatOverlay.style.display = 'none';
    cbtShell.style.display = 'flex';
    bindAntiCheatListeners();
    
    // Start timer interval loop NOW
    cbtInterval = setInterval(updateTimerClock, 1000);
    showToast('Ujian Dimulai. Layar Terkunci Anti-Cheat.', 'success');
  };

  startLockdownBtn.addEventListener('click', startExamFlow);
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

// ----------------------------------------------------
// 14. PRD & PRD2 HELPER & NEW VIEW FUNCTIONS
// ----------------------------------------------------

// Pure SVG QR Code Generator
function generateSVGQRCode(text) {
  const size = 25;
  let cells = Array(size).fill(0).map(() => Array(size).fill(false));
  
  const addFinder = (row, col) => {
    for (let r = 0; r < 7; r++) {
      for (let c = 0; c < 7; c++) {
        if (r === 0 || r === 6 || c === 0 || c === 6 || (r >= 2 && r <= 4 && c >= 2 && c <= 4)) {
          cells[row + r][col + c] = true;
        }
      }
    }
  };
  addFinder(0, 0);
  addFinder(0, size - 7);
  addFinder(size - 7, 0);

  let hash = 0;
  for (let i = 0; i < text.length; i++) {
    hash = (hash << 5) - hash + text.charCodeAt(i);
    hash |= 0;
  }
  
  let seed = Math.abs(hash);
  const pseudoRand = () => {
    seed = (seed * 9301 + 49297) % 233280;
    return seed / 233280;
  };

  for (let r = 0; r < size; r++) {
    for (let c = 0; c < size; c++) {
      const inTopLeft = r < 8 && c < 8;
      const inTopRight = r < 8 && c >= size - 8;
      const inBottomLeft = r >= size - 8 && c < 8;
      if (!inTopLeft && !inTopRight && !inBottomLeft) {
        cells[r][c] = pseudoRand() > 0.45;
      }
    }
  }

  const rectSize = 8;
  const svgSize = size * rectSize;
  let rectsSvg = '';
  for (let r = 0; r < size; r++) {
    for (let c = 0; c < size; c++) {
      if (cells[r][c]) {
        rectsSvg += `<rect x="${c * rectSize}" y="${r * rectSize}" width="${rectSize}" height="${rectSize}" fill="#0f172a" />`;
      }
    }
  }

  return `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${svgSize} ${svgSize}" width="200" height="200" style="background:#ffffff; padding:10px; border-radius:12px;">
      <rect width="100%" height="100%" fill="#ffffff" />
      ${rectsSvg}
    </svg>
  `;
}

// Modal Item Analysis
function openItemAnalysisModal(questionId) {
  const q = db.get('questions').find(item => item.id === questionId);
  if (!q) return;

  const stats = db.getQuestionAnalysis(questionId);

  const html = `
    <div class="modal-header">
      <h3 class="modal-title">Analisis Butir Soal - ${q.code}</h3>
      <button class="modal-close" data-close-modal><i data-lucide="x"></i></button>
    </div>
    <div class="modal-body" style="display:flex; flex-direction:column; gap:20px;">
      <div style="background:var(--neutral-100); border:1px solid var(--neutral-300); border-radius:8px; padding:12px; font-size:13px;">
        <strong>Pertanyaan:</strong>
        <div style="margin-top:4px;">${q.questionText}</div>
      </div>

      <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
        <div class="card" style="padding:16px;">
          <span style="font-size:11px; font-weight:700; color:var(--neutral-600); text-transform:uppercase;">Indeks Tingkat Kesulitan (P)</span>
          <div style="font-size:28px; font-weight:800; color:var(--primary); margin-top:4px;">${stats.difficultyIndex}</div>
          <span style="font-size:11px; color:var(--neutral-600);">${stats.correctCount} dari ${stats.totalAttempts} siswa menjawab benar</span>
        </div>

        <div class="card" style="padding:16px;">
          <span style="font-size:11px; font-weight:700; color:var(--neutral-600); text-transform:uppercase;">Indeks Daya Beda (D)</span>
          <div style="font-size:28px; font-weight:800; color:#10b981; margin-top:4px;">${stats.discriminationIndex}</div>
          <span style="font-size:11px; color:var(--neutral-600);">Membedakan kelompok atas & bawah</span>
        </div>
      </div>

      <div>
        <h4 style="font-size:13px; font-weight:700; margin-bottom:8px;">Distribusi Jawaban Siswa</h4>
        <div style="display:flex; flex-direction:column; gap:8px; font-size:12px;">
          ${Object.entries(stats.choiceDistribution).map(([option, count]) => {
            const pct = stats.totalAttempts > 0 ? Math.round((count / stats.totalAttempts) * 100) : 0;
            return `
              <div>
                <div style="display:flex; justify-content:space-between; margin-bottom:2px;">
                  <span>Pilihan ${option}</span>
                  <strong>${count} siswa (${pct}%)</strong>
                </div>
                <div class="analysis-bar-outer">
                  <div class="analysis-bar-inner" style="width:${pct}%;"></div>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      </div>

      <div style="background:#eef2ff; border-left:4px solid var(--primary); padding:12px; border-radius:4px; font-size:12px;">
        <strong style="color:var(--primary);">Rekomendasi Kualitas Soal (${stats.quality}):</strong>
        <p style="margin-top:4px; color:var(--neutral-700);">${stats.recommendation}</p>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" data-close-modal>Tutup</button>
    </div>
  `;

  openModal(html, 'modal-large');
}

// Modal Exam Code & QR Code
function openQRCodeModal(packageId) {
  const pkg = db.get('packages').find(p => p.id === packageId);
  if (!pkg) return;

  const examCode = pkg.examCode || `EXAM-INF-${Math.floor(1000 + Math.random()*9000)}`;
  const qrSvg = generateSVGQRCode(JSON.stringify({ examCode, duration: pkg.duration || 60, name: pkg.name }));

  const html = `
    <div class="modal-header">
      <h3 class="modal-title">Kode Ujian & QR Code Paket Ujian</h3>
      <button class="modal-close" data-close-modal><i data-lucide="x"></i></button>
    </div>
    <div class="modal-body" style="display:flex; flex-direction:column; align-items:center; gap:16px; text-align:center;">
      <div class="qr-code-card">
        <h3 style="font-size:16px; color:var(--neutral-900); margin:0;">${pkg.name}</h3>
        <p style="font-size:12px; color:var(--neutral-600); margin:0;">Pindai QR Code menggunakan Aplikasi Android Ujian Siswa</p>
        
        <div class="qr-code-svg-wrap">
          ${qrSvg}
        </div>
        
        <div style="background:var(--neutral-100); border:1px solid var(--neutral-400); padding:10px 20px; border-radius:8px;">
          <div style="font-size:11px; font-weight:700; color:var(--neutral-600); text-transform:uppercase;">KODE PAKET UJIAN</div>
          <div style="font-size:24px; font-weight:900; color:var(--primary); letter-spacing:2px; margin-top:2px;">${examCode}</div>
          <div style="font-size:12px; font-weight:600; color:var(--neutral-700); margin-top:4px;">Durasi Ujian: ${pkg.duration || 60} Menit</div>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" data-close-modal>Tutup</button>
      <button class="btn btn-primary" onclick="window.print()"><i data-lucide="printer"></i> Cetak QR Code</button>
    </div>
  `;

  openModal(html);
}

// Modal Student Permission Management
function openStudentPermissionModal(packageId) {
  const pkg = db.get('packages').find(p => p.id === packageId);
  if (!pkg) return;

  const students = db.get('users').filter(u => u.role === 'SISWA');

  const html = `
    <div class="modal-header">
      <h3 class="modal-title">Kelola Izin Pengerjaan Siswa - ${pkg.name}</h3>
      <button class="modal-close" data-close-modal><i data-lucide="x"></i></button>
    </div>
    <div class="modal-body" style="display:flex; flex-direction:column; gap:16px; max-height:60vh; overflow-y:auto;">
      <p style="font-size:13px; color:var(--neutral-600);">
        Pendidik dapat memberi izin atau memblokir akses 1 siswa spesifik untuk mengikuti paket ujian ini.
      </p>
      
      <div class="table-wrapper">
        <table class="table" style="font-size:13px;">
          <thead>
            <tr>
              <th>Nama Siswa</th>
              <th>Kelas</th>
              <th>Status Izin Ujian</th>
              <th style="text-align:right;">Aksi Izin</th>
            </tr>
          </thead>
          <tbody>
            ${students.map(s => {
              const isAllowed = db.canStudentTakeExam(s.id, pkg.id);
              const statusBadge = isAllowed
                ? '<span class="badge badge-success">IZIN DIBERIKAN</span>'
                : '<span class="badge badge-danger">BLOCKED / DITOLAK</span>';
              return `
                <tr>
                  <td style="font-weight:700;">${s.name}</td>
                  <td>Kelas X</td>
                  <td>${statusBadge}</td>
                  <td style="text-align:right;">
                    <button class="btn ${isAllowed ? 'btn-danger' : 'btn-primary'} btn-toggle-student-perm" data-student-id="${s.id}" data-allowed="${isAllowed}" style="padding:6px 12px; font-size:11px;">
                      ${isAllowed ? 'Blokir Akses' : 'Berikan Izin'}
                    </button>
                  </td>
                </tr>
              `;
            }).join('') || '<tr><td colspan="4" style="text-align:center;">Tidak ada data siswa.</td></tr>'}
          </tbody>
        </table>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-secondary" data-close-modal>Selesai</button>
    </div>
  `;

  openModal(html, 'modal-large');

  document.querySelectorAll('.btn-toggle-student-perm').forEach(btn => {
    btn.addEventListener('click', () => {
      const studentId = btn.dataset.studentId;
      const currentlyAllowed = btn.dataset.allowed === 'true';
      db.toggleStudentExamPermission(pkg.id, studentId, !currentlyAllowed);
      showToast(`Izin siswa berhasil diperbarui.`, 'success');
      openStudentPermissionModal(packageId);
    });
  });
}

// Multi-School Management Page
function renderKelolaSekolah(mount) {
  const schools = db.get('schools');

  mount.innerHTML = `
    <div class="action-row">
      <div></div>
      <button class="btn btn-primary" id="btn-tambah-sekolah">
        <i data-lucide="plus"></i>
        <span>Tambah Sekolah Baru</span>
      </button>
    </div>

    <div class="card">
      <div class="card-body" style="padding:0;">
        <div class="table-wrapper">
          <table class="table">
            <thead>
              <tr>
                <th>Nama Sekolah</th>
                <th>NPSN</th>
                <th>Alamat & Kontak</th>
                <th>Tahun Ajaran</th>
                <th style="text-align:right;">Aksi</th>
              </tr>
            </thead>
            <tbody>
              ${schools.map(s => `
                <tr>
                  <td>
                    <div style="display:flex; align-items:center; gap:12px;">
                      <img src="${s.logo || 'https://images.unsplash.com/photo-1592280771190-3e2e4d571952?w=150&h=150&fit=crop&q=80'}" style="width:36px; height:36px; border-radius:50%; object-fit:cover;">
                      <div style="font-weight:700; color:var(--neutral-900);">${s.name}</div>
                    </div>
                  </td>
                  <td><span class="badge badge-neutral">${s.npsn || '-'}</span></td>
                  <td>
                    <div style="font-size:12px; color:var(--neutral-700);">${s.address || '-'}</div>
                    <div style="font-size:11px; color:var(--neutral-500);">${s.email || '-'} | Telp: ${s.phone || '-'}</div>
                  </td>
                  <td><span class="badge badge-primary">${s.academicYear || '2025/2026'}</span></td>
                  <td style="text-align:right;">
                    <button class="btn btn-secondary btn-icon edit-sch" data-id="${s.id}" title="Edit Sekolah"><i data-lucide="edit"></i></button>
                    ${schools.length > 1 ? `<button class="btn btn-danger btn-icon del-sch" data-id="${s.id}" title="Hapus Sekolah"><i data-lucide="trash-2"></i></button>` : ''}
                  </td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  document.getElementById('btn-tambah-sekolah').addEventListener('click', () => {
    const name = prompt('Masukkan Nama Sekolah Baru:');
    if (!name) return;
    const npsn = prompt('Masukkan NPSN Sekolah:');
    const academicYear = prompt('Masukkan Tahun Ajaran (contoh: 2025/2026):', '2025/2026');

    db.insert('schools', {
      name,
      npsn: npsn || '10299999',
      academicYear: academicYear || '2025/2026',
      address: 'Jl. Utama Sekolah',
      email: `info@${name.toLowerCase().replace(/[^a-z0-9]/g, '')}.sch.id`,
      phone: '0623-12345',
      logo: 'https://images.unsplash.com/photo-1592280771190-3e2e4d571952?w=150&h=150&fit=crop&q=80'
    });

    showToast('Sekolah baru berhasil ditambahkan!', 'success');
    renderKelolaSekolah(mount);
  });

  mount.querySelectorAll('.edit-sch').forEach(btn => {
    btn.addEventListener('click', () => {
      const sch = schools.find(s => s.id === btn.dataset.id);
      const newName = prompt('Ubah Nama Sekolah:', sch.name);
      if (newName) {
        db.update('schools', sch.id, { name: newName });
        showToast('Data sekolah diperbarui.', 'success');
        renderKelolaSekolah(mount);
      }
    });
  });

  mount.querySelectorAll('.del-sch').forEach(btn => {
    btn.addEventListener('click', () => {
      if (confirm('Apakah Anda yakin ingin menghapus sekolah ini?')) {
        db.delete('schools', btn.dataset.id);
        showToast('Sekolah berhasil dihapus.', 'success');
        renderKelolaSekolah(mount);
      }
    });
  });
}

// Privileges Matrix Control Page
function renderPrivileges(mount) {
  const roles = ['ADMIN_SEKOLAH', 'GURU', 'REVIEWER', 'SISWA'];

  const privilegeKeys = [
    { key: 'canExport', label: 'Ekspor Berkas Word & PDF' },
    { key: 'canImport', label: 'Impor Soal Excel & CSV' },
    { key: 'canAutoGenerate', label: 'Generator Paket Soal Otomatis' },
    { key: 'canPracticeExam', label: 'Latihan Ujian Mandiri' },
    { key: 'requireReviewer', label: 'Wajib Persetujuan Reviewer' },
    { key: 'canStudentScanQR', label: 'Siswa Boleh Scan QR / Input Kode' }
  ];

  mount.innerHTML = `
    <div class="card">
      <div class="card-header">
        <h3 class="card-title">Matriks Pengaturan Hak Akses & Privilege Fitur Peran</h3>
      </div>
      <div class="card-body" style="padding:0;">
        <div class="table-wrapper">
          <table class="table privilege-matrix-table">
            <thead>
              <tr>
                <th>Fitur & Privilege System</th>
                ${roles.map(r => `<th>${r.replace('_', ' ')}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${privilegeKeys.map(p => `
                <tr>
                  <td style="font-weight:700; color:var(--neutral-900);">${p.label}</td>
                  ${roles.map(role => {
                    const isChecked = db.can(role, p.key);
                    return `
                      <td>
                        <label class="toggle-switch">
                          <input type="checkbox" class="priv-chk" data-role="${role}" data-key="${p.key}" ${isChecked ? 'checked' : ''}>
                          <span class="slider-switch"></span>
                        </label>
                      </td>
                    `;
                  }).join('')}
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  mount.querySelectorAll('.priv-chk').forEach(chk => {
    chk.addEventListener('change', () => {
      const role = chk.dataset.role;
      const key = chk.dataset.key;
      const val = chk.checked;
      db.setPrivilege(role, key, val);
      showToast(`Privilege "${key}" untuk peran ${role} diperbarui ke: ${val ? 'AKTIF' : 'NONAKTIF'}.`, 'success');
    });
  });
}

// Android Exam App Kiosk Mode Simulator Page
function renderAndroidExamAppScreen(packageId) {
  const packages = db.get('packages');
  const activePackage = packages.find(p => p.id === packageId) || packages[0];

  appMount.innerHTML = `
    <div class="android-kiosk-wrapper">
      <div class="android-kiosk-frame">
        <div class="kiosk-notch">
          <div class="kiosk-camera-lens"></div>
        </div>

        <div class="kiosk-screen">
          <div class="kiosk-status-bar">
            <span>08:00 AM</span>
            <div style="display:flex; gap:6px; align-items:center;">
              <i data-lucide="wifi" style="width:12px; height:12px;"></i>
              <i data-lucide="shield-check" style="width:12px; height:12px;"></i>
              <span>100%</span>
            </div>
          </div>

          <div class="kiosk-app-header">
            <div style="font-size:11px; font-weight:800; letter-spacing:1px; color:#a5b4fc;">APLIKASI UJIAN ANDROID KIOSK</div>
            <h2 style="font-size:18px; font-weight:800; margin-top:4px; color:white;">BankSoalPro Exam Agent</h2>
          </div>

          <div class="kiosk-app-body">
            <div class="card" style="padding:16px; background:#eef2ff; border:1px solid #c7d2fe;">
              <div style="font-size:12px; font-weight:700; color:#3730a3;">Status Siswa Active</div>
              <div style="font-size:14px; font-weight:800; color:#1e1b4b; margin-top:2px;">${activeUser.name}</div>
              <div style="font-size:11px; color:#4338ca;">Perangkat Terverifikasi (Kiosk Lockdown Enabled)</div>
            </div>

            <!-- Mode Switch Tabs -->
            <div style="display:flex; gap:8px; border-bottom:1px solid var(--neutral-300); padding-bottom:8px;">
              <button class="btn btn-primary" id="btn-kiosk-tab-qr" style="flex:1; font-size:11px; padding:8px;">
                <i data-lucide="qr-code"></i> Scan QR Code
              </button>
              <button class="btn btn-secondary" id="btn-kiosk-tab-code" style="flex:1; font-size:11px; padding:8px;">
                <i data-lucide="key"></i> Input Kode Paket
              </button>
            </div>

            <!-- Pane 1: QR Scanner Sim -->
            <div id="kiosk-pane-qr" style="display:flex; flex-direction:column; gap:12px;">
              <div class="kiosk-scan-box" id="trigger-scan-qr">
                <i data-lucide="camera" style="width:40px; height:40px; color:#4f46e5; margin-bottom:8px;"></i>
                <div style="font-size:13px; font-weight:800; color:#1e1b4b;">Arahkan Kamera ke QR Code</div>
                <div style="font-size:11px; color:#6366f1; margin-top:4px;">Klik untuk mensimulasikan pemindaian cepat QR Code Ujian</div>
              </div>
            </div>

            <!-- Pane 2: Manual Code Input -->
            <div id="kiosk-pane-code" style="display:none; flex-direction:column; gap:12px;">
              <div class="form-group">
                <label class="form-label" for="kiosk-code-input">Masukkan Kode Paket Ujian</label>
                <input type="text" id="kiosk-code-input" class="form-input" style="text-transform:uppercase; font-size:18px; font-weight:800; text-align:center; letter-spacing:2px;" placeholder="EXAM-INF-8921">
              </div>
              <button class="btn btn-primary" id="btn-submit-kiosk-code" style="width:100%;">
                <span>Verifikasi & Mulai Ujian</span>
                <i data-lucide="arrow-right"></i>
              </button>
            </div>

            <div style="border-top:1px dashed var(--neutral-300); padding-top:12px;">
              <div style="font-size:11px; font-weight:700; color:var(--neutral-600); margin-bottom:6px;">Pilihan Ujian Tersedia:</div>
              <div style="display:flex; flex-direction:column; gap:6px;">
                ${packages.map(p => `
                  <button class="btn btn-secondary quick-start-pkg" data-id="${p.id}" style="justify-content:space-between; font-size:11px; padding:8px 12px; text-align:left;">
                    <span>${p.name}</span>
                    <span class="badge badge-primary">${p.duration || 60}m</span>
                  </button>
                `).join('')}
              </div>
            </div>
          </div>

          <div class="kiosk-nav-bar">
            <a href="#dashboard" style="color:inherit;"><i data-lucide="home"></i></a>
            <i data-lucide="shield"></i>
            <i data-lucide="square"></i>
          </div>
        </div>
      </div>
    </div>
  `;

  lucide.createIcons();

  const tabQr = document.getElementById('btn-kiosk-tab-qr');
  const tabCode = document.getElementById('btn-kiosk-tab-code');
  const paneQr = document.getElementById('kiosk-pane-qr');
  const paneCode = document.getElementById('kiosk-pane-code');

  tabQr.addEventListener('click', () => {
    tabQr.className = 'btn btn-primary';
    tabCode.className = 'btn btn-secondary';
    paneQr.style.display = 'flex';
    paneCode.style.display = 'none';
  });

  tabCode.addEventListener('click', () => {
    tabCode.className = 'btn btn-primary';
    tabQr.className = 'btn btn-secondary';
    paneCode.style.display = 'flex';
    paneQr.style.display = 'none';
  });

  const launchExam = (pkgId) => {
    const pkg = db.get('packages').find(p => p.id === pkgId);
    if (!pkg) {
      showToast('Paket ujian tidak ditemukan.', 'error');
      return;
    }

    if (!db.canStudentTakeExam(activeUser.id, pkg.id)) {
      showToast('AKSES DITOLAK: Guru belum memberikan izin pengerjaan paket ujian ini untuk Anda!', 'error');
      return;
    }

    showToast(`Memulai ujian "${pkg.name}" via Android Exam Kiosk...`, 'success');
    window.location.hash = `#cbt?id=${pkg.id}`;
  };

  document.getElementById('trigger-scan-qr').addEventListener('click', () => {
    launchExam(activePackage.id);
  });

  document.getElementById('btn-submit-kiosk-code').addEventListener('click', () => {
    const code = document.getElementById('kiosk-code-input').value.trim().toUpperCase();
    const pkg = db.get('packages').find(p => p.examCode === code) || activePackage;
    launchExam(pkg.id);
  });

  mount.querySelectorAll('.quick-start-pkg').forEach(btn => {
    btn.addEventListener('click', () => launchExam(btn.dataset.id));
  });
}

// ----------------------------------------------------
// 14. PAGE: PROFIL SAYA & KELOLA USER (Admin)
// ----------------------------------------------------
function renderProfilSaya(mount) {
  const subjects = db.get('subjects');
  const userSubjects = (activeUser.subjectIds || []).map(id => {
    const sub = subjects.find(s => s.id === id);
    return sub ? sub.name : id;
  }).join(', ') || 'Belum diatur';

  mount.innerHTML = `
    <div class="card" style="max-width:600px; margin:0 auto;">
      <div class="card-header" style="display:flex; gap:16px; align-items:center;">
        <div style="width:64px; height:64px; border-radius:50%; background:var(--primary); color:white; display:flex; align-items:center; justify-content:center; font-size:24px; font-weight:700;">
          ${activeUser.name.charAt(0)}
        </div>
        <div>
          <h3 class="card-title" style="font-size:20px; margin-bottom:4px;">${activeUser.name}</h3>
          <span class="badge badge-primary">${activeUser.role}</span>
        </div>
      </div>
      <div class="card-body" style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-group">
          <label class="form-label">Email</label>
          <input type="text" class="form-input" value="${activeUser.email}" disabled>
        </div>
        <div class="form-group">
          <label class="form-label">NIP / NISN</label>
          <input type="text" class="form-input" value="${activeUser.nip || activeUser.nisn || '-'}" disabled>
        </div>
        <div class="form-group">
          <label class="form-label">Status Akun</label>
          <input type="text" class="form-input" value="${activeUser.status}" disabled>
        </div>
        ${activeUser.role === 'GURU' ? `
        <div class="form-group">
          <label class="form-label">Mata Pelajaran yang Diampu</label>
          <input type="text" class="form-input" value="${userSubjects}" disabled>
        </div>
        ` : ''}
      </div>
    </div>
  `;
}

// ----------------------------------------------------
// PAGE: KELOLA SISWA (CRUD + Ekspor/Impor CSV)
// ----------------------------------------------------
function renderKelolaSiswa(mount) {
  if (activeUser.role !== 'SUPER_ADMIN' && activeUser.role !== 'ADMIN_SEKOLAH') {
    mount.innerHTML = '<p style="color:red;font-weight:bold;padding:20px;">Akses Ditolak.</p>';
    return;
  }
  const schools = db.get('schools') || [];
  let allSiswa = (db.get('users') || []).filter(u => u.role === 'SISWA');
  if (activeUser.role === 'ADMIN_SEKOLAH') {
    allSiswa = allSiswa.filter(u => u.schoolId === activeUser.schoolId);
  }

  mount.innerHTML = `
    <div style="display:flex; flex-direction:column; gap:20px;">
      <div class="card">
        <div class="card-body" style="display:flex; gap:12px; flex-wrap:wrap; align-items:center; padding:16px;">
          <button class="btn btn-primary" id="ks-btn-add"><i data-lucide="plus"></i> Tambah Siswa</button>
          <button class="btn btn-secondary" id="ks-btn-export"><i data-lucide="download"></i> Ekspor CSV</button>
          <label class="btn btn-secondary" style="cursor:pointer; margin:0;" title="Impor dari file CSV">
            <i data-lucide="upload"></i> Impor CSV
            <input type="file" id="ks-import-input" accept=".csv" style="display:none;">
          </label>
          <span style="font-size:12px; color:var(--neutral-500); margin-left:auto;">Total: <strong>${allSiswa.length}</strong> siswa</span>
        </div>
      </div>
      <div class="card">
        <div class="card-header"><h3 class="card-title">Daftar Siswa</h3></div>
        <div class="card-body" style="padding:0; overflow-x:auto;">
          <table class="data-table">
            <thead>
              <tr>
                <th>#</th><th>Nama Siswa</th><th>NISN</th><th>Email</th><th>Kelas</th><th>Sekolah</th><th>Status</th><th>Aksi</th>
              </tr>
            </thead>
            <tbody>
              ${allSiswa.length === 0
                ? '<tr><td colspan="8" style="text-align:center;padding:32px;color:var(--neutral-500);">Belum ada data siswa.</td></tr>'
                : allSiswa.map((s, idx) => {
                    const school = schools.find(sc => sc.id === s.schoolId)?.name || '-';
                    return `<tr>
                      <td style="color:var(--neutral-500);">${idx + 1}</td>
                      <td><strong>${s.name}</strong></td>
                      <td style="font-family:monospace;">${s.nisn || '-'}</td>
                      <td>${s.email}</td>
                      <td>${s.kelas || '-'}</td>
                      <td style="font-size:12px;">${school}</td>
                      <td><span class="badge ${s.status === 'AKTIF' ? 'badge-success' : 'badge-danger'}">${s.status}</span></td>
                      <td>
                        <button class="btn btn-secondary btn-icon ks-edit" data-id="${s.id}" title="Edit"><i data-lucide="edit"></i></button>
                        <button class="btn btn-danger btn-icon ks-delete" data-id="${s.id}" title="Hapus"><i data-lucide="trash-2"></i></button>
                      </td>
                    </tr>`;
                  }).join('')}
            </tbody>
          </table>
        </div>
      </div>
      <div class="card" style="border:1px dashed var(--neutral-300);">
        <div class="card-body" style="padding:14px;">
          <p style="font-size:12px; color:var(--neutral-600); margin:0;">
            <strong>Format CSV Impor:</strong>
            <code style="background:var(--neutral-100);padding:2px 6px;border-radius:4px;">nama,nisn,email,password,kelas,schoolId,status</code>
          </p>
        </div>
      </div>
    </div>
  `;
  refreshIcons();

  mount.querySelector('#ks-btn-add').onclick = () => openSiswaModal(null, mount);

  mount.querySelectorAll('.ks-edit').forEach(btn => {
    btn.onclick = () => openSiswaModal(btn.dataset.id, mount);
  });

  mount.querySelectorAll('.ks-delete').forEach(btn => {
    btn.onclick = () => {
      if (confirm('Yakin ingin menghapus data siswa ini secara permanen?')) {
        db.delete('users', btn.dataset.id);
        showToast('Data siswa berhasil dihapus.', 'success');
        renderKelolaSiswa(mount);
      }
    };
  });

  mount.querySelector('#ks-btn-export').onclick = () => {
    const header = 'nama,nisn,email,password,kelas,schoolId,status';
    const rows = allSiswa.map(s => `"${s.name}","${s.nisn || ''}","${s.email}","${s.password || ''}","${s.kelas || ''}","${s.schoolId || ''}","${s.status || 'AKTIF'}"`);
    const csv = [header, ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `data_siswa_${Date.now()}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    showToast('Data siswa berhasil diekspor.', 'success');
  };

  mount.querySelector('#ks-import-input').onchange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const lines = ev.target.result.split('\n').filter(l => l.trim());
      let imported = 0;
      lines.forEach((line, idx) => {
        if (idx === 0 && line.toLowerCase().includes('nama')) return;
        const parts = line.split(',').map(p => p.trim().replace(/^"|"$/g, ''));
        if (parts.length < 3) return;
        const [nama, nisn, email, password, kelas, schoolId, status] = parts;
        if (!nama || !email) return;
        db.insert('users', {
          name: nama, nisn: nisn || '', email,
          password: password || 'siswa123', kelas: kelas || '',
          schoolId: schoolId || (activeUser.role === 'ADMIN_SEKOLAH' ? activeUser.schoolId : ''),
          role: 'SISWA', status: status || 'AKTIF'
        });
        imported++;
      });
      showToast(`${imported} data siswa berhasil diimpor.`, 'success');
      renderKelolaSiswa(mount);
    };
    reader.readAsText(file);
  };
}

function openSiswaModal(userId = null, parentMount = null) {
  const gModal = document.getElementById('global-modal');
  const mContent = document.getElementById('modal-content-container');
  if (!gModal || !mContent) { showToast('Komponen modal tidak ditemukan!', 'error'); return; }

  const isEdit = !!userId;
  const s = isEdit ? (db.get('users').find(u => u.id === userId) || {}) : {};
  const schools = db.get('schools') || [];
  const isSuperAdmin = activeUser.role === 'SUPER_ADMIN';

  let schoolOpts = '';
  if (isSuperAdmin) {
    schoolOpts = '<option value="">-- Pilih Sekolah --</option>';
    schools.forEach(sc => {
      schoolOpts += `<option value="${sc.id}" ${s.schoolId === sc.id ? 'selected' : ''}>${sc.name}</option>`;
    });
  } else {
    const mySchool = schools.find(x => x.id === activeUser.schoolId);
    schoolOpts = mySchool ? `<option value="${mySchool.id}" selected>${mySchool.name}</option>` : '';
  }

  mContent.innerHTML = `
    <div style="padding:28px; min-width:420px; max-width:500px; width:100%;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
        <h3 style="font-size:18px; font-weight:800; color:var(--neutral-900); margin:0;">${isEdit ? '&#x270F;&#xFE0F; Edit Siswa' : '&#x2795; Tambah Siswa Baru'}</h3>
        <button type="button" id="sm-close" style="background:none;border:1px solid var(--neutral-300);border-radius:8px;padding:6px 10px;cursor:pointer;font-size:16px;">&times;</button>
      </div>
      <form id="sm-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group"><label class="form-label">Nama Lengkap *</label><input type="text" id="sm-name" class="form-input" value="${s.name || ''}" placeholder="Nama lengkap siswa" required></div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group"><label class="form-label">NISN</label><input type="text" id="sm-nisn" class="form-input" value="${s.nisn || ''}" placeholder="0012345678"></div>
          <div class="form-group"><label class="form-label">Kelas</label><input type="text" id="sm-kelas" class="form-input" value="${s.kelas || ''}" placeholder="X-IPA-1"></div>
        </div>
        <div class="form-group"><label class="form-label">Email *</label><input type="email" id="sm-email" class="form-input" value="${s.email || ''}" placeholder="email@siswa.sch.id" required></div>
        <div class="form-group">
          <label class="form-label">Kata Sandi ${isEdit ? '(kosongkan jika tidak diubah)' : '*'}</label>
          <input type="text" id="sm-pass" class="form-input" value="${isEdit ? (s.password || '') : ''}" placeholder="${isEdit ? 'Biarkan kosong jika tidak ingin diubah' : 'Kata sandi siswa'}">
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label class="form-label">Sekolah *</label>
            <select id="sm-school" class="form-input" ${!isSuperAdmin ? 'disabled' : ''}>${schoolOpts}</select>
          </div>
          <div class="form-group">
            <label class="form-label">Status</label>
            <select id="sm-status" class="form-input">
              <option value="AKTIF" ${(s.status || 'AKTIF') === 'AKTIF' ? 'selected' : ''}>Aktif</option>
              <option value="NONAKTIF" ${s.status === 'NONAKTIF' ? 'selected' : ''}>Nonaktif</option>
            </select>
          </div>
        </div>
        <div style="display:flex; gap:12px; margin-top:8px; justify-content:flex-end;">
          <button type="button" id="sm-cancel" class="btn btn-secondary">Batal</button>
          <button type="submit" class="btn btn-primary">&#x1F4BE; Simpan</button>
        </div>
      </form>
    </div>
  `;
  gModal.style.display = 'flex';

  const closeModal = () => { gModal.style.display = 'none'; };
  mContent.querySelector('#sm-close').onclick = closeModal;
  mContent.querySelector('#sm-cancel').onclick = closeModal;
  gModal.onclick = (e) => { if (e.target === gModal) closeModal(); };

  mContent.querySelector('#sm-form').onsubmit = (e) => {
    e.preventDefault();
    const nama = mContent.querySelector('#sm-name').value.trim();
    const email = mContent.querySelector('#sm-email').value.trim();
    const pass = mContent.querySelector('#sm-pass').value.trim();
    const nisn = mContent.querySelector('#sm-nisn').value.trim();
    const kelas = mContent.querySelector('#sm-kelas').value.trim();
    const schoolEl = mContent.querySelector('#sm-school');
    const schoolId = isSuperAdmin ? (schoolEl ? schoolEl.value : '') : activeUser.schoolId;
    const status = mContent.querySelector('#sm-status').value;

    if (!nama || !email) { showToast('Nama dan Email wajib diisi!', 'error'); return; }
    if (!isEdit && !pass) { showToast('Kata sandi wajib diisi untuk siswa baru!', 'error'); return; }
    if (!schoolId) { showToast('Pilih sekolah terlebih dahulu!', 'error'); return; }

    const data = { ...s, id: isEdit ? s.id : ('stu-' + Date.now()), name: nama, email, nisn, kelas, schoolId, status, role: 'SISWA' };
    if (pass) data.password = pass;

    if (isEdit) { db.update('users', s.id, data); showToast(`Data "${nama}" berhasil diperbarui.`, 'success'); }
    else { db.insert('users', data); showToast(`Siswa "${nama}" berhasil ditambahkan.`, 'success'); }

    closeModal();
    const mountEl = parentMount || document.getElementById('shell-content-mount');
    if (mountEl) renderKelolaSiswa(mountEl);
  };
}

// ============================================================
// KELOLA PENGGUNA (Super Admin & Admin Sekolah)
// ============================================================
function renderKelolaUser(mount) {
  if (activeUser.role !== 'SUPER_ADMIN' && activeUser.role !== 'ADMIN_SEKOLAH') {
    mount.innerHTML = '<p style="color:red; font-weight:bold; padding:20px;">Akses Ditolak: Fitur khusus Super Admin dan Admin Sekolah.</p>';
    return;
  }

  let allUsers = db.get('users') || [];
  if (activeUser.role === 'ADMIN_SEKOLAH') {
    allUsers = allUsers.filter(u => u.schoolId === activeUser.schoolId && u.role !== 'SUPER_ADMIN');
  }
  const schools = db.get('schools') || [];

  const roleLabel = { SUPER_ADMIN: 'Super Admin', ADMIN_SEKOLAH: 'Admin Sekolah', GURU: 'Guru', REVIEWER: 'Reviewer Soal', SISWA: 'Siswa' };
  const roleBadge = { SUPER_ADMIN: 'badge-primary', ADMIN_SEKOLAH: 'badge-warning', GURU: 'badge-secondary', REVIEWER: 'badge-info', SISWA: 'badge-success' };

  mount.innerHTML = `
    <div class="card">
      <div class="card-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
        <h3 class="card-title">Daftar Semua Pengguna</h3>
        <button class="btn btn-primary" id="btn-add-global-user"><i data-lucide="plus"></i> Tambah Pengguna Baru</button>
      </div>
      <div class="card-body" style="padding:0; overflow-x:auto;">
        <table class="data-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Nama &amp; NIP/NISN</th>
              <th>Email</th>
              <th>Role</th>
              <th>Asal Sekolah</th>
              <th>Status</th>
              <th>Aksi</th>
            </tr>
          </thead>
          <tbody>
            ${allUsers.length === 0
              ? '<tr><td colspan="7" style="text-align:center;padding:32px;color:var(--neutral-500);">Belum ada data pengguna.</td></tr>'
              : allUsers.map((u, idx) => {
                  const school = schools.find(s => s.id === u.schoolId)?.name || (u.role === 'SUPER_ADMIN' ? 'Semua Sekolah' : '-');
                  const badge = roleBadge[u.role] || 'badge-secondary';
                  const label = roleLabel[u.role] || u.role;
                  return `<tr>
                    <td style="color:var(--neutral-500);">${idx + 1}</td>
                    <td>
                      <strong>${u.name}</strong><br>
                      <span style="font-size:11px; color:var(--neutral-500);">${u.nip || u.nisn || u.email}</span>
                    </td>
                    <td>${u.email}</td>
                    <td><span class="badge ${badge}">${label}</span></td>
                    <td><span style="font-size:12px;">${school}</span></td>
                    <td><span class="badge ${u.status === 'AKTIF' ? 'badge-success' : 'badge-danger'}">${u.status}</span></td>
                    <td>
                      <button class="btn btn-secondary btn-icon edit-global-user" data-id="${u.id}" title="Edit Data"><i data-lucide="edit"></i></button>
                      ${u.id !== activeUser.id ? `<button class="btn btn-danger btn-icon delete-global-user" data-id="${u.id}" title="Hapus"><i data-lucide="trash-2"></i></button>` : ''}
                    </td>
                  </tr>`;
                }).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
  refreshIcons();

  mount.querySelector('#btn-add-global-user').onclick = () => openUserModal(null, mount);

  mount.querySelectorAll('.edit-global-user').forEach(btn => {
    btn.onclick = () => openUserModal(btn.dataset.id, mount);
  });

  mount.querySelectorAll('.delete-global-user').forEach(btn => {
    btn.onclick = () => {
      if (confirm('Yakin ingin menghapus pengguna ini secara permanen dari sistem?')) {
        db.delete('users', btn.dataset.id);
        showToast('Pengguna berhasil dihapus.', 'success');
        renderKelolaUser(mount);
      }
    };
  });
}

function openUserModal(userId = null, parentMount = null) {
  const gModal = document.getElementById('global-modal');
  const mContent = document.getElementById('modal-content-container');
  if (!gModal || !mContent) { showToast('Komponen modal tidak ditemukan!', 'error'); return; }

  const isEdit = !!userId;
  const u = isEdit ? (db.get('users').find(user => user.id === userId) || {}) : {};
  const schools = db.get('schools') || [];

  const isSuperAdmin = activeUser.role === 'SUPER_ADMIN';
  const isAdminSekolah = activeUser.role === 'ADMIN_SEKOLAH';

  let schoolOptionsHtml = '';
  if (isSuperAdmin) {
    schoolOptionsHtml = '<option value="">-- Tidak Terikat Sekolah --</option>';
    schools.forEach(s => {
      schoolOptionsHtml += `<option value="${s.id}" ${u.schoolId === s.id ? 'selected' : ''}>${s.name}</option>`;
    });
  } else if (isAdminSekolah) {
    const mySchool = schools.find(s => s.id === activeUser.schoolId);
    schoolOptionsHtml = mySchool ? `<option value="${mySchool.id}" selected>${mySchool.name}</option>` : '';
  }

  let roleOptionsHtml = '';
  if (isSuperAdmin) {
    roleOptionsHtml = `
      <option value="SUPER_ADMIN" ${u.role === 'SUPER_ADMIN' ? 'selected' : ''}>Super Admin</option>
      <option value="ADMIN_SEKOLAH" ${u.role === 'ADMIN_SEKOLAH' ? 'selected' : ''}>Admin Sekolah</option>
      <option value="GURU" ${(!u.role || u.role === 'GURU') ? 'selected' : ''}>Guru</option>
      <option value="REVIEWER" ${u.role === 'REVIEWER' ? 'selected' : ''}>Reviewer Soal</option>
      <option value="SISWA" ${u.role === 'SISWA' ? 'selected' : ''}>Siswa</option>
    `;
  } else {
    roleOptionsHtml = `
      <option value="GURU" ${(!u.role || u.role === 'GURU') ? 'selected' : ''}>Guru</option>
      <option value="REVIEWER" ${u.role === 'REVIEWER' ? 'selected' : ''}>Reviewer Soal</option>
      <option value="SISWA" ${u.role === 'SISWA' ? 'selected' : ''}>Siswa</option>
    `;
  }

  mContent.innerHTML = `
    <div style="padding:28px; min-width:440px; max-width:520px; width:100%;">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
        <h3 style="font-size:18px; font-weight:800; color:var(--neutral-900); margin:0;">
          ${isEdit ? '&#x270F;&#xFE0F; Edit Pengguna' : '&#x2795; Tambah Pengguna Baru'}
        </h3>
        <button type="button" id="um-close-btn" style="background:none;border:1px solid var(--neutral-300);border-radius:8px;padding:6px 10px;cursor:pointer;font-size:16px;">&times;</button>
      </div>
      <form id="um-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label class="form-label">Nama Lengkap *</label>
          <input type="text" id="um-name" class="form-input" placeholder="Nama lengkap pengguna" value="${u.name || ''}" required>
        </div>
        <div class="form-group">
          <label class="form-label">Email *</label>
          <input type="email" id="um-email" class="form-input" placeholder="email@sekolah.sch.id" value="${u.email || ''}" required>
        </div>
        <div class="form-group">
          <label class="form-label">Kata Sandi ${isEdit ? '(kosongkan jika tidak diubah)' : '*'}</label>
          <input type="text" id="um-pass" class="form-input" placeholder="${isEdit ? 'Biarkan kosong jika tidak ingin mengubah' : 'Masukkan kata sandi'}" value="${isEdit ? (u.password || '') : ''}">
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:14px;">
          <div class="form-group">
            <label class="form-label">Role *</label>
            <select id="um-role" class="form-input" required>${roleOptionsHtml}</select>
          </div>
          <div class="form-group">
            <label class="form-label">Status</label>
            <select id="um-status" class="form-input">
              <option value="AKTIF" ${(u.status || 'AKTIF') === 'AKTIF' ? 'selected' : ''}>Aktif</option>
              <option value="NONAKTIF" ${u.status === 'NONAKTIF' ? 'selected' : ''}>Nonaktif</option>
            </select>
          </div>
        </div>
        <div class="form-group">
          <label class="form-label">Sekolah Tempat Bekerja/Belajar</label>
          <select id="um-school" class="form-input" ${isAdminSekolah ? 'disabled' : ''}>${schoolOptionsHtml}</select>
          ${isSuperAdmin ? '<small style="color:var(--neutral-500);font-size:11px;">Biarkan kosong untuk akun Super Admin yang tidak terikat sekolah.</small>' : ''}
        </div>
        <div class="form-group">
          <label class="form-label">NIP / NISN (opsional)</label>
          <input type="text" id="um-nip" class="form-input" placeholder="NIP untuk Guru/Admin, NISN untuk Siswa" value="${u.nip || u.nisn || ''}">
        </div>
        <div style="display:flex; gap:12px; margin-top:8px; justify-content:flex-end;">
          <button type="button" id="um-cancel-btn" class="btn btn-secondary">Batal</button>
          <button type="submit" class="btn btn-primary">&#x1F4BE; Simpan Data</button>
        </div>
      </form>
    </div>
  `;
  gModal.style.display = 'flex';

  const closeModal = () => { gModal.style.display = 'none'; };
  mContent.querySelector('#um-close-btn').onclick = closeModal;
  mContent.querySelector('#um-cancel-btn').onclick = closeModal;
  gModal.onclick = (e) => { if (e.target === gModal) closeModal(); };

  mContent.querySelector('#um-form').onsubmit = (e) => {
    e.preventDefault();

    const nameVal = mContent.querySelector('#um-name').value.trim();
    const emailVal = mContent.querySelector('#um-email').value.trim();
    const passVal = mContent.querySelector('#um-pass').value.trim();
    const roleVal = mContent.querySelector('#um-role').value;
    const statusVal = mContent.querySelector('#um-status').value;
    const schoolEl = mContent.querySelector('#um-school');
    const nipEl = mContent.querySelector('#um-nip');
    const nipVal = nipEl ? nipEl.value.trim() : '';
    const schoolVal = isAdminSekolah ? activeUser.schoolId : (schoolEl ? schoolEl.value : '');

    if (!nameVal || !emailVal) { showToast('Nama dan Email wajib diisi!', 'error'); return; }
    if (!isEdit && !passVal) { showToast('Kata sandi wajib diisi untuk pengguna baru!', 'error'); return; }
    if (!schoolVal && roleVal !== 'SUPER_ADMIN') {
      showToast('Pilih sekolah terlebih dahulu!', 'error');
      return;
    }

    const updatedUser = {
      ...u,
      id: isEdit ? u.id : ('usr-' + Date.now()),
      name: nameVal,
      email: emailVal,
      role: roleVal,
      schoolId: schoolVal || null,
      status: statusVal
    };
    if (passVal) updatedUser.password = passVal;
    if (nipVal) {
      if (roleVal === 'SISWA') { updatedUser.nisn = nipVal; delete updatedUser.nip; }
      else { updatedUser.nip = nipVal; delete updatedUser.nisn; }
    }

    if (isEdit) {
      db.update('users', u.id, updatedUser);
      showToast(`Data "${nameVal}" berhasil diperbarui.`, 'success');
    } else {
      db.insert('users', updatedUser);
      showToast(`Pengguna "${nameVal}" berhasil ditambahkan.`, 'success');
    }
    closeModal();
    const mountEl = parentMount || document.getElementById('shell-content-mount');
    if (mountEl) renderKelolaUser(mountEl);
  };
}
