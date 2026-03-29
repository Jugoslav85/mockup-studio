// ─────────────────────────────────────────────────────────────
//  MOCKUP STUDIO — app.js
// ─────────────────────────────────────────────────────────────

// ── DOM ──────────────────────────────────────────────────────
const iframe           = document.getElementById("photopea");
const inPsd            = document.getElementById("in-psd");
const inImg            = document.getElementById("in-img");
const btnGen           = document.getElementById("btn-generate");
const psdStatus        = document.getElementById("psd-status");
const psdGrid          = document.getElementById("psd-grid");
const imgGrid          = document.getElementById("img-grid");
const imgCountBadge    = document.getElementById("img-count-badge");
const psdCountBadge    = document.getElementById("psd-count-badge");
const panelImg         = document.getElementById("panel-img");
const progressStrip    = document.getElementById("progress-strip");
const progressBar      = document.getElementById("progress-bar");
const progressLabel    = document.getElementById("progress-label");
const progressPct      = document.getElementById("progress-pct");
const galleryArea      = document.getElementById("gallery-area");
const resultsContainer = document.getElementById("results-container");
const btnDownloadAll   = document.getElementById("btn-download-all");
const btnHint          = document.getElementById("btn-hint");
const systemPip        = document.getElementById("system-pip");

// ── DATA ─────────────────────────────────────────────────────
let psdQueue         = [];
let imgQueue         = [];
let generatedMockups = {};
let totalExpected    = 0;
let totalGenerated   = 0;
let totalErrors      = 0;

// ── STATE ─────────────────────────────────────────────────────
let currentPsdIndex     = 0;
let photopeaLoadTimeout = null; // Photopea initial load watchdog
let retryCount          = 0;
let APP_STATE           = "IDLE";
let currentImgBase64    = null;
let batchId             = 0;  // incremented each run; stale ArrayBuffer results are dropped
let slotRetryCount      = 0;  // per-slot retry counter for timing failures
const MAX_SLOT_RETRIES  = 2;  // retry a slot up to 2 times before marking as error
const MIN_RESULT_BYTES  = 2048; // ArrayBuffers smaller than this are considered empty/corrupt

// ── PER-OPERATION TIMER OBJECT ────────────────────────────────
// Every Photopea operation gets its own cancellable context.
// Replacing currentOp via makeOp() instantly invalidates all prior callbacks.
let currentOp = null;

function makeOp() {
  // Cancel any timers still running from the previous operation
  if (currentOp) {
    if (currentOp._interval)  clearInterval(currentOp._interval);
    if (currentOp._timeout)   clearTimeout(currentOp._timeout);
    if (currentOp._heartbeat) clearInterval(currentOp._heartbeat);
    currentOp._dead = true;
  }
  const op = {
    id:         Math.random().toString(36).slice(2), // unique opId embedded in every Photopea script
    _interval:  null,
    _timeout:   null,
    _heartbeat: null,
    _dead:      false,
    _deadlineMs: 0,
    _waitDocTarget:   null,
    _waitDocCallback: null,
    _waitPollTimer:   null,
  };
  currentOp = op;
  return op;
}

function cancelOp() {
  if (currentOp) {
    if (currentOp._interval)  clearInterval(currentOp._interval);
    if (currentOp._timeout)   clearTimeout(currentOp._timeout);
    if (currentOp._heartbeat) clearInterval(currentOp._heartbeat);
    if (currentOp._waitPollTimer) clearTimeout(currentOp._waitPollTimer);
    currentOp._dead = true;
    currentOp = null;
  }
}

// Adaptive timeout — scales with PSD file size and slot count
function adaptiveTimeout(baseSec, psdIdx, slotCount) {
  const fileSizeMb = (psdQueue[psdIdx]?.file?.size || 0) / (1024 * 1024);
  const sizeBonus  = Math.min(fileSizeMb * 1500, 30000); // up to 30s extra for large files
  const slotBonus  = (slotCount || 1) * 3000;
  return baseSec * 1000 + sizeBonus + slotBonus;
}

// ── PREMIUM STATE ─────────────────────────────────────────────
// Trigger: ?premium=true in URL, or Ctrl+Shift+P to toggle, or "Simulate Payment" modal button
function isPremium() {
  return sessionStorage.getItem('ms_premium') === '1';
}

function setPremium(val) {
  if (val) sessionStorage.setItem('ms_premium', '1');
  else sessionStorage.removeItem('ms_premium');
  updatePremiumUI();
}

function updatePremiumUI() {
  const badge = document.getElementById('tier-badge');
  const premium = isPremium();
  if (badge) {
    badge.textContent  = premium ? '✦ Pro' : 'Free';
    badge.className    = 'tier-badge ' + (premium ? 'tier-pro' : 'tier-free');
    badge.title        = premium ? 'Pro — click to switch to Free' : 'Free — click to switch to Pro';
  }
  // Download All button
  if (btnDownloadAll) {
    btnDownloadAll.classList.toggle('btn-locked', !premium);
    btnDownloadAll.title = premium ? '' : 'Upgrade to Pro to download all at once';
  }
  // Re-apply locked state to any existing group download buttons
  document.querySelectorAll('.btn-small').forEach(btn => {
    btn.classList.toggle('btn-locked', !premium);
    if (!premium) btn.title = 'Upgrade to Pro to download groups';
  });
}

// ── UPGRADE MODAL ─────────────────────────────────────────────
function showUpgradeModal() {
  let d = document.getElementById('ms-upgrade');
  if (!d) {
    d = document.createElement('div');
    d.id = 'ms-upgrade';
    d.innerHTML = `
      <div class="upgrade-backdrop"></div>
      <div class="upgrade-box">
        <div class="upgrade-header">
          <div class="upgrade-logo">✦ Pro</div>
          <button class="upgrade-close" id="upgrade-close">×</button>
        </div>
        <div class="upgrade-body">
          <div class="upgrade-title">Unlock Mockup Studio Pro</div>
          <div class="upgrade-sub">You're on the free plan. Upgrade to remove watermarks and unlock bulk downloads.</div>
          <ul class="upgrade-features">
            <li><span class="uf-check">✓</span> No watermarks on exports</li>
            <li><span class="uf-check">✓</span> Download all mockups as ZIP</li>
            <li><span class="uf-check">✓</span> Download by group</li>
            <li><span class="uf-check">✓</span> Priority template access</li>
          </ul>
          <div class="upgrade-price">From <strong>$9 / month</strong> · cancel anytime</div>
        </div>
        <div class="upgrade-actions">
          <button class="upgrade-btn-cancel" id="upgrade-cancel">Maybe later</button>
          <button class="upgrade-btn-pay" id="upgrade-pay">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><rect x="1" y="4" width="22" height="16" rx="2"/><line x1="1" y1="10" x2="23" y2="10"/></svg>
            Simulate Payment
          </button>
        </div>
        <div class="upgrade-note">This is a demo — clicking "Simulate Payment" grants Pro access for this session.</div>
      </div>`;
    document.body.appendChild(d);
    d.querySelector('.upgrade-backdrop').addEventListener('click', () => d.classList.remove('show'));
    document.getElementById('upgrade-close').addEventListener('click', () => d.classList.remove('show'));
    document.getElementById('upgrade-cancel').addEventListener('click', () => d.classList.remove('show'));
    document.getElementById('upgrade-pay').addEventListener('click', () => {
      d.classList.remove('show');
      setPremium(true);
      showToast('✦ Pro unlocked — watermarks removed, bulk download enabled', 'success');
    });
    document.addEventListener('keydown', e => { if (e.key === 'Escape') d.classList.remove('show'); });
  }
  d.classList.add('show');
}

// ── WATERMARK ─────────────────────────────────────────────────
async function applyWatermark(blob) {
  return new Promise((resolve) => {
    const img = new Image();
    const url = URL.createObjectURL(blob);
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width  = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      URL.revokeObjectURL(url);

      // Watermark text
      const text     = 'MOCKUP STUDIO';
      const fontSize = Math.max(24, Math.round(Math.min(canvas.width, canvas.height) * 0.055));
      ctx.font        = `800 ${fontSize}px Syne, system-ui, sans-serif`;
      ctx.textAlign   = 'center';
      ctx.textBaseline = 'middle';

      // Shadow for readability on any background
      ctx.shadowColor   = 'rgba(0,0,0,0.55)';
      ctx.shadowBlur    = fontSize * 0.6;
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 0;

      ctx.fillStyle = 'rgba(255,255,255,0.52)';
      ctx.fillText(text, canvas.width / 2, canvas.height / 2);

      // Second pass — thin outline for contrast
      ctx.shadowBlur  = 0;
      ctx.strokeStyle = 'rgba(0,0,0,0.18)';
      ctx.lineWidth   = fontSize * 0.04;
      ctx.strokeText(text, canvas.width / 2, canvas.height / 2);

      const fmt  = window.outputFormat || 'png';
      const mime = fmt.startsWith('jpg') ? 'image/jpeg' : 'image/png';
      const q    = fmt.startsWith('jpg') ? parseFloat(fmt.split(':')[1] || '1') : undefined;
      canvas.toBlob(watermarkedBlob => resolve(watermarkedBlob), mime, q);
    };
    img.onerror = () => { URL.revokeObjectURL(url); resolve(blob); }; // fallback — return original
    img.src = url;
  });
}

// ── PSD PREVIEW ───────────────────────────────────────────────
let previewQueue   = [];
let previewRunning = false;
let psdPreviewMode = false;

const BUSY_STATES = ["RUNNING","LOADING_PSD","OPENING_SO","INJECTING","SAVING"];

// ── HELPERS ──────────────────────────────────────────────────


// ── DIAGNOSTIC LOG ────────────────────────────────────────────
// On-screen batch log — collapsed by default, shows during batch.
// Each entry: timestamp, level (info/ok/warn/error/dim), message.
const _diagStart = { time: 0 };

function diag(msg, level = 'info') {
  // Always log to console too
  const method = level === 'error' ? 'error' : level === 'warn' ? 'warn' : 'log';
  console[method]('[Diag]', msg);

  const panel = document.getElementById('diag-panel');
  const body  = document.getElementById('diag-body');
  if (!panel || !body) return;

  const elapsed = _diagStart.time ? ((Date.now() - _diagStart.time) / 1000).toFixed(1) + 's' : '0.0s';
  const line = document.createElement('div');
  line.className = `diag-line ${level}`;
  line.innerHTML = `<span class="diag-time">${elapsed}</span><span class="diag-msg">${msg}</span>`;
  body.appendChild(line);
  // Auto-scroll to bottom
  body.scrollTop = body.scrollHeight;

  // Update dot colour
  const dot = document.getElementById('diag-dot');
  if (dot) {
    if (level === 'error') dot.className = 'diag-dot error';
    else if (level === 'warn' && dot.className !== 'diag-dot error') dot.className = 'diag-dot active';
    else if (dot.className === 'diag-dot') dot.className = 'diag-dot active';
  }
}

window.diagClear = function() {
  const body = document.getElementById('diag-body');
  if (body) body.innerHTML = '';
  const dot = document.getElementById('diag-dot');
  if (dot) dot.className = 'diag-dot';
};

function diagShow() {
  const panel = document.getElementById('diag-panel');
  if (panel) panel.classList.remove('hidden');
  _diagStart.time = Date.now();
  window.diagClear();
  const dot = document.getElementById('diag-dot');
  if (dot) dot.className = 'diag-dot active';
}

function diagDone() {
  const dot = document.getElementById('diag-dot');
  if (dot) dot.className = 'diag-dot';
}

function sendMessage(payload) { iframe.contentWindow.postMessage(payload, "*"); }

// Disable pointer-events on the iframe while processing so it doesn't
// intercept scroll hit-tests and cause jank.
function setIframeBusy(busy) {
  iframe.style.pointerEvents = busy ? 'none' : '';
}

let _rafProgressPct = undefined, _rafProgressLabel = undefined, _rafProgressPending = false;
function setProgress(pct, label) {
  if (pct !== undefined) _rafProgressPct = pct;
  if (label)             _rafProgressLabel = label;
  if (_rafProgressPending) return;
  _rafProgressPending = true;
  requestAnimationFrame(() => {
    _rafProgressPending = false;
    if (_rafProgressPct !== undefined) {
      progressBar.style.width = Math.round(_rafProgressPct) + "%";
      // Show "X / Y" counter alongside percentage
      const done  = totalGenerated + totalErrors;
      const total = totalExpected || 0;
      const counter = total > 0 ? `${done} / ${total}` : '';
      if (progressPct) progressPct.textContent = counter;
    }
    if (_rafProgressLabel && progressLabel) progressLabel.textContent = _rafProgressLabel;
  });
}

function setPip(text, state) {
  systemPip.textContent = text;
  const dot = document.querySelector('.pip-dot');
  if (!dot) return;
  const c = state === 'busy'  ? '#fb923c'
          : state === 'done'  ? '#7effc4'
          : state === 'error' ? '#ff6b6b'
          : state === 'load'  ? '#a78bfa'
          : 'var(--accent)';
  dot.style.background = c;
  dot.style.boxShadow  = `0 0 8px ${c}`;
}

function revokeUrls(urls) {
  (urls || []).forEach(u => { try { if (u) URL.revokeObjectURL(u); } catch(_) {} });
}

function showToast(msg, type = 'info') {
  let t = document.getElementById('ms-toast');
  if (!t) { t = document.createElement('div'); t.id = 'ms-toast'; document.body.appendChild(t); }
  t.textContent = msg;
  t.className = `ms-toast ms-toast-${type} show`;
  clearTimeout(t._hide);
  t._hide = setTimeout(() => t.classList.remove('show'), 3800);
}

function showConfirm(title, body, confirmLabel, onConfirm) {
  let d = document.getElementById('ms-confirm');
  if (!d) {
    d = document.createElement('div');
    d.id = 'ms-confirm';
    d.innerHTML = `
      <div class="confirm-backdrop"></div>
      <div class="confirm-box">
        <div class="confirm-title" id="confirm-title"></div>
        <div class="confirm-body" id="confirm-body"></div>
        <div class="confirm-actions">
          <button class="confirm-cancel" id="confirm-cancel">Cancel</button>
          <button class="confirm-ok" id="confirm-ok"></button>
        </div>
      </div>`;
    document.body.appendChild(d);
    d.querySelector('.confirm-backdrop').addEventListener('click', () => d.classList.remove('show'));
    document.getElementById('confirm-cancel').addEventListener('click', () => d.classList.remove('show'));
    document.addEventListener('keydown', e => { if (e.key === 'Escape') d.classList.remove('show'); });
  }
  document.getElementById('confirm-title').textContent = title;
  document.getElementById('confirm-body').textContent  = body;
  const okBtn = document.getElementById('confirm-ok');
  okBtn.textContent = confirmLabel;
  okBtn.onclick = () => { d.classList.remove('show'); onConfirm(); };
  d.classList.add('show');
}

function sanitiseFilename(name) {
  return name.replace(/[/\\:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
}

function getExt() {
  return (window.outputFormat || 'png').startsWith('jpg') ? 'jpg' : 'png';
}

// Prefix / suffix naming
function applyNaming(baseName) {
  return baseName; // settings removed — no prefix/suffix for now
}

// ── GENERATE BUTTON STATE ─────────────────────────────────────

function countReadyInputs() {
  return inputs.filter(input =>
    input.length > 0 &&
    input.every(slot => slot.imgIdx !== null && slot.imgIdx !== undefined && imgQueue[slot.imgIdx])
  ).length;
}

function updateGenerateButton() {
  const hasPsd     = psdQueue.length > 0;
  const isIdle     = ["IDLE","DONE","STOPPED"].includes(APP_STATE);
  const previewing = previewRunning || previewQueue.length > 0;
  const readyInputs = countReadyInputs();
  const isDone     = APP_STATE === "DONE";

  btnGen.disabled = !(hasPsd && readyInputs > 0 && isIdle && !previewing) && !isDone;

  // Button always says "Start Batch" — completion footer handles post-batch actions
  const bc = btnGen.querySelector('.btn-content');
  if (bc && !bc.textContent.trim().includes('Start Batch')) {
    bc.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><polygon points="5,3 19,12 5,21"/></svg> Start Batch`;
  }

  if (!hasPsd && !imgQueue.length) btnHint.textContent = "Upload PSD templates and design images to begin";
  else if (!hasPsd)               btnHint.textContent = "Upload PSD templates to continue";
  else if (!imgQueue.length)      btnHint.textContent = "Upload design images to continue";
  else if (previewing)            btnHint.textContent = "Generating PSD previews, please wait…";
  else if (!inputs.length)        btnHint.textContent = "Use Auto-fill or Add input to assign images";
  else if (!readyInputs)          btnHint.textContent = "Assign images to slots — use Auto-fill to get started";
  else                            btnHint.textContent = `${readyInputs} input${readyInputs>1?'s':''} ready across ${psdQueue.length} template${psdQueue.length>1?'s':''}`;
}

// ── DRAG-TO-REORDER ───────────────────────────────────────────

function enableDragReorder(container, queue, onReorder) {
  // ── Mouse drag (desktop) ──────────────────────────────────
  let dragSrc = null;
  container.addEventListener('dragstart', e => {
    dragSrc = e.target.closest('[draggable="true"]');
    if (dragSrc) setTimeout(() => dragSrc.classList.add('dragging'), 0);
  });
  container.addEventListener('dragend', () => {
    document.querySelectorAll('.dragging,.drag-over').forEach(el => el.classList.remove('dragging','drag-over'));
    dragSrc = null;
  });
  container.addEventListener('dragover', e => {
    e.preventDefault();
    const t = e.target.closest('[draggable="true"]');
    if (t && t !== dragSrc) {
      document.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));
      t.classList.add('drag-over');
    }
  });
  container.addEventListener('drop', e => {
    e.preventDefault();
    // Don't reorder if this is an image-assign drag (destined for a slot)
    if (e.dataTransfer.getData('dragType') === 'imgAssign') return;
    const t = e.target.closest('[draggable="true"]');
    if (!t || t === dragSrc || !dragSrc) return;
    const si = +dragSrc.dataset.idx, ti = +t.dataset.idx;
    const [m] = queue.splice(si, 1);
    queue.splice(ti, 0, m);
    onReorder();
  });

  // ── Touch drag (mobile) ───────────────────────────────────
  let touchSrc = null, touchClone = null, touchSrcIdx = null;

  container.addEventListener('touchstart', e => {
    const el = e.target.closest('[draggable="true"]');
    if (!el) return;
    touchSrc    = el;
    touchSrcIdx = +el.dataset.idx;
    el.classList.add('dragging');

    // Create a visual ghost clone
    const rect  = el.getBoundingClientRect();
    touchClone  = el.cloneNode(true);
    touchClone.style.cssText = `
      position:fixed; z-index:9999; pointer-events:none; opacity:0.85;
      width:${rect.width}px; height:${rect.height}px;
      top:${rect.top}px; left:${rect.left}px;
      border-radius:8px; box-shadow:0 8px 32px rgba(0,0,0,0.5);
      transform:scale(1.04); transition:none;
    `;
    document.body.appendChild(touchClone);
  }, { passive: true });

  container.addEventListener('touchmove', e => {
    if (!touchClone) return;
    e.preventDefault();
    const touch = e.touches[0];
    const rect  = touchClone.getBoundingClientRect();
    touchClone.style.top  = (touch.clientY - rect.height / 2) + 'px';
    touchClone.style.left = (touch.clientX - rect.width  / 2) + 'px';

    // Highlight target
    touchClone.style.display = 'none';
    const elUnder = document.elementFromPoint(touch.clientX, touch.clientY);
    touchClone.style.display = '';
    const target = elUnder?.closest('[draggable="true"]');
    document.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));
    if (target && target !== touchSrc) target.classList.add('drag-over');
  }, { passive: false });

  container.addEventListener('touchend', e => {
    if (!touchSrc || !touchClone) return;
    const touch = e.changedTouches[0];

    // Clean up clone
    touchClone.remove();
    touchClone = null;
    touchSrc.classList.remove('dragging');
    document.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));

    // Find drop target
    touchClone = null;
    const elUnder = document.elementFromPoint(touch.clientX, touch.clientY);
    const target  = elUnder?.closest('[draggable="true"]');
    if (target && target !== touchSrc) {
      const ti = +target.dataset.idx;
      const [m] = queue.splice(touchSrcIdx, 1);
      queue.splice(ti, 0, m);
      onReorder();
    }

    touchSrc = null;
    touchSrcIdx = null;
  }, { passive: true });
}

// ── DRAG-TO-SCROLL ────────────────────────────────────────────

function enableDragScroll(el) {
  let down = false, startX, sl;
  el.addEventListener('mousedown', e => {
    if (e.target.closest('button,img')) return;
    down = true; startX = e.pageX - el.offsetLeft; sl = el.scrollLeft; el.style.userSelect = 'none';
  });
  el.addEventListener('mouseleave', () => { down = false; });
  el.addEventListener('mouseup',    () => { down = false; el.style.userSelect = ''; });
  el.addEventListener('mousemove',  e => {
    if (!down) return;
    e.preventDefault();
    el.scrollLeft = sl - (e.pageX - el.offsetLeft - startX) * 1.4;
  });
}

// ── LIGHTBOX ──────────────────────────────────────────────────

function openLightbox(url, name) {
  let lb = document.getElementById('ms-lightbox');
  if (!lb) {
    lb = document.createElement('div');
    lb.id = 'ms-lightbox';
    lb.innerHTML = `
      <div class="lb-backdrop"></div>
      <div class="lb-frame">
        <img id="lb-img" src="" alt="">
        <div class="lb-bar">
          <span class="lb-name" id="lb-name"></span>
          <div class="lb-actions">
            <button class="lb-btn" id="lb-copy">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>
              Copy
            </button>
            <button class="lb-btn lb-btn-dl" id="lb-dl">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7,10 12,15 17,10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Download
            </button>
            <button class="lb-close" id="lb-close">×</button>
          </div>
        </div>
      </div>`;
    document.body.appendChild(lb);
    lb.querySelector('.lb-backdrop').addEventListener('click', closeLightbox);
    document.getElementById('lb-close').addEventListener('click', closeLightbox);
    document.addEventListener('keydown', e => { if (e.key === 'Escape') closeLightbox(); });
  }
  document.getElementById('lb-img').src = url;
  document.getElementById('lb-name').textContent = `${name}.${getExt()}`;
  document.getElementById('lb-dl').onclick   = () => downloadSingle(url, name, getExt());
  document.getElementById('lb-copy').onclick = () => copyToClipboard(url);
  lb.classList.add('show');
  document.body.style.overflow = 'hidden';
}

function closeLightbox() {
  const lb = document.getElementById('ms-lightbox');
  if (lb) lb.classList.remove('show');
  document.body.style.overflow = '';
}

async function copyToClipboard(url) {
  try {
    const res  = await fetch(url);
    const blob = await res.blob();
    await navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })]);
    showToast('Image copied to clipboard', 'success');
  } catch(_) {
    showToast('Copy not supported in this browser — try downloading', 'warn');
  }
}
window.copyToClipboard = copyToClipboard;

// ── INIT ─────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', function () {
  // ── Premium init ──────────────────────────────────────────
  const premiumParam = new URLSearchParams(window.location.search).get('premium');
  if (premiumParam === 'true')  setPremium(true);
  if (premiumParam === 'false') setPremium(false);
  updatePremiumUI();

  // Click the tier badge to toggle (most reliable method)
  const tierBadge = document.getElementById('tier-badge');
  if (tierBadge) {
    tierBadge.style.cursor = 'pointer';
    tierBadge.addEventListener('click', () => {
      const next = !isPremium();
      setPremium(next);
      showToast(next ? '✦ Pro mode enabled' : 'Switched to Free plan', next ? 'success' : 'info');
    });
  }

  // Ctrl+Shift+T toggles premium (T = Tier, avoids browser-reserved combos)
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'T') {
      e.preventDefault();
      const next = !isPremium();
      setPremium(next);
      showToast(next ? '✦ Pro mode enabled' : 'Switched to Free plan', next ? 'success' : 'info');
    }
  });
  setPip("Loading…", "load");
  iframe.src = "https://www.photopea.com#" + JSON.stringify({
    environment: {
      theme:        0,     // dark theme
      localstorage: false, // don't persist state to localStorage — big RAM/speed win
      autosave:     0,     // no autosave timer
      showtools:    false, // hide tool panels — less DOM work per frame
      menus:        0,     // hide menus
      intro:        0,     // skip intro screen
    }
  });

  // If Photopea hasn't sent "done" within 20s, warn the user
  photopeaLoadTimeout = setTimeout(() => {
    if (APP_STATE === "IDLE" && systemPip.textContent === "Loading…") {
      setPip("Unavailable", "error");
      showToast('⚠️ Photopea is taking too long to load. Check your connection or try disabling ad blockers.', 'warn');
    }
  }, 20000);

  inPsd.addEventListener('change', handlePsdUpload);
  inImg.addEventListener('change', handleImagesSelect);
  btnGen.addEventListener('click', handleGenClick);
  btnDownloadAll.addEventListener('click', () => {
    if (!isPremium()) { showUpgradeModal(); return; }
    downloadAllZip();
  });
  document.getElementById("btn-stop").addEventListener('click', stopBatch);

  document.getElementById("btn-run-again").addEventListener('click', runAgain);

  // OS-level drag-and-drop on upload zones — visual feedback + actual file handling
  const dropPsd = document.getElementById('drop-psd');
  const dropImg = document.getElementById('drop-img');

  ['dragenter','dragover'].forEach(evt => {
    dropPsd.addEventListener(evt, e => { e.preventDefault(); dropPsd.classList.add('drag-active'); });
    dropImg.addEventListener(evt, e => { e.preventDefault(); dropImg.classList.add('drag-active'); });
  });
  ['dragleave'].forEach(evt => {
    dropPsd.addEventListener(evt, () => dropPsd.classList.remove('drag-active'));
    dropImg.addEventListener(evt, () => dropImg.classList.remove('drag-active'));
  });

  // PSD drop
  dropPsd.addEventListener('drop', e => {
    e.preventDefault();
    dropPsd.classList.remove('drag-active');
    const files = Array.from(e.dataTransfer.files).filter(f => f.name.toLowerCase().endsWith('.psd'));
    if (!files.length) { showToast('Please drop .psd files here', 'warn'); return; }
    const existing = new Set(psdQueue.map(f => f.name));
    const dupes = files.filter(f => existing.has(f.name));
    if (dupes.length) showToast(`⚠️ Skipped ${dupes.length} duplicate${dupes.length>1?'s':''}: ${dupes.map(d=>d.name).join(', ')}`, 'warn');
    files.forEach(f => { if (!existing.has(f.name)) psdQueue.push({ file: f, name: f.name, previewUrl: null }); });
    renderPsdThumbs();
    psdCountBadge.textContent = `${psdQueue.length} file${psdQueue.length>1?'s':''}`;
    psdCountBadge.classList.remove('hidden');
    if (psdStatus) psdStatus.textContent = `${psdQueue.length} template${psdQueue.length>1?'s':''} loaded`;
    inImg.disabled = false;
    if (APP_STATE === "DONE") { hideCompletionFooter(); showToast('Previous results are still saved below.', 'info'); }
    updateGenerateButton();
    schedulePsdPreviews();
  });

  // Image drop
  dropImg.addEventListener('drop', e => {
    e.preventDefault();
    dropImg.classList.remove('drag-active');
    const files = Array.from(e.dataTransfer.files).filter(f => f.type.startsWith('image/'));
    if (!files.length) { showToast('Please drop image files here', 'warn'); return; }
    const existing = new Set(imgQueue.map(f => f.name));
    const dupes = files.filter(f => existing.has(f.name));
    if (dupes.length) showToast(`⚠️ Skipped ${dupes.length} duplicate${dupes.length>1?'s':''}: ${dupes.map(d=>d.name).join(', ')}`, 'warn');
    const newFiles = files.filter(f => !existing.has(f.name));
    newFiles.forEach(f => addImageToQueue(f, () => {
      renderImgThumbs();
      updateGenerateButton();
    }));
    if (newFiles.length) {
      renderImgThumbs();
      renderSlotPanels();
      imgCountBadge.textContent = `${imgQueue.length} file${imgQueue.length>1?'s':''}`;
      imgCountBadge.classList.remove('hidden');
    }
    updateGenerateButton();
  });

  // Cmd/Ctrl+Enter shortcut
  document.addEventListener('keydown', e => {
    if ((e.metaKey || e.ctrlKey) && e.key === 'Enter' && !btnGen.disabled) handleGenClick();
  });

  enableDragReorder(psdGrid, psdQueue, renderPsdThumbs);
  enableDragReorder(imgGrid, imgQueue, renderImgThumbs);
  updateGenerateButton();

  // ── Library template via sessionStorage ───────────────────────
  window._pendingLibraryPsd = JSON.parse(sessionStorage.getItem('pendingTemplate') || 'null');
  if (window._pendingLibraryPsd) {
    sessionStorage.removeItem('pendingTemplate');
    showToast('Template found — waiting for Photopea to load…', 'info');
  }
});

// ── Load a template from library.json by id ───────────────────
async function loadTemplateFromLibrary(entry) {
  showToast('Loading template from library…', 'info');
  try {
    const psdRes = await fetch(entry.psd);
    if (!psdRes.ok) throw new Error(`Could not fetch ${entry.psd} (${psdRes.status})`);
    const blob = await psdRes.blob();
    const file = new File([blob], entry.id + '.psd', { type: 'application/octet-stream' });

    const existing = new Set(psdQueue.map(f => f.name));
    if (!existing.has(file.name)) {
      psdQueue.push({ file, name: file.name, previewUrl: null });
      renderPsdThumbs();
      psdCountBadge.textContent = `${psdQueue.length} file${psdQueue.length > 1 ? 's' : ''}`;
      psdCountBadge.classList.remove('hidden');
      if (psdStatus) psdStatus.textContent = `${psdQueue.length} template${psdQueue.length > 1 ? 's' : ''} loaded`;
      updateGenerateButton();
      schedulePsdPreviews();
    }
    showToast(`✓ "${entry.name}" loaded — add your images to begin`, 'success');
  } catch(err) {
    console.error('loadTemplateFromLibrary:', err);
    showToast('Failed to load template: ' + err.message, 'warn');
  }
}

function handleGenClick() {
  if (BUSY_STATES.includes(APP_STATE)) return;
  // If coming from DONE state, reset cleanly before starting a new batch.
  // This ensures Photopea gets a proper clean-slate signal (CLEANUP_DONE)
  // rather than inheriting any stale state from the previous run.
  if (APP_STATE === "DONE") {
    APP_STATE = "IDLE";
    progressStrip.classList.add("hidden");
    setProgress(0, "");
    document.getElementById('progress-strip')?.classList.remove('complete');
    setPip("Ready", "ready");
  }
  startBatch();
}

// Re-run with exactly the same files — ask before wiping existing results
function runAgain() {
  if (BUSY_STATES.includes(APP_STATE)) return;
  const hasResults = Object.values(generatedMockups).some(arr => arr.length > 0);
  if (hasResults) {
    showConfirm(
      'Run Again?',
      'This will clear all existing mockups and re-run with the same files.',
      'Clear & Run Again',
      () => _doRunAgain()
    );
  } else {
    _doRunAgain();
  }
}

function _doRunAgain() {
  APP_STATE = "IDLE";
  clearMockupsSilent();
  galleryArea.classList.add("hidden");
  btnDownloadAll.disabled = true;
  document.getElementById("btn-run-again").classList.add("hidden");
  progressStrip.classList.add("hidden");
  setProgress(0, "");
  startBatch();
}

// ── 1. PSD UPLOAD ────────────────────────────────────────────

function handlePsdUpload(e) {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  const existing = new Set(psdQueue.map(f => f.name));
  const dupes = files.filter(f => existing.has(f.name));
  if (dupes.length) showToast(`⚠️ Skipped ${dupes.length} duplicate${dupes.length>1?'s':''}: ${dupes.map(d=>d.name).join(', ')}`, 'warn');
  files.forEach(f => { if (!existing.has(f.name)) psdQueue.push({ file: f, name: f.name, previewUrl: null }); });
  renderPsdThumbs();
  psdCountBadge.textContent = `${psdQueue.length} file${psdQueue.length>1?'s':''}`;
  psdCountBadge.classList.remove("hidden");
  if (psdStatus) psdStatus.textContent = `${psdQueue.length} template${psdQueue.length>1?'s':''} loaded`;
  inImg.disabled = false;
  inPsd.value = "";
  if (APP_STATE === "DONE") { hideCompletionFooter(); showToast('Previous results are still saved below.', 'info'); }
  updateGenerateButton();
  schedulePsdPreviews();
}

function renderPsdThumbs() {
  psdGrid.innerHTML = "";
  psdQueue.forEach((item, i) => {
    const div = document.createElement("div");
    div.className = "psd-thumb-box";
    div.id = `psd-box-${i}`;
    div.draggable = true;
    div.dataset.idx = i;
    const preview = item.previewUrl
      ? `<img src="${item.previewUrl}" class="psd-preview-img" alt="preview">`
      : `<span class="psd-icon">PSD</span>`;
    div.innerHTML = `
      <div class="psd-order-badge">${i + 1}</div>
      <div class="drag-handle" title="Drag to reorder">⠿</div>
      ${preview}
      <span class="psd-name">${item.name.replace(/\.psd$/i,'')}</span>
      <button class="thumb-remove" onclick="removePsd(${i})" title="Remove">×</button>`;
    psdGrid.appendChild(div);
  });
}

// ── SLOT PANELS ──────────────────────────────────────────────
// Global input list — each input is an array of {psdIdx, soName, imgIdx}
// one slot per SO per PSD, in PSD order.
let inputs = [];  // inputs[inputIdx][slotIdx] = {psdIdx, soName, imgIdx}

// Called each time a PSD finishes scanning — adds its SO slots to all inputs
function onPsdScanComplete(psdIdx) {
  const psd = psdQueue[psdIdx];
  if (!psd?.soNames?.length) return;

  if (inputs.length === 0) {
    // First PSD — create the first empty input
    inputs = [[]];
  }

  // Append this PSD's SO slots to every existing input
  inputs.forEach(input => {
    psd.soNames.forEach(soName => {
      // Only add if not already present for this psdIdx+soName
      const exists = input.some(s => s.psdIdx === psdIdx && s.soName === soName);
      if (!exists) input.push({ psdIdx, soName, imgIdx: null });
    });
  });

  renderSlotPanels();
  updateGenerateButton();
}

// When a PSD is removed — strip its slots from all inputs, clean up empty inputs
function onPsdRemoved(psdIdx) {
  inputs.forEach(input => {
    for (let i = input.length - 1; i >= 0; i--) {
      if (input[i].psdIdx === psdIdx) input.splice(i, 1);
      else if (input[i].psdIdx > psdIdx) input[i].psdIdx--;
    }
  });
  // Remove fully empty inputs (no slots at all) — but keep at least one
  inputs = inputs.filter(input => input.length > 0);
  if (inputs.length === 0 && psdQueue.length > 0) inputs = [[]];
  renderSlotPanels();
  updateGenerateButton();
}

function renderSlotPanels() {
  const container = document.getElementById('slot-panels');
  if (!container) return;
  container.innerHTML = '';

  const scannedPsds = psdQueue.filter(p => p.soNames?.length);
  if (!scannedPsds.length) { container.classList.add('hidden'); return; }
  container.classList.remove('hidden');

  const panel = document.createElement('div');
  panel.className = 'slot-panel glass-panel';
  panel.id = 'slot-panel-main';

  // ── Top action bar ──────────────────────────────────────
  const actionBar = document.createElement('div');
  actionBar.className = 'slot-action-bar';
  actionBar.innerHTML = `
    <div class="slot-action-bar-left">
      <button class="slot-help-btn" id="btn-slot-help">
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
        How does this work?
      </button>
      <button class="slot-clear-all-btn" id="btn-clear-inputs">
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="3,6 5,6 21,6"/><path d="M19,6l-1,14a2,2,0,0,1-2,2H8a2,2,0,0,1-2-2L5,6"/></svg>
        Clear inputs
      </button>
    </div>`;
  actionBar.querySelector('#btn-slot-help').addEventListener('click', showSlotHelp);
  actionBar.querySelector('#btn-clear-inputs').addEventListener('click', () => {
    showConfirm(
      'Clear all inputs?',
      'This will remove all input columns and image assignments.',
      'Clear all',
      () => clearAllInputs()
    );
  });
  panel.appendChild(actionBar);

  // ── Table ───────────────────────────────────────────────
  const tableWrap = document.createElement('div');
  tableWrap.className = 'slot-table-wrap';

  // Left label column
  const labelCol = document.createElement('div');
  labelCol.className = 'slot-label-col';
  const cornerCell = document.createElement('div');
  cornerCell.className = 'slot-corner-cell';
  labelCol.appendChild(cornerCell);

  psdQueue.forEach((psd, psdIdx) => {
    if (!psd.soNames?.length) return;
    const psdShortName = psd.name.replace(/\.psd$/i,'');
    psd.soNames.forEach((soName, soIdx) => {
      const label = document.createElement('div');
      const isFirst = soIdx === 0;
      label.className = `slot-row-label${isFirst ? ' psd-group-first' : ''}`;
      if (isFirst) label.dataset.psdName = psdShortName;
      label.innerHTML = `
        ${psd.previewUrl
          ? `<img src="${psd.previewUrl}" class="slot-label-thumb">`
          : `<div class="slot-label-icon">PSD</div>`}
        <div class="slot-label-text">
          <div class="slot-label-psd">${psdShortName}</div>
          <div class="slot-label-so">${soName}</div>
        </div>`;
      labelCol.appendChild(label);
    });
  });
  tableWrap.appendChild(labelCol);

  // Scrollable input columns
  const inputsScroll = document.createElement('div');
  inputsScroll.className = 'slot-inputs-scroll';
  inputsScroll.id = 'slot-inputs-scroll';
  enableDragScroll(inputsScroll);

  inputs.forEach((input, inputIdx) => {
    inputsScroll.appendChild(buildInputColumn(input, inputIdx));
  });

  tableWrap.appendChild(inputsScroll);
  panel.appendChild(tableWrap);

  _renderEmptyStateHint(panel);
  container.appendChild(panel);
}

// Build a full input column DOM node
function buildInputColumn(input, inputIdx) {
  const col = document.createElement('div');
  col.className = 'slot-input-col';
  col.id = `slot-col-${inputIdx}`;

  const colHeader = document.createElement('div');
  colHeader.className = 'slot-col-header';
  colHeader.innerHTML = `
    <span class="slot-input-label">Input ${inputIdx + 1}</span>
    <button class="slot-input-remove" title="Remove column">×</button>`;
  colHeader.querySelector('.slot-input-remove').addEventListener('click', () => removeInput(inputIdx));
  col.appendChild(colHeader);

  let lastPsdIdx = -1;
  input.forEach((slot, slotIdx) => {
    const isFirstOfPsd = slot.psdIdx !== lastPsdIdx;
    lastPsdIdx = slot.psdIdx;
    col.appendChild(buildSlotCell(inputIdx, slotIdx, slot, isFirstOfPsd));
  });

  return col;
}


// ── Image picker modal ────────────────────────────────────────
// Opens when a slot is clicked — shows all uploaded images in a grid.
// User can pick one or upload a new file.
function openImagePicker(inputIdx, slotIdx) {
  // Remove existing picker if any
  document.getElementById('img-picker-modal')?.remove();

  const currentIdx = inputs[inputIdx]?.[slotIdx]?.imgIdx;

  const backdrop = document.createElement('div');
  backdrop.className = 'img-picker-backdrop';
  backdrop.id = 'img-picker-modal';

  let selectedIdx = currentIdx ?? null;

  function buildGrid() {
    const items = imgQueue.map((item, i) => {
      const div = document.createElement('div');
      div.className = 'img-picker-item' + (i === selectedIdx ? ' selected' : '');
      div.innerHTML = `<img src="${item.url}" loading="lazy"><div class="img-picker-item-name">${item.name}</div>`;
      div.addEventListener('click', () => {
        selectedIdx = i;
        backdrop.querySelectorAll('.img-picker-item').forEach(el => el.classList.remove('selected'));
        div.classList.add('selected');
        confirmBtn.disabled = false;
      });
      return div;
    });
    grid.innerHTML = '';
    if (items.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'img-picker-empty';
      empty.textContent = 'No images uploaded yet — use the button below to add one.';
      grid.appendChild(empty);
    } else {
      items.forEach(el => grid.appendChild(el));
    }
  }

  backdrop.innerHTML = `
    <div class="img-picker-box">
      <div class="img-picker-header">
        <div class="img-picker-title">Choose image for this slot</div>
        <button class="img-picker-close">✕</button>
      </div>
      <div class="img-picker-grid" id="img-picker-grid"></div>
      <div class="img-picker-footer">
        <button class="img-picker-upload-btn">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17,8 12,3 7,8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
          Upload new image
        </button>
        <button class="img-picker-confirm-btn" ${selectedIdx === null ? 'disabled' : ''}>
          Assign →
        </button>
      </div>
    </div>`;

  document.body.appendChild(backdrop);

  const grid = backdrop.querySelector('.img-picker-grid');
  const confirmBtn = backdrop.querySelector('.img-picker-confirm-btn');

  buildGrid();

  // Close on backdrop click
  backdrop.addEventListener('click', e => { if (e.target === backdrop) backdrop.remove(); });
  backdrop.querySelector('.img-picker-close').addEventListener('click', () => backdrop.remove());
  document.addEventListener('keydown', function escClose(e) {
    if (e.key === 'Escape') { backdrop.remove(); document.removeEventListener('keydown', escClose); }
  });

  // Confirm
  confirmBtn.addEventListener('click', () => {
    if (selectedIdx === null) return;
    inputs[inputIdx][slotIdx].imgIdx = selectedIdx;
    updateSlotCell(inputIdx, slotIdx);
    updateGenerateButton();
    backdrop.remove();
  });

  // Upload new
  backdrop.querySelector('.img-picker-upload-btn').addEventListener('click', () => {
    const fp = document.createElement('input');
    fp.type = 'file'; fp.accept = 'image/*';
    fp.onchange = ev => {
      const f = ev.target.files[0]; if (!f) return;
      const existing = imgQueue.find(q => q.name === f.name);
      let idx;
      if (existing) {
        idx = imgQueue.indexOf(existing);
      } else {
        addImageToQueue(f, () => {
          buildGrid(); // refresh grid once resize is done
          updateGenerateButton();
        });
        idx = imgQueue.length - 1;
        imgCountBadge.textContent = `${imgQueue.length} file${imgQueue.length>1?'s':''}`;
        imgCountBadge.classList.remove('hidden');
        renderImgThumbs();
      }
      // Pre-select the newly added image
      selectedIdx = idx;
      buildGrid();
      confirmBtn.disabled = false;
    };
    fp.click();
  });
}

// Build a single slot cell DOM node
function buildSlotCell(inputIdx, slotIdx, slot, isFirstOfPsd) {
  const img = slot.imgIdx !== null && slot.imgIdx !== undefined ? imgQueue[slot.imgIdx] : null;
  const zone = document.createElement('div');
  zone.className = `slot-drop-zone${img ? ' slot-filled' : ''}${isFirstOfPsd ? ' psd-group-first' : ''}`;
  zone.id = `slot-cell-${inputIdx}-${slotIdx}`;

  if (img) {
    zone.innerHTML = `
      <img src="${img.url}" class="slot-assigned-img">
      <button class="slot-clear-btn" title="Clear">×</button>`;
    zone.querySelector('.slot-clear-btn').addEventListener('click', e => {
      e.stopPropagation();
      inputs[inputIdx][slotIdx].imgIdx = null;
      updateSlotCell(inputIdx, slotIdx);
      updateGenerateButton();
    });
    zone.addEventListener('click', e => {
      if (e.target.closest('.slot-clear-btn')) return; // clear btn handled above
      openImagePicker(inputIdx, slotIdx);
    });
  } else {
    zone.innerHTML = `
      <div class="slot-empty-content">
        <svg class="slot-drop-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21,15 16,10 5,21"/></svg>
        <div class="slot-drop-hint">Click or drop image</div>
      </div>`;
    zone.addEventListener('click', () => openImagePicker(inputIdx, slotIdx));
  }

  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    // Reject reorder drags — only accept image-assign drags
    if (e.dataTransfer.getData('dragType') !== 'imgAssign') return;
    const imgIdx = parseInt(e.dataTransfer.getData('imgIdx'));
    if (!isNaN(imgIdx)) {
      inputs[inputIdx][slotIdx].imgIdx = imgIdx;
      updateSlotCell(inputIdx, slotIdx);
      updateGenerateButton();
    }
  });

  return zone;
}

// ── Targeted cell update — only replaces the one changed cell ──
function updateSlotCell(inputIdx, slotIdx) {
  const slot = inputs[inputIdx]?.[slotIdx];
  if (!slot) return;
  const existing = document.getElementById(`slot-cell-${inputIdx}-${slotIdx}`);
  if (!existing) { renderSlotPanels(); return; } // fallback if DOM desynced
  const isFirstOfPsd = slotIdx === 0 || inputs[inputIdx][slotIdx - 1]?.psdIdx !== slot.psdIdx;
  const newCell = buildSlotCell(inputIdx, slotIdx, slot, isFirstOfPsd);
  existing.replaceWith(newCell);
  const panel = document.getElementById('slot-panel-main');
  if (panel) _renderEmptyStateHint(panel);
}

// ── Empty state hint ───────────────────────────────────────────
function _renderEmptyStateHint(panel) {
  const existing = panel.querySelector('.slot-first-time-hint');
  if (existing) existing.remove();
  const anyFilled = inputs.some(input => input.some(slot => slot.imgIdx !== null && slot.imgIdx !== undefined));
  if (!anyFilled) {
    const emptyState = document.createElement('div');
    emptyState.className = 'slot-first-time-hint';
    emptyState.innerHTML = `
      <div class="slot-fth-inner">
        <div class="slot-fth-icon">
          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4"><polygon points="13,2 3,14 12,14 11,22 21,10 12,10 13,2"/></svg>
        </div>
        <div class="slot-fth-text">
          <strong>Press Auto-fill to get started</strong>
          <span>Auto-fill assigns your uploaded images to templates automatically. You can also drag images from the panel above into any slot individually.</span>
        </div>
        <button class="slot-fth-btn" id="btn-fth-autofill">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><polygon points="13,2 3,14 12,14 11,22 21,10 12,10 13,2"/></svg>
          Auto-fill now
        </button>
      </div>`;
    emptyState.querySelector('#btn-fth-autofill').addEventListener('click', autoFillInputs);
    panel.appendChild(emptyState);
  }
}

function addInput() {
  // If no inputs exist yet (e.g. after clearing with no PSDs scanned), do nothing silently
  if (!inputs.length) {
    showToast('Upload and scan a PSD template first', 'warn');
    return;
  }
  const newInput = inputs[0].map(slot => ({ ...slot, imgIdx: null }));
  inputs.push(newInput);
  // Targeted append — no full rebuild
  const scroll = document.getElementById('slot-inputs-scroll');
  if (scroll) {
    scroll.appendChild(buildInputColumn(newInput, inputs.length - 1));
    const panel = document.getElementById('slot-panel-main');
    if (panel) _renderEmptyStateHint(panel);
  } else {
    renderSlotPanels();
  }
  updateGenerateButton();
}

function removeInput(inputIdx) {
  // Allow removing all columns — structural change needs full rebuild
  inputs.splice(inputIdx, 1);
  renderSlotPanels();
  updateGenerateButton();
}

function clearAllInputs() {
  // Reset to one empty input column — keeps the table visible with slots to fill
  if (inputs.length > 0) {
    inputs = [inputs[0].map(slot => ({ ...slot, imgIdx: null }))];
  } else {
    inputs = [];
  }
  renderSlotPanels();
  updateGenerateButton();
  showToast('All inputs cleared', 'info');
}

// Auto-fill logic:
// - Single-SO PSDs: one input per image, that image fills the slot
// - Multi-SO PSDs: images grouped by SO count, each group fills one input column
//   e.g. 3-SO PSD + 6 images → Input 1 gets images 1+2+3, Input 2 gets images 4+5+6
// - Mixed: number of inputs = max groups needed across all PSDs
//   Single-SO PSDs fill their slot with the first image of each group
function autoFillInputs() {
  if (!imgQueue.length) { showToast('Upload images first', 'warn'); return; }

  const scanned = psdQueue.filter(p => p.soNames?.length);
  if (!scanned.length) return;

  // Max SOs across all PSDs — determines group size
  const maxSos = Math.max(...scanned.map(p => p.soNames.length));

  // Number of inputs = ceil(imgQueue.length / maxSos)
  const numInputs = Math.ceil(imgQueue.length / maxSos);

  inputs = [];

  for (let inputIdx = 0; inputIdx < numInputs; inputIdx++) {
    const input = [];
    psdQueue.forEach((psd, psdIdx) => {
      if (!psd.soNames?.length) return;
      const soCount = psd.soNames.length;
      psd.soNames.forEach((soName, soIdx) => {
        let imgIdx = null;
        if (soCount === 1) {
          // Single-SO: use first image of this group
          const groupStart = inputIdx * maxSos;
          imgIdx = groupStart < imgQueue.length ? groupStart : null;
        } else {
          // Multi-SO: distribute images across SO slots
          const groupStart = inputIdx * soCount;
          const candidate = groupStart + soIdx;
          imgIdx = candidate < imgQueue.length ? candidate : null;
        }
        input.push({ psdIdx, soName, imgIdx });
      });
    });
    inputs.push(input);
  }

  renderSlotPanels();
  updateGenerateButton();
  showToast(`✓ Created ${inputs.length} input${inputs.length > 1 ? 's' : ''}`, 'success');
}

function showSlotHelp() {
  let d = document.getElementById('slot-help-modal');
  if (!d) {
    d = document.createElement('div');
    d.id = 'slot-help-modal';
    d.innerHTML = `
      <div class="slot-help-backdrop"></div>
      <div class="slot-help-box">
        <div class="slot-help-header">
          <div class="slot-help-title">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
            How Mockup Inputs Work
          </div>
          <button class="slot-help-close">×</button>
        </div>
        <div class="slot-help-body">
          <div class="slot-help-item">
            <div class="slot-help-icon">1</div>
            <div>
              <strong>Each column is one output mockup.</strong>
              When you start the batch, each input column generates one image per PSD template.
            </div>
          </div>
          <div class="slot-help-item">
            <div class="slot-help-icon">2</div>
            <div>
              <strong>Drag or click to assign images.</strong>
              Drag any image from the queue above into a slot, or click an empty slot to upload directly.
            </div>
          </div>
          <div class="slot-help-item">
            <div class="slot-help-icon">3</div>
            <div>
              <strong>⚡ Auto-fill handles the mapping for you.</strong>
              For single-design templates, one input is created per image.
              For multi-design templates (e.g. a set of 3 posters), images are grouped — images 1-3 fill Input 1, images 4-6 fill Input 2, and so on.
            </div>
          </div>
          <div class="slot-help-item">
            <div class="slot-help-icon">4</div>
            <div>
              <strong>The SO layer name tells you which slot is which.</strong>
              Layer names like "Right", "Left", "Middle" come directly from your PSD file.
            </div>
          </div>
        </div>
      </div>`;
    document.body.appendChild(d);
    d.querySelector('.slot-help-backdrop').addEventListener('click', () => d.classList.remove('show'));
    d.querySelector('.slot-help-close').addEventListener('click', () => d.classList.remove('show'));
    document.addEventListener('keydown', e => { if (e.key === 'Escape') d.classList.remove('show'); });
  }
  d.classList.add('show');
}

window.removePsd = function(i) {
  revokeUrls([psdQueue[i]?.previewUrl]);
  psdQueue.splice(i, 1);
  renderPsdThumbs();
  onPsdRemoved(i);
  const n = psdQueue.length;
  if (n) {
    psdCountBadge.textContent = `${n} file${n>1?'s':''}`;
  } else {
    psdCountBadge.classList.add("hidden");
    if (psdStatus) psdStatus.textContent = "Waiting…";
    // image panel stays enabled — user can upload in any order
  }
  updateGenerateButton();
};

// ── PSD PREVIEWS ─────────────────────────────────────────────

function schedulePsdPreviews() {
  psdQueue.forEach((item, i) => { if (!item.previewUrl && !previewQueue.includes(i)) previewQueue.push(i); });
  if (!previewRunning) runNextPreview();
}

function runNextPreview() {
  if (BUSY_STATES.includes(APP_STATE)) { previewRunning = false; updateGenerateButton(); return; }
  const idx = previewQueue.shift();
  if (idx === undefined) { previewRunning = false; updateGenerateButton(); return; }
  if (!psdQueue[idx]) { runNextPreview(); return; }
  previewRunning   = true;
  psdPreviewMode   = true;
  window._previewIdx = idx;
  // Mark thumbnail as loading
  const box = document.getElementById(`psd-box-${idx}`);
  if (box) box.classList.add('previewing');
  sendMessage(`try{while(app.documents.length>0){app.activeDocument=app.documents[0];app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}app.echoToOE(JSON.stringify({type:"PREVIEW_READY"}));}catch(e){app.echoToOE(JSON.stringify({type:"PREVIEW_READY"}));}`);
}

// ── IMAGE RESIZE ON UPLOAD ───────────────────────────────────
// Resize once at upload time; store base64 on the item so injection
// is instant — no async canvas work during batch processing.
const IMG_MAX_PX = 2000;

function addImageToQueue(file, onDone) {
  const item = {
    file,
    name:             file.name,
    url:              URL.createObjectURL(file),
    fitMode:          window.currentFitMode || 'fill',
    fitModeOverridden: false,
    base64:           null,   // filled async below
    _resizing:        true,
  };
  imgQueue.push(item);

  const objectURL = URL.createObjectURL(file);
  const tempImg   = new Image();
  tempImg.onload = () => {
    URL.revokeObjectURL(objectURL);
    const { naturalWidth: w, naturalHeight: h } = tempImg;
    const scale   = (w > IMG_MAX_PX || h > IMG_MAX_PX) ? IMG_MAX_PX / Math.max(w, h) : 1;
    const canvas  = document.createElement('canvas');
    canvas.width  = Math.round(w * scale);
    canvas.height = Math.round(h * scale);
    canvas.getContext('2d').drawImage(tempImg, 0, 0, canvas.width, canvas.height);
    const hasAlpha = file.type === 'image/png' || file.type === 'image/webp';
    // Strip any whitespace/newlines — some browsers insert them in long base64 strings,
    // which breaks Photopea's script parser when the data URL is embedded in a JS string.
    item.base64    = canvas.toDataURL(hasAlpha ? 'image/png' : 'image/jpeg', 0.95).replace(/\s/g, '');
    item._resizing = false;
    const origKB   = Math.round(file.size / 1024);
    const newKB    = Math.round(item.base64.length * 0.75 / 1024);
    console.log(`[IMG] ${file.name}: ${origKB}KB → ${newKB}KB (${canvas.width}×${canvas.height})`);
    if (onDone) onDone(item);
  };
  tempImg.onerror = () => {
    URL.revokeObjectURL(objectURL);
    // Fallback: read as-is
    const reader = new FileReader();
    reader.onload = ev => {
      item.base64    = ev.target.result.replace(/\s/g, '');
      item._resizing = false;
      if (onDone) onDone(item);
    };
    reader.readAsDataURL(file);
  };
  tempImg.src = objectURL;
  return item;
}

// ── 2. IMAGE UPLOAD ──────────────────────────────────────────

function handleImagesSelect(e) {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  const existing = new Set(imgQueue.map(f => f.name));
  const dupes = files.filter(f => existing.has(f.name));
  if (dupes.length) showToast(`⚠️ Skipped ${dupes.length} duplicate${dupes.length>1?'s':''}: ${dupes.map(d=>d.name).join(', ')}`, 'warn');
  const newFiles = files.filter(f => !existing.has(f.name));
  if (newFiles.length && APP_STATE === "DONE") {
    showToast('Previous results are still saved below — scroll down to download them.', 'info');
    hideCompletionFooter();
  }
  newFiles.forEach(f => addImageToQueue(f, () => {
    renderImgThumbs();
    updateGenerateButton();
  }));
  if (newFiles.length) {
    renderImgThumbs();   // render immediately (shows spinner/placeholder)
    renderSlotPanels();
    imgCountBadge.textContent = `${imgQueue.length} file${imgQueue.length>1?'s':''}`;
    imgCountBadge.classList.remove("hidden");
  }
  inImg.value = "";
  updateGenerateButton();
}

function renderImgThumbs() {
  imgGrid.innerHTML = "";
  imgQueue.forEach((item, i) => {
    const wrap = document.createElement("div");
    wrap.className = "img-thumb-wrap";
    wrap.draggable = true;
    wrap.dataset.idx = i;
    const pills = ['fit','fill','stretch'].map(m =>
      `<button class="fit-pill${item.fitMode===m?' active':''}" data-idx="${i}" data-fit="${m}" onclick="setImgFit(${i},'${m}')">${m}</button>`
    ).join('');
    wrap.innerHTML = `
      <div class="drag-handle" title="Drag to reorder">⠿</div>
      <img src="${item.url}" class="mini-thumb" title="${item.name}" loading="lazy">
      <div class="fit-pills">${pills}</div>
      <button class="thumb-remove" onclick="removeImg(${i})" title="Remove">×</button>`;
    // Allow dragging to slot panels
    wrap.addEventListener('dragstart', ev => {
      ev.dataTransfer.setData('imgIdx', String(i));
      ev.dataTransfer.setData('dragType', 'imgAssign');
      ev.dataTransfer.effectAllowed = 'copyMove';
    });
    imgGrid.appendChild(wrap);
  });
}

window.setImgFit = function(i, mode) {
  if (!imgQueue[i]) return;
  imgQueue[i].fitMode          = mode;
  imgQueue[i].fitModeOverridden = true;
  document.querySelectorAll(`.fit-pill[data-idx="${i}"]`).forEach(p => p.classList.toggle('active', p.dataset.fit === mode));
};

window.removeImg = function(i) {
  revokeUrls([imgQueue[i]?.url]);
  imgQueue.splice(i, 1);
  // Fix stale imgIdx references in global inputs
  inputs.forEach(input => {
    input.forEach(slot => {
      if (slot.imgIdx === i) slot.imgIdx = null;
      else if (slot.imgIdx !== null && slot.imgIdx > i) slot.imgIdx--;
    });
  });
  renderImgThumbs();
  renderSlotPanels();
  const n = imgQueue.length;
  imgCountBadge.textContent = n ? `${n} file${n>1?'s':''}` : '';
  if (!n) imgCountBadge.classList.add("hidden");
  updateGenerateButton();
};

// ── 3. BATCH ─────────────────────────────────────────────────

// Flattened list of {psdIdx, inputIdx, input} built at batch start
let batchPlan = [];
let batchPlanIndex = 0;
// Within one input, we process slots sequentially
let currentSlotIndex = 0;

function buildInputKey(inputIdx) {
  const input = inputs[inputIdx];
  if (!input?.length) return `Input ${inputIdx + 1}`;
  // Collect unique image names used in this input
  const imgNames = [...new Set(
    input
      .filter(s => s.imgIdx !== null && s.imgIdx !== undefined && imgQueue[s.imgIdx])
      .map(s => imgQueue[s.imgIdx].name.replace(/\.[^.]+$/, ''))  // strip extension
  )];
  // Collect unique PSD names
  const psdNames = [...new Set(
    input
      .filter(s => psdQueue[s.psdIdx])
      .map(s => psdQueue[s.psdIdx].name.replace(/\.psd$/i, ''))
  )];
  if (!imgNames.length) return `Input ${inputIdx + 1}`;
  return `${imgNames.join(' + ')} — ${psdNames.join(', ')}`;
}

function startBatch() {
  if (!psdQueue.length || !inputs.length) return;

  // PSD-first batch plan: for each PSD, for each ready input, collect its slots for that PSD
  const readyInputIdxs = inputs
    .map((input, idx) => ({ input, idx }))
    .filter(({ input }) =>
      input.length > 0 &&
      input.every(s => s.imgIdx !== null && s.imgIdx !== undefined && imgQueue[s.imgIdx])
    );

  if (!readyInputIdxs.length) return;

  batchPlan = [];
  psdQueue.forEach((psd, psdIdx) => {
    if (!psd.soNames?.length) return;
    readyInputIdxs.forEach(({ input, idx: inputIdx }) => {
      const slots = input.filter(s => s.psdIdx === psdIdx);
      if (slots.length) batchPlan.push({ psdIdx, inputIdx, slots });
    });
  });

  if (!batchPlan.length) return;

  document.getElementById('completion-footer')?.classList.add('hidden');
  clearMockupsSilent();
  batchId++;
  totalExpected  = batchPlan.length;
  totalGenerated = 0;
  totalErrors    = 0;

  readyInputIdxs.forEach(({ idx }) => {
    const key = buildInputKey(idx);
    if (!generatedMockups[key]) generatedMockups[key] = [];
  });

  btnGen.disabled = true;
  inPsd.disabled  = true;
  inImg.disabled  = true;
  galleryArea.classList.remove("hidden");
  progressStrip.classList.remove("hidden");
  btnDownloadAll.disabled = true;
  document.getElementById("btn-run-again").classList.add("hidden");

  setPip("Processing", "busy");
  document.getElementById("btn-stop").classList.remove("hidden");
  document.getElementById('progress-strip')?.classList.remove('complete');
  // Grey out controls that are irrelevant during processing
  document.getElementById('btn-ctrl-autofill')?.setAttribute('disabled', '');
  document.getElementById('btn-ctrl-add')?.setAttribute('disabled', '');
  document.querySelectorAll('.btn-clear-panel').forEach(b => b.setAttribute('disabled', ''));
  document.querySelectorAll('.seg-btn[data-fmt]').forEach(b => b.setAttribute('disabled', ''));
  document.querySelectorAll('#fit-control .seg-btn').forEach(b => b.setAttribute('disabled', ''));
  document.getElementById('jpg-quality')?.setAttribute('disabled', '');
  document.getElementById('drop-psd')?.classList.add('drop-disabled');
  document.getElementById('drop-img')?.classList.add('drop-disabled');

  // ── Abort any in-flight preview cleanly ──────────────────
  // Clear preview queue so no further previews start
  previewQueue = [];
  previewRunning = false;
  psdPreviewMode = false;

  APP_STATE        = "RUNNING";
  batchPlanIndex   = 0;
  currentPsdIndex  = 0;
  currentSlotIndex = 0;

  // Small delay to let any in-flight Photopea preview scripts finish
  // before we send our first batch cleanup command
  setIframeBusy(true);
  document.querySelector('.upload-row')?.classList.add('panels-collapsed');
  document.getElementById('slot-panels')?.classList.add('panels-collapsed');
  document.querySelector('.controls-bar')?.classList.add('panels-collapsed');
  hideCompletionFooter();
  diagShow();
  diag(`Batch started — ${batchPlan.length} item${batchPlan.length!==1?'s':''} across ${psdQueue.length} PSD${psdQueue.length!==1?'s':''}`, 'info');
  // Show proactive "stay on this tab" notice at batch start
  (function() {
    const el = document.getElementById('stay-notice');
    if (el) el.classList.remove('hidden');
  })();
  setTimeout(loadNextPlanItem, 800);
}


function loadNextPlanItem() {
  if (APP_STATE === "STOPPED") return;
  if (batchPlanIndex >= batchPlan.length) { finishAll(); return; }

  const plan = batchPlan[batchPlanIndex];
  const psd  = psdQueue[plan.psdIdx];
  currentPsdIndex  = plan.psdIdx;
  currentSlotIndex = 0;

  setProgress((batchPlanIndex / totalExpected) * 100, `Loading: ${psd.name} — Input ${plan.inputIdx + 1}`);
  APP_STATE = "LOADING_PSD";
  diag(`─── Item ${batchPlanIndex + 1}/${totalExpected}: ${psd.name} × Input ${plan.inputIdx + 1}`, 'info');

  document.querySelectorAll(".psd-thumb-box").forEach(el => el.classList.remove("active"));
  const box = document.getElementById(`psd-box-${plan.psdIdx}`);
  if (box) box.classList.add("active");

  sendMessage(`try{while(app.documents.length>0){app.activeDocument=app.documents[0];app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}app.echoToOE(JSON.stringify({type:"CLEANUP_DONE"}));}catch(e){app.echoToOE(JSON.stringify({type:"CLEANUP_DONE"}));}`);
}

function uploadPsdFile() {
  const reader = new FileReader();
  reader.onload = function() { sendMessage(this.result); };
  reader.readAsArrayBuffer(psdQueue[currentPsdIndex].file);
}

function scanPsd() {
  // The slot system already knows the SO names from the preview scan.
  // We just need to confirm the PSD document is ready in Photopea,
  // apply placeholder hiding once, then start the slot loop.
  setProgress(undefined, "Preparing layers…");
  retryCount = 0;

  const op        = makeOp();
  const plan      = batchPlan[batchPlanIndex];
  const timeoutMs = adaptiveTimeout(30, currentPsdIndex, plan?.slots?.length || 1);

  // Run placeholder hiding and confirm doc is ready — no SO scan needed
  sendMessage(`try{
    app.displayDialogs=DialogModes.NO;
    var doc=app.activeDocument;
    function hidePH(p){for(var i=0;i<p.layers.length;i++){var l=p.layers[i];var n=l.name.toLowerCase();if(n.indexOf("placeholder")!==-1||n.indexOf("delete")!==-1||n.indexOf("preview")!==-1||n.indexOf("your design")!==-1||n.indexOf("replace me")!==-1||n.indexOf("remove")!==-1||n.indexOf("instruction")!==-1||n.indexOf("guide")!==-1||n.indexOf("watermark")!==-1||n.indexOf("promo")!==-1){try{l.visible=false;}catch(e2){}}}}
    hidePH(doc);
    app.echoToOE(JSON.stringify({type:"SCAN_OK",opId:"${op.id}"}));
  }catch(e){app.echoToOE(JSON.stringify({type:"SCAN_OK",opId:"${op.id}"}));}`);

  // Adaptive timeout: if we never hear back, skip this PSD
  op._timeout = setTimeout(() => {
    if (op._dead || APP_STATE === "STOPPED") return;
    op._dead = true;
    showToast(`⚠️ Skipped "${psdQueue[currentPsdIndex]?.name}" — PSD did not respond`, 'warn');
    const box = document.getElementById(`psd-box-${currentPsdIndex}`);
    if (box) { box.classList.remove("active"); box.classList.add("done"); }
    batchPlanIndex++;
    loadNextPlanItem();
  }, timeoutMs);
}

function startSlotLoop() {
  // Called after PSD is loaded and initial scan confirms SOs are present
  currentSlotIndex = 0;
  processNextSlot();
}

function processNextSlot(isRetry) {
  if (APP_STATE === "STOPPED") return;
  const plan  = batchPlan[batchPlanIndex];
  const slots = plan.slots;

  if (currentSlotIndex >= slots.length) {
    // All slots filled — export the final composite
    exportFinalMockup();
    return;
  }

  // Reset retry counter when moving to a new slot (not a retry)
  if (!isRetry) slotRetryCount = 0;

  const slot = slots[currentSlotIndex];
  const img  = imgQueue[slot.imgIdx];
  if (!img) { currentSlotIndex++; processNextSlot(); return; }

  const done = batchPlanIndex;
  setProgress(
    (done / totalExpected) * 100,
    `Input ${plan.inputIdx + 1} — "${slot.soName}" → ${psdQueue[plan.psdIdx]?.name || ''}`
  );

  // base64 was prepared at upload time by addImageToQueue.
  // If still resizing (edge case: batch started immediately after upload),
  // wait briefly and retry.
  if (img._resizing || !img.base64) {
    setTimeout(processNextSlot, 150);
    return;
  }
  currentImgBase64 = img.base64;
  selectAndOpenSlot(slot.soName);
}

// ── Wait for Photopea doc count to reach target before proceeding ──
// Uses a closure-captured operation so concurrent calls can never overwrite each other.
// Polls up to 40 times (~8 seconds) then proceeds to avoid deadlock.
function waitForDocCount(target, callback) {
  if (APP_STATE === "STOPPED") return;
  const capturedOp = currentOp;
  let attempts = 0;

  function poll() {
    if (APP_STATE === "STOPPED") return;
    if (capturedOp && capturedOp._dead) return; // this op was superseded — bail silently
    if (attempts >= 40) { callback(); return; }  // safety valve
    attempts++;
    const opTag = capturedOp ? capturedOp.id : '';
    sendMessage(`app.echoToOE(JSON.stringify({type:"POLL_DOCS",count:app.documents.length,opId:"${opTag}"}));`);
    const t = setTimeout(poll, 200);
    if (capturedOp) capturedOp._waitPollTimer = t;
  }

  if (capturedOp) {
    capturedOp._waitDocTarget   = target;
    capturedOp._waitDocCallback = callback;
    capturedOp._waitPollTimer   = null;
  }
  poll();
}

// Called from the message router when POLL_DOCS arrives for the current op
function handleDocCountWait(count, op) {
  if (!op || op._dead) return;
  if (op._waitPollTimer) { clearTimeout(op._waitPollTimer); op._waitPollTimer = null; }
  if (count <= op._waitDocTarget) {
    const cb = op._waitDocCallback;
    op._waitDocTarget = op._waitDocCallback = null;
    if (cb) cb();
  }
  // If target not yet reached, the poll() closure's own setTimeout will fire again
}

function selectAndOpenSlot(soName) {
  APP_STATE = "OPENING_SO";
  const escaped = soName.replace(/\\/g,'\\\\').replace(/"/g,'\\"');
  const plan    = batchPlan[batchPlanIndex];
  const op      = makeOp();

  // Use a very generous flat timeout for SO open — executeAction() BLOCKS Photopea
  // so heartbeat pings go unanswered while it executes. The adaptive timeout alone.
  // Any POLL_DOCS reply resets the deadline, acting as a liveness signal.
  const fileSizeMb  = (psdQueue[currentPsdIndex]?.file?.size || 0) / (1024 * 1024);
  const sizeBonus   = Math.min(fileSizeMb * 2000, 60000); // up to 60s extra for big files
  const flatTimeout = 60000 + sizeBonus; // 60s base — generous for slow PSDs
  op._deadlineMs    = Date.now() + flatTimeout;

  diag(`Opening SO: "${soName}"… (timeout: ${Math.round(flatTimeout/1000)}s)`, 'info');
  const _soOpenStart = Date.now();

  // No heartbeat during SO open — Photopea is blocked by executeAction.
  // Poll doc count every 800ms instead; each reply extends deadline.
  let _pollCount = 0;
  op._interval = setInterval(() => {
    if (op._dead) return;
    _pollCount++;
    if (_pollCount % 5 === 0) { // log every ~4s
      const elapsed = ((Date.now() - _soOpenStart) / 1000).toFixed(1);
      diag(`  SO waiting… ${elapsed}s elapsed, deadline in ${Math.round((op._deadlineMs - Date.now())/1000)}s`, 'dim');
    }
    sendMessage(`app.echoToOE(JSON.stringify({type:"POLL_DOCS",count:app.documents.length,opId:"${op.id}"}));`);
  }, 800);

  // Deadline checker — each POLL_DOCS reply that arrives extends deadline by 10s
  // (handled in message router POLL_DOCS case). Fires every 5s.
  function checkDeadline() {
    if (op._dead) return;
    if (Date.now() < op._deadlineMs) { op._timeout = setTimeout(checkDeadline, 5000); return; }
    op._dead = true;
    diag(`  ⚠️ SO open timeout after ${Math.round((Date.now() - (_soOpenStart||Date.now()))/1000)}s`, 'warn');
    if (slotRetryCount < MAX_SLOT_RETRIES) {
      slotRetryCount++;
      diag(`  Retrying (${slotRetryCount}/${MAX_SLOT_RETRIES})…`, 'warn');
      showToast(`⚠️ SO open timeout — retrying (${slotRetryCount}/${MAX_SLOT_RETRIES})`, 'warn');
      sendMessage(`try{while(app.documents.length>0){app.activeDocument=app.documents[0];app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}}catch(e){}`);
      setTimeout(() => { if (APP_STATE !== "STOPPED") loadNextPlanItem(); }, 1500);
    } else {
      addErrorToGallery('SO open timeout after retries');
    }
  }
  op._timeout = setTimeout(checkDeadline, 5000);

  // Select the named SO layer then open it
  const script = `
    try{
      app.displayDialogs=DialogModes.NO;
      var doc=app.activeDocument;
      function findByName(p,n){for(var i=0;i<p.layers.length;i++){var l=p.layers[i];if(l.name===n)return l;if(l.layers&&l.layers.length>0){var f=findByName(l,n);if(f)return f;}}return null;}
      var target=findByName(doc,"${escaped}");
      if(!target){
        function findFirstSO(p){if(!p||!p.layers)return null;for(var i=0;i<p.layers.length;i++){var l=p.layers[i];if(l.visible!==false&&l.kind==LayerKind.SMARTOBJECT)return l;if(l.layers&&l.layers.length>0){var f=findFirstSO(l);if(f)return f;}}return null;}
        target=findFirstSO(doc);
      }
      if(!target)throw "SO not found: ${escaped}";
      doc.activeLayer=target;
      var desc=new ActionDescriptor();
      executeAction(stringIDToTypeID("placedLayerEditContents"),desc,3);
      app.echoToOE(JSON.stringify({type:"SO_OPEN_CMD",opId:"${op.id}"}));
    }catch(e){app.echoToOE(JSON.stringify({type:"ERROR",msg:e.toString(),opId:"${op.id}"}));}
  `;
  sendMessage(script);
}

function exportFinalMockup() {
  APP_STATE = "SAVING";
  const thisBatch = batchId;
  const op  = makeOp();
  const fmt     = window.outputFormat || 'png';
  const saveCmd = fmt.startsWith('jpg')
    ? `app.activeDocument.saveToOE("${fmt}");`
    : `app.activeDocument.saveToOE("png");`;

  const script = `
    try{
      app.displayDialogs=DialogModes.NO;
      var md=app.documents[0];
      app.activeDocument=md;
      // Hide opaque cover layers above smart objects
      var soIdx=-1;
      for(var i=0;i<md.layers.length;i++){if(md.layers[i].kind===LayerKind.SMARTOBJECT){soIdx=i;break;}}
      if(soIdx>0){
        for(var j=0;j<soIdx;j++){
          var ly=md.layers[j];
          try{if(ly.visible&&ly.blendMode===BlendMode.NORMAL&&ly.opacity===100&&
              ly.kind!==LayerKind.SMARTOBJECT&&ly.kind!==LayerKind.BRIGHTNESSCONTRAST&&
              ly.kind!==LayerKind.CHANNELMIXER&&ly.kind!==LayerKind.COLORBALANCE&&
              ly.kind!==LayerKind.CURVES&&ly.kind!==LayerKind.EXPOSURE&&
              ly.kind!==LayerKind.GRADIENTMAP&&ly.kind!==LayerKind.HUESATURATION&&
              ly.kind!==LayerKind.INVERSION&&ly.kind!==LayerKind.LEVELS&&
              ly.kind!==LayerKind.PHOTOFILTER&&ly.kind!==LayerKind.POSTERIZE&&
              ly.kind!==LayerKind.SELECTIVECOLOR&&ly.kind!==LayerKind.THRESHOLD&&
              ly.kind!==LayerKind.VIBRANCE){ly.visible=false;}}catch(e2){}
        }
      }
      try{md.flatten();}catch(fe){}
      ${saveCmd}
    }catch(e){app.echoToOE(JSON.stringify({type:"ERROR",msg:e.toString(),opId:"${op.id}"}));}
  `;
  window._currentBatchId = thisBatch;
  sendMessage(script);
}

function verifyAndInject() {
  APP_STATE = "INJECTING";
  const op        = makeOp();
  const timeoutMs = adaptiveTimeout(40, currentPsdIndex, batchPlan[batchPlanIndex]?.slots?.length || 1);

  // JSON.stringify the data URL to produce a safely escaped JS string literal —
  // prevents Photopea's parser from choking on any special character sequences
  // that can appear in base64 (especially in large PNG/WebP images).
  const safeUrl = JSON.stringify(currentImgBase64);
  const script = `try{if(app.documents.length<2)throw "Tab Missing";app.activeDocument=app.documents[app.documents.length-1];app.open(${safeUrl},"INJECTED_LAYER",true);app.echoToOE(JSON.stringify({type:"INJECT_DONE",opId:"${op.id}"}));}catch(e){app.echoToOE(JSON.stringify({type:"ERROR",msg:e.toString(),opId:"${op.id}"}));}`;
  sendMessage(script);

  // Heartbeat — extend deadline while Photopea is still responding
  op._heartbeat = setInterval(() => {
    if (op._dead) return;
    sendMessage(`app.echoToOE(JSON.stringify({type:"HEARTBEAT",opId:"${op.id}"}));`);
  }, 4000);

  op._timeout = setTimeout(() => {
    if (op._dead || APP_STATE !== "INJECTING") return;
    op._dead = true;
    diag(`  ⚠️ Injection timeout`, 'warn');
    if (slotRetryCount < MAX_SLOT_RETRIES) {
      slotRetryCount++;
      diag(`  Retrying injection (${slotRetryCount}/${MAX_SLOT_RETRIES})…`, 'warn');
      showToast(`⚠️ Injection timeout — retrying (${slotRetryCount}/${MAX_SLOT_RETRIES})`, 'warn');
      sendMessage(`try{while(app.documents.length>0){app.activeDocument=app.documents[0];app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}}catch(e){}`);
      setTimeout(() => { if (APP_STATE !== "STOPPED") loadNextPlanItem(); }, 1500);
    } else {
      addErrorToGallery('Injection timeout after retries');
    }
  }, timeoutMs);
}

function processAndSave() {
  APP_STATE = "SAVING";
  const op  = makeOp();

  const plan    = batchPlan[batchPlanIndex];
  const slot    = plan.slots[currentSlotIndex];
  const img     = imgQueue[slot.imgIdx];
  const fitMode = (img?.fitModeOverridden ? img.fitMode : null) || window.currentFitMode || 'fill';

  let resizeLine;
  if (fitMode === 'fit') {
    resizeLine = `var dw=doc.width;var dh=doc.height;var lw=l.bounds[2]-l.bounds[0];var lh=l.bounds[3]-l.bounds[1];if(lw>0&&lh>0){var s=Math.min(dw/lw,dh/lh);l.resize(s*100,s*100);}`;
  } else if (fitMode === 'stretch') {
    resizeLine = `var dw=doc.width;var dh=doc.height;var lw=l.bounds[2]-l.bounds[0];var lh=l.bounds[3]-l.bounds[1];if(lw>0&&lh>0){l.resize((dw/lw)*100,(dh/lh)*100);}`;
  } else {
    resizeLine = `var dw=doc.width;var dh=doc.height;var lw=l.bounds[2]-l.bounds[0];var lh=l.bounds[3]-l.bounds[1];if(lw>0&&lh>0){var s=Math.max(dw/lw,dh/lh)*1.02;l.resize(s*100,s*100);}`;
  }

  const timeoutMs = adaptiveTimeout(40, currentPsdIndex, plan?.slots?.length || 1);

  const script = `
    try{
      app.displayDialogs=DialogModes.NO;
      var doc=app.activeDocument;
      var l=doc.activeLayer;
      try{l.rasterize();}catch(e){}
      ${resizeLine}
      var b=l.bounds;
      l.translate((doc.width/2)-((b[0]+b[2])/2),(doc.height/2)-((b[1]+b[3])/2));
      doc.save();
      doc.close();
      app.activeDocument=app.documents[0];
      app.echoToOE(JSON.stringify({type:"SLOT_SAVED",opId:"${op.id}"}));
    }catch(e){app.echoToOE(JSON.stringify({type:"ERROR",msg:e.toString(),opId:"${op.id}"}));}
  `;
  sendMessage(script);

  // Adaptive timeout — close docs first, then error and move on
  op._timeout = setTimeout(() => {
    if (op._dead || APP_STATE !== "SAVING") return;
    op._dead = true;
    console.warn('processAndSave timeout');
    try { sendMessage(`try{app.displayDialogs=DialogModes.NO;while(app.documents.length>0){app.activeDocument=app.documents[0];app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}}catch(e){}`); } catch(_) {}
    setTimeout(() => addErrorToGallery('Save timeout'), 800);
  }, timeoutMs);
}

// ── 4. GALLERY ───────────────────────────────────────────────

async function addResultToGallery(blob) {
  if (APP_STATE === "STOPPED" || APP_STATE === "IDLE") return;
  const plan = batchPlan[batchPlanIndex];
  if (!plan) return;

  // Apply watermark for free users
  if (!isPremium()) blob = await applyWatermark(blob);

  const key     = buildInputKey(plan.inputIdx);
  const psdName = psdQueue[plan.psdIdx]?.name;
  if (!generatedMockups[key]) generatedMockups[key] = [];
  const url = URL.createObjectURL(blob);
  const firstSlot = plan.slots[0];
  const imgName = firstSlot && imgQueue[firstSlot.imgIdx] ? imgQueue[firstSlot.imgIdx].name : null;

  // Generate a small preview thumbnail (400px wide JPEG) so the gallery
  // <img> tags decode a fraction of the full export's pixel data.
  // Full blob is kept for download/lightbox only.
  let thumbUrl = url; // fallback to full if canvas fails
  try {
    const blobUrl = url;
    await new Promise(resolve => {
      const tmpImg = new Image();
      tmpImg.onload = () => {
        const THUMB_W = 400;
        const scale   = THUMB_W / tmpImg.naturalWidth;
        const canvas  = document.createElement('canvas');
        canvas.width  = THUMB_W;
        canvas.height = Math.round(tmpImg.naturalHeight * scale);
        canvas.getContext('2d').drawImage(tmpImg, 0, 0, canvas.width, canvas.height);
        thumbUrl = canvas.toDataURL('image/jpeg', 0.82);
        resolve();
      };
      tmpImg.onerror = resolve; // fallback
      tmpImg.src = blobUrl;
    });
  } catch(_) {}

  generatedMockups[key].push({ psd: psdName, blob, url, thumbUrl, error: false, imgName });
  totalGenerated++;
  if (totalGenerated - totalErrors >= 1) btnDownloadAll.disabled = false; // FIX 8 — unlock as first result arrives
  updateGalleryCount(); // FIX 10
  updateGroupUI(key, plan);
  // Mark PSD done if this was its last input
  const box = document.getElementById(`psd-box-${plan.psdIdx}`);
  const allDone = batchPlan.slice(batchPlanIndex + 1).every(p => p.psdIdx !== plan.psdIdx);
  if (box && allDone) { box.classList.remove("active"); box.classList.add("done"); }
  batchPlanIndex++;
  loadNextPlanItem();
}

function addErrorToGallery(errMsg = '') {
  const plan = batchPlan[batchPlanIndex];
  if (!plan) { batchPlanIndex++; loadNextPlanItem(); return; }
  const key     = buildInputKey(plan.inputIdx);
  const psdName = psdQueue[plan.psdIdx]?.name;
  if (!generatedMockups[key]) generatedMockups[key] = [];
  generatedMockups[key].push({ psd: psdName, blob: null, url: null, error: true, errMsg });
  totalErrors++;
  updateGalleryCount();
  updateGroupUI(key, plan);
  batchPlanIndex++;
  // Always force-close all Photopea docs before advancing — ensures the next
  // item starts from a clean 0-doc state regardless of what state Photopea is in.
  // We don't wait for CLEANUP_DONE here — just fire-and-forget, then delay to give
  // Photopea time to process the close before the next item's cleanup runs.
  try {
    sendMessage(`try{app.displayDialogs=DialogModes.NO;while(app.documents.length>0){app.activeDocument=app.documents[0];app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}}catch(e){}`);
  } catch(_) {}
  setTimeout(loadNextPlanItem, 800); // longer delay to let Photopea settle
}

function updateGroupUI(key, plan) {
  const safeId  = CSS.escape(key.replace(/\W/g,'_'));
  const psdName = psdQueue[plan.psdIdx]?.name || '';
  let group = document.getElementById(`group-${safeId}`);

  if (!group) {
    group = document.createElement("div");
    group.className = "result-group";
    group.id = `group-${safeId}`;
    group.innerHTML = `
      <div class="group-header">
        <div class="group-info">
          <div class="group-title">${key}</div>
        </div>
        <button class="btn-small" id="dlbtn-${safeId}" disabled>↓ Download</button>
      </div>
      <div class="mockup-scroll" id="scroll-${safeId}"></div>`;
    resultsContainer.appendChild(group);
    enableDragScroll(group.querySelector('.mockup-scroll'));
    if (resultsContainer.children.length === 1) {
      setTimeout(() => galleryArea.scrollIntoView({ behavior: 'smooth', block: 'start' }), 150);
    }
  }

  const entries = generatedMockups[key];
  const latest  = entries[entries.length - 1];
  const scroll  = group.querySelector('.mockup-scroll');
  const ext     = getExt();

  if (latest.error) {
    const card = document.createElement("div");
    card.className = "mockup-card error-card";
    const errDetail = latest.errMsg && latest.errMsg.toLowerCase().includes('smart')
      ? 'No smart object found'
      : latest.errMsg && latest.errMsg.toLowerCase().includes('tab')
      ? 'PSD tab did not open'
      : 'Generation failed';
    card.innerHTML = `
      <div class="error-card-inner">
        <div class="error-icon">✕</div>
        <div class="error-label">${psdName.replace(/\.psd$/i,'')}</div>
        <div class="error-sub">${errDetail}</div>
        <div class="error-hint">Check PSD has a smart object layer</div>
      </div>`;
    scroll.appendChild(card);
  } else {
    const rawPsd    = psdName.replace(/\.psd$/i,'');
    // Use the plan item passed to updateGroupUI directly (not batchPlanIndex which may have advanced)
    const psdSlots  = plan?.slots || [];
    const imgName   = psdSlots.length > 0 && imgQueue[psdSlots[0].imgIdx]
      ? imgQueue[psdSlots[0].imgIdx].name.replace(/\.[^.]+$/, '')
      : null;
    // filename = imageName_psdName (or just psdName if no image name available)
    const rawBase   = imgName ? `${imgName}_${rawPsd}` : rawPsd;
    const namedBase = applyNaming(rawBase);
    const card = document.createElement("div");
    card.className = "mockup-card";
    card.dataset.url  = latest.url;
    card.dataset.name = namedBase;
    card.dataset.ext  = ext;
    card.innerHTML = `
      <img data-src="${latest.thumbUrl || latest.url}" alt="${rawBase}" class="mockup-thumb mockup-thumb-lazy">
      <div class="card-actions">
        <button class="card-btn card-btn-copy" title="Copy to clipboard">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>
        </button>
        <button class="card-btn card-btn-dl" title="Download">↓</button>
      </div>`;
    // Wire lazy load: swap data-src → src when card enters viewport
    const lazyImg = card.querySelector('.mockup-thumb-lazy');
    if (lazyImg) {
      if (!window._thumbObserver) {
        window._thumbObserver = new IntersectionObserver(entries => {
          entries.forEach(entry => {
            if (entry.isIntersecting) {
              const img = entry.target;
              if (img.dataset.src) { img.src = img.dataset.src; delete img.dataset.src; }
              window._thumbObserver.unobserve(img);
            }
          });
        }, { rootMargin: '200px' });
      }
      window._thumbObserver.observe(lazyImg);
    }
    card.querySelector('.mockup-thumb').addEventListener('click', () => openLightbox(latest.url, namedBase));
    card.querySelector('.card-btn-copy').addEventListener('click', e => { e.stopPropagation(); copyToClipboard(latest.url); });
    card.querySelector('.card-btn-dl').addEventListener('click', e => { e.stopPropagation(); downloadSingle(latest.url, namedBase, ext); });
    const dlBtn = document.getElementById(`dlbtn-${safeId}`);
    if (dlBtn) {
      dlBtn.disabled = false;
      dlBtn.classList.toggle('btn-locked', !isPremium());
      dlBtn.title = isPremium() ? '' : 'Upgrade to Pro to download groups';
      dlBtn.onclick = () => {
        if (!isPremium()) { showUpgradeModal(); return; }
        downloadGroup(key);
      };
    }
    scroll.appendChild(card);
  }
}

// ── 5. FINISH / STOP / CLEAR ──────────────────────────────────

// FIX 10 — update gallery count badge
function updateGalleryCount() {
  const el = document.getElementById('gallery-count');
  if (!el) return;
  const ok   = totalGenerated - totalErrors;
  const errs = totalErrors;
  if (ok + errs === 0) { el.textContent = ''; return; }
  el.textContent = errs > 0 ? `(${ok} done, ${errs} failed)` : `(${ok})`;
  el.className   = errs > 0 ? 'gallery-count gallery-count-warn' : 'gallery-count';
}


// ── Reset session ─────────────────────────────────────────────
function resetSession() {
  if (BUSY_STATES.includes(APP_STATE)) {
    showToast('Stop the current batch before resetting', 'warn'); return;
  }
  showConfirm(
    'Reset session?',
    'This will clear all uploaded files and generated mockups. This cannot be undone.',
    'Reset everything',
    () => {
      window.clearPsds();
      window.clearImages();
      window.clearMockups();
      APP_STATE = 'IDLE';
      progressStrip.classList.add('hidden');
      setProgress(0, '');
      setPip('Ready', 'ready');
      hideCompletionFooter();
      document.getElementById('stay-notice')?.classList.add('hidden');
      document.querySelector('.controls-bar')?.classList.remove('panels-collapsed');
      updateGenerateButton();
      window.scrollTo({ top: 0, behavior: 'smooth' });
      showToast('Session reset', 'info');
    }
  );
}
window.resetSession = resetSession;


// ── Generate footer / completion footer visibility helpers ────
// The generate footer hides while the completion footer is showing
// so they don't visually overlap.
function showCompletionFooter() {
  document.getElementById('completion-footer')?.classList.remove('hidden');
  document.querySelector('.generate-footer')?.classList.add('footer-hidden');
}
function hideCompletionFooter() {
  document.getElementById('completion-footer')?.classList.add('hidden');
  document.querySelector('.generate-footer')?.classList.remove('footer-hidden');
}

function finishAll() {
  if (APP_STATE === "STOPPED") return;
  currentImgBase64 = null; // free after full batch complete
  const ok = totalGenerated - totalErrors;
  setProgress(100, "Complete!");
  setPip("Done", "done");

  // Re-enable everything
  btnGen.disabled = false;
  inPsd.disabled  = false;
  inImg.disabled  = false;
  document.getElementById('btn-ctrl-autofill')?.removeAttribute('disabled');
  document.getElementById('btn-ctrl-add')?.removeAttribute('disabled');
  document.querySelectorAll('.btn-clear-panel').forEach(b => b.removeAttribute('disabled'));
  document.querySelectorAll('.seg-btn[data-fmt]').forEach(b => b.removeAttribute('disabled'));
  document.querySelectorAll('#fit-control .seg-btn').forEach(b => b.removeAttribute('disabled'));
  document.getElementById('jpg-quality')?.removeAttribute('disabled');
  document.getElementById('drop-psd')?.classList.remove('drop-disabled');
  document.getElementById('drop-img')?.classList.remove('drop-disabled');
  // Generate button stays as "Start Batch" — completion footer handles next actions
  document.getElementById("btn-stop").classList.add("hidden");
  document.getElementById("btn-run-again").classList.add("hidden"); // use completion footer instead
  btnDownloadAll.disabled = false;
  btnDownloadAll.classList.toggle('btn-locked', !isPremium());

  document.querySelectorAll('.btn-small[disabled]').forEach(btn => {
    const sid  = btn.id.replace('dlbtn-','');
    const name = Object.keys(generatedMockups).find(k => CSS.escape(k.replace(/\W/g,'_')) === sid);
    if (name && generatedMockups[name].some(e => !e.error)) {
      btn.disabled = false;
      btn.classList.toggle('btn-locked', !isPremium());
      btn.title = isPremium() ? '' : 'Upgrade to Pro to download groups';
    }
  });

  APP_STATE = "DONE";
  diag(`Batch complete — ${totalGenerated - totalErrors} ok, ${totalErrors} failed`, totalErrors > 0 ? 'warn' : 'ok');
  diagDone();
  setIframeBusy(false);
  document.getElementById('stay-notice')?.classList.add('hidden');
  // Keep upload + slot panels visible but collapse controls bar (locked during batch)
  document.querySelector('.upload-row')?.classList.remove('panels-collapsed');
  document.getElementById('slot-panels')?.classList.remove('panels-collapsed');
  document.querySelector('.controls-bar')?.classList.remove('panels-collapsed');

  const msg = totalErrors > 0
    ? `✓ ${ok} mockup${ok !== 1 ? 's' : ''} generated — ${totalErrors} failed`
    : `✓ Batch complete — ${ok} mockup${ok !== 1 ? 's' : ''} generated`;
  showToast(msg, totalErrors > 0 ? 'warn' : 'success');

  // Persistent completion summary in progress strip
  const labelEl = document.getElementById('progress-label');
  const pctEl   = document.getElementById('progress-pct');
  if (labelEl) labelEl.textContent = totalErrors > 0
    ? `${ok} mockup${ok !== 1?'s':''} generated · ${totalErrors} failed`
    : `${ok} mockup${ok !== 1?'s':''} generated`;
  if (pctEl) pctEl.textContent = `${ok} / ${totalExpected}`;
  document.getElementById('progress-strip')?.classList.add('complete');

  // Completion footer — Download All + Run Again
  showCompletionFooter();
  const footer = document.getElementById('completion-footer');
  if (footer) {
    const compText = document.getElementById('completion-text');
    if (compText) compText.textContent = totalErrors > 0
      ? `${ok} mockup${ok!==1?'s':''} ready · ${totalErrors} failed`
      : `${ok} mockup${ok!==1?'s':''} ready`;
    const dlBtn = footer.querySelector('.completion-dl-btn');
    if (dlBtn) {
      dlBtn.disabled = !isPremium() || ok === 0;
      dlBtn.classList.toggle('btn-locked', !isPremium());
      dlBtn.onclick = () => { if (!isPremium()) { showUpgradeModal(); return; } downloadAllZip(); };
    }
    const againBtn = footer.querySelector('.completion-again-btn');
    if (againBtn) againBtn.onclick = () => runAgain();
  }

  // Resume any pending PSD previews
  if (previewQueue.length) schedulePsdPreviews();
}


function stopBatch() {
  if (!BUSY_STATES.includes(APP_STATE)) return;
  APP_STATE = "STOPPED";
  setIframeBusy(false);
  cancelOp();
  hideCompletionFooter();
  document.getElementById('stay-notice')?.classList.add('hidden');
  document.querySelector('.upload-row')?.classList.remove('panels-collapsed');
  document.getElementById('slot-panels')?.classList.remove('panels-collapsed');
  document.querySelector('.controls-bar')?.classList.remove('panels-collapsed');
  setPip("Stopped", "error");
  progressStrip.classList.add("hidden");
  const bc = btnGen.querySelector('.btn-content');
  if (bc) bc.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><polygon points="5,3 19,12 5,21"/></svg> Start Batch`;
  btnGen.disabled = false;
  inPsd.disabled  = false;
  inImg.disabled  = false;
  document.getElementById('btn-ctrl-autofill')?.removeAttribute('disabled');
  document.getElementById('btn-ctrl-add')?.removeAttribute('disabled');
  document.querySelectorAll('.btn-clear-panel').forEach(b => b.removeAttribute('disabled'));
  document.querySelectorAll('.seg-btn[data-fmt]').forEach(b => b.removeAttribute('disabled'));
  document.querySelectorAll('#fit-control .seg-btn').forEach(b => b.removeAttribute('disabled'));
  document.getElementById('jpg-quality')?.removeAttribute('disabled');
  document.getElementById('drop-psd')?.classList.remove('drop-disabled');
  document.getElementById('drop-img')?.classList.remove('drop-disabled');
  document.getElementById("btn-stop").classList.add("hidden");
  document.querySelectorAll('.btn-small[disabled]').forEach(btn => {
    const sid  = btn.id.replace('dlbtn-','');
    const name = Object.keys(generatedMockups).find(k => CSS.escape(k.replace(/\W/g,'_')) === sid);
    if (name && generatedMockups[name].some(e => !e.error)) {
      btn.disabled = false;
      btn.classList.toggle('btn-locked', !isPremium());
      btn.title = isPremium() ? '' : 'Upgrade to Pro to download groups';
    }
  });
  updateGenerateButton();
  showToast('Batch stopped. Partial results available below.', 'warn');
}

function clearMockupsSilent() {
  // Reset gallery count
  const countEl = document.getElementById('gallery-count');
  if (countEl) countEl.textContent = '';
  // Revoke all blob URLs before wiping to free memory
  Object.values(generatedMockups).forEach(arr =>
    arr.forEach(m => { try { if (m.url) URL.revokeObjectURL(m.url); } catch(_) {} })
  );
  generatedMockups = {};
  totalGenerated   = 0;
  totalErrors      = 0;
  resultsContainer.innerHTML = "";
}

window.clearPsds = function() {
  if (BUSY_STATES.includes(APP_STATE)) return;
  revokeUrls(psdQueue.map(p => p.previewUrl));
  psdQueue = []; previewQueue = []; inputs = []; batchPlan = []; batchPlanIndex = 0;
  psdGrid.innerHTML = "";
  psdCountBadge.classList.add("hidden");
  if (psdStatus) psdStatus.textContent = "Waiting…";
  renderSlotPanels();
  updateGenerateButton();
};

window.clearImages = function() {
  if (BUSY_STATES.includes(APP_STATE)) return;
  revokeUrls(imgQueue.map(i => i.url));
  imgQueue = []; imgGrid.innerHTML = "";
  imgCountBadge.classList.add("hidden");
  updateGenerateButton();
};

window.clearMockups = function() {
  clearMockupsSilent();
  galleryArea.classList.add("hidden");
  btnDownloadAll.disabled = true;
  hideCompletionFooter();
};

// ── 6. DOWNLOADS ──────────────────────────────────────────────

window.downloadSingle = (url, name, ext = 'png') => {
  const a = document.createElement("a"); a.href = url; a.download = `${name}.${ext}`; a.click();
};

// ── ZIP via Web Worker (non-blocking) ──────────────────────────
// Serialises blobs to ArrayBuffers, ships to worker, receives zip blob back.
// Keeps main thread free so the UI doesn't freeze during compression.
function _runZipWorker(files, zipName) {
  showToast('⏳ Building zip…', 'info');
  const workerSrc = `
    importScripts('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
    self.onmessage = async function(e) {
      const { files, folders } = e.data;
      const zip = new JSZip();
      if (folders) {
        for (const [folderName, items] of Object.entries(folders)) {
          const folder = zip.folder(folderName);
          for (const item of items) folder.file(item.name, item.buffer);
        }
      } else {
        for (const item of files) zip.file(item.name, item.buffer);
      }
      const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
      self.postMessage(blob);
    };
  `;
  const worker = new Worker(URL.createObjectURL(new Blob([workerSrc], { type: 'application/javascript' })));
  worker.onmessage = e => {
    const blob   = e.data;
    const dlUrl  = URL.createObjectURL(blob);
    const a      = document.createElement('a');
    a.href       = dlUrl;
    a.download   = zipName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(dlUrl), 10000);
    worker.terminate();
    showToast('✓ Zip ready', 'success');
  };
  worker.onerror = err => {
    console.error('Zip worker error', err);
    showToast('Zip failed — try again', 'error');
    worker.terminate();
  };
  return worker;
}

window.downloadGroup = name => {
  const ext     = getExt();
  const entries = (generatedMockups[name]||[]).filter(m => !m.error && m.blob);
  const reads   = entries.map((m, i) => {
    const psdBase   = m.psd.replace(/\.psd$/i,'');
    const imgBase   = m.imgName ? m.imgName.replace(/\.[^.]+$/, '') : null;
    const rawBase   = imgBase ? `${imgBase}_${psdBase}` : psdBase;
    const namedBase = applyNaming(rawBase) || `mockup_${i+1}`;
    return m.blob.arrayBuffer().then(buf => ({ name: `${namedBase}.${ext}`, buffer: buf }));
  });
  Promise.all(reads).then(files => {
    const w = _runZipWorker(files, `${applyNaming(name.replace(/\.\w+$/,''))}_Group.zip`);
    w.postMessage({ files });
  });
};

function downloadAllZip() {
  const ext     = getExt();
  const folderReads = {};
  const allReads = [];
  Object.keys(generatedMockups).forEach(name => {
    const folderName = sanitiseFilename(applyNaming(name.replace(/\.\w+$/, '')));
    folderReads[folderName] = [];
    (generatedMockups[name]||[]).filter(m => !m.error && m.blob).forEach((m, i) => {
      const psdBase   = m.psd.replace(/\.psd$/i, '');
      const imgBase   = m.imgName ? m.imgName.replace(/\.[^.]+$/, '') : null;
      const rawBase   = imgBase ? `${imgBase}_${psdBase}` : psdBase;
      const fileName  = `${applyNaming(rawBase) || `mockup_${i+1}`}.${ext}`;
      const p = m.blob.arrayBuffer().then(buf => {
        folderReads[folderName].push({ name: fileName, buffer: buf });
      });
      allReads.push(p);
    });
  });
  Promise.all(allReads).then(() => {
    const w = _runZipWorker(null, 'All_Mockups.zip');
    w.postMessage({ folders: folderReads });
  });
}

// ── BACKGROUND TAB RESILIENCE ────────────────────────────────
// Two complementary techniques to prevent Chrome from throttling timers
// when the tab is in the background:
//
// 1. Web Locks API — holds a lock while the batch runs, which signals to
//    the browser that this tab has active work and should not be aggressively
//    throttled. Works in Chrome/Edge/Firefox.
//
// 2. Silent Audio Context — playing an inaudible audio buffer in a loop is
//    a well-established trick that prevents timer throttling in most browsers.
//    The buffer is 1 sample of silence so it costs nothing.
(function() {
  let _lockRelease  = null;
  let _audioCtx     = null;
  let _audioSource  = null;

  function startBackgroundGuard() {
    // Web Locks
    if (navigator.locks && !_lockRelease) {
      navigator.locks.request('mockup-studio-batch', { mode: 'exclusive' }, lock => {
        return new Promise(resolve => { _lockRelease = resolve; });
      });
    }
    // Silent audio loop
    if (!_audioCtx) {
      try {
        _audioCtx = new (window.AudioContext || window.webkitAudioContext)();
        const buf = _audioCtx.createBuffer(1, 1, 22050); // 1 sample silence
        function loop() {
          if (!_audioCtx) return;
          _audioSource = _audioCtx.createBufferSource();
          _audioSource.buffer = buf;
          _audioSource.connect(_audioCtx.destination);
          _audioSource.onended = loop;
          _audioSource.start();
        }
        loop();
      } catch(_) {} // silently fail if audio blocked
    }
  }

  function stopBackgroundGuard() {
    if (_lockRelease) { _lockRelease(); _lockRelease = null; }
    if (_audioSource) { try { _audioSource.onended = null; _audioSource.stop(); } catch(_) {} _audioSource = null; }
    if (_audioCtx)    { try { _audioCtx.close(); } catch(_) {} _audioCtx = null; }
  }

  // Patch startBatch, finishAll, stopBatch
  const _origStartBatch = startBatch;
  startBatch = function() { startBackgroundGuard(); _origStartBatch.apply(this, arguments); };
  const _origFinishAll = finishAll;
  finishAll = function() { stopBackgroundGuard(); _origFinishAll.apply(this, arguments); };
  const _origStopBatch = stopBatch;
  stopBatch = function() { stopBackgroundGuard(); _origStopBatch.apply(this, arguments); };

  window._startBackgroundGuard = startBackgroundGuard;
  window._stopBackgroundGuard  = stopBackgroundGuard;
})();

// ── MEMORY PRESSURE MONITOR ──────────────────────────────────
// Chrome exposes window.performance.memory; other browsers return undefined.
// We poll every 30s and show a persistent warning banner if used heap exceeds
// 75% of the limit — a reliable signal the tab is approaching OOM territory.
(function() {
  const WARN_RATIO  = 0.75;   // warn at 75% heap usage
  const POLL_MS     = 30000;
  let   warned      = false;

  function checkMemory() {
    const mem = window.performance?.memory;
    if (!mem) return; // not supported
    const ratio = mem.usedJSHeapSize / mem.jsHeapSizeLimit;
    if (ratio >= WARN_RATIO && !warned) {
      warned = true;
      showToast(
        '⚠️ High memory usage — consider refreshing the tab before starting a new batch to avoid slowdowns.',
        'warn'
      );
      // Show persistent banner
      let banner = document.getElementById('mem-warn-banner');
      if (!banner) {
        banner = document.createElement('div');
        banner.id = 'mem-warn-banner';
        banner.style.cssText = 'position:fixed;bottom:56px;left:50%;transform:translateX(-50%);background:#92400e;color:#fef3c7;padding:8px 18px;border-radius:8px;font-size:12px;z-index:9998;display:flex;align-items:center;gap:10px;box-shadow:0 4px 20px rgba(0,0,0,0.4);';
        banner.innerHTML = '⚠️ High memory — refresh tab between batches for best performance <button onclick="this.parentElement.remove()" style="background:none;border:none;color:inherit;cursor:pointer;font-size:14px;margin-left:4px;">✕</button>';
        document.body.appendChild(banner);
      }
    } else if (ratio < WARN_RATIO - 0.1) {
      // Memory recovered (e.g. after clear) — reset warning
      warned = false;
      document.getElementById('mem-warn-banner')?.remove();
    }
    setTimeout(checkMemory, POLL_MS);
  }

  // Start polling after 10s (give app time to load)
  setTimeout(checkMemory, 10000);
})();

// ── 7. MESSAGE ROUTER ────────────────────────────────────────

// ── TAB VISIBILITY: DEADLINE FREEZE ──────────────────────────
// When the tab is hidden, freeze the current op deadline so Chrome's timer
// throttling doesn't cause false timeouts. The batch continues naturally —
// no pause, no resume, no state changes.
(function() {
  let _hiddenAt = null;

  document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
      _hiddenAt = Date.now();
    } else {
      if (_hiddenAt !== null && currentOp && !currentOp._dead) {
        // Extend deadline by however long the tab was hidden
        currentOp._deadlineMs += (Date.now() - _hiddenAt);
      }
      _hiddenAt = null;
    }
  });
})();

window.addEventListener('message', function(e) {

  // ── Preview pipeline ────────────────────────────────────
  if (psdPreviewMode) {
    if (typeof e.data === "string") {
      try {
        const msg = JSON.parse(e.data);
        if (msg.type === "PREVIEW_READY") {
          const item = psdQueue[window._previewIdx];
          if (!item) { psdPreviewMode = false; setTimeout(runNextPreview, 200); return; }
          const reader = new FileReader();
          reader.onload = function() { sendMessage(this.result); };
          reader.readAsArrayBuffer(item.file);
          return;
        }
        if (msg.type === "PREVIEW_FAIL") {
          psdPreviewMode = false;
          sendMessage(`try{app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}catch(e){}`);
          setTimeout(runNextPreview, 300); return;
        }
        if (msg.type === "PREVIEW_SO_NAMES") {
          const idx = window._previewIdx;
          if (idx !== undefined && psdQueue[idx]) {
            psdQueue[idx].soNames = msg.names.length > 0 ? msg.names : ['Design'];
            onPsdScanComplete(idx);
          }
          return;
        }
      } catch(_) {}
    }
    if (e.data === "done") {
      // PSD loaded — hide placeholders, scan SO names, export thumbnail — all in one script
      sendMessage(`
        app.displayDialogs=DialogModes.NO;
        try{
          var doc=app.activeDocument;
          function hidePH(p){for(var i=0;i<p.layers.length;i++){var l=p.layers[i];var n=l.name.toLowerCase();if(n.indexOf("placeholder")!==-1||n.indexOf("delete")!==-1||n.indexOf("preview")!==-1||n.indexOf("your design")!==-1||n.indexOf("replace me")!==-1||n.indexOf("remove")!==-1||n.indexOf("instruction")!==-1||n.indexOf("guide")!==-1||n.indexOf("watermark")!==-1||n.indexOf("promo")!==-1){try{l.visible=false;}catch(e2){}}}}
          hidePH(doc);
          function findAllSO(p,r){if(!p||!p.layers)return;for(var i=0;i<p.layers.length;i++){var l=p.layers[i];if(l.visible===false)continue;if(l.kind==LayerKind.SMARTOBJECT)r.push(l.name);else if(l.layers&&l.layers.length>0)findAllSO(l,r);}}
          var soNames=[];findAllSO(doc,soNames);
          app.echoToOE(JSON.stringify({type:"PREVIEW_SO_NAMES",names:soNames}));
          try{doc.flatten();}catch(e){}
          doc.saveToOE("png");
        }catch(ex){app.echoToOE(JSON.stringify({type:"PREVIEW_FAIL"}));}
      `);
      return;
    }
    if (e.data instanceof ArrayBuffer) {
      const idx = window._previewIdx;
      if (idx !== undefined && psdQueue[idx]) {
        revokeUrls([psdQueue[idx].previewUrl]);
        psdQueue[idx].previewUrl = URL.createObjectURL(new Blob([e.data],{type:"image/png"}));
        renderPsdThumbs();
      }
      // Clear loading state
      const box = document.getElementById(`psd-box-${window._previewIdx}`);
      if (box) box.classList.remove('previewing');
      psdPreviewMode = false;
      sendMessage(`try{app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);}catch(e){}`);
      setTimeout(runNextPreview, 300);
      return;
    }
    return;
  }

  // ── Photopea ready ───────────────────────────────────────
  if (e.data === "done" && APP_STATE === "IDLE") {
    clearTimeout(photopeaLoadTimeout);
    setPip("Ready", "ready");
    updateGenerateButton();
    // If we arrived from the library, load the template now that Photopea is ready
    if (window._pendingLibraryPsd) {
      const entry = window._pendingLibraryPsd;
      window._pendingLibraryPsd = null;
      loadTemplateFromLibrary(entry);
    }
    return;
  }

  // ── Batch pipeline ───────────────────────────────────────
  if (APP_STATE === "STOPPED") return;

  if (e.data === "done" && APP_STATE === "LOADING_PSD") { scanPsd(); return; }

  if (typeof e.data === "string") {
    try {
      const msg = JSON.parse(e.data);

      // ── Heartbeat reply — extend the current op's deadline ──
      if (msg.type === "HEARTBEAT") {
        if (currentOp && !currentOp._dead && msg.opId === currentOp.id) {
          currentOp._deadlineMs = Date.now() + 30000; // extend by 30s on each heartbeat
        }
        return;
      }

      // ── Staleness guard — discard messages from superseded operations ──
      // CLEANUP_DONE carries no opId (it fires before any op is created) — always allow.
      if (msg.opId && currentOp && msg.opId !== currentOp.id) return;
      if (msg.opId && !currentOp) return; // op was already cancelled

      if (msg.type === "CLEANUP_DONE") {
        uploadPsdFile();
      }
      else if (msg.type === "SCAN_OK") {
        cancelOp();
        startSlotLoop();
      }
      else if (msg.type === "MULTI_SO") {
        cancelOp();
        startSlotLoop();
      }
      else if (msg.type === "POLL_DOCS") {
        const op = currentOp;
        // Every POLL_DOCS reply is proof Photopea is alive — extend deadline
        if (op && !op._dead && APP_STATE === "OPENING_SO") {
          op._deadlineMs = Math.max(op._deadlineMs, Date.now() + 10000);
        }
        if (op && op._waitDocCallback != null) {
          // In a waitForDocCount closure — let the closure handle it
          handleDocCountWait(msg.count, op);
        } else if (msg.count >= 2 && APP_STATE === "OPENING_SO") {
          // SO tab has opened — transition state FIRST to prevent duplicate triggers
          // from queued POLL_DOCS messages arriving in the same event loop tick.
          APP_STATE = "INJECTING";
          diag(`  SO opened (${msg.count} docs) — injecting image…`, 'ok');
          if (op) {
            if (op._interval)  { clearInterval(op._interval);  op._interval  = null; }
            if (op._heartbeat) { clearInterval(op._heartbeat); op._heartbeat = null; }
            if (op._timeout)   { clearTimeout(op._timeout);    op._timeout   = null; }
            op._dead = true;
          }
          // Scale settle delay to template complexity
          const plan = batchPlan[batchPlanIndex];
          const soCount = plan ? plan.slots.length : 1;
          const delay = soCount >= 6 ? 1400 : soCount >= 3 ? 1000 : 700;
          setTimeout(verifyAndInject, delay);
        }
      }
      else if (msg.type === "INJECT_DONE") {
        diag(`  Image injected — saving slot…`, 'ok');
        if (currentOp) {
          if (currentOp._heartbeat) { clearInterval(currentOp._heartbeat); currentOp._heartbeat = null; }
          if (currentOp._timeout)   { clearTimeout(currentOp._timeout);    currentOp._timeout   = null; }
          currentOp._dead = true;
        }
        setTimeout(processAndSave, 200);
      }
      else if (msg.type === "SLOT_SAVED") {
        diag(`  Slot saved ✓`, 'ok');
        cancelOp();
        slotRetryCount = 0;
        currentSlotIndex++;
        // Wait for the SO sub-document to close before processing the next slot
        waitForDocCount(1, () => processNextSlot(false));
      }
      else if (msg.type === "ERROR") {
        cancelOp();
        diag(`  ✕ Photopea error: ${msg.msg} (state: ${APP_STATE})`, 'error');
        addErrorToGallery(msg.msg || '');
      }
    } catch(_) {}
    return;
  }

  if (e.data instanceof ArrayBuffer) {
    // Drop stale ArrayBuffers from a previous or stopped batch
    if (APP_STATE === "STOPPED" || APP_STATE === "IDLE") return;
    if (window._currentBatchId !== batchId) return;
    const fmt  = window.outputFormat || 'png';
    const mime = fmt.startsWith('jpg') ? 'image/jpeg' : 'image/png';

    // ── Retry if result looks empty/corrupt (too small to be a real image) ──
    if (e.data.byteLength < MIN_RESULT_BYTES) {
      slotRetryCount++;
      if (slotRetryCount <= MAX_SLOT_RETRIES) {
        console.warn(`Result too small (${e.data.byteLength} bytes) — retrying slot (attempt ${slotRetryCount})`);
        showToast(`⚠️ Slot result was empty — retrying… (${slotRetryCount}/${MAX_SLOT_RETRIES})`, 'warn');
        // Re-run this slot from scratch
        const plan = batchPlan[batchPlanIndex];
        const slot = plan?.slots[currentSlotIndex - 1] || plan?.slots[currentSlotIndex];
        if (slot) {
          currentSlotIndex = Math.max(0, currentSlotIndex - 1);
          // Small extra delay before retry on slow machines
          setTimeout(() => processNextSlot(true), 600);
        } else {
          addResultToGallery(new Blob([e.data], {type: mime})); // no slot to retry
        }
        return;
      }
      console.warn('Max retries reached for slot — accepting result anyway');
    }

    slotRetryCount = 0; // reset on successful result
    diag(`  Export received (${Math.round(e.data.byteLength/1024)}KB) → gallery`, 'ok');
    addResultToGallery(new Blob([e.data], {type: mime}));
  }
});
