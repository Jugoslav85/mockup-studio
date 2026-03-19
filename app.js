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
let currentPsdIndex  = 0;
let pollInterval     = null;
let retryCount       = 0;
let APP_STATE        = "IDLE";
let currentImgBase64 = null;
let batchId          = 0;  // incremented each run; stale Photopea callbacks are ignored

// ── PSD PREVIEW ───────────────────────────────────────────────
let previewQueue   = [];
let previewRunning = false;
let psdPreviewMode = false;

const BUSY_STATES = ["RUNNING","LOADING_PSD","OPENING_SO","INJECTING","SAVING"];

// ── HELPERS ──────────────────────────────────────────────────

function sendMessage(payload) { iframe.contentWindow.postMessage(payload, "*"); }

function setProgress(pct, label) {
  if (pct !== undefined) {
    progressBar.style.width = Math.round(pct) + "%";
    if (progressPct) progressPct.textContent = Math.round(pct) + "%";
  }
  if (label && progressLabel) progressLabel.textContent = label;
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
  const val  = (document.getElementById('naming-value')?.value || '').trim();
  const mode = document.querySelector('input[name="naming-mode"]:checked')?.value || 'none';
  if (!val || mode === 'none') return baseName;
  if (mode === 'prefix') return val + baseName;
  // suffix: insert before extension
  const dot = baseName.lastIndexOf('.');
  return dot > 0 ? baseName.slice(0, dot) + val + baseName.slice(dot) : baseName + val;
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

  const bc = btnGen.querySelector('.btn-content');
  if (isDone && bc) {
    bc.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><polygon points="5,3 19,12 5,21"/></svg> New Batch`;
    btnGen.disabled = false;
  } else if (!isDone && bc && bc.textContent.includes('New')) {
    bc.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><polygon points="5,3 19,12 5,21"/></svg> Start Batch`;
  }

  if (!hasPsd)          btnHint.textContent = "Upload PSDs & images to begin";
  else if (previewing)  btnHint.textContent = "Generating PSD previews, please wait…";
  else if (!readyInputs) btnHint.textContent = "Assign images to slots to begin";
  else                  btnHint.textContent = `${readyInputs} input${readyInputs>1?'s':''} ready across ${psdQueue.length} template${psdQueue.length>1?'s':''}`;
}

// ── DRAG-TO-REORDER ───────────────────────────────────────────

function enableDragReorder(container, queue, onReorder) {
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
    const t = e.target.closest('[draggable="true"]');
    if (!t || t === dragSrc || !dragSrc) return;
    const si = +dragSrc.dataset.idx, ti = +t.dataset.idx;
    const [m] = queue.splice(si, 1);
    queue.splice(ti, 0, m);
    onReorder();
  });
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

window.onload = function () {
  setPip("Loading…", "load");
  iframe.src = "https://www.photopea.com#" + JSON.stringify({ environment: { theme: 0 } });

  inPsd.addEventListener('change', handlePsdUpload);
  inImg.addEventListener('change', handleImagesSelect);
  btnGen.addEventListener('click', handleGenClick);
  btnDownloadAll.addEventListener('click', downloadAllZip);
  document.getElementById("btn-stop").addEventListener('click', stopBatch);
  document.getElementById("btn-run-again").addEventListener('click', runAgain);

  // OS-level drag-and-drop visual feedback on upload zones
  ['dragenter','dragover'].forEach(evt => {
    document.getElementById('drop-psd').addEventListener(evt, e => { e.preventDefault(); e.currentTarget.classList.add('drag-active'); });
    document.getElementById('drop-img').addEventListener(evt, e => { e.preventDefault(); e.currentTarget.classList.add('drag-active'); });
  });
  ['dragleave','drop'].forEach(evt => {
    document.getElementById('drop-psd').addEventListener(evt, () => document.getElementById('drop-psd').classList.remove('drag-active'));
    document.getElementById('drop-img').addEventListener(evt, () => document.getElementById('drop-img').classList.remove('drag-active'));
  });

  // Cmd/Ctrl+Enter shortcut
  document.addEventListener('keydown', e => {
    if ((e.metaKey || e.ctrlKey) && e.key === 'Enter' && !btnGen.disabled) handleGenClick();
  });

  enableDragReorder(psdGrid, psdQueue, renderPsdThumbs);
  enableDragReorder(imgGrid, imgQueue, renderImgThumbs);
  updateGenerateButton();
};

// New Batch vs Start Batch
function handleGenClick() {
  if (APP_STATE === "DONE") {
    // Reset to idle — keep files loaded, scroll to top
    APP_STATE = "IDLE";
    inPsd.disabled = false;
    inImg.disabled = false;
    progressStrip.classList.add("hidden");
    setProgress(0, "");
    document.getElementById("btn-run-again").classList.add("hidden");
    setPip("Ready", "ready");
    window.scrollTo({ top: 0, behavior: 'smooth' });
    updateGenerateButton();
  } else {
    startBatch();
  }
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
  panelImg.classList.remove("disabled");
  inImg.disabled = false;
  inPsd.value = "";
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

  // ── Top action bar ────────────────────────────────────────
  const actionBar = document.createElement('div');
  actionBar.className = 'slot-action-bar';
  actionBar.innerHTML = `
    <div class="slot-action-bar-left">
      <span class="slot-panel-title-text">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
        Mockup Inputs
      </span>
      <span class="slot-hint-text">Drag images into slots, or click to upload</span>
      <button class="slot-help-btn" id="btn-slot-help" title="How does this work?">?</button>
    </div>
    <div class="slot-panel-actions">
      <button class="btn-slot-autofill" id="btn-autofill">⚡ Auto-fill</button>
      <button class="btn-slot-add" id="btn-add-input">+ Input</button>
    </div>`;
  actionBar.querySelector('#btn-autofill').addEventListener('click', autoFillInputs);
  actionBar.querySelector('#btn-add-input').addEventListener('click', addInput);
  actionBar.querySelector('#btn-slot-help').addEventListener('click', showSlotHelp);
  panel.appendChild(actionBar);

  // ── Table: rows = PSDs, columns = inputs ─────────────────
  // Layout: fixed left column (PSD labels) + scrolling right area (input columns)
  const tableWrap = document.createElement('div');
  tableWrap.className = 'slot-table-wrap';

  // Left column — PSD labels with thumbnail
  const labelCol = document.createElement('div');
  labelCol.className = 'slot-label-col';

  // Top-left corner cell — blank, aligns with input headers
  const cornerCell = document.createElement('div');
  cornerCell.className = 'slot-corner-cell';
  labelCol.appendChild(cornerCell);

  psdQueue.forEach((psd, psdIdx) => {
    if (!psd.soNames?.length) return;
    const psdShortName = psd.name.replace(/\.psd$/i,'');
    psd.soNames.forEach((soName, soIdx) => {
      const label = document.createElement('div');
      // First SO of each PSD gets the group divider + name badge
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

  // Scrollable right area — input columns
  const inputsScroll = document.createElement('div');
  inputsScroll.className = 'slot-inputs-scroll';
  enableDragScroll(inputsScroll);

  inputs.forEach((input, inputIdx) => {
    const col = document.createElement('div');
    col.className = 'slot-input-col';

    // Column header
    const colHeader = document.createElement('div');
    colHeader.className = 'slot-col-header';
    colHeader.innerHTML = `
      <span class="slot-input-label">Input ${inputIdx + 1}</span>
      <button class="slot-input-remove" title="Remove">×</button>`;
    colHeader.querySelector('.slot-input-remove').addEventListener('click', () => removeInput(inputIdx));
    col.appendChild(colHeader);

    // One slot per PSD×SO row
    // Track first-slot-of-PSD for the group divider line
    let lastPsdIdx = -1;
    input.forEach((slot, slotIdx) => {
      const img = slot.imgIdx !== null && slot.imgIdx !== undefined ? imgQueue[slot.imgIdx] : null;
      const isFirstOfPsd = slot.psdIdx !== lastPsdIdx;
      lastPsdIdx = slot.psdIdx;

      const zone = document.createElement('div');
      zone.className = `slot-drop-zone${img ? ' slot-filled' : ''}${isFirstOfPsd ? ' psd-group-first' : ''}`;

      if (img) {
        zone.innerHTML = `
          <img src="${img.url}" class="slot-assigned-img">
          <button class="slot-clear-btn" title="Clear">×</button>`;
        zone.querySelector('.slot-clear-btn').addEventListener('click', e => {
          e.stopPropagation();
          inputs[inputIdx][slotIdx].imgIdx = null;
          renderSlotPanels(); updateGenerateButton();
        });
      } else {
        zone.innerHTML = `
          <div class="slot-empty-content">
            <div class="slot-drop-hint">Drop or click</div>
          </div>`;
        zone.addEventListener('click', () => {
          const fp = document.createElement('input');
          fp.type = 'file'; fp.accept = 'image/*';
          fp.onchange = ev => {
            const f = ev.target.files[0]; if (!f) return;
            const existing = imgQueue.find(q => q.name === f.name);
            let idx;
            if (existing) {
              idx = imgQueue.indexOf(existing);
            } else {
              imgQueue.push({ file: f, name: f.name, url: URL.createObjectURL(f), fitMode: window.currentFitMode || 'fill', fitModeOverridden: false });
              idx = imgQueue.length - 1;
              renderImgThumbs();
              imgCountBadge.textContent = `${imgQueue.length} file${imgQueue.length>1?'s':''}`;
              imgCountBadge.classList.remove("hidden");
            }
            inputs[inputIdx][slotIdx].imgIdx = idx;
            renderSlotPanels(); updateGenerateButton();
          };
          fp.click();
        });
      }

      zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
      zone.addEventListener('drop', e => {
        e.preventDefault(); zone.classList.remove('drag-over');
        const imgIdx = parseInt(e.dataTransfer.getData('imgIdx'));
        if (!isNaN(imgIdx)) {
          inputs[inputIdx][slotIdx].imgIdx = imgIdx;
          renderSlotPanels(); updateGenerateButton();
        }
      });

      col.appendChild(zone);
    });

    inputsScroll.appendChild(col);
  });

  tableWrap.appendChild(inputsScroll);
  panel.appendChild(tableWrap);
  container.appendChild(panel);
}

function addInput() {
  if (!inputs.length) return;
  // Duplicate structure of first input with all slots empty
  const newInput = inputs[0].map(slot => ({ ...slot, imgIdx: null }));
  inputs.push(newInput);
  renderSlotPanels();
}

function removeInput(inputIdx) {
  if (inputs.length <= 1) return;
  inputs.splice(inputIdx, 1);
  renderSlotPanels();
  updateGenerateButton();
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
    panelImg.classList.add("disabled");
    inImg.disabled = true;
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

// ── 2. IMAGE UPLOAD ──────────────────────────────────────────

function handleImagesSelect(e) {
  const files = Array.from(e.target.files);
  if (!files.length) return;
  const existing = new Set(imgQueue.map(f => f.name));
  const dupes = files.filter(f => existing.has(f.name));
  if (dupes.length) showToast(`⚠️ Skipped ${dupes.length} duplicate${dupes.length>1?'s':''}: ${dupes.map(d=>d.name).join(', ')}`, 'warn');
  files.forEach(f => {
    if (!existing.has(f.name))
      imgQueue.push({ file: f, name: f.name, url: URL.createObjectURL(f), fitMode: window.currentFitMode || 'fill', fitModeOverridden: false });
  });
  renderImgThumbs();
  renderSlotPanels();
  imgCountBadge.textContent = `${imgQueue.length} file${imgQueue.length>1?'s':''}`;
  imgCountBadge.classList.remove("hidden");
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
      ev.dataTransfer.effectAllowed = 'copy';
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
  psdPreviewMode   = false;
  APP_STATE        = "RUNNING";
  batchPlanIndex   = 0;
  currentPsdIndex  = 0;
  currentSlotIndex = 0;
  loadNextPlanItem();
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
  if (pollInterval) { clearInterval(pollInterval); clearTimeout(pollInterval); }

  // Run placeholder hiding and confirm doc is ready — no SO scan needed
  sendMessage(`try{
    app.displayDialogs=DialogModes.NO;
    var doc=app.activeDocument;
    function hidePH(p){for(var i=0;i<p.layers.length;i++){var l=p.layers[i];var n=l.name.toLowerCase();if(n.indexOf("placeholder")!==-1||n.indexOf("delete")!==-1||n.indexOf("preview")!==-1||n.indexOf("your design")!==-1||n.indexOf("replace me")!==-1||n.indexOf("remove")!==-1||n.indexOf("instruction")!==-1||n.indexOf("guide")!==-1||n.indexOf("watermark")!==-1||n.indexOf("promo")!==-1){try{l.visible=false;}catch(e2){}}}}
    hidePH(doc);
    app.echoToOE(JSON.stringify({type:"SCAN_OK"}));
  }catch(e){app.echoToOE(JSON.stringify({type:"SCAN_OK"}));}`);

  // Timeout safety: if we never hear back, skip this PSD
  pollInterval = setTimeout(() => {
    if (APP_STATE === "STOPPED") return;
    showToast(`⚠️ Skipped "${psdQueue[currentPsdIndex]?.name}" — PSD did not respond`, 'warn');
    const box = document.getElementById(`psd-box-${currentPsdIndex}`);
    if (box) { box.classList.remove("active"); box.classList.add("done"); }
    batchPlanIndex++;
    loadNextPlanItem();
  }, 12000);
}

function startSlotLoop() {
  // Called after PSD is loaded and initial scan confirms SOs are present
  currentSlotIndex = 0;
  processNextSlot();
}

function processNextSlot() {
  if (APP_STATE === "STOPPED") return;
  const plan  = batchPlan[batchPlanIndex];
  const slots = plan.slots;

  if (currentSlotIndex >= slots.length) {
    // All slots filled — export the final composite
    exportFinalMockup();
    return;
  }

  const slot = slots[currentSlotIndex];
  const img  = imgQueue[slot.imgIdx];
  if (!img) { currentSlotIndex++; processNextSlot(); return; }

  const done = batchPlanIndex;
  setProgress(
    (done / totalExpected) * 100,
    `Input ${plan.inputIdx + 1} — "${slot.soName}" → ${psdQueue[plan.psdIdx]?.name || ''}`
  );

  const reader = new FileReader();
  reader.onload = ev => {
    currentImgBase64 = ev.target.result;
    selectAndOpenSlot(slot.soName);
  };
  reader.readAsDataURL(img.file);
}

function selectAndOpenSlot(soName) {
  APP_STATE = "OPENING_SO";
  const escaped = soName.replace(/\\/g,'\\\\').replace(/"/g,'\\"');
  // Select the named SO layer, then open it
  const script = `
    try{
      app.displayDialogs=DialogModes.NO;
      var doc=app.activeDocument;
      function findByName(p,n){for(var i=0;i<p.layers.length;i++){var l=p.layers[i];if(l.name===n)return l;if(l.layers&&l.layers.length>0){var f=findByName(l,n);if(f)return f;}}return null;}
      var target=findByName(doc,"${escaped}");
      if(!target){
        // Fallback: find first visible SO
        function findFirstSO(p){if(!p||!p.layers)return null;for(var i=0;i<p.layers.length;i++){var l=p.layers[i];if(l.visible!==false&&l.kind==LayerKind.SMARTOBJECT)return l;if(l.layers&&l.layers.length>0){var f=findFirstSO(l);if(f)return f;}}return null;}
        target=findFirstSO(doc);
      }
      if(!target)throw "SO not found: ${escaped}";
      doc.activeLayer=target;
      var desc=new ActionDescriptor();
      executeAction(stringIDToTypeID("placedLayerEditContents"),desc,3);
      app.echoToOE(JSON.stringify({type:"SO_OPEN_CMD"}));
    }catch(e){app.echoToOE(JSON.stringify({type:"ERROR",msg:e.toString()}));}
  `;
  sendMessage(script);
  if (pollInterval) { clearInterval(pollInterval); clearTimeout(pollInterval); }
  pollInterval = setInterval(() => {
    sendMessage(`app.echoToOE(JSON.stringify({type:"POLL_DOCS",count:app.documents.length}));`);
  }, 500);
}

function exportFinalMockup() {
  APP_STATE = "SAVING";
  const thisBatch = batchId;
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
    }catch(e){app.echoToOE(JSON.stringify({type:"ERROR",msg:e.toString(),bid:${thisBatch}}));}
  `;
  window._currentBatchId = thisBatch;
  sendMessage(script);
}

function verifyAndInject() {
  APP_STATE = "INJECTING";
  if (pollInterval) { clearInterval(pollInterval); clearTimeout(pollInterval); }
  const script = `try{if(app.documents.length<2)throw "Tab Missing";app.activeDocument=app.documents[app.documents.length-1];app.open("${currentImgBase64}","INJECTED_LAYER",true);app.echoToOE(JSON.stringify({type:"INJECT_DONE"}));}catch(e){app.echoToOE(JSON.stringify({type:"ERROR",msg:e.toString()}));}`;
  sendMessage(script);
}

function processAndSave() {
  APP_STATE = "SAVING";
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
      app.echoToOE(JSON.stringify({type:"SLOT_SAVED"}));
    }catch(e){app.echoToOE(JSON.stringify({type:"ERROR",msg:e.toString()}));}
  `;
  sendMessage(script);
}

// ── 4. GALLERY ───────────────────────────────────────────────

function addResultToGallery(blob) {
  if (APP_STATE === "STOPPED" || APP_STATE === "IDLE") return;
  const plan = batchPlan[batchPlanIndex];
  if (!plan) return;
  const key     = buildInputKey(plan.inputIdx);
  const psdName = psdQueue[plan.psdIdx]?.name;
  if (!generatedMockups[key]) generatedMockups[key] = [];
  const url = URL.createObjectURL(blob);
  generatedMockups[key].push({ psd: psdName, blob, url, error: false });
  totalGenerated++;
  updateGroupUI(key, plan);
  // Mark PSD done if this was its last input
  const box = document.getElementById(`psd-box-${plan.psdIdx}`);
  const allDone = batchPlan.slice(batchPlanIndex + 1).every(p => p.psdIdx !== plan.psdIdx);
  if (box && allDone) { box.classList.remove("active"); box.classList.add("done"); }
  batchPlanIndex++;
  loadNextPlanItem();
}

function addErrorToGallery() {
  const plan = batchPlan[batchPlanIndex];
  if (!plan) { batchPlanIndex++; loadNextPlanItem(); return; }
  const key     = buildInputKey(plan.inputIdx);
  const psdName = psdQueue[plan.psdIdx]?.name;
  if (!generatedMockups[key]) generatedMockups[key] = [];
  generatedMockups[key].push({ psd: psdName, blob: null, url: null, error: true });
  totalErrors++;
  updateGroupUI(key, plan);
  batchPlanIndex++;
  setTimeout(loadNextPlanItem, 200);
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
    card.innerHTML = `
      <div class="error-card-inner">
        <div class="error-icon">✕</div>
        <div class="error-label">${psdName.replace(/\.psd$/i,'')}</div>
        <div class="error-sub">Failed</div>
      </div>`;
    scroll.appendChild(card);
  } else {
    const rawPsd    = psdName.replace(/\.psd$/i,'');
    // Get the image name from the slot assigned to this PSD in this plan item
    const planItem  = batchPlan[batchPlanIndex - 1] || batchPlan[batchPlanIndex];
    const psdSlots  = planItem?.slots || [];
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
      <img src="${latest.url}" alt="${rawBase}" loading="lazy" class="mockup-thumb">
      <div class="card-actions">
        <button class="card-btn card-btn-copy" title="Copy to clipboard">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>
        </button>
        <button class="card-btn card-btn-dl" title="Download">↓</button>
      </div>`;
    card.querySelector('.mockup-thumb').addEventListener('click', () => openLightbox(latest.url, namedBase));
    card.querySelector('.card-btn-copy').addEventListener('click', e => { e.stopPropagation(); copyToClipboard(latest.url); });
    card.querySelector('.card-btn-dl').addEventListener('click', e => { e.stopPropagation(); downloadSingle(latest.url, namedBase, ext); });
    const dlBtn = document.getElementById(`dlbtn-${safeId}`);
    if (dlBtn) { dlBtn.disabled = false; dlBtn.onclick = () => downloadSingle(latest.url, namedBase, ext); }
    scroll.appendChild(card);
  }
}

// ── 5. FINISH / STOP / CLEAR ──────────────────────────────────

function finishAll() {
  if (APP_STATE === "STOPPED") return;
  const ok = totalGenerated - totalErrors;
  setProgress(100, "Complete!");
  setPip("Done", "done");

  // Re-enable everything
  btnGen.disabled = false;
  inPsd.disabled  = false;
  inImg.disabled  = false;
  const bc = btnGen.querySelector('.btn-content');
  if (bc) bc.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><polygon points="5,3 19,12 5,21"/></svg> New Batch`;

  document.getElementById("btn-stop").classList.add("hidden");
  document.getElementById("btn-run-again").classList.remove("hidden");
  btnDownloadAll.disabled = false;

  document.querySelectorAll('.btn-small[disabled]').forEach(btn => {
    const sid  = btn.id.replace('dlbtn-','');
    const name = Object.keys(generatedMockups).find(k => CSS.escape(k.replace(/\W/g,'_')) === sid);
    if (name && generatedMockups[name].some(e => !e.error)) btn.disabled = false;
  });

  APP_STATE = "DONE";

  const msg = totalErrors > 0
    ? `✓ ${ok} mockup${ok !== 1 ? 's' : ''} generated — ${totalErrors} failed`
    : `✓ Batch complete — ${ok} mockup${ok !== 1 ? 's' : ''} generated`;
  showToast(msg, totalErrors > 0 ? 'warn' : 'success');

  // Resume any pending PSD previews
  if (previewQueue.length) schedulePsdPreviews();
}

function stopBatch() {
  if (!BUSY_STATES.includes(APP_STATE)) return;
  APP_STATE = "STOPPED";
  if (pollInterval) { clearInterval(pollInterval); clearTimeout(pollInterval); }
  setPip("Stopped", "error");
  progressStrip.classList.add("hidden");
  const bc = btnGen.querySelector('.btn-content');
  if (bc) bc.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="currentColor"><polygon points="5,3 19,12 5,21"/></svg> Start Batch`;
  btnGen.disabled = false;
  inPsd.disabled  = false;
  inImg.disabled  = false;
  document.getElementById("btn-stop").classList.add("hidden");
  document.querySelectorAll('.btn-small[disabled]').forEach(btn => {
    const sid  = btn.id.replace('dlbtn-','');
    const name = Object.keys(generatedMockups).find(k => CSS.escape(k.replace(/\W/g,'_')) === sid);
    if (name && generatedMockups[name].some(e => !e.error)) btn.disabled = false;
  });
  updateGenerateButton();
  showToast('Batch stopped. Partial results available below.', 'warn');
}

function clearMockupsSilent() {
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
  panelImg.classList.add("disabled"); inImg.disabled = true;
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
};

// ── 6. DOWNLOADS ──────────────────────────────────────────────

window.downloadSingle = (url, name, ext = 'png') => {
  const a = document.createElement("a"); a.href = url; a.download = `${name}.${ext}`; a.click();
};

window.downloadGroup = name => {
  const zip = new JSZip();
  const ext = getExt();
  (generatedMockups[name]||[]).filter(m => !m.error && m.blob).forEach(m => {
    const namedBase = applyNaming(m.psd.replace(/\.psd$/i,''));
    zip.file(`${namedBase}.${ext}`, m.blob);
  });
  zip.generateAsync({type:"blob"}).then(c => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(c);
    a.download = `${applyNaming(name.replace(/\.\w+$/,''))}_Group.zip`;
    a.click();
  });
};

function downloadAllZip() {
  const zip = new JSZip();
  const ext = getExt();
  Object.keys(generatedMockups).forEach(name => {
    const folder = zip.folder(sanitiseFilename(applyNaming(name.replace(/\.\w+$/, ''))));
    (generatedMockups[name]||[]).filter(m => !m.error && m.blob).forEach(m => {
      folder.file(`${applyNaming(m.psd.replace(/\.psd$/i,''))}.${ext}`, m.blob);
    });
  });
  zip.generateAsync({type:"blob"}).then(c => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(c); a.download = "All_Mockups.zip"; a.click();
  });
}

// ── 7. MESSAGE ROUTER ────────────────────────────────────────

window.onmessage = function(e) {

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
    setPip("Ready", "ready");
    updateGenerateButton();
    return;
  }

  // ── Batch pipeline ───────────────────────────────────────
  if (APP_STATE === "STOPPED") return;

  if (e.data === "done" && APP_STATE === "LOADING_PSD") { scanPsd(); return; }

  if (typeof e.data === "string") {
    try {
      const msg = JSON.parse(e.data);
      if      (msg.type === "CLEANUP_DONE")                uploadPsdFile();
      else if (msg.type === "SCAN_OK")                     { clearInterval(pollInterval); clearTimeout(pollInterval); startSlotLoop(); }
      else if (msg.type === "MULTI_SO")                    { clearInterval(pollInterval); clearTimeout(pollInterval); startSlotLoop(); }
      else if (msg.type === "POLL_DOCS" && msg.count >= 2) { clearInterval(pollInterval); clearTimeout(pollInterval); setTimeout(verifyAndInject, 800); }
      else if (msg.type === "INJECT_DONE")                 setTimeout(processAndSave, 200);
      else if (msg.type === "SLOT_SAVED")                  { currentSlotIndex++; setTimeout(processNextSlot, 200); }
      else if (msg.type === "ERROR")                       addErrorToGallery();
    } catch(_) {}
    return;
  }

  if (e.data instanceof ArrayBuffer) {
    // Drop stale ArrayBuffers from a previous or stopped batch
    if (APP_STATE === "STOPPED" || APP_STATE === "IDLE") return;
    if (window._currentBatchId !== batchId) return;
    const fmt  = window.outputFormat || 'png';
    const mime = fmt.startsWith('jpg') ? 'image/jpeg' : 'image/png';
    addResultToGallery(new Blob([e.data], {type: mime}));
  }
};
