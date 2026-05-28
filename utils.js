// ── Utilidades compartidas AClímate ──────────────────────────────────

// ── Gestión de OTs ────────────────────────────────────────────────────

// Genera un número de OT correlativo (OT-AÑO-NNN)
function generarNumOT() {
  const año = new Date().getFullYear();
  const ots = JSON.parse(localStorage.getItem('ots_local') || '[]');
  const delAño = ots.filter(o => o.id.startsWith('OT-' + año));
  const siguiente = String(delAño.length + 1).padStart(3, '0');
  return `OT-${año}-${siguiente}`;
}

// Guarda o actualiza una OT en localStorage
function guardarOT(ot) {
  const ots = JSON.parse(localStorage.getItem('ots_local') || '[]');
  const idx = ots.findIndex(o => o.id === ot.id);
  if (idx >= 0) { ots[idx] = ot; } else { ots.unshift(ot); }
  localStorage.setItem('ots_local', JSON.stringify(ots));
}

// Carga OTs: primero intenta GAS, fallback a localStorage
async function cargarOTs() {
  // Siempre devuelve al menos las locales de inmediato
  let ots = JSON.parse(localStorage.getItem('ots_local') || '[]');

  if (!navigator.onLine || CONFIG.GAS_URL.includes('TU_SCRIPT_ID')) return ots;

  try {
    const resp = await fetch(CONFIG.GAS_URL + '?action=ots', { cache: 'no-store' });
    const json = await resp.json();
    if (json.status === 'ok' && Array.isArray(json.data)) {
      // Merge: prioridad a las del servidor
      const ids = new Set(json.data.map(o => o.id));
      const localesNuevas = ots.filter(o => !ids.has(o.id));
      const merged = [...json.data, ...localesNuevas];
      localStorage.setItem('ots_local', JSON.stringify(merged));
      ots = merged;
    }
  } catch(e) {}

  return ots;
}

// Poblar un <select> con las OTs abiertas
async function poblarSelectOT(selectEl, { incluirGeneral = false } = {}) {
  selectEl.innerHTML = '<option value="">— Cargando OTs... —</option>';
  const ots = await cargarOTs();
  const abiertas = ots.filter(o => o.estado === 'Abierta');

  selectEl.innerHTML = '<option value="">— Seleccionar OT —</option>';
  if (incluirGeneral) {
    const opt = document.createElement('option');
    opt.value = 'GENERAL';
    opt.textContent = 'Gasto general operativo';
    opt.dataset.cliente = 'N/A';
    opt.dataset.tipo = '';
    selectEl.appendChild(opt);
  }
  if (!abiertas.length) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = '— No hay OTs abiertas —';
    opt.disabled = true;
    selectEl.appendChild(opt);
    return;
  }
  abiertas.forEach(ot => {
    const opt = document.createElement('option');
    opt.value = ot.id;
    opt.textContent = `${ot.id} — ${ot.cliente}`;
    opt.dataset.cliente = ot.cliente;
    opt.dataset.tipo    = ot.tipo || '';
    opt.dataset.tecnico = ot.tecnico || '';
    selectEl.appendChild(opt);
  });
}

// Conecta un select de OT con los campos auto-fill (cliente, tipo, técnico)
function conectarOTSelect(selectEl, { clienteDisplay, clienteHidden, tipoSelect, tecnicoSelect } = {}) {
  selectEl.addEventListener('change', function() {
    const opt = this.options[this.selectedIndex];
    const cliente = opt?.dataset?.cliente || '';
    const tipo    = opt?.dataset?.tipo    || '';
    const tecnico = opt?.dataset?.tecnico || '';

    if (clienteDisplay) clienteDisplay.textContent = cliente || '—';
    if (clienteHidden)  clienteHidden.value = cliente;
    if (tipoSelect && tipo) tipoSelect.value = tipo;
    if (tecnicoSelect && tecnico) tecnicoSelect.value = tecnico;
  });
}

// ── Auto-guardado ─────────────────────────────────────────────────────
function autoSave(formId, formEl) {
  const data = {};
  const inputs = formEl.querySelectorAll('input, select, textarea');
  inputs.forEach(el => {
    if (el.name) data[el.name] = el.value;
  });
  localStorage.setItem('draft_' + formId, JSON.stringify(data));
  const bar = document.getElementById('autosave-bar');
  if (bar) {
    bar.textContent = '💾 Borrador guardado — ' + new Date().toLocaleTimeString('es-CR');
    bar.classList.add('show');
  }
}

function restoreDraft(formId, formEl) {
  const raw = localStorage.getItem('draft_' + formId);
  if (!raw) return;
  try {
    const data = JSON.parse(raw);
    Object.entries(data).forEach(([name, val]) => {
      const el = formEl.querySelector(`[name="${name}"]`);
      if (el) el.value = val;
    });
  } catch (e) {}
}

function clearDraft(formId) {
  localStorage.removeItem('draft_' + formId);
}

// ── Toast ──────────────────────────────────────────────────────────────
function showToast(msg, type = 'default', duration = 3500) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  t.className = 'show ' + type;
  clearTimeout(t._timer);
  t._timer = setTimeout(() => { t.className = ''; }, duration);
}

// ── Loading overlay ─────────────────────────────────────────────────
function showLoading(msg = 'Enviando formulario...') {
  const el = document.getElementById('loading-overlay');
  if (!el) return;
  el.querySelector('p').textContent = msg;
  el.classList.add('show');
}
function hideLoading() {
  const el = document.getElementById('loading-overlay');
  if (el) el.classList.remove('show');
}

// ── Pantalla de éxito ──────────────────────────────────────────────
function showSuccess(msg = 'Formulario enviado correctamente') {
  hideLoading();
  const sc = document.getElementById('success-screen');
  if (!sc) return;
  const p = sc.querySelector('p');
  if (p) p.textContent = msg;
  sc.classList.add('show');
}

// ── Validación de formulario ────────────────────────────────────────
function validarForm(formEl) {
  let ok = true;
  formEl.querySelectorAll('[required]').forEach(el => {
    const campo = el.closest('.campo');
    if (!campo) return;
    const val = el.value.trim();
    if (!val) {
      campo.classList.add('error');
      ok = false;
    } else {
      campo.classList.remove('error');
    }
  });

  // Validar ítems con clase item-row
  formEl.querySelectorAll('.item-row').forEach(row => {
    const selEl = row.querySelector('.ac-hidden');
    const qtyEl = row.querySelector('.item-qty');
    if (selEl && !selEl.value) {
      row.querySelector('.autocomplete-input').style.borderColor = 'var(--rojo)';
      ok = false;
    }
    if (qtyEl && (!qtyEl.value || parseInt(qtyEl.value) < 1)) {
      qtyEl.style.borderColor = 'var(--rojo)';
      ok = false;
    }
  });

  // Validar fotos obligatorias
  formEl.querySelectorAll('.foto-required').forEach(el => {
    const campo = el.closest('.campo');
    if (!campo) return;
    if (!el._hasFile) {
      campo.classList.add('error');
      ok = false;
    } else {
      campo.classList.remove('error');
    }
  });

  if (!ok) showToast('Completá todos los campos obligatorios', 'warn');
  return ok;
}

// ── Comprimir foto ──────────────────────────────────────────────────
function comprimirFoto(file, maxDim, quality) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        let w = img.width, h = img.height;
        if (w > maxDim || h > maxDim) {
          if (w > h) { h = Math.round(h * maxDim / w); w = maxDim; }
          else       { w = Math.round(w * maxDim / h); h = maxDim; }
        }
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ── Setup de campo de foto ───────────────────────────────────────────
function setupFotoField(inputEl, previewContainer, onUpdate) {
  inputEl._hasFile = false;
  inputEl._fotos = [];

  function actualizarContador() {
    const n = inputEl._fotos.length;
    const txt = inputEl.closest('.foto-zona')?.querySelector('.nombre-archivo');
    if (txt) txt.textContent = n === 0 ? '' : n === 1 ? '1 foto agregada' : n + ' fotos agregadas';
    inputEl._hasFile = n > 0;
    inputEl.closest('.campo')?.classList.remove('error');
    if (onUpdate) onUpdate(n);
  }

  function renderPreviews() {
    if (!previewContainer) return;
    previewContainer.innerHTML = '';
    inputEl._fotos.forEach((foto, idx) => {
      const wrap = document.createElement('div');
      wrap.style.cssText = 'position:relative;display:inline-block;margin:2px;';
      const img = document.createElement('img');
      img.src = foto.data;
      img.style.cssText = 'display:block;border-radius:4px;';
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = '✕';
      btn.style.cssText = 'position:absolute;top:2px;right:2px;background:rgba(0,0,0,.6);color:#fff;border:none;border-radius:50%;width:20px;height:20px;font-size:.65rem;cursor:pointer;line-height:1;padding:0;';
      btn.onclick = () => {
        inputEl._fotos.splice(idx, 1);
        renderPreviews();
        actualizarContador();
      };
      wrap.appendChild(img);
      wrap.appendChild(btn);
      previewContainer.appendChild(wrap);
    });
  }

  inputEl.addEventListener('change', async () => {
    const files = Array.from(inputEl.files);
    if (!files.length) return;
    for (const file of files) {
      try {
        const b64 = await comprimirFoto(file, CONFIG.FOTO_MAX_DIM, CONFIG.FOTO_QUALITY);
        inputEl._fotos.push({ nombre: file.name, data: b64 });
      } catch(e) {}
    }
    renderPreviews();
    actualizarContador();
    // Limpiar el input para permitir seleccionar el mismo archivo de nuevo
    inputEl.value = '';
  });
}

// ── Autocomplete de catálogo ─────────────────────────────────────────
function crearAutocomplete(wrap, catalogo, onSelect) {
  const input    = wrap.querySelector('.autocomplete-input');
  const dropdown = wrap.querySelector('.autocomplete-dropdown');
  const hidden   = wrap.querySelector('.ac-hidden');
  const selected = wrap.querySelector('.ac-selected');
  let focusIdx = -1;

  function render(query) {
    const q = query.toLowerCase();
    const matches = catalogo
      .filter(it => it.codigo.toLowerCase().includes(q) || it.nombre.toLowerCase().includes(q))
      .slice(0, 50);
    dropdown.innerHTML = '';
    focusIdx = -1;
    if (!matches.length) { dropdown.classList.remove('open'); return; }
    matches.forEach((it, i) => {
      const div = document.createElement('div');
      div.className = 'ac-option';
      div.innerHTML =
        `<div class="ac-code">${it.codigo}</div>` +
        `<div class="ac-name">${it.nombre}</div>` +
        `<div class="ac-cat">${it.categoria}</div>`;
      div.addEventListener('mousedown', e => {
        e.preventDefault();
        seleccionar(it);
      });
      dropdown.appendChild(div);
    });
    dropdown.classList.add('open');
  }

  function seleccionar(it) {
    hidden.value = it.codigo;
    input.value  = it.codigo + ' — ' + it.nombre;
    dropdown.classList.remove('open');
    if (selected) {
      selected.textContent = '✓ ' + it.nombre;
      selected.classList.add('show');
    }
    input.closest('.campo')?.classList.remove('error');
    input.style.borderColor = '';
    if (onSelect) onSelect(it);
  }

  input.addEventListener('input', () => render(input.value));
  input.addEventListener('focus', () => { if (input.value) render(input.value); });
  input.addEventListener('blur',  () => setTimeout(() => dropdown.classList.remove('open'), 150));
  input.addEventListener('keydown', e => {
    const opts = dropdown.querySelectorAll('.ac-option');
    if (e.key === 'ArrowDown') { focusIdx = Math.min(focusIdx + 1, opts.length - 1); }
    else if (e.key === 'ArrowUp') { focusIdx = Math.max(focusIdx - 1, 0); }
    else if (e.key === 'Enter' && focusIdx >= 0) { e.preventDefault(); opts[focusIdx]?.dispatchEvent(new Event('mousedown')); }
    else if (e.key === 'Escape') { dropdown.classList.remove('open'); }
    opts.forEach((o, i) => o.classList.toggle('focused', i === focusIdx));
    if (focusIdx >= 0) opts[focusIdx]?.scrollIntoView({ block: 'nearest' });
  });
}

// ── Multi-ítems ──────────────────────────────────────────────────────
let _itemCounter = 0;

function crearItemRow(catalogoUsado, container) {
  _itemCounter++;
  const idx = _itemCounter;
  const row = document.createElement('div');
  row.className = 'item-row';
  row.dataset.idx = idx;
  row.innerHTML = `
    <div class="item-row-header">
      <span class="item-num">Ítem #${idx}</span>
      <button type="button" class="btn-rm-item" title="Eliminar">✕</button>
    </div>
    <div class="campo" style="margin-bottom:0">
      <label>Código / Descripción <span class="req">*</span></label>
      <div class="autocomplete-wrap">
        <input class="autocomplete-input" type="text" placeholder="Buscar código o nombre..." autocomplete="off">
        <input class="ac-hidden" type="hidden" name="item_codigo_${idx}">
        <div class="autocomplete-dropdown"></div>
      </div>
      <span class="ac-selected"></span>
    </div>
    <div class="item-qty-row">
      <label>Cantidad <span class="req">*</span></label>
      <input class="item-qty" type="number" name="item_qty_${idx}" min="1" step="1" placeholder="0">
    </div>`;

  row.querySelector('.btn-rm-item').addEventListener('click', () => {
    row.remove();
    renumerarItems(container);
  });

  crearAutocomplete(row.querySelector('.autocomplete-wrap'), catalogoUsado, null);
  container.appendChild(row);
}

function renumerarItems(container) {
  container.querySelectorAll('.item-row').forEach((row, i) => {
    const num = row.querySelector('.item-num');
    if (num) num.textContent = 'Ítem #' + (i + 1);
  });
}

function recogerItems(container) {
  const items = [];
  container.querySelectorAll('.item-row').forEach(row => {
    const codigo = row.querySelector('.ac-hidden')?.value || '';
    const nombre = row.querySelector('.ac-selected')?.textContent.replace('✓ ', '') || '';
    const qty    = row.querySelector('.item-qty')?.value || '';
    if (codigo && qty) items.push({ codigo, nombre, cantidad: parseInt(qty) });
  });
  return items;
}

// ── Enviar a Google Apps Script ──────────────────────────────────────
async function enviarAGAS(payload) {
  const resp = await fetch(CONFIG.GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(payload)
  });
  const json = await resp.json();
  if (json.status !== 'ok') throw new Error(json.message || 'Error desconocido');
  return json;
}

// Guarda en cola offline si no hay conexión
async function enviarConFallback(payload, formId) {
  if (!navigator.onLine) {
    guardarEnCola(formId, payload);
    showToast('Sin conexión — guardado localmente', 'warn');
    showSuccess('Guardado localmente. Se enviará cuando haya conexión.');
    return;
  }
  try {
    await enviarAGAS(payload);
  } catch (e) {
    guardarEnCola(formId, payload);
    throw e;
  }
}

function guardarEnCola(formId, payload) {
  const cola = JSON.parse(localStorage.getItem('cola_envio') || '[]');
  cola.push({ formId, payload, ts: Date.now() });
  localStorage.setItem('cola_envio', JSON.stringify(cola));
}

// Reintenta envíos pendientes al recuperar conexión
window.addEventListener('online', async () => {
  const cola = JSON.parse(localStorage.getItem('cola_envio') || '[]');
  if (!cola.length) return;
  const pendientes = [...cola];
  localStorage.removeItem('cola_envio');
  for (const item of pendientes) {
    try { await enviarAGAS(item.payload); }
    catch (e) { guardarEnCola(item.formId, item.payload); }
  }
});

// ── Fecha/hora actual formateada (zona Costa Rica) ───────────────────
const TZ = 'America/Costa_Rica';

function fechaHoy() {
  return new Date().toLocaleDateString('en-CA', { timeZone: TZ }); // YYYY-MM-DD
}

function horaAhora() {
  return new Date().toLocaleTimeString('es-CR', {
    timeZone: TZ, hour: '2-digit', minute: '2-digit', hour12: true
  });
}

function timestampCR() {
  return new Date().toLocaleString('es-CR', {
    timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true
  });
}
