'use strict';

// ── State ────────────────────────────────────────────────────────────────────
let state = { possession: [], used: [], rendidos: [] };
// item shape: { prefix, number, code, type: 'CABLE'|'BOTELLA', destino?, fecha?, nota? }

let filters     = { possession: '', used: '', rendidos: '' };
let typeFilters = { possession: 'ALL', used: 'ALL', rendidos: 'ALL' };
let selected    = { possession: new Set(), used: new Set() };
let pendingDelete = null;

// ── Persistence ──────────────────────────────────────────────────────────────
function loadState() {
  try {
    const s = localStorage.getItem('worktracker-v1');
    if (s) state = JSON.parse(s);
    if (!state.rendidos) state.rendidos = [];
    ['possession','used','rendidos'].forEach(list => {
      state[list].forEach(i => { if (!i.type) i.type = 'CABLE'; });
    });
  } catch (_) {}
}

function saveState() {
  localStorage.setItem('worktracker-v1', JSON.stringify(state));
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function makeItem(prefix, numStr, type) {
  const p = prefix.trim().toUpperCase();
  return { prefix: p, number: numStr, code: p + numStr, type };
}

function allCodes() {
  return new Set([...state.possession, ...state.used, ...state.rendidos].map(i => i.code));
}

function sortItems(items) {
  return [...items].sort((a, b) => {
    const pc = a.prefix.localeCompare(b.prefix);
    if (pc !== 0) return pc;
    const tc = a.type.localeCompare(b.type);
    if (tc !== 0) return tc;
    return parseInt(a.number, 10) - parseInt(b.number, 10);
  });
}

function groupByPrefix(items) {
  const g = {};
  items.forEach(item => { (g[item.prefix] = g[item.prefix] || []).push(item); });
  return g;
}

function escapeHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function typeBadgeHtml(type) {
  const cls  = type === 'CABLE' ? 'type-cable' : 'type-botella';
  const icon = type === 'CABLE' ? '🔌' : '🍾';
  return `<span class="type-badge ${cls}">${icon} ${type}</span>`;
}

function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

// ── Add items ────────────────────────────────────────────────────────────────
function addToPossession(items) {
  const existing = allCodes();
  let added = 0, dupes = 0;
  items.forEach(item => {
    if (existing.has(item.code)) { dupes++; }
    else { state.possession.push(item); existing.add(item.code); added++; }
  });
  if (added) saveState();
  return { added, dupes };
}

// ── Destino dialog ────────────────────────────────────────────────────────────
let pendingMove = null;

function askDestino(codes) {
  pendingMove = { codes };
  const n = codes.length;
  document.getElementById('destino-title').textContent    = n === 1 ? `Destino — ${codes[0]}` : `Destino (${n} precintos)`;
  document.getElementById('destino-subtitle').textContent = 'Podés dejar el destino en blanco si no aplica.';
  document.getElementById('destino-input').value = '';
  document.getElementById('destino-fecha').value = todayISO();
  document.getElementById('destino-overlay').classList.remove('hidden');
  setTimeout(() => document.getElementById('destino-input').focus(), 80);
}

function executeMove() {
  if (!pendingMove) return;
  const destino = document.getElementById('destino-input').value.trim();
  const fecha   = document.getElementById('destino-fecha').value || todayISO();
  const { codes } = pendingMove;
  pendingMove = null;
  document.getElementById('destino-overlay').classList.add('hidden');
  codes.forEach(code => {
    const idx = state.possession.findIndex(i => i.code === code);
    if (idx !== -1) {
      const item = state.possession.splice(idx, 1)[0];
      item.destino = destino;
      item.fecha   = fecha;
      state.used.push(item);
    }
  });
  selected.possession.clear();
  saveState(); render();
  showToast(`${codes.length} precinto(s) marcado(s) como utilizado`);
}

document.getElementById('destino-ok').addEventListener('click', executeMove);

document.getElementById('destino-cancel').addEventListener('click', () => {
  pendingMove = null;
  document.getElementById('destino-overlay').classList.add('hidden');
});

document.getElementById('destino-overlay').addEventListener('click', e => {
  if (e.target === e.currentTarget) { pendingMove = null; e.currentTarget.classList.add('hidden'); }
});

document.getElementById('destino-input').addEventListener('keydown', e => {
  if (e.key === 'Enter')  executeMove();
  if (e.key === 'Escape') { pendingMove = null; document.getElementById('destino-overlay').classList.add('hidden'); }
});

// ── Nota dialog ──────────────────────────────────────────────────────────────
let pendingNota = null;

function askNota(code) {
  const item = [...state.used, ...state.rendidos].find(i => i.code === code);
  if (!item) return;
  pendingNota = code;
  document.getElementById('nota-title').textContent = `Nota — ${code}`;
  document.getElementById('nota-input').value = item.nota || '';
  document.getElementById('nota-overlay').classList.remove('hidden');
  setTimeout(() => document.getElementById('nota-input').focus(), 80);
}

function executeNota() {
  if (!pendingNota) return;
  const item = [...state.used, ...state.rendidos].find(i => i.code === pendingNota);
  if (item) item.nota = document.getElementById('nota-input').value.trim();
  pendingNota = null;
  document.getElementById('nota-overlay').classList.add('hidden');
  saveState(); render(); renderHistorial();
}

document.getElementById('nota-ok').addEventListener('click', executeNota);
document.getElementById('nota-cancel').addEventListener('click', () => {
  pendingNota = null;
  document.getElementById('nota-overlay').classList.add('hidden');
});
document.getElementById('nota-overlay').addEventListener('click', e => {
  if (e.target === e.currentTarget) { pendingNota = null; e.currentTarget.classList.add('hidden'); }
});
document.getElementById('nota-input').addEventListener('keydown', e => {
  if (e.key === 'Escape') { pendingNota = null; document.getElementById('nota-overlay').classList.add('hidden'); }
});

// ── Move / Return ─────────────────────────────────────────────────────────────
function moveToUsed(code) { askDestino([code]); }

function returnToPossession(code) {
  const idx = state.used.findIndex(i => i.code === code);
  if (idx === -1) return;
  const item = state.used.splice(idx, 1)[0];
  delete item.destino; delete item.fecha; delete item.nota;
  state.possession.push(item);
  selected.used.delete(code);
  saveState(); render();
}

function moveSelectedToUsed() {
  if (!selected.possession.size) { showToast('No hay elementos seleccionados'); return; }
  askDestino([...selected.possession]);
}

function returnSelectedToPossession() {
  if (!selected.used.size) { showToast('No hay elementos seleccionados'); return; }
  const codes = [...selected.used];
  codes.forEach(code => {
    const idx = state.used.findIndex(i => i.code === code);
    if (idx !== -1) {
      const item = state.used.splice(idx, 1)[0];
      delete item.destino; delete item.fecha; delete item.nota;
      state.possession.push(item);
    }
  });
  selected.used.clear();
  saveState(); render();
  showToast(`${codes.length} precinto(s) devuelto(s) a En posesión`);
}

// ── Rendir ────────────────────────────────────────────────────────────────────
function rendirSelected() {
  if (!selected.used.size) { showToast('No hay elementos seleccionados'); return; }
  const codes = [...selected.used];
  codes.forEach(code => {
    const idx = state.used.findIndex(i => i.code === code);
    if (idx !== -1) {
      const item = state.used.splice(idx, 1)[0];
      item.fechaRendido = todayISO();
      state.rendidos.unshift(item); // más reciente primero
    }
  });
  selected.used.clear();
  saveState(); render();
  showToast(`${codes.length} precinto(s) rendido(s) al historial`);
}

document.getElementById('rendir-btn').addEventListener('click', rendirSelected);

// ── Historial modal ───────────────────────────────────────────────────────────
function openHistorial() {
  document.getElementById('historial-overlay').classList.remove('hidden');
  renderHistorial();
}

function closeHistorial() {
  document.getElementById('historial-overlay').classList.add('hidden');
}

document.getElementById('btn-historial').addEventListener('click', openHistorial);
document.getElementById('historial-close').addEventListener('click', closeHistorial);
document.getElementById('historial-overlay').addEventListener('click', e => {
  if (e.target === e.currentTarget) closeHistorial();
});

document.getElementById('search-rendidos').addEventListener('input', e => {
  filters.rendidos = e.target.value;
  renderHistorial();
});

function renderHistorial() {
  const container = document.getElementById('list-rendidos');
  const text = filters.rendidos.toLowerCase();
  const type = typeFilters.rendidos;

  const items = state.rendidos.filter(i =>
    (!text || i.code.toLowerCase().includes(text) ||
              (i.destino||'').toLowerCase().includes(text) ||
              (i.nota||'').toLowerCase().includes(text)) &&
    (type === 'ALL' || i.type === type)
  );

  const nAll    = state.rendidos.length;
  const nCable  = state.rendidos.filter(i => i.type === 'CABLE').length;
  const nBottle = state.rendidos.filter(i => i.type === 'BOTELLA').length;
  document.getElementById('historial-count').textContent = `${nAll} precinto(s)`;
  document.getElementById('type-all-r').textContent    = `Todos (${nAll})`;
  document.getElementById('type-cable-r').textContent  = `🔌 CABLE (${nCable})`;
  document.getElementById('type-bottle-r').textContent = `🍾 BOTELLA (${nBottle})`;

  if (!items.length) {
    container.innerHTML = `<div class="empty-state">
      <div class="empty-icon">${text ? '🔍' : '📭'}</div>
      <p>${text ? `Sin resultados para "<strong>${escapeHtml(text)}</strong>"` : 'No hay precintos rendidos todavía.'}</p>
    </div>`;
    return;
  }

  let html = '';
  items.forEach(item => {
    const safeCode = escapeHtml(item.code);
    html += `<div class="item item-rendido" data-code="${safeCode}" data-list="rendidos">
      <div class="item-main">
        <span class="item-code">${safeCode}</span>
        <div class="item-meta">
          ${typeBadgeHtml(item.type)}
          ${item.destino     ? `<span class="item-destino">📍 ${escapeHtml(item.destino)}</span>` : ''}
          ${item.fecha       ? `<span class="item-fecha">🗓 ${escapeHtml(item.fecha)}</span>` : ''}
          ${item.fechaRendido? `<span class="item-fecha-rendido">✅ ${escapeHtml(item.fechaRendido)}</span>` : ''}
        </div>
        ${item.nota ? `<div class="item-nota">📝 ${escapeHtml(item.nota)}</div>` : ''}
      </div>
      <div class="item-actions">
        <button class="btn-sm btn-nota"      data-action="nota"   title="Agregar nota">📝</button>
        <button class="btn-sm btn-secondary" data-action="return" title="Volver a Utilizado">↩</button>
        <button class="btn-sm btn-danger"    data-action="delete" title="Eliminar registro">✕</button>
      </div>
    </div>`;
  });
  container.innerHTML = html;
}

// Event delegation for historial list
document.getElementById('list-rendidos').addEventListener('click', e => {
  const btn = e.target.closest('button[data-action]');
  if (!btn) return;
  const item = btn.closest('.item');
  if (!item) return;
  const code   = item.dataset.code;
  const action = btn.dataset.action;
  if (action === 'nota')   askNota(code);
  if (action === 'return') returnRendidoToUsed(code);
  if (action === 'delete') askDeleteRendido(code);
});

function returnRendidoToUsed(code) {
  const idx = state.rendidos.findIndex(i => i.code === code);
  if (idx === -1) return;
  const item = state.rendidos.splice(idx, 1)[0];
  delete item.fechaRendido;
  state.used.push(item);
  saveState(); render(); renderHistorial();
  showToast(`${code} devuelto a Utilizado`);
}

function askDeleteRendido(code) {
  pendingDelete = { listName: 'rendidos', codes: [code] };
  openConfirm('Eliminar registro', `¿Eliminar el registro <strong>${escapeHtml(code)}</strong> del historial? Esta acción no se puede deshacer.`);
}

// ── Delete ────────────────────────────────────────────────────────────────────
function askDeleteSingle(listName, code) {
  pendingDelete = { listName, codes: [code] };
  openConfirm('Eliminar precinto', `¿Eliminar <strong>${escapeHtml(code)}</strong>? Esta acción no se puede deshacer.`);
}

function askDeleteSelected(listName) {
  const count = selected[listName].size;
  if (!count) { showToast('No hay elementos seleccionados'); return; }
  pendingDelete = { listName, codes: [...selected[listName]] };
  openConfirm('Eliminar selección', `¿Eliminar <strong>${count} precinto(s)</strong> seleccionados? Esta acción no se puede deshacer.`);
}

function executeDelete() {
  if (!pendingDelete) return;
  const { listName, codes } = pendingDelete;
  const codeSet = new Set(codes);
  state[listName] = state[listName].filter(i => !codeSet.has(i.code));
  codes.forEach(c => selected[listName]?.delete(c));
  pendingDelete = null;
  saveState(); render();
  if (listName === 'rendidos') renderHistorial();
  showToast(`${codes.length} precinto(s) eliminado(s)`);
}

// ── Selection ─────────────────────────────────────────────────────────────────
function toggleSelect(listName, code, checked) {
  checked ? selected[listName].add(code) : selected[listName].delete(code);
  const item = document.querySelector(`.item[data-code="${CSS.escape(code)}"][data-list="${listName}"]`);
  if (item) item.classList.toggle('selected', checked);
}

function getVisibleItems(listName) {
  const src  = state[listName];
  const text = filters[listName].toLowerCase();
  const type = typeFilters[listName];
  return src.filter(i =>
    (!text || i.code.toLowerCase().includes(text)) &&
    (type === 'ALL' || i.type === type)
  );
}

function toggleSelectAll(listName) {
  const visible = getVisibleItems(listName);
  const allSel  = visible.length > 0 && visible.every(i => selected[listName].has(i.code));
  visible.forEach(i => allSel ? selected[listName].delete(i.code) : selected[listName].add(i.code));
  render();
}

// ── Render lists ──────────────────────────────────────────────────────────────
function renderList(containerId, listName) {
  const container = document.getElementById(containerId);
  const items  = getVisibleItems(listName);
  const sorted = sortItems(items);

  const all     = state[listName];
  const nCable  = all.filter(i => i.type === 'CABLE').length;
  const nBottle = all.filter(i => i.type === 'BOTELLA').length;
  const pfx = listName === 'possession' ? 'p' : 'u';
  document.getElementById(`type-all-${pfx}`).textContent    = `Todos (${all.length})`;
  document.getElementById(`type-cable-${pfx}`).textContent  = `🔌 CABLE (${nCable})`;
  document.getElementById(`type-bottle-${pfx}`).textContent = `🍾 BOTELLA (${nBottle})`;

  const filter = filters[listName];
  if (!sorted.length) {
    container.innerHTML = `<div class="empty-state">
      <div class="empty-icon">${filter ? '🔍' : '📭'}</div>
      <p>${filter ? `Sin resultados para "<strong>${escapeHtml(filter)}</strong>"` : 'No hay precintos aquí todavía.'}</p>
    </div>`;
    return;
  }

  const groups   = groupByPrefix(sorted);
  const prefixes = Object.keys(groups).sort((a,b) => a.localeCompare(b));
  const isUsed   = listName === 'used';

  let html = '';
  prefixes.forEach(prefix => {
    const groupItems = groups[prefix];
    const nC = groupItems.filter(i => i.type === 'CABLE').length;
    const nB = groupItems.filter(i => i.type === 'BOTELLA').length;
    const summary = [nC && `🔌 ${nC}`, nB && `🍾 ${nB}`].filter(Boolean).join(' &nbsp; ');
    html += `<div class="prefix-group">
      <div class="prefix-group-header">${escapeHtml(prefix)} &nbsp;·&nbsp; ${summary}</div>`;
    groupItems.forEach(item => {
      const code     = item.code;
      const checked  = selected[listName].has(code);
      const safeCode = escapeHtml(code);
      const actions  = isUsed
        ? `<button class="btn-sm btn-nota"      data-action="nota"   title="Agregar nota">📝</button>
           <button class="btn-sm btn-secondary" data-action="return" title="Devolver a En posesión">↩</button>
           <button class="btn-sm btn-danger"    data-action="delete" title="Eliminar">✕</button>`
        : `<button class="btn-sm btn-success"   data-action="move"   title="Marcar como utilizado">✓</button>
           <button class="btn-sm btn-danger"    data-action="delete" title="Eliminar">✕</button>`;
      html += `<div class="item${checked?' selected':''}" data-code="${safeCode}" data-list="${listName}">
        <input type="checkbox" ${checked?'checked':''} data-cb="${listName}" data-code="${safeCode}" aria-label="${safeCode}">
        <div class="item-main">
          <span class="item-code">${safeCode}</span>
          <div class="item-meta">
            ${typeBadgeHtml(item.type)}
            ${item.destino ? `<span class="item-destino">📍 ${escapeHtml(item.destino)}</span>` : ''}
            ${item.fecha   ? `<span class="item-fecha">🗓 ${escapeHtml(item.fecha)}</span>` : ''}
          </div>
          ${item.nota ? `<div class="item-nota">📝 ${escapeHtml(item.nota)}</div>` : ''}
        </div>
        <div class="item-actions">${actions}</div>
      </div>`;
    });
    html += `</div>`;
  });
  container.innerHTML = html;
}

function render() {
  renderList('list-possession', 'possession');
  renderList('list-used', 'used');
  document.getElementById('badge-possession').textContent = state.possession.length;
  document.getElementById('badge-used').textContent       = state.used.length;
}

// ── Type filter pills ─────────────────────────────────────────────────────────
function setTypeFilter(listName, type) {
  typeFilters[listName] = type;
  if (listName !== 'rendidos') selected[listName].clear();
  document.querySelectorAll(`.type-filter-btn[data-list="${listName}"]`).forEach(btn => {
    btn.classList.toggle('active', btn.dataset.type === type);
  });
  if (listName === 'rendidos') renderHistorial();
  else render();
}

document.querySelectorAll('.type-filter-btn').forEach(btn => {
  btn.addEventListener('click', () => setTypeFilter(btn.dataset.list, btn.dataset.type));
});

// ── Event delegation for lists ────────────────────────────────────────────────
['list-possession','list-used'].forEach(id => {
  document.getElementById(id).addEventListener('click', e => {
    const btn  = e.target.closest('button[data-action]');
    if (!btn) return;
    const item = btn.closest('.item');
    if (!item) return;
    const code     = item.dataset.code;
    const listName = item.dataset.list;
    const action   = btn.dataset.action;
    if (action === 'move')   moveToUsed(code);
    if (action === 'return') returnToPossession(code);
    if (action === 'delete') askDeleteSingle(listName, code);
    if (action === 'nota')   askNota(code);
  });

  document.getElementById(id).addEventListener('change', e => {
    const cb = e.target.closest('input[data-cb]');
    if (!cb) return;
    toggleSelect(cb.dataset.cb, cb.dataset.code, cb.checked);
  });
});

// ── Search ───────────────────────────────────────────────────────────────────
document.getElementById('search-possession').addEventListener('input', e => {
  filters.possession = e.target.value; selected.possession.clear(); render();
});
document.getElementById('search-used').addEventListener('input', e => {
  filters.used = e.target.value; selected.used.clear(); render();
});

// ── Bulk action buttons ───────────────────────────────────────────────────────
document.getElementById('select-all-possession').addEventListener('click', () => toggleSelectAll('possession'));
document.getElementById('select-all-used').addEventListener('click',       () => toggleSelectAll('used'));
document.getElementById('mark-used-btn').addEventListener('click',         moveSelectedToUsed);
document.getElementById('return-btn').addEventListener('click',            returnSelectedToPossession);
document.getElementById('delete-possession-btn').addEventListener('click', () => askDeleteSelected('possession'));
document.getElementById('delete-used-btn').addEventListener('click',       () => askDeleteSelected('used'));

// ── Confirm dialog ───────────────────────────────────────────────────────────
function openConfirm(title, msg) {
  document.getElementById('confirm-title').textContent = title;
  document.getElementById('confirm-message').innerHTML = msg;
  document.getElementById('confirm-overlay').classList.remove('hidden');
}

document.getElementById('confirm-cancel').addEventListener('click', () => {
  pendingDelete = null;
  document.getElementById('confirm-overlay').classList.add('hidden');
});
document.getElementById('confirm-ok').addEventListener('click', () => {
  document.getElementById('confirm-overlay').classList.add('hidden');
  executeDelete();
});
document.getElementById('confirm-overlay').addEventListener('click', e => {
  if (e.target === e.currentTarget) { pendingDelete = null; e.currentTarget.classList.add('hidden'); }
});

// ── Tab navigation ────────────────────────────────────────────────────────────
function switchTab(tabName) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tabName));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.id === 'tab-' + tabName));
}
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => switchTab(btn.dataset.tab));
});

// ── Type selector in forms ────────────────────────────────────────────────────
function getSelectedType(formId) {
  return document.querySelector(`#${formId} .type-toggle-btn.active`)?.dataset.type || 'CABLE';
}
document.querySelectorAll('.type-toggle-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    btn.closest('.type-toggle').querySelectorAll('.type-toggle-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  });
});

// ── Range form ────────────────────────────────────────────────────────────────
document.getElementById('range-form').addEventListener('submit', e => {
  e.preventDefault();
  const prefix   = document.getElementById('range-prefix').value.trim();
  const startStr = document.getElementById('range-start').value.trim();
  const endStr   = document.getElementById('range-end').value.trim();
  const type     = getSelectedType('range-form');

  if (!prefix || !startStr || !endStr) { showToast('Completa todos los campos'); return; }
  const startVal = parseInt(startStr, 10), endVal = parseInt(endStr, 10);
  if (isNaN(startVal) || isNaN(endVal))  { showToast('Los números no son válidos'); return; }
  if (startVal > endVal)                 { showToast('El número inicial debe ser ≤ al final'); return; }
  if (endVal - startVal > 10000)         { showToast('El rango no puede superar 10 000 elementos'); return; }

  const padLen = startStr.length;
  const items  = [];
  for (let n = startVal; n <= endVal; n++) {
    items.push(makeItem(prefix, String(n).padStart(padLen, '0'), type));
  }

  const { added, dupes } = addToPossession(items);
  let msg = `${added} precinto(s) ${type} agregado(s)`;
  if (dupes) msg += `, ${dupes} ya existían (omitidos)`;
  showToast(msg);
  if (added) {
    document.getElementById('range-start').value = '';
    document.getElementById('range-end').value   = '';
    render(); switchTab('possession');
  }
});

// ── Single form ───────────────────────────────────────────────────────────────
document.getElementById('single-form').addEventListener('submit', e => {
  e.preventDefault();
  const prefix = document.getElementById('single-prefix').value.trim();
  const numStr = document.getElementById('single-number').value.trim();
  const type   = getSelectedType('single-form');

  if (!prefix || !numStr)          { showToast('Completa todos los campos'); return; }
  if (isNaN(parseInt(numStr, 10))) { showToast('El número no es válido');    return; }

  const { added } = addToPossession([makeItem(prefix, numStr, type)]);
  if (added) {
    showToast(`Precinto ${type} agregado`);
    document.getElementById('single-number').value = '';
    render(); switchTab('possession');
  } else {
    showToast('El precinto ya existe');
  }
});

// ── Export / Import ───────────────────────────────────────────────────────────
document.getElementById('btn-export').addEventListener('click', () => {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), {
    href: url, download: `precintos-${new Date().toISOString().slice(0,10)}.json`
  });
  a.click(); URL.revokeObjectURL(url);
});

document.getElementById('btn-import').addEventListener('click', () => {
  document.getElementById('import-file').click();
});

document.getElementById('import-file').addEventListener('change', e => {
  const file = e.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const data = JSON.parse(ev.target.result);
      if (Array.isArray(data.possession) && Array.isArray(data.used)) {
        state = { possession: data.possession, used: data.used, rendidos: data.rendidos || [] };
        ['possession','used','rendidos'].forEach(l => state[l].forEach(i => { if (!i.type) i.type = 'CABLE'; }));
        selected    = { possession: new Set(), used: new Set() };
        typeFilters = { possession: 'ALL', used: 'ALL', rendidos: 'ALL' };
        filters     = { possession: '', used: '', rendidos: '' };
        document.getElementById('search-possession').value = '';
        document.getElementById('search-used').value       = '';
        document.getElementById('search-rendidos').value   = '';
        saveState(); render(); renderHistorial();
        showToast('Datos importados correctamente');
      } else { showToast('Archivo JSON inválido'); }
    } catch (_) { showToast('Error al leer el archivo'); }
  };
  reader.readAsText(file); e.target.value = '';
});

// ── Toast ─────────────────────────────────────────────────────────────────────
let toastTimer = null;
function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 3000);
}

// ── Service Worker ────────────────────────────────────────────────────────────
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  });
}

// ── Init ──────────────────────────────────────────────────────────────────────
loadState();
render();
