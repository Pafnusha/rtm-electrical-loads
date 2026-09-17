import { lookupKr } from './kr-table.js';
import { exportToExcel } from './excel-export.js';

const DEMO_ROWS = [
  { name: 'Токарные станки', n: 8, pnUnit: 7.5, ki: 0.14, cosPhi: 0.4, ks: 0.25 },
  { name: 'Сварочные трансформаторы', n: 4, pnUnit: 25, ki: 0.35, cosPhi: 0.5, ks: 0.35 },
  { name: 'Вентиляторы', n: 6, pnUnit: 5.5, ki: 0.7, cosPhi: 0.8, ks: 0.65 },
  { name: 'Кран мостовой', n: 2, pnUnit: 22, ki: 0.2, cosPhi: 0.5, ks: 0.3 },
  { name: 'Освещение цеха', n: 1, pnUnit: 18, ki: 0.85, cosPhi: 0.95, ks: 0.9 }
];

const state = {
  mode: 'rtm', // 'rtm' | 'demand'
  krTable: 'table1', // 'table1' | 'table2'
  un: 0.4,
  krOverride: false,
  krManual: 1.0,
  groupKs: 0.5,
  rows: DEMO_ROWS.map((r) => ({ ...r, id: uid() }))
};

function uid() {
  return 'r' + Math.random().toString(36).slice(2, 10);
}

function clampCos(c) {
  return Math.max(-0.999999, Math.min(0.999999, c));
}

function tgFromCos(cosPhi) {
  const c = Number(cosPhi);
  if (!Number.isFinite(c) || Math.abs(c) >= 1) return 0;
  return Math.tan(Math.acos(clampCos(c)));
}

function round(x, d = 4) {
  if (!Number.isFinite(x)) return 0;
  const m = 10 ** d;
  return Math.round(x * m) / m;
}

function compute() {
  const rows = state.rows.map((r) => {
    const n = Number(r.n) || 0;
    const pnUnit = Number(r.pnUnit) || 0;
    const ki = Number(r.ki) || 0;
    const cosPhi = Number(r.cosPhi) || 0;
    const ks = Number(r.ks) || 0;
    const Pn = n * pnUnit;
    const KiPn = ki * Pn;
    const tg = tgFromCos(cosPhi);
    const nPn2 = n * pnUnit * pnUnit;
    return { ...r, n, pnUnit, ki, cosPhi, ks, Pn, KiPn, tg, nPn2 };
  });

  const Pn = rows.reduce((s, r) => s + r.Pn, 0);
  const KiPn = rows.reduce((s, r) => s + r.KiPn, 0);
  const sumNPn2 = rows.reduce((s, r) => s + r.nPn2, 0);
  const KiAvg = Pn > 0 ? KiPn / Pn : 0;

  let ne = 0;
  if (sumNPn2 > 0 && Pn > 0) {
    ne = (Pn * Pn) / sumNPn2;
    // Если все одинаковые — ne ≈ sum(n); ограничим разумным минимумом
    if (ne < 1 && Pn > 0) ne = 1;
  }
  // Не больше суммы n (по смыслу РТМ при упрощённых проверках)
  const sumN = rows.reduce((s, r) => s + r.n, 0);
  if (sumN > 0 && ne > sumN * 1.0001) {
    // допускаем расчётное ne > n для разнотипных — по основной формуле это нормально
  }

  const tgWeightedNum = rows.reduce((s, r) => s + r.KiPn * r.tg, 0);
  const tgAvg = KiPn > 0 ? tgWeightedNum / KiPn : 0;

  let kr = 1;
  if (state.krOverride) {
    kr = Number(state.krManual) || 1;
  } else {
    kr = lookupKr(state.krTable, ne, KiAvg);
  }

  let Pp; let Qp;
  if (state.mode === 'demand') {
    const ks = Number(state.groupKs) || 0;
    Pp = ks * Pn;
    Qp = Pp * tgAvg;
  } else {
    Pp = kr * KiPn;
    Qp = ne <= 10 ? 1.1 * KiPn * tgAvg : KiPn * tgAvg;
  }

  const Sp = Math.sqrt(Pp * Pp + Qp * Qp);
  const Un = Number(state.un) || 0.4;
  const Ip = Un > 0 ? Sp / (Math.sqrt(3) * Un) : 0;

  return {
    rows,
    Pn, KiPn, KiAvg, sumNPn2, ne, tgAvg, kr, Pp, Qp, Sp, Ip, Un, sumN
  };
}

/* ——— DOM ——— */
const $ = (sel) => document.querySelector(sel);
const tbody = () => $('#ep-tbody');

function fmt(x, d = 3) {
  if (!Number.isFinite(x)) return '—';
  return Number(x).toLocaleString('ru-RU', {
    minimumFractionDigits: 0,
    maximumFractionDigits: d
  });
}

function renderRows() {
  const tb = tbody();
  tb.innerHTML = '';
  const demand = state.mode === 'demand';
  state.rows.forEach((r, idx) => {
    const tr = document.createElement('tr');
    tr.dataset.id = r.id;
    tr.innerHTML = `
      <td class="col-num">${idx + 1}</td>
      <td><input type="text" data-f="name" value="${esc(r.name)}" placeholder="Наименование ЭП"></td>
      <td><input type="number" data-f="n" min="0" step="1" value="${r.n}"></td>
      <td><input type="number" data-f="pnUnit" min="0" step="0.01" value="${r.pnUnit}"></td>
      <td><input type="number" data-f="ki" min="0" max="1" step="0.01" value="${r.ki}" ${demand ? 'disabled' : ''}></td>
      <td><input type="number" data-f="cosPhi" min="0" max="1" step="0.01" value="${r.cosPhi}"></td>
      <td><input type="number" data-f="ks" min="0" max="1" step="0.01" value="${r.ks}" ${demand ? '' : 'disabled'} title="Кс строки (справочно; в режиме спроса используется Кс группы)"></td>
      <td class="col-calc" data-c="Pn">—</td>
      <td class="col-calc" data-c="KiPn">—</td>
      <td class="col-calc" data-c="tg">—</td>
      <td class="col-actions">
        <button type="button" class="btn-icon" data-act="dup" title="Дублировать">⧉</button>
        <button type="button" class="btn-icon danger" data-act="del" title="Удалить">×</button>
      </td>`;
    tb.appendChild(tr);
  });
  updateCalcCells();
}

function esc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;');
}

function updateCalcCells() {
  const t = compute();
  const trs = tbody().querySelectorAll('tr');
  t.rows.forEach((r, i) => {
    const tr = trs[i];
    if (!tr) return;
    tr.querySelector('[data-c="Pn"]').textContent = fmt(r.Pn, 2);
    tr.querySelector('[data-c="KiPn"]').textContent = fmt(r.KiPn, 2);
    tr.querySelector('[data-c="tg"]').textContent = fmt(r.tg, 3);
  });

  $('#out-Pn').textContent = fmt(t.Pn, 2);
  $('#out-KiPn').textContent = fmt(t.KiPn, 2);
  $('#out-KiAvg').textContent = fmt(t.KiAvg, 3);
  $('#out-ne').textContent = fmt(t.ne, 2);
  $('#out-kr').textContent = fmt(t.kr, 3);
  $('#out-tg').textContent = fmt(t.tgAvg, 3);
  $('#out-Pp').textContent = fmt(t.Pp, 2);
  $('#out-Qp').textContent = fmt(t.Qp, 2);
  $('#out-Sp').textContent = fmt(t.Sp, 2);
  $('#out-Ip').textContent = fmt(t.Ip, 1);

  const krNote = $('#kr-note');
  if (state.mode === 'demand') {
    krNote.textContent = 'В режиме спроса Кр не используется (Рр = Кс · ΣPн).';
  } else if (state.krOverride) {
    krNote.textContent = 'Кр задан вручную.';
  } else {
    const tbl = state.krTable === 'table2' ? 'табл. 2' : 'табл. 1';
    krNote.textContent = `Кр по ${tbl} РТМ (интерполяция по nэ=${fmt(t.ne, 2)}, Ки=${fmt(t.KiAvg, 3)}).`;
  }
}

function syncControlsFromState() {
  $('#mode-select').value = state.mode;
  $('#kr-table-select').value = state.krTable;
  $('#un-input').value = state.un;
  $('#kr-override').checked = state.krOverride;
  $('#kr-manual').value = state.krManual;
  $('#kr-manual').disabled = !state.krOverride;
  $('#group-ks').value = state.groupKs;
  document.body.dataset.mode = state.mode;
  renderRows();
}

function bind() {
  $('#mode-select').addEventListener('change', (e) => {
    state.mode = e.target.value;
    renderRows();
  });
  $('#kr-table-select').addEventListener('change', (e) => {
    state.krTable = e.target.value;
    updateCalcCells();
  });
  $('#un-input').addEventListener('input', (e) => {
    state.un = Number(e.target.value) || 0.4;
    updateCalcCells();
  });
  $('#kr-override').addEventListener('change', (e) => {
    state.krOverride = e.target.checked;
    $('#kr-manual').disabled = !state.krOverride;
    updateCalcCells();
  });
  $('#kr-manual').addEventListener('input', (e) => {
    state.krManual = Number(e.target.value) || 1;
    updateCalcCells();
  });
  $('#group-ks').addEventListener('input', (e) => {
    state.groupKs = Number(e.target.value) || 0;
    updateCalcCells();
  });

  tbody().addEventListener('input', (e) => {
    const inp = e.target.closest('input[data-f]');
    if (!inp) return;
    const tr = inp.closest('tr');
    const row = state.rows.find((r) => r.id === tr.dataset.id);
    if (!row) return;
    const f = inp.dataset.f;
    if (f === 'name') row.name = inp.value;
    else row[f] = inp.value === '' ? 0 : Number(inp.value);
    updateCalcCells();
  });

  tbody().addEventListener('click', (e) => {
    const btn = e.target.closest('[data-act]');
    if (!btn) return;
    const tr = btn.closest('tr');
    const id = tr.dataset.id;
    const idx = state.rows.findIndex((r) => r.id === id);
    if (idx < 0) return;
    if (btn.dataset.act === 'del') {
      if (state.rows.length <= 1) {
        state.rows[0] = { id: uid(), name: '', n: 1, pnUnit: 0, ki: 0.2, cosPhi: 0.8, ks: 0.5 };
      } else {
        state.rows.splice(idx, 1);
      }
      renderRows();
    } else if (btn.dataset.act === 'dup') {
      const src = state.rows[idx];
      state.rows.splice(idx + 1, 0, { ...src, id: uid() });
      renderRows();
    }
  });

  $('#btn-add').addEventListener('click', () => {
    state.rows.push({
      id: uid(), name: '', n: 1, pnUnit: 0, ki: 0.2, cosPhi: 0.8, ks: 0.5
    });
    renderRows();
  });

  $('#btn-save-json').addEventListener('click', () => {
    const payload = {
      version: 1,
      mode: state.mode,
      krTable: state.krTable,
      un: state.un,
      krOverride: state.krOverride,
      krManual: state.krManual,
      groupKs: state.groupKs,
      rows: state.rows.map(({ name, n, pnUnit, ki, cosPhi, ks }) => ({
        name, n, pnUnit, ki, cosPhi, ks
      }))
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    downloadBlob(blob, 'rtm-loads.json');
  });

  $('#btn-load-json').addEventListener('click', () => $('#file-json').click());
  $('#file-json').addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    try {
      const data = JSON.parse(await file.text());
      if (Array.isArray(data.rows)) {
        state.rows = data.rows.map((r) => ({
          id: uid(),
          name: r.name || '',
          n: Number(r.n) || 0,
          pnUnit: Number(r.pnUnit) || 0,
          ki: Number(r.ki) || 0,
          cosPhi: Number(r.cosPhi) || 0.8,
          ks: Number(r.ks) || 0.5
        }));
      }
      if (data.mode) state.mode = data.mode;
      if (data.krTable) state.krTable = data.krTable;
      if (data.un != null) state.un = Number(data.un) || 0.4;
      if (data.krOverride != null) state.krOverride = !!data.krOverride;
      if (data.krManual != null) state.krManual = Number(data.krManual) || 1;
      if (data.groupKs != null) state.groupKs = Number(data.groupKs) || 0.5;
      syncControlsFromState();
    } catch (err) {
      alert('Не удалось загрузить JSON: ' + err.message);
    }
  });

  $('#btn-excel').addEventListener('click', async () => {
    const btn = $('#btn-excel');
    btn.disabled = true;
    btn.textContent = 'Формирование…';
    try {
      const t = compute();
      const blob = await exportToExcel({
        rows: state.rows,
        meta: {
          mode: state.mode,
          krTable: state.krTable,
          un: state.un,
          krOverride: state.krOverride,
          krManual: state.krManual,
          groupKs: state.groupKs
        },
        totals: t
      });
      downloadBlob(blob, 'rtm-electrical-loads.xlsx');
    } catch (err) {
      alert('Ошибка экспорта Excel: ' + err.message);
      console.error(err);
    } finally {
      btn.disabled = false;
      btn.textContent = 'Скачать Excel';
    }
  });

  $('#btn-reset').addEventListener('click', () => {
    if (!confirm('Загрузить демо-пример (5 строк)?')) return;
    state.rows = DEMO_ROWS.map((r) => ({ ...r, id: uid() }));
    state.mode = 'rtm';
    state.krTable = 'table1';
    state.un = 0.4;
    state.krOverride = false;
    state.krManual = 1;
    state.groupKs = 0.5;
    syncControlsFromState();
  });
}

function downloadBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

document.addEventListener('DOMContentLoaded', () => {
  bind();
  syncControlsFromState();
});
