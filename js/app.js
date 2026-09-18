import { lookupKr } from './kr-table.js';
import { exportToExcel } from './excel-export.js';

const TR_SERIES = [25, 40, 63, 100, 160, 250, 400, 630, 1000, 1600, 2500, 4000, 6300, 10000];

const DEMO_ROWS = [
  { name: 'Сети внутриплощадочные. Электрообогрев трубопроводных коллекторов', n: 1, pnUnit: 340, ki: 0.75, cosPhi: 0.98, ks: 0.75 },
  { name: 'Электрообогрев', n: 1, pnUnit: 200, ki: 0.75, cosPhi: 0.98, ks: 0.75 },
  { name: 'Собственные нужды азотной станции', n: 1, pnUnit: 40, ki: 0.8, cosPhi: 0.85, ks: 0.8 },
  { name: 'Азотная станция. Насосы (2 раб. + 1 рез.)', n: 2, pnUnit: 216, ki: 0.8, cosPhi: 0.8, ks: 0.8 },
  { name: 'УПТ. Насосы 2046-P-301A,B,C (2 раб. + 1 рез.)', n: 2, pnUnit: 132, ki: 1, cosPhi: 0.8, ks: 1 },
  { name: 'Технологические нагрузки УПТ', n: 1, pnUnit: 180, ki: 0.7, cosPhi: 0.75, ks: 0.7 }
];

const state = {
  mode: 'rtm',
  krTable: 'table1',
  un: 0.4,
  krOverride: false,
  krManual: 1.0,
  groupKs: 0.5,
  ko: 0.9,
  cosTarget: 0.95,
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

function pickTransformer(s) {
  if (!(s > 0)) return 0;
  const found = TR_SERIES.find((x) => x >= s - 1e-9);
  return found || Math.ceil(s);
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
    const Qrow = KiPn * tg;
    const nPn2 = n * pnUnit * pnUnit;
    return { ...r, n, pnUnit, ki, cosPhi, ks, Pn, KiPn, tg, Qrow, nPn2 };
  });

  const Pn = rows.reduce((s, r) => s + r.Pn, 0);
  const KiPn = rows.reduce((s, r) => s + r.KiPn, 0);
  const Qsum = rows.reduce((s, r) => s + r.Qrow, 0);
  const sumNPn2 = rows.reduce((s, r) => s + r.nPn2, 0);
  const KiAvg = Pn > 0 ? KiPn / Pn : 0;
  const sumN = rows.reduce((s, r) => s + r.n, 0);

  let ne = 0;
  if (sumNPn2 > 0 && Pn > 0) {
    ne = (Pn * Pn) / sumNPn2;
    if (ne < 1 && Pn > 0) ne = 1;
  }

  const tgWeightedNum = rows.reduce((s, r) => s + r.KiPn * r.tg, 0);
  const tgAvg = KiPn > 0 ? tgWeightedNum / KiPn : 0;

  let kr = 1;
  if (state.krOverride) kr = Number(state.krManual) || 1;
  else kr = lookupKr(state.krTable, ne, KiAvg);

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
  const cos1 = Sp > 0 ? Pp / Sp : 0;

  const ko = Number(state.ko);
  const koSafe = Number.isFinite(ko) ? ko : 1;
  const PpKo = Pp * koSafe;
  const QpKo = Qp * koSafe;
  const SpKo = Math.sqrt(PpKo * PpKo + QpKo * QpKo);
  const IpKo = Un > 0 ? SpKo / (Math.sqrt(3) * Un) : 0;

  const cos2 = Number(state.cosTarget) || 0.95;
  const tg2 = tgFromCos(cos2);
  const tg1 = PpKo > 0 ? QpKo / PpKo : tgAvg;
  let Qc = PpKo * (tg1 - tg2);
  if (Qc < 0) Qc = 0;
  const Qp2 = QpKo - Qc;
  const Sp2 = Math.sqrt(PpKo * PpKo + Qp2 * Qp2);
  const Ip2 = Un > 0 ? Sp2 / (Math.sqrt(3) * Un) : 0;
  const cosAfter = Sp2 > 0 ? PpKo / Sp2 : 0;
  const Str = pickTransformer(Sp2);

  return {
    rows, Pn, KiPn, Qsum, KiAvg, sumNPn2, ne, tgAvg, kr, Pp, Qp, Sp, Ip, Un, sumN,
    cos1, ko: koSafe, PpKo, QpKo, SpKo, IpKo, cos2, tg1, tg2, Qc, Qp2, Sp2, Ip2, cosAfter, Str
  };
}

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
      <td class="col-calc" data-c="Pn">—</td>
      <td><input type="number" data-f="ki" min="0" max="1" step="0.01" value="${r.ki}" ${demand ? 'disabled' : ''}></td>
      <td><input type="number" data-f="cosPhi" min="0" max="1" step="0.01" value="${r.cosPhi}"></td>
      <td class="col-calc" data-c="tg">—</td>
      <td class="col-calc" data-c="KiPn">—</td>
      <td class="col-calc" data-c="Qrow">—</td>
      <td class="col-calc" data-c="nPn2">—</td>
      <td><input type="number" data-f="ks" min="0" max="1" step="0.01" value="${r.ks}" ${demand ? '' : 'disabled'}></td>
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
    .replace(/&/g, '&')
    .replace(/"/g, '"')
    .replace(/</g, '<');
}

function setTxt(id, val, d) {
  const el = $(id);
  if (el) el.textContent = fmt(val, d);
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
    tr.querySelector('[data-c="Qrow"]').textContent = fmt(r.Qrow, 2);
    tr.querySelector('[data-c="nPn2"]').textContent = fmt(r.nPn2, 0);
  });

  setTxt('#out-Pn', t.Pn, 2);
  setTxt('#out-KiPn', t.KiPn, 2);
  setTxt('#out-KiAvg', t.KiAvg, 3);
  setTxt('#out-cos1', t.cos1, 3);
  setTxt('#out-ne', t.ne, 2);
  setTxt('#out-kr', t.kr, 3);
  setTxt('#out-tg', t.tgAvg, 3);
  setTxt('#out-Pp', t.Pp, 2);
  setTxt('#out-Qp', t.Qp, 2);
  setTxt('#out-Sp', t.Sp, 2);
  setTxt('#out-Ip', t.Ip, 1);
  setTxt('#out-PpKo', t.PpKo, 2);
  setTxt('#out-QpKo', t.QpKo, 2);
  setTxt('#out-SpKo', t.SpKo, 2);
  setTxt('#out-tg2', t.tg2, 3);
  setTxt('#out-Qc', t.Qc, 2);
  setTxt('#out-Qp2', t.Qp2, 2);
  setTxt('#out-Sp2', t.Sp2, 2);
  setTxt('#out-Ip2', t.Ip2, 1);
  setTxt('#out-Str', t.Str, 0);

  const krNote = $('#kr-note');
  if (state.mode === 'demand') {
    krNote.textContent = 'Режим спроса: Рр = Кс · ΣPн. Итого идёт в расчёт КУ.';
  } else if (state.krOverride) {
    krNote.textContent = 'Кр задан вручную. Итого Рр/Рр·Ко используется для КУ.';
  } else {
    const tbl = state.krTable === 'table2' ? 'табл. 2' : 'табл. 1';
    krNote.textContent = `Итого: nэ=${fmt(t.ne, 2)}, Ки=${fmt(t.KiAvg, 3)}, Кр по ${tbl} = ${fmt(t.kr, 3)}.`;
  }

  const ku = $('#ku-note');
  if (ku) {
    ku.textContent =
      `cosφ₁=${fmt(t.cos1, 3)} (tgφ₁=${fmt(t.tg1, 3)}) → cosφ₂=${fmt(t.cos2, 3)} (tgφ₂=${fmt(t.tg2, 3)}). ` +
      `Qку=${fmt(t.Qc, 1)} квар. После КУ: S=${fmt(t.Sp2, 1)} кВ·А, рекомендуется ТМ ${fmt(t.Str, 0)} кВ·А.`;
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
  if ($('#ko-input')) $('#ko-input').value = state.ko;
  if ($('#cos-target')) $('#cos-target').value = state.cosTarget;
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
  $('#ko-input').addEventListener('input', (e) => {
    state.ko = Number(e.target.value);
    if (!Number.isFinite(state.ko)) state.ko = 1;
    updateCalcCells();
  });
  $('#cos-target').addEventListener('input', (e) => {
    state.cosTarget = Number(e.target.value) || 0.95;
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
    state.rows.push({ id: uid(), name: '', n: 1, pnUnit: 0, ki: 0.2, cosPhi: 0.8, ks: 0.5 });
    renderRows();
  });

  $('#btn-save-json').addEventListener('click', () => {
    const payload = {
      version: 2,
      mode: state.mode,
      krTable: state.krTable,
      un: state.un,
      krOverride: state.krOverride,
      krManual: state.krManual,
      groupKs: state.groupKs,
      ko: state.ko,
      cosTarget: state.cosTarget,
      rows: state.rows.map(({ name, n, pnUnit, ki, cosPhi, ks }) => ({ name, n, pnUnit, ki, cosPhi, ks }))
    };
    downloadBlob(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' }), 'rtm-loads.json');
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
      if (data.ko != null) state.ko = Number(data.ko);
      if (data.cosTarget != null) state.cosTarget = Number(data.cosTarget) || 0.95;
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
          groupKs: state.groupKs,
          ko: state.ko,
          cosTarget: state.cosTarget
        },
        totals: t
      });
      downloadBlob(blob, 'nagruzki-ktp.xlsx');
    } catch (err) {
      alert('Ошибка экспорта Excel: ' + err.message);
      console.error(err);
    } finally {
      btn.disabled = false;
      btn.textContent = 'Скачать Excel';
    }
  });

  $('#btn-reset').addEventListener('click', () => {
    if (!confirm('Загрузить демо промысла (УПТ / азот / обогрев)?')) return;
    state.rows = DEMO_ROWS.map((r) => ({ ...r, id: uid() }));
    state.mode = 'rtm';
    state.krTable = 'table1';
    state.un = 0.4;
    state.krOverride = false;
    state.krManual = 1;
    state.groupKs = 0.5;
    state.ko = 0.9;
    state.cosTarget = 0.95;
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
