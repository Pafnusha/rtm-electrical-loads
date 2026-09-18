/**
 * Экспорт расчёта в .xlsx: исходные, итоги, КУ.
 */

const EXCELJS_CDN = [
  'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
  'https://unpkg.com/exceljs@4.4.0/dist/exceljs.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js'
];

function excelLib() {
  return (typeof ExcelJS !== 'undefined' && ExcelJS) || (typeof window !== 'undefined' && window.ExcelJS) || null;
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = () => resolve(src);
    s.onerror = () => reject(new Error(src));
    document.head.appendChild(s);
  });
}

async function ensureExcelJS() {
  if (excelLib()) return excelLib();
  for (const src of EXCELJS_CDN) {
    try {
      await loadScript(src);
      if (excelLib()) return excelLib();
    } catch (e) { /* next */ }
  }
  throw new Error('ExcelJS не загружен. Откройте: https://cdn.jsdelivr.net/gh/Pafnusha/rtm-electrical-loads@main/index.html');
}

export async function exportToExcel(opts) {
  const { rows, meta, totals } = opts;
  const ExcelJS = await ensureExcelJS();

  const wb = new ExcelJS.Workbook();
  wb.creator = 'РТМ калькулятор НТД';
  wb.created = new Date();

  const wsSrc = wb.addWorksheet('Исходные', { views: [{ state: 'frozen', ySplit: 1 }] });
  const wsTot = wb.addWorksheet('Итоги и КУ');
  const wsHelp = wb.addWorksheet('Справка');

  wsSrc.addRow([
    'Наименование', 'n', 'Pн_ед, кВт', 'Ки', 'cos φ',
    'Pн, кВт', 'Ки·Pн, кВт', 'tg φ', 'Q=КиPн·tg', 'n·Pн²', 'Кс'
  ]);
  wsSrc.getRow(1).font = { bold: true };
  wsSrc.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF7' } };

  const n = Math.max(rows.length, 1);
  for (let i = 0; i < n; i++) {
    const r = rows[i] || { name: '', n: 0, pnUnit: 0, ki: 0, cosPhi: 0.8, ks: 0.5 };
    const rowIdx = i + 2;
    const excelRow = wsSrc.getRow(rowIdx);
    excelRow.getCell(1).value = r.name || '';
    excelRow.getCell(2).value = Number(r.n) || 0;
    excelRow.getCell(3).value = Number(r.pnUnit) || 0;
    excelRow.getCell(4).value = Number(r.ki) || 0;
    excelRow.getCell(5).value = Number(r.cosPhi) || 0;
    excelRow.getCell(11).value = Number(r.ks) || 0;

    const pn = (Number(r.n) || 0) * (Number(r.pnUnit) || 0);
    const kiPn = (Number(r.ki) || 0) * pn;
    let tg = 0;
    const c = Number(r.cosPhi) || 0;
    if (c > -1 && c < 1) tg = Math.tan(Math.acos(Math.max(-1, Math.min(1, c))));
    const q = kiPn * tg;
    const nPn2 = (Number(r.n) || 0) * Math.pow(Number(r.pnUnit) || 0, 2);
    excelRow.getCell(6).value = { formula: `B${rowIdx}*C${rowIdx}`, result: pn };
    excelRow.getCell(7).value = { formula: `D${rowIdx}*F${rowIdx}`, result: kiPn };
    excelRow.getCell(8).value = {
      formula: `IF(OR(E${rowIdx}>=1,E${rowIdx}<=-1),0,TAN(ACOS(MAX(-1,MIN(1,E${rowIdx})))))`,
      result: tg
    };
    excelRow.getCell(9).value = { formula: `G${rowIdx}*H${rowIdx}`, result: q };
    excelRow.getCell(10).value = { formula: `B${rowIdx}*C${rowIdx}^2`, result: nPn2 };
  }

  const lastData = n + 1;
  const totRow = lastData + 1;
  wsSrc.getCell(totRow, 1).value = 'ИТОГО';
  wsSrc.getCell(totRow, 1).font = { bold: true };
  wsSrc.getCell(totRow, 6).value = { formula: `SUM(F2:F${lastData})`, result: totals.Pn };
  wsSrc.getCell(totRow, 7).value = { formula: `SUM(G2:G${lastData})`, result: totals.KiPn };
  wsSrc.getCell(totRow, 9).value = { formula: `SUM(I2:I${lastData})`, result: totals.Qsum };
  wsSrc.getCell(totRow, 10).value = { formula: `SUM(J2:J${lastData})`, result: totals.sumNPn2 };
  wsSrc.getRow(totRow).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2B3' } };

  wsSrc.columns = [
    { width: 42 }, { width: 8 }, { width: 12 }, { width: 8 }, { width: 10 },
    { width: 12 }, { width: 12 }, { width: 10 }, { width: 14 }, { width: 14 }, { width: 8 }
  ];

  wsTot.getColumn(1).width = 44;
  wsTot.getColumn(2).width = 18;
  wsTot.getColumn(3).width = 52;

  const put = (row, label, formula, result, note) => {
    wsTot.getCell(row, 1).value = label;
    if (formula != null) wsTot.getCell(row, 2).value = { formula, result: result ?? null };
    else wsTot.getCell(row, 2).value = result;
    if (note) wsTot.getCell(row, 3).value = note;
  };

  wsTot.getCell(1, 1).value = 'Параметр';
  wsTot.getCell(1, 2).value = 'Значение';
  wsTot.getCell(1, 3).value = 'Примечание';
  wsTot.getRow(1).font = { bold: true };
  wsTot.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF7' } };

  const mode = meta.mode || 'rtm';
  const un = Number(meta.un) || 0.4;
  const krManual = !!meta.krOverride;
  const krVal = Number(meta.krManual != null ? meta.krManual : totals.kr) || 1;
  const groupKs = Number(meta.groupKs) || 0.5;
  const ko = Number(meta.ko != null ? meta.ko : totals.ko) || 1;
  const cos2 = Number(meta.cosTarget != null ? meta.cosTarget : totals.cos2) || 0.95;

  put(2, 'Режим', null, mode === 'demand' ? 'Коэффициент спроса' : 'РТМ');
  put(3, 'Таблица Кр', null, meta.krTable === 'table2' ? 'Табл. 2' : 'Табл. 1');
  put(4, 'Uн, кВ', null, un);
  put(5, 'ΣPн, кВт', `SUM(Исходные!F2:F${lastData})`, totals.Pn);
  put(6, 'ΣКи·Pн, кВт', `SUM(Исходные!G2:G${lastData})`, totals.KiPn);
  put(7, 'Ки ср', 'IF(B5>0,B6/B5,0)', totals.KiAvg);
  put(8, 'Σ(n·Pн²)', `SUM(Исходные!J2:J${lastData})`, totals.sumNPn2);
  put(9, 'nэ', 'IF(B8>0,B5^2/B8,0)', totals.ne, 'nэ = (ΣPн)² / Σ(n·Pн_ед²)');
  put(10, 'tgφ ср', `IF(B6>0,SUMPRODUCT(Исходные!G2:G${lastData},Исходные!H2:H${lastData})/B6,0)`, totals.tgAvg);
  put(11, 'Кр', null, krVal, krManual ? 'Ручной Кр' : 'По табл. РТМ');
  put(12, 'Кс группы', null, groupKs);

  if (mode === 'demand') {
    put(13, 'Рр ИТОГО, кВт', 'B12*B5', totals.Pp, 'Рр = Кс · ΣPн');
    put(14, 'Qр ИТОГО, квар', 'B13*B10', totals.Qp);
  } else {
    put(13, 'Рр ИТОГО, кВт', 'B11*B6', totals.Pp, 'Рр = Кр · ΣКиPн');
    put(14, 'Qр ИТОГО, квар', 'IF(B9<=10,1.1*B6*B10,B6*B10)', totals.Qp, '1,1 при nэ≤10');
  }
  put(15, 'Sр ИТОГО, кВ·А', 'SQRT(B13^2+B14^2)', totals.Sp);
  put(16, 'Iр, А', 'IF(B4>0,B15/(SQRT(3)*B4),0)', totals.Ip);
  put(17, 'cosφ₁', 'IF(B15>0,B13/B15,0)', totals.cos1);

  put(19, 'Ко одновременности', null, ko, 'Задаёт проектировщик');
  put(20, 'Рр·Ко, кВт', 'B13*B19', totals.PpKo);
  put(21, 'Qр·Ко, квар', 'B14*B19', totals.QpKo);
  put(22, 'Sр·Ко, кВ·А', 'SQRT(B20^2+B21^2)', totals.SpKo);

  put(24, 'cosφ₂ проектировщика', null, cos2, 'Целевой cosφ после КУ');
  put(25, 'tgφ₂', 'IF(OR(B24>=1,B24<=-1),0,TAN(ACOS(MAX(-1,MIN(1,B24)))))', totals.tg2);
  put(26, 'tgφ₁ по итогу', 'IF(B20>0,B21/B20,0)', totals.tg1);
  put(27, 'Qку, квар', 'MAX(0,B20*(B26-B25))', totals.Qc, 'Qку = Рр·Ко · (tg1−tg2)');
  put(28, 'Q после КУ, квар', 'B21-B27', totals.Qp2);
  put(29, 'S после КУ, кВ·А', 'SQRT(B20^2+B28^2)', totals.Sp2);
  put(30, 'I после КУ, А', 'IF(B4>0,B29/(SQRT(3)*B4),0)', totals.Ip2);
  put(31, 'Sтр реком., кВ·А', null, totals.Str, 'Ближайший стандартный ТМ ≥ S после КУ');

  [13, 14, 15, 27, 29, 31].forEach((i) => {
    wsTot.getRow(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2B3' } };
  });

  const helpLines = [
    ['Ведомость нагрузок по РТМ 36.18.32.4-92 + КУ'],
    [''],
    ['ИТОГО по строкам входит в Рр, Qр, Sр.'],
    ['Рр = Кр · Σ(Ки·Pн); Qр = 1,1·ΣКиPн·tg при nэ≤10.'],
    ['Итого с Ко: Рр' = Рр·Ко; Qр' = Qр·Ко.'],
    ['Qку = Рр' · (tgφ₁ − tgφ₂), где tgφ₂ из cosφ проектировщика.'],
    ['После КУ: Q = Qр' − Qку; S = √(Рр'² + Q²).']
  ];
  helpLines.forEach((line, idx) => { wsHelp.getCell(idx + 1, 1).value = line[0]; });
  wsHelp.getColumn(1).width = 96;

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}
