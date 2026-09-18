/**
 * Экспорт расчёта в .xlsx с живыми формулами Excel (ExcelJS).
 * Листы: «Исходные», «Итоги», «Справка».
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
    } catch (e) {
      /* next CDN */
    }
  }
  throw new Error('ExcelJS не загружен. Откройте калькулятор через jsDelivr, не raw.githack: https://cdn.jsdelivr.net/gh/Pafnusha/rtm-electrical-loads@main/index.html');
}

/**
 * @param {object} opts
 * @param {Array} opts.rows
 * @param {object} opts.meta
 * @param {object} opts.totals
 */
export async function exportToExcel(opts) {
  const { rows, meta, totals } = opts;
  const ExcelJS = await ensureExcelJS();

  const wb = new ExcelJS.Workbook();
  wb.creator = 'РТМ 36.18.32.4-92 калькулятор';
  wb.created = new Date();

  const wsSrc = wb.addWorksheet('Исходные', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });
  const wsTot = wb.addWorksheet('Итоги');
  const wsHelp = wb.addWorksheet('Справка');

  const headers = [
    'Наименование', 'n', 'Pн_ед, кВт', 'Ки', 'cos φ',
    'Pн, кВт', 'Ки·Pн, кВт', 'tg φ', 'n·Pн²', 'Кс'
  ];
  wsSrc.addRow(headers);
  wsSrc.getRow(1).font = { bold: true };
  wsSrc.getRow(1).fill = {
    type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF7' }
  };

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
    excelRow.getCell(10).value = Number(r.ks) || 0;

    const pn = (Number(r.n) || 0) * (Number(r.pnUnit) || 0);
    const kiPn = (Number(r.ki) || 0) * pn;
    let tg = 0;
    const c = Number(r.cosPhi) || 0;
    if (c > -1 && c < 1) tg = Math.tan(Math.acos(Math.max(-1, Math.min(1, c))));
    const nPn2 = (Number(r.n) || 0) * Math.pow(Number(r.pnUnit) || 0, 2);
    excelRow.getCell(6).value = { formula: `B${rowIdx}*C${rowIdx}`, result: pn };
    excelRow.getCell(7).value = { formula: `D${rowIdx}*F${rowIdx}`, result: kiPn };
    excelRow.getCell(8).value = {
      formula: `IF(OR(E${rowIdx}>=1,E${rowIdx}<=-1),0,TAN(ACOS(MAX(-1,MIN(1,E${rowIdx})))))`,
      result: tg
    };
    excelRow.getCell(9).value = { formula: `B${rowIdx}*C${rowIdx}^2`, result: nPn2 };
  }

  const lastData = n + 1;
  wsSrc.columns = [
    { width: 28 }, { width: 8 }, { width: 12 }, { width: 8 }, { width: 10 },
    { width: 12 }, { width: 12 }, { width: 10 }, { width: 14 }, { width: 8 }
  ];

  wsTot.getColumn(1).width = 36;
  wsTot.getColumn(2).width = 18;
  wsTot.getColumn(3).width = 48;

  const put = (row, label, formula, result, note) => {
    wsTot.getCell(row, 1).value = label;
    if (formula != null) {
      wsTot.getCell(row, 2).value = { formula, result: result ?? null };
    } else {
      wsTot.getCell(row, 2).value = result;
    }
    if (note) wsTot.getCell(row, 3).value = note;
  };

  wsTot.getCell(1, 1).value = 'Параметр';
  wsTot.getCell(1, 2).value = 'Значение';
  wsTot.getCell(1, 3).value = 'Примечание';
  wsTot.getRow(1).font = { bold: true };
  wsTot.getRow(1).fill = {
    type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF7' }
  };

  const mode = meta.mode || 'rtm';
  const un = Number(meta.un) || 0.4;
  const krManual = !!meta.krOverride;
  const krVal = Number(meta.krManual != null ? meta.krManual : totals.kr) || 1;
  const groupKs = Number(meta.groupKs) || 0.5;

  put(2, 'Режим расчёта', null, mode === 'demand' ? 'Коэффициент спроса' : 'РТМ (Ки, Кр, nэ)');
  put(3, 'Таблица Кр', null,
    meta.krTable === 'table2' ? 'Табл. 2 (шины НН / шинопроводы)' : 'Табл. 1 (питающие сети)',
    'РТМ 36.18.32.4-92');
  put(4, 'Uн, кВ', null, un, 'Номинальное напряжение');

  put(5, 'ΣPн, кВт', `SUM(Исходные!F2:F${lastData})`, totals.Pn);
  put(6, 'ΣКи·Pн, кВт', `SUM(Исходные!G2:G${lastData})`, totals.KiPn);
  put(7, 'Ки ср', `IF(B5>0,B6/B5,0)`, totals.KiAvg);
  put(8, 'Σ(n·Pн²)', `SUM(Исходные!I2:I${lastData})`, totals.sumNPn2);
  put(9, 'nэ', `IF(B8>0,B5^2/B8,0)`, totals.ne, 'nэ = (ΣPн)² / Σ(n·Pн_ед²)');
  put(
    10,
    'tgφ ср (взвеш.)',
    `IF(B6>0,SUMPRODUCT(Исходные!G2:G${lastData},Исходные!H2:H${lastData})/B6,0)`,
    totals.tgAvg
  );

  if (krManual || mode === 'demand') {
    put(11, 'Кр', null, krVal, krManual ? 'Ручной ввод Кр' : 'В режиме спроса не используется');
  } else {
    put(11, 'Кр', null, krVal,
      'По табл. РТМ (билинейная интерполяция в калькуляторе). При правке данных в Excel обновите Кр вручную или пересчитайте на сайте.');
  }

  put(12, 'Кс (группы)', null, groupKs, 'Для режима «Коэффициент спроса»');

  if (mode === 'demand') {
    put(13, 'Рр, кВт', `B12*B5`, totals.Pp, 'Рр = Кс · ΣPн');
    put(14, 'Qр, квар', `B13*B10`, totals.Qp, 'Qр = Рр · tgφ');
  } else {
    put(13, 'Рр, кВт', `B11*B6`, totals.Pp, 'Рр = Кр · ΣКи·Pн');
    put(
      14,
      'Qр, квар',
      `IF(B9<=10,1.1*B6*B10,B6*B10)`,
      totals.Qp,
      'Qр = 1,1·КиPн·tgφ при nэ≤10; иначе КиPн·tgφ'
    );
  }

  put(15, 'Sр, кВ·А', `SQRT(B13^2+B14^2)`, totals.Sp);
  put(16, 'Iр, А', `IF(B4>0,B15/(SQRT(3)*B4),0)`, totals.Ip, 'Iр = Sр / (√3 · Uн)');

  const helpLines = [
    ['Калькулятор электрических нагрузок по РТМ 36.18.32.4-92'],
    [''],
    ['Основные формулы (метод коэффициента использования):'],
    ['Pн_i = n · Pн_ед'],
    ['Ки·Pн_i = Ки · Pн_i'],
    ['tgφ_i = tan(arccos(cosφ))'],
    ['ΣPн = Σ Pн_i;  ΣКиPн = Σ Ки·Pн_i;  Ки_ср = ΣКиPн / ΣPн'],
    ['nэ = (ΣPн)² / Σ(n · Pн_ед²)'],
    ['Кр — по табл. 1 (питающие сети) или табл. 2 (шины НН / шинопроводы)'],
    ['Рр = Кр · ΣКиPн'],
    ['Qр = 1,1 · ΣКиPн · tgφ_ср при nэ ≤ 10;  Qр = ΣКиPн · tgφ_ср при nэ > 10'],
    ['Sр = √(Рр² + Qр²);  Iр = Sр / (√3 · Uн)'],
    [''],
    ['Режим коэффициента спроса: Рр = Кс · ΣPн; Qр = Рр · tgφ'],
    [''],
    ['Лист «Исходные»: колонки A–E и J — исходные данные; F–I — формулы.'],
    ['Лист «Итоги»: суммы и расчётные величины — формулы со ссылками на «Исходные».'],
    ['Кр берётся из таблицы на сайте (интерполяция). При изменении nэ/Ки в Excel'],
    ['обновите ячейку Кр вручную или пересчитайте на сайте и скачайте файл снова.'],
    [''],
    ['Полный текст нормативного документа не воспроизводится. См. официальный РТМ 36.18.32.4-92.']
  ];
  helpLines.forEach((line, idx) => {
    wsHelp.getCell(idx + 1, 1).value = line[0];
  });
  wsHelp.getColumn(1).width = 90;

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}
