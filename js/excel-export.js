/**
 * Экспорт расчёта в .xlsx с живыми формулами Excel (ExcelJS).
 * Листы: «Исходные», «Итоги», «Справка».
 */

/**
 * @param {object} opts
 * @param {Array} opts.rows — строки ЭП {name,n,pnUnit,ki,cosPhi,ks}
 * @param {object} opts.meta — {mode, krTable, un, krOverride, krManual, groupKs}
 * @param {object} opts.totals — вычисленные итоги (для кэшированных значений)
 */
export async function exportToExcel(opts) {
  const { rows, meta, totals } = opts;
  if (typeof ExcelJS === 'undefined') {
    throw new Error('ExcelJS не загружен');
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'РТМ 36.18.32.4-92 калькулятор';
  wb.created = new Date();

  const wsSrc = wb.addWorksheet('Исходные', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });
  const wsTot = wb.addWorksheet('Итоги');
  const wsHelp = wb.addWorksheet('Справка');

  // ——— Исходные ———
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

    // F: Pн = n * Pн_ед
    excelRow.getCell(6).value = { formula: `B${rowIdx}*C${rowIdx}` };
    // G: Ки·Pн
    excelRow.getCell(7).value = { formula: `D${rowIdx}*F${rowIdx}` };
    // H: tgφ = TAN(ACOS(cosφ)) — защита от |cos|>1
    excelRow.getCell(8).value = {
      formula: `IF(OR(E${rowIdx}>=1,E${rowIdx}<=-1),0,TAN(ACOS(MAX(-1,MIN(1,E${rowIdx})))))`
    };
    // I: n·Pн_ед²
    excelRow.getCell(9).value = { formula: `B${rowIdx}*C${rowIdx}^2` };
    // J: Кс (для режима спроса)
    excelRow.getCell(10).value = Number(r.ks) || 0;

    // Кэшированные значения (result) для просмотрщиков без пересчёта
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

  // ——— Итоги ———
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

  // ΣPн
  put(5, 'ΣPн, кВт', `SUM(Исходные!F2:F${lastData})`, totals.Pn);
  // ΣКиPн
  put(6, 'ΣКи·Pн, кВт', `SUM(Исходные!G2:G${lastData})`, totals.KiPn);
  // Киср
  put(7, 'Ки ср', `IF(B5>0,B6/B5,0)`, totals.KiAvg);
  // Σ(n·Pн²)
  put(8, 'Σ(n·Pн²)', `SUM(Исходные!I2:I${lastData})`, totals.sumNPn2);
  // nэ
  put(9, 'nэ', `IF(B8>0,B5^2/B8,0)`, totals.ne, 'nэ = (ΣPн)² / Σ(n·Pн_ед²)');

  // tgφ ср взвешенный: SUMPRODUCT(G,H)/ΣG
  put(
    10,
    'tgφ ср (взвеш.)',
    `IF(B6>0,SUMPRODUCT(Исходные!G2:G${lastData},Исходные!H2:H${lastData})/B6,0)`,
    totals.tgAvg
  );

  // Кр
  if (krManual || mode === 'demand') {
    put(11, 'Кр', null, krVal, krManual ? 'Ручной ввод Кр' : 'В режиме спроса не используется');
  } else {
    // Значение из таблицы (как число) + примечание об интерполяции
    put(11, 'Кр', null, krVal,
      'По табл. РТМ (билинейная интерполяция в калькуляторе). При правке данных в Excel обновите Кр вручную или пересчитайте на сайте.');
  }

  // Кс группы
  put(12, 'Кс (группы)', null, groupKs, 'Для режима «Коэффициент спроса»');

  if (mode === 'demand') {
    // Pp = Ks * Pn
    put(13, 'Рр, кВт', `B12*B5`, totals.Pp, 'Рр = Кс · ΣPн');
    // Qp = Pp * tgφ
    put(14, 'Qр, квар', `B13*B10`, totals.Qp, 'Qр = Рр · tgφ');
  } else {
    // Pp = Kr * KiPn
    put(13, 'Рр, кВт', `B11*B6`, totals.Pp, 'Рр = Кр · ΣКи·Pн');
    // Qp = 1.1*KiPn*tg if ne<=10 else KiPn*tg
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

  // ——— Справка ———
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
