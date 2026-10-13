import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadPane } from './harness.mjs';

const REGIONS = ['Nord', 'Sud', 'Est', 'Ouest'];

// A sales table: header row then `count` rows whose values derive from the row
// number, so every expected figure can be recomputed independently.
function sales(count) {
  const rows = [['Id', 'Région', 'Année', 'Montant', 'Statut']];
  for (let i = 1; i <= count; i++) {
    rows.push([i, REGIONS[i % 4], 2023 + (i % 3), (i % 100) + 0.5, i % 10 === 0 ? 'Annulé' : 'Payé']);
  }
  return rows;
}

const excel = (book, extra = {}) => loadPane({ host: 'Excel', excel: book, url: 'https://files.test/Ventes.xlsx', ...extra });
const book = (rows, more = []) => ({ active: 'Ventes', sheets: [{ name: 'Ventes', rows }, ...more] });

async function readAll(pane, first) {
  const pages = [];
  for (let target = first, guard = 0; guard < 2000; guard++) {
    const page = await pane.exec('office_read', target ? { target } : {});
    pages.push(page);
    if (!page.next) return pages;
    target = page.next;
  }
  throw new Error('office_read never reached the end');
}

test('the outline lists every sheet with its size and its header row', async () => {
  const pane = excel({
    active: 'Ventes',
    selection: 'B2:C3',
    sheets: [
      { name: 'Ventes', rows: sales(1000) },
      { name: 'Annexe 1', rows: [['Clé', 'Valeur'], ['a', 1]] },
      { name: 'Vide', rows: [] },
      { name: 'Masquée', rows: [['x']], hidden: true },
    ],
  });
  const outline = await pane.exec('office_outline', {});
  assert.equal(outline.host, 'Excel');
  assert.equal(outline.active, 'Ventes');
  assert.equal(outline.selection, 'Ventes!B2:C3');
  assert.equal(outline.sheetCount, 4);
  assert.deepEqual(outline.sheets[0], {
    name: 'Ventes',
    visibility: 'Visible',
    range: 'Ventes!A1:E1001',
    rows: 1001,
    columns: 5,
    firstRow: ['Id', 'Région', 'Année', 'Montant', 'Statut'],
  });
  assert.deepEqual(outline.sheets[2], { name: 'Vide', visibility: 'Visible', empty: true });
  assert.equal(outline.sheets[3].visibility, 'Hidden');
});

test('the outline of a 200,000-row sheet reads its header row and nothing else', async () => {
  const pane = excel(book(sales(200000)));
  const outline = await pane.exec('office_outline', {});
  assert.equal(outline.sheets[0].rows, 200001);
  assert.equal(pane.stats.cellsSent, 5);
});

test('reading page after page returns every row once, in order', async () => {
  const rows = sales(3000);
  const pane = excel(book(rows));
  const pages = await readAll(pane);
  assert.ok(pages.length > 1);
  assert.deepEqual(pages.flatMap((page) => page.values), rows);
  assert.equal(pages[0].range.split(':')[0], 'A1');
  assert.equal(pages[0].rowsLeft, rows.length - pages[0].values.length);
});

test('a read stays within the size limit and never pulls more than a bounded block of cells', async () => {
  const pane = excel(book(sales(200000)));
  const page = await pane.exec('office_read', { target: 'Ventes' });
  assert.ok(JSON.stringify(page.values).length <= pane.get('MAX_OUTPUT_CHARS'));
  assert.ok(pane.stats.largestRead <= pane.get('READ_CELLS'), `pulled ${pane.stats.largestRead} cells at once`);
  assert.match(page.next, /^Ventes!A\d+:E200001$/);
});

test('whole columns mean the used rows, not a million empty ones', async () => {
  const rows = sales(50);
  const pane = excel(book(rows));
  const page = await pane.exec('office_read', { target: 'Ventes!B:C' });
  assert.equal(page.range, 'B1:C51');
  assert.deepEqual(page.values, rows.map((row) => row.slice(1, 3)));
  assert.equal(page.next, undefined);
});

test('an explicit range is returned as asked, on the named sheet', async () => {
  const pane = excel(book(sales(50), [{ name: 'Annexe 1', rows: [['a', 'b'], ['c', 'd'], ['e', 'f']] }]));
  const page = await pane.exec('office_read', { target: "'Annexe 1'!A2:B3" });
  assert.equal(page.sheet, 'Annexe 1');
  assert.equal(page.range, 'A2:B3');
  assert.deepEqual(page.values, [['c', 'd'], ['e', 'f']]);
});

test('positions stay right for a table that does not start in A1, across pages', async () => {
  const rows = sales(3000);
  const pane = excel({ active: 'Ventes', sheets: [{ name: 'Ventes', rows, top: 4, left: 2 }] });
  const pages = await readAll(pane);
  assert.equal(pages[0].range.split(':')[0], 'C5');
  assert.deepEqual(pages.flatMap((page) => page.values), rows);
  const last = pages.at(-1).range.split(':')[1];
  assert.equal(last, `G${4 + rows.length}`);
});

test('a sheet whose name needs quoting can be continued with the position handed back', async () => {
  const rows = sales(3000);
  const pane = excel({ active: 'Ventes', sheets: [{ name: 'Ventes', rows: [] }, { name: "Budget d'équipe 2025", rows }] });
  const pages = await readAll(pane, "Budget d'équipe 2025");
  assert.ok(pages[0].next.startsWith("'Budget d''équipe 2025'!"));
  assert.deepEqual(pages.flatMap((page) => page.values), rows);
});

test('formulas are reported with the cell that holds them, next to the computed values', async () => {
  const pane = excel({
    active: 'Calc',
    sheets: [{ name: 'Calc', rows: [['Prix', 'Qté', 'Total'], [10, 3, 30], [5, 2, 10]], formulas: { C2: '=A2*B2', C3: '=A3*B3' } }],
  });
  const page = await pane.exec('office_read', {});
  assert.deepEqual(page.values[1], [10, 3, 30]);
  assert.deepEqual(page.formulas, [
    { cell: 'C2', formula: '=A2*B2' },
    { cell: 'C3', formula: '=A3*B3' },
  ]);
});

test('a sheet wider than a read says how many columns were left out', async () => {
  const wide = [Array.from({ length: 250 }, (_, i) => `c${i}`), Array.from({ length: 250 }, (_, i) => i)];
  const pane = excel(book(wide));
  const page = await pane.exec('office_read', {});
  assert.equal(page.values[0].length, pane.get('READ_COLS'));
  assert.equal(page.columnsCut, 250 - pane.get('READ_COLS'));
});

test('an empty sheet and an empty range read as empty', async () => {
  const pane = excel(book([['a']], [{ name: 'Vide', rows: [] }]));
  assert.deepEqual(await pane.exec('office_read', { target: 'Vide' }), {
    host: 'Excel',
    title: 'Ventes.xlsx',
    sheet: 'Vide',
    empty: true,
    values: [],
  });
  assert.equal((await pane.exec('office_read', { target: 'Ventes!X50:Z60' })).empty, true);
});

test('an unknown sheet is an error', async () => {
  const pane = excel(book([['a']]));
  await assert.rejects(pane.exec('office_read', { target: 'Nope!A1' }), /ItemNotFound: Nope/);
});

test('search looks through every sheet and returns each hit with its row', async () => {
  const rows = sales(5000);
  rows[4321][4] = 'Litige en cours';
  const pane = excel(book(rows, [{ name: 'Notes', rows: [['Sujet'], ['Rien'], ['Dossier en LITIGE']] }]));
  const found = await pane.exec('office_search', { query: 'litige' });
  assert.equal(found.matches, 2);
  assert.deepEqual(found.results, [
    { at: 'Ventes!E4322', row: rows[4321] },
    { at: 'Notes!A3', row: ['Dossier en LITIGE'] },
  ]);
  // A hit position is directly readable.
  const page = await pane.exec('office_read', { target: found.results[0].at });
  assert.deepEqual(page.values, [['Litige en cours']]);
});

test('search counts every match but only brings back the rows of the first ones', async () => {
  const pane = excel(book(sales(200000)));
  const found = await pane.exec('office_search', { query: 'annulé', limit: 10 });
  assert.equal(found.matches, 20000);
  assert.equal(found.results.length, 10);
  assert.deepEqual(found.results.map((hit) => hit.at).slice(0, 3), ['Ventes!E11', 'Ventes!E21', 'Ventes!E31']);
  assert.ok(pane.stats.cellsSent <= 10 * pane.get('ROW_CONTEXT'), 'only the rows of the returned hits are read');
});

test('search can be narrowed to a sheet or to a range', async () => {
  const pane = excel(book(sales(100), [{ name: 'Notes', rows: [['Nord et Sud']] }]));
  const sheet = await pane.exec('office_search', { query: 'nord', in: 'Notes' });
  assert.deepEqual(sheet.results.map((hit) => hit.at), ['Notes!A1']);
  const range = await pane.exec('office_search', { query: 'nord', in: 'Ventes!B2:B10' });
  assert.deepEqual(range.results.map((hit) => hit.at), ['Ventes!B5', 'Ventes!B9']);
});

test('search with no match is an empty answer, not an error', async () => {
  const pane = excel(book(sales(10)));
  const found = await pane.exec('office_search', { query: 'introuvable' });
  assert.equal(found.matches, 0);
  assert.deepEqual(found.results, []);
});

// Reference implementation of the aggregation, written the obvious way.
function expected(rows, op, valueCol, groupCol, keep = () => true) {
  const groups = new Map();
  rows.slice(1).filter(keep).forEach((row) => {
    const value = row[valueCol];
    if (typeof value !== 'number') return;
    const key = String(row[groupCol]);
    groups.set(key, [...(groups.get(key) || []), value]);
  });
  const figure = (values) =>
    ({
      sum: values.reduce((a, b) => a + b, 0),
      count: values.length,
      average: values.reduce((a, b) => a + b, 0) / values.length,
      min: Math.min(...values),
      max: Math.max(...values),
    })[op];
  return new Map([...groups].map(([key, values]) => [key, figure(values)]));
}

test('a total per group over 200,000 rows matches a straightforward computation', async () => {
  const rows = sales(200000);
  const pane = excel(book(rows));
  const out = await pane.exec('excel_compute', { op: 'sum', value: 'Montant', group_by: 'Région' });
  const reference = expected(rows, 'sum', 3, 1);
  assert.equal(out.rowsRead, 200000);
  assert.equal(out.partial, undefined);
  assert.equal(out.groupCount, 4);
  for (const group of out.groups) {
    assert.ok(Math.abs(group.value - reference.get(group.key)) < 1e-6, group.key);
    assert.equal(group.rows, 50000);
  }
  assert.deepEqual(out.groups.map((g) => g.value), [...out.groups.map((g) => g.value)].sort((a, b) => b - a), 'largest first');
});

test('an aggregation reads the columns it needs and no other, in bounded blocks', async () => {
  const pane = excel(book(sales(200000)));
  await pane.exec('excel_compute', { op: 'sum', value: 'Montant', group_by: 'Région' });
  assert.equal(pane.stats.cellsSent, 5 + 2 * 200000, 'the header row, then two columns out of five');
  assert.ok(pane.stats.largestRead <= pane.get('COMPUTE_CHUNK'));
});

test('every operation gives the same figure as the reference, with and without a filter', async () => {
  const rows = sales(5000);
  const pane = excel(book(rows));
  for (const op of ['sum', 'count', 'average', 'min', 'max']) {
    const all = await pane.exec('excel_compute', { op, value: 'Montant', group_by: 'Année' });
    const reference = expected(rows, op, 3, 2);
    assert.equal(all.groupCount, reference.size, op);
    for (const group of all.groups) assert.ok(Math.abs(group.value - reference.get(group.key)) < 1e-6, `${op} ${group.key}`);

    const filtered = await pane.exec('excel_compute', {
      op,
      value: 'Montant',
      group_by: 'Année',
      filter: { column: 'Région', equals: 'nord ' },
    });
    const kept = expected(rows, op, 3, 2, (row) => row[1] === 'Nord');
    for (const group of filtered.groups) {
      assert.ok(Math.abs(group.value - kept.get(group.key)) < 1e-6, `${op} filtered ${group.key}`);
    }
  }
});

test('a single figure is returned when there is no grouping', async () => {
  const rows = sales(1000);
  const pane = excel(book(rows));
  const total = await pane.exec('excel_compute', { op: 'sum', value: 'Montant' });
  assert.ok(Math.abs(total.result - rows.slice(1).reduce((sum, row) => sum + row[3], 0)) < 1e-6);
  assert.equal(total.rows, 1000);
  const cancelled = await pane.exec('excel_compute', { op: 'count', filter: { column: 'Statut', equals: 'Annulé' } });
  assert.equal(cancelled.result, 100);
  const by2024 = await pane.exec('excel_compute', { op: 'count', filter: { column: 'Année', equals: 2024 } });
  assert.equal(by2024.result, rows.filter((row) => row[2] === 2024).length, 'a numeric filter matches numeric cells');
});

test('a filter that matches nothing gives zero, or no figure for an average', async () => {
  const pane = excel(book(sales(100)));
  const filter = { column: 'Région', equals: 'Centre' };
  assert.equal((await pane.exec('excel_compute', { op: 'sum', value: 'Montant', filter })).result, 0);
  assert.equal((await pane.exec('excel_compute', { op: 'count', filter })).result, 0);
  assert.equal((await pane.exec('excel_compute', { op: 'average', value: 'Montant', filter })).result, null);
});

test('text and empty cells in a numeric column are left out of the figures', async () => {
  const pane = excel(book([['Client', 'Montant'], ['a', 10], ['b', 'n/a'], ['c', ''], ['d', 30], ['e', null]]));
  const sum = await pane.exec('excel_compute', { op: 'sum', value: 'Montant' });
  assert.deepEqual([sum.result, sum.rows], [40, 2]);
  const average = await pane.exec('excel_compute', { op: 'average', value: 'Montant' });
  assert.equal(average.result, 20);
  const filled = await pane.exec('excel_compute', { op: 'count', value: 'Montant' });
  assert.equal(filled.result, 3, 'count with a column counts its non-empty cells');
});

test('a column can be named by its letter, including when the table is not in A1', async () => {
  const rows = sales(200);
  const pane = excel({ active: 'Ventes', sheets: [{ name: 'Ventes', rows, top: 2, left: 3 }] });
  // The table starts in D3: "Montant" is in column G.
  const byLetter = await pane.exec('excel_compute', { op: 'sum', value: 'G' });
  const byName = await pane.exec('excel_compute', { op: 'sum', value: 'montant' });
  assert.equal(byLetter.result, byName.result);
  await assert.rejects(pane.exec('excel_compute', { op: 'sum', value: 'A' }), /"A" is not a column of Ventes/);
});

test('an unknown column is refused with the list of the real ones', async () => {
  const pane = excel(book(sales(10)));
  await assert.rejects(
    pane.exec('excel_compute', { op: 'sum', value: 'Chiffre' }),
    /value "Chiffre" is not a column of Ventes — columns: Id, Région, Année, Montant, Statut/
  );
  await assert.rejects(pane.exec('excel_compute', { op: 'median', value: 'Montant' }), /op must be one of/);
  await assert.rejects(pane.exec('excel_compute', { op: 'sum' }), /value \(the column to sum\) is required/);
  await assert.rejects(pane.exec('excel_compute', { op: 'count', filter: { column: 'Statut' } }), /filter needs column and equals/);
});

test('more groups than fit in an answer: the largest ones, and how many there are in all', async () => {
  const rows = [['Client', 'Montant']];
  for (let i = 1; i <= 500; i++) rows.push([`Client ${i}`, i]);
  const pane = excel(book(rows));
  const out = await pane.exec('excel_compute', { op: 'sum', value: 'Montant', group_by: 'Client' });
  assert.equal(out.groupCount, 500);
  assert.equal(out.groups.length, pane.get('COMPUTE_GROUPS'));
  assert.deepEqual(out.groups[0], { key: 'Client 500', value: 500, rows: 1 });
});

test('an aggregation that runs out of time returns figures flagged as partial, with the rows covered', async () => {
  const rows = sales(200000);
  const pane = excel(book(rows), { msPerSync: 4000 });
  const out = await pane.exec('excel_compute', { op: 'count', group_by: 'Région' });
  assert.equal(out.partial, true);
  assert.equal(out.rowsTotal, 200000);
  assert.ok(out.rowsRead > 0 && out.rowsRead < 200000);
  assert.equal(out.groups.reduce((sum, group) => sum + group.rows, 0), out.rowsRead, 'the figures cover exactly the rows read');
});

test('a block of values is written from its top-left cell, on the requested sheet', async () => {
  const workbook = book([['a']], [{ name: 'Cible', rows: [] }]);
  const pane = excel(workbook);
  const done = await pane.exec('excel_write_range', { sheet: 'Cible', address: 'B2', values: [['Nom', 'Total'], ['x', '=1+1']] });
  assert.equal(done.address, 'Cible!B2:C3');
  assert.deepEqual(workbook.sheets[1].rows[1].slice(1), ['Nom', 'Total']);
  assert.deepEqual(workbook.sheets[1].rows[2].slice(1), ['x', '=1+1']);
  assert.deepEqual(workbook.sheets[0].rows, [['a']], 'the active sheet is untouched');
});

test('a ragged or empty block is refused before anything is written', async () => {
  const workbook = book([['a']]);
  const pane = excel(workbook);
  for (const values of [[], [[1, 2], [3]], 'text', undefined]) {
    await assert.rejects(pane.exec('excel_write_range', { address: 'A1', values }), /non-empty rectangular 2D array/);
  }
  await assert.rejects(pane.exec('excel_write_range', { values: [[1]] }), /address is required/);
  assert.equal(pane.stats.syncs, 0);
  assert.deepEqual(workbook.sheets[0].rows, [['a']]);
});

test('a sheet is added without taking the user away from the one on screen', async () => {
  const workbook = book([['a']]);
  const pane = excel(workbook);
  const done = await pane.exec('excel_add_sheet', { name: ' Synthèse ' });
  assert.deepEqual(done, { ok: true, sheet: 'Synthèse' });
  assert.deepEqual(workbook.sheets.map((sheet) => sheet.name), ['Ventes', 'Synthèse']);
  assert.equal(workbook.active, 'Ventes');
  // The new sheet is reachable by name right away.
  await pane.exec('excel_write_range', { sheet: 'Synthèse', address: 'A1', values: [['Total']] });
  assert.deepEqual(workbook.sheets[1].rows, [['Total']]);
  await assert.rejects(pane.exec('excel_add_sheet', { name: 'Ventes' }), /ItemAlreadyExists/);
  await assert.rejects(pane.exec('excel_add_sheet', { name: '  ' }), /name is required/);
});

test('a chart is created on the sheet that holds the range, with the requested type', async () => {
  const workbook = book(sales(12));
  const pane = excel(workbook);
  await pane.exec('excel_add_chart', { address: 'B1:D13', type: 'line', title: 'Ventes par mois' });
  assert.equal(workbook.sheets[0].charts.length, 1);
  assert.equal(workbook.sheets[0].charts[0].type, 'Line');
  assert.equal(workbook.sheets[0].charts[0].source, 'Ventes!B1:D13');
  assert.equal(workbook.sheets[0].charts[0].title.text, 'Ventes par mois');
  await assert.rejects(pane.exec('excel_add_chart', { address: 'A1:B2', type: 'radar' }), /type must be one of column, bar, line, pie, area, scatter/);
});

test('the context shared with the agent is the top of the active sheet, flagged as partial', async () => {
  const pane = excel(book(sales(200000)));
  const doc = await pane.get('readDocument')();
  assert.ok(doc.content.startsWith('Sheet: Ventes\nRange: A1:E'));
  assert.equal(doc.more, true);
  assert.ok(pane.stats.largestRead <= pane.get('READ_CELLS'));
});
