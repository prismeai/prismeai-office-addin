// Excel: outline, read, search, compute — then the tools that change the workbook.

const MAX_SHEETS = 50;
const MAX_HEADERS = 30;
const READ_CELLS = 3000; // cells loaded per read, before the size cut
const READ_COLS = 100;
const MAX_FORMULAS = 100;
const ROW_CONTEXT = 20; // columns of the row shown with a search hit
const COMPUTE_CHUNK = 20000; // rows per round trip
const COMPUTE_GROUPS = 50;
const A1_RANGE =
  /^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$|^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d{1,7}:\$?\d{1,7}$/;

function columnName(index) {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

function columnIndex(name) {
  return name
    .toUpperCase()
    .split('')
    .reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function sheetRef(name) {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

// "Sales!A1:C9" → { sheet, address } · "A1:C9" → { address } · "Sales" → { sheet }.
function splitTarget(target) {
  const text = given(target) ? String(target).trim() : '';
  if (!text) return {};
  const unquote = (name) => name.replace(/^'|'$/g, '').replace(/''/g, "'");
  const bang = text.lastIndexOf('!');
  if (bang >= 0) return { sheet: unquote(text.slice(0, bang)), address: text.slice(bang + 1) || undefined };
  if (A1_RANGE.test(text)) return { address: text };
  return { sheet: unquote(text) };
}

// "C2:D5" / "C2" → 0-based bounds, or null when it isn't a cell range.
function parseA1(address) {
  const match = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(address.replace(/\$/g, '').toUpperCase());
  if (!match) return null;
  const r0 = +match[2] - 1;
  const c0 = columnIndex(match[1]);
  return { r0, c0, r1: match[4] ? +match[4] - 1 : r0, c1: match[3] ? columnIndex(match[3]) : c0 };
}

function targetSheet(ctx, name) {
  return name ? ctx.workbook.worksheets.getItem(name) : ctx.workbook.worksheets.getActiveWorksheet();
}

async function excelOutline() {
  return Excel.run(async (ctx) => {
    officeStep = 'listing the sheets';
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name,items/visibility');
    const active = sheets.getActiveWorksheet();
    active.load('name');
    const selection = ctx.workbook.getSelectedRange();
    selection.load('address');
    await ctx.sync();
    const items = sheets.items.slice(0, MAX_SHEETS);
    officeStep = 'measuring the sheets';
    const used = items.map((sheet) => {
      const range = sheet.getUsedRangeOrNullObject(true);
      range.load(['address', 'rowCount', 'columnCount']);
      return range;
    });
    await ctx.sync();
    // First row of each sheet, as a hint of what the columns hold.
    officeStep = 'reading the header rows';
    const headers = used.map((range) => {
      if (range.isNullObject) return null;
      const row = range.getCell(0, 0).getResizedRange(0, Math.min(range.columnCount, MAX_HEADERS) - 1);
      row.load('values');
      return row;
    });
    await ctx.sync();
    officeStep = '';
    return {
      active: active.name,
      selection: selection.address,
      sheetCount: sheets.items.length,
      sheets: items.map((sheet, i) =>
        used[i].isNullObject
          ? { name: sheet.name, visibility: sheet.visibility, empty: true }
          : {
              name: sheet.name,
              visibility: sheet.visibility,
              range: used[i].address,
              rows: used[i].rowCount,
              columns: used[i].columnCount,
              firstRow: headers[i].values[0],
            }
      ),
    };
  });
}

async function excelRead({ target }, maxChars = MAX_OUTPUT_CHARS) {
  const { sheet, address } = splitTarget(target);
  return Excel.run(async (ctx) => {
    const ws = targetSheet(ctx, sheet);
    ws.load('name');
    // Only the used part: "A:C" must not mean a million empty rows.
    const used = (address ? ws.getRange(address) : ws.getRange()).getUsedRangeOrNullObject(true);
    used.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
    await ctx.sync();
    if (used.isNullObject) return { sheet: ws.name, empty: true, values: [] };
    const cols = Math.min(used.columnCount, READ_COLS);
    const rows = Math.min(used.rowCount, Math.max(1, Math.floor(READ_CELLS / cols)));
    const range = ws.getRangeByIndexes(used.rowIndex, used.columnIndex, rows, cols);
    range.load(['values', 'formulas']);
    await ctx.sync();
    const values = range.values;
    const cellFormulas = range.formulas;
    let kept = 0;
    let chars = 0;
    while (kept < rows) {
      const size = JSON.stringify(values[kept]).length + 1; // + the separator
      if (kept && chars + size > maxChars) break;
      chars += size;
      kept++;
    }
    const first = columnName(used.columnIndex);
    const last = columnName(used.columnIndex + cols - 1);
    const top = used.rowIndex + 1;
    const formulas = [];
    for (let r = 0; r < kept && formulas.length < MAX_FORMULAS; r++) {
      cellFormulas[r].forEach((formula, c) => {
        if (typeof formula === 'string' && formula[0] === '=' && formulas.length < MAX_FORMULAS) {
          formulas.push({ cell: columnName(used.columnIndex + c) + (top + r), formula });
        }
      });
    }
    return {
      sheet: ws.name,
      range: `${first}${top}:${last}${top + kept - 1}`,
      values: values.slice(0, kept),
      formulas: formulas.length ? formulas : undefined,
      rowsLeft: used.rowCount - kept || undefined,
      columnsCut: used.columnCount > cols ? used.columnCount - cols : undefined,
      next:
        kept < used.rowCount
          ? `${sheetRef(ws.name)}!${first}${top + kept}:${last}${used.rowIndex + used.rowCount}`
          : undefined,
    };
  });
}

async function excelSearch({ query, scope, limit }) {
  const { sheet, address } = splitTarget(scope);
  return Excel.run(async (ctx) => {
    let sheets;
    if (sheet || address) {
      sheets = [targetSheet(ctx, sheet)];
    } else {
      const all = ctx.workbook.worksheets;
      all.load('items/name');
      await ctx.sync();
      sheets = all.items.slice(0, MAX_SHEETS);
    }
    officeStep = 'searching the cells';
    // Excel does the search: no cell crosses the wire but the hits.
    const found = sheets.map((ws) => {
      ws.load('name');
      const areas = (address ? ws.getRange(address) : ws).findAllOrNullObject(query, {
        completeMatch: false,
        matchCase: false,
      });
      areas.load('address,cellCount');
      return areas;
    });
    await ctx.sync();
    let matches = 0;
    const hits = [];
    found.forEach((areas, i) => {
      if (areas.isNullObject) return;
      matches += areas.cellCount;
      for (const part of areas.address.split(',')) {
        if (hits.length >= limit) return;
        const bounds = parseA1(part.slice(part.lastIndexOf('!') + 1));
        if (!bounds) continue;
        for (let r = bounds.r0; r <= bounds.r1 && hits.length < limit; r++) {
          for (let c = bounds.c0; c <= bounds.c1 && hits.length < limit; c++) hits.push({ ws: sheets[i], r, c });
        }
      }
    });
    officeStep = 'reading the matching rows';
    hits.forEach((hit) => {
      hit.row = hit.ws.getRangeByIndexes(hit.r, 0, 1, ROW_CONTEXT);
      hit.row.load('values');
    });
    await ctx.sync();
    officeStep = '';
    return {
      query,
      matches,
      results: hits.map((hit) => {
        const row = hit.row.values[0];
        while (row.length && row[row.length - 1] === '') row.pop();
        return { at: `${sheetRef(hit.ws.name)}!${columnName(hit.c)}${hit.r + 1}`, row };
      }),
    };
  });
}

const COMPUTE_OPS = ['sum', 'count', 'average', 'min', 'max'];
// Aggregates a table without handing its rows to the agent: only the columns
// involved are read, by chunks, and only the figures go back.
async function excelCompute({ sheet, op, value, group_by: groupBy, filter }) {
  if (!COMPUTE_OPS.includes(op)) throw new Error(`op must be one of ${COMPUTE_OPS.join(', ')}`);
  if (op !== 'count' && !value) throw new Error(`value (the column to ${op}) is required`);
  if (filter && (!filter.column || filter.equals === undefined)) throw new Error('filter needs column and equals');
  const deadline = Date.now() + READ_BUDGET_MS;
  return Excel.run(async (ctx) => {
    officeStep = 'measuring the sheet';
    const ws = targetSheet(ctx, sheet);
    ws.load('name');
    const used = ws.getUsedRangeOrNullObject(true);
    used.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
    await ctx.sync();
    if (used.isNullObject) throw new Error(`sheet ${ws.name} is empty`);
    const head = ws.getRangeByIndexes(used.rowIndex, used.columnIndex, 1, used.columnCount);
    head.load('values');
    await ctx.sync();
    const headers = head.values[0].map((cell) => String(cell).trim());
    // A column is named by its header, or by its letter.
    const column = (name, label) => {
      const text = String(name).trim();
      let offset = headers.findIndex((header) => header.toLowerCase() === text.toLowerCase());
      if (offset < 0 && /^[A-Za-z]{1,3}$/.test(text)) offset = columnIndex(text) - used.columnIndex;
      if (offset < 0 || offset >= used.columnCount) {
        throw new Error(`${label} "${text}" is not a column of ${ws.name} — columns: ${headers.filter(Boolean).join(', ')}`);
      }
      return offset;
    };
    const cols = {
      value: value ? column(value, 'value') : null,
      group: groupBy ? column(groupBy, 'group_by') : null,
      filter: filter ? column(filter.column, 'filter.column') : null,
    };
    const wanted = filter ? String(filter.equals).trim().toLowerCase() : '';
    const groups = new Map();
    const total = used.rowCount - 1; // first row = headers
    let done = 0;
    officeStep = 'reading the columns';
    while (done < total && Date.now() < deadline) {
      const count = Math.min(COMPUTE_CHUNK, total - done);
      const read = (offset) => {
        if (offset === null) return null;
        const range = ws.getRangeByIndexes(used.rowIndex + 1 + done, used.columnIndex + offset, count, 1);
        range.load('values');
        return range;
      };
      const ranges = [read(cols.value), read(cols.group), read(cols.filter)];
      await ctx.sync();
      const [values, keys, tests] = ranges.map((range) => range && range.values);
      for (let i = 0; i < count; i++) {
        if (tests && String(tests[i][0]).trim().toLowerCase() !== wanted) continue;
        const cell = values ? values[i][0] : 1;
        if (op === 'count' ? cell === '' : typeof cell !== 'number') continue;
        const key = keys ? String(keys[i][0]) : '';
        let acc = groups.get(key);
        if (!acc) groups.set(key, (acc = { rows: 0, sum: 0, min: Infinity, max: -Infinity }));
        acc.rows++;
        if (typeof cell === 'number') {
          acc.sum += cell;
          acc.min = Math.min(acc.min, cell);
          acc.max = Math.max(acc.max, cell);
        }
      }
      done += count;
    }
    officeStep = '';
    const figure = (acc) => {
      const raw = { sum: acc.sum, count: acc.rows, average: acc.sum / acc.rows, min: acc.min, max: acc.max }[op];
      return Number(raw.toPrecision(12));
    };
    const out = { sheet: ws.name, op, value: value || undefined, rowsRead: done };
    if (filter) out.filter = filter;
    // Out of time: the figures cover the first `rowsRead` rows only.
    if (done < total) Object.assign(out, { partial: true, rowsTotal: total });
    if (groupBy) {
      const all = [...groups].map(([key, acc]) => ({ key, value: figure(acc), rows: acc.rows }));
      all.sort((a, b) => b.value - a.value);
      Object.assign(out, { group_by: groupBy, groupCount: all.length, groups: all.slice(0, COMPUTE_GROUPS) });
    } else {
      const acc = groups.get('');
      Object.assign(out, { result: acc ? figure(acc) : op === 'count' || op === 'sum' ? 0 : null, rows: acc ? acc.rows : 0 });
    }
    return out;
  });
}

async function excelWriteRange({ address, values, sheet }) {
  if (!address) throw new Error('address is required');
  const rows = Array.isArray(values) ? values : null;
  if (!rows || !rows.length || !rows.every((r) => Array.isArray(r) && r.length === rows[0].length)) {
    throw new Error('values must be a non-empty rectangular 2D array');
  }
  return Excel.run(async (ctx) => {
    // `address` is the top-left anchor; the range grows to fit `values`.
    const range = targetSheet(ctx, sheet)
      .getRange(address)
      .getCell(0, 0)
      .getResizedRange(rows.length - 1, rows[0].length - 1);
    range.values = rows; // a string starting with "=" is written as a formula
    range.load('address');
    await ctx.sync();
    return { ok: true, address: range.address };
  });
}

async function excelAddSheet({ name }) {
  if (typeof name !== 'string' || !name.trim()) throw new Error('name is required');
  return Excel.run(async (ctx) => {
    // Created first, read in a second round trip: asking for the name of the
    // sheet in the batch that creates it killed the pane in Excel on the web
    // (three times out of three; never since the two are apart).
    // Not activated either: the user keeps the sheet they are looking at.
    const sheet = ctx.workbook.worksheets.add(name.trim());
    await ctx.sync();
    sheet.load('name');
    await ctx.sync();
    return { ok: true, sheet: sheet.name };
  });
}

const CHART_TYPES = {
  column: 'ColumnClustered',
  bar: 'BarClustered',
  line: 'Line',
  pie: 'Pie',
  area: 'Area',
  scatter: 'XYScatter',
};
async function excelAddChart({ address, type, title, sheet }) {
  if (!address) throw new Error('address is required');
  const chartType = CHART_TYPES[type || 'column'];
  if (!chartType) throw new Error(`type must be one of ${Object.keys(CHART_TYPES).join(', ')}`);
  return Excel.run(async (ctx) => {
    const target = targetSheet(ctx, sheet);
    const chart = target.charts.add(chartType, target.getRange(address), 'Auto');
    if (title) chart.title.text = String(title);
    chart.load('name');
    await ctx.sync();
    return { ok: true, chart: chart.name };
  });
}
