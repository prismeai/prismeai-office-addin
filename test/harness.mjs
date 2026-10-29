// Loads the task pane scripts the way an Office webview does (the classic
// scripts of taskpane.html, in its order), against in-memory documents that stand in for Word, Excel,
// PowerPoint and Outlook.
//
// The stand-ins reproduce what matters for the code under test, as measured on
// the real hosts: data only exists after `load` + `sync`, every `sync` is a
// round trip that costs time, a PowerPoint batch holding one shape without a
// text frame is rejected as a whole and its context is then unusable. They also
// count what crosses the wire, so a test can assert that a tool stays cheap on
// a large document, not only that it answers correctly.
import fs from 'node:fs';
import vm from 'node:vm';

const ROOT = new URL('../', import.meta.url);
// The local scripts of the page, in the order the page loads them: a script
// missing from taskpane.html, or loaded too early, fails the tests.
const SCRIPTS = [...fs.readFileSync(new URL('taskpane.html', ROOT), 'utf8').matchAll(/<script src="([^":]+)"><\/script>/g)]
  .map((match) => match[1])
  .map((file) => ({ file, source: fs.readFileSync(new URL(file, ROOT), 'utf8') }));

const HOSTS = { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint', Outlook: 'Outlook' };

// Objects built inside the vm have another realm's prototypes: compare plain data.
export const plain = (value) => (value === undefined ? undefined : JSON.parse(JSON.stringify(value)));

const fieldsOf = (spec) =>
  (typeof spec === 'string' ? spec : Array.isArray(spec) ? spec.join(',') : spec.select || '')
    .split(',')
    .map((field) => field.trim().replace(/^items\//, ''))
    .filter(Boolean);

// ── Word ────────────────────────────────────────────────────────────────────
// doc.paragraphs: [{ text, style? ('Heading1'…), comments? }]
function fakeWord(doc, env) {
  const stats = env.stats;
  const context = () => {
    const pending = [];
    const edits = [];

    const paragraphProxy = (model, loaded = new Set()) => {
      const need = (field) => {
        if (!loaded.has(field)) throw new Error(`PropertyNotLoaded: ${field}`);
      };
      return {
        get text() {
          need('text');
          return model.text;
        },
        get outlineLevel() {
          need('outlineLevel');
          const match = /^Heading(\d)$/.exec(model.style || '');
          return match ? +match[1] : 10;
        },
        get styleBuiltIn() {
          need('styleBuiltIn');
          return model.style || 'Normal';
        },
        set styleBuiltIn(value) {
          model.style = value;
        },
        load(spec) {
          const fields = fieldsOf(spec);
          pending.push(() => {
            fields.forEach((field) => loaded.add(field));
            if (fields.includes('text')) {
              stats.paragraphsSent++;
              stats.charsSent += model.text.length;
            }
          });
        },
        search: (text) => search([model], text),
        insertParagraph(text, where) {
          if (where !== 'After') throw new Error('unsupported location ' + where);
          const added = { text };
          doc.paragraphs.splice(doc.paragraphs.indexOf(model) + 1, 0, added);
          return paragraphProxy(added);
        },
      };
    };

    // Case-sensitive, like the add-in asks for (matchCase: true).
    const search = (models, text) => {
      const items = [];
      models.forEach((model) => {
        for (let i = model.text.indexOf(text); i >= 0; i = model.text.indexOf(text, i + text.length)) {
          const index = i;
          items.push({
            insertText(replacement, mode) {
              if (mode !== 'Replace') throw new Error('unsupported mode ' + mode);
              edits.push({ model, index, length: text.length, replacement });
            },
            insertComment(comment) {
              (model.comments = model.comments || []).push(comment);
            },
            paragraphs: { getLast: () => paragraphProxy(model) },
          });
        }
      });
      return { load() {}, items };
    };

    const collection = { items: [] };
    collection.load = (spec) => {
      const fields = fieldsOf(spec);
      pending.push(() => {
        const windowed = typeof spec === 'object' && !Array.isArray(spec) && !env.options.ignoreWindow;
        const skip = windowed ? spec.skip || 0 : 0;
        const models = windowed ? doc.paragraphs.slice(skip, skip + (spec.top ?? Infinity)) : doc.paragraphs;
        collection.items = models.map((model) => paragraphProxy(model, new Set(fields)));
        if (fields.includes('text')) {
          stats.paragraphsSent += models.length;
          stats.charsSent += models.reduce((sum, model) => sum + model.text.length, 0);
        }
      });
    };

    return {
      document: {
        body: {
          paragraphs: collection,
          search: (text) => search(doc.paragraphs, text),
          insertParagraph(text, where) {
            if (where !== 'End') throw new Error('unsupported location ' + where);
            const added = { text };
            doc.paragraphs.push(added);
            return paragraphProxy(added);
          },
        },
      },
      async sync() {
        env.roundTrip();
        pending.splice(0).forEach((apply) => apply());
        // Later occurrences first, so earlier offsets stay valid.
        edits
          .splice(0)
          .sort((a, b) => b.index - a.index)
          .forEach(({ model, index, length, replacement }) => {
            model.text = model.text.slice(0, index) + replacement + model.text.slice(index + length);
          });
      },
    };
  };
  return { run: async (task) => task(context()) };
}

// ── Excel ───────────────────────────────────────────────────────────────────
// book: { active, selection?, sheets: [{ name, rows, top?, left?, formulas? ({ B2: '=A2*2' }) }] }
const SHEET_ROWS = 1048576;
const SHEET_COLS = 16384;
const MAX_PAYLOAD_CELLS = 5000000; // a real host refuses a response that large

const letters = (index) => {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
};
const number = (name) => [...name].reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0) - 1;
const quoted = (name) => (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);

function bounds(address) {
  const text = address.replace(/\$/g, '').toUpperCase();
  let match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(text);
  if (match) {
    const r = +match[2] - 1;
    const c = number(match[1]);
    return { r, c, nr: (match[4] ? +match[4] - 1 : r) - r + 1, nc: (match[3] ? number(match[3]) : c) - c + 1 };
  }
  match = /^([A-Z]+):([A-Z]+)$/.exec(text);
  if (match) return { r: 0, c: number(match[1]), nr: SHEET_ROWS, nc: number(match[2]) - number(match[1]) + 1 };
  match = /^(\d+):(\d+)$/.exec(text);
  if (match) return { r: +match[1] - 1, c: 0, nr: +match[2] - +match[1] + 1, nc: SHEET_COLS };
  throw new Error('InvalidArgument: ' + address);
}

function fakeExcel(book, env) {
  const stats = env.stats;
  const NULL = { isNullObject: true, load() {} };
  const top = (sheet) => sheet.top || 0;
  const left = (sheet) => sheet.left || 0;
  const width = (sheet) => {
    if (sheet.measured !== sheet.rows.length) {
      sheet.width = sheet.rows.reduce((max, row) => Math.max(max, row.length), 0);
      sheet.measured = sheet.rows.length;
    }
    return sheet.width;
  };
  const cell = (sheet, r, c) => {
    const row = sheet.rows[r - top(sheet)];
    const value = row ? row[c - left(sheet)] : undefined;
    return value === undefined || value === null ? '' : value;
  };
  const formula = (sheet, r, c) => (sheet.formulas || {})[letters(c) + (r + 1)] || cell(sheet, r, c);

  const context = () => {
    const pending = [];

    const range = (sheet, r, c, nr, nc) => {
      const loaded = new Set();
      const matrix = (read) => {
        const out = [];
        for (let i = 0; i < nr; i++) {
          const row = [];
          for (let j = 0; j < nc; j++) row.push(read(sheet, r + i, c + j));
          out.push(row);
        }
        return out;
      };
      const need = (field) => {
        if (!loaded.has(field)) throw new Error(`PropertyNotLoaded: ${field}`);
      };
      const used = () => {
        if (!sheet.rows.length || !width(sheet)) return NULL;
        const r0 = Math.max(r, top(sheet));
        const c0 = Math.max(c, left(sheet));
        const r1 = Math.min(r + nr, top(sheet) + sheet.rows.length);
        const c1 = Math.min(c + nc, left(sheet) + width(sheet));
        return r1 > r0 && c1 > c0 ? range(sheet, r0, c0, r1 - r0, c1 - c0) : NULL;
      };
      const self = {
        isNullObject: false,
        rowIndex: r,
        columnIndex: c,
        rowCount: nr,
        columnCount: nc,
        cellCount: nr * nc,
        get address() {
          const end = nr * nc > 1 ? `:${letters(c + nc - 1)}${r + nr}` : '';
          return `${quoted(sheet.name)}!${letters(c)}${r + 1}${end}`;
        },
        load(spec) {
          const fields = fieldsOf(spec);
          pending.push(() => {
            const heavy = fields.filter((field) => field === 'values' || field === 'formulas').length;
            if (heavy && nr * nc > MAX_PAYLOAD_CELLS) throw new Error('ResponsePayloadSizeLimitExceeded');
            stats.cellsSent += heavy * nr * nc;
            stats.largestRead = Math.max(stats.largestRead, heavy ? nr * nc : 0);
            fields.forEach((field) => loaded.add(field));
          });
        },
        get values() {
          need('values');
          return matrix(cell);
        },
        set values(rows) {
          pending.push(() => {
            rows.forEach((row, i) =>
              row.forEach((value, j) => {
                const y = r + i - top(sheet);
                while (sheet.rows.length <= y) sheet.rows.push([]);
                sheet.rows[y][c + j - left(sheet)] = value;
              })
            );
          });
        },
        get formulas() {
          need('formulas');
          return matrix(formula);
        },
        getUsedRangeOrNullObject: used,
        getCell: (i, j) => range(sheet, r + i, c + j, 1, 1),
        getResizedRange: (rows, cols) => range(sheet, r, c, nr + rows, nc + cols),
        findAllOrNullObject(query, criteria) {
          if (criteria.completeMatch || criteria.matchCase) throw new Error('unsupported criteria');
          const area = used();
          if (area.isNullObject) return NULL;
          const needle = query.toLowerCase();
          // Vertical runs of hits form one area, as Excel reports them.
          const areas = [];
          let count = 0;
          for (let x = area.columnIndex; x < area.columnIndex + area.columnCount; x++) {
            let start = -1;
            for (let y = area.rowIndex; y <= area.rowIndex + area.rowCount; y++) {
              const hit =
                y < area.rowIndex + area.rowCount && String(cell(sheet, y, x)).toLowerCase().includes(needle);
              if (hit && start < 0) start = y;
              if (!hit && start >= 0) {
                areas.push({ c: x, r0: start, r1: y - 1 });
                count += y - start;
                start = -1;
              }
            }
          }
          if (!areas.length) return NULL;
          areas.sort((a, b) => a.r0 - b.r0 || a.c - b.c);
          return {
            isNullObject: false,
            load() {},
            cellCount: count,
            address: areas
              .map(({ c: x, r0, r1 }) => {
                const end = r1 > r0 ? `:${letters(x)}${r1 + 1}` : '';
                return `${quoted(sheet.name)}!${letters(x)}${r0 + 1}${end}`;
              })
              .join(','),
          };
        },
      };
      return self;
    };

    const worksheet = (sheet) => ({
      load() {},
      name: sheet.name,
      visibility: sheet.hidden ? 'Hidden' : 'Visible',
      getRange(address) {
        if (!address) return range(sheet, 0, 0, SHEET_ROWS, SHEET_COLS);
        const b = bounds(address);
        return range(sheet, b.r, b.c, b.nr, b.nc);
      },
      getRangeByIndexes: (r, c, nr, nc) => range(sheet, r, c, nr, nc),
      getUsedRangeOrNullObject: () => range(sheet, 0, 0, SHEET_ROWS, SHEET_COLS).getUsedRangeOrNullObject(),
      findAllOrNullObject: (query, criteria) =>
        range(sheet, 0, 0, SHEET_ROWS, SHEET_COLS).findAllOrNullObject(query, criteria),
      activate() {
        book.active = sheet.name;
      },
      charts: {
        add(type, source, seriesBy) {
          const chart = { type, source: source.address, seriesBy, title: { text: '' }, name: 'Chart 1', load() {} };
          (sheet.charts = sheet.charts || []).push(chart);
          return chart;
        },
      },
    });
    const find = (name) => {
      const sheet = book.sheets.find((candidate) => candidate.name === name);
      if (!sheet) throw new Error(`ItemNotFound: ${name}`);
      return sheet;
    };

    return {
      workbook: {
        worksheets: {
          load() {},
          get items() {
            return book.sheets.map(worksheet);
          },
          getActiveWorksheet: () => worksheet(find(book.active)),
          getItem: (name) => worksheet(find(name)),
          add(name) {
            if (book.sheets.some((sheet) => sheet.name === name)) throw new Error('ItemAlreadyExists');
            const sheet = { name, rows: [] };
            book.sheets.push(sheet);
            return worksheet(sheet);
          },
        },
        getSelectedRange() {
          const b = bounds(book.selection || 'A1');
          return range(find(book.active), b.r, b.c, b.nr, b.nc);
        },
      },
      async sync() {
        env.roundTrip();
        pending.splice(0).forEach((apply) => apply());
      },
    };
  };
  return { run: async (task) => task(context()) };
}

// ── PowerPoint ──────────────────────────────────────────────────────────────
// deck.slides: [{ shapes: [{ id, name, type, kind?, contained?, text }] }]
// `text: null` = a shape without a text frame (picture, chart, table…).
// options.unlisted: slide numbers PowerPoint on the web has not loaded yet —
// they list no shape until a slide is brought on screen.
// options.onScreen: the slide the user has selected.
// options.slowSelect: a promise the next selection waits on before PowerPoint answers.
function fakePowerPoint(deck, env) {
  const stats = env.stats;
  const ids = new WeakMap();
  let lastId = 255;
  const slideId = (slide) => {
    if (!ids.has(slide)) ids.set(slide, String(++lastId));
    return ids.get(slide);
  };
  const numberOf = (id) => deck.slides.findIndex((slide) => slideId(slide) === id) + 1;
  let selection = env.options.onScreen ? [slideId(deck.slides[env.options.onScreen - 1])] : [];
  env.onScreen = () => selection.map(numberOf);
  const context = () => {
    let broken = false;
    let selecting = false;
    let placeholderLoads = 0;
    const textLoads = [];
    const counts = [];
    const edits = [];

    const shapeProxy = (shape, index) => ({
      id: shape ? shape.id : undefined,
      name: shape ? shape.name : undefined,
      type: shape ? shape.type : undefined,
      placeholderFormat: {
        load() {
          placeholderLoads++;
        },
        type: shape ? shape.kind : undefined,
        containedType: shape ? shape.contained ?? null : undefined,
      },
      textFrame: {
        textRange: {
          load() {
            // A slide the host has not finished loading refuses its text.
            textLoads.push((env.options.unloaded || []).includes(index + 1) ? null : shape);
          },
          get text() {
            return shape.text;
          },
          set text(value) {
            edits.push(() => {
              if (!shape || shape.text === null) throw new Error('InvalidArgument');
              shape.text = value;
            });
          },
        },
        set autoSizeSetting(value) {
          edits.push(() => {
            if (env.options.noAutoSize) throw new Error('NotImplemented');
            shape.autoSize = value;
          });
        },
      },
    });

    const slideProxy = (index) => ({
      load() {},
      get id() {
        return slideId(deck.slides[index]);
      },
      shapes: {
        load() {
          edits.push(() => {
            if (!deck.slides[index]) throw new Error('InvalidArgument');
            stats.shapesListed += deck.slides[index].shapes.length;
          });
        },
        get items() {
          if ((env.options.unlisted || []).includes(index + 1)) return [];
          return deck.slides[index].shapes.map((shape) => shapeProxy(shape, index));
        },
        getItem: (id) => shapeProxy(deck.slides[index].shapes.find((shape) => shape.id === id), index),
        addTextBox(text) {
          const shape = { id: String(900 + deck.slides[index].shapes.length), name: 'TextBox', type: 'TextBox', text };
          edits.push(() => deck.slides[index].shapes.push(shape));
          return { load() {}, id: shape.id };
        },
      },
      delete() {
        edits.push(() => deck.slides.splice(index, 1));
      },
      moveTo(position) {
        edits.push(() => deck.slides.splice(position, 0, deck.slides.splice(index, 1)[0]));
      },
    });

    return {
      presentation: {
        slides: {
          getCount() {
            const count = {};
            counts.push(count);
            return count;
          },
          getItemAt: slideProxy,
        },
        getSelectedSlides() {
          const selected = { items: [], load() {} };
          edits.push(() => {
            selected.items = selection.map((id) => ({ id }));
          });
          return selected;
        },
        setSelectedSlides(slideIds) {
          selecting = true;
          edits.push(() => {
            selection = [...slideIds];
            stats.selections.push([...slideIds].map(numberOf));
            // A slide on screen makes the host load that part of the deck.
            env.options.unlisted = [];
          });
        },
      },
      async sync() {
        env.roundTrip();
        if (selecting && env.options.slowSelect) {
          const answer = env.options.slowSelect;
          env.options.slowSelect = null;
          await answer;
        }
        selecting = false;
        // What a failed batch leaves behind can't be synced again.
        if (broken) throw new Error("Cannot read properties of null (reading 'statement')");
        const fail = (message) => {
          broken = true;
          placeholderLoads = 0;
          textLoads.length = 0;
          throw new Error(message);
        };
        if (placeholderLoads && env.options.noPlaceholderInfo) fail('NotImplemented');
        placeholderLoads = 0;
        if (textLoads.some((shape) => !shape || shape.text === null)) fail('InvalidArgument');
        stats.textsSent += textLoads.splice(0).length;
        try {
          edits.splice(0).forEach((apply) => apply());
        } catch (err) {
          fail(err.message);
        }
        counts.splice(0).forEach((count) => {
          count.value = deck.slides.length;
        });
      },
    };
  };
  return { run: async (task) => task(context()) };
}

// ── The pane ────────────────────────────────────────────────────────────────
// loadPane({ host, platform, word | excel | powerpoint | mail, selection, msPerSync, debug, options })
export function loadPane(setup = {}) {
  const env = {
    options: setup.options || {},
    stats: { syncs: 0, paragraphsSent: 0, charsSent: 0, cellsSent: 0, largestRead: 0, textsSent: 0, shapesListed: 0, selections: [] },
    now: 1700000000000,
    roundTrip() {
      env.stats.syncs++;
      env.now += setup.msPerSync || 0;
    },
  };
  const timers = [];
  const waiting = new Map();
  const sent = [];
  const listeners = []; // the pane's handlers for messages from the chat frame
  const forms = []; // compose forms Outlook was asked to open
  const elements = new Map();
  const element = (id) => {
    if (!elements.has(id)) {
      elements.set(id, {
        value: '',
        style: {},
        classList: { add() {}, remove() {}, toggle() {} },
        addEventListener() {},
        appendChild() {},
        contentWindow: {
          postMessage(message) {
            sent.push(message);
            const done = waiting.get(message.correlationId);
            if (message.type === 'tool-result' && done) done(plain(message));
          },
        },
      });
    }
    return elements.get(id);
  };

  const mail = setup.mail;
  const sandbox = {
    console: { log() {}, warn() {}, error() {} },
    URLSearchParams,
    document: { getElementById: element, createElement: () => element(Symbol()), head: element('head') },
    location: { search: setup.debug ? '?debug=1' : '', hostname: 'addin.example.test' },
    localStorage: { getItem: () => null, setItem() {} },
    addEventListener(type, listener) {
      if (type === 'message') listeners.push(listener);
    },
    // Timers only fire when a test says so: no test waits on a real clock.
    setTimeout(callback, delay) {
      timers.push({ callback, due: env.now + delay, live: true });
      return timers.length;
    },
    clearTimeout(id) {
      if (timers[id - 1]) timers[id - 1].live = false;
    },
    // The real Date, on the test's clock.
    Date: class extends Date {
      static now() {
        return env.now;
      }
    },
    Office: {
      HostType: HOSTS,
      PlatformType: { OfficeOnline: 'OfficeOnline', PC: 'PC', Mac: 'Mac' },
      AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
      CoercionType: { Text: 'text', SlideRange: 'slideRange' },
      EventType: { DocumentSelectionChanged: 'selection', ItemChanged: 'item' },
      onReady: (callback) => callback({ host: setup.host || null }),
      context: {
        platform: setup.platform || 'OfficeOnline',
        requirements: { isSetSupported: () => !env.options.oldHost },
        document: {
          url: setup.url || 'https://files.test/docs/Report.docx?web=1',
          addHandlerAsync() {},
          getSelectedDataAsync(type, callback) {
            if (setup.selection === undefined) {
              callback({ status: 'failed', error: { message: 'The current selection is not supported.' } });
            } else {
              callback({ status: 'succeeded', value: setup.selection });
            }
          },
        },
        mailbox: {
          addHandlerAsync() {},
          displayNewAppointmentForm: (form) => forms.push({ kind: 'meeting', ...form }),
          displayNewMessageForm: (form) => forms.push({ kind: 'message', ...form }),
          item: mail && {
            subject: mail.subject,
            from: mail.from,
            to: mail.to,
            body: { getAsync: (type, callback) => callback({ status: 'succeeded', value: mail.body }) },
          },
        },
      },
    },
    Word: fakeWord(setup.word || { paragraphs: [] }, env),
    Excel: fakeExcel(setup.excel || { active: 'Sheet1', sheets: [{ name: 'Sheet1', rows: [] }] }, env),
    PowerPoint: fakePowerPoint(setup.powerpoint || { slides: [] }, env),
  };
  sandbox.window = sandbox;
  const context = vm.createContext(sandbox);
  SCRIPTS.forEach(({ file, source }) => vm.runInContext(source, context, { filename: file }));
  // Replies go to the chat frame once it has introduced itself.
  vm.runInContext("frameOrigin = 'https://instance.test'; cfg = { instance: 'https://instance.test', agent: 'agent-1', api: 'https://api.instance.test/v2' }", context);

  let calls = 0;
  const pane = {
    stats: env.stats,
    sent,
    forms,
    // A top-level binding of the script: a function, a constant…
    get: (name) => vm.runInContext(name, context),
    // The tool, without the relay around it. Rejects with the tool's own error.
    exec: async (tool, args) => plain(await pane.get('execTool')(tool, args)),
    // The full path of a relayed call: resolves with the message sent back to the chat.
    call(name, args) {
      const correlationId = `call-${++calls}`;
      return new Promise((resolve) => {
        waiting.set(correlationId, resolve);
        pane.get('handleToolExec')({ correlationId, name, args });
      });
    },
    // A message from the chat frame, as the pane receives it.
    frame(message, origin = 'https://instance.test') {
      listeners.forEach((listener) => listener({ origin, data: { source: 'prisme-embed-frame', ...message } }));
    },
    results: () => sent.filter((message) => message.type === 'tool-result').map(plain),
    // The slides selected in PowerPoint, by number.
    onScreen: () => env.onScreen(),
    // Moves the clock forward and fires the timers that are due.
    elapse(ms) {
      env.now += ms;
      timers.filter((timer) => timer.live && timer.due <= env.now).forEach((timer) => {
        timer.live = false;
        timer.callback();
      });
    },
  };
  return pane;
}

// Lets promise chains started by the code under test run to completion.
export const settle = () => new Promise((resolve) => setImmediate(resolve));
