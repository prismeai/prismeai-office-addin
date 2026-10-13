// The tools as the agent sees them: the reading verbs that are the same in
// every application, the selection, and the table of every tool by name.
// Loaded after the files of each application.

async function selectedText() {
  const T = Office.HostType;
  if (host === T.Excel) {
    // Text coercion only handles a single cell → read the selected range.
    return Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(['values', 'cellCount']);
      await ctx.sync();
      return range.cellCount > 1 || range.values[0][0] !== '' ? rowsToText(range.values) : '';
    });
  }
  if (host === T.Outlook) {
    const item = Office.context.mailbox.item;
    // Only compose mode exposes the selection.
    if (!item || typeof item.getSelectedDataAsync !== 'function') return '';
    const res = await officeAsync((cb) => item.getSelectedDataAsync(Office.CoercionType.Text, cb));
    return (res && res.data) || '';
  }
  // A selection that isn't text (a slide thumbnail, an image…) is refused by
  // the host: that is an empty selection, not a failure.
  const text = await officeAsync((cb) =>
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, cb)
  ).catch(() => '');
  return String(text || '');
}

async function insertText({ text, location }) {
  if (typeof text !== 'string') throw new Error('text is required');
  const T = Office.HostType;
  if (host === T.Word) {
    await Word.run(async (ctx) => {
      if (location === 'start' || location === 'end') {
        ctx.document.body.insertText(text, location === 'start' ? 'Start' : 'End');
      } else {
        ctx.document.getSelection().insertText(text, 'Replace');
      }
      await ctx.sync();
    });
    return { ok: true };
  }
  if (host === T.Outlook) {
    const item = Office.context.mailbox.item;
    if (!item || typeof item.body.setSelectedDataAsync !== 'function') {
      throw new Error('the message is open in read mode — use outlook_draft_reply');
    }
    await officeAsync((cb) =>
      item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, cb)
    );
    return { ok: true };
  }
  // Excel (active cell) and PowerPoint (selected text box).
  await officeAsync((cb) =>
    Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, cb)
  );
  return { ok: true };
}

// The beginning of the open document / mail, as { title, content, more } —
// `more` when the document goes on past what is returned.
async function readDocument() {
  const T = Office.HostType;
  if (host === T.Word) {
    const { paragraphs, next } = await wordRead({}, CONTEXT_CHARS);
    return { title: documentTitle(), content: paragraphs.map((p) => p.text).join('\n'), more: !!next };
  }
  if (host === T.Excel) {
    const { sheet, range, values, next } = await excelRead({}, CONTEXT_CHARS);
    return {
      title: documentTitle(),
      content: `Sheet: ${sheet}\nRange: ${range || ''}\n\n${rowsToText(values)}`,
      more: !!next,
    };
  }
  if (host === T.PowerPoint) {
    const { total, slides } = await readSlides(0, CONTEXT_SLIDES, Date.now() + READ_BUDGET_MS);
    const text = (slide) =>
      slide.shapes
        .map((shape) => shape.text)
        .filter(Boolean)
        .join('\n');
    return {
      title: documentTitle(),
      content: slides.map((slide) => `--- Slide ${slide.slide} ---\n${text(slide)}`).join('\n\n'),
      more: total > slides.length,
    };
  }
  if (host === T.Outlook) {
    const mail = await readMail();
    if (!mail) return { title: 'Outlook', content: '' };
    return {
      title: mail.subject || 'Outlook message',
      content: `From: ${mail.from}\nTo: ${mail.to.join(', ')}\n\n${mail.body}`,
      more: mail.body.length > CONTEXT_CHARS,
    };
  }
  throw new Error('not running inside an Office application');
}

function located(body) {
  const title = host === Office.HostType.Outlook ? undefined : documentTitle();
  return { host: hostName(), title, ...body };
}

// What the document is made of, and how big.
async function officeOutline(args) {
  const T = Office.HostType;
  if (host === T.Word) return located(await wordOutline(args));
  if (host === T.Excel) return located(await excelOutline());
  if (host === T.PowerPoint) return located(await powerPointOutline(args));
  const mail = await openMail();
  return located({ subject: mail.subject, from: mail.from, to: mail.to, characters: mail.body.length });
}

// One slice of it.
async function officeRead(args) {
  const T = Office.HostType;
  if (host === T.Word) return located(await wordRead(args));
  if (host === T.Excel) return located(await excelRead(args));
  if (host === T.PowerPoint) return located(await powerPointRead(args));
  return located(await mailRead(args));
}

// Where a text appears in it.
async function officeSearch({ query, in: scope, limit }) {
  if (typeof query !== 'string' || !query.trim()) throw new Error('query is required');
  if (query.length > MAX_SEARCH || /[\r\n]/.test(query)) {
    throw new Error(`query must be a single line of at most ${MAX_SEARCH} characters`);
  }
  const max = Math.min(Math.max(Number(limit) || SEARCH_LIMIT, 1), MAX_SEARCH_LIMIT);
  const args = { query: query.trim(), scope, limit: max };
  const T = Office.HostType;
  if (host === T.Word) return located(await wordSearch(args));
  if (host === T.Excel) return located(await excelSearch(args));
  if (host === T.PowerPoint) return located(await powerPointSearch(args));
  return located(await mailSearch(args));
}

// Every tool the pane runs: the application it needs (none = any) and what
// runs it. The names are the ones of the tool catalogue on the instance.
const TOOLS = {
  office_outline: { run: officeOutline },
  office_read: { run: officeRead },
  office_search: { run: officeSearch },
  office_get_selection: {
    run: async () => ({ host: hostName(), text: (await selectedText()).slice(0, MAX_OUTPUT_CHARS) }),
  },
  office_insert_text: { run: insertText },
  word_replace_text: { host: 'Word', run: wordReplaceText },
  word_insert_paragraph: { host: 'Word', run: wordInsertParagraph },
  word_add_comment: { host: 'Word', run: wordAddComment },
  excel_compute: { host: 'Excel', run: excelCompute },
  excel_write_range: { host: 'Excel', run: excelWriteRange },
  excel_add_sheet: { host: 'Excel', run: excelAddSheet },
  excel_add_chart: { host: 'Excel', run: excelAddChart },
  powerpoint_write_slide: { host: 'PowerPoint', run: powerPointWriteSlide },
  powerpoint_add_slide: { host: 'PowerPoint', run: powerPointAddSlide },
  powerpoint_delete_slide: { host: 'PowerPoint', run: powerPointDeleteSlide },
  powerpoint_move_slide: { host: 'PowerPoint', run: powerPointMoveSlide },
  outlook_draft_reply: { host: 'Outlook', run: outlookDraftReply },
  outlook_draft_new: { host: 'Outlook', run: outlookDraftNew },
  outlook_draft_meeting: { host: 'Outlook', run: outlookDraftMeeting },
};

async function execTool(tool, args) {
  if (!host) throw new Error('not running inside an Office application');
  const entry = Object.prototype.hasOwnProperty.call(TOOLS, tool) ? TOOLS[tool] : null;
  if (!entry) throw new Error('unknown tool ' + tool);
  if (entry.host && entry.host !== hostName()) throw new Error(`${tool} is not available in ${hostName()}`);
  return entry.run(args || {});
}
