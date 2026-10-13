// What every tool shares: the Office host the pane runs in, the helpers around
// Office.js, and the rules of reading at scale. Loaded first.

let host = null; // Office.HostType.* — null when opened outside Office (plain chat)

const HOST_LABELS = { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint', Outlook: 'Outlook' };
function hostName() {
  const T = Office.HostType;
  if (host === T.Word) return HOST_LABELS.Word;
  if (host === T.Excel) return HOST_LABELS.Excel;
  if (host === T.PowerPoint) return HOST_LABELS.PowerPoint;
  if (host === T.Outlook) return HOST_LABELS.Outlook;
  return null;
}

// Callback-style Office API → promise (rejects on a failed AsyncResult).
function officeAsync(call) {
  return new Promise((resolve, reject) => {
    call((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(new Error((res.error && res.error.message) || 'Office call failed'));
    });
  });
}

function rowsToText(values) {
  return values.map((row) => row.join('\t')).join('\n');
}

function documentTitle() {
  const url = (Office.context.document && Office.context.document.url) || '';
  // Local path or web URL → keep the file name only.
  const name = decodeURIComponent(url.split(/[\\/]/).pop().split('?')[0] || '');
  return name || hostName() + ' document';
}

// ── Reading at scale ─────────────────────────────────────────────────────────
// A client document can be hundreds of pages or hundreds of thousands of rows,
// so nothing here returns "the document": the agent gets its outline, then
// reads a slice or searches, the way one works on a large file with ls / head /
// grep. Every position handed back ("p:120", "Sales!A1:L50", "slide:40") is
// accepted as is by the next call.
const MAX_OUTPUT_CHARS = 12000; // content of one read
const MAX_RESULT_CHARS = 40000; // hard stop on any tool result
const CONTEXT_CHARS = 8000; // beginning of the document shared as context
const READ_BUDGET_MS = 15000;
const SEARCH_LIMIT = 20;
const MAX_SEARCH_LIMIT = 50;
const MAX_SEARCH = 255; // Word's search takes no more, and no line break
// Where a long Office call currently is — reported when a tool call times out.
let officeStep = '';
// Figures of the running tool call (round trips, fallbacks), for the debug mode.
let toolDiag = {};

const given = (value) => value !== undefined && value !== null;

// "p:12" / "p:12-40" → { from: 12, to: 40 | null }, 1-based. No target = from the start.
function parseSpan(target, prefix) {
  if (!given(target) || target === '') return { from: 1, to: null };
  const match = new RegExp(`^${prefix}:(\\d+)(?:-(\\d+))?$`).exec(String(target).trim());
  if (!match || +match[1] < 1 || (match[2] && +match[2] < +match[1])) {
    throw new Error(`expected a position such as "${prefix}:12" or "${prefix}:12-40", got "${target}"`);
  }
  return { from: +match[1], to: match[2] ? +match[2] : null };
}

// The text around a match, on one line.
function snippet(text, index, length) {
  const start = Math.max(0, index - 80);
  const end = Math.min(text.length, index + length + 80);
  return (start > 0 ? '…' : '') + text.slice(start, end).replace(/\s+/g, ' ') + (end < text.length ? '…' : '');
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
}
