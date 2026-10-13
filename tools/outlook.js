// Outlook: the open email, read in parts and searched — then the unsent
// drafts (reply, new email, meeting) opened in Outlook's own forms.

// Mail fields are plain values in read mode and async getters in compose mode.
function mailField(field) {
  if (field && typeof field.getAsync === 'function') {
    return officeAsync((cb) => field.getAsync(cb)).catch(() => undefined);
  }
  return Promise.resolve(field);
}

function formatAddress(a) {
  if (!a) return '';
  return a.displayName && a.emailAddress ? `${a.displayName} <${a.emailAddress}>` : a.emailAddress || a.displayName || '';
}

async function readMail() {
  const item = Office.context.mailbox.item;
  if (!item) return null;
  const [subject, from, to, body] = await Promise.all([
    mailField(item.subject),
    mailField(item.from),
    mailField(item.to),
    officeAsync((cb) => item.body.getAsync(Office.CoercionType.Text, cb)).catch(() => ''),
  ]);
  return {
    subject: subject || '',
    from: formatAddress(from),
    to: (to || []).map(formatAddress).filter(Boolean),
    body: body || '',
  };
}

async function openMail() {
  const mail = await readMail();
  if (!mail) throw new Error('no email is open');
  return mail;
}

async function mailRead({ target }) {
  const mail = await openMail();
  const span = parseSpan(target, 'c'); // character positions in the body
  const stop = Math.min(mail.body.length, span.to || Infinity);
  const end = Math.min(stop, span.from - 1 + MAX_OUTPUT_CHARS);
  return {
    subject: mail.subject,
    from: mail.from,
    to: mail.to,
    body: mail.body.slice(span.from - 1, end),
    next: end < stop ? `c:${end + 1}${span.to ? `-${span.to}` : ''}` : undefined,
  };
}

async function mailSearch({ query, limit }) {
  const { body } = await openMail();
  const haystack = body.toLowerCase();
  const needle = query.toLowerCase();
  const results = [];
  let matches = 0;
  for (let i = haystack.indexOf(needle); i >= 0; i = haystack.indexOf(needle, i + needle.length)) {
    matches++;
    if (results.length < limit) results.push({ at: `c:${i + 1}`, text: snippet(body, i, needle.length) });
  }
  return { query, matches, results };
}

// One address or several → a list of strings.
const addressList = (value) => (Array.isArray(value) ? value : value ? [value] : []).map(String);

async function outlookDraftNew({ to, cc, subject, body }) {
  if (typeof body !== 'string') throw new Error('body is required');
  const mailbox = Office.context.mailbox;
  if (!mailbox || typeof mailbox.displayNewMessageForm !== 'function') {
    throw new Error('this version of Outlook cannot open a new message from the add-in');
  }
  // Opens Outlook's own compose form, unsent — the user reviews and sends.
  mailbox.displayNewMessageForm({
    toRecipients: addressList(to),
    ccRecipients: addressList(cc),
    subject: String(subject || ''),
    htmlBody: escapeHtml(body),
  });
  return { ok: true, sent: false };
}

// "2026-10-14T10:00" (the user's local time) or any ISO 8601 date → Date.
function meetingDate(value, label) {
  const date = new Date(value);
  if (typeof value !== 'string' || Number.isNaN(date.getTime())) {
    throw new Error(`${label} must be a date and time such as "2026-10-14T10:00"`);
  }
  return date;
}

async function outlookDraftMeeting({ attendees, optional, subject, body, start, end, location }) {
  const mailbox = Office.context.mailbox;
  if (!mailbox || typeof mailbox.displayNewAppointmentForm !== 'function') {
    throw new Error('this version of Outlook cannot open a new meeting from the add-in');
  }
  const form = {
    requiredAttendees: addressList(attendees),
    optionalAttendees: addressList(optional),
    subject: String(subject || ''),
    body: String(body || ''),
    location: String(location || ''),
  };
  // Without a start, Outlook proposes its own next free slot.
  if (given(start)) {
    form.start = meetingDate(start, 'start');
    form.end = given(end) ? meetingDate(end, 'end') : new Date(form.start.getTime() + 30 * 60000);
    if (form.end <= form.start) throw new Error('end must be after start');
  } else if (given(end)) {
    throw new Error('end needs a start');
  }
  // Opens Outlook's own meeting form, unsent — the user reviews and sends.
  mailbox.displayNewAppointmentForm(form);
  return { ok: true, sent: false };
}

async function outlookDraftReply({ body, replyAll }) {
  if (typeof body !== 'string') throw new Error('body is required');
  const item = Office.context.mailbox.item;
  if (!item || typeof item.displayReplyForm !== 'function') {
    throw new Error('no message open in read mode');
  }
  // Opens Outlook's own reply form, unsent — the user reviews and sends.
  if (replyAll) item.displayReplyAllForm(escapeHtml(body));
  else item.displayReplyForm(escapeHtml(body));
  return { ok: true, sent: false };
}
