/**
 * office/host.ts: the only file that talks to Office.js.
 *
 * Everything else in the pane works from an `EmailContext` (see crm/match.ts),
 * which this module produces either from the real open item or, when there is
 * no Outlook host, from whatever the user types into the test-mode bar. That
 * one seam is what makes the add-in runnable and testable in a plain browser,
 * and it is the reason the same page on GitHub Pages is both the taskpane and
 * the demo.
 *
 * Host detection: `Office.onReady()` resolves with `{ host: 'Outlook' }` inside
 * Outlook and `{ host: null }` (or never exists, if the CDN script failed)
 * anywhere else. We wait for it with a timeout so a slow CDN cannot hang the
 * pane forever; after the timeout we assume no host and fall back to test mode.
 *
 * Reading the item: read mode gives plain values; compose mode gives async
 * getters. `readItem` normalises both into one promise. The body is fetched
 * only on request (`readBody`), because it is the expensive call and the pane
 * only needs it when the user opts in to capture it.
 */

import type { EmailAddress, EmailContext } from '../crm/match';

export type HostKind = 'outlook' | 'none';

/**
 * Watch for the host, reporting possibly TWICE.
 *
 * Office.onReady inside Outlook can take longer than any reasonable spinner
 * (the host bootstraps its scripts on first load), and a plain timeout race
 * gets it wrong in the worst way: the pane falls into browser test mode and
 * STAYS there even though it is sitting inside Outlook. That is exactly the
 * bug this replaces. So: after `timeoutMs` with no answer, report 'none' so
 * the user is not staring at a spinner, but KEEP LISTENING; if onReady later
 * resolves with a host, report 'outlook' and the app upgrades in place.
 *
 * `onChange` therefore fires once ('none' or 'outlook') or twice ('none' then
 * 'outlook'). It never downgrades.
 */
export function watchHost(onChange: (kind: HostKind) => void, timeoutMs = 4000): void {
  const office = getOffice();
  if (!office || typeof office.onReady !== 'function') { onChange('none'); return; }
  let reported: HostKind | null = null;
  const report = (kind: HostKind) => {
    if (reported === 'outlook') return;          // never downgrade
    if (reported === kind) return;               // no duplicate reports
    reported = kind;
    onChange(kind);
  };
  const timer = setTimeout(() => report('none'), timeoutMs);
  try {
    office.onReady(info => { clearTimeout(timer); report(info?.host ? 'outlook' : 'none'); })
      .catch?.(() => { clearTimeout(timer); report('none'); });
  } catch { clearTimeout(timer); report('none'); }
}

/** One-shot form of watchHost, kept for tests and simple callers. */
export function detectHost(timeoutMs = 3000): Promise<HostKind> {
  return new Promise(resolve => {
    let settled = false;
    watchHost(kind => { if (!settled) { settled = true; resolve(kind); } }, timeoutMs);
  });
}

/** The global, if any. Wrapped so tests can install a stub on window. */
export function getOffice(): OfficeNamespace | undefined {
  return (globalThis as unknown as { Office?: OfficeNamespace }).Office;
}

function toAddress(a: OfficeEmailAddressDetails | undefined | null): EmailAddress | undefined {
  if (!a || !a.emailAddress) return undefined;
  return { name: a.displayName ?? '', email: a.emailAddress };
}

/** Read-mode arrays vs compose-mode getters, normalised to a promise of addresses. */
function readRecipients(field: OfficeEmailAddressDetails[] | OfficeRecipients | undefined): Promise<EmailAddress[]> {
  if (!field) return Promise.resolve([]);
  if (Array.isArray(field)) return Promise.resolve(field.map(toAddress).filter((x): x is EmailAddress => !!x));
  return new Promise(resolve => {
    try {
      field.getAsync(r => resolve(r.status === 'succeeded' && r.value ? r.value.map(toAddress).filter((x): x is EmailAddress => !!x) : []));
    } catch { resolve([]); }
  });
}

function readSubject(field: string | OfficeComposeSubject | undefined): Promise<string> {
  if (field === undefined || field === null) return Promise.resolve('');
  if (typeof field === 'string') return Promise.resolve(field);
  return new Promise(resolve => {
    try { field.getAsync(r => resolve(r.status === 'succeeded' && r.value ? r.value : '')); } catch { resolve(''); }
  });
}

/**
 * Build an EmailContext from the currently open Outlook item, or null if there
 * is no item (the pane opened with nothing selected).
 */
export async function readItem(): Promise<EmailContext | null> {
  const item = getOffice()?.context?.mailbox?.item;
  if (!item) return null;
  // Compose mode is recognisable by the async getters; read mode has plain values.
  const isCompose = !!item.to && !Array.isArray(item.to);
  const [subject, to, cc] = await Promise.all([readSubject(item.subject), readRecipients(item.to), readRecipients(item.cc)]);
  const from = toAddress(item.from) ?? toAddress(item.sender);
  const itemId = item.internetMessageId || item.itemId || undefined;
  return {
    mode: isCompose ? 'compose' : 'read',
    ...(itemId ? { itemId } : {}),
    subject,
    ...(from && !isCompose ? { from } : {}),
    to,
    cc,
  };
}

/** Fetch the plain-text body on demand. Resolves '' on any failure. */
export function readBody(): Promise<string> {
  const item = getOffice()?.context?.mailbox?.item;
  if (!item?.body) return Promise.resolve('');
  return new Promise(resolve => {
    try { item.body!.getAsync('text', r => resolve(r.status === 'succeeded' && r.value ? r.value : '')); } catch { resolve(''); }
  });
}

/** The signed-in user's address, so matching can skip "me". */
export function ownAddress(): string | undefined {
  return getOffice()?.context?.mailbox?.userProfile?.emailAddress || undefined;
}

/**
 * Subscribe to the pinned-pane "item changed" event. Returns an unsubscribe
 * that is a no-op if the host does not support it. A pinned taskpane stays
 * open while the user moves between emails; without this it would keep
 * showing the first one.
 */
export function onItemChanged(handler: () => void): () => void {
  const office = getOffice();
  const mailbox = office?.context?.mailbox;
  const eventType = office?.EventType?.ItemChanged;
  if (!mailbox?.addHandlerAsync || !eventType) return () => {};
  try { mailbox.addHandlerAsync(eventType, handler); } catch { /* unsupported host */ }
  // Office has no removeHandler in this slice; the pane lives for the session.
  return () => {};
}
