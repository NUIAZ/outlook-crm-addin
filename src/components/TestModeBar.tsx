/**
 * components/TestModeBar.tsx: the controls shown when there is no Outlook host.
 *
 * Outside Outlook the pane has no open email, so this bar lets you describe one:
 * pick a sample, or type a sender, recipients and subject by hand, and choose
 * read or compose mode. The result is a normal EmailContext; nothing below the
 * bar knows it came from a form rather than from Office.js. That is the point:
 * the whole add-in is exercisable in a browser tab, which is how the hosted
 * demo works and how the tests drive the UI.
 */

import { useState } from 'react';
import type { EmailAddress, EmailContext } from '../crm/match';
import { SAMPLE_EMAILS } from '../crm/seed';

interface Props {
  readonly value: EmailContext;
  readonly onChange: (ctx: EmailContext) => void;
}

/** "Name <addr>, addr2" -> addresses. Tolerant: bare addresses, commas, semicolons. */
export function parseAddressList(text: string): EmailAddress[] {
  return text
    .split(/[;,]/)
    .map(s => s.trim())
    .filter(Boolean)
    .map(s => {
      const m = /^(.*?)\s*<([^>]+)>$/.exec(s);
      return m ? { name: m[1].replace(/"/g, '').trim(), email: m[2].trim() } : { name: '', email: s };
    });
}

function formatAddressList(list: readonly EmailAddress[]): string {
  return list.map(a => (a.name ? `${a.name} <${a.email}>` : a.email)).join(', ');
}

export function sampleToContext(sampleId: string): EmailContext {
  const s = SAMPLE_EMAILS.find(x => x.id === sampleId) ?? SAMPLE_EMAILS[0];
  return {
    mode: 'read',
    itemId: `<${s.id}@sample.example>`,
    subject: s.subject,
    from: s.from,
    to: [...s.to],
    cc: [...s.cc],
    body: s.body,
  };
}

export function TestModeBar({ value, onChange }: Props) {
  const [fromText, setFromText] = useState(value.from ? formatAddressList([value.from]) : '');
  const [toText, setToText] = useState(formatAddressList(value.to));
  const [ccText, setCcText] = useState(formatAddressList(value.cc));
  const [subject, setSubject] = useState(value.subject);
  const [mode, setMode] = useState<'read' | 'compose'>(value.mode);
  const [body, setBody] = useState(value.body ?? '');

  const apply = (next: Partial<{ fromText: string; toText: string; ccText: string; subject: string; mode: 'read' | 'compose'; body: string }>) => {
    const f = next.fromText ?? fromText;
    const t = next.toText ?? toText;
    const c = next.ccText ?? ccText;
    const sub = next.subject ?? subject;
    const m = next.mode ?? mode;
    const b = next.body ?? body;
    const from = parseAddressList(f)[0];
    onChange({
      mode: m,
      // A typed email gets a synthetic, stable id from its subject + sender so the
      // duplicate check still works in test mode (log it twice, get the warning).
      itemId: `<test-${hash(`${f}|${sub}`)}@test.example>`,
      subject: sub,
      ...(from && m === 'read' ? { from } : {}),
      to: parseAddressList(t),
      cc: parseAddressList(c),
      body: b,
    });
  };

  const loadSample = (id: string) => {
    const ctx = sampleToContext(id);
    setFromText(ctx.from ? formatAddressList([ctx.from]) : '');
    setToText(formatAddressList(ctx.to));
    setCcText(formatAddressList(ctx.cc));
    setSubject(ctx.subject);
    setMode('read');
    setBody(ctx.body ?? '');
    onChange(ctx);
  };

  return (
    <section className="testbar" aria-labelledby="testbar-heading">
      <h2 id="testbar-heading">Test mode: no Outlook host detected</h2>
      <p className="muted">
        Describe an email and the pane behaves as if it were open in Outlook. Pick a sample or type your own.
      </p>
      <div className="row">
        <label>
          Sample email
          <select onChange={e => loadSample(e.target.value)} defaultValue="">
            <option value="" disabled>Choose…</option>
            {SAMPLE_EMAILS.map(s => <option key={s.id} value={s.id}>{s.label}</option>)}
          </select>
        </label>
        <fieldset className="mode">
          <legend>Mode</legend>
          <label><input type="radio" name="mode" checked={mode === 'read'} onChange={() => { setMode('read'); apply({ mode: 'read' }); }} /> Reading an email</label>
          <label><input type="radio" name="mode" checked={mode === 'compose'} onChange={() => { setMode('compose'); apply({ mode: 'compose' }); }} /> Writing one (compose)</label>
        </fieldset>
      </div>
      {mode === 'read' && (
        <label>From
          <input value={fromText} onChange={e => { setFromText(e.target.value); apply({ fromText: e.target.value }); }} placeholder="Name <someone@company.example>" />
        </label>
      )}
      <label>To
        <input value={toText} onChange={e => { setToText(e.target.value); apply({ toText: e.target.value }); }} placeholder="addresses, comma separated" />
      </label>
      <label>CC
        <input value={ccText} onChange={e => { setCcText(e.target.value); apply({ ccText: e.target.value }); }} placeholder="optional" />
      </label>
      <label>Subject
        <input value={subject} onChange={e => { setSubject(e.target.value); apply({ subject: e.target.value }); }} />
      </label>
      <label>Body (only used if you tick "include the email body")
        <textarea rows={3} value={body} onChange={e => { setBody(e.target.value); apply({ body: e.target.value }); }} />
      </label>
    </section>
  );
}

/** Tiny string hash for synthetic ids; not cryptographic, just stable. */
function hash(s: string): string {
  let h = 2166136261;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
  return (h >>> 0).toString(36);
}
