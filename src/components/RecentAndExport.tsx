/**
 * components/RecentAndExport.tsx: what the pane has written, and where it went.
 *
 * This is the honest bottom of the pane. "Recently added" lists every record
 * the add-in created (those carrying an Outlook item id, plus contacts created
 * from the pane). "Export" shows the same records as the JSON payloads the real
 * add-in would have POSTed to the CRM API, endpoint by endpoint, so a reader
 * can see the transport boundary: everything above this component is
 * production logic; the only thing the demo replaces is the function that
 * receives these payloads. "Reset demo data" restores the seed.
 */

import { useState } from 'react';
import type { CrmData } from '../crm/types';
import type { CrmStore } from '../crm/store';

interface Props {
  readonly store: CrmStore;
  readonly data: CrmData;
  readonly onReset: () => void;
}

export function RecentAndExport({ store, data, onReset }: Props) {
  const [open, setOpen] = useState(false);
  const [showJson, setShowJson] = useState(false);
  const added = store.addedByAddIn();
  const count = added.activities.length + added.followUps.length + added.tasks.length + added.contacts.length;
  const nameOf = (contactId: string) => {
    const c = data.contacts.find(x => x.id === contactId);
    return c ? `${c.firstName} ${c.lastName}` : contactId;
  };

  const payloads = [
    ...added.contacts.map(c => ({ endpoint: 'POST /api/contacts', body: c })),
    ...added.activities.map(a => ({ endpoint: 'POST /api/activities', body: a })),
    ...added.followUps.map(f => ({ endpoint: 'POST /api/followups', body: f })),
    ...added.tasks.map(t => ({ endpoint: 'POST /api/tasks', body: t })),
  ];

  return (
    <section className="card recent" aria-labelledby="recent-heading">
      <h2 id="recent-heading">
        <button type="button" className="disclosure" aria-expanded={open} onClick={() => setOpen(o => !o)}>
          {open ? '▾' : '▸'} Recently added by this pane ({count})
        </button>
      </h2>
      {open && (
        <>
          {count === 0 ? <p className="muted small">Nothing yet. Log an email, add a follow-up or a task, or create a contact, and it will appear here.</p> : (
            <ul className="plain small">
              {added.contacts.map(c => <li key={c.id}>Contact: {c.firstName} {c.lastName} ({c.email})</li>)}
              {added.activities.map(a => <li key={a.id}>Activity: {a.subject} on {nameOf(a.contactId)}</li>)}
              {added.followUps.map(f => <li key={f.id}>Follow-up: {f.subject} on {nameOf(f.contactId)}, due {f.dueOn}</li>)}
              {added.tasks.map(t => <li key={t.id}>Task: {t.subject} on {nameOf(t.contactId)}{t.dueOn ? `, due ${t.dueOn}` : ''}</li>)}
            </ul>
          )}
          <p className="small muted">
            Where this went: this host's localStorage, under <code>outlook-crm-addin:v1</code>. Per device, per host, visible to nobody else. The real add-in would have posted each record to the CRM API instead:
          </p>
          <div className="row">
            <button type="button" onClick={() => setShowJson(s => !s)} disabled={count === 0}>{showJson ? 'Hide' : 'Show'} what would have been sent</button>
            <button type="button" className="danger" onClick={onReset}>Reset demo data</button>
          </div>
          {showJson && count > 0 && (
            <pre className="export" aria-label="Payloads the real add-in would have sent">{JSON.stringify(payloads, null, 2)}</pre>
          )}
        </>
      )}
    </section>
  );
}
