/**
 * components/ActionForms.tsx: Log email / Follow-up / Task, against the matched
 * contact (and the opportunity picked on the context card, if any).
 *
 * Each action is a small form that calls a pure builder from crm/actions.ts and
 * hands the record to the store. Before "Log email" runs, the duplicate report
 * for this item is shown; if anything already came from the same email the
 * button says so and the user chooses to proceed (warn-but-allow, see
 * actions.ts). Body capture is a checkbox, off by default, and fetches the
 * body only when ticked.
 */

import { useState } from 'react';
import { FOLLOW_UP_TIMEFRAMES, type CrmData } from '../crm/types';
import type { EmailContext, ParticipantMatch } from '../crm/match';
import { buildEmailActivity, buildFollowUp, buildTask, findDuplicates } from '../crm/actions';
import type { CrmStore } from '../crm/store';

interface Props {
  readonly store: CrmStore;
  readonly data: CrmData;
  readonly ctx: EmailContext;
  readonly target: ParticipantMatch;
  readonly opportunityId: string | undefined;
  /** Fetches the body on demand (Office or test mode); resolves '' if none. */
  readonly loadBody: () => Promise<string>;
  readonly onDone: (message: string) => void;
}

type Tab = 'log' | 'followup' | 'task';

export function ActionForms({ store, data, ctx, target, opportunityId, loadBody, onDone }: Props) {
  const [tab, setTab] = useState<Tab>('log');
  const dup = findDuplicates(data, ctx.itemId);

  return (
    <section className="card actions" aria-labelledby="actions-heading">
      <h2 id="actions-heading">Actions for this email</h2>
      <div role="tablist" aria-label="Action" className="tabs">
        {(['log', 'followup', 'task'] as const).map(t => (
          <button key={t} role="tab" type="button" aria-selected={tab === t} className={tab === t ? 'active' : ''} onClick={() => setTab(t)}>
            {t === 'log' ? 'Log email' : t === 'followup' ? 'Follow-up' : 'Task'}
          </button>
        ))}
      </div>

      {dup.any && (
        <p className="warn" role="status">
          Already logged from this email: {dup.activities.length} activit{dup.activities.length === 1 ? 'y' : 'ies'}, {dup.followUps.length} follow-up{dup.followUps.length === 1 ? '' : 's'}, {dup.tasks.length} task{dup.tasks.length === 1 ? '' : 's'}. You can still add more; this is a warning, not a block.
        </p>
      )}

      {tab === 'log' && <LogForm {...{ store, ctx, target, opportunityId, loadBody, onDone, warned: dup.any }} />}
      {tab === 'followup' && <FollowUpForm {...{ store, ctx, target, opportunityId, onDone }} />}
      {tab === 'task' && <TaskForm {...{ store, ctx, target, opportunityId, onDone }} />}
    </section>
  );
}

function LogForm({ store, ctx, target, opportunityId, loadBody, onDone, warned }: Omit<Props, 'data'> & { warned: boolean }) {
  const [note, setNote] = useState('');
  const [includeBody, setIncludeBody] = useState(false);
  const [busy, setBusy] = useState(false);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    setBusy(true);
    const body = includeBody ? await loadBody() : undefined;
    const activity = buildEmailActivity({ ctx: { ...ctx, body }, target, opportunityId, includeBody, note });
    store.addActivity(activity);
    setBusy(false);
    setNote('');
    onDone(`Logged "${activity.subject}" to ${target.contact.firstName} ${target.contact.lastName}.`);
  };

  return (
    <form onSubmit={submit} className="form">
      <p className="small">
        An activity of kind Email on <strong>{target.contact.firstName} {target.contact.lastName}</strong> and <strong>{target.company.name}</strong>
        {opportunityId ? ', linked to the selected opportunity' : ''}. Subject: <em>{ctx.subject || '(no subject)'}</em>.
      </p>
      <label>Note (optional)
        <textarea rows={2} value={note} onChange={e => setNote(e.target.value)} placeholder="What matters about this email" />
      </label>
      <label className="check">
        <input type="checkbox" checked={includeBody} onChange={e => setIncludeBody(e.target.checked)} />
        Include the email body in the note (capped at 1,500 characters)
      </label>
      <button type="submit" disabled={busy}>{warned ? 'Log anyway' : 'Log email'}</button>
    </form>
  );
}

function FollowUpForm({ store, ctx, target, opportunityId, onDone }: Pick<Props, 'store' | 'ctx' | 'target' | 'opportunityId' | 'onDone'>) {
  const [timeframeId, setTimeframeId] = useState(FOLLOW_UP_TIMEFRAMES[2].id);
  const [subject, setSubject] = useState('');

  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    const f = buildFollowUp({ ctx, target, opportunityId, timeframeId, subject });
    store.addFollowUp(f);
    setSubject('');
    onDone(`Follow-up "${f.subject}" due ${f.dueOn}.`);
  };

  return (
    <form onSubmit={submit} className="form">
      <label>When
        <select value={timeframeId} onChange={e => setTimeframeId(e.target.value)}>
          {FOLLOW_UP_TIMEFRAMES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
        </select>
      </label>
      <label>Subject (optional; defaults to "Follow up: {ctx.subject || '(no subject)'}")
        <input value={subject} onChange={e => setSubject(e.target.value)} />
      </label>
      <button type="submit">Create follow-up</button>
    </form>
  );
}

function TaskForm({ store, ctx, target, opportunityId, onDone }: Pick<Props, 'store' | 'ctx' | 'target' | 'opportunityId' | 'onDone'>) {
  const [subject, setSubject] = useState(ctx.subject);
  const [dueOn, setDueOn] = useState('');

  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    const t = buildTask({ ctx, target, opportunityId, subject, dueOn });
    store.addTask(t);
    onDone(`Task "${t.subject}"${t.dueOn ? ` due ${t.dueOn}` : ''} created.`);
  };

  return (
    <form onSubmit={submit} className="form">
      <label>Subject
        <input value={subject} onChange={e => setSubject(e.target.value)} />
      </label>
      <label>Due date (optional)
        <input type="date" value={dueOn} onChange={e => setDueOn(e.target.value)} />
      </label>
      <button type="submit">Create task</button>
    </form>
  );
}
