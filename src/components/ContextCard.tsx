/**
 * components/ContextCard.tsx: who this is and what is going on with them.
 *
 * Shown once a participant has matched. Last activity, the user's open
 * follow-ups for the contact, and the company's open opportunities (pick one to
 * link the action to). If several people on the email matched, a switcher lets
 * the user change the primary. Everything here is read-only; the actions live
 * in ActionForms.
 */

import type { CrmData, Opportunity } from '../crm/types';
import type { MatchResult, ParticipantMatch } from '../crm/match';

interface Props {
  readonly data: CrmData;
  readonly match: MatchResult;
  readonly primary: ParticipantMatch;
  readonly onChoosePrimary: (m: ParticipantMatch) => void;
  readonly selectedOpportunityId: string | undefined;
  readonly onSelectOpportunity: (id: string | undefined) => void;
}

const money = (n: number) => n.toLocaleString(undefined, { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
const when = (iso: string) => new Date(iso).toLocaleDateString(undefined, { dateStyle: 'medium' });

export function ContextCard({ data, match, primary, onChoosePrimary, selectedOpportunityId, onSelectOpportunity }: Props) {
  const { contact, company, role } = primary;
  const activities = data.activities.filter(a => a.contactId === contact.id).sort((a, b) => b.occurredAt.localeCompare(a.occurredAt));
  const last = activities[0];
  const followUps = data.followUps.filter(f => f.contactId === contact.id && !f.done).sort((a, b) => a.dueOn.localeCompare(b.dueOn));
  const opportunities: Opportunity[] = data.opportunities.filter(o => o.companyId === company.id);
  const today = new Date().toISOString().slice(0, 10);

  return (
    <section className="card context" aria-labelledby="context-heading">
      <h2 id="context-heading">
        {contact.firstName} {contact.lastName}
        <span className="muted"> at {company.name}</span>
      </h2>
      <p className="muted small">
        {contact.title ? `${contact.title} · ` : ''}{contact.email}
        {' · '}matched from <strong>{role === 'from' ? 'sender' : role.toUpperCase()}</strong>
      </p>

      {match.all.length > 1 && (
        <div className="switcher">
          <span className="small">Also on this email: </span>
          {match.all.filter(m => m !== primary).map(m => (
            <button key={m.contact.id} type="button" className="link" onClick={() => onChoosePrimary(m)}>
              {m.contact.firstName} {m.contact.lastName} ({m.role === 'from' ? 'sender' : m.role.toUpperCase()})
            </button>
          ))}
        </div>
      )}

      <dl className="facts">
        <dt>Last activity</dt>
        <dd>{last ? <>{last.kind}: {last.subject} <span className="muted">({when(last.occurredAt)})</span></> : <span className="muted">none yet</span>}</dd>
        <dt>Open follow-ups</dt>
        <dd>
          {followUps.length === 0 ? <span className="muted">none</span> : (
            <ul className="plain">
              {followUps.map(f => (
                <li key={f.id} className={f.dueOn < today ? 'overdue' : ''}>
                  {f.subject} <span className="muted">(due {f.dueOn}{f.dueOn < today ? ', overdue' : ''})</span>
                </li>
              ))}
            </ul>
          )}
        </dd>
        <dt>Open opportunities at {company.name}</dt>
        <dd>
          {opportunities.length === 0 ? <span className="muted">none</span> : (
            <fieldset className="opps">
              <legend className="sr-only">Link this email to an opportunity</legend>
              <label className="opp">
                <input type="radio" name="opportunity" checked={!selectedOpportunityId} onChange={() => onSelectOpportunity(undefined)} />
                <span>Do not link to an opportunity</span>
              </label>
              {opportunities.map(o => (
                <label key={o.id} className="opp">
                  <input type="radio" name="opportunity" checked={selectedOpportunityId === o.id} onChange={() => onSelectOpportunity(o.id)} />
                  <span>{o.name} <span className="muted">({o.stage}, {money(o.amount)})</span></span>
                </label>
              ))}
            </fieldset>
          )}
        </dd>
      </dl>
    </section>
  );
}
