/**
 * components/CreateContact.tsx: "this person is not in the CRM; add them?"
 *
 * Shown when no participant matched. Pre-filled from the address (name split
 * into first/last, or derived from the local part) and, when the sender's
 * domain belongs to a known company, that company is pre-selected: the most
 * common real case is a new person at an existing customer. A new company can
 * be typed instead.
 *
 * The internal version ran the same server-side validation and duplicate rules
 * as the web app here. In the demo the duplicate rule is local: an address
 * already in the CRM is refused with a message (that one is a block, not a
 * warning, because two contacts with one address is a data error, not a
 * judgement call).
 */

import { useState } from 'react';
import type { Company, CrmData } from '../crm/types';
import { normalizeEmail, splitName, type EmailAddress, type MatchResult } from '../crm/match';
import type { CrmStore } from '../crm/store';
import { newId } from '../crm/actions';

interface Props {
  readonly store: CrmStore;
  readonly data: CrmData;
  readonly match: MatchResult;
  readonly candidate: EmailAddress;
  readonly onCreated: (message: string) => void;
}

export function CreateContact({ store, data, match, candidate, onCreated }: Props) {
  const guess = splitName(candidate.name, candidate.email);
  const [firstName, setFirstName] = useState(guess.firstName);
  const [lastName, setLastName] = useState(guess.lastName);
  const [title, setTitle] = useState('');
  const [companyId, setCompanyId] = useState<string>(match.domainCompany?.id ?? '');
  const [newCompany, setNewCompany] = useState('');
  const [error, setError] = useState<string | null>(null);

  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    const email = normalizeEmail(candidate.email);
    if (data.contacts.some(c => normalizeEmail(c.email) === email)) {
      setError(`A contact with ${email} already exists.`);
      return;
    }
    if (!firstName.trim()) { setError('First name is required.'); return; }
    let company: Company | undefined = data.companies.find(c => c.id === companyId);
    if (!company) {
      if (!newCompany.trim()) { setError('Choose a company or type a new one.'); return; }
      // The real add-in had no inline "create company"; it is a demo-only
      // convenience so the stranger case can complete. The domain comes from
      // the address, which is also how domain matching will find the next
      // person from the same place.
      company = { id: newId('co'), name: newCompany.trim(), domain: email.split('@')[1] ?? '', industry: 'Unknown' };
      store.addCompany(company);
    }
    store.addContact({
      id: newId('c'), firstName: firstName.trim(), lastName: lastName.trim(), email,
      companyId: company.id, ...(title.trim() ? { title: title.trim() } : {}), source: 'Email',
    });
    onCreated(`Created ${firstName} ${lastName} at ${company.name}.`);
  };

  return (
    <section className="card create" aria-labelledby="create-heading">
      <h2 id="create-heading">Not in the CRM yet</h2>
      <p className="small">
        <strong>{candidate.name || candidate.email}</strong>{candidate.name ? ` <${candidate.email}>` : ''} did not match a contact
        {match.domainCompany ? <>, but the domain belongs to <strong>{match.domainCompany.name}</strong>, so it is probably someone new there.</> : '.'}
        {' '}Add them and then log the email.
      </p>
      <form onSubmit={submit} className="form">
        <div className="row">
          <label>First name <input value={firstName} onChange={e => setFirstName(e.target.value)} required /></label>
          <label>Last name <input value={lastName} onChange={e => setLastName(e.target.value)} /></label>
        </div>
        <label>Title (optional) <input value={title} onChange={e => setTitle(e.target.value)} /></label>
        <label>Company
          <select value={companyId} onChange={e => setCompanyId(e.target.value)}>
            <option value="">New company (type below)</option>
            {data.companies.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </label>
        {!companyId && (
          <label>New company name <input value={newCompany} onChange={e => setNewCompany(e.target.value)} /></label>
        )}
        {error && <p className="error" role="alert">{error}</p>}
        <button type="submit">Create contact</button>
      </form>
    </section>
  );
}
