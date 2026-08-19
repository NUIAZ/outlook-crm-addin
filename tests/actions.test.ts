/**
 * tests/actions.test.ts: the record builders and the duplicate rule.
 */
import { describe, expect, it } from 'vitest';
import { buildSeed } from '../src/crm/seed';
import { matchEmail, type EmailContext } from '../src/crm/match';
import { BODY_CAPTURE_LIMIT, buildEmailActivity, buildFollowUp, buildTask, clipBody, dueDateFor, findDuplicates } from '../src/crm/actions';
import { createStore } from '../src/crm/store';

const NOW = new Date('2026-08-19T15:30:00Z');
const data = buildSeed(NOW);
const me = 'you@yourcompany.example';
const email: EmailContext = {
  mode: 'read', itemId: '<abc123@mail.example>', subject: 'Re: Depot telematics rollout',
  from: { name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' },
  to: [{ name: 'You', email: me }], cc: [],
  body: 'Thanks for the call.\n\nCould you   send the revised quote?',
};
const target = matchEmail(data, email, [me]).primary!;

describe('buildEmailActivity', () => {
  it('links contact, company and the chosen opportunity, and stamps the item id', () => {
    const a = buildEmailActivity({ ctx: email, target, opportunityId: 'o-1', includeBody: false }, NOW, 'a-test');
    expect(a).toMatchObject({ id: 'a-test', contactId: 'c-1', companyId: 'co-bluefin', opportunityId: 'o-1', kind: 'Email', subject: 'Re: Depot telematics rollout', itemId: '<abc123@mail.example>' });
    expect(a.occurredAt).toBe(NOW.toISOString());
    expect(a.notes).toContain('Received email. From Priya Raman.');
    expect(a.notes).not.toContain('Email body');
  });

  it('omits opportunityId and itemId keys when absent (clean payloads)', () => {
    const a = buildEmailActivity({ ctx: { ...email, itemId: undefined }, target, includeBody: false }, NOW, 'x');
    expect('opportunityId' in a).toBe(false);
    expect('itemId' in a).toBe(false);
  });

  it('captures the body only when asked, collapsed and capped', () => {
    const a = buildEmailActivity({ ctx: email, target, includeBody: true, note: 'board prep' }, NOW, 'x');
    expect(a.notes).toContain('board prep');
    expect(a.notes).toContain('Email body: Thanks for the call. Could you send the revised quote?');
    const long = 'x'.repeat(BODY_CAPTURE_LIMIT + 500);
    expect(clipBody(long).length).toBe(BODY_CAPTURE_LIMIT);
    expect(clipBody(long).endsWith('…')).toBe(true);
  });

  it('describes a sent email by its recipients in compose mode', () => {
    const draft: EmailContext = { mode: 'compose', subject: 'Quote', to: [{ name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' }], cc: [] };
    const t = matchEmail(data, draft, [me]).primary!;
    const a = buildEmailActivity({ ctx: draft, target: t, includeBody: false }, NOW, 'x');
    expect(a.notes).toContain('Sent email. To Priya Raman.');
  });

  it('falls back to "(no subject)"', () => {
    const a = buildEmailActivity({ ctx: { ...email, subject: '' }, target, includeBody: false }, NOW, 'x');
    expect(a.subject).toBe('(no subject)');
  });
});

describe('follow-ups and tasks', () => {
  it('computes due dates from the admin timeframes', () => {
    expect(dueDateFor('tomorrow', NOW)).toBe('2026-08-20');
    expect(dueDateFor('1w', NOW)).toBe('2026-08-26');
    expect(dueDateFor('bogus', NOW)).toBe('2026-08-26'); // unknown -> 1 week
  });
  it('builds a follow-up with a default subject and the item id', () => {
    const f = buildFollowUp({ ctx: email, target, timeframeId: '3d' }, NOW, 'f-x');
    expect(f).toMatchObject({ id: 'f-x', contactId: 'c-1', subject: 'Follow up: Re: Depot telematics rollout', dueOn: '2026-08-22', done: false, itemId: '<abc123@mail.example>' });
  });
  it('builds a task with the email subject prefilled and optional due date', () => {
    expect(buildTask({ ctx: email, target }, NOW, 't-x').subject).toBe('Re: Depot telematics rollout');
    expect('dueOn' in buildTask({ ctx: email, target }, NOW, 't-x')).toBe(false);
    expect(buildTask({ ctx: email, target, subject: 'Send deck', dueOn: '2026-09-01' }, NOW, 't-x')).toMatchObject({ subject: 'Send deck', dueOn: '2026-09-01' });
  });
});

describe('duplicate rule: warn, never block', () => {
  it('reports nothing for a fresh item, and nothing without an item id', () => {
    expect(findDuplicates(data, '<never-seen@x>').any).toBe(false);
    expect(findDuplicates(data, undefined).any).toBe(false);
  });
  it('reports every record that came from the same item, across all three kinds', () => {
    const store = createStore(null, () => NOW);
    store.addActivity(buildEmailActivity({ ctx: email, target, includeBody: false }, NOW, 'a1'));
    store.addFollowUp(buildFollowUp({ ctx: email, target, timeframeId: '1w' }, NOW, 'f1'));
    store.addTask(buildTask({ ctx: email, target }, NOW, 't1'));
    const d = findDuplicates(store.getData(), email.itemId);
    expect(d.any).toBe(true);
    expect(d.activities.map(a => a.id)).toEqual(['a1']);
    expect(d.followUps.map(f => f.id)).toEqual(['f1']);
    expect(d.tasks.map(t => t.id)).toEqual(['t1']);
  });
  it('logging the same email twice is allowed; both records exist afterwards', () => {
    const store = createStore(null, () => NOW);
    store.addActivity(buildEmailActivity({ ctx: email, target, includeBody: false }, NOW, 'a1'));
    store.addActivity(buildEmailActivity({ ctx: email, target, includeBody: false }, NOW, 'a2'));
    expect(findDuplicates(store.getData(), email.itemId).activities).toHaveLength(2);
  });
});
