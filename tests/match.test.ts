/**
 * tests/match.test.ts: who on the email is in the CRM?
 *
 * These are the rules a salesperson would describe out loud: "it should find
 * the sender; if the sender is an assistant and the buyer is CC'd, it should
 * find the buyer too; if it is someone new at a customer, it should at least
 * know the company". Each rule is a test. All pure, no DOM.
 */
import { describe, expect, it } from 'vitest';
import { buildSeed } from '../src/crm/seed';
import { domainOf, matchEmail, normalizeEmail, participantsInOrder, splitName, type EmailContext } from '../src/crm/match';

const data = buildSeed(new Date('2026-08-19T12:00:00Z'));
const me = 'you@yourcompany.example';
const ctx = (partial: Partial<EmailContext>): EmailContext => ({ mode: 'read', subject: 'Hello', to: [{ name: 'You', email: me }], cc: [], ...partial });

describe('normalisation helpers', () => {
  it('lower-cases and trims addresses', () => {
    expect(normalizeEmail('  Priya.Raman@Bluefin-Logistics.EXAMPLE ')).toBe('priya.raman@bluefin-logistics.example');
  });
  it('extracts the domain, or empty for junk', () => {
    expect(domainOf('a@B.Example')).toBe('b.example');
    expect(domainOf('nonsense')).toBe('');
  });
});

describe('read mode: sender first', () => {
  it('matches a known sender as primary, from the sender field', () => {
    const r = matchEmail(data, ctx({ from: { name: 'Priya Raman', email: 'PRIYA.RAMAN@bluefin-logistics.example' } }), [me]);
    expect(r.primary?.contact.id).toBe('c-1');
    expect(r.primary?.role).toBe('from');
    expect(r.primary?.company.name).toBe('Bluefin Logistics');
    expect(r.unknownCandidate).toBeNull();
  });

  it('skips my own address even when it is in To', () => {
    const r = matchEmail(data, ctx({ from: { name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' } }), [me]);
    expect(r.all.map(m => m.address.email)).not.toContain(me);
  });

  it('assistant sends, buyer on CC: sender is primary, buyer is also listed', () => {
    const r = matchEmail(data, ctx({
      from: { name: 'Hannah Okafor', email: 'hannah.okafor@bluefin-logistics.example' },
      cc: [{ name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' }],
    }), [me]);
    expect(r.primary?.contact.id).toBe('c-3');
    expect(r.all.map(m => m.contact.id)).toEqual(['c-3', 'c-1']);
    expect(r.all[1].role).toBe('cc');
  });

  it('unknown sender, known buyer on CC: the buyer becomes primary and the sender is the create candidate', () => {
    const r = matchEmail(data, ctx({
      from: { name: 'New Assistant', email: 'new.assistant@bluefin-logistics.example' },
      cc: [{ name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' }],
    }), [me]);
    expect(r.primary?.contact.id).toBe('c-1');
    expect(r.unknownCandidate?.email).toBe('new.assistant@bluefin-logistics.example');
    // Someone matched, so the domain guess is not surfaced as the headline.
    expect(r.domainCompany).toBeNull();
  });

  it('unknown person at a known company: nobody matches, but the company is guessed from the domain', () => {
    const r = matchEmail(data, ctx({ from: { name: 'Jordan Ellis', email: 'jordan.ellis@granitepeak.example' } }), [me]);
    expect(r.primary).toBeNull();
    expect(r.all).toHaveLength(0);
    expect(r.domainCompany?.name).toBe('Granite Peak Outfitters');
    expect(r.unknownCandidate?.email).toBe('jordan.ellis@granitepeak.example');
  });

  it('complete stranger: no match, no company guess, still a create candidate', () => {
    const r = matchEmail(data, ctx({ from: { name: 'Sam', email: 'sam@northbend-plumbing.example' } }), [me]);
    expect(r.primary).toBeNull();
    expect(r.domainCompany).toBeNull();
    expect(r.unknownCandidate?.email).toBe('sam@northbend-plumbing.example');
  });

  it('does not list the same address twice if it appears in two fields', () => {
    const p = { name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' };
    const r = matchEmail(data, ctx({ from: p, cc: [p] }), [me]);
    expect(r.all).toHaveLength(1);
  });
});

describe('compose mode: recipients, not sender', () => {
  it('ignores any from and matches To then CC in order', () => {
    const r = matchEmail(data, {
      mode: 'compose', subject: 'Draft',
      from: { name: 'Me', email: me },
      to: [{ name: 'Wes', email: 'wes.thorne@granitepeak.example' }],
      cc: [{ name: 'Maya', email: 'maya.lindqvist@granitepeak.example' }],
    }, [me]);
    expect(r.all.map(m => m.contact.id)).toEqual(['c-4', 'c-5']);
    expect(r.primary?.role).toBe('to');
  });

  it('participantsInOrder omits the sender in compose mode', () => {
    const order = participantsInOrder({ mode: 'compose', subject: '', from: { name: 'Me', email: me }, to: [{ name: 'A', email: 'a@x.example' }], cc: [] });
    expect(order.map(p => p.role)).toEqual(['to']);
  });
});

describe('splitName', () => {
  it('splits first and last, keeping honorifics with the first name', () => {
    expect(splitName('Dr. Anita Shah', 'x@y.example')).toEqual({ firstName: 'Dr. Anita', lastName: 'Shah' });
  });
  it('handles a single token and strips quotes and angle brackets', () => {
    expect(splitName('"Madonna"', 'm@y.example')).toEqual({ firstName: 'Madonna', lastName: '' });
  });
  it('derives a name from the local part when the display name is empty', () => {
    expect(splitName('', 'jordan.ellis@granitepeak.example')).toEqual({ firstName: 'Jordan', lastName: 'Ellis' });
    expect(splitName('', 'sam@x.example')).toEqual({ firstName: 'Sam', lastName: '' });
  });
});
