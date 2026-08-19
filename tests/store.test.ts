/**
 * tests/store.test.ts: where the data goes, and that it comes back.
 */
import { describe, expect, it } from 'vitest';
import { STORAGE_KEY, createStore } from '../src/crm/store';
import { buildSeed } from '../src/crm/seed';

const NOW = new Date('2026-08-19T12:00:00Z');

/** In-memory Storage so tests do not depend on jsdom's localStorage behaviour. */
function memStorage(initial: Record<string, string> = {}) {
  const m = new Map(Object.entries(initial));
  return { getItem: (k: string) => m.get(k) ?? null, setItem: (k: string, v: string) => { m.set(k, v); }, removeItem: (k: string) => { m.delete(k); }, _map: m };
}

describe('store', () => {
  it('seeds on first run and persists the seed immediately', () => {
    const s = memStorage();
    const store = createStore(s, () => NOW);
    expect(store.wasRestored()).toBe(false);
    expect(store.getData().contacts).toHaveLength(18);
    expect(s._map.has(STORAGE_KEY)).toBe(true);
  });

  it('restores from storage on the next run, including records the pane added', () => {
    const s = memStorage();
    const first = createStore(s, () => NOW);
    first.addTask({ id: 't-9', contactId: 'c-1', companyId: 'co-bluefin', subject: 'Call back', done: false, itemId: '<m@x>' });
    const second = createStore(s, () => NOW);
    expect(second.wasRestored()).toBe(true);
    expect(second.getData().tasks.some(t => t.id === 't-9')).toBe(true);
  });

  it('ignores a corrupt or wrong-version record and reseeds', () => {
    const corrupt = createStore(memStorage({ [STORAGE_KEY]: '{not json' }), () => NOW);
    expect(corrupt.wasRestored()).toBe(false);
    const old = createStore(memStorage({ [STORAGE_KEY]: JSON.stringify({ version: 0, data: { contacts: [] } }) }), () => NOW);
    expect(old.wasRestored()).toBe(false);
    expect(old.getData().contacts).toHaveLength(18);
  });

  it('works with no storage at all (in-memory session)', () => {
    const store = createStore(null, () => NOW);
    store.addContact({ id: 'c-99', firstName: 'A', lastName: 'B', email: 'a@b.example', companyId: 'co-bluefin', source: 'Email' });
    expect(store.getData().contacts.some(c => c.id === 'c-99')).toBe(true);
  });

  it('notifies subscribers on every write and stops after unsubscribe', () => {
    const store = createStore(null, () => NOW);
    let n = 0;
    const off = store.subscribe(() => { n++; });
    store.addCompany({ id: 'co-x', name: 'X', domain: 'x.example', industry: 'Test' });
    store.addTask({ id: 't', contactId: 'c-1', companyId: 'co-bluefin', subject: 's', done: false });
    expect(n).toBe(2);
    off();
    store.reset();
    expect(n).toBe(2);
  });

  it('addedByAddIn lists only records that came from the pane, newest first', () => {
    const store = createStore(null, () => NOW);
    store.addActivity({ id: 'a-x', contactId: 'c-1', companyId: 'co-bluefin', kind: 'Email', subject: 'first', occurredAt: NOW.toISOString(), itemId: '<1@x>' });
    store.addActivity({ id: 'a-y', contactId: 'c-1', companyId: 'co-bluefin', kind: 'Email', subject: 'second', occurredAt: NOW.toISOString(), itemId: '<2@x>' });
    store.addContact({ id: 'c-new', firstName: 'N', lastName: 'P', email: 'n@p.example', companyId: 'co-bluefin', source: 'Email' });
    const added = store.addedByAddIn();
    expect(added.activities.map(a => a.id)).toEqual(['a-y', 'a-x']);   // seed activities have no itemId
    expect(added.contacts.map(c => c.id)).toEqual(['c-new']);         // seed contacts have no source
  });

  it('reset restores the seed exactly', () => {
    const store = createStore(null, () => NOW);
    store.addTask({ id: 't', contactId: 'c-1', companyId: 'co-bluefin', subject: 's', done: false });
    store.reset();
    expect(store.getData()).toEqual(buildSeed(NOW));
  });
});
