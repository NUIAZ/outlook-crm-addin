/**
 * tests/host.test.ts: the Office.js seam, driven through the stub.
 *
 * The things most likely to be wrong in an add-in are exactly the ones a
 * browser-only dev loop never hits: compose mode's async getters, the sender
 * fallback, the RFC message id preference, and what happens when Office.js is
 * missing or slow. Each is a test here.
 */
import { afterEach, describe, expect, it } from 'vitest';
import { installOffice, uninstallOffice } from './office-stub';
import { detectHost, onItemChanged, ownAddress, readBody, readItem, watchHost } from '../src/office/host';

afterEach(() => uninstallOffice());

describe('detectHost', () => {
  it("reports 'none' when Office.js is absent", async () => {
    uninstallOffice();
    expect(await detectHost()).toBe('none');
  });
  it("reports 'none' when Office.js loaded but has no host (plain browser)", async () => {
    installOffice({ host: null });
    expect(await detectHost()).toBe('none');
  });
  it("reports 'outlook' inside Outlook", async () => {
    installOffice({ host: 'Outlook' });
    expect(await detectHost()).toBe('outlook');
  });
  it("times out to 'none' if onReady never resolves (slow or broken CDN)", async () => {
    installOffice({ hang: true });
    expect(await detectHost(50)).toBe('none');
  });
});

describe('watchHost: the slow-host upgrade', () => {
  // The real-world failure this guards: Outlook's onReady landing AFTER the
  // fallback timeout. The pane must first report 'none' (test mode appears)
  // and then 'outlook' (the pane upgrades in place), in that order.
  it("reports 'none' at the timeout, then upgrades to 'outlook' when onReady lands late", async () => {
    installOffice({ readyAfterMs: 80 });
    const reports: string[] = [];
    watchHost(k => reports.push(k), 20);
    await new Promise(r => setTimeout(r, 150));
    expect(reports).toEqual(['none', 'outlook']);
  });
  it('reports only once when onReady beats the timeout', async () => {
    installOffice({});
    const reports: string[] = [];
    watchHost(k => reports.push(k), 1000);
    await new Promise(r => setTimeout(r, 50));
    expect(reports).toEqual(['outlook']);
  });
  it('never downgrades and never duplicates', async () => {
    installOffice({ readyAfterMs: 30 });
    const reports: string[] = [];
    watchHost(k => reports.push(k), 10);
    await new Promise(r => setTimeout(r, 120));
    expect(reports.filter(r => r === 'outlook')).toHaveLength(1);
    expect(reports[reports.length - 1]).toBe('outlook');
  });
});

describe('readItem', () => {
  it('returns null with no open item', async () => {
    installOffice({ item: null });
    expect(await readItem()).toBeNull();
  });

  it('normalises a read-mode item: plain values, sender, both recipients, RFC id', async () => {
    installOffice({ item: {
      mode: 'read', from: { displayName: 'Priya Raman', emailAddress: 'priya.raman@bluefin-logistics.example' },
      subject: 'Re: quote', to: [{ displayName: 'Me', emailAddress: 'me@mycompany.example' }], cc: [{ displayName: '', emailAddress: 'cc@x.example' }],
      internetMessageId: '<rfc@mail.example>', itemId: 'AAMk-exchange-id',
    } });
    const ctx = await readItem();
    expect(ctx).toEqual({
      mode: 'read', itemId: '<rfc@mail.example>', subject: 'Re: quote',
      from: { name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' },
      to: [{ name: 'Me', email: 'me@mycompany.example' }], cc: [{ name: '', email: 'cc@x.example' }],
    });
  });

  it('falls back to the Exchange item id when there is no RFC id, and to sender when from is missing', async () => {
    installOffice({ item: { mode: 'read', sender: { displayName: 'S', emailAddress: 's@x.example' }, subject: 'x', to: [], cc: [], itemId: 'AAMk1' } });
    const ctx = await readItem();
    expect(ctx?.itemId).toBe('AAMk1');
    expect(ctx?.from?.email).toBe('s@x.example');
  });

  it('normalises a compose-mode item: async getters, no sender, no item id', async () => {
    installOffice({ item: { mode: 'compose', subject: 'Draft to Wes', to: [{ displayName: 'Wes', emailAddress: 'wes.thorne@granitepeak.example' }], cc: [] } });
    const ctx = await readItem();
    expect(ctx?.mode).toBe('compose');
    expect(ctx?.subject).toBe('Draft to Wes');
    expect(ctx?.to).toEqual([{ name: 'Wes', email: 'wes.thorne@granitepeak.example' }]);
    expect(ctx?.from).toBeUndefined();
    expect(ctx?.itemId).toBeUndefined();
  });
});

describe('readBody, ownAddress, onItemChanged', () => {
  it('reads the plain-text body on demand, and returns empty without an item', async () => {
    installOffice({ item: { mode: 'read', subject: 's', to: [], cc: [], body: 'hello body' } });
    expect(await readBody()).toBe('hello body');
    uninstallOffice();
    expect(await readBody()).toBe('');
  });
  it('exposes the signed-in address', () => {
    installOffice({ userEmail: 'me@mycompany.example' });
    expect(ownAddress()).toBe('me@mycompany.example');
    uninstallOffice();
    expect(ownAddress()).toBeUndefined();
  });
  it('subscribes to ItemChanged and the handler fires', () => {
    const h = installOffice({});
    let fired = 0;
    onItemChanged(() => { fired++; });
    h.fireItemChanged();
    expect(fired).toBe(1);
  });
  it('is a no-op without a host', () => {
    uninstallOffice();
    expect(() => onItemChanged(() => {})).not.toThrow();
  });
});
