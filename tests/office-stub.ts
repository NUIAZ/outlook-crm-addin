/**
 * tests/office-stub.ts: a minimal fake `Office` global for testing host.ts.
 *
 * Covers exactly the slice office.d.ts declares: onReady, context.mailbox.item
 * in read or compose shape, userProfile, addHandlerAsync, and EventType. Tests
 * install it with `installOffice(...)` before importing or calling the host
 * module, and remove it with `uninstallOffice()` so the "no host" path is also
 * exercised. Nothing here is clever; it exists so the real detection and
 * normalisation code runs in jsdom instead of being mocked away.
 */

export interface StubAddress { displayName: string; emailAddress: string }

export interface StubReadItem {
  mode: 'read';
  from?: StubAddress;
  sender?: StubAddress;
  subject: string;
  to: StubAddress[];
  cc: StubAddress[];
  internetMessageId?: string;
  itemId?: string;
  body?: string;
}

export interface StubComposeItem {
  mode: 'compose';
  subject: string;
  to: StubAddress[];
  cc: StubAddress[];
  body?: string;
}

export interface InstallOptions {
  host?: string | null;
  item?: StubReadItem | StubComposeItem | null;
  userEmail?: string;
  /** Make onReady never resolve, to test the detection timeout. */
  hang?: boolean;
}

/** Installs window.Office. Returns a handle to fire ItemChanged from a test. */
export function installOffice(opts: InstallOptions = {}): { fireItemChanged: () => void; setItem: (i: InstallOptions['item']) => void } {
  const handlers: (() => void)[] = [];
  let current = opts.item ?? null;

  const itemObject = () => {
    if (!current) return null;
    if (current.mode === 'read') {
      const r = current;
      return {
        from: r.from, sender: r.sender, subject: r.subject, to: r.to, cc: r.cc,
        internetMessageId: r.internetMessageId, itemId: r.itemId,
        body: { getAsync: (_t: string, cb: (res: { status: 'succeeded'; value: string }) => void) => cb({ status: 'succeeded', value: r.body ?? '' }) },
      };
    }
    const c = current;
    return {
      subject: { getAsync: (cb: (res: { status: 'succeeded'; value: string }) => void) => cb({ status: 'succeeded', value: c.subject }) },
      to: { getAsync: (cb: (res: { status: 'succeeded'; value: StubAddress[] }) => void) => cb({ status: 'succeeded', value: c.to }) },
      cc: { getAsync: (cb: (res: { status: 'succeeded'; value: StubAddress[] }) => void) => cb({ status: 'succeeded', value: c.cc }) },
      body: { getAsync: (_t: string, cb: (res: { status: 'succeeded'; value: string }) => void) => cb({ status: 'succeeded', value: c.body ?? '' }) },
    };
  };

  const mailbox = {
    get item() { return itemObject(); },
    userProfile: { emailAddress: opts.userEmail ?? 'me@mycompany.example', displayName: 'Me' },
    addHandlerAsync: (_type: unknown, handler: () => void) => { handlers.push(handler); },
  };

  const Office = {
    onReady: (cb?: (info: { host: string | null }) => void) => {
      if (opts.hang) return new Promise<{ host: string | null }>(() => {});
      const info = { host: opts.host === undefined ? 'Outlook' : opts.host };
      cb?.(info);
      return Promise.resolve(info);
    },
    context: { mailbox },
    EventType: { ItemChanged: 'itemChanged' },
  };
  (globalThis as unknown as { Office: unknown }).Office = Office;
  return {
    fireItemChanged: () => { for (const h of handlers) h(); },
    setItem: i => { current = i ?? null; },
  };
}

export function uninstallOffice(): void {
  delete (globalThis as unknown as { Office?: unknown }).Office;
}
