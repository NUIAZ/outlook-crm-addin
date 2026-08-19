/**
 * crm/store.ts: the demo's CRM "backend", which is localStorage.
 *
 * WHERE THE DATA GOES (the question everyone asks):
 *   Everything you add in the pane (a logged email, a follow-up, a task, a new
 *   contact) is written to `localStorage` under one key, on the machine the
 *   pane is running on, inside whichever host is showing it. No network request
 *   is made; there is no server. That means:
 *     - it persists across opens of the pane, so logging the same email twice
 *       really does trip the duplicate warning;
 *     - it is per device AND per host: Outlook desktop, Outlook on the web, and
 *       a browser tab each have their own storage for this origin, so a record
 *       made in one will not show in another;
 *     - clearing site data for the origin removes it;
 *     - nobody else can see it.
 *   "Reset demo data" restores the seed. "Export" shows the records the real
 *   add-in would have POSTed to the CRM API, so the transport boundary is
 *   visible rather than hidden.
 *
 * In the internal version this file was a thin HTTP client; every function
 * below had the same name and signature and called an endpoint instead. The
 * components and the pure logic in match.ts / actions.ts did not change.
 *
 * Design notes:
 *   - The store is a tiny observable: `subscribe` + `getData`, so React can use
 *     `useSyncExternalStore` and re-render on writes without a state library.
 *   - Writes are whole-object replace-and-persist. The dataset is small; clarity
 *     beats cleverness.
 *   - A schema version is stored alongside the data, so a future seed change
 *     can migrate or reset rather than render garbage.
 */

import { buildSeed } from './seed';
import type { Activity, Company, Contact, CrmData, FollowUp, Task } from './types';

export const STORAGE_KEY = 'outlook-crm-addin:v1';
const SCHEMA_VERSION = 1;

interface Persisted {
  readonly version: number;
  readonly data: CrmData;
  readonly savedAt: string;
}

type Listener = () => void;

export interface CrmStore {
  getData(): CrmData;
  subscribe(listener: Listener): () => void;
  addActivity(a: Activity): void;
  addFollowUp(f: FollowUp): void;
  addTask(t: Task): void;
  addContact(c: Contact): void;
  /** Demo-only: the real CRM created companies on their own screen. */
  addCompany(c: Company): void;
  /** Records created by the add-in (those carrying an itemId), newest first. */
  addedByAddIn(): { activities: Activity[]; followUps: FollowUp[]; tasks: Task[]; contacts: Contact[] };
  reset(): void;
  /** True if the current data came from storage rather than a fresh seed. */
  wasRestored(): boolean;
}

/**
 * Create a store. `storage` is injectable so tests can pass an in-memory map
 * and so a host without localStorage (some Office webviews in private mode
 * block it) degrades to an in-memory session instead of crashing.
 */
export function createStore(storage: Pick<Storage, 'getItem' | 'setItem' | 'removeItem'> | null = safeLocalStorage(), now: () => Date = () => new Date()): CrmStore {
  let data: CrmData;
  let restored = false;
  const listeners = new Set<Listener>();

  const persist = () => {
    if (!storage) return;
    const rec: Persisted = { version: SCHEMA_VERSION, data, savedAt: now().toISOString() };
    try { storage.setItem(STORAGE_KEY, JSON.stringify(rec)); } catch { /* quota or blocked: keep going in memory */ }
  };
  const set = (next: CrmData) => { data = next; persist(); for (const l of listeners) l(); };

  // Load: a valid persisted record of the current schema wins; anything else
  // (absent, corrupt, old version) starts from the seed.
  try {
    const raw = storage?.getItem(STORAGE_KEY);
    const parsed = raw ? (JSON.parse(raw) as Partial<Persisted>) : null;
    if (parsed && parsed.version === SCHEMA_VERSION && parsed.data && Array.isArray(parsed.data.contacts)) {
      data = parsed.data as CrmData;
      restored = true;
    } else {
      data = buildSeed(now());
      persist();
    }
  } catch {
    data = buildSeed(now());
  }

  return {
    getData: () => data,
    subscribe(l) { listeners.add(l); return () => { listeners.delete(l); }; },
    addActivity(a) { set({ ...data, activities: [...data.activities, a] }); },
    addFollowUp(f) { set({ ...data, followUps: [...data.followUps, f] }); },
    addTask(t) { set({ ...data, tasks: [...data.tasks, t] }); },
    addContact(c) { set({ ...data, contacts: [...data.contacts, c] }); },
    addCompany(c) { set({ ...data, companies: [...data.companies, c] }); },
    addedByAddIn() {
      const byNewest = <T extends { id: string }>(xs: readonly T[]) => [...xs].reverse();
      return {
        activities: byNewest(data.activities.filter(a => !!a.itemId)),
        followUps: byNewest(data.followUps.filter(f => !!f.itemId)),
        tasks: byNewest(data.tasks.filter(t => !!t.itemId)),
        contacts: byNewest(data.contacts.filter(c => c.source === 'Email')),
      };
    },
    reset() { restored = false; set(buildSeed(now())); },
    wasRestored: () => restored,
  };
}

/** localStorage if it is usable here, else null (private-mode webviews can throw). */
export function safeLocalStorage(): Storage | null {
  try {
    const probe = '__outlook-crm-addin-probe__';
    window.localStorage.setItem(probe, '1');
    window.localStorage.removeItem(probe);
    return window.localStorage;
  } catch {
    return null;
  }
}

/** One shared store for the app. Tests build their own with createStore(). */
let shared: CrmStore | null = null;
export function getStore(): CrmStore {
  if (!shared) shared = createStore();
  return shared;
}
