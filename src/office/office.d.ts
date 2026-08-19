/**
 * office/office.d.ts: ambient types for the slice of Office.js this add-in uses.
 *
 * Hand-written on purpose. @types/office-js is large and the pane touches a
 * dozen members; typing exactly those makes the contract with the host visible
 * in one screen, and keeps `any` out. Swap for the official types if the
 * add-in grows.
 *
 * Two things about Outlook's object model that shape the host adapter:
 *   - READ mode exposes plain values (`item.from`, `item.to`, `item.subject`),
 *     COMPOSE mode exposes async getters (`item.to.getAsync(cb)`), because a
 *     draft's recipients can change under you. The adapter normalises both.
 *   - `internetMessageId` (the RFC 5322 Message-ID) exists only for received
 *     mail and is stable across mailboxes; `itemId` is the Exchange id. The
 *     duplicate check prefers the former so the same email logged from two
 *     mailboxes is still recognised.
 */

interface OfficeEmailAddressDetails {
  displayName: string;
  emailAddress: string;
}

interface OfficeAsyncResult<T> {
  status: 'succeeded' | 'failed';
  value?: T;
  error?: { message?: string };
}

interface OfficeBody {
  getAsync(coercionType: 'text' | 'html', callback: (result: OfficeAsyncResult<string>) => void): void;
}

/** Compose-mode recipients and subject are async getter objects. */
interface OfficeRecipients {
  getAsync(callback: (result: OfficeAsyncResult<OfficeEmailAddressDetails[]>) => void): void;
}
interface OfficeComposeSubject {
  getAsync(callback: (result: OfficeAsyncResult<string>) => void): void;
}

interface OfficeMailboxItem {
  /** Read mode only. `sender` is the on-behalf-of fallback. */
  from?: OfficeEmailAddressDetails;
  sender?: OfficeEmailAddressDetails;
  /** String in read mode; async getter in compose. */
  subject?: string | OfficeComposeSubject;
  /** Arrays in read mode; async Recipients objects in compose. */
  to?: OfficeEmailAddressDetails[] | OfficeRecipients;
  cc?: OfficeEmailAddressDetails[] | OfficeRecipients;
  /** RFC Message-ID; received mail only. */
  internetMessageId?: string;
  /** Exchange item id; read mode. */
  itemId?: string;
  body?: OfficeBody;
}

interface OfficeMailbox {
  item?: OfficeMailboxItem | null;
  /** The signed-in mailbox owner: used to skip "me" when matching. */
  userProfile?: { emailAddress?: string; displayName?: string };
  addHandlerAsync?(eventType: unknown, handler: () => void, callback?: (result: unknown) => void): void;
}

interface OfficeNamespace {
  onReady(callback?: (info: { host: string | null }) => void): Promise<{ host: string | null }>;
  context?: { mailbox?: OfficeMailbox };
  EventType?: { ItemChanged: unknown };
}

/** Undefined when office.js did not load (plain browser, blocked CDN). */
declare const Office: OfficeNamespace | undefined;
