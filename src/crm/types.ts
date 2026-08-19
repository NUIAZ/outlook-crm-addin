/**
 * crm/types.ts: the shape of the CRM the pane talks to.
 *
 * In the internal version these were DTOs from a real CRM API. Here they are
 * the same shapes backed by an in-memory dataset persisted to localStorage
 * (see store.ts), so the pane's logic is identical and only the transport
 * differs. Keeping the types in one file is what makes that swap honest: the
 * UI and the matching/action logic never import anything from the store's
 * persistence layer, only these shapes.
 */

/** A company a contact works for. */
export interface Company {
  readonly id: string;
  readonly name: string;
  /** Email domain used for domain-level matching when the person is unknown. */
  readonly domain: string;
  readonly industry: string;
}

/** A person in the CRM. */
export interface Contact {
  readonly id: string;
  readonly firstName: string;
  readonly lastName: string;
  /** Stored lower-cased; matching is case-insensitive on both sides. */
  readonly email: string;
  readonly companyId: string;
  readonly title?: string;
  /** Where the contact came from, e.g. "Email" when created from the pane. */
  readonly source?: string;
}

export type OpportunityStage = 'Prospecting' | 'Qualified' | 'Proposal' | 'Negotiation';

/** A deal with a company; the pane lets you link a logged email to one. */
export interface Opportunity {
  readonly id: string;
  readonly companyId: string;
  readonly name: string;
  readonly stage: OpportunityStage;
  readonly amount: number;
}

/**
 * Something that happened with a contact. Logged emails become activities.
 * `itemId` is the Outlook item id of the email that produced it, which is what
 * the duplicate check keys on: log the same email twice and the pane warns.
 */
export interface Activity {
  readonly id: string;
  readonly contactId: string;
  readonly companyId: string;
  readonly opportunityId?: string;
  readonly kind: 'Email' | 'Call' | 'Meeting' | 'Note';
  readonly subject: string;
  readonly notes?: string;
  /** ISO date-time. */
  readonly occurredAt: string;
  /** Outlook item id, present when the activity came from the add-in. */
  readonly itemId?: string;
}

/** A reminder to get back to someone. */
export interface FollowUp {
  readonly id: string;
  readonly contactId: string;
  readonly companyId: string;
  readonly opportunityId?: string;
  readonly subject: string;
  /** ISO date. */
  readonly dueOn: string;
  readonly done: boolean;
  readonly itemId?: string;
}

/** A to-do, optionally dated. */
export interface Task {
  readonly id: string;
  readonly contactId: string;
  readonly companyId: string;
  readonly opportunityId?: string;
  readonly subject: string;
  /** ISO date, or absent for "no due date". */
  readonly dueOn?: string;
  readonly done: boolean;
  readonly itemId?: string;
}

/** Everything the pane reads and writes, as one serialisable object. */
export interface CrmData {
  readonly companies: readonly Company[];
  readonly contacts: readonly Contact[];
  readonly opportunities: readonly Opportunity[];
  readonly activities: readonly Activity[];
  readonly followUps: readonly FollowUp[];
  readonly tasks: readonly Task[];
}

/** Admin-managed follow-up timeframes, as the internal app had them. */
export const FOLLOW_UP_TIMEFRAMES: readonly { readonly id: string; readonly label: string; readonly days: number }[] = [
  { id: 'tomorrow', label: 'Tomorrow', days: 1 },
  { id: '3d', label: 'In 3 days', days: 3 },
  { id: '1w', label: 'In a week', days: 7 },
  { id: '2w', label: 'In two weeks', days: 14 },
  { id: '1m', label: 'In a month', days: 30 },
];
