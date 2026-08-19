/**
 * crm/match.ts: who on this email is in the CRM?
 *
 * Pure functions over the CRM data and a description of the email. No Office,
 * no React, no storage: tests/match.test.ts exercises them directly.
 *
 * The rule, in order:
 *   1. In READ mode, the sender first. That is who you are replying to.
 *   2. Then To, then CC, in the order Outlook gives them. This is what catches
 *      the assistant-writes-buyer-is-CC'd case: Hannah sends, Priya is on CC,
 *      and Priya is the contact that matters.
 *   3. In COMPOSE mode there is no sender worth matching (it is you), so the
 *      recipients are matched in To, CC order.
 * The first exact email match wins and becomes the "primary" contact; every
 * other participant that matched is listed as well so the user can switch.
 *
 * If nobody matches by address, we still try to match the sender's DOMAIN
 * against company domains. That does not identify a person, but it lets the
 * pane say "this is someone new at Granite Peak Outfitters" and pre-fill the
 * company when creating them, which is the most common real-world case: a new
 * person at an existing customer.
 */

import type { Company, Contact, CrmData } from './types';

/** One address on the email, as Office.js gives it to us (name may be empty). */
export interface EmailAddress {
  readonly name: string;
  readonly email: string;
}

/** The parts of the open item the matcher needs. Built by office/host.ts. */
export interface EmailContext {
  readonly mode: 'read' | 'compose';
  /** Outlook item id in read mode; absent for an unsaved draft. */
  readonly itemId?: string;
  readonly subject: string;
  readonly from?: EmailAddress;
  readonly to: readonly EmailAddress[];
  readonly cc: readonly EmailAddress[];
  /** Plain-text body, only fetched when the user opts in to capture it. */
  readonly body?: string;
}

export interface ParticipantMatch {
  readonly address: EmailAddress;
  /** Which field the address came from. */
  readonly role: 'from' | 'to' | 'cc';
  readonly contact: Contact;
  readonly company: Company;
}

export interface MatchResult {
  /** The contact the pane acts on by default, or null if nobody matched. */
  readonly primary: ParticipantMatch | null;
  /** Every participant that matched, primary first. */
  readonly all: readonly ParticipantMatch[];
  /**
   * When nobody matched by address but an address's domain belongs to a known
   * company, that company, so "create contact" can pre-fill it. Null otherwise.
   */
  readonly domainCompany: Company | null;
  /** The address we would create a contact for if the user chooses to. */
  readonly unknownCandidate: EmailAddress | null;
}

/** Lower-cased, trimmed; matching is never case-sensitive. */
export function normalizeEmail(email: string): string {
  return email.trim().toLowerCase();
}

/** The part after the @, lower-cased, or '' if malformed. */
export function domainOf(email: string): string {
  const at = normalizeEmail(email).lastIndexOf('@');
  return at < 0 ? '' : normalizeEmail(email).slice(at + 1);
}

/**
 * Participants in matching priority order for the given mode. Exported because
 * the UI uses the same order to list "everyone on this email".
 */
export function participantsInOrder(ctx: EmailContext): readonly { address: EmailAddress; role: 'from' | 'to' | 'cc' }[] {
  const out: { address: EmailAddress; role: 'from' | 'to' | 'cc' }[] = [];
  if (ctx.mode === 'read' && ctx.from?.email) out.push({ address: ctx.from, role: 'from' });
  for (const a of ctx.to) if (a.email) out.push({ address: a, role: 'to' });
  for (const a of ctx.cc) if (a.email) out.push({ address: a, role: 'cc' });
  return out;
}

/**
 * Match an email against the CRM.
 *
 * `ownAddresses` lets the caller exclude the user's own mailbox (in compose,
 * Office gives you yourself in no field, but in read mode you are usually in
 * To, and you are not a contact to log against).
 */
export function matchEmail(data: CrmData, ctx: EmailContext, ownAddresses: readonly string[] = []): MatchResult {
  const own = new Set(ownAddresses.map(normalizeEmail));
  const byEmail = new Map<string, Contact>();
  for (const c of data.contacts) byEmail.set(normalizeEmail(c.email), c);
  const companyById = new Map(data.companies.map(c => [c.id, c]));
  const companyByDomain = new Map(data.companies.map(c => [c.domain.toLowerCase(), c]));

  const all: ParticipantMatch[] = [];
  const seen = new Set<string>();
  let unknownCandidate: EmailAddress | null = null;
  let domainCompany: Company | null = null;

  for (const p of participantsInOrder(ctx)) {
    const key = normalizeEmail(p.address.email);
    if (!key || own.has(key) || seen.has(key)) continue;
    seen.add(key);
    const contact = byEmail.get(key);
    if (contact) {
      const company = companyById.get(contact.companyId);
      if (company) all.push({ address: p.address, role: p.role, contact, company });
      continue;
    }
    // First unknown participant is the one we offer to create. In read mode
    // that is almost always the sender, which is the intent.
    if (!unknownCandidate) {
      unknownCandidate = p.address;
      domainCompany = companyByDomain.get(domainOf(key)) ?? null;
    }
  }

  return {
    primary: all[0] ?? null,
    all,
    // Only surface the domain guess when nobody matched outright; if a known
    // contact is on the email the "new person" flow is secondary.
    domainCompany: all.length === 0 ? domainCompany : null,
    unknownCandidate,
  };
}

/**
 * Split a display name into first/last for the create-contact form. "Dr. Anita
 * Shah" -> { first: "Dr. Anita", last: "Shah" }; a single token goes in first.
 * If the name is empty, derive something from the local part of the address
 * ("jordan.ellis" -> "Jordan Ellis") so the form is never blank.
 */
export function splitName(name: string, email: string): { firstName: string; lastName: string } {
  const clean = name.replace(/["<>]/g, '').trim();
  if (clean) {
    const parts = clean.split(/\s+/);
    if (parts.length === 1) return { firstName: parts[0], lastName: '' };
    return { firstName: parts.slice(0, -1).join(' '), lastName: parts[parts.length - 1] };
  }
  const local = normalizeEmail(email).split('@')[0] ?? '';
  const tokens = local.split(/[._-]+/).filter(Boolean).map(t => t.charAt(0).toUpperCase() + t.slice(1));
  if (tokens.length === 0) return { firstName: '', lastName: '' };
  if (tokens.length === 1) return { firstName: tokens[0], lastName: '' };
  return { firstName: tokens.slice(0, -1).join(' '), lastName: tokens[tokens.length - 1] };
}
