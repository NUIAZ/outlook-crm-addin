/**
 * crm/actions.ts: the three things the pane can do with an email, as pure
 * builders and checks. The store applies the results; these never touch it.
 *
 *   logEmail     -> an Activity on the contact (+ company, + optional opportunity)
 *   createFollowUp -> a FollowUp due after an admin-managed timeframe
 *   createTask   -> a Task, subject prefilled from the email, optional due date
 *
 * The duplicate rule. Every record produced from an email carries the Outlook
 * `itemId`. Before logging, `findDuplicates` looks for existing records with
 * the same itemId. The pane then WARNS and lets the user proceed anyway. That
 * "warn-but-allow" choice is deliberate and matches how the rest of the
 * internal CRM handled duplicates: the software can notice, but it does not
 * know whether you meant to log the same thread against two opportunities.
 * Blocking would be wrong as often as it is right; warning is always right.
 *
 * Body capture is opt-in and capped: an email body can be long, it can contain
 * things the user did not intend to copy into a shared system, and an
 * activity note is for the gist. The cap is a constant so it is easy to find.
 */

import { FOLLOW_UP_TIMEFRAMES, type Activity, type CrmData, type FollowUp, type Task } from './types';
import type { EmailContext, ParticipantMatch } from './match';

/** Maximum characters of email body appended to an activity note. */
export const BODY_CAPTURE_LIMIT = 1500;

export interface LogEmailInput {
  readonly ctx: EmailContext;
  readonly target: ParticipantMatch;
  readonly opportunityId?: string;
  /** Whether the user ticked "include the email body". */
  readonly includeBody: boolean;
  /** Free-text note the user typed, appended before any captured body. */
  readonly note?: string;
}

/** Records already in the CRM that came from this same Outlook item. */
export interface DuplicateReport {
  readonly activities: readonly Activity[];
  readonly followUps: readonly FollowUp[];
  readonly tasks: readonly Task[];
  readonly any: boolean;
}

export function findDuplicates(data: CrmData, itemId: string | undefined): DuplicateReport {
  if (!itemId) return { activities: [], followUps: [], tasks: [], any: false };
  const activities = data.activities.filter(a => a.itemId === itemId);
  const followUps = data.followUps.filter(f => f.itemId === itemId);
  const tasks = data.tasks.filter(t => t.itemId === itemId);
  return { activities, followUps, tasks, any: activities.length + followUps.length + tasks.length > 0 };
}

/** Clip a body for capture: collapse whitespace, cut at the limit with an ellipsis. */
export function clipBody(body: string | undefined, limit = BODY_CAPTURE_LIMIT): string {
  const flat = (body ?? '').replace(/\s+/g, ' ').trim();
  if (flat.length <= limit) return flat;
  return flat.slice(0, limit - 1).trimEnd() + '…';
}

/** Stable-enough id for demo records; the real CRM assigned these server-side. */
export function newId(prefix: string, now: Date = new Date(), salt = Math.random()): string {
  return `${prefix}-${now.getTime().toString(36)}-${Math.floor(salt * 1e6).toString(36)}`;
}

/** Build the Activity for "Log email". Pure; `now` and `id` are injectable for tests. */
export function buildEmailActivity(input: LogEmailInput, now: Date = new Date(), id = newId('a', now)): Activity {
  const { ctx, target, opportunityId, includeBody, note } = input;
  const parts: string[] = [];
  if (note?.trim()) parts.push(note.trim());
  if (includeBody) {
    const clipped = clipBody(ctx.body);
    if (clipped) parts.push(`Email body: ${clipped}`);
  }
  const who = ctx.mode === 'read'
    ? `From ${ctx.from?.name || ctx.from?.email || 'unknown sender'}`
    : `To ${[...ctx.to, ...ctx.cc].map(a => a.name || a.email).filter(Boolean).join(', ') || 'recipients'}`;
  parts.unshift(`${ctx.mode === 'read' ? 'Received' : 'Sent'} email. ${who}.`);
  return {
    id,
    contactId: target.contact.id,
    companyId: target.company.id,
    ...(opportunityId ? { opportunityId } : {}),
    kind: 'Email',
    subject: ctx.subject || '(no subject)',
    notes: parts.join('\n'),
    occurredAt: now.toISOString(),
    ...(ctx.itemId ? { itemId: ctx.itemId } : {}),
  };
}

/** Due date for a timeframe id, as an ISO date. Unknown ids fall back to 1 week. */
export function dueDateFor(timeframeId: string, now: Date = new Date()): string {
  const tf = FOLLOW_UP_TIMEFRAMES.find(t => t.id === timeframeId) ?? FOLLOW_UP_TIMEFRAMES[2];
  const d = new Date(now.getTime() + tf.days * 24 * 60 * 60 * 1000);
  return d.toISOString().slice(0, 10);
}

export interface FollowUpInput {
  readonly ctx: EmailContext;
  readonly target: ParticipantMatch;
  readonly opportunityId?: string;
  readonly timeframeId: string;
  /** Defaults to the email subject. */
  readonly subject?: string;
}

export function buildFollowUp(input: FollowUpInput, now: Date = new Date(), id = newId('f', now)): FollowUp {
  const { ctx, target, opportunityId, timeframeId, subject } = input;
  return {
    id,
    contactId: target.contact.id,
    companyId: target.company.id,
    ...(opportunityId ? { opportunityId } : {}),
    subject: (subject ?? '').trim() || `Follow up: ${ctx.subject || '(no subject)'}`,
    dueOn: dueDateFor(timeframeId, now),
    done: false,
    ...(ctx.itemId ? { itemId: ctx.itemId } : {}),
  };
}

export interface TaskInput {
  readonly ctx: EmailContext;
  readonly target: ParticipantMatch;
  readonly opportunityId?: string;
  readonly subject?: string;
  /** ISO date or empty for none. */
  readonly dueOn?: string;
}

export function buildTask(input: TaskInput, now: Date = new Date(), id = newId('t', now)): Task {
  const { ctx, target, opportunityId, subject, dueOn } = input;
  return {
    id,
    contactId: target.contact.id,
    companyId: target.company.id,
    ...(opportunityId ? { opportunityId } : {}),
    subject: (subject ?? '').trim() || ctx.subject || '(no subject)',
    ...(dueOn ? { dueOn } : {}),
    done: false,
    ...(ctx.itemId ? { itemId: ctx.itemId } : {}),
  };
}

/**
 * What the real add-in would have POSTed. The demo shows this under "Export"
 * so the transport boundary is visible: everything above this line is the
 * same in production; only the function that receives the payload differs.
 */
export interface OutboundPayload {
  readonly endpoint: '/api/activities' | '/api/followups' | '/api/tasks' | '/api/contacts';
  readonly body: unknown;
}
