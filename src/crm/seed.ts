/**
 * crm/seed.ts: the fictional CRM the demo ships with.
 *
 * Six companies, eighteen contacts, a handful of open opportunities, and some
 * history, so the context card has something to show the first time you open
 * the pane on an email. Every name, company, domain and address is invented.
 * Domains use the reserved `.example` top-level domain (RFC 2606), so nothing
 * here can ever resolve to a real mailbox.
 *
 * Dates are relative to "now" at seed time so "last activity 3 days ago" keeps
 * reading sensibly no matter when the demo is first opened.
 */

import type { CrmData } from './types';

const DAY = 24 * 60 * 60 * 1000;
const iso = (ms: number) => new Date(ms).toISOString();
const day = (ms: number) => new Date(ms).toISOString().slice(0, 10);

/** Build the seed relative to `now` (injectable for deterministic tests). */
export function buildSeed(now: Date = new Date()): CrmData {
  const t = now.getTime();
  return {
    companies: [
      { id: 'co-bluefin', name: 'Bluefin Logistics', domain: 'bluefin-logistics.example', industry: 'Transportation' },
      { id: 'co-granite', name: 'Granite Peak Outfitters', domain: 'granitepeak.example', industry: 'Retail' },
      { id: 'co-sable', name: 'Sable & Co. Roasters', domain: 'sableroasters.example', industry: 'Food and beverage' },
      { id: 'co-meridian', name: 'Meridian Dental Group', domain: 'meridiandental.example', industry: 'Healthcare' },
      { id: 'co-tallgrass', name: 'Tallgrass Seed Company', domain: 'tallgrassseed.example', industry: 'Agriculture' },
      { id: 'co-orchard', name: 'Orchard Lane Bakery', domain: 'orchardlane.example', industry: 'Food and beverage' },
    ],
    contacts: [
      { id: 'c-1', firstName: 'Priya', lastName: 'Raman', email: 'priya.raman@bluefin-logistics.example', companyId: 'co-bluefin', title: 'Operations Director' },
      { id: 'c-2', firstName: 'Tomás', lastName: 'Alvarez', email: 'tomas.alvarez@bluefin-logistics.example', companyId: 'co-bluefin', title: 'Fleet Manager' },
      { id: 'c-3', firstName: 'Hannah', lastName: 'Okafor', email: 'hannah.okafor@bluefin-logistics.example', companyId: 'co-bluefin', title: 'Executive Assistant' },
      { id: 'c-4', firstName: 'Wesley', lastName: 'Thorne', email: 'wes.thorne@granitepeak.example', companyId: 'co-granite', title: 'Owner' },
      { id: 'c-5', firstName: 'Maya', lastName: 'Lindqvist', email: 'maya.lindqvist@granitepeak.example', companyId: 'co-granite', title: 'Buyer' },
      { id: 'c-6', firstName: 'Desmond', lastName: 'Park', email: 'desmond.park@granitepeak.example', companyId: 'co-granite', title: 'Store Manager' },
      { id: 'c-7', firstName: 'Aurelie', lastName: 'Sable', email: 'aurelie@sableroasters.example', companyId: 'co-sable', title: 'Founder' },
      { id: 'c-8', firstName: 'Ken', lastName: 'Moriarty', email: 'ken.moriarty@sableroasters.example', companyId: 'co-sable', title: 'Head of Wholesale' },
      { id: 'c-9', firstName: 'Beatriz', lastName: 'Nunes', email: 'beatriz.nunes@sableroasters.example', companyId: 'co-sable', title: 'Accounts Payable' },
      { id: 'c-10', firstName: 'Dr. Anita', lastName: 'Shah', email: 'anita.shah@meridiandental.example', companyId: 'co-meridian', title: 'Practice Owner' },
      { id: 'c-11', firstName: 'Graham', lastName: 'Whitfield', email: 'graham.whitfield@meridiandental.example', companyId: 'co-meridian', title: 'Office Manager' },
      { id: 'c-12', firstName: 'Lena', lastName: 'Kowalczyk', email: 'lena.k@meridiandental.example', companyId: 'co-meridian', title: 'Procurement' },
      { id: 'c-13', firstName: 'Ruth', lastName: 'Abernathy', email: 'ruth.abernathy@tallgrassseed.example', companyId: 'co-tallgrass', title: 'CEO' },
      { id: 'c-14', firstName: 'Caleb', lastName: 'Osei', email: 'caleb.osei@tallgrassseed.example', companyId: 'co-tallgrass', title: 'Agronomist' },
      { id: 'c-15', firstName: 'Ingrid', lastName: 'Halvorsen', email: 'ingrid.halvorsen@tallgrassseed.example', companyId: 'co-tallgrass', title: 'Logistics Coordinator' },
      { id: 'c-16', firstName: 'Marco', lastName: 'DeLuca', email: 'marco@orchardlane.example', companyId: 'co-orchard', title: 'Owner' },
      { id: 'c-17', firstName: 'Sunita', lastName: 'Bhatt', email: 'sunita.bhatt@orchardlane.example', companyId: 'co-orchard', title: 'Wholesale Manager' },
      { id: 'c-18', firstName: 'Owen', lastName: 'Fitzgerald', email: 'owen.fitzgerald@orchardlane.example', companyId: 'co-orchard', title: 'Bookkeeper' },
    ],
    opportunities: [
      { id: 'o-1', companyId: 'co-bluefin', name: 'Depot telematics rollout', stage: 'Proposal', amount: 48000 },
      { id: 'o-2', companyId: 'co-bluefin', name: 'Q4 driver tablets', stage: 'Qualified', amount: 12500 },
      { id: 'o-3', companyId: 'co-granite', name: 'Point-of-sale refresh, 3 stores', stage: 'Negotiation', amount: 27300 },
      { id: 'o-4', companyId: 'co-sable', name: 'Wholesale ordering portal', stage: 'Prospecting', amount: 9800 },
      { id: 'o-5', companyId: 'co-meridian', name: 'Patient reminder integration', stage: 'Proposal', amount: 15200 },
      { id: 'o-6', companyId: 'co-tallgrass', name: 'Seed lot traceability', stage: 'Qualified', amount: 64000 },
    ],
    activities: [
      { id: 'a-1', contactId: 'c-1', companyId: 'co-bluefin', opportunityId: 'o-1', kind: 'Meeting', subject: 'Telematics scope review', occurredAt: iso(t - 3 * DAY) },
      { id: 'a-2', contactId: 'c-1', companyId: 'co-bluefin', kind: 'Email', subject: 'Re: depot count for the quote', occurredAt: iso(t - 9 * DAY) },
      { id: 'a-3', contactId: 'c-4', companyId: 'co-granite', opportunityId: 'o-3', kind: 'Call', subject: 'Pricing call', occurredAt: iso(t - 1 * DAY) },
      { id: 'a-4', contactId: 'c-7', companyId: 'co-sable', kind: 'Note', subject: 'Met at the trade show', occurredAt: iso(t - 21 * DAY) },
      { id: 'a-5', contactId: 'c-10', companyId: 'co-meridian', opportunityId: 'o-5', kind: 'Email', subject: 'Integration questions', occurredAt: iso(t - 5 * DAY) },
      { id: 'a-6', contactId: 'c-13', companyId: 'co-tallgrass', opportunityId: 'o-6', kind: 'Meeting', subject: 'Site visit', occurredAt: iso(t - 14 * DAY) },
    ],
    followUps: [
      { id: 'f-1', contactId: 'c-1', companyId: 'co-bluefin', opportunityId: 'o-1', subject: 'Send revised telematics quote', dueOn: day(t + 2 * DAY), done: false },
      { id: 'f-2', contactId: 'c-4', companyId: 'co-granite', opportunityId: 'o-3', subject: 'Confirm store 3 install window', dueOn: day(t + 1 * DAY), done: false },
      { id: 'f-3', contactId: 'c-13', companyId: 'co-tallgrass', subject: 'Share traceability case study', dueOn: day(t - 2 * DAY), done: false },
    ],
    tasks: [
      { id: 't-1', contactId: 'c-10', companyId: 'co-meridian', opportunityId: 'o-5', subject: 'Draft integration SOW', dueOn: day(t + 5 * DAY), done: false },
    ],
  };
}

/**
 * Sample emails for manual test mode (when the pane runs outside Outlook).
 * Three deliberately different cases: a known sender; an unknown sender at a
 * known company's domain; and an assistant sending on behalf of a CC'd buyer,
 * which is the case the To/CC matching exists for.
 */
export const SAMPLE_EMAILS = [
  {
    id: 'sample-known',
    label: 'Known sender (Priya at Bluefin)',
    from: { name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' },
    to: [{ name: 'You', email: 'you@yourcompany.example' }],
    cc: [],
    subject: 'Re: Depot telematics rollout, revised quote',
    body: 'Thanks for the call. Could you send the revised quote with the two extra depots included? We would like to take it to the board next Thursday.',
  },
  {
    id: 'sample-unknown',
    label: 'Unknown sender at a known company (new person at Granite Peak)',
    from: { name: 'Jordan Ellis', email: 'jordan.ellis@granitepeak.example' },
    to: [{ name: 'You', email: 'you@yourcompany.example' }],
    cc: [],
    subject: 'Introduction: taking over purchasing from Maya',
    body: 'Hi, Maya has moved to a new role and I will be your contact for the point-of-sale project going forward. Can we set up a call this week?',
  },
  {
    id: 'sample-cc',
    label: 'Assistant writing, real buyer on CC (Hannah for Priya)',
    from: { name: 'Hannah Okafor', email: 'hannah.okafor@bluefin-logistics.example' },
    to: [{ name: 'You', email: 'you@yourcompany.example' }],
    cc: [{ name: 'Priya Raman', email: 'priya.raman@bluefin-logistics.example' }],
    subject: 'Scheduling: board prep with Priya',
    body: 'Priya asked me to find 30 minutes before Thursday for the telematics numbers. Does Tuesday at 2 work?',
  },
  {
    id: 'sample-stranger',
    label: 'Complete stranger (no match at all)',
    from: { name: 'Sam Rivera', email: 'sam.rivera@northbend-plumbing.example' },
    to: [{ name: 'You', email: 'you@yourcompany.example' }],
    cc: [],
    subject: 'Quote request',
    body: 'Hello, we are a 12-van plumbing outfit and would like to talk about vehicle tracking.',
  },
] as const;
