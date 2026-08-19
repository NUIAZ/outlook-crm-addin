/**
 * App.tsx: the taskpane.
 *
 * Flow:
 *   1. Detect the host. Inside Outlook, read the open item and subscribe to
 *      item changes (the pane can be pinned and followed between emails).
 *      Anywhere else, show the test-mode bar and take the email from there.
 *   2. Match the email's participants against the CRM (crm/match.ts).
 *   3. If someone matched: context card + actions. If nobody did: offer to
 *      create the first unknown participant as a contact, pre-filled.
 *   4. "Recently added" and "Export" at the bottom show exactly what the pane
 *      wrote and what the real add-in would have posted; "Reset demo data"
 *      puts the seed back.
 *
 * State lives in two places only: the open EmailContext (React state) and the
 * CRM data (the store, read through useSyncExternalStore so every write
 * re-renders the pane). No global state library; the pane is small on purpose.
 */

import { useCallback, useEffect, useMemo, useState, useSyncExternalStore } from 'react';
import { getStore } from './crm/store';
import { matchEmail, type EmailContext, type ParticipantMatch } from './crm/match';
import { onItemChanged, ownAddress, readBody, readItem, watchHost, type HostKind } from './office/host';
import { ConnectivityGate } from './office/ConnectivityGate';
import { TestModeBar, sampleToContext } from './components/TestModeBar';
import { ContextCard } from './components/ContextCard';
import { ActionForms } from './components/ActionForms';
import { CreateContact } from './components/CreateContact';
import { RecentAndExport } from './components/RecentAndExport';

export function App() {
  const store = getStore();
  const data = useSyncExternalStore(store.subscribe, store.getData, store.getData);

  const [host, setHost] = useState<HostKind | 'detecting'>('detecting');
  const [ctx, setCtx] = useState<EmailContext | null>(null);
  const [primaryOverride, setPrimaryOverride] = useState<ParticipantMatch | null>(null);
  const [opportunityId, setOpportunityId] = useState<string | undefined>(undefined);
  const [toast, setToast] = useState<string | null>(null);

  // 1. Host detection, then read the item (and keep reading it when pinned).
  //    watchHost can report twice: 'none' quickly (so the user is not staring
  //    at a spinner while Outlook bootstraps), then 'outlook' when onReady
  //    finally lands, at which point the pane upgrades in place and the test
  //    bar disappears. Falling into test mode INSIDE Outlook and staying there
  //    was a real bug; see office/host.ts.
  useEffect(() => {
    let cancelled = false;
    watchHost(kind => {
      if (cancelled) return;
      setHost(kind);
      if (kind === 'outlook') {
        const load = async () => { const item = await readItem(); if (!cancelled) { setCtx(item); setPrimaryOverride(null); setOpportunityId(undefined); } };
        void load();
        onItemChanged(() => { void load(); });
      } else {
        setCtx(sampleToContext('sample-known'));
      }
    });
    return () => { cancelled = true; };
  }, []);

  // 2. Match.
  const own = useMemo(() => (host === 'outlook' ? [ownAddress()].filter((x): x is string => !!x) : ['you@yourcompany.example']), [host]);
  const match = useMemo(() => (ctx ? matchEmail(data, ctx, own) : null), [data, ctx, own]);
  const primary = primaryOverride && match?.all.includes(primaryOverride) ? primaryOverride : match?.primary ?? null;

  // Body loader: Office in Outlook, the typed body in test mode.
  const loadBody = useCallback(async () => (host === 'outlook' ? readBody() : ctx?.body ?? ''), [host, ctx]);

  const say = (m: string) => { setToast(m); window.setTimeout(() => setToast(null), 4000); };

  if (host === 'detecting') return <p className="gate" role="status">Starting…</p>;

  return (
    <ConnectivityGate>
      <div className="pane">
        <header className="top">
          <h1>
            {/* The same mark as the ribbon button (public/icons), so the pane is
                recognisably the thing you clicked. Served relative to the page so
                it works from Pages and from localhost alike. */}
            <img className="logo" src="./icons/icon-64.png" width="24" height="24" alt="" />
            Log to CRM
          </h1>
          <span className={`host ${host}`}>{host === 'outlook' ? 'In Outlook' : 'Browser test mode'}</span>
        </header>

        {host === 'none' && ctx && <TestModeBar value={ctx} onChange={c => { setCtx(c); setPrimaryOverride(null); setOpportunityId(undefined); }} />}

        {toast && <p className="toast" role="status">{toast}</p>}

        {!ctx && host === 'outlook' && (
          <p className="card muted">Open an email (or start a draft) and the pane will look its sender and recipients up in the CRM.</p>
        )}

        {ctx && match && (
          <>
            <p className="email-line small">
              {ctx.mode === 'read' ? 'Reading' : 'Writing'}: <strong>{ctx.subject || '(no subject)'}</strong>
              {ctx.mode === 'read' && ctx.from ? <> from {ctx.from.name || ctx.from.email}</> : null}
              {ctx.mode === 'compose' ? <> to {[...ctx.to, ...ctx.cc].map(a => a.name || a.email).join(', ') || 'nobody yet'}</> : null}
            </p>

            {primary ? (
              <>
                <ContextCard data={data} match={match} primary={primary} onChoosePrimary={m => { setPrimaryOverride(m); setOpportunityId(undefined); }} selectedOpportunityId={opportunityId} onSelectOpportunity={setOpportunityId} />
                <ActionForms store={store} data={data} ctx={ctx} target={primary} opportunityId={opportunityId} loadBody={loadBody} onDone={say} />
              </>
            ) : match.unknownCandidate ? (
              <CreateContact store={store} data={data} match={match} candidate={match.unknownCandidate} onCreated={say} />
            ) : (
              <p className="card muted">Nobody on this email to look up yet.</p>
            )}
          </>
        )}

        <RecentAndExport store={store} data={data} onReset={() => { store.reset(); setPrimaryOverride(null); setOpportunityId(undefined); say('Demo data reset to the seed.'); }} />

        <footer className="foot small muted">
          Fictional CRM, stored only in this host's localStorage. Nothing is sent anywhere.{' '}
          <a href="https://github.com/NUIAZ/outlook-crm-addin" target="_blank" rel="noopener noreferrer">Source and how to sideload</a>
        </footer>
      </div>
    </ConnectivityGate>
  );
}
