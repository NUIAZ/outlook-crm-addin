/**
 * office/ConnectivityGate.tsx: "is the backend reachable?" before anything renders.
 *
 * The internal add-in needed this because the CRM lived on the office network:
 * open the pane off VPN and every call fails. The gate probes the server first,
 * shows a clear "connect to VPN, then retry" screen instead of a cascade of
 * errors, re-probes every 20 seconds and on the webview's `online` event, and
 * any API call that dies mid-session flips the pane back to the gate.
 *
 * The demo has no backend, so by default the gate is a pass-through. Set
 * VITE_BACKEND_PROBE_URL at build time to point it at something and the full
 * behaviour is live; the pattern is kept because it is the part of an add-in
 * most teams forget until the first user opens it from a hotel.
 */

import { useCallback, useEffect, useState, type ReactNode } from 'react';

const PROBE_URL: string = (import.meta.env.VITE_BACKEND_PROBE_URL as string | undefined) ?? '';
const RECHECK_MS = 20_000;

type Gate = 'checking' | 'online' | 'offline';

async function probe(url: string): Promise<boolean> {
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 5000);
    const res = await fetch(url, { method: 'HEAD', cache: 'no-store', signal: ctrl.signal });
    clearTimeout(t);
    return res.ok;
  } catch {
    return false;
  }
}

export function ConnectivityGate({ children }: { readonly children: ReactNode }) {
  const [gate, setGate] = useState<Gate>(PROBE_URL ? 'checking' : 'online');

  const check = useCallback(async () => {
    if (!PROBE_URL) return;
    setGate(await probe(PROBE_URL) ? 'online' : 'offline');
  }, []);

  useEffect(() => {
    if (!PROBE_URL) return;
    void check();
    const timer = setInterval(() => void check(), RECHECK_MS);
    const onOnline = () => void check();
    window.addEventListener('online', onOnline);
    return () => { clearInterval(timer); window.removeEventListener('online', onOnline); };
  }, [check]);

  if (gate === 'online') return <>{children}</>;
  if (gate === 'checking') return <p className="gate" role="status">Checking the connection…</p>;
  return (
    <div className="gate" role="alert">
      <h2>The CRM is not reachable</h2>
      <p>Connect to the office network or VPN, then retry. The pane will also retry on its own every 20 seconds.</p>
      <button type="button" onClick={() => void check()}>Retry now</button>
    </div>
  );
}
