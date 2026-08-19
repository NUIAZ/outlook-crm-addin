/**
 * vitest.setup.ts: clear localStorage between tests so the persistence tests
 * are independent of execution order. Nothing else is global on purpose: the
 * Office stub is installed per test so host detection is really tested.
 */
import { beforeEach } from 'vitest';
beforeEach(() => { try { localStorage.clear(); } catch { /* unavailable: tests that care inject their own */ } });
