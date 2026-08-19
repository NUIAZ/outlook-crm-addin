/**
 * tests/manifest.test.ts: the manifest is the contract with Outlook, and
 * Exchange rejects it for reasons that only show up at sideload time. Catch
 * the ones we can at test time: every URL https and on our origin, every icon
 * it references actually shipped, both surfaces present, the id well-formed.
 */
import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const xml = readFileSync(resolve('public/manifest.xml'), 'utf8');
const ORIGIN = 'https://nuiaz.github.io/outlook-crm-addin/';

describe('manifest.xml', () => {
  it('has a well-formed GUID id and a version', () => {
    expect(xml).toMatch(/<Id>[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}<\/Id>/);
    expect(xml).toMatch(/<Version>\d+\.\d+\.\d+\.\d+<\/Version>/);
  });

  it('uses only https URLs, all on the Pages origin or GitHub', () => {
    const urls = [...xml.matchAll(/DefaultValue="(https?:\/\/[^"]+)"/g)].map(m => m[1]);
    expect(urls.length).toBeGreaterThan(5);
    for (const u of urls) {
      expect(u.startsWith('https://'), u).toBe(true);
      expect(u.startsWith(ORIGIN) || u.startsWith('https://github.com/NUIAZ/'), u).toBe(true);
    }
    // XML namespace URIs (schemas.microsoft.com) are legitimately http:; only
    // resource URLs matter to Exchange, and those are the DefaultValue ones.
    expect(xml).not.toMatch(/DefaultValue="http:\/\//);
  });

  it('ships every icon it references', () => {
    const icons = [...xml.matchAll(new RegExp(`${ORIGIN.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')}icons/([\\w.-]+\\.png)`, 'g'))].map(m => m[1]);
    expect(new Set(icons).size).toBeGreaterThanOrEqual(5);
    for (const f of icons) expect(existsSync(resolve('public/icons', f)), f).toBe(true);
  });

  it('declares both the read and the compose surfaces, and supports pinning on read', () => {
    expect(xml).toContain('MessageReadCommandSurface');
    expect(xml).toContain('MessageComposeCommandSurface');
    expect(xml).toContain('<SupportsPinning>true</SupportsPinning>');
    expect(xml).toContain('<Permissions>ReadItem</Permissions>');
  });

  it('points the taskpane at the site root, which index.html serves', () => {
    expect(xml).toContain(`<bt:Url id="Taskpane.Url" DefaultValue="${ORIGIN}" />`);
    expect(existsSync(resolve('index.html'))).toBe(true);
  });
});
