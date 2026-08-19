// scripts/screenshots.mjs: captures docs/*.png from a running copy of the pane.
// Playwright is not a dependency; point PLAYWRIGHT_DIR at a project that has it.
//   SITE_URL=http://localhost:5173/ PLAYWRIGHT_DIR=../x/node_modules/playwright BROWSER_CHANNEL=msedge node scripts/screenshots.mjs
// Default SITE_URL is the Pages deployment. The pane is captured at taskpane
// width (400px) in light and dark, for the known-sender sample, after logging
// once (so the duplicate warning shows), and for the unknown-sender sample.
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { mkdirSync } from 'node:fs';

const pw = process.env.PLAYWRIGHT_DIR ? pathToFileURL(resolve(process.env.PLAYWRIGHT_DIR, 'index.mjs')).href : 'playwright';
const { chromium } = await import(pw);
const root = resolve(import.meta.dirname, '..');
const out = resolve(root, 'docs');
mkdirSync(out, { recursive: true });
const site = process.env.SITE_URL || 'https://nuiaz.github.io/outlook-crm-addin/';
const browser = await chromium.launch({ channel: process.env.BROWSER_CHANNEL || undefined });

async function shoot(name, { scheme = 'light', prepare, full = true } = {}) {
  const page = await browser.newPage({ viewport: { width: 400, height: 900 }, colorScheme: scheme });
  await page.goto(site, { waitUntil: 'networkidle' });
  await page.getByRole('heading', { name: /Priya Raman/ }).waitFor();
  if (prepare) await prepare(page);
  await page.waitForTimeout(300);
  await page.screenshot({ path: resolve(out, name), fullPage: full });
  console.log('wrote', name);
  await page.close();
}

// Hide the test-mode bar for the hero so it reads like the pane inside Outlook.
const hideTestBar = async page => { await page.addStyleTag({ content: '.testbar{display:none}' }); };
await shoot('hero.png', { prepare: hideTestBar });
await shoot('hero-dark.png', { scheme: 'dark', prepare: hideTestBar });
await shoot('with-testbar.png');
await shoot('duplicate-warning.png', { prepare: async page => {
  await hideTestBar(page);
  await page.getByRole('button', { name: 'Log email' }).click();
  await page.getByText(/Already logged from this email/).waitFor();
} });
await shoot('create-contact.png', { prepare: async page => {
  await page.getByLabel('Sample email').selectOption('sample-unknown');
  await page.getByRole('heading', { name: /Not in the CRM yet/ }).waitFor();
  await hideTestBar(page);
} });
await browser.close();
