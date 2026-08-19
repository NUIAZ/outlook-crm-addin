/**
 * tests/app.test.tsx: the pane end to end in jsdom, in both hosts.
 *
 * In "browser" (no Office) the pane must fall into test mode, load the first
 * sample, show Priya's context card, and let the user log the email, and the
 * second log of the same email must show the duplicate warning. In "Outlook"
 * (stub installed) the pane must read the real item, and follow an ItemChanged
 * event. These tests render the real App against a real store (in-memory).
 */
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { installOffice, uninstallOffice } from './office-stub';

// A fresh App + store module per test: the shared store is module-level.
async function mountApp() {
  const { App } = await import('../src/App');
  return render(<App />);
}

beforeEach(() => { localStorage.clear(); });
afterEach(() => { uninstallOffice(); });

describe('browser test mode', () => {
  it('falls back to test mode, loads a sample, and shows the matched contact', async () => {
    uninstallOffice();
    await mountApp();
    expect(await screen.findByText(/Browser test mode/)).toBeTruthy();
    expect(await screen.findByRole('heading', { name: /Priya Raman/ })).toBeTruthy();
    expect(screen.getByText(/Open opportunities at Bluefin Logistics/)).toBeTruthy();
  });

  it('logs the email, then warns on the second log of the same email (warn-but-allow)', async () => {
    uninstallOffice();
    const user = userEvent.setup();
    await mountApp();
    await screen.findByRole('heading', { name: /Priya Raman/ });
    const actions = within(screen.getByRole('region', { name: /Actions for this email/ }));
    await user.click(actions.getByRole('button', { name: 'Log email' }));
    expect(await screen.findByText(/Logged "Re: Depot telematics rollout, revised quote" to Priya Raman/)).toBeTruthy();
    // Second time: the warning is shown and the button changes, but still works.
    expect(await actions.findByText(/Already logged from this email/)).toBeTruthy();
    await user.click(actions.getByRole('button', { name: 'Log anyway' }));
    // Both records exist in the store (visible via Recently added).
    await user.click(screen.getByRole('button', { name: /Recently added by this pane \(2\)/ }));
    expect(screen.getAllByText(/Activity: Re: Depot telematics rollout/)).toHaveLength(2);
  });

  it('switching to the unknown-sender sample offers to create a contact at the guessed company', async () => {
    uninstallOffice();
    const user = userEvent.setup();
    await mountApp();
    await screen.findByRole('heading', { name: /Priya Raman/ });
    await user.selectOptions(screen.getByLabelText('Sample email'), 'sample-unknown');
    expect(await screen.findByRole('heading', { name: /Not in the CRM yet/ })).toBeTruthy();
    // The paragraph names the guessed company and the dropdown pre-selects it.
    expect(screen.getByText(/the domain belongs to/).textContent).toContain('Granite Peak Outfitters');
    expect((screen.getByLabelText('Company') as HTMLSelectElement).value).toBe('co-granite');
    expect((screen.getByLabelText('First name') as HTMLInputElement).value).toBe('Jordan');
    await user.click(screen.getByRole('button', { name: 'Create contact' }));
    // Now they match, and the context card appears for them.
    expect(await screen.findByRole('heading', { name: /Jordan Ellis/ })).toBeTruthy();
  });

  it('compose mode matches recipients instead of the sender', async () => {
    uninstallOffice();
    const user = userEvent.setup();
    await mountApp();
    await screen.findByRole('heading', { name: /Priya Raman/ });
    await user.click(screen.getByLabelText(/Writing one \(compose\)/));
    // To was "You" (own address) in the sample; set a known recipient.
    const to = screen.getByLabelText('To');
    await user.clear(to);
    await user.type(to, 'Wesley Thorne <wes.thorne@granitepeak.example>');
    expect(await screen.findByRole('heading', { name: /Wesley Thorne/ })).toBeTruthy();
    expect(screen.getByText(/matched from/).textContent).toMatch(/TO/);
  });
});

describe('inside Outlook', () => {
  it('reads the open item and follows ItemChanged', async () => {
    const h = installOffice({ item: {
      mode: 'read', from: { displayName: 'Ruth Abernathy', emailAddress: 'ruth.abernathy@tallgrassseed.example' },
      subject: 'Traceability pilot', to: [{ displayName: 'Me', emailAddress: 'me@mycompany.example' }], cc: [], internetMessageId: '<r1@x>',
    } });
    await mountApp();
    expect(await screen.findByText(/In Outlook/)).toBeTruthy();
    expect(await screen.findByRole('heading', { name: /Ruth Abernathy/ })).toBeTruthy();
    expect(screen.queryByText(/Browser test mode/)).toBeNull();

    h.setItem({ mode: 'read', from: { displayName: 'Marco DeLuca', emailAddress: 'marco@orchardlane.example' }, subject: 'Bread order', to: [], cc: [], internetMessageId: '<m1@x>' });
    h.fireItemChanged();
    expect(await screen.findByRole('heading', { name: /Marco DeLuca/ })).toBeTruthy();
  });

  it('shows a prompt when the pane opens with nothing selected', async () => {
    installOffice({ item: null });
    await mountApp();
    expect(await screen.findByText(/Open an email \(or start a draft\)/)).toBeTruthy();
  });

  it('captures the body from Office when the box is ticked', async () => {
    installOffice({ item: {
      mode: 'read', from: { displayName: 'Priya Raman', emailAddress: 'priya.raman@bluefin-logistics.example' },
      subject: 'Numbers', to: [], cc: [], internetMessageId: '<b1@x>', body: 'The body text from Office',
    } });
    const user = userEvent.setup();
    await mountApp();
    await screen.findByRole('heading', { name: /Priya Raman/ });
    await user.click(screen.getByLabelText(/Include the email body/));
    await user.click(screen.getByRole('button', { name: 'Log email' }));
    await screen.findByText(/Logged "Numbers"/);
    await user.click(screen.getByRole('button', { name: /Recently added/ }));
    await user.click(screen.getByRole('button', { name: /Show what would have been sent/ }));
    await waitFor(() => expect(screen.getByLabelText(/Payloads/).textContent).toContain('The body text from Office'));
  });
});
