# Security

This is a demo of an Outlook add-in, and add-ins run inside your mailbox, so the
posture is worth stating plainly.

## What sideloading this add-in means

An Outlook add-in with `ReadItem` permission can read the message you currently have
open (sender, recipients, subject, body on request) while its pane is showing, and if
you pin the pane, every message you open after that. That is true of every mail add-in;
it is what the "Log to CRM" feature needs. Outlook warns you when you sideload a custom
add-in for exactly this reason. Sideload only from origins you trust; this one is
served from `https://nuiaz.github.io/outlook-crm-addin/`, under the same GitHub account
that owns this repository, and a test pins every URL in the manifest to that origin.

## What the demo does with what it reads

Nothing leaves the host. There are no network requests in the pane except one optional
`HEAD` probe in the connectivity gate, which is off unless `VITE_BACKEND_PROBE_URL` is
set at build time (it is not, for the published build). What you read and what you add
is written to the host's `localStorage` under one key, on your machine, readable only
by this origin. The email body is fetched only if you tick "include the email body" and
is capped at 1,500 characters before storage. **Reset demo data** removes everything the
pane stored; clearing site data for the origin does the same.

## Permissions and scope

- Manifest `<Permissions>ReadItem</Permissions>`: the least-privileged level that can
  read the open item. Not `ReadWriteItem`, not `ReadWriteMailbox`.
- `AppDomains` is the Pages origin only. Every URL in the manifest is HTTPS.
- Two extension points (read and compose surfaces), one taskpane, no event-based
  activation, no mailbox-wide access.

## Code-level controls

- No `innerHTML`, `dangerouslySetInnerHTML`, `eval`, or `document.write`. Everything
  from an email (subject, names, addresses, body) is rendered by React as text, and
  the Export drawer shows `JSON.stringify` output inside a `<pre>`. Subject lines are
  written by other people; the pane treats them as hostile.
- Two runtime dependencies (react, react-dom). The build is static files with no
  secrets, no environment variables, and no server.
- A Content Security Policy in `index.html`: same-origin for everything except
  Office.js. See the next section.

## The one external script

`office.js` is loaded from `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`.
Microsoft requires add-ins to load it from that CDN (bundling is not permitted) and
updates it in place, so Subresource Integrity cannot be used. The CSP allows that host
for `script-src` and `connect-src` (Office.js fetches its own locale and telemetry
files) and nothing else off-origin. If the pane ever stops loading inside Outlook after
a host update, the CSP is the first thing to check; loosen `connect-src` for the Office
host before anything else.

`frame-ancestors` is intentionally absent: browsers ignore it in a `<meta>` policy, and
GitHub Pages cannot send HTTP headers. If you self-host, set it as a header to the
Outlook origins.

## What the real add-in had to do that the demo does not

The internal version talked to a live CRM. The controls that mattered there, and that
this demo does not need, were:

- Authentication in the Outlook webview via the CRM's cookie flow (`Secure`,
  `HttpOnly`, `SameSite`), with silent refresh, and CSRF protection on the API.
- All validation and duplicate rules enforced **server-side**; the pane's checks were
  conveniences, never the gate.
- Authorisation on the API (owner/team visibility) so the pane only saw what the user
  was allowed to see.
- HTTPS with a certificate the client trusts; Exchange rejects any `http://` URL in a
  manifest.
- A deliberate privacy decision about body capture (opt-in, capped, appended to a note)
  owned by the CRM's administrators, not by the add-in author.
- No caching of CRM data in the webview's `localStorage`. The demo can, because its
  data is fictional.

## Reporting

If you find something, open an issue on this repository or message the author. There
is no bounty; there is gratitude.
