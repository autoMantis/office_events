# Email to Event — Outlook add-in

Read an email, press one button, get a prefilled appointment form. Subject from the
email, meeting link detected, body carried over. You add attendees and save.

Live at **https://events.jorgemenas.com** — sideload `manifest.xml` from there, no local
server needed. Everything runs client-side via Office.js; nothing is sent to any server
except the one-time static file load. No secrets, no backend, no data collection.

## How it decides

| Email contains | Location | Body |
|---|---|---|
| Teams join link | `Microsoft Teams Meeting` | Join link at the top + email body |
| Zoom / Meet / Webex / GoTo / Whereby | the link itself | Join link at the top + email body |
| no link | empty | email body |

Body is truncated at 2000 characters. Start defaults to the next `:00`/`:30`,
30 minutes — change it in the pane before creating.

## Files

| File | What |
|---|---|
| `manifest.xml` | Add-in manifest (XML — works on classic Outlook desktop and web) |
| `taskpane.html` / `taskpane.js` | The pane UI and Office.js glue |
| `parse.js` | Link detection + body building. No Office.js, so it's unit-testable |
| `test-parse.js` | `node test-parse.js` — the only test |
| `assets/` | Ribbon icons |

## Use it (production)

Download `manifest.xml` from <https://events.jorgemenas.com/manifest.xml> (or the repo),
then sideload:

- **Outlook on the web / new Outlook**: Settings → General → Manage add-ins →
  *My add-ins* → *Add a custom add-in* → *Add from file* → pick `manifest.xml`.
- **Classic Outlook desktop**: same dialog, reachable from *Get Add-ins* on the ribbon.

Open any email → **Email to Event** on the ribbon.

## Local dev

```bash
npm install
npm run certs
npm start
```

Open <https://localhost:3000/taskpane.html> once in a browser and accept the cert if
prompted. Point a copy of `manifest.xml` at `localhost:3000` to sideload the dev version
(don't edit the committed one — that one is pinned to production).

## Deploy

```bash
npx wrangler deploy
```

`wrangler.jsonc` publishes the repo root as static assets (no server code, no billing
per request). Custom domain `events.jorgemenas.com` is attached in the Cloudflare
dashboard under this Worker's Settings → Domains & Routes.

## "Add-in Error — something went wrong and we couldn't start this add-in"

1. Does <https://events.jorgemenas.com/taskpane.html> load in a browser? If not, redeploy.
2. Still failing: remove the add-in from *My add-ins* and sideload `manifest.xml`
   again — Outlook caches the manifest.

## Known limits

- Office.js cannot flip the **Teams meeting** toggle on the new appointment form. A
  Teams link is put in the body and the location says `Microsoft Teams Meeting`;
  clicking the link joins. If you want a *real* Teams-organised meeting, toggle it
  yourself in the form — it will add its own link.
- No date/time parsing from the email text. Deliberate: "call me at 5" produces more
  wrong appointments than right ones. Add it in `parse.js` → `defaultSlot` if you
  change your mind.
- Requires Mailbox requirement set 1.5+.

## Test

```bash
npm test
```
