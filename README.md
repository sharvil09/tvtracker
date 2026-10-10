# WatchTogether — Google-only edition

Recovered from the user's deployed Sheets backend. Google Apps Script now serves `Index.html` and invokes the backend through `google.script.run`. There is no Cloudflare Worker, proxy secret, or browser-supplied identity. Google supplies the authenticated user's email; every API call checks it against the two configured accounts before accessing the database.

## Install in your existing Apps Script project

1. Open your [WatchTogether sheet](https://docs.google.com/spreadsheets/d/18dy64N5JlpWg5tFZxIf4KzpHLYxj6nQjunsA1JrheAI/edit) → Extensions → Apps Script.
2. Replace `Code.gs` with `apps-script/Code.gs` from this repository.
3. Click **+ → HTML**, name it **Index**, and replace its contents with `apps-script/Index.html`.
4. Project Settings → Show appsscript.json manifest file in editor. Replace that file with `apps-script/appsscript.json`.
5. Keep the existing Script Properties: SPREADSHEET_ID, USER_A_EMAIL, USER_A_NAME, USER_B_EMAIL, USER_B_NAME, TMDB_READ_TOKEN. Optional REGION and TIMEZONE default to US and America/New_York. PROXY_SECRET is no longer used. No Google OAuth client ID or client secret is required by this version.
6. Share the spreadsheet with the second configured Google account as an **Editor**. Keep general access **Restricted**. Both users can directly edit the shared database, so this design is for two trusted people.
7. Run `setupWatchTogether` while signed in as one of the configured accounts. It preserves existing table records, validates headers, and creates any missing tabs. Authorize spreadsheet, external requests (TMDB), and account-email access.
8. **Deploy → Manage deployments → Edit (pencil)**. Choose **New version**. Set **Execute as: User accessing the web app** and **Who has access: Anyone with Google account** (Google's label can vary by account; require signed-in Google users). Click **Deploy**. Both users must authorize access when they first open it. If this setting cannot be changed on your deployment, create a new Web app deployment with these settings and use its new URL.
9. Open the resulting `/exec` URL. Updating the existing deployment normally retains your existing URL:
   https://script.google.com/macros/s/AKfycbydNFQNFpdym04WQ_vhNjuOaRfi2Cv0eNdl7gZAb4xQR1O9BPV1Nkq-tSXBj5xk3IeZZQ/exec

The new `userinfo.email` scope requires renewed authorization. The second user must also have spreadsheet edit access. Do not use **Execute as: Me**: Google may conceal the caller's email in that mode; the backend rejects missing identities. Do not configure anonymous access. Google Workspace policies and OAuth verification requirements can restrict deployment or authorization; follow Google's displayed guidance for your own private app rather than approving unrelated permissions.

## App behavior

- Search TMDB for films and series; add to watchlist, watching, or watched.
- Track each title independently. Episode increments follow canonical aired episodes across seasons.
- Optional four-tier ratings and at most three pairwise comparisons. Logging watched status never forces a rating.
- Separate libraries for the two verified accounts.
- Upcoming episodes, per-show follow preferences, metadata refresh, and regional streaming availability.
- Profile preferences, streaming services, CSV import in batches of five, and ambiguous-match review.
- Polling every 20 seconds while visible, skipped during local saves or open dialogs.
- Locked writes and stable request journal IDs preserve the recovered backend's retry protection. Sheets still has no multi-row transactions; refresh after an interrupted edit and inspect state before trying a new edit.

This is a Google-hosted web app. Apps Script's iframe hosting does not provide the former installable offline PWA. Automatic playback logging, Trakt, push notifications, and offline persistence are deferred. No fixture catalog or viewing history is included.

## Verify after deployment

Sign in as account A, search for a title, add it, reload, and confirm persistence. Advance a series episode; verify its row in `tracking`. Sign in as account B and confirm its library is separate. Try a third account: backend access must be denied. Import a small CSV and resolve an uncertain match. Check mobile layout.

Local checks: `node --test tests/*.test.cjs`. These checks mock Google services. They do not verify the user's live deployment, permissions, TMDB token, or sheet data.
