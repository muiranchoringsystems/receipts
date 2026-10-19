# Muir Receipts — PWA

Snap a card receipt, pick the expense type, add the reason, send. The app uploads the photo (or PDF)
plus a `_meta.json` sidecar straight into the Finance SharePoint folder
`9_Admin/Finance/Credit Card Receipts/Incoming` as the signed-in user. The scheduled Claude run
(skill `card-receipts-register`) picks it up from there.

No app store. Users open the link once, "Add to Home Screen", and get a Muir icon on iOS and Android.


## Security model (read before go-live)
- **Access**: nothing reaches SharePoint without a Microsoft sign-in to the Muir tenant (single-tenant registration; tenant MFA / conditional access apply). No secrets in the code.
- **Staying signed in**: Microsoft caps SPA refresh tokens at 24 h (not configurable). The app refreshes sign-in when it opens; with a live Microsoft session (Stay signed in / KMSI, or a Conditional Access persistent-browser + sign-in-frequency policy) that is a brief bounce with no password. How often a password/MFA is actually asked is set by the tenant session policy, not the app.
- **Lost phone**: Entra › Users › the person › **Revoke sessions** (kills the session and refresh tokens; existing access tokens lapse within the hour, usually sooner where SharePoint enforces continuous access evaluation). Reset the password too. Nothing unsent on the phone can be uploaded after that.
- **No lost receipts**: every send is written to the phone's queue before any upload; an expired sign-in never navigates away mid-send.
- **Scope**: delegated `Files.ReadWrite.All` lets a signed-in user's token reach anything *that user* can reach. The app only writes to `Incoming`, so the risk is code tampering: **protect the hosting account** (GitHub 2FA, Max sole owner, no collaborators).
- **Entra › Enterprise applications › Muir Receipts › Properties › Assignment required = Yes**, assign only cardholders.
- **CSP** is set via a `<meta>` tag in `index.html` (GitHub Pages can't send headers): scripts from self only; network only to login.microsoftonline.com, graph.microsoft.com, *.sharepoint.com. Verified: external scripts are refused.
- Later: ask IT whether a "selected sites" Graph permission can replace `Files.ReadWrite.All` for this app.
- Mailbox: set `receipts@` to accept mail from inside the organisation only.

## Files

| File | Purpose |
|---|---|
| `index.html` | the app (single screen flow: sign in → capture → details → done) |
| `app.js` | logic: MSAL sign-in, image downscale, Graph upload, offline queue (IndexedDB), recent list |
| `config.js` | **edit this**: Entra client/tenant IDs, SharePoint drive/folder IDs, expense-type list |
| `sw.js` | service worker: caches the shell so the icon opens instantly / offline |
| `manifest.webmanifest`, `icons/` | PWA identity (name, colours, icons) |
| `msal-browser.min.js` | Microsoft auth library (vendored, v3.28.1) |
| `staticwebapp.config.json` | Azure Static Web Apps routing + security headers |

## Setup (IT, ~30 minutes)

### 1. Entra app registration (so the app can sign users in and write to SharePoint as them)
1. Entra admin centre › App registrations › **New registration**. Name `Muir Receipts`. Supported account types: *this organisation only*. Redirect URI: platform **Single-page application**, URI = the site URL you will host at (e.g. `https://receipts.muir.com.au/` or the Static Web App URL). Register.
2. Copy **Application (client) ID** and **Directory (tenant) ID** into `config.js`.
3. API permissions › Add › Microsoft Graph › **Delegated**: `User.Read`, `Files.ReadWrite.All`. Click **Grant admin consent** (so staff never see a consent prompt).
4. Authentication › confirm the SPA redirect URI; if you later move hosts, add the new URI here too. No client secret is needed (PKCE).

### 2. SharePoint permissions
Cardholders need to write to `Incoming` and nothing else in Finance. On the `Incoming` folder: *Manage access* › stop inheriting › grant the **Cardholders** security group **Contribute** (or *Edit*); leave the rest of `Credit Card Receipts` finance-only. Users without access get a clear "Graph 403" in the app.

### 3. Host it (any HTTPS static host; Azure Static Web Apps free tier recommended)
- Azure portal › Create **Static Web App** › Free plan › deployment: *Other* (upload) or connect a GitHub repo containing this folder. The `staticwebapp.config.json` is picked up automatically.
- Or with the CLI: `npm i -g @azure/static-web-apps-cli` then `swa deploy ./ --env production` from this folder.
- Optional custom domain `receipts.muir.com.au` (CNAME, free SSL). Put the final URL in the Entra redirect URI (step 1).
- Alternative hosts that also work: GitHub Pages (private repo + Pages), Cloudflare Pages, an IIS/nginx static site. Only `index.html`'s CSP in `staticwebapp.config.json` is Azure-specific; replicate the headers on other hosts if you want them.


### Hosting on GitHub Pages (Muir has no Azure subscription — 20 Oct 2026)
1. Create a GitHub account for Muir (e.g. `muir-engineering`) and a **public** repository named `receipts`.
   Public is fine: the code holds no secrets (an SPA client ID and tenant ID are public by design; access is
   enforced by Microsoft sign-in and SharePoint permissions).
2. "Add file › Upload files" → drag in the *contents* of this folder (index.html, app.js, config.js, sw.js,
   manifest.webmanifest, msal-browser.min.js, icons/). `staticwebapp.config.json` is Azure-only and can be left out.
3. Settings › Pages → Source: *Deploy from a branch*, branch `main`, folder `/ (root)` → Save. The site appears at
   `https://<account>.github.io/receipts/` within a minute or two.
4. In the Entra app registration add **both** SPA redirect URIs: `https://<account>.github.io/receipts/` and
   `https://<account>.github.io/receipts/index.html`.
5. To release an update: upload the changed files over the old ones (bump `APP_VERSION` and `CACHE` first).

### 4. Roll out
Send staff the link. iOS: open in **Safari** › Share › *Add to Home Screen*. Android (Chrome): menu › *Install app* / *Add to Home screen*. First open asks them to sign in with their Muir Microsoft account; after that the icon goes straight to the camera button.

## Changing the expense-type list without redeploying
Put a JSON file in `Credit Card Receipts` (e.g. `app-config.json`) with any of `expenseTypes`, `attendeeTypes`, `projects`, `reasonPlaceholder`, and set its item id in `config.js › configFileId`. The app merges it over the defaults on launch. Keep the type names identical to the register's Lookups sheet.

## What lands in `Incoming`
```
20261019-141205_max.buckley_1.jpg
20261019-141205_max.buckley_2.jpg        (extra pages, if any)
20261019-141205_max.buckley_meta.json    (uploaded last = set is complete)
```
`_meta.json` (schema `muir-receipt/1`): submitter name + UPN, captured_at, expense_type, reason,
attendees, project, card (company|personal), foreign_currency, files[]. The Power Automate flow for
emailed receipts uses the same `YYYYMMDD-HHmmss_<sender>` prefix with a `_body.txt` sidecar, so the
register run treats both sources alike.

## "No time now" — send the photo, add details later
On the details screen, **No time now — send photo, add details later** uploads the photo(s) immediately with
`_meta.json` `status: "pending_details"` (so finance already has the receipt and the register shows it as
*Needs info*), and keeps a local draft in the app's **To finish** list. Tapping the draft reopens the details
screen; **Send details** re-uploads `_meta.json` with `status: "complete"` (replace). If the run has already
filed the photos, it treats the new meta as an update to that Receipt ID.

## Behaviour notes
- Photos are downscaled to 1600 px on the long edge, JPEG 85% (~300–900 KB); EXIF rotation is applied. PDFs are uploaded untouched; files over 4 MB use a Graph upload session.
- Offline or failed uploads queue in the browser (IndexedDB) and send automatically when the app is next opened online; a yellow banner shows the count.
- "Reason" is mandatory (min 3 chars); attendees are mandatory for Entertainment and Travel-Meals (FBT).
- Sign-in uses redirect flow (works inside installed PWAs on iOS where popups don't). Tokens cache in localStorage; *Sign out* in the footer clears them.
- Demo mode for previews: open `index.html?demo=1` — no sign-in, uploads are simulated and logged to the console.

## Testing locally
`python3 -m http.server 8765` in this folder, then open `http://localhost:8765/index.html?demo=1`. Service workers and the camera input work on `localhost` without HTTPS.

## Release checklist
Bump `APP_VERSION` in `app.js` and `CACHE` in `sw.js`, redeploy. Installed apps pick the new version up on next open (network-first shell).
