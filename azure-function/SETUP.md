# Intake Form Azure Function — Setup Guide

Follow these in order. Each step says exactly what to click/create. Nothing
here can be done by Claude directly — all of it happens in the Azure
Portal, Entra admin center, SharePoint, and GitHub, outside this repo.

> Adapted from the standalone `FPHM Intake Form` project's own SETUP.md
> after that project was merged into this repo (2026-09-24) — see
> `CLAUDE.md`'s "Intake Forms" section. Steps 1-5 (SharePoint list, Azure
> AD app, Function deploy, pointing the form at it) are unchanged from the
> original. Steps 6-7 (hosting the frontend, registering a separate
> redirect URI) are **no longer needed as separate steps** — the intake
> form and staff view are already part of this repo's existing GitHub
> Pages deployment and existing Azure AD app registration; see the notes
> at the end of this file. The standalone project's own GitHub repo
> (`FPHM2026/prayer-intake`) was deleted 2026-09-25 once this merge was
> confirmed complete — this file is now the only copy of these steps.

---

## 1. Create the SharePoint list

**Fastest way — run the script:** the standalone project's
`scripts/Create-IntakeList.ps1` creates the list and every column below in
one go (installs the PnP.PowerShell module if you don't have it, prompts
an interactive Microsoft 365 sign-in, safe to re-run) — that script wasn't
part of the files merged into this repo, so pull it from the local
`FPHM Intake Form` project folder if you still have it, or recreate the
list by hand below.

**By hand:** Go to
`https://creeksidechurch.sharepoint.com/sites/FreedomPrayer` (the
same site the Scheduler already uses) → Site contents → New → List →
Blank list.

- **Name**: `IntakeResponses`
- Add these columns (Site contents → IntakeResponses → gear icon → List
  settings → Create column, or the "+ Add column" button on the list view):

| Column name | Type | Notes |
|---|---|---|
| `Token` | Single line of text | The resume secret (a random UUID) — never shown to staff, only used internally |
| `RecipientName` | Single line of text | Mirrors the "Name" answer once entered, so this shows in the list view |
| `RecipientEmail` | Single line of text | Mirrors the "Email" answer |
| `Status` | Choice | Values: `InProgress`, `Submitted`. Default: `InProgress` |
| `ResponsesJSON` | Multiple lines of text | **Plain text** (not Rich text — same gotcha as the Scheduler's Notes fields, see this repo's CLAUDE.md gotcha #5). Open the column's advanced settings and enable **"Allow unlimited length"** if you see that option — the JSON blob of all 88 answers is small (a few KB) but default text-column limits can be tight. |
| `SignatureDataUrl` | Multiple lines of text | **Plain text**, also enable "Allow unlimited length" if available. Holds the signature as a base64 PNG data URL. |
| `SubmittedAt` | Date and Time | Include time |

The list's built-in `Title` column is used automatically (set to the
recipient's name once known, or "New intake" before that) — nothing extra
needed there.

No item-level permission changes are needed here (unlike BlackoutDates in
the Scheduler) — writes only ever come from the app-only Azure Function
below, which uses its own site-level grant, not a signed-in user's
permissions. The staff "Intake Forms" tab reads this list via the
Scheduler's normal delegated Graph session instead — see this repo's
CLAUDE.md, "Intake Forms" section, for how the Ministers SharePoint group
gets Read-only access.

---

## 2. Register a new Azure AD app for the Function (app-only, tightly scoped)

This is **separate** from the Scheduler's existing app registration
(`549f5207-d7f8-4924-9bde-30532d90c1d2`) on purpose: the intake data is far
more sensitive (health, abuse, addiction, spiritual disclosures), so it
gets its own credential, scoped to only this one SharePoint site, that
only the Function ever holds — never shipped to a browser.

1. Entra admin center → App registrations → New registration.
   - Name: `FPHM Intake Function` (or similar)
   - Supported account types: single tenant (same as the Scheduler's app)
   - Redirect URI: leave blank — this app never signs a user in, so it
     needs no redirect URI at all.
2. After creation, note the **Application (client) ID** and confirm the
   **Directory (tenant) ID** matches the Scheduler's
   (`b9e447ab-c42c-4b04-92c8-d629b0ce320d`).
3. Certificates & secrets → New client secret → copy the **Value**
   immediately (it's only shown once). This is `CLIENT_SECRET` below.
4. API permissions → Add a permission → Microsoft Graph → **Application
   permissions** (not Delegated) → search `Sites.Selected` → add it.
5. Click **Grant admin consent** for the tenant (needs a Global
   Administrator, same requirement as the Scheduler's original
   `Sites.ReadWrite.All` consent — see this repo's CLAUDE.md gotcha #7).

`Sites.Selected` alone grants **no site access yet** — that's the point:
step 3 below explicitly grants this one app write access to just the one
SharePoint site, nothing else in the tenant.

---

## 3. Grant the new app access to just the FreedomPrayer site

Use [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
signed in as a Global Administrator (or any Graph call authenticated the
same way — PowerShell with `Connect-MgGraph` works too):

1. `GET https://graph.microsoft.com/v1.0/sites/creeksidechurch.sharepoint.com:/sites/FreedomPrayer`
   → copy the `id` field (looks like `creeksidechurch.sharepoint.com,GUID,GUID`).
2. `POST https://graph.microsoft.com/v1.0/sites/{that id}/permissions`
   with body:
   ```json
   {
     "roles": ["write"],
     "grantedToIdentities": [{
       "application": {
         "id": "THE-NEW-APP-CLIENT-ID-FROM-STEP-2",
         "displayName": "FPHM Intake Function"
       }
     }]
   }
   ```
3. A `201 Created` response confirms it. This app can now read/write this
   one site via app-only auth — nothing else on the tenant.

---

## 4. Create and deploy the Azure Function App

1. Azure Portal → Create a resource → Function App.
   - Runtime stack: **Node.js**, version 20 LTS
   - Hosting: **Consumption (Serverless)** plan — this workload is tiny
     (a church intake form), the free grant covers it comfortably.
   - Region: whichever is closest to your church.
   - Name it something like `fphm-intake-func` — your Function App's URL
     will be `https://fphm-intake-func.azurewebsites.net`.
2. Once created, go to the Function App → Configuration → Application
   settings → add these (New application setting for each):
   - `TENANT_ID` = `b9e447ab-c42c-4b04-92c8-d629b0ce320d`
   - `CLIENT_ID` = the new app's client ID (step 2)
   - `CLIENT_SECRET` = the secret value (step 2)
   - `SITE_URL` = `https://creeksidechurch.sharepoint.com/sites/FreedomPrayer`
   - `LIST_NAME` = `IntakeResponses`
   - Save, and let it restart.
3. Function App → CORS (under Settings/API) → add your GitHub Pages
   origin, `https://FPHM2026.github.io` (origin only, no path) →
   Save. This is what lets the public form (a different origin) call
   these endpoints from the browser.
4. Deploy the code in this repo's `azure-function/` folder:
   - Easiest: install the [Azure Functions VS Code extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azurefunctions),
     open the `azure-function` folder in VS Code, sign in to Azure, and use
     "Deploy to Function App" from the extension's sidebar, picking the
     Function App you just created.
   - Or, with [Azure Functions Core Tools](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local)
     and the Azure CLI installed:
     ```bash
     cd "azure-function"
     npm install
     func azure functionapp publish fphm-intake-func
     ```
5. Copy `local.settings.json.example` to `local.settings.json` if you want
   to test locally first (`func start`) — never commit `local.settings.json`
   itself, it holds the real secret.
6. Verify: `POST https://fphm-intake-func.azurewebsites.net/api/intake/start`
   with body `{}` should return `{"token": "..."}`. Check the
   `IntakeResponses` list for a new "New intake" item.

---

## 5. Point the public form at the deployed Function

Edit `intake/js/apiClient.js` in this repo, change:
```js
const FUNCTION_BASE_URL = "https://REPLACE-WITH-YOUR-FUNCTION-APP.azurewebsites.net/api";
```
to your actual Function App URL + `/api`, e.g.
`https://fphm-intake-func.azurewebsites.net/api`. Commit and push
(`preview/` first, then promote to production per this repo's normal
Deployment workflow in CLAUDE.md).

---

## 6. Hosting the frontend — already done, no action needed

Unlike the standalone project (which needed its own new GitHub repo +
Pages setup), the intake form now lives inside this repo's existing
GitHub Pages deployment:
- Public intake link: `https://FPHM2026.github.io/prayer-scheduler/intake/`
- Staff view: the "Intake Forms" tab inside
  `https://FPHM2026.github.io/prayer-scheduler/` (production, once
  promoted) or `.../preview/` (already live there)

Nothing to create here — deploying this repo's `main` branch (as already
happens for the Scheduler) serves `/intake/` automatically.

---

## 7. Staff view's sign-in — already covered, no new redirect URI needed

The standalone project needed a new redirect URI for a separate
`admin/index.html`. That's not the case here: the staff "Intake Forms" tab
is built into `index.html`/`preview/index.html` directly, which already
have their redirect URIs registered under the Scheduler's existing Azure
AD app (`549f5207-d7f8-4924-9bde-30532d90c1d2`) — no extra Entra
configuration needed for this step.

---

## 8. Test end to end

1. Open the public link, fill in a few sections, click "Save & continue
   later", copy the link, open it in a different browser/incognito window
   — it should resume exactly where you left off.
2. Finish the form, sign, submit.
3. Open the Scheduler (preview or production, once promoted), sign in
   with a Microsoft 365 account that has access to the FreedomPrayer
   site, go to the "Intake Forms" tab, confirm the submission shows up
   under "Submitted" with the right answers and signature, and that
   "Print / Save as PDF" produces a clean printable page.
4. From the minister portal, sign in as a minister assigned to that
   recipient and confirm the "View intake form" link on their session
   card works.
