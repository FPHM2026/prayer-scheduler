# Deploying the Intake Worker

Replaces the earlier Azure Function plan (see `azure-function/` in this
repo's history for the abandoned version) — same job, no new billing
required since you already have a Cloudflare account (the
`fphm-whatsapp-notify` Worker uses the same pattern).

## 1. Install Wrangler (Cloudflare's CLI) and sign in

From this `cloudflare-worker/` folder:
```bash
npm install
npx wrangler login
```
This opens a browser to authorize Wrangler against your existing
Cloudflare account — the same one the WhatsApp Worker already lives in.

## 2. Set the client secret (kept out of source control)

```bash
npx wrangler secret put CLIENT_SECRET
```
When prompted, paste the client secret **value** for the "FPHM Intake
Form" Azure AD app registration (Client ID
`5580dcc8-7837-48fe-a0cf-92eda3959cb0`) — the same one already set up for
`Sites.Selected` access to the FreedomPrayer SharePoint site. If that
secret has expired by the time you deploy, generate a new one in the
Entra admin center first (App registrations → FPHM Intake Form →
Certificates & secrets).

Everything else the Worker needs (`TENANT_ID`, `CLIENT_ID`, `SITE_URL`,
`LIST_NAME`, `ALLOWED_ORIGINS`) is already in `wrangler.toml` as plain
`[vars]` — nothing secret in any of those.

## 3. Deploy

```bash
npx wrangler deploy
```
This prints the Worker's live URL, something like:
```
https://fphm-intake-func.<your-subdomain>.workers.dev
```

## 4. Point the public form at it

Edit `../intake/js/apiClient.js` in this repo, set:
```js
const FUNCTION_BASE_URL = "https://fphm-intake-func.<your-subdomain>.workers.dev/api";
```
(note the trailing `/api` — the Worker's routes are `/api/intake/start`
etc., matching what `apiClient.js` already calls). Commit and push.

## 5. Verify

```bash
curl -X POST https://fphm-intake-func.<your-subdomain>.workers.dev/api/intake/start -H "Content-Type: application/json" -d "{}"
```
Should return `{"token":"..."}`. Check the `IntakeResponses` SharePoint
list for a new "New intake" item.
