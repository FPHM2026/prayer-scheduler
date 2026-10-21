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

## Push notifications (added 2026-10-21)

The same Worker also sends phone notifications to ministers when they're
booked on a session. One-time setup, in this order:

1. **Create the SharePoint list.** On the FreedomPrayer site, create a
   list named exactly `PushSubscriptions` with two columns besides Title:
   - `MinisterId` (Number)
   - `SubscriptionJSON` (Multiple lines of text, **Plain text**)
   Ministers don't need any permission on it; only this Worker touches it.
2. **Make the push keys** (from this `cloudflare-worker/` folder):
   ```bash
   npx web-push generate-vapid-keys
   ```
   It prints a Public Key and a Private Key.
3. **Store them as Worker secrets:**
   ```bash
   npx wrangler secret put VAPID_PUBLIC_KEY
   npx wrangler secret put VAPID_PRIVATE_KEY
   ```
   Paste each key when asked. Keep the same keys from then on: changing
   them silently breaks every phone that already turned notifications on.
4. **Deploy:** `npx wrangler deploy`
5. **Check it:** open
   `https://fphm-intake-func.ajjamoore.workers.dev/api/push/key` in a
   browser. It should show `{"publicKey":"..."}`.
