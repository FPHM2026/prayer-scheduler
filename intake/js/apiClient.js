/* =========================================================================
   API CLIENT — talks to the Cloudflare Worker backend (cloudflare-worker/)
   which holds the app-only Graph credential and writes to the
   IntakeResponses SharePoint list. The public form never talks to
   Graph/SharePoint directly (it has no sign-in), only to these four
   endpoints. (Originally an Azure Function - switched to Cloudflare
   Workers 2026-09-25 to avoid opening a new Azure subscription; see
   cloudflare-worker/DEPLOY.md.)
========================================================================= */

const FUNCTION_BASE_URL = "https://fphm-intake-func.ajjamoore.workers.dev/api";

const FPHM_API = {
  async start(recipientName, recipientEmail) {
    return post("/intake/start", { recipientName, recipientEmail });
  },
  async load(token) {
    return post("/intake/load", { token });
  },
  async save(token, responses) {
    return post("/intake/save", { token, responses });
  },
  async submit(token, responses, signatureDataUrl) {
    return post("/intake/submit", { token, responses, signatureDataUrl });
  }
};

async function post(path, body) {
  const res = await fetch(FUNCTION_BASE_URL + path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Request to ${path} failed (${res.status}): ${text.slice(0, 300)}`);
  }
  return res.json();
}
