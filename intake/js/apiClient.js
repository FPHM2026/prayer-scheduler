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
  },
  // The staff-editable question schema (see the Form Editor tab in
  // preview/index.html and the IntakeFormSchema SharePoint list). GET,
  // not POST - a pure read, no body.
  async schema() {
    const res = await fetch(FUNCTION_BASE_URL + "/intake/schema");
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(`Request to /intake/schema failed (${res.status}): ${text.slice(0, 300)}`);
    }
    return res.json();
  },
  // Checks for an already-in-progress form belonging to the same person
  // (matched on name + email + country of birth together, not name
  // alone - see the Worker handler's own comment), so the public form
  // can offer to resume that one instead of silently creating a second
  // copy when someone starts over on a different device.
  async findDuplicate(name, email, countryOfBirth, token) {
    return post("/intake/find-duplicate", { name, email, countryOfBirth, token });
  },
  // Lets the visitor delete their own form (in progress or already
  // submitted) given only their own token.
  async deleteForm(token) {
    return post("/intake/delete", { token });
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
    const err = new Error(`Request to ${path} failed (${res.status}): ${text.slice(0, 300)}`);
    // Callers (boot()'s resume-from-token path in particular) need to tell
    // "the server is unreachable, fall back to local data" apart from "the
    // server answered and this token genuinely doesn't exist any more"
    // (e.g. staff deleted an abandoned in-progress form) - only the first
    // one should fall back to offline mode; the second should start fresh.
    err.status = res.status;
    throw err;
  }
  return res.json();
}
