/* =========================================================================
   FPHM INTAKE FUNCTION — Cloudflare Worker

   Ported from the original Azure Function (see the standalone "FPHM
   Intake Form" project's azure-function/ for the pre-port version, kept
   for reference) after deciding not to open a new Azure subscription
   just for this. Same job: hold the app-only Graph credential
   (Sites.Selected, granted to just the FreedomPrayer site - see that
   project's SETUP.md steps 2-3, still accurate) so the anonymous public
   intake form can create/update/submit its SharePoint list item without
   any staff/minister credential ever touching a browser the public can
   reach.

   Four routes, same shape as before:
     POST /api/intake/start   { recipientName?, recipientEmail? } -> { token }
     POST /api/intake/load    { token } -> { status, responses, recipientName, submittedAt }
     POST /api/intake/save    { token, responses } -> { ok: true }
     POST /api/intake/submit  { token, responses, signatureDataUrl } -> { ok: true }
========================================================================= */

const JSON_HEADERS = { "Content-Type": "application/json" };

function corsHeaders(env, request) {
  const allowed = (env.ALLOWED_ORIGINS || "").split(",").map((s) => s.trim()).filter(Boolean);
  const origin = request.headers.get("Origin") || "";
  const allowOrigin = allowed.includes(origin) ? origin : allowed[0] || "*";
  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Max-Age": "86400"
  };
}

function jsonResponse(body, status, extraHeaders) {
  return new Response(JSON.stringify(body), { status, headers: { ...JSON_HEADERS, ...extraHeaders } });
}

// ---- Graph client (app-only, client credentials) ------------------------
// Module-level cache: persists across requests handled by the same warm
// isolate (Workers reuse isolates similarly to how the old Azure Function
// reused warm instances), re-fetched on a cold start otherwise.
let cachedToken = null;
let cachedTokenExpiry = 0;
let cachedSiteId = null;
let cachedListId = null;

async function getAppToken(env) {
  if (cachedToken && Date.now() < cachedTokenExpiry - 60000) return cachedToken;
  const res = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env.CLIENT_ID,
      client_secret: env.CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    })
  });
  if (!res.ok) throw new Error("Failed to get app token: " + (await res.text()));
  const data = await res.json();
  cachedToken = data.access_token;
  cachedTokenExpiry = Date.now() + data.expires_in * 1000;
  return cachedToken;
}

async function graphFetch(env, path, options = {}) {
  const token = await getAppToken(env);
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    ...options,
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json", ...(options.headers || {}) }
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph ${options.method || "GET"} ${path} failed (${res.status}): ${text.slice(0, 500)}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function resolveSiteAndList(env) {
  if (cachedSiteId && cachedListId) return { siteId: cachedSiteId, listId: cachedListId };
  const siteUrl = new URL(env.SITE_URL);
  const site = await graphFetch(env, `/sites/${siteUrl.hostname}:${siteUrl.pathname}`);
  cachedSiteId = site.id;
  const listName = env.LIST_NAME || "IntakeResponses";
  const listsResp = await graphFetch(env, `/sites/${cachedSiteId}/lists?$select=id,displayName`);
  const match = listsResp.value.find((l) => l.displayName === listName);
  if (!match) throw new Error(`List "${listName}" not found on site.`);
  cachedListId = match.id;
  return { siteId: cachedSiteId, listId: cachedListId };
}

// Token is our own generated crypto.randomUUID() - safe to interpolate
// directly into an OData filter (no quotes/special chars to escape).
async function findItemByToken(env, token) {
  const { siteId, listId } = await resolveSiteAndList(env);
  // Token isn't an indexed SharePoint column, so Graph refuses to filter on
  // it (400 invalidRequest) unless explicitly told the query might be slow
  // on a large list - fine at this list's size (a church intake form, not
  // millions of rows).
  const resp = await graphFetch(
    env,
    `/sites/${siteId}/lists/${listId}/items?$expand=fields&$filter=fields/Token eq '${token}'`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  return resp.value[0] || null;
}

async function createItem(env, fields) {
  const { siteId, listId } = await resolveSiteAndList(env);
  return graphFetch(env, `/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    body: JSON.stringify({ fields })
  });
}

async function updateItem(env, itemId, fields) {
  const { siteId, listId } = await resolveSiteAndList(env);
  return graphFetch(env, `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fields)
  });
}

// ---- route handlers -----------------------------------------------------
async function handleStart(env, request) {
  const body = await request.json().catch(() => ({}));
  const token = crypto.randomUUID();
  const fields = {
    Title: body.recipientName || "New intake",
    Token: token,
    RecipientName: body.recipientName || "",
    RecipientEmail: body.recipientEmail || "",
    Status: "InProgress",
    ResponsesJSON: "{}"
  };
  await createItem(env, fields);
  return { token };
}

async function handleLoad(env, request) {
  const body = await request.json();
  if (!body.token) return { __status: 400, error: "token required" };
  const item = await findItemByToken(env, body.token);
  if (!item) return { __status: 404, error: "not found" };
  let responses = {};
  try {
    responses = JSON.parse(item.fields.ResponsesJSON || "{}");
  } catch (e) {}
  return {
    status: item.fields.Status || "InProgress",
    responses,
    recipientName: item.fields.RecipientName || "",
    submittedAt: item.fields.SubmittedAt || null
  };
}

async function handleSave(env, request) {
  const body = await request.json();
  if (!body.token) return { __status: 400, error: "token required" };
  const item = await findItemByToken(env, body.token);
  if (!item) return { __status: 404, error: "not found" };
  const responses = body.responses || {};
  const fields = { ResponsesJSON: JSON.stringify(responses) };
  if (responses.name) {
    fields.RecipientName = String(responses.name);
    fields.Title = String(responses.name);
  }
  if (responses.email) fields.RecipientEmail = String(responses.email);
  await updateItem(env, item.id, fields);
  return { ok: true };
}

async function handleSubmit(env, request) {
  const body = await request.json();
  if (!body.token) return { __status: 400, error: "token required" };
  const item = await findItemByToken(env, body.token);
  if (!item) return { __status: 404, error: "not found" };
  const responses = body.responses || {};
  const fields = {
    ResponsesJSON: JSON.stringify(responses),
    SignatureDataUrl: body.signatureDataUrl || "",
    Status: "Submitted",
    SubmittedAt: new Date().toISOString()
  };
  if (responses.name) {
    fields.RecipientName = String(responses.name);
    fields.Title = String(responses.name);
  }
  if (responses.email) fields.RecipientEmail = String(responses.email);
  await updateItem(env, item.id, fields);
  return { ok: true };
}

const ROUTES = {
  "/api/intake/start": handleStart,
  "/api/intake/load": handleLoad,
  "/api/intake/save": handleSave,
  "/api/intake/submit": handleSubmit
};

export default {
  async fetch(request, env, ctx) {
    const cors = corsHeaders(env, request);

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: cors });
    }

    const url = new URL(request.url);
    const handler = ROUTES[url.pathname];
    if (!handler) return jsonResponse({ error: "not found" }, 404, cors);
    if (request.method !== "POST") return jsonResponse({ error: "method not allowed" }, 405, cors);

    try {
      const result = await handler(env, request);
      const status = result.__status || 200;
      if (result.__status) delete result.__status;
      return jsonResponse(result, status, cors);
    } catch (e) {
      return jsonResponse({ error: e.message }, 500, cors);
    }
  }
};
