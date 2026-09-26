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

   Routes:
     POST /api/intake/start          { recipientName?, recipientEmail? } -> { token }
     POST /api/intake/load           { token } -> { status, responses, recipientName, submittedAt }
     POST /api/intake/save           { token, responses } -> { ok: true }
     POST /api/intake/submit         { token, responses, signatureDataUrl } -> { ok: true }
     GET  /api/intake/schema         -> the live staff-editable question schema
     POST /api/intake/find-duplicate { name, email, countryOfBirth, token } ->
       { found, token? } - added 2026-09-26 so the public form can offer
       "resume that one" when the same person starts a second copy (e.g.
       on a different device); see the handler's own comment for why it
       requires all three fields to match, not just name.
     POST /api/intake/delete         { token } -> { ok: true } - added
       2026-09-26 so a visitor can delete their own form given only their
       own token; InProgress only, 403 once Submitted (see the handler's
       own comment).
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

async function resolveSiteId(env) {
  if (cachedSiteId) return cachedSiteId;
  const siteUrl = new URL(env.SITE_URL);
  const site = await graphFetch(env, `/sites/${siteUrl.hostname}:${siteUrl.pathname}`);
  cachedSiteId = site.id;
  return cachedSiteId;
}

// listId cache keyed by display name, so the Worker can talk to more than
// one list (IntakeResponses for responses, IntakeFormSchema for the
// staff-editable question schema) without re-resolving the site each time.
const listIdCache = {};
async function resolveListId(env, listName) {
  if (listIdCache[listName]) return listIdCache[listName];
  const siteId = await resolveSiteId(env);
  const listsResp = await graphFetch(env, `/sites/${siteId}/lists?$select=id,displayName`);
  const match = listsResp.value.find((l) => l.displayName === listName);
  if (!match) throw new Error(`List "${listName}" not found on site.`);
  listIdCache[listName] = match.id;
  return match.id;
}

async function resolveSiteAndList(env) {
  const siteId = await resolveSiteId(env);
  const listId = await resolveListId(env, env.LIST_NAME || "IntakeResponses");
  return { siteId, listId };
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

async function deleteItem(env, itemId) {
  const { siteId, listId } = await resolveSiteAndList(env);
  return graphFetch(env, `/sites/${siteId}/lists/${listId}/items/${itemId}`, {
    method: "DELETE"
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
    submittedAt: item.fields.SubmittedAt || null,
    // Added 2026-09-26 so a visitor revisiting their resume link after
    // submitting (no in-memory copy left from the submit flow itself)
    // can still see/print their signature, not just their answers.
    signatureDataUrl: item.fields.SignatureDataUrl || ""
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

// Lets the visitor delete their own form, given only their own token -
// the same trust boundary load/save/submit already use (the token is a
// random UUID only the visitor and whoever they shared their resume link
// with would know; see CLAUDE.md's Intake Forms section on this data's
// sensitivity). InProgress only, deliberately - once a form is
// Submitted, staff/ministers may already be relying on it for a
// scheduled session, so the recipient shouldn't be able to pull it out
// from under them. Enforced here, not just by hiding the delete button
// once Submitted client-side, so it can't be bypassed by calling this
// endpoint directly.
async function handleDelete(env, request) {
  const body = await request.json().catch(() => ({}));
  if (!body.token) return { __status: 400, error: "token required" };
  const item = await findItemByToken(env, body.token);
  if (!item) return { __status: 404, error: "not found" };
  if (item.fields.Status === "Submitted") {
    return { __status: 403, error: "This form has already been submitted and can no longer be deleted." };
  }
  await deleteItem(env, item.id);
  return { ok: true };
}

// Checks for an already-InProgress intake response belonging to the same
// person, so the public form can offer "resume that one" instead of
// silently letting them start a second copy (the common case: someone
// starts on one device, comes back on another without their saved
// resume link). Requires name + email + country of birth to ALL match
// an existing InProgress item (not just name) before handing back its
// token - an explicit, deliberately-chosen tradeoff given this data's
// sensitivity (see CLAUDE.md's Intake Forms section): name alone is
// easy to guess/share and would let a stranger hijack someone else's
// in-progress form; this three-factor match is a much smaller target.
async function handleFindDuplicate(env, request) {
  const body = await request.json().catch(() => ({}));
  const { name, email, countryOfBirth, token } = body;
  if (!name || !email || !countryOfBirth) return { __status: 400, error: "name, email, countryOfBirth required" };
  const { siteId, listId } = await resolveSiteAndList(env);
  const resp = await graphFetch(
    env,
    `/sites/${siteId}/lists/${listId}/items?$expand=fields&$filter=fields/Status eq 'InProgress'`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  const normName = name.trim().toLowerCase();
  const normEmail = email.trim().toLowerCase();
  const normCountry = countryOfBirth.trim().toLowerCase();
  for (const item of resp.value) {
    const f = item.fields;
    if (f.Token === token) continue; // never match the caller's own in-progress item
    if ((f.RecipientName || "").trim().toLowerCase() !== normName) continue;
    if ((f.RecipientEmail || "").trim().toLowerCase() !== normEmail) continue;
    let responses = {};
    try {
      responses = JSON.parse(f.ResponsesJSON || "{}");
    } catch (e) {
      continue;
    }
    if ((responses.countryOfBirth || "").trim().toLowerCase() !== normCountry) continue;
    return { found: true, token: f.Token };
  }
  return { found: false };
}

// Read-only: the current staff-editable question schema, for the public
// form to render from instead of its own hardcoded default. Cached
// in-memory per warm isolate (same pattern as the Graph token/site/list
// caches above) - a schema edit takes effect for new isolates immediately,
// and for warm ones within a few minutes as they naturally recycle.
let cachedSchema = null;
let cachedSchemaExpiry = 0;
async function handleSchema(env) {
  if (cachedSchema && Date.now() < cachedSchemaExpiry) return cachedSchema;
  const listId = await resolveListId(env, "IntakeFormSchema");
  const siteId = await resolveSiteId(env);
  const resp = await graphFetch(env, `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=1`);
  const item = resp.value[0];
  if (!item) return { __status: 404, error: "no schema found" };
  let schema;
  try {
    schema = JSON.parse(item.fields.SchemaJSON);
  } catch (e) {
    return { __status: 500, error: "stored schema is not valid JSON" };
  }
  cachedSchema = schema;
  cachedSchemaExpiry = Date.now() + 5 * 60 * 1000; // 5 minutes
  return schema;
}

const ROUTES = {
  "/api/intake/start": handleStart,
  "/api/intake/load": handleLoad,
  "/api/intake/save": handleSave,
  "/api/intake/submit": handleSubmit,
  "/api/intake/schema": handleSchema,
  "/api/intake/find-duplicate": handleFindDuplicate,
  "/api/intake/delete": handleDelete
};
const GET_ROUTES = new Set(["/api/intake/schema"]);

export default {
  async fetch(request, env, ctx) {
    const cors = corsHeaders(env, request);

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: cors });
    }

    const url = new URL(request.url);
    const handler = ROUTES[url.pathname];
    if (!handler) return jsonResponse({ error: "not found" }, 404, cors);
    const allowedMethod = GET_ROUTES.has(url.pathname) ? "GET" : "POST";
    if (request.method !== allowedMethod) return jsonResponse({ error: "method not allowed" }, 405, cors);

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
