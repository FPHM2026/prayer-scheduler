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

   Minister push notifications (added 2026-10-21) - see the "Push
   notifications" section further down for the full picture:
     GET  /api/push/key         -> { publicKey } (VAPID, for the portal)
     POST /api/push/subscribe   { subscription } + Bearer <Graph token>
     POST /api/push/unsubscribe { endpoint }     + Bearer <Graph token>
     POST /api/push/test        {}               + Bearer <Graph token>
     POST /api/push/session     { sessionId, ministerIds?, kind?, when,
                                  recipientName? (canceled/removed only) }
                                                 + Bearer <Graph token>
     POST /api/messages/list    {}               + Bearer <Graph token>
     POST /api/messages/read    { ids?: [], unread?: true }  + Bearer <Graph token>
     POST /api/messages/delete  { ids: [] }      + Bearer <Graph token>
       (added 2026-10-21) the portal's Messages tab - every session
       notification is also saved per minister, so one missed on the
       phone (or never pushed, if they haven't turned push on) still
       shows up in the app. See "Minister messages" below.
========================================================================= */

const JSON_HEADERS = { "Content-Type": "application/json" };

function corsHeaders(env, request) {
  const allowed = (env.ALLOWED_ORIGINS || "").split(",").map((s) => s.trim()).filter(Boolean);
  const origin = request.headers.get("Origin") || "";
  const allowOrigin = allowed.includes(origin) ? origin : allowed[0] || "*";
  return {
    "Access-Control-Allow-Origin": allowOrigin,
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
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

/* =========================================================================
   PUSH NOTIFICATIONS (added 2026-10-21)

   Ministers turn on notifications once in the portal; the portal sends
   its browser push subscription here, and we keep it in the
   PushSubscriptions SharePoint list (Title = minister name, MinisterId
   Number, SubscriptionJSON plain-text multiline). When an admin books a
   session in the Scheduler, the Scheduler calls /api/push/session and we
   push to every subscribed device of the ministers on that session.

   Who is calling: every push route except /key needs the caller's own
   Microsoft Graph token (the one the Scheduler/portal already hold). We
   check it by calling Graph /me with it - Graph itself validates the
   signature - and require the token's tenant to be ours. Subscribe/
   unsubscribe/test then match that sign-in to a minister's SignInEmail,
   so nobody can register a device as somebody else. /session only ever
   pushes a real session's details to that session's own team.

   Web Push itself (RFC 8291 payload encryption + RFC 8292 VAPID) is done
   by hand with WebCrypto below - no npm dependency. VAPID_PUBLIC_KEY and
   VAPID_PRIVATE_KEY are Worker secrets (see DEPLOY.md), VAPID_SUBJECT a
   plain var in wrangler.toml.
========================================================================= */
const PUSH_LIST = "PushSubscriptions";
const te = new TextEncoder();

function b64uToBytes(s) {
  s = String(s).replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  return Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
}
function bytesToB64u(buf) {
  let s = "";
  for (const b of new Uint8Array(buf)) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function concatBytes(...parts) {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let i = 0;
  for (const p of parts) { out.set(p, i); i += p.length; }
  return out;
}
async function hkdf(salt, ikm, info, length) {
  const key = await crypto.subtle.importKey("raw", ikm, "HKDF", false, ["deriveBits"]);
  return new Uint8Array(await crypto.subtle.deriveBits({ name: "HKDF", hash: "SHA-256", salt, info }, key, length * 8));
}

// RFC 8291 "aes128gcm" encryption of one payload for one subscription.
async function encryptPushPayload(subscription, plaintext) {
  const uaPublic = b64uToBytes(subscription.keys.p256dh);
  const authSecret = b64uToBytes(subscription.keys.auth);
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const local = await crypto.subtle.generateKey({ name: "ECDH", namedCurve: "P-256" }, true, ["deriveBits"]);
  const localPublic = new Uint8Array(await crypto.subtle.exportKey("raw", local.publicKey));
  const uaKey = await crypto.subtle.importKey("raw", uaPublic, { name: "ECDH", namedCurve: "P-256" }, false, []);
  const shared = new Uint8Array(await crypto.subtle.deriveBits({ name: "ECDH", public: uaKey }, local.privateKey, 256));
  const ikm = await hkdf(authSecret, shared, concatBytes(te.encode("WebPush: info\0"), uaPublic, localPublic), 32);
  const cek = await hkdf(salt, ikm, te.encode("Content-Encoding: aes128gcm\0"), 16);
  const nonce = await hkdf(salt, ikm, te.encode("Content-Encoding: nonce\0"), 12);
  const aesKey = await crypto.subtle.importKey("raw", cek, "AES-GCM", false, ["encrypt"]);
  // One record: the payload plus the 0x02 "last record" delimiter.
  const cipher = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv: nonce }, aesKey, concatBytes(plaintext, new Uint8Array([2]))));
  const rs = new Uint8Array(4);
  new DataView(rs.buffer).setUint32(0, 4096);
  return concatBytes(salt, rs, new Uint8Array([localPublic.length]), localPublic, cipher);
}

// RFC 8292 VAPID header for one push service origin.
async function vapidAuthorization(env, endpoint) {
  const header = bytesToB64u(te.encode(JSON.stringify({ typ: "JWT", alg: "ES256" })));
  const claims = bytesToB64u(te.encode(JSON.stringify({
    aud: new URL(endpoint).origin,
    exp: Math.floor(Date.now() / 1000) + 12 * 3600,
    sub: env.VAPID_SUBJECT || "https://fphm2026.github.io/prayer-scheduler/"
  })));
  const pub = b64uToBytes(env.VAPID_PUBLIC_KEY);
  const key = await crypto.subtle.importKey("jwk", {
    kty: "EC", crv: "P-256", d: env.VAPID_PRIVATE_KEY,
    x: bytesToB64u(pub.slice(1, 33)), y: bytesToB64u(pub.slice(33, 65)), ext: true
  }, { name: "ECDSA", namedCurve: "P-256" }, false, ["sign"]);
  const sig = await crypto.subtle.sign({ name: "ECDSA", hash: "SHA-256" }, key, te.encode(header + "." + claims));
  return `vapid t=${header}.${claims}.${bytesToB64u(sig)}, k=${env.VAPID_PUBLIC_KEY}`;
}

// Returns the push service's HTTP status; 404/410 mean the subscription is
// gone for good (app removed, permission revoked) and should be deleted.
async function sendPush(env, subscription, message) {
  const body = await encryptPushPayload(subscription, te.encode(JSON.stringify(message)));
  const res = await fetch(subscription.endpoint, {
    method: "POST",
    headers: {
      Authorization: await vapidAuthorization(env, subscription.endpoint),
      "Content-Encoding": "aes128gcm",
      "Content-Type": "application/octet-stream",
      TTL: "86400",
      Urgency: "high"
    },
    body
  });
  return res.status;
}

function requirePushConfigured(env) {
  if (!env.VAPID_PUBLIC_KEY || !env.VAPID_PRIVATE_KEY) {
    return { __status: 503, error: "Push notifications aren't set up on the Worker yet (missing VAPID keys)." };
  }
  return null;
}

// The caller's identity, from their own Graph token. Graph validates the
// token; we additionally require it to come from our own tenant.
async function callerFromRequest(env, request) {
  const auth = request.headers.get("Authorization") || "";
  const token = auth.startsWith("Bearer ") ? auth.slice(7) : "";
  if (!token) return null;
  try {
    const claims = JSON.parse(atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
    if (claims.tid !== env.TENANT_ID) return null;
  } catch (e) {
    return null;
  }
  const res = await fetch("https://graph.microsoft.com/v1.0/me?$select=userPrincipalName,mail", {
    headers: { Authorization: "Bearer " + token }
  });
  if (!res.ok) return null;
  const me = await res.json();
  return { upn: (me.userPrincipalName || me.mail || "").trim().toLowerCase() };
}

async function listAllItems(env, listName) {
  const siteId = await resolveSiteId(env);
  const listId = await resolveListId(env, listName);
  let url = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=999`;
  const items = [];
  while (url) {
    const resp = await graphFetch(env, url);
    items.push(...resp.value);
    url = resp["@odata.nextLink"] ? resp["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
  }
  return items;
}

async function ministerForCaller(env, request) {
  const caller = await callerFromRequest(env, request);
  if (!caller) return { error: { __status: 401, error: "Please sign in again." } };
  const ministers = await listAllItems(env, "PrayerMinisters");
  const m = ministers.find((i) => (i.fields.SignInEmail || "").trim().toLowerCase() === caller.upn);
  if (!m) return { error: { __status: 403, error: "Your sign-in isn't on the minister roster." } };
  return { minister: { id: parseInt(m.id, 10), name: m.fields.Title || "" } };
}

async function pushSubscriptions(env) {
  const items = await listAllItems(env, PUSH_LIST);
  return items.map((i) => {
    let sub = null;
    try { sub = JSON.parse(i.fields.SubscriptionJSON || "null"); } catch (e) {}
    return { itemId: i.id, ministerId: parseInt(i.fields.MinisterId, 10), sub };
  }).filter((x) => x.sub && x.sub.endpoint && x.sub.keys);
}

async function pushListItem(env, method, itemId, fields) {
  const siteId = await resolveSiteId(env);
  const listId = await resolveListId(env, PUSH_LIST);
  if (method === "POST") {
    return graphFetch(env, `/sites/${siteId}/lists/${listId}/items`, { method, body: JSON.stringify({ fields }) });
  }
  if (method === "PATCH") {
    return graphFetch(env, `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, { method, body: JSON.stringify(fields) });
  }
  return graphFetch(env, `/sites/${siteId}/lists/${listId}/items/${itemId}`, { method: "DELETE" });
}

// Sends one message to many subscriptions, deleting any the push service
// says are gone. Returns how many were delivered.
async function pushToSubscriptions(env, subs, message) {
  let sent = 0;
  await Promise.all(subs.map(async (s) => {
    try {
      const status = await sendPush(env, s.sub, message);
      if (status >= 200 && status < 300) sent++;
      else if (status === 404 || status === 410) await pushListItem(env, "DELETE", s.itemId).catch(() => {});
    } catch (e) {
      // One bad subscription shouldn't stop the rest.
    }
  }));
  return sent;
}

async function handlePushKey(env) {
  if (!env.VAPID_PUBLIC_KEY) return { __status: 503, error: "not configured" };
  return { publicKey: env.VAPID_PUBLIC_KEY };
}

async function handlePushSubscribe(env, request) {
  const notReady = requirePushConfigured(env);
  if (notReady) return notReady;
  const body = await request.json().catch(() => ({}));
  const sub = body.subscription;
  if (!sub || !sub.endpoint || !sub.keys || !sub.keys.p256dh || !sub.keys.auth) return { __status: 400, error: "subscription required" };
  const { minister, error } = await ministerForCaller(env, request);
  if (error) return error;
  const clean = { endpoint: sub.endpoint, keys: { p256dh: sub.keys.p256dh, auth: sub.keys.auth } };
  const fields = { Title: minister.name, MinisterId: minister.id, SubscriptionJSON: JSON.stringify(clean) };
  const existing = (await pushSubscriptions(env)).find((s) => s.sub.endpoint === sub.endpoint);
  if (existing) await pushListItem(env, "PATCH", existing.itemId, fields);
  else await pushListItem(env, "POST", null, fields);
  return { ok: true };
}

async function handlePushUnsubscribe(env, request) {
  const body = await request.json().catch(() => ({}));
  if (!body.endpoint) return { __status: 400, error: "endpoint required" };
  const { minister, error } = await ministerForCaller(env, request);
  if (error) return error;
  const mine = (await pushSubscriptions(env)).filter((s) => s.sub.endpoint === body.endpoint && s.ministerId === minister.id);
  for (const s of mine) await pushListItem(env, "DELETE", s.itemId);
  return { ok: true };
}

async function handlePushTest(env, request) {
  const notReady = requirePushConfigured(env);
  if (notReady) return notReady;
  const { minister, error } = await ministerForCaller(env, request);
  if (error) return error;
  const subs = (await pushSubscriptions(env)).filter((s) => s.ministerId === minister.id);
  const sent = await pushToSubscriptions(env, subs, {
    title: "Notifications are working",
    body: "You'll get a notification like this when you're booked for a prayer session.",
    tag: "fphm-test"
  });
  return { ok: true, sent };
}

// Called by the Scheduler after a session is saved or removed.
//   kind "booked"/"changed" (default): recipient and team come from the
//     saved session in SharePoint, not the caller; ministerIds optionally
//     narrows it to some of that team.
//   kind "canceled" (session removed) / "removed" (minister taken off the
//     team): the session no longer holds that team (or no longer exists),
//     so the caller supplies ministerIds and recipientName. Only an admin
//     can call this - a signed-in minister is refused - so a minister
//     can't send fake cancellations to others.
// `when` is the admin's own browser-formatted date/time (the Worker runs
// in UTC and doesn't know the church's time zone), trimmed to a sane length.
const SESSION_TITLES = {
  booked: "You've been booked for a prayer session",
  changed: "Prayer session time changed",
  canceled: "Prayer session canceled",
  removed: "You've been taken off a prayer session"
};
async function handlePushSession(env, request) {
  const notReady = requirePushConfigured(env);
  if (notReady) return notReady;
  const caller = await callerFromRequest(env, request);
  if (!caller) return { __status: 401, error: "Please sign in again." };
  const ministers = await listAllItems(env, "PrayerMinisters");
  if (ministers.some((m) => (m.fields.SignInEmail || "").trim().toLowerCase() === caller.upn)) {
    return { __status: 403, error: "Only the scheduling admin can send session notifications." };
  }
  const body = await request.json().catch(() => ({}));
  const kind = SESSION_TITLES[body.kind] ? body.kind : "booked";
  const sessionId = parseInt(body.sessionId, 10) || null;
  const requested = Array.isArray(body.ministerIds) ? body.ministerIds.map((x) => parseInt(x, 10)).filter((x) => !isNaN(x)) : null;
  let targets, recipient;
  if (kind === "canceled" || kind === "removed") {
    const known = new Set(ministers.map((m) => parseInt(m.id, 10)));
    targets = new Set((requested || []).filter((id) => known.has(id)));
    recipient = String(body.recipientName || "a recipient").slice(0, 80);
  } else {
    if (!sessionId) return { __status: 400, error: "sessionId required" };
    const siteId = await resolveSiteId(env);
    const listId = await resolveListId(env, "PrayerSessions");
    const item = await graphFetch(env, `/sites/${siteId}/lists/${listId}/items/${sessionId}?$expand=fields`);
    const f = item.fields || {};
    if (f.Status !== "Scheduled" || !f.SessionDate) return { ok: true, sent: 0, skipped: "not a scheduled session" };
    const assigned = String(f.AssignedMinisterIDs || "").replace(/<[^>]*>/g, "").split(",").map((x) => parseInt(x.trim(), 10)).filter((x) => !isNaN(x));
    const only = requested ? new Set(requested) : null;
    targets = new Set(assigned.filter((id) => !only || only.has(id)));
    recipient = String(f.RecipientName || "a recipient").slice(0, 80);
  }
  if (!targets.size) return { ok: true, sent: 0 };
  const when = String(body.when || "").slice(0, 80);
  const message = {
    title: SESSION_TITLES[kind],
    body: when ? `${recipient} · ${when}` : recipient,
    tag: "fphm-session-" + (sessionId || "x") + "-" + kind
  };
  // Saved for every minister on the list, pushed or not. A missing
  // MinisterMessages list must never stop the push itself.
  await Promise.all([...targets].map((ministerId) =>
    saveMinisterMessage(env, ministerId, message.title, message.body, sessionId).catch(() => {})
  ));
  const subs = (await pushSubscriptions(env)).filter((s) => targets.has(s.ministerId));
  if (!subs.length) return { ok: true, sent: 0 };
  const sent = await pushToSubscriptions(env, subs, message);
  return { ok: true, sent };
}

/* =========================================================================
   MINISTER MESSAGES (added 2026-10-21)

   The portal's Messages tab. Each session notification is also stored in
   the MinisterMessages SharePoint list (Title = message title, MinisterId
   Number, Body plain-text multiline, SessionId Number, IsRead Yes/No) -
   only this Worker reads or writes it, so ministers need no permission on
   it. A minister only ever sees and marks their own, matched from their
   Graph token the same way as the push routes.
========================================================================= */
const MESSAGES_LIST = "MinisterMessages";

async function saveMinisterMessage(env, ministerId, title, body, sessionId) {
  const siteId = await resolveSiteId(env);
  const listId = await resolveListId(env, MESSAGES_LIST);
  return graphFetch(env, `/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    body: JSON.stringify({ fields: { Title: title, MinisterId: ministerId, Body: body, SessionId: sessionId || null, IsRead: false } })
  });
}

async function myMessageItems(env, ministerId) {
  const siteId = await resolveSiteId(env);
  const listId = await resolveListId(env, MESSAGES_LIST);
  const resp = await graphFetch(
    env,
    `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=999&$filter=fields/MinisterId eq ${parseInt(ministerId, 10)}`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  return resp.value;
}

async function handleMessagesList(env, request) {
  const { minister, error } = await ministerForCaller(env, request);
  if (error) return error;
  const items = await myMessageItems(env, minister.id);
  const messages = items.map((i) => ({
    id: parseInt(i.id, 10),
    title: i.fields.Title || "",
    body: i.fields.Body || "",
    sessionId: i.fields.SessionId ? parseInt(i.fields.SessionId, 10) : null,
    read: !!i.fields.IsRead,
    created: i.createdDateTime
  })).sort((a, b) => String(b.created).localeCompare(String(a.created))).slice(0, 100);
  return { messages };
}

// Marks the caller's own messages read: the given ids, or all unread ones.
async function handleMessagesRead(env, request) {
  const body = await request.json().catch(() => ({}));
  const { minister, error } = await ministerForCaller(env, request);
  if (error) return error;
  // unread: true flips messages back to unread (only for the ids given).
  const markRead = !body.unread;
  const wanted = Array.isArray(body.ids) ? new Set(body.ids.map((x) => parseInt(x, 10))) : null;
  if (!markRead && !wanted) return { __status: 400, error: "ids required" };
  const items = (await myMessageItems(env, minister.id))
    .filter((i) => !!i.fields.IsRead !== markRead && (!wanted || wanted.has(parseInt(i.id, 10))));
  const siteId = await resolveSiteId(env);
  const listId = await resolveListId(env, MESSAGES_LIST);
  await Promise.all(items.map((i) =>
    graphFetch(env, `/sites/${siteId}/lists/${listId}/items/${i.id}/fields`, { method: "PATCH", body: JSON.stringify({ IsRead: markRead }) })
  ));
  return { ok: true, marked: items.length };
}

// Deletes only the caller's own messages; ids belonging to anyone else are ignored.
async function handleMessagesDelete(env, request) {
  const body = await request.json().catch(() => ({}));
  const { minister, error } = await ministerForCaller(env, request);
  if (error) return error;
  if (!Array.isArray(body.ids) || !body.ids.length) return { __status: 400, error: "ids required" };
  const wanted = new Set(body.ids.map((x) => parseInt(x, 10)));
  const items = (await myMessageItems(env, minister.id)).filter((i) => wanted.has(parseInt(i.id, 10)));
  const siteId = await resolveSiteId(env);
  const listId = await resolveListId(env, MESSAGES_LIST);
  await Promise.all(items.map((i) =>
    graphFetch(env, `/sites/${siteId}/lists/${listId}/items/${i.id}`, { method: "DELETE" })
  ));
  return { ok: true, deleted: items.length };
}

const ROUTES = {
  "/api/intake/start": handleStart,
  "/api/intake/load": handleLoad,
  "/api/intake/save": handleSave,
  "/api/intake/submit": handleSubmit,
  "/api/intake/schema": handleSchema,
  "/api/intake/find-duplicate": handleFindDuplicate,
  "/api/intake/delete": handleDelete,
  "/api/push/key": handlePushKey,
  "/api/push/subscribe": handlePushSubscribe,
  "/api/push/unsubscribe": handlePushUnsubscribe,
  "/api/push/test": handlePushTest,
  "/api/push/session": handlePushSession,
  "/api/messages/list": handleMessagesList,
  "/api/messages/read": handleMessagesRead,
  "/api/messages/delete": handleMessagesDelete
};
const GET_ROUTES = new Set(["/api/intake/schema", "/api/push/key"]);

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
