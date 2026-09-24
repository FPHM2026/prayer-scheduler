/* =========================================================================
   GRAPH CLIENT — app-only (client credentials) access to the
   IntakeResponses SharePoint list. This is the ONLY thing that lets the
   anonymous public intake form persist data: it holds its own tightly
   scoped credential (Sites.Selected, granted to just this one SharePoint
   site — see SETUP.md), separate from the FPHM Scheduler's broad delegated
   staff/minister sign-in.

   Module-level caches (token, site id, list id) persist across warm
   invocations of the same Function App instance — fine on Consumption
   plan; each just gets re-fetched on a cold start.
========================================================================= */

let cachedToken = null;
let cachedTokenExpiry = 0;

async function getAppToken() {
  if (cachedToken && Date.now() < cachedTokenExpiry - 60000) return cachedToken;
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
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

async function graphFetch(path, options = {}) {
  const token = await getAppToken();
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

let cachedSiteId = null;
let cachedListId = null;

async function resolveSiteAndList() {
  if (cachedSiteId && cachedListId) return { siteId: cachedSiteId, listId: cachedListId };
  const siteUrl = new URL(process.env.SITE_URL);
  const site = await graphFetch(`/sites/${siteUrl.hostname}:${siteUrl.pathname}`);
  cachedSiteId = site.id;
  const listName = process.env.LIST_NAME || "IntakeResponses";
  const listsResp = await graphFetch(`/sites/${cachedSiteId}/lists?$select=id,displayName`);
  const match = listsResp.value.find((l) => l.displayName === listName);
  if (!match) throw new Error(`List "${listName}" not found on site.`);
  cachedListId = match.id;
  return { siteId: cachedSiteId, listId: cachedListId };
}

// Token is our own generated crypto.randomUUID() — safe to interpolate
// directly into an OData filter (no quotes/special chars to escape).
async function findItemByToken(token) {
  const { siteId, listId } = await resolveSiteAndList();
  const resp = await graphFetch(
    `/sites/${siteId}/lists/${listId}/items?$expand=fields&$filter=fields/Token eq '${token}'`
  );
  return resp.value[0] || null;
}

async function createItem(fields) {
  const { siteId, listId } = await resolveSiteAndList();
  return graphFetch(`/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    body: JSON.stringify({ fields })
  });
}

async function updateItem(itemId, fields) {
  const { siteId, listId } = await resolveSiteAndList();
  return graphFetch(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: "PATCH",
    body: JSON.stringify(fields)
  });
}

module.exports = { graphFetch, resolveSiteAndList, findItemByToken, createItem, updateItem };
