# Prayer Ministry Scheduler — Project Context

## What this is
A scheduling app for a church prayer ministry team, coordinating one-on-one prayer
sessions between external recipients and internal prayer ministers. Built as a
static HTML/JS single-page app, hosted on GitHub Pages, backed by SharePoint Lists
via Microsoft Graph API (not Power Apps, not raw SharePoint hosting — both were
tried and abandoned; see "History" below).

## Architecture
- **Frontend**: single self-contained `index.html` (vanilla JS, no build step, no
  framework). All UI, styling, and logic live in this one file.
- **Hosting**: GitHub Pages, org repo `FPHM2026/prayer-scheduler`, deployed from
  the `main` branch root. Live at `https://FPHM2026.github.io/prayer-scheduler/`.
- **Auth**: MSAL.js (`@azure/msal-browser`, loaded via jsDelivr CDN,
  `cacheLocation: "sessionStorage"`) — admins sign in with their normal Microsoft
  365 account via a popup.
- **Data**: Microsoft Graph API (`https://graph.microsoft.com/v1.0`), talking
  directly to SharePoint Lists on the church's site. Not the older SharePoint REST
  API (`_api/web/...`) — that doesn't support CORS from an external origin; Graph
  does, via a registered Azure AD app.

## Config values already in the code (in `index.html`, top of the `<script>`)
- Azure AD Client ID: `549f5207-d7f8-4924-9bde-30532d90c1d2`
- Azure AD Tenant ID: `b9e447ab-c42c-4b04-92c8-d629b0ce320d`
- SharePoint site URL (hardcoded, `SITE_URL` constant):
  `https://creeksidechurch.sharepoint.com/sites/FreedomPrayer`
- Redirect URI registered on the Azure AD app: must exactly match the GitHub
  Pages URL above (trailing slash included).
- Graph API permission granted: `Sites.ReadWrite.All` (delegated, admin consent
  already granted for the org).

## SharePoint Lists (the actual data store)
- **PrayerMinisters**: Title (name), Email, Phone, Status (Choice: Active/Inactive)
- **PrayerSessions**: Title, SessionDate, SessionEndDate, RecipientName,
  RecipientContact, AssignedMinisterIDs (comma-separated minister IDs, plain
  text), LeadMinisterIDs (comma-separated subset of the above who are "Lead"
  rather than "Support"), LocationName (plain text, not a true Lookup column —
  written as plain text matching a Locations list Title, chosen deliberately to
  avoid Graph's LookupId complexity), Notes (plain text — must be Plain text
  format, not Rich text, or it stores HTML), ApptType (Choice: First/Follow-up),
  Status (Choice: Scheduled/Completed/**Waiting**), PreviousSessionId (Number,
  links a follow-up session back to the one it followed), WaCreated (Yes/No),
  WaLink (plain text), **Priority (Yes/No, default No)** — flags a Waiting entry
  to the top of the Planning tab regardless of how long they've been waiting.
  A Waiting-status item has SessionDate/SessionEndDate/LocationName/
  AssignedMinisterIDs/LeadMinisterIDs blank until it's actually scheduled.
  "Days waiting" is computed client-side from the item's own SharePoint
  `createdDateTime` metadata — there is no dedicated DateAdded field for this.
- **BlackoutDates**: Title (minister name), BlackoutDate (Date), **EndDate
  (Date, optional)** — blank/null means a single-day blackout; set means an
  inclusive date range. Notes (Plain text)
- **Locations**: Title (location name) — deliberately simple, just a name
- **Prospects**: retired. The app no longer reads this list at all — replaced
  by PrayerSessions items with Status="Waiting" (see above). The list itself
  still exists in SharePoint with whatever old data was in it; run
  `FPHM Scheduler Import Prep/migrate-prospects-to-waiting.console.js` once
  (same browser-console pattern as the historical session import) to copy any
  remaining entries over as Waiting sessions, then archive/delete the
  Prospects list yourself whenever you're ready. Not done automatically —
  Claude has no SharePoint write access outside a session you're driving.

## Features implemented
Every list-style tab (Schedule, Completed Sessions, Minister Overview,
Blackout Dates) is grouped into collapsible accordion cards by recipient or
minister, collapsed by default, with a "Session #N" chip computed once per
recipient across their full chronological history (any status, including a
Waiting entry) — not per-view. A shared bottom-sheet modal pattern (full
width, slides up from the bottom, capped at 90vh) is used for every form
(New Session, Minister, Blackout Date); the confirm dialog and Add Location
stay as small centered modals. Main nav is a sticky, single-row,
horizontally-scrolling tab strip on mobile/tablet; desktop restores the
fuller header and lets tabs wrap.

- Schedule tab: grouped by recipient, search, Upcoming/All/Past filter,
  soonest-upcoming-first sort
- New/Edit/Follow-up session form: recipient info, date + day-of-week label,
  quick-select start times (Sun 4:00/5:00/5:30pm, Tue 6:00pm — curated presets,
  not auto-generated from a window), custom start/end time (hour/half-hour only,
  validated both via `step="1800"` and an explicit save-time check), location
  dropdown with inline "add new location", appointment type (First/Follow-up,
  smart-defaulted), status (Scheduled/Completed — hidden on new sessions, shown
  when editing), minister multi-select (2+ required) with Lead/Support role
  toggle per minister, live "also on this date" panel, "last session with this
  recipient" hint (name-match based — recipients aren't persistent records),
  per-minister upcoming-session-count badge, blackout-date flagging,
  minister-specific double-booking prevention (two sessions CAN overlap in time
  if they share no ministers — only blocks if a specific minister would be in
  two places at once), notes, WhatsApp group-created toggle + optional invite
  link. The same modal doubles as the Waiting-list form (see Planning below) —
  a `sessKind` flag ('session' vs 'waiting') shows/hides the scheduling-only
  fields rather than being a separate form.
- Completed Sessions tab: grouped by recipient, read-only, search, plus a
  collapsible stats block (YTD count, same-period-last-year delta, unique
  recipients, first-time vs. returning split, monthly trend chart for the
  current year)
- Planning tab: PrayerSessions items with Status="Waiting" (no separate
  Prospects list — see above), sortable by longest-waiting or highest session
  number, priority-flagged entries always pinned to the top, flag/schedule/
  edit/delete actions per entry. "Schedule Session" reopens the same session
  modal with the scheduling fields revealed, converting the same record.
- Minister Overview tab: grouped by minister, Active/Inactive sub-tabs (Active
  hides anyone with zero sessions in the selected range), search, date range
  defaulting to year-to-date, collapsible stats block (hours this range vs.
  the same calendar range last year, active ministers serving, avg hours per
  minister), per-minister upcoming sessions
- Prayer Ministers tab: search, Active/Inactive sub-tabs, one-click A→Z/Z→A
  sort toggle, activate/deactivate/edit/delete per minister
- Blackout Dates tab: grouped by minister, soonest first, Add/Edit modal with
  a single-date/date-range toggle (writes BlackoutDate + EndDate), edit/delete
  per entry
- Sign Out button (added after a multi-user bug — see "Known gotchas")
- SVG icons instead of emoji for actions, keyboard-accessible minister
  selection (tabindex + Enter/Space), visible focus rings

## Minister portal (`/portal/`)
A second, deliberately small single-file app (`portal/index.html`) for prayer
ministers themselves — "my upcoming sessions" (read-only) and "my time off"
(add/edit/delete their own BlackoutDates entries). Supersedes the earlier
"Forms + Power Automate, no accounts" idea below — instead, each minister
gets their own **unlicensed** Entra ID account (a directory identity, no
M365 seat cost) so they can sign in for real.

- **Same Azure AD app registration** as the main scheduler (same Client ID) —
  a second registration isn't needed. Delegated Graph permissions only ever
  allow what the *signed-in user's own SharePoint permissions* allow, so a
  minister account with restricted SharePoint access genuinely can't read or
  write more than it's been granted, even though the app itself still
  requests the same broad `Sites.ReadWrite.All`. The real access boundary is
  SharePoint's own permissions on a "Ministers" group, granted directly in
  SharePoint (not in this code, not in Entra):
  - PrayerMinisters: Read (lets the portal resolve "which minister is this"
    by matching sign-in email, and show co-minister names on a session)
  - PrayerSessions: Read (SharePoint can't restrict "only sessions you're
    assigned to" — item-level permissions key off who *created* the item,
    and staff create every session, so ministers can technically read the
    full upcoming schedule; the portal's own UI just filters to theirs)
  - BlackoutDates: Contribute, with **item-level permissions** set to "Read
    items created by the user" / "Create and edit items created by the
    user" — this one *is* genuinely enforced, since each minister is the
    creator of their own time-off request
  - Locations: no access needed (LocationName is already plain text on the
    session item)
- Needs its own redirect URI on the same app registration:
  `https://FPHM2026.github.io/prayer-scheduler/portal/`
- **Identity matching**: the signed-in account's email (`account.username`)
  must exactly match (case-insensitive) that minister's Email field in
  PrayerMinisters, or the portal shows a "couldn't find your profile"
  message instead of a broken page. Keep that field accurate when adding a
  new minister account.
- Creating the actual Entra accounts (Microsoft 365 admin center → Add a
  user → "Create user without product license") and the SharePoint group/
  permission setup above are manual admin steps outside this repo — Claude
  has no tenant admin access to do them.

## Not yet built
- **Calendar (month grid) view** — was planned but never built in this HTML
  version. The agenda/Schedule view covers the "chronological list" requirement
  on its own; the calendar grid is a nice-to-have, not yet started.

## Known gotchas (hard-won, don't reintroduce these bugs)
1. **Graph list item IDs are strings, not numbers.** Comma-separated ID fields
   (AssignedMinisterIDs etc.) get parsed with `parseInt`. If you load a list's
   raw `id` without also converting it to a number, you get silent `===`
   mismatches everywhere (checkboxes that never show as checked, blank IDs on
   save). Fix already applied: `listItems()` does
   `{ ...i.fields, id: parseInt(i.id, 10) }` — note `id` comes *after* the
   fields spread, because Graph's `fields` object sometimes has its own stray
   `id`-like key that would otherwise silently win.
2. **MSAL v3 requires `await msalInstance.initialize()`** before any other MSAL
   call, or you get `uninitialized_public_client_application`. Already handled
   via a cached `msalReady` promise awaited at the top of `init()` and the
   sign-in handler.
3. **`interaction_in_progress` errors** happen when a previous sign-in attempt
   didn't complete cleanly (closed popup, page reload mid-flow) and MSAL's
   internal flag gets stuck — worse with `localStorage` since it persists
   across sessions/users on a shared device. Fixed by using `sessionStorage`
   instead, plus auto-clearing the stuck flag and prompting a retry if it
   happens again.
4. **`wireApp()` must only run once.** It attaches all click handlers; if
   called twice (e.g. after a sign-out/sign-in cycle) every handler fires
   twice. Guarded with an `appWired` flag.
5. **SharePoint "Multiple lines of text" columns default to Rich Text**, which
   stores values as raw HTML — shows up as literal `<div class="ExternalClass...">`
   garbage in the app. Always set these to Plain Text in SharePoint's column
   settings (Notes fields on PrayerSessions, BlackoutDates, Prospects all hit
   this).
6. **SharePoint Online forces `.html` files in a document library to download
   instead of rendering**, and blocks custom script on modern sites by default.
   This is why the app is hosted on GitHub Pages, not inside SharePoint itself.
   Don't try to move it back into a SharePoint doc library — it doesn't work,
   full stop, regardless of what's inside the file.
7. **Azure AD admin consent** for `Sites.ReadWrite.All` needs a Global
   Administrator or similarly privileged role — a regular user can register the
   app itself but can't grant tenant-wide consent. Already granted; if you ever
   re-register the app or change scopes, you'll need that consent step redone.

## History (why it's built this way, not some other way)
1. Started as a plan to host raw HTML/JS directly in a SharePoint document
   library — blocked (see gotcha #6).
2. Pivoted to Power Apps canvas app + Power Automate — this was fully designed
   (see git history / prior conversation) and functional, but the user wanted
   production hosting without Power Apps.
3. Rebuilt as this GitHub Pages + Microsoft Graph app, which is the current,
   live version. The Power Apps build is now obsolete and can be abandoned.

## Style/tone notes
The person building this (via chat, not yet via Claude Code) preferred very
granular step-by-step instructions and iterative debugging — screenshot the
exact error, fix one thing, retest. If continuing that pattern, small verified
steps tend to work better than large unverified changes for this project,
given how much of the above list was discovered through exactly that process.
