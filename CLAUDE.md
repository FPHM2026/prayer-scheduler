# FPHM Scheduler — Project Context

## What this is
A scheduling app for a church prayer ministry team, coordinating one-on-one prayer
sessions between external recipients and internal prayer ministers. Built as a
static HTML/JS single-page app, hosted on GitHub Pages, backed by SharePoint Lists
via Microsoft Graph API (not Power Apps, not raw SharePoint hosting — both were
tried and abandoned; see "History" below). A companion minister-facing portal
(`portal/index.html`) lets prayer ministers see their own sessions and manage
their own time off without needing a licensed Microsoft 365 seat — see "Minister
portal" below.

## Architecture
- **Frontend**: three self-contained single-file apps (vanilla JS, no build step,
  no framework, all UI/styling/logic per file): `index.html` (production admin
  scheduler), `preview/index.html` (staging copy for testing unreleased changes
  against real data), `portal/index.html` (minister-facing companion app).
- **Hosting**: GitHub Pages, org repo `FPHM2026/prayer-scheduler` (public),
  deployed from the `main` branch root — `/preview/` and `/portal/` are just
  subfolders in that same deployment, not separate hosting setups:
  - Production scheduler: `https://FPHM2026.github.io/prayer-scheduler/`
  - Minister portal: `https://FPHM2026.github.io/prayer-scheduler/portal/`
  - Preview build of the scheduler: `https://FPHM2026.github.io/prayer-scheduler/preview/`
    — see "Deployment workflow" below before editing this file.
- **Auth**: MSAL.js (`@azure/msal-browser`, loaded via jsDelivr CDN,
  `cacheLocation: "sessionStorage"`) — everyone signs in with their normal
  Microsoft 365 account via a popup, admins and ministers alike.
- **Data**: Microsoft Graph API (`https://graph.microsoft.com/v1.0`), talking
  directly to SharePoint Lists on the church's site. Not the older SharePoint REST
  API (`_api/web/...`) — that doesn't support CORS from an external origin; Graph
  does, via a registered Azure AD app.

## Deployment workflow (preview → production)
For anything bigger than a trivial/safe fix, build and test in
`preview/index.html` first, never `index.html` directly:
1. Edit `preview/index.html`. It carries blocks production doesn't have —
   a `#previewModeBanner` (HTML near the top of `<body>`, CSS a few lines
   above it), and a **Changes tab** (`panel-changes`, its nav button, the
   `CHANGE_LOG`/`CHANGE_LOG_ARCHIVE` arrays + `renderChangeLog()`) — strip
   all of this back out if copying preview's content wholesale into
   production. The Changes tab is the first tab and the default-active one
   specifically so a non-technical reviewer's first click lands on "what
   changed," not buried after everything else — add one new dated entry to
   `CHANGE_LOG` (newest at the top) on every preview publish, written in
   plain language describing what to actually go check, not a commit log.
   The tab has a Recent/Archive switch: `CHANGE_LOG` is "Recent" (only what's
   live in preview but not yet in production); `CHANGE_LOG_ARCHIVE` is
   "Archive" (already shipped, kept for reference). As step 6 below
   promotes entries to production, move those same `CHANGE_LOG` entries
   onto the *top* of `CHANGE_LOG_ARCHIVE` instead of deleting them —
   `CHANGE_LOG` should end up empty right after a full promotion, since
   preview and production are back in sync at that point.
2. Bump `const APP_VERSION = "..."` in preview/index.html (date + counter,
   e.g. `"2026-09-15.1"`). Two independent checks read this: each page polls
   its own deployed copy's APP_VERSION and prompts a reload if what's loaded
   is stale; separately, production's own page polls preview/index.html's
   APP_VERSION and shows a small "preview build differs" badge whenever it
   doesn't match production's — the two files should carry the *same*
   APP_VERSION once (and only once) they're fully back in sync content-wise.
3. Test locally: `python -m http.server <port>` in the repo root + the
   Claude_Browser tools, injecting mock `account`/`ministers`/`sessions`/
   `blackouts`/`locations` directly in the console to bypass real MSAL/Graph
   calls. This validates the app's own logic; it can NOT validate Entra/
   SharePoint permission configuration — that needs the user testing live
   against the real tenant with a real account (see "Style/tone notes").
4. Commit + push preview/index.html straight to `main` — GitHub Pages serves
   both files from the same branch root, so `/preview/` is just a URL path.
5. Have the user test against real data at the live preview URL.
6. To promote: either copy preview/index.html over index.html wholesale and
   strip the preview-banner block back out (when *all* of preview's changes
   are ready), or cherry-pick specific edits into index.html directly (when
   preview has multiple features in flight and only some are ready). Bump
   index.html's own APP_VERSION to match. Verify locally again before
   pushing. In the same pass, move whichever `CHANGE_LOG` entries just went
   live onto the top of `CHANGE_LOG_ARCHIVE` in preview/index.html (see step 1).
7. After every push to `main`, also fast-forward the `preview` git branch:
   `git checkout preview && git merge main -m "Sync preview branch" && git
   push origin preview && git checkout main`. GitHub Pages does NOT serve
   from this branch — hosting is always from `main` regardless — keeping it
   synced is just a clean history checkpoint of when preview and production
   last matched.

`portal/index.html` doesn't go through this dance — it's pushed directly to
`main`, since it's a fully separate file that never touches the production
scheduler's own code.

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
- **PrayerMinisters**: Title (name), Email (their real, checkable contact
  address — what staff use to reach them), Phone, **SignInEmail** (their
  unlicensed Entra ID sign-in address for the minister portal — not a real
  inbox, deliberately a separate field from Email so giving someone portal
  access never overwrites how you actually contact them), Status (Choice:
  Active/Inactive), **Gender (Choice: Male/Female, added 2026-09-17)** — feeds
  the same-sex team-preference matching on the Schedule & Planning tab, see
  "Planning Slots" under "Features implemented" below.
- **PrayerSessions**: Title, SessionDate, SessionEndDate, RecipientName,
  RecipientContact, AssignedMinisterIDs (comma-separated minister IDs, plain
  text), LeadMinisterIDs (comma-separated subset of the above who are "Lead"
  rather than "Support"), LocationName (plain text, not a true Lookup column —
  written as plain text matching a Locations list Title, chosen deliberately to
  avoid Graph's LookupId complexity), Notes (plain text — must be Plain text
  format, not Rich text, or it stores HTML), ApptType (Choice: First/
  Follow-up/Drop-in/Training — the last two are used by the Drop-in/
  Training quick-add buttons, see "Features implemented" below), Status
  (Choice: Scheduled/Completed/
  **Waiting**), PreviousSessionId (Number,
  links a follow-up session back to the one it followed), WaCreated (Yes/No),
  WaLink (plain text), **Priority (Yes/No, default No)** — flags a Waiting entry
  to the top of the Planning tab regardless of how long they've been waiting.
  **Contacted (Yes/No, default No)** — added to SharePoint 2026-09-17.
  Tracks whether the admin has reached out to a Waiting entry (see the
  Planning tab bullet under "Features implemented" below).
  **RecipientGender (Choice: Male/Female, added 2026-09-17)** and
  **GenderPreference (Choice: Mixed/Same-sex, default Mixed, added
  2026-09-17)** — the recipient's own gender and whether they want a
  same-sex team; used by the Planning Slots feature (see below).
  A Waiting-status item has SessionDate/SessionEndDate/LocationName/
  AssignedMinisterIDs/LeadMinisterIDs blank until it's actually scheduled.
  "Days waiting" is computed client-side from the item's own SharePoint
  `createdDateTime` metadata — there is no dedicated DateAdded field for this.
- **BlackoutDates**: Title (minister name), BlackoutDate (Date), **EndDate
  (Date, optional)** — blank/null means a single-day blackout; set means an
  inclusive date range. Notes (Plain text)
- **Locations**: Title (location name) — deliberately simple, just a name.
  Managed from the Settings tab's Locations sub-view (add/delete) since
  2026-09-19 — see "Features implemented" below for the deletion-safety
  fallback that keeps a deleted location's name on any session that already
  used it.
- **PlanningSlots** (added 2026-09-17): Title, SlotDate (Date), StartTime/
  EndTime (plain text, "HH:MM" 24-hour — not a true Time column), Status
  (Choice: Open/Tentative/Booked/**Drop-In** — note the hyphen and capital
  I; don't confuse with PrayerSessions' own ApptType="Drop-in", lowercase i,
  a separate field on a separate list), ClaimedWaitingId (Number, the
  claimed PrayerSessions Waiting item's ID), ClaimedDate (Date),
  ProposedMinisterIDs / ProposedLeadIDs (comma-separated, same plain-text-ID
  convention as AssignedMinisterIDs/LeadMinisterIDs above), LinkedSessionId
  (Number, the real PrayerSessions item once Booked), Notes (Plain text).
  See "Planning Slots" under "Features implemented" for how the states flow.
- **TimeWindows** (added 2026-09-19): backs the Settings tab's Quick Slots
  and Sessions sub-views. Title (free label, e.g. "Sunday", "Drop-In",
  "Default Location"), Kind (Choice: Weekday/DropIn/DefaultLocation),
  DayOfWeek (Number, 0=Sunday…6=Saturday — only set on Weekday rows, blank
  on the other two kinds), StartTime/EndTime (plain text, "HH:MM" 24-hour,
  same convention as PlanningSlots above — used on Weekday and DropIn rows,
  blank on DefaultLocation), LocationName (plain text matching a Locations
  list Title, same convention as PrayerSessions' own LocationName — used on
  DropIn and DefaultLocation rows, blank on Weekday). Exactly one Weekday
  row per configured default day, exactly one DropIn row, exactly one
  DefaultLocation row — no separate "enabled" flag, presence as a Weekday
  row *is* enabled for that day.
- **Prospects**: retired. The app no longer reads this list at all — replaced
  by PrayerSessions items with Status="Waiting" (see above). The list itself
  still exists in SharePoint with whatever old data was in it; run
  `FPHM Scheduler Import Prep/migrate-prospects-to-waiting.console.js` once
  (same browser-console pattern as the historical session import) to copy any
  remaining entries over as Waiting sessions, then archive/delete the
  Prospects list yourself whenever you're ready. Not done automatically —
  Claude has no SharePoint write access outside a session you're driving.

## Features implemented
Every list-style tab (Schedule, Completed Sessions, Prayer Ministers,
Blackout Dates) is grouped into collapsible accordion cards by recipient or
minister, collapsed by default, with a "Session #N" chip computed once per
recipient across their full chronological history (any status, including a
Waiting entry) — not per-view. A shared bottom-sheet modal pattern (full
width, slides up from the bottom, capped at 90vh) is used for every form
(New Session, Minister, Blackout Date); the confirm dialog and Add Location
stay as small centered modals. Main nav is a sticky, single-row,
horizontally-scrolling tab strip on mobile/tablet; desktop restores the
fuller header and lets tabs wrap. Every search field across both this app
and the portal has an inline clear (×) button, shown only once there's text
to clear. Every card-facing date (session rows, group summaries, stats
ranges) includes the year, via the shared `fmtDate`/`fmtShort` helpers.

- Schedule tab: grouped by recipient, search, Upcoming/All/Past filter,
  soonest-upcoming-first sort, plus "+ Drop-in" and "+ Training" quick-add
  buttons alongside "+ New Session" — see the Drop-in/Training bullet below
  for what those prefill.
- New/Edit/Follow-up session form: recipient info, date + day-of-week label,
  a quick-select start-time button auto-generated from that date's weekday
  default (one preset per configured day — see the Settings tab's Quick Slots
  section below; "No preset window for this day" shows instead for a day with
  no configured default), custom start/end time (plain native time inputs, no
  granularity restriction), location dropdown with inline "add new location",
  appointment type (First/Follow-up/
  Drop-in/Training, smart-defaulted), status (Scheduled/Completed — hidden on
  new sessions, shown when editing), minister multi-select (2+ required) with
  Lead/Support role toggle per minister, live "also on this date" panel,
  per-minister upcoming-session-count badge, blackout-date flagging,
  minister-specific double-booking prevention (two sessions CAN overlap in time
  if they share no ministers — only blocks if a specific minister would be in
  two places at once), notes, WhatsApp group-created toggle + optional invite
  link. The same modal doubles as the Waiting-list form (see Planning below) —
  a `sessKind` flag ('session' vs 'waiting') shows/hides the scheduling-only
  fields rather than being a separate form.
  - **Recipient session history + team pre-fill**: once the typed recipient
    name matches prior PrayerSessions records (name-match based — recipients
    aren't persistent records), a collapsed-by-default panel expands into
    every matching prior session, most recent first — date with weekday
    (`fmtDate`), ministers with Lead/Support (`ministerNamesOf`), and that
    session's notes — each row with a "Use this team" action. Replaces the
    old one-line "last session with this recipient" hint. The most recent
    session's team auto-applies for a genuinely new session only (not an
    edit, not a follow-up, not a Drop-in/Training preset — each of those
    already has its own correct minister source and must not be overridden),
    visibly instead of silently, so a one-off substitution is something the
    admin can catch and override rather than something that quietly becomes
    the new default. Same recipient match as the autocomplete above; zero
    new data — pure client-side logic against sessions already loaded.
  - **"+ Drop-in" / "+ Training" quick-add buttons**: prefill the form for a
    shared group event — recipient name set to "Sunday Drop-In" / "Training",
    Appointment Type set to match, every Active minister pre-selected as
    Support, Recipient Contact, Recipient Gender, and Team Preference all
    hidden (none apply to a group session). Drop-In further defaults Start/
    End Time and Location from the Settings tab's own Sunday Drop-In row
    (Quick Slots section) — falls back to 5:00–8:15pm / "Kids' Wing" by name
    if that row hasn't been configured yet. Excluded from every statistic
    (the Completed Sessions combo
    card, the Prayer Ministers tab's hours/serving numbers) since assigning
    every minister to the same session would otherwise wildly inflate those —
    the underlying session records still show up normally in the plain
    lists, just not in the aggregates.
- Completed Sessions tab: grouped by recipient, read-only, search, plus a
  collapsible combo stats card — YTD / Last Year / Lifetime columns for
  session count and unique recipients (with a first-time vs. follow-up
  breakdown per column and a % delta badge on the Last Year column), a
  "Since &lt;Month Year&gt;" lifetime range label computed from the earliest
  completed session on record, and a monthly trend chart comparing this
  year to last. Same combo card design ported to the portal's own Completed
  Sessions tab, scoped to just that minister's sessions.
- Planning tab: **superseded 2026-09-17 by the Schedule & Planning tab below**
  — the old Schedule and Planning tabs' HTML/JS are both still fully intact
  in `preview/index.html`, just deliberately unlinked from the main nav (see
  the comment above the `<nav>` block) as an easy-to-restore fallback, not
  deleted. What follows describes that retained-but-hidden code, unchanged:
  PrayerSessions items with Status="Waiting" (no separate
  Prospects list — see above), sortable by longest-waiting or highest session
  number, priority-flagged entries always pinned to the top, flag/schedule/
  edit/delete actions per entry. "Schedule Session" reopens the same session
  modal with the scheduling fields revealed, converting the same record.
  A phone-icon "Contacted" toggle (`toggleContacted`) plus a small badge next
  to Priority records whether the admin has reached out — a manual flag for
  an external action, same pattern as `WaCreated`'s WhatsApp toggle; the app
  never sends anything itself. Fully independent of Status — toggling it
  doesn't move an entry out of Waiting, only "Schedule Session" does that.
- **Schedule & Planning tab** (added 2026-09-17, promoted to production the
  same day): replaces the old
  Schedule + Planning tabs with one merged view backed by the new
  PlanningSlots list, so every slot state — Open, Tentative, Booked, or the
  recurring Drop-In — is one collapsible accordion card grouped by week, with
  a Booked card's expanded body showing the same facts the old Schedule tab
  row did (recipient, contact, team, notes). Out of scope for this change,
  per explicit decision: the minister-facing `portal/index.html`'s own
  Schedule tab is untouched.
  - **Capacity**: a normal date's slot count is `floor(activePMs / 2)` where
    activePMs excludes anyone blacked out that specific date (matched by
    minister *name*, same convention as `renderMinisterList` elsewhere — not
    an ID join) — a session always needs exactly 2 PMs, so an odd one out
    just sits that date out.
  - **Quick Add Slots**: a modal with "Next 4 Weeks"/"Next Calendar Month"
    range presets (or a manual date range), a day-of-week filter, and a live
    per-date breakdown before confirming. The number of slots it creates per
    date is capacity minus however many already exist there, clamped at
    0 — idempotent by construction, so re-running it over an
    already-provisioned or now-unavailable date adds nothing.
  - **4th Sunday = Drop-In**: `isFourthSunday()` overrides the normal
    capacity math entirely for that one date each month — a single Drop-In
    slot instead of the usual 2-PM-pair slots, timed from the same
    Settings-configured Sunday Drop-In row as the "+ Drop-in" button above
    (fallback 5:00–8:15pm if unconfigured).
  - **Custom slots**: a "+ Add custom slot" affordance on any date that
    already has calculated slots lets the admin add one more, with a Type
    picker — Regular (a claimable PlanningSlot, same as always), Sunday
    Drop-In (also slot-based, mirrors Quick Add's own Drop-In slot), or
    Training (opens the normal session modal pre-filled with that date
    instead — Training has no slot concept anywhere else in the app, so this
    reuses the existing save path and its conflict checks rather than
    duplicating them). Regular/Drop-In use the same minister-double-booking
    rule as everywhere else (blocks only on an actual time *overlap*, not
    merely the same day).
  - **Claiming**: an Open slot's picker lists every unclaimed Waiting
    candidate; expanding one shows their prior-session history (if any) to
    reuse a past team via "Use this team," or "Claim with auto-pick" to
    build a fresh team from who's actually free for that slot's date/time
    (excluding anyone already committed to an overlapping Tentative/Booked/
    Drop-In slot elsewhere that day).
  - **Gender preference** (Male/Female on PrayerMinisters; RecipientGender +
    GenderPreference on PrayerSessions/Waiting, see schema above): the
    capacity/Quick-Add math stays gender-agnostic on purpose — the
    preference only affects availability *after* someone with a same-sex
    requirement is being matched, not before. An Open slot's header always
    shows live "Available now: N PMs (XF, YM)." Auto-picking a team for a
    Same-sex-preference candidate with no usable history is a **hard
    block** if fewer than 2 of the needed gender are actually free right
    now (mirrors the existing "at least 2 PMs" hard-block on the session
    form). Reusing a *specific* historical team is always allowed even if
    it doesn't match, just flagged with a non-blocking "⚠ doesn't match…"
    warning instead (mirrors the existing Blackout Dates soft-override
    pattern) — a deliberate, visible admin choice is never second-guessed.
  - **Booking**: "Book This Session" on a Tentative slot opens the normal
    session modal pre-filled from the slot (date/time/team), converting the
    claimed Waiting record in place into a real session on save — same
    underlying mechanism as the old `scheduleFromWaiting`. On save, the
    slot is PATCHed to Booked + LinkedSessionId + the session's actual
    saved times. A Drop-In slot's own "Record Drop-In Session" button
    reuses the existing "+Drop-in" preset (recipient/appt-type/full active
    roster prefilled) the same way, just also pinning the date and PATCHing
    the slot on save.
- Prayer Ministers tab (merged Minister Overview + roster management into
  one tab): grouped by minister, Active/Inactive sub-tabs, search, date range
  defaulting to year-to-date, collapsible stats block (hours this range vs.
  the same calendar range last year, active ministers serving, avg hours per
  minister), a sort toggle cycling Next Appointment ⇄ A→Z, per-minister
  upcoming sessions, activate/deactivate/edit/delete actions per minister
- Blackout Dates tab: grouped by minister, soonest first, Add/Edit modal with
  a single-date/date-range toggle (writes BlackoutDate + EndDate), edit/delete
  per entry
- **Reports tab** (added 2026-09-19; real report content added 2026-09-21,
  corrected the same day against the actual punch-list spec — see "Known
  gotchas" #14): a date range picker only — no report-type picker, since
  one Generate always produces all three types together. Date-range presets
  (This Year, Last Year, Last 3 Years, All Time — same `.range-preset`
  pattern Quick Add Slots uses) fill the From/To fields without typing;
  "Last 3 Years" is the last 3 *full* calendar years, excluding the
  current partial one, so it's a clean complement to "Last Year" instead of
  overlapping it. Generate computes against the live `sessions`/`ministers`
  already loaded (Completed only — same retrospective convention
  `completedStats()` uses) and opens a standalone, printable page in a new
  tab (Blob + `window.open`, not a `data:` URI — more reliable across
  browsers for a full HTML document); the new tab has no access back into
  the app, so every number is baked into plain HTML before it opens,
  nothing there re-fetches or recomputes anything.
  - **Three distinct sections on one page** — Freedom Sessions, Sunday
    Drop-In, Training — stacked, never merged, since their totals are kept
    intentionally separate.
  - **Each section is a year-by-year table**, not cards/tiles: one row per
    calendar year touched by the chosen range (a partial first/last year
    only counts what's actually in range), Sessions + a **Running Total**
    column (shaded background, bold — visually distinct so it's never
    mistaken for a single year's own count; cumulative from 0 at the start
    of *this report's* range, not a lifetime total). Freedom Sessions adds
    Unique Recipients / First-time / Follow-up columns (same per-
    *recipient*, not per-session, classification `completedStats()` uses);
    Drop-In and Training skip those three — every occurrence shares one
    literal recipName ("Sunday Drop-In"/"Training"), so there's no
    individual-recipient concept to split.
  - **Team-wide table shown first**; a collapsed-by-default "Show minister
    breakdown" toggle per section (native `<details>`/`<summary>`) expands
    into sessions + times-as-Lead per minister, tagging anyone whose
    current PrayerMinisters Status isn't Active with a small "Inactive"
    label — the breakdown includes everyone with a session in range
    regardless of current roster status. The page's own Print button
    force-sets every `<details>` on the page open before calling
    `window.print()` — collapsed content doesn't print, so this is
    load-bearing, not cosmetic.
  - Each section shows its own empty-state row when it has no completed
    sessions in range, rather than a misleading all-zeros table.
  - **Not built, still a real option**: Location is a clean, trustworthy
    report dimension now (an admin-managed list via Settings, not free
    text) — "sessions by location" as an optional breakdown/filter is
    reasonable to add later, just wasn't part of this pass.
- **Settings tab** (added 2026-09-19): admin-only configuration, no code
  changes needed, backed by a new `TimeWindows` SharePoint list (see schema
  below). Three sub-tabs:
  - *Quick Slots*: two cards — **Default Days** (which weekdays have a
    default Freedom Session window and what time each runs; add/edit/remove
    a day; feeds both the session form's quick-select button and Quick Add
    Slots' own day-chip defaults) and **Sunday Drop-In** (default Start/End
    Time + Location for Drop-In slots and sessions, singleton row).
  - *Locations*: add/delete locations that feed the location dropdown
    elsewhere. Deletion is allowed even if a location is in use (with a
    confirmation) — existing sessions keep displaying their original
    location name via a `locationRawName` fallback captured at load time,
    since the live id-based lookup would otherwise go blank once the
    location's gone.
  - *Sessions*: the default location used when creating a brand-new session
    before a room has been decided (singleton row; falls back to "To Be
    Determined" by name, then the first location, if unconfigured).
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
  - BlackoutDates: Contribute. Item-level "Read"/"Create and edit" were
    originally set to "items created by the user," intending SharePoint
    itself to enforce "ministers only see/manage their own time off" — this
    turned out **not to be reliably honored via Microsoft Graph** even for
    Full Control users (see gotcha #11), so both were relaxed to "all
    items." The "only your own" UX is now entirely client-side filtering in
    both apps, not a SharePoint-enforced boundary — an accepted tradeoff,
    not a hard security guarantee.
  - Locations: no access needed (LocationName is already plain text on the
    session item)
- Needs its own redirect URI on the same app registration, under the
  **Single-page application** platform specifically (see gotcha #8):
  `https://FPHM2026.github.io/prayer-scheduler/portal/`
- **Identity matching**: the signed-in account's email (`account.username`)
  must exactly match (case-insensitive) that minister's **SignInEmail**
  field in PrayerMinisters — not the Email field, which is their real
  contact address and is never touched by this matching — or the portal
  shows a "couldn't find your profile" message instead of a broken page.
  Set SignInEmail from the main app's Add/Edit Minister form when creating
  a minister's Entra account.
- Creating the actual Entra accounts (Microsoft 365 admin center → Add a
  user → "Create user without product license") and the SharePoint group/
  permission setup above are manual admin steps outside this repo — Claude
  has no tenant admin access to do them.
- **Sign-in redirect (production, in index.html)**: a minister who signs
  into the *main scheduler* — not the portal — gets redirected straight to
  `/portal/` instead of landing in the admin UI. `afterSignIn()` checks the
  signed-in account's email against every minister's SignInEmail right
  after resolving the SharePoint site, before ever showing the admin UI; a
  match redirects via `window.location.href` (relative `portal/` path), no
  match proceeds into the admin app as normal. This means one URL — the
  production scheduler link — can be handed out to everyone; SharePoint's
  own permissions remain the real access boundary, this redirect is a UX
  nicety on top, not a security control.

### Portal UI (mirrors the main app's component style)
Three tabs — Schedule, Completed Sessions, Blackout Dates — reusing the main
app's accordion-card/bottom-sheet-modal CSS verbatim so the two apps feel
like one product, not two:
- Schedule / Completed Sessions: same grouped-by-recipient cards as the main
  app, filtered to sessions the signed-in minister is assigned to; Schedule
  has an Upcoming/All/Past range filter, both have search (with a clear-×
  button, same as every search field in both apps).
- Blackout Dates: minister manages their own time off (add/edit/delete),
  with an Upcoming/All/Past range filter defaulting to Upcoming so old time
  off doesn't clutter the default view — a range entry still in progress
  (started before today, ends after) stays classified as Upcoming until its
  whole span is behind today. The "Next unavailable" summary always reflects
  the true next upcoming entry regardless of which range is currently being
  browsed below it.

## Not yet built
- **Calendar (month grid) view** — was planned but never built in this HTML
  version. The agenda/Schedule view covers the "chronological list" requirement
  on its own; the calendar grid is a nice-to-have, not yet started.
- **Location as a report dimension** — Reports' Session/Training/Drop-in
  Statistics (built 2026-09-21, see "Features implemented") don't break
  down or filter by Location yet, even though it's now a clean,
  admin-managed list via Settings rather than free text. A reasonable
  optional addition, not started.

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
8. **AADSTS9002326 "Cross-origin token redemption" error** happens when a
   redirect URI is registered under the **Web** platform in Entra instead of
   **Single-page application** — Web assumes a confidential server-side
   client and blocks the cross-origin token exchange MSAL's popup/silent
   flows need. Every redirect URI this app uses (root, `/preview/`,
   `/portal/`) must live under the Single-page application platform
   section, never Web — this has regressed more than once (a URI ending up
   duplicated under both platforms, or added under the wrong one by habit).
   If it happens again: Entra admin center → App registrations → this app →
   Authentication → check both the Web and Single-page application sections
   for the exact URI that's failing.
9. **Two `position:sticky;top:0` elements don't stack, they overlap.** If a
   banner sits above the main sticky nav in DOM order, giving both the same
   `top:0` makes whichever has the higher z-index render on top of the
   other once scrolled, rather than the nav pushing below the banner — fix
   is to only make ONE of them sticky (the preview banner deliberately
   isn't), not to keep raising z-index.
10. **Native `<input type="time">` renders 24-hour or 12-hour AM/PM
    depending on the page's locale, not anything CSS can control.** A bare
    `lang="en"` is ambiguous; `lang="en-US"` on `<html>` gets Chromium
    browsers (Chrome/Edge) to render 12-hour AM/PM. Firefox/Safari mostly
    follow the OS locale instead and may not be affected either way.
11. **SharePoint item-level "Read/Create/Edit items created by the user"
    restrictions are NOT reliably honored for Full Control users when
    accessed via Microsoft Graph**, even though the classic SharePoint web
    UI does honor that exemption for them. Confirmed empirically on
    BlackoutDates: an admin with Full Control (via group membership AND a
    direct grant) saw zero items until "Read access" was changed from "Read
    items created by the user" to "Read all items." Pragmatic fix used
    here: relax item-level restrictions to "all items" wherever they're
    set, and rely entirely on each app's own client-side filtering for the
    "you only see your own X" UX — an accepted tradeoff, not a SharePoint-
    enforced boundary.
12. **Filtering displayed items by a secondary control (date range) AFTER
    determining a search match can silently hide matches instead of
    showing them.** The Schedule tab's "search by minister or recipient"
    matched correctly, but then filtered the matched recipient's session
    list down to whatever the Upcoming/Past range allowed — if a matched
    minister's sessions all fell outside that range, the whole group
    vanished with no indication a match had even occurred. Fixed by
    skipping the range filter entirely while a search query is active —
    search overrides browsing filters, not the other way around.
13. **A wholesale preview → production copy carries `preview/index.html`'s
    own relative-path assumptions with it.** `PORTAL_URL` is `"../portal/"`
    in preview (one folder deeper than the repo root) but must be
    `"portal/"` in production — copying the file verbatim silently points
    production's minister sign-in redirect one level too high. Also strip
    the `#previewModeBanner` CSS/markup block AND the comment that explains
    it (the comment sits just above the CSS rule it describes, easy to
    leave orphaned if only the rule itself gets removed), and the entire
    **Changes tab** (`panel-changes` section, its nav button, `CHANGE_LOG`/
    `renderChangeLog()`) — production never had any of these. Check all of
    it by hand after every promotion; nothing catches this automatically.

## History (why it's built this way, not some other way)
1. Started as a plan to host raw HTML/JS directly in a SharePoint document
   library — blocked (see gotcha #6).
2. Pivoted to Power Apps canvas app + Power Automate — this was fully designed
   (see git history / prior conversation) and functional, but the user wanted
   production hosting without Power Apps.
3. Rebuilt as this GitHub Pages + Microsoft Graph app, which is the current,
   live version. The Power Apps build is now obsolete and can be abandoned.
4. Added the minister portal (unlicensed Entra accounts, see "Minister
   portal" above) once it became clear ministers needed real self-service
   access to their own schedule and time off, without per-seat licensing
   cost — this superseded an earlier "Microsoft Forms + Power Automate, no
   accounts needed" idea that was considered and dropped in favor of real
   accounts with real sign-in.

## Style/tone notes
Now being built via Claude Code, not just chat. The person building this
prefers granular step-by-step instructions and iterative debugging —
screenshot the exact error, fix one thing, retest — and for SharePoint/Entra
configuration steps specifically (outside this repo, outside Claude's own
access) wants clear numbered instructions. Small verified steps tend to work
better than large unverified changes for this project, given how much of the
gotchas list above was discovered through exactly that process.

Real end-to-end testing — signing in with a real test minister account
against live SharePoint/Entra — has repeatedly caught bugs that local
mock-data testing structurally cannot catch, since mocking bypasses the
real auth/permission layer entirely: the AADSTS9002326 redirect-URI-
platform gotcha (#8), the SharePoint item-level-permission/Graph
unreliability (#11), the sticky-banner overlap (#9). Treat mock-data
testing (injecting `account`/`ministers`/`sessions`/etc. via the browser
console) as validating the app's own logic; treat real-account testing as
the only way to validate the Entra/SharePoint configuration around it — the
two are not substitutes for each other.

Feature requests get tracked in a separate Claude-authored artifact ("FPHM
Scheduler — Punch List," a claude.ai artifact, not a file in this repo) that
the user shares a link to — often with real design detail (exact columns,
which UI pattern to reuse, what to explicitly exclude) well beyond what
gets repeated in the chat message asking for the work. The Reports tab was
first built from a reasonable-sounding but incomplete guess when that
spec wasn't checked first, and had to be substantially reworked once it
was (see gotcha-style lesson: when a request references "requirements,"
"the punch list," "what was requested," or similar and no link is in the
current conversation, ask for the artifact link before designing anything
non-trivial, rather than proposing options and building from whichever one
sounds closest).
