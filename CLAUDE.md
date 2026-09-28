# FPHM Scheduler — Project Context

## What this is
A scheduling app for a church prayer ministry team, coordinating one-on-one prayer
sessions between external recipients and internal prayer ministers. Built as a
static HTML/JS single-page app, hosted on GitHub Pages, backed by SharePoint Lists
via Microsoft Graph API (not Power Apps, not raw SharePoint hosting — both were
tried and abandoned; see "History" below). A companion minister-facing portal
(`portal/index.html`) lets prayer ministers see their own sessions and manage
their own time off without needing a licensed Microsoft 365 seat — see "Minister
portal" below. A third piece, `intake/index.html`, is the public-facing digital
intake questionnaire recipients fill out before their first session — no sign-in,
save-and-resume via a link, digital signature — with staff and minister views
of completed forms built into this same app; see "Intake Forms" below.

## Architecture
- **Frontend**: four apps, mostly self-contained single files (vanilla JS, no
  build step, no framework): `index.html` (production admin scheduler),
  `preview/index.html` (staging copy for testing unreleased changes against
  real data), `portal/index.html` (minister-facing companion app),
  `intake/index.html` (public recipient-facing intake form). The one
  deliberate exception to "all logic per file": `js/formConfig.js` (the
  intake question schema) and `js/formEngine.js` (its rendering/validation
  logic) are shared via `<script src="../js/...">` — so editing a question
  is ever only a one-file change — across `preview/index.html`,
  `portal/index.html`, and `intake/index.html` as of the Intake Forms merge
  (2026-09-24). **`index.html` (production) does not load either shared
  file yet and has no Intake Forms tab** — the merge commit was explicitly
  a "preview build," not yet promoted; see "Intake Forms" below and the
  "Not yet done" bullet there.
- **Hosting**: GitHub Pages, org repo `FPHM2026/prayer-scheduler` (public),
  deployed from the `main` branch root — `/preview/`, `/portal/` and
  `/intake/` are just subfolders in that same deployment, not separate
  hosting setups:
  - Production scheduler: `https://FPHM2026.github.io/prayer-scheduler/`
  - Minister portal: `https://FPHM2026.github.io/prayer-scheduler/portal/`
  - Public intake form: `https://FPHM2026.github.io/prayer-scheduler/intake/`
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
1. Edit `preview/index.html`. It carries one block production doesn't have
   — a `#previewModeBanner` (HTML near the top of `<body>`, CSS a few lines
   above it) — strip that back out if copying preview's content wholesale
   into production. The **Changes tab** (`panel-changes`, its nav button,
   the `CHANGE_LOG`/`CHANGE_LOG_ARCHIVE` arrays + `renderChangeLog()`) is
   NOT preview-only, unlike the banner: it originally was meant to be
   stripped at promotion too, but that step got missed on the 2026-09-27
   promotion, and the user decided to just keep it in production rather
   than have it stripped back out — so as of that date it ships in both,
   unconditionally, moved to the *end* of the tab list in both (it was
   previously first and default-active specifically so a reviewer's first
   click landed on "what changed" — no longer the case now that it's
   permanent in both builds). Still add one new dated entry to `CHANGE_LOG`
   (newest at the top) on every preview publish, written in plain language
   describing what to actually go check, not a commit log. The tab has a
   Recent/Archive switch: `CHANGE_LOG` is "Recent" (only what's live in
   preview but not yet in production); `CHANGE_LOG_ARCHIVE` is "Archive"
   (already shipped, kept for reference). As step 6 below promotes entries
   to production, move those same `CHANGE_LOG` entries onto the *top* of
   `CHANGE_LOG_ARCHIVE` instead of deleting them — `CHANGE_LOG` should end
   up empty right after a full promotion, since preview and production are
   back in sync at that point.
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
  "Planning Slots" under "Features implemented" below. ~~PortalPassword~~ —
  briefly added 2026-09-27, then removed the same day once the password
  design settled on a shared, non-stored default (see "Add/Edit Minister
  creates the actual sign-in account" under "Features implemented" below).
  If you added this SharePoint column while following an earlier version of
  this doc, it's safe to delete — nothing reads or writes it anymore.
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
  **IntakeResponseId (Number, optional, added 2026-09-26)** — a
  staff-confirmed link from this session/Waiting entry to a specific
  IntakeResponses item, overriding the plain name-match
  `findIntakeForRecipient()` otherwise falls back to. See "Intake form
  linking" under "Intake Forms" below.
  **ConfirmedMinisterIDs and DeclinedMinisterIDs (both Single line of text,
  Plain text — MUST be added to SharePoint by hand, not created
  automatically; ConfirmedMinisterIDs added 2026-09-27, DeclinedMinisterIDs
  added the same day right after)** — mutually exclusive comma-separated
  subsets of AssignedMinisterIDs: who's confirmed, and who's declined, from
  their own portal card, that they're actually coming — same shape as
  LeadMinisterIDs. Both only ever written by `setMyAttendance()` in
  `portal/index.html` — the signed-in minister setting/clearing their own
  attendance response, never anyone else's. See "Minister attendance
  confirmation" under "Features implemented" below.
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
- **IntakeResponses** (added 2026-09-24, see "Intake Forms" below for the
  full feature): Title (recipient name once known, else "New intake"),
  Token (Text, a random UUID — the anonymous recipient's resume secret,
  never shown to staff/ministers), RecipientName / RecipientEmail (Text,
  mirrored from the `name`/`email` answer keys on every save),
  Status (Choice: InProgress/Submitted), ResponsesJSON (Multiple lines of
  text, **Plain text**, "Allow unlimited length" enabled — the entire
  answers object as JSON, keyed by question id from `js/formConfig.js`;
  this is *why* adding/removing a question never needs a SharePoint schema
  change), SignatureDataUrl (Multiple lines of text, Plain text, unlimited
  length — base64 PNG from the signature canvas), SubmittedAt (Date and
  Time, set only on final submit). The Ministers SharePoint group has Read
  only on this list (granted separately from its Contribute on
  BlackoutDates) — see "Intake Forms" for why that boundary is
  client-side-only, same as everywhere else ministers' access is scoped.
- **IntakeFormSchema** (added 2026-09-25): a single-item list holding the
  ENTIRE intake question schema as one JSON blob — Title (always
  "Current"), SchemaJSON (Multiple lines of text, Plain text, unlimited
  length — same JSON shape as `js/formConfig.js`'s
  `window.FPHM_INTAKE_CONFIG`), SchemaVersion (Number, incremented on every
  save, purely informational — no code reads it). This is what the "Edit
  Intake Questions" tab writes to and what `intake/index.html` fetches
  (via the Cloudflare Worker's `/api/intake/schema`) and both
  `preview/index.html`/`portal/index.html` fetch (via their own delegated
  Graph session) to override the hardcoded default from `js/formConfig.js`
  at runtime — see "Intake Forms" below for the full mechanism. No new
  SharePoint permission grant was needed for this one: staff already have
  full read/write via existing delegated access, and it's covered by the
  same app-only `Sites.Selected` grant the Worker already had for
  IntakeResponses.
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
    - **Recipient DETAILS (Contact/Gender/Team Preference) auto-fill for
      Waiting too** (added 2026-09-26): gated separately from the team
      auto-apply above (`historyDetailsAutoOk`, not `historyAutoApplyOk`)
      since these describe the *recipient*, not a team assignment — they're
      just as useful adding someone to the Waiting list as starting a new
      session, unlike the minister team (a Waiting entry has no team yet,
      see the PrayerSessions schema notes above). Sources from the single
      most recent matching record of *any* status, not just real sessions —
      a Waiting entry already carries these same three fields from when it
      was first created, so someone whose only prior record is another
      Waiting entry still gets their details back. Both the typeahead pick
      (`pickRecipNameAt`) and the history panel's own auto-apply
      (`renderRecipientHistory`) share the one `historyAutoName`
      once-per-distinct-name tracker so neither path re-fires or clobbers a
      manual edit while the admin keeps typing a still-matching name.
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
  - **Contacted stays available after claiming, right in the slot's
    summary row** (`renderSlotCard()`, added 2026-09-28, at the user's
    request — moved from the expanded body to the summary the same day,
    also at the user's request, so it's usable without expanding the
    accordion): a Tentative slot's collapsed header row shows the same
    Contacted icon-button the Unplaced Candidates card shows, next to the
    toggle button — previously it disappeared the moment a waiting
    candidate was claimed onto a slot, since only the Unplaced Candidates
    view rendered it, and the first fix only added it back inside the
    expanded `renderTentativeBody()`, not the summary. Structurally it's a
    **sibling** of the `.slot-header` toggle button, not nested inside it —
    a `<button>` can't nest inside another `<button>` — wrapped together in
    a flex row, with `event.stopPropagation()` on the Contacted button so
    clicking it doesn't also toggle the accordion open/closed. Same
    `toggleContacted()` call either way; the expanded body still shows a
    small "Contacted" tag next to the name for context, just not a second
    copy of the button.
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
  - **Add/Edit Minister can create and manage the actual sign-in account**
    (added 2026-09-27, promoted to production the same day):
    previously, creating a minister's unlicensed Entra account was entirely
    a manual step in the Microsoft 365 admin center (still documented under
    "Minister portal" below) — this makes the app itself do it via Graph,
    gated behind two delegated scopes (added to `GRAPH_SCOPES`):
    `User.ReadWrite.All` and `Directory.AccessAsUser.All`. Both need their
    own admin consent in Entra (a manual step, same category as the
    original `Sites.ReadWrite.All` consent in gotcha #7) for a signed-in
    admin who actually holds a role that can manage users (User
    Administrator or similar). **Both scopes are required**:
    `User.ReadWrite.All` alone creates accounts fine (`POST /users`) but
    gets a 403 `Authorization_RequestDenied` on the password-reset `PATCH`
    specifically — confirmed by extensive live troubleshooting the same day
    (ruled out: stale token, missing role, wrong app/tenant, guest-type or
    hybrid-synced target, a Continuous-Access-Evaluation claims challenge,
    even reproduced identically through Microsoft's own Graph Explorer)
    before a Microsoft Q&A thread surfaced `Directory.AccessAsUser.All` as
    the actual missing piece — adding it fixed the PATCH immediately.
    - **"Also create their sign-in account"** — a checkbox on Add Minister,
      checked by default (unchecking it just saves the roster entry, same
      as before this feature existed). Only shown when the minister doesn't
      already have one (`signInEmail` blank) — an existing account can only
      have its password changed here, not be recreated.
    - **`generateSignInEmail()`** suggests `firstname.lastname@` +
      `PORTAL_ACCOUNT_DOMAIN` (`creeksidechurch.ca`) into the Sign-In Email
      field as the admin types the name — a starting point, not a hard
      rule; it's an ordinary editable field, so a collision with an
      existing tenant account just needs a hand-adjustment before saving.
      Never overwrites a value already there (editing an existing minister,
      or a value the admin has already hand-typed — tracked via a
      `dataset.touched` flag on the field, not a second piece of state).
    - **The password is always temporary, and always the same shared,
      year-stamped default** — `defaultTempPassword()` returns
      `'FPHM' + currentYear + '!'` (e.g. `FPHM2026!`). Both
      `createMinisterAccount()` and `updateMinisterAccountPassword()` set
      `forceChangePasswordNextSignIn: true`, so Entra always forces the
      minister to pick their own real password the first time they sign
      in — this was tried first as an admin-typed, persisted password
      (`forceChangePasswordNextSignIn: false`) and Graph rejected that
      combination outright (see above), so there was never a "real,
      ongoing" password for an admin to usefully store or look up anyway.
      Since it's always temporary, it's deliberately **not masked and not
      stored anywhere** — no SharePoint column for it (an earlier
      `PortalPassword` column was added, then removed the same day once
      this design settled; safe to delete if you added it). Add Minister
      prefills the plain-text field with this year's default (still
      editable if you want something else for one person); Edit Minister
      shows a **"Reset to a new temporary password"** button instead of an
      always-editable field, so a save never silently changes an existing
      minister's password by accident — clicking it reveals the field
      prefilled with this year's default and saving commits the reset.
    - **Account creation also adds the new minister to
      `ENTRA_MINISTERS_GROUP_ID`, an Entra security group** (named "Prayer
      Ministers" in the tenant; added 2026-09-27, same day) — this group,
      not the Entra account by itself, is the actual access boundary for
      the minister portal: the SharePoint site's own permissions are
      granted to this group rather than to individual users, so skipping
      this step would leave someone with a working sign-in but a blank
      portal. `addMinisterToEntraGroup()` is a plain delegated Graph call
      (`POST /groups/{id}/members/$ref`), gated behind one more scope
      (`GroupMember.ReadWrite.All`, added to `GRAPH_SCOPES`) — **its own
      admin consent step**, same category as the other scopes above.
      **This replaced a same-day-earlier, now-removed design** that used a
      classic SharePoint site permission group instead: Graph v1.0 has no
      endpoint for classic SharePoint site-group membership, so that
      version called the SharePoint REST API directly and needed its own
      token, an `EnsureUser` call, and a retry loop of up to 10 attempts /
      12 seconds apart (~2 minutes worst case) to work around SharePoint
      taking well over a minute to recognize a just-created Entra account.
      Switching the actual access grant from a SharePoint group to an Entra
      group sidesteps that propagation lag entirely — Graph's own group
      membership has no such delay — at the cost of one manual,
      one-time reconfiguration: the SharePoint site's permission grant had
      to be pointed at the new Entra group instead of the old SharePoint
      group. With account creation now doing both the Entra account and
      the group membership, adding a minister to the team is just filling
      out the form — no manual step left afterward, and no multi-minute
      wait either.
    - **Any failure in this chain blocks the whole save**, not just a
      warning: if the checkbox is checked and either the Entra account
      creation or the group add fails (a taken UPN, a 403 from a signed-in
      account that doesn't actually hold the right role, etc.), the
      minister record itself is NOT saved either — the error surfaces in
      the modal, distinguishes which step failed (the group-add error
      explicitly says the Entra account WAS created, so the admin knows to
      add them to the group by hand rather than retry account creation),
      and the admin fixes the issue and retries, rather than the roster
      silently ending up with a minister record that has no matching
      account/access and no clear sign anything went wrong.
    - **Deleting a minister removes them from the Entra group, but
      deliberately never touches their Entra account** (`deleteMinister()`
      / `removeMinisterFromEntraGroup()`, added 2026-09-27, same day). This
      is an explicit choice, not an oversight: at least one person on this
      roster is also a church employee, so a roster deletion must never
      disable or delete their actual Microsoft 365 account — revoking
      group membership (their portal access) is the only thing a roster
      change should ever be able to do to their account.
      `removeMinisterFromEntraGroup()` resolves the user's directory object
      id first (Graph's `DELETE /groups/{id}/members/{id}/$ref` needs that,
      not a UPN). Unlike account creation above, a failed group removal
      does **not** block the roster deletion — it still proceeds, with an
      `alert()` telling the admin the person may still have portal access
      and to remove them from the group by hand. Blocking here would trap
      the admin unable to remove someone from the roster over a transient
      API failure, which is worse than the leftover-access edge case it
      would prevent. Note this is deliberately narrower than the
      Active/Inactive toggle elsewhere on this tab, which does not touch
      group membership at all — only an actual roster **deletion** revokes
      portal access today.
- **Minister attendance confirmation/decline** (added and promoted to
  production 2026-09-27, `portal/index.html` + `preview/index.html`/
  `index.html` — see `ConfirmedMinisterIDs`/`DeclinedMinisterIDs` in the
  schema section above): a Scheduled session's card on a minister's own
  portal schedule shows a bright orange primary "I'm coming to this
  session" button and a secondary "I can't make it" button (`.btn`
  matching the design system's primary-action color) when they haven't
  responded yet. Once they pick one, that button becomes its own
  undo-labeled state — "✓ You're confirmed — tap to undo" (secondary) or
  "✗ You declined — tap to undo" (`.btn.danger`) — and the other option
  disappears; `setMyAttendance(sessionId, 'confirmed'|'declined')` fully
  reversible either way, per explicit request (plans change, in either
  direction). Confirmed and declined are mutually exclusive — setting one
  always clears the other — and only ever write the *signed-in* minister's
  own id; the buttons only render at all when the session is actually
  assigned to them (`s.ministers.find(m=>m.id===me.id)`). Decline
  (`DeclinedMinisterIDs`) was added as a follow-up the same day, right
  after confirm shipped, at the user's request.
  `ministerNamesOf()` — shared logic duplicated in `portal/index.html` and
  `preview/index.html`, same as everywhere else minister names get
  rendered — appends a plain-text ` ✓` after a confirmed minister's name,
  or ` ✗` after a declined one. Deliberately plain text, not an HTML/SVG
  icon: every caller wraps the result in `escapeHtml()`. This makes a
  minister's response visible in three places at once: their own portal
  card, their co-assigned minister(s)' portal cards (so a team can see who
  else is actually coming, at the user's explicit request), and every
  session listing in the admin Scheduler. **Also requires ministers to
  have Contribute access to PrayerSessions** (upgraded from Read — a
  manual SharePoint permission change, confirmed live 2026-09-27 after the
  write initially failed 403 with only Read access; same category as the
  Contribute access ministers already have on BlackoutDates). Promoted to
  production `index.html` the same day, after live verification with the
  new column(s) and the permission change in place.
- Blackout Dates tab: grouped by minister, soonest first, Add/Edit modal with
  a single-date/date-range toggle (writes BlackoutDate + EndDate), edit/delete
  per entry
  - **"All Ministers" bulk option on Add (not Edit)** (`populateBlackoutMinisterSelect()`/
    `saveBlackout()`, added 2026-09-28, at the user's request: "if I blackout
    Tuesday Sept 29 and add Quick Slots, no quick slots would be created"):
    picking "— All Ministers (blocks this date entirely) —" writes one
    ordinary per-minister BlackoutDates row for every currently-Active
    minister, same date/range/notes — **not** a special record type or a
    new SharePoint column. This works because `capacityForDate()` (used by
    Quick Add Slots) already computed `Math.floor(available.length/2)` off
    "every Active minister not blacked out that day" — once that set is
    empty, capacity was already 0, Quick Add Slots already added nothing.
    The only gap being closed here is the one-click UX; the actual
    blackout-awareness in slot generation already existed and needed no
    code change. Only offered when adding a brand-new blackout — there's
    no coherent "switch to All Ministers" action on an existing
    single-minister row, so it's not in the dropdown when editing one.
    **Known tradeoff**: since these are N ordinary rows, not one shared
    flag, a minister added or reactivated *after* an "All Ministers"
    blackout was created won't automatically be covered by it — that date
    would show capacity again once the new/returning minister is counted
    as available. Acceptable for a one-time event (holiday, building
    closure) decided close to the date; if this becomes a real problem,
    revisit as a genuine shared/global blackout concept instead.
- **Reports tab** (added 2026-09-19; real report content added 2026-09-21,
  corrected the same day against the actual punch-list spec — see
  "Style/tone notes" below on checking the punch-list artifact first): a
  date range picker only — no report-type picker, since
  one Generate always produces all three types together. Date-range presets
  (This Year, Last Year, Last Year + This Year, All Time — same
  `.range-preset` pattern Quick Add Slots uses) fill the From/To fields
  without typing; "Last Year + This Year" (swapped in for the original
  "Last 3 Years" on 2026-09-24 at the user's request — an earlier attempt
  the same day, "Last Year to Date" meaning just last year's own YTD
  slice, wasn't what was wanted) is Jan 1 of last year through *today*,
  covering all of last year plus this year to date in one range — a
  multi-year range, so it also triggers the Total row described below.
  Generate computes against the live `sessions`/`ministers`
  already loaded (Completed only — same retrospective convention
  `completedStats()` uses) and opens a standalone, printable page in a new
  tab (Blob + `window.open`, not a `data:` URI — more reliable across
  browsers for a full HTML document); the new tab has no access back into
  the app, so every number is baked into plain HTML before it opens,
  nothing there re-fetches or recomputes anything.
  - **Three distinct sections on one page** — Freedom Sessions, Sunday
    Drop-In, Training — stacked, never merged, since their totals are kept
    intentionally separate.
  - **Each section is a year-by-year table**, not cards/tiles, newest year
    first: one row per calendar year touched by the chosen range (a partial
    first/last year only counts what's actually in range), Sessions, and
    (Freedom Sessions only) Unique Recipients / First-time / Follow-up
    columns (same per-*recipient*, not per-session, classification
    `completedStats()` uses); Drop-In and Training skip those three — every
    occurrence shares one literal recipName ("Sunday Drop-In"/"Training"),
    so there's no individual-recipient concept to split. A trailing **Hours**
    column (added 2026-09-24) totals `durationHours()` once per *session* in
    that year — not per minister on it, so it isn't the same number as the
    minister-table's own Hours column further down, which sums per minister
    and double-counts a session for every minister who was on it. This is
    "how much ministry time happened," not "how much time any one minister
    gave."
    - **No running-total column.** Shipped one earlier on 2026-09-24
      (shaded background, bold, cumulative from 0 at the start of the
      report's range) then removed it the same day — the shading meant to
      set it apart from the plain Sessions column just read as confusing,
      and the user asked for it gone outright rather than restyled.
    - **A "Total" row instead, only when the range spans more than one
      year** (added 2026-09-24): a single non-expandable `<tr
      class="total-row">` appended after the last (oldest) year, shaded
      like the old running-total column was, with the whole range's own
      Sessions/recipient-split/Hours — computed fresh from the *entire*
      from/to session list via `reportSessionsFor()`/`reportRecipientSplit()`,
      not summed from the per-year rows, so a recipient who shows up in two
      different years is still counted once in Unique Recipients rather
      than twice. A single-year (or partial-year) range shows no Total row
      at all, since it would just repeat that one year's own numbers.
  - **Click a year row to expand it** (added 2026-09-24) into month rows
    **in that same year-table** (one `<tr class="month-row">` per month,
    January through the current month for the current year — including
    that month while still partial, not padded out with months that
    haven't happened), each filling in the *same columns* the year row has
    (Sessions, and for Freedom Sessions also Unique Recipients/First-time/
    Follow-up) rather than a separate nested mini-table off to the side.
    Directly underneath the year row — its own `<tr class="detail-row
    minister-section-row">`, always visible, never hidden behind
    expanding the year's months first (changed 2026-09-27 at the user's
    request — it wasn't clear whose minister breakdown you were looking
    at when you had to open the year to even find the toggle) — sits that
    year's own "Show minister breakdown for {year}" toggle, explicit year
    number included in the label for the same reason. Deliberately **not**
    one breakdown merged across the whole selected range; a multi-year
    range repeats Year → its Ministers → next Year → its Ministers, a
    subtle top border on the minister row marking it as belonging to the
    year above rather than blending into the row below. The year row's
    own `data-group="y-{type}-{year}"` toggle (`toggleGroup(id)`) now
    covers *only* the month rows — the minister row isn't in that group
    and keeps whatever open/closed state the admin left it in regardless
    of the year being expanded or collapsed (including via the "Expand/
    Collapse All Years" buttons, which likewise never touch it).
    - **The minister breakdown itself stays a separate nested `<table
      class="minister-table">`** (tried folding it into the year-table's
      own rows on 2026-09-24 — reverted the same day: the user wanted the
      minister list visually separated as its own sub-table, not blended
      into the year rows). It's themed with the exact same palette as the
      year-table (`#f1ede4` border color, uppercase `#6b6b6b` header
      treatment, `#aa5a3c` chevron) instead of the slightly muted sub-table
      colors it originally shipped with, so it reads as the same table
      family even though it's structurally its own `<table>`. Columns:
      Sessions, then — for Freedom Sessions only — the same **Unique
      Recipients/First-time/Follow-up** split the top-level year/month
      rows show (added 2026-09-24, computed from just that minister's own
      sessions via `reportRecipientSplit()`, not the whole year's list),
      then **As Lead** and **Hours** (`durationHours()` summed — a
      session's full duration credited to everyone on it, not divided
      among them) — As Lead stays even on Freedom Sessions where the
      recipient-split columns are also present, it's never replaced by
      them. Tags anyone whose current PrayerMinisters Status isn't Active
      with a small "Inactive" label (the breakdown includes everyone with
      a session in that year regardless of current roster status). Each
      *minister* row is itself click-to-expand (single-row `toggleRow(id)`,
      not `toggleGroup`) into their own month-by-month numbers — a small
      nested `<table class="month-table">` with the same column set as the
      minister row itself (Sessions, the Freedom-Sessions-only recipient
      split, As Lead, Hours) — one structural level further down since
      it's inside an already-separate table.
    - The page's own Print button removes `hidden` from every collapsed
      element before calling `window.print()` — collapsed content doesn't
      print, so this is load-bearing, not cosmetic; re-verified against
      the full nesting (year → month rows in-table; year → minister
      toggle → minister-table rows → each minister's own month-table)
      after the minister breakdown moved back out to its own table.
  - Each section shows its own empty-state row when it has no completed
    sessions in range, rather than a misleading all-zeros table.
  - **Mobile: every table scrolls horizontally within its own card**
    (added 2026-09-24, `.table-scroll` div wrapping each `<table>`, plus a
    `min-width` on the table itself so columns keep their natural width
    instead of getting squeezed illegibly thin) — a Freedom Sessions
    year-table can run 7 columns and the minister-table 8, too wide for a
    phone screen no matter how tight the padding gets, and it was
    overflowing past the card's right edge before this. Also added a
    `<meta name="viewport" content="width=device-width, initial-scale=1">`
    to the generated report's own `<head>` (it had none — without it a
    phone renders the page at a zoomed-out desktop-width viewport instead
    of actual screen width), and a `@media (max-width:640px)` block that
    trims the card's own padding and the drill-down indentation so more
    of that limited width goes to content.
  - **Gotcha hit while building this**: the generated report page embeds
    its own `<script>...<\/script>` (for `toggleRow`) inside the *main
    app's* template-literal string. Writing a literal `</script>` inside
    that string — even though it's just string content to the JS parser —
    makes the **browser's HTML tokenizer** end the outer `<script>` tag
    right there, since HTML parsing doesn't know about JS string
    boundaries; everything after it in the file silently stops being
    treated as script. Must stay written as `<\/script>` (or equivalent)
    wherever a literal `</script>` needs to appear inside a `<script>`
    block's own source text. See "Known gotchas" #14.
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
  SharePoint's own list-level permissions, granted to the "Prayer
  Ministers" **Entra security group** (`ENTRA_MINISTERS_GROUP_ID`, see
  gotcha below — group *membership* is automated as of 2026-09-27, but
  those SharePoint permission grants themselves are still configured
  directly in SharePoint, not in this code or in Entra):
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
- **Creating the actual Entra account, adding it to the "Prayer Ministers"
  Entra group, and resetting an existing account's password are all
  automatable from Add/Edit Minister** (added and promoted to production
  2026-09-27 — see "Add/Edit Minister can create and manage the actual
  sign-in account" under "Features implemented" above), for whichever
  admin's own signed-in account holds a role that can manage users AND
  is an owner of (or otherwise permitted to manage membership of) the
  `ENTRA_MINISTERS_GROUP_ID` group. Two manual, one-time setup steps still
  have to happen outside this repo before that works at all:
  1. **Grant admin consent for THREE delegated Microsoft Graph scopes** on
     the Scheduler's Azure AD app registration (Client ID
     `549f5207-d7f8-4924-9bde-30532d90c1d2`) — Entra admin center → App
     registrations → this app → API permissions → Add a permission →
     Microsoft Graph → Delegated → add `User.ReadWrite.All`,
     `Directory.AccessAsUser.All` (see "Features implemented" above for why
     account creation alone doesn't need this one but password resets do),
     and `GroupMember.ReadWrite.All` (needed to add/remove the Entra group's
     membership) → **Grant admin consent for [tenant]** once, covering all
     three (same Global-Admin-or-equivalent requirement as the original
     `Sites.ReadWrite.All` consent, gotcha #7).
  2. **Point the SharePoint site's permissions at the Entra group** instead
     of (or in addition to) any individual users — the site's actual
     access grant has to go to `ENTRA_MINISTERS_GROUP_ID` for group
     membership to mean anything; this is a normal SharePoint "Advanced
     permissions settings" change, not something this app can do.

  No SharePoint schema change is needed for any of this (the password is
  never stored — see above). With both steps done, adding a minister to
  the team really is just filling out the form — no manual group step left
  afterward, and (unlike the classic-SharePoint-group version this
  replaced) no multi-minute wait either.
- **Sign-in redirect, both directions**: whichever of the two apps you sign
  into, a mismatched account gets redirected to the other one instead of
  landing somewhere with nothing useful for them. Neither check is a
  security control — SharePoint/Entra group permissions remain the real
  access boundary — this is purely a UX nicety so one URL can be handed out
  to everyone and they land in the right place.
  - **Scheduler → Portal** (production, in `index.html`): a minister who
    signs into the *main scheduler* gets redirected straight to `/portal/`
    instead of landing in the admin UI. `afterSignIn()` checks the
    signed-in account's email against every minister's SignInEmail right
    after resolving the SharePoint site, before ever showing the admin UI;
    a match redirects via `window.location.href` (relative `portal/`
    path), no match proceeds into the admin app as normal.
  - **Portal → Scheduler** (`portal/index.html`, added 2026-09-27 at the
    user's request — previously an admin who signed into the portal by
    mistake had to sign out and separately sign into the Scheduler):
    mirrors the check above in reverse. The portal's own `afterSignIn()`
    resolves the site/lists, checks the signed-in account's email against
    every minister's SignInEmail, and redirects to `../` (the Scheduler)
    on NO match, before ever showing the portal UI — same
    resolve-then-check-then-redirect shape as the Scheduler's own version,
    including its `#loadError` limitation (that element lives inside the
    still-hidden main UI, so an error from `resolveSiteAndLists()` itself,
    before the redirect check even runs, is written somewhere not yet
    visible — a pre-existing quirk in the Scheduler's version too,
    deliberately mirrored rather than fixed here as an unrelated change).

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

## Intake Forms (`/intake/`, `js/formConfig.js` + `js/formEngine.js`)
Digital replacement for the ministry's old Microsoft Forms intake
questionnaire — 88 questions across 17 sections, with real branching logic
(marital status swaps in a whole different sub-section; ~20 "if Yes, please
explain" conditionals) reverse-engineered from the original form. Originally
built as a separate project/repo (`FPHM Intake Form`), merged into this repo
2026-09-24 once it became clear intake responses needed to be visible from
both the staff scheduler and the minister portal, matched to a recipient's
existing session history.

- **`js/formConfig.js`** holds the HARDCODED DEFAULT question schema —
  used only as a fallback if the live schema (see below) can't be fetched
  (Worker unreachable, malformed data). As of 2026-09-25 it is **not**
  the actual source of truth anymore: staff edit questions live through
  the "Edit Intake Questions" tab (`preview/index.html`), which writes to
  the `IntakeFormSchema` SharePoint list, not this file. The shape
  (`visibleIf: {id, equals/in/includes/notEquals}` on a question or a
  whole section, mirroring how the marital-status sub-sections work) is
  unchanged and still documented in this file's own header comment — the
  live editor produces schemas in exactly this shape, and the fallback
  needs to stay structurally compatible with whatever the editor can
  produce. Every app that touches intake data (`intake/index.html`,
  `index.html`, `preview/index.html`, `portal/index.html`) loads it via
  `<script src="../js/formConfig.js">` — a deliberate, sole exception to
  this repo's "single self-contained file" convention, since duplicating 88
  questions across four files would directly defeat "one place to edit
  questions."
- **`js/formEngine.js`** — shared visibility/validation logic plus
  `FPHM.renderResponsesHtml(responses)`, which every intake-answer view
  (public print/PDF, staff detail, minister modal) calls to turn a stored
  `responses` object into the same formatted HTML, styled by the `.ans-*`
  CSS rules each app defines locally (matching its own palette). Both
  print consumers (`intake/index.html`'s own `#printArea`, and
  `preview/index.html`'s Intake Forms tab detail view) give `.ans-row` a
  print-only two-column layout (added 2026-09-26) - a narrow ~30% label
  column plus a wide answer column, since a long free-text answer wraps
  far more than its label does, and stacking label-above-answer (the
  on-screen layout, unchanged) used noticeably more vertical space and
  pushed the printout to more pages than it needed. This also defined
  `.no-print`/`@media print` for `preview/index.html`'s main app for the
  first time - the `no-print` class was already sprinkled through its
  markup (including on `openIntakeDetail`'s own toolbar) but had never
  actually done anything outside the separate Reports-generator page,
  since no CSS rule for it existed there. Every
  internal reference to `SECTIONS` reads `window.FPHM_INTAKE_CONFIG.SECTIONS`
  live on each call rather than a value captured once at load time — this
  is what actually makes the live-schema override below work: swapping
  `window.FPHM_INTAKE_CONFIG` after this script has already run takes
  effect on the very next render, no reload needed. `FPHM.applyLiveSchema
  (schema)` does that swap (accepts a parsed object or a JSON string, throws
  if it doesn't look like a real schema so a caller can catch it and keep
  the hardcoded default instead of rendering a broken form).
- **Live schema loading (staff-editable, added 2026-09-25)** — the actual
  question schema lives in the `IntakeFormSchema` SharePoint list (see
  above), not in code. Each app fetches it differently, matching how it
  already talks to SharePoint elsewhere: `intake/index.html` calls
  `fetchLiveSchema()` in `boot()`, which hits the Worker's `GET
  /api/intake/schema` (the Worker caches the parsed schema in memory for 5
  minutes per warm isolate — a saved edit can take up to that long to
  reach a given edge location, not instant); `preview/index.html`/
  `index.html`/`portal/index.html` fetch it as just another list via their
  existing delegated Graph session in `loadAll()`, no new permission
  needed. All three call `FPHM.applyLiveSchema()` on success and silently
  keep the `js/formConfig.js` default on failure — same offline-tolerant
  philosophy as everything else here. `INTRO_TEXT`/`LIABILITY_TEXT` in
  `intake/index.html` are read via `introText()`/`liabilityText()`
  functions rather than a destructured constant, for the same
  live-override reason as `formEngine.js`'s `SECTIONS`.
- **"Edit Intake Questions" tab** (`preview/index.html`, staff-facing,
  not yet promoted to production `index.html`) — add/remove/reorder
  sections and questions, edit labels/type/required/options/min-max, and a
  guided (not raw-JSON) picker for `visibleIf` branching: pick a target
  question (only ones earlier in the form — a forward reference to an
  unanswered question can't drive visibility), a comparison
  (equals/notEquals/in/includes), and a value (a dropdown of the target's
  own options where it has fixed choices, free text otherwise). Edits a
  working copy (`schemaDraft`, deep-cloned on load) — nothing touches the
  live schema until "Save changes", which validates first (no duplicate
  question ids; every `visibleIf.id` must reference a question that still
  exists) and refuses to save with a specific error message if either
  check fails, PATCHes `IntakeFormSchema`, then calls
  `FPHM.applyLiveSchema()` locally so this app's own Intake Forms tab
  reflects the change immediately too, not just the public form on its
  next fetch. Text-field edits (label wording, option lists, intro/
  liability text) update `schemaDraft` in place WITHOUT a full re-render —
  same lesson as the public form's own answer inputs (see git history,
  index.html was originally re-rendering on every keystroke and losing
  focus/cursor position); only structural changes (add/remove/reorder/type
  change/`visibleIf` target change) call `renderFormEditor()`.
- **`intake/index.html`** — the public form. No sign-in (recipients have no
  Microsoft 365 account); autosave + a "Save & continue later" link
  (`?token=<uuid>`) with graceful local-only fallback if the network drops;
  canvas signature pad (mouse + touch); print/PDF view of the final answers.
  Talks to a Cloudflare Worker (`cloudflare-worker/`: `wrangler.toml`,
  `src/index.js`, `DEPLOY.md` — not part of this GitHub Pages deployment,
  deployed separately via `wrangler deploy`, live and verified working at
  `https://fphm-intake-func.ajjamoore.workers.dev`), which holds its own
  tightly-scoped app-only Graph credential (`Sites.Selected`, granted to
  just this one SharePoint site) so it can create/update/submit the
  recipient's session without any staff/minister credential ever touching
  a browser the public can reach. **Originally built as an Azure
  Function** (`azure-function/SETUP.md` in this repo still has that
  abandoned setup guide, preserved 2026-09-25 when the standalone `FPHM
  Intake Form` project's GitHub repo was deleted — see "Not yet done"
  below — kept only as historical reference, not the live path) —
  switched to Cloudflare Workers the same day after discovering the
  account with the right Entra/SharePoint roles had no Azure
  subscription, and the user didn't want to open one just for this;
  Cloudflare was already in use for the WhatsApp notification feature, so
  this reuses that same free-tier account rather than adding a new one.
  The `Sites.Selected` grant to the "FPHM Intake Form" Entra app
  registration (Client ID `5580dcc8-7837-48fe-a0cf-92eda3959cb0`) carried
  over unchanged — only where the credential runs changed, not the
  credential itself or the SharePoint-side permission model.
- **Staff view** — a new "Intake Forms" tab, currently in
  `preview/index.html` only (not yet promoted to production `index.html` —
  see "Not yet done" below): Submitted/In-progress list, search, full
  detail view, print/PDF. Reads `IntakeResponses` directly via the same
  delegated Graph session as every other tab (`Sites.ReadWrite.All`,
  already consented) — no Worker involved for reads, staff already have
  real permissions.
- **Recipient matching** — same convention as the existing "recipient
  session history" panel in the session modal: matched by name only (no
  persistent recipient record to join on). `findIntakeForRecipient()` in
  `preview/index.html` powers a banner in the session modal
  (`renderIntakeMatchPanel()`, hooked into `renderRecipientHistory()`) that
  jumps straight to a matched intake form.
- **Intake form linking (added 2026-09-26)** — the plain name-match above
  is only ever a suggestion now, not the final word: a Booked slot's card
  and each Waiting/Unplaced Candidates entry in the Schedule & Planning
  tab show a persistent `renderIntakeLinkBlock()` panel that reads/writes
  that session's own `IntakeResponseId` column (see the PrayerSessions
  schema above). If unset, it shows the name-match suggestion with a
  one-click "Link this form"; either way, "Change"/"Choose a different
  form" opens a small search-by-name-or-email picker
  (`renderIntakeLinkPicker()`) for the misspelling/duplicate-name case,
  and "Unlink" clears it back to the plain suggestion. A linked form whose
  own name doesn't match the recipient's is still shown (staff explicitly
  chose it — an override, not a bug) but flagged with a visible mismatch
  warning. `portal/index.html`'s `findMyIntakeForRecipient()` prefers the
  same link when one of the minister's own sessions for that recipient has
  it set, falling back to its own name match otherwise. **Requires the
  IntakeResponseId column to actually exist in SharePoint** — added
  manually by the user (same as Priority/Contacted before it, see history
  above), not by this app; the mapping code reads it safely as null if the
  column isn't there yet, same convention as those two.
- **Preventing accidental duplicate in-progress forms (added
  2026-09-26)** — the two-device case (someone starts the form, doesn't
  save/use their resume link, then opens it fresh on another device) used
  to silently create two separate InProgress records for the same person.
  Once the Personal Information section (name, email, country of birth)
  is complete, the public form calls the Worker's new
  `/api/intake/find-duplicate` endpoint; if another InProgress record
  matches all three fields, it offers "Continue that one" (switches over
  via `resumeWithToken()`, the same load path a saved resume link uses)
  or "Keep this new one". Deliberately requires all three fields to
  match, not name alone — a name-only match would let anyone who knows or
  guesses a name pull up a stranger's in-progress form and its answers,
  which this data's sensitivity can't take on. Staff can clean up
  whatever's left from before this existed (or the rare case someone
  declines to resume) via the "Delete this in-progress form" action on
  the Intake Forms tab's In-progress list (`deleteIntakeForm()`) — offered
  only for InProgress forms, never Submitted ones. That delete action
  surfaced a real bug the same day: a visitor whose token got deleted
  this way came back to a 404 on load, which the code treated exactly
  like a network failure - it fell into "offline" mode and just stayed
  there forever, since the record it was waiting to sync to no longer
  existed. Fixed by having `apiClient.js`'s `post()` tag a failed
  request with its HTTP status, so `boot()`/`resumeWithToken()`/`doSave()`
  can tell "genuinely offline" apart from "that token doesn't exist any
  more" and call `recoverFromDeletedToken()` on the latter - gets a fresh
  token and re-adopts whatever's saved locally under the old one (or
  lands on a clean Intro screen if there's nothing local), instead of
  leaving the visitor stuck.
- **Visitor can delete their own in-progress form** (added 2026-09-26) —
  a "Delete my answers" link sits next to "Save & continue later" in the
  sticky top bar while filling the form out, calling
  `confirmAndDeleteForm()`, which hits the Worker's new `POST
  /api/intake/delete { token }` (`handleDelete()`) and then resets to a
  clean Intro screen. Trust boundary is the same one load/save/submit
  already use: knowing the token (a random UUID) is proof enough, no
  separate check. **InProgress only, not Submitted** - originally also
  offered on the thank-you screen after submitting, removed the same day
  at the user's explicit request: once a form is Submitted, staff/
  ministers may already be relying on it for a scheduled session, so the
  recipient shouldn't be able to pull it out from under them. Enforced in
  both places, not just by hiding the button - `handleDelete()` itself
  returns 403 for a Submitted item, so it can't be bypassed by calling
  the endpoint directly.
- **Long free-text answers auto-grow instead of clipping/scrolling**
  (added 2026-09-26) — every `<textarea>` (the "please explain"
  questions, and the checkboxes "Other, please specify" box) grows to
  fit what's typed via `autoGrowTextarea()`, called on every keystroke
  and once per question on render (so a prefilled multi-line answer
  starts at the right height too). This also fixed a real layout bug:
  the "Other" box used to be a plain `<input>` sized by CSS `width:100%`
  inside a `display:flex` (row) `.choice-option` with no `flex-wrap` -
  without a line break, that width was computed against the whole row's
  space *in addition to* the checkbox and label already sitting there,
  so it visibly overflowed past the card's right edge instead of
  wrapping onto its own line. Fixed by adding `flex-wrap:wrap` to
  `.choice-option` and `flex:1 1 100%` to `.other-text`, which now forces
  it onto its own full-width line the way `margin-top:6px` alone never
  actually did.
- **Minister portal** — `portal/index.html` adds a "View intake form" link
  on a recipient's Schedule/Completed Sessions group, but **only once
  staff has explicitly linked that session to a specific intake form**
  (tightened 2026-09-26, at the user's explicit request — an earlier
  version also fell back to a plain name match when nothing was
  explicitly linked, which meant a minister could see a matching form
  before any admin had approved it; there is deliberately no such
  fallback any more, `findMyIntakeForRecipient()` returns null unless
  `intakeResponseId` is set on one of that minister's own sessions for
  that recipient). `findMyIntakeForRecipient()` also still requires the
  recipient to actually appear in `mySessions()` before it'll match
  anything — **this is a client-side filter only, not a SharePoint-
  enforced boundary**: the Ministers group has Read on the whole
  IntakeResponses list (Graph doesn't reliably enforce per-item permissions
  even for privileged users — see gotcha #11), so a minister with the portal
  open could technically fetch any recipient's answers directly via Graph.
  Explicitly accepted for this data despite its sensitivity (health, abuse,
  addiction, spiritual/occult disclosures) — decided 2026-09-24 rather than
  silently inherited from the BlackoutDates precedent. Revisit if a stronger
  boundary is ever wanted (would mean routing minister reads through a
  server-side check instead of direct Graph access).
  **Submitted forms only** (also tightened 2026-09-26) — an in-progress
  form isn't a finished record yet, and a minister has no business seeing
  a recipient's answers before they've actually finished and signed. The
  view modal is otherwise unchanged read-only (no edit/link/unlink here,
  ever - that stays admin-only in `preview/index.html`) except for a new
  **"Print / Save as PDF"** button
  (`printIntakeView()`) — the modal itself can't reliably print in place
  (a `position:fixed` overlay behaves inconsistently across browsers'
  print engines), so this builds the same document shape as the
  recipient's own printout (answers, then the Liability Release wording
  and signature together at the bottom) into a separate, normally-hidden
  `#intakePrintArea`, isolated for print via a `.printing-intake` body
  class (`visibility:hidden` on everything else, that one area repositioned
  to the page origin) rather than the modal itself.
- **Not yet done**: **not promoted to production** — the merge commit was
  explicitly a preview build; `index.html` has no Intake Forms tab and
  doesn't load `js/formConfig.js`/`js/formEngine.js` yet (see "Architecture"
  above). Follow the normal "Deployment workflow" above to promote once
  it's been tested at the live preview URL. **The standalone `FPHM Intake
  Form` project's GitHub repo (`FPHM2026/prayer-intake`) was deleted
  2026-09-25** (at the user's request, this merge already confirmed
  complete) — its Azure deploy guide was copied into this repo first as
  `azure-function/SETUP.md` before deletion, so nothing was lost, even
  though that path was abandoned for Cloudflare the same day; the local
  folder (`H:\My Drive\Claude\FPHM Intake Form`) still exists on disk as
  an unlinked historical reference if ever needed, just not on GitHub
  anymore. Unlike that abandoned path, **the Cloudflare Worker actually
  is deployed and live-verified** — `intake/js/apiClient.js`'s
  `FUNCTION_BASE_URL` points at it for real, so the public form is no
  longer running in its offline fallback mode for real users. The **"Edit
  Intake Questions" tab is also preview-only so far** — same promotion
  step needed. It can edit `INTRO_TEXT` (the framework-explanation intro
  screen, added 2026-09-25 - `renderFormEditorIntro()`, its own small
  editor above the sections list since it isn't a question and doesn't
  live in `SECTIONS`) and, as of 2026-09-26, `LIABILITY_TEXT` too (the
  liability release wording on the final signature screen -
  `renderFormEditorLiability()`, same pattern as the intro editor). Both
  became rich text later the same day - a `contenteditable` box
  (`.rte-editable`) with a `document.execCommand()`-based toolbar
  (`richTextToolbarHtml()`: bold/italic/underline/lists/clear formatting,
  no external library) instead of a plain textarea, storing real HTML in
  `INTRO_TEXT.body`/`LIABILITY_TEXT.body` rather than an array of
  paragraphs or a plain pre-wrapped string. `FPHM.introBodyHtml()` /
  `FPHM.liabilityBodyHtml()` (`js/formEngine.js`) render that HTML and
  transparently upgrade either legacy shape the first time it's loaded
  into the editor or shown on the public form, via an `isHtmlBody()`
  heuristic (a real tag present means "already migrated"). Rendered
  trusted/unescaped everywhere - safe, since it's staff-authored through
  this same editor, the same trust level as every other part of the
  schema staff already fully control. The public form's own print/PDF
  view (`#printArea` on the thank-you screen) now also shows the
  Liability Release title, wording, and the recipient's own signature
  image - added 2026-09-26 alongside the rich-text change (initially
  placed above the answers, moved to the bottom next to the signature
  the same day per explicit request, so the printout reads answers
  first then what was agreed to and signed, together, at the end).
  Needed two follow-on changes to actually have a signature available to
  show: the Worker's `handleLoad()` now returns `signatureDataUrl` too
  (previously only `handleSubmit()`'s request body ever saw it - a
  revisit via a resume link had no way to get it back), and
  `state.signatureDataUrl` is now populated on a successful submit, not
  just the offline-fallback path that already saved it to
  `localStorage`. The same 2026-09-26 change also added a "type your name instead" mode to
  the public form's signature step (`js/formEngine.js`'s
  `attachSignaturePad` gained `setMode()`/`setTypedText()`): a visitor can
  toggle between drawing with mouse/touch and typing their name, which
  gets rendered onto the same signature canvas in a cursive Google Font
  ("Dancing Script", loaded via `<link>` in `intake/index.html`) so the
  stored `SignatureDataUrl` stays a plain PNG data URL either way — every
  downstream consumer (staff Intake Forms tab, minister portal modal)
  needed zero changes. There's also no raw-JSON fallback for a
  `visibleIf` shape the guided picker can't express (per the explicit
  2026-09-25 decision to keep the picker guided-only rather than exposing
  the schema directly); either would mean hand-editing `js/formConfig.js`'s
  shape or extending
  the editor later if that's ever actually needed.

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
    leave orphaned if only the rule itself gets removed) — production never
    had this one. Check it by hand after every promotion; nothing catches
    this automatically.
    **Historical note, no longer applicable as of 2026-09-27**: this gotcha
    used to also cover stripping the **Changes tab** (`panel-changes`
    section, its nav button, `CHANGE_LOG`/`renderChangeLog()`,
    `wireApp()`'s `$('#changesSubtabs')` binding) out of every promotion —
    missing any one piece of that surface threw `TypeError: Cannot read
    properties of null (reading 'querySelectorAll')`, caught by
    `afterSignIn()`'s try/catch and shown to every signed-in user as a
    misleading "Could not load data... Check the SharePoint site URL"
    error. That whole strip step is now **moot**: the Changes tab ships in
    both preview and production unconditionally as of the 2026-09-27
    promotion (see step 1 of the promotion process above) — there is
    nothing left to strip or grep for on this front.
    **Hit again on 2026-09-26/27, different bug, same root cause**:
    `preview/index.html` lives one folder deeper than the repo root, so
    every path it references to something also at the repo root has to
    start with `../` — `PORTAL_URL`, `INTAKE_URL`, the `<script
    src="../js/formConfig.js">`/`formEngine.js` tags, and the "Open
    Intake Form" link's `href="../intake/"` all correctly say `../` in
    preview. Every one of them is **wrong** once wholesale-copied into
    `index.html` at the actual repo root — `../js/formConfig.js` from
    there resolves one level *above* the repo, a silent 404 (the intake
    schema/engine just never loads, breaking the Intake Forms tab with no
    obvious error) rather than a crash. This will keep recurring on every
    future wholesale promotion for the same reason `PORTAL_URL` does —
    **after every promotion, grep production `index.html` for `\.\./`**
    (a literal `../`) and fix every hit *except* the one harmless comment
    at the top of the file that just mentions `(../js/)` descriptively.
14. **A literal `</script>` inside a JS string breaks the *enclosing*
    `<script>` tag, even though it's just string content to the JS parser.**
    The Reports feature builds a whole standalone HTML page (with its own
    `<script>...</script>` for `toggleRow`) inside a template literal in
    this app's own script. The **browser's HTML tokenizer** scans for the
    literal character sequence `</script` to end a script element — it has
    no concept of JS string/template-literal boundaries, so the first
    `</script>` anywhere in the raw file, even one meant as plain string
    content three template-literals deep, silently ends the outer
    `<script>` tag right there. Everything after it in the file stops being
    treated as script (functions defined below it just don't exist -
    `ReferenceError: X is not defined` on things that are clearly right
    there in the source), with no build step or linter to catch it, since
    the file is syntactically valid JS on its own and only breaks once
    embedded in HTML. Fix: write it as `<\/script>` (or any other escape
    that changes the source text without changing the resulting string)
    anywhere a literal `</script>` needs to appear inside a `<script>`
    block's own content. Applies symmetrically to nested `<style>` tags
    too, though this codebase hasn't hit that case yet.

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

## Working on this repo across multiple chat sessions
This repo now has several features developed in parallel across different
Claude Code chats at once (e.g. the Intake Forms merge and a WhatsApp
booking-notification feature landed around the same time, 2026-09-24/25).
Two sessions editing the same checked-out working directory at once step on
each other's uncommitted changes — this was discovered the hard way when one
session's readiness check turned up another session's in-progress,
uncommitted `preview/index.html` edits sitting in the shared directory.

**Standing policy: start a new feature in its own git worktree**, not the
main checkout. Use the `EnterWorktree` tool (name it after the feature,
e.g. `intake-forms`, `whatsapp-notify`) at the start of a session doing
non-trivial new work in this repo — this project's own CLAUDE.md content
counts as the "project instructions" that tool checks for, so a session
reading this file should treat starting a worktree as the default for new
feature work here, not something that needs to be asked for every time.
Small doc-only or single-line fixes to the main checkout are fine without
one. When a feature is done, merge its branch into `main` (or `preview`,
per the promotion workflow above) normally and clean up the worktree.
