# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

- **Admins (ministry coordinators, e.g. Andrew):** run the whole ministry from the admin Scheduler: book prayer sessions, match recipients to two-minister teams, manage the waiting list and pre-intake, read intake forms, write session and intake call notes, and produce statistics for church leadership. They use it on a computer for planning and on a phone (Home Screen app) on the go.
- **Prayer ministers (PMs):** volunteers on unlicensed Microsoft accounts who use "My FPHM" (the portal) to see their own sessions, confirm or decline attendance, add time off, read linked intake forms and earlier session notes, and (as Lead) write session notes. Mostly on phones, often the installed Home Screen app.
- **Church leadership (lower priority):** receives ministry statistics from the Statistics tab's Word reports; does not use the app.
- **Prayer recipients (PRs):** members of the public who fill in the intake form before their first session. No account, any device, often in a vulnerable moment; the form asks about health, abuse, addiction and spiritual history.

## Product Purpose

Coordinate one-on-one Freedom Prayer sessions between recipients and the ministry's team of volunteer prayer ministers, replacing spreadsheets, Microsoft Forms and Power Apps. Success means every recipient moves from enquiry to intake to a booked session with the right team without anyone falling through the cracks, ministers always know where they need to be, and leadership gets accurate statistics.

## Positioning

Built for this one ministry's own process (Freedom Sessions, Sunday Drop-In, Journey to Freedom, two-minister teams with a Lead, same-sex team preference, pre-intake to waiting list to booked), on the church's existing Microsoft 365 and SharePoint, with no per-seat licence cost for volunteers.

## Operating Context

- Four pages on GitHub Pages: admin Scheduler (`index.html`, with `preview/` kept in sync), My FPHM portal (`portal/`), public intake form (`intake/`), and a demo with invented data (`demo/`).
- Data lives in SharePoint lists via Microsoft Graph; sign-in is Microsoft 365. A Cloudflare Worker handles the public intake form and push notifications.
- Admin and portal are installable Home Screen apps; push notifications reach ministers' and admins' phones.
- Releases go straight to production with a version number and a Changes tab entry; the public intake form needs Andrew's OK first.

## Capabilities and Constraints

- Plain HTML/CSS/JS in single files, no build step or framework.
- Everything must work at every viewport, phone through desktop (Andrew's standing rule).
- No icon-only controls; every icon has a visible text label.
- 8px spacing scale, with 4px only between an icon and its label.
- No sideways-scrolling tab bars (Andrew rejected it).
- Shared intake schema in `js/formConfig.js` / `js/formEngine.js`; questions are edited in the app, not in code.
- Vocabulary: PR = prayer recipient, PM = prayer minister, "My FPHM" = the minister portal, "Session #N" = a recipient's Nth session.
- Access limits for ministers are partly client-side only (SharePoint/Graph limitation); this is an accepted tradeoff.

## Brand Commitments

- Name: Freedom Prayer Healing Ministry (FPHM).
- Existing design system artifact: Playfair Display + Jost, terracotta and charcoal, cream page background (kept on purpose as part of the brand).
- Voice: plain, warm, everyday language; no jargon in help text or Changes entries.

## Evidence on Hand

- Real data is private (recipients' names, intake answers and notes are sensitive). Use the demo's invented data for screenshots and examples; never show or invent real people.
- Ministry statistics report spec and Leader's Copy define the reporting numbers.

## Product Principles

1. Nobody falls through the cracks: every recipient is always in a visible state (Pre-Intake, Waiting, Tentative, Booked, Completed).
2. Protect recipients' privacy: ministers see only what an admin has linked or marked visible, and names are hidden on intake forms.
3. Plain and obvious over clever: labelled buttons, plain words, explanations where behaviour isn't obvious.
4. Phones are first-class: ministers mostly use the app on a phone.

## Accessibility & Inclusion

Users of all ages and comfort levels with technology. Being able to make text larger matters (confirmed by Andrew): keep and extend text-size controls (the A−/A+ buttons on intake views) and make sure layouts hold up with larger text. Visible labels, focus rings and touch-friendly targets. No formal standard (e.g. WCAG level) set yet.
