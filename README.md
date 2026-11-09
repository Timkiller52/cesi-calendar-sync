# CESI Calendar Sync

Read-only converter from the publicly published CESI Outlook timetable to an Apple Calendar-compatible ICS subscription. GitHub Actions regenerates the feed about every 30 minutes; GitHub Pages hosts it.

## Setup

1. Create a **public** GitHub repository named `cesi-calendar-sync` (no initial README).
2. Upload `sync.py`, `README.md`, and `.github/workflows/update.yml` while preserving paths, or push them with Git.
3. Under **Settings → Pages**, set **Build and deployment → Source** to **GitHub Actions**.
4. Under **Actions → Update CESI calendar**, select **Run workflow**.
5. When the workflow succeeds and Pages deploys, subscribe to `https://YOUR_USERNAME.github.io/cesi-calendar-sync/calendar.ics` using Apple Calendar's **File → New Calendar Subscription** (Mac). Choose iCloud as the location to sync to iPhone.

## Notes

- Anyone who discovers your GitHub Pages URL can read the timetable. Do not use this with confidential calendars.
- The endpoint is an undocumented published Outlook endpoint and may change or stop working.
- Scheduled GitHub Actions runs may be delayed or skipped, and Apple Calendar controls its own refresh cadence.
- The script fails rather than overwriting a working feed if fetching fails or no events are returned.
- Rolling range: previous 30 days to next 365 days; older events disappear as they fall out of this range.
- If you change the repository name, update the subscription URL.
- If an event's Exchange ID changes on modification, Apple Calendar may treat it as a new event.
