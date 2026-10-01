# Target Escalation Dashboard

Built on the Best Buy dashboard v4: same framework and styling. One page with three tabs:

- **Escalations**: Target's CM work orders from Maximo (`index.html`). Read only for now.
- **Store Visits** and **Directory**: from v4 (`visits.html`, shown in a frame under the shared header), now fed by the Power BI export.

The banner holds the title, the three tabs and the four imports. Each import box is named for its file and says what it fills in:

- **export** is Maximo's work order export. It fills the escalation list and each work order's details. It's recognized by its column headings (`work_order_number`, `work_type`, `site`, ...), so the file name doesn't matter. Only CM rows are loaded; PM rows are skipped, and a file with no CM rows stops the import. Status codes: APPR = New, VDECWO = Open, VINPRG = Completed, VDECWC = Closed, not removed. A status code not on that list, a missing column, or a date not in Maximo's format (`09/30/2026 06:28 AM EDT`) stops the import with a message, and nothing is loaded. A work order is past due when its LOS Finish (`repair_by`) has passed and it's New or Open. The list is sorted by LOS Finish, oldest first. Each work order shows Problem Code (`problem`), Status Date (`earliest_start`), Report Date (`arrive_by`) and LOS Finish (`repair_by`).
- **data** is Power BI's export (sheet "Export"). It fills each store's rep, DM (`FOM Territory`) and whether the rep visited that week. Only the Target Break-Fix Continuity form is kept, and only the waves that have started (this week and earlier). Stores are matched to Maximo by Chain Store # (the site without its "T"). Power BI has no status column, so a visit is Completed once its report is entered, else Sched this week when its date falls this week. Dates are read as the serial numbers in the file, because SheetJS shifts date cells by the time zone's old offset. A Power BI export is cut off at 150,000 rows, so filter it to Target and the recent weeks. The file is read once, by the Store Visits frame; each escalation's "Rep and visits" card asks the frame for the store's last four weeks.
- **Interactives Program Presence and Functionality** is the report CS emails out. It fills each escalation's "Rep's answers" card: the store's visit date, how many displays are present and working, and every display answered No with the reason given. Only its "Presence and Functionality" sheet is read, from the row headed Visit Date down; the trend and pivot sheets are ignored. Stores are matched to Maximo by Store # (the site without its "T"). A display whose DPCI number appears in the work order's description is listed first and marked. A missing column, or a Present? answer that isn't Yes or No, stops the import.
- **People Report** is the rep directory. It fills the Directory tab's rep details.

Everything is read in the browser. No workbook data is saved; only display preferences (theme, open tab, open filter panels) are remembered. There is no export.

Pushing to `main` publishes the site to the Cloudflare Worker `target-escalations-dashboard`, through Cloudflare's Git integration (Workers Builds) on this repo. It serves `index.html`, `visits.html` and `xlsx.full.min.js`; `.assetsignore` keeps the other repo files off the site.
