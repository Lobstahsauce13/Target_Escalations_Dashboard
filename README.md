# Target Escalation Dashboard

Built on the Best Buy dashboard v4: same framework and styling. One page with three tabs:

- **Escalations**: Target's PM work orders from Maximo (`index.html`). Read only for now.
- **Store Visits** and **Directory**: unchanged from v4 (`visits.html`, shown in a frame under the shared header).

The banner holds the title, the three tabs and every import: Data, Current visits, Previous visits, People report.

**Data** is Maximo's work order export. It's recognized by its column headings (`work_order_number`, `work_type`, `site`, ...), so the file name doesn't matter. Only PM rows are loaded; CM rows are skipped. Status codes: APPR = New, VDECWO = Open, VINPRG = Completed, VDECWC = Closed, not removed. A status code not on that list, a missing column, or a date not in Maximo's format (`09/30/2026 06:28 AM EDT`) stops the import with a message, and nothing is loaded. A work order is past due when its LOS Finish (`repair_by`) has passed and it's New or Open. The list is sorted by LOS Finish, oldest first.

Everything is read in the browser. No workbook data is saved; only display preferences (theme, open tab, open filter panels, visit column mapping) are remembered. There is no export.

Pushing to `main` publishes the site to Cloudflare Pages (`target-escalations`) through `.github/workflows/deploy.yml`. It needs a Pages project named `target-escalations` and the repo secrets `CLOUDFLARE_API_TOKEN` (permission Account > Cloudflare Pages > Edit) and `CLOUDFLARE_ACCOUNT_ID`.
