# Best Buy Dashboard v4

Live at https://bestbuy-escalations-v4.pages.dev (Cloudflare Pages).

One page with three tabs:

- **Escalations**: the v3 escalations dashboard (`index.html`).
- **Store Visits**: the Store Work Lookup page (`visits.html`), shown in a frame under the shared header. Its left panel matches Escalations (store search, Filters, count, Clear) over a list of the stores they leave. The list narrows with each digit typed, matching the start of a BDS Store ID or chain store #. Picking a store, or narrowing to one, shows its header with the assignment status chart on the right and its call forms in the table below.
- **Directory**: the rep lookup, in the same frame. Searching a rep by name or person ID, or picking one from the list, shows their contact details from the People report and all their assignments this week (Monday to Sunday).

The banner holds the title, the three tabs and every import: FEDR, Current visits, Previous visits, People report. The bar under it holds the counts for the open tab (the escalation pills on Escalations, the assignment counts for the picked store or the whole list on Store Visits), Summary and Undo all my edits. Current visits (the Priority Pivot) feeds all three tabs; the People report feeds the Directory and the rep phone numbers in the tables. A store typed on Escalations carries over to Store Visits, and one typed on Store Visits carries back. There is no export.

Everything is read in the browser. No workbook data is saved; only display preferences (theme, open tab, open filter panels, visit column mapping) are remembered.

Pushing to `main` publishes the site to Cloudflare Pages (`bestbuy-escalations-v4`) through `.github/workflows/deploy.yml`, which needs the repo secrets `CLOUDFLARE_API_TOKEN` and `CLOUDFLARE_ACCOUNT_ID`.
