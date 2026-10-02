---
name: run-target-escalations-dashboard
description: Run, start, drive, or screenshot the Target Escalation Dashboard (index.html + visits.html) in headless Chromium with synthetic Excel files. Use to check a change in the real page, load the five imports, switch tabs, or take light/dark screenshots.
---

The dashboard is static HTML that reads Excel files in the browser (SheetJS); there is no build and no backend. An agent drives it with `.claude/skills/run-target-escalations-dashboard/driver.mjs`: it serves the repo itself, opens `index.html` in headless Chromium (1440x900, New York time zone), and runs one command per line from stdin. Test data comes from `make_samples.py` in the same folder (fake stores and people).

All paths are relative to the repo root.

## Prerequisites

Node 22 and the global Playwright 1.56 with Chromium in `/opt/pw-browsers` are already in the cloud container; the driver finds Playwright with `npm root -g`. The sample generator needs openpyxl:

```bash
python3 -m pip install -q openpyxl
```

## Make sample files

Dates are relative to today (this week's Monday, LOS Finish in a few days), so regenerate each session instead of reusing old ones. It writes into `./samples` of the current directory.

```bash
mkdir -p /tmp/target-samples && (cd /tmp/target-samples && python3 /home/claude/target_escalations_dashboard/.claude/skills/run-target-escalations-dashboard/make_samples.py)
```

You get `export_ok`, `data_ok`, `presence_ok`, `tracker_ok`, `parts_ok` (good files), broken variants for each (`export_bad_status`, `data_missing_col`, `presence_bad_answer`, ...), and `expected.json` with the counts the page should show (`named_wo` is the work order whose description names a DPCI; it's store T0581).

## Run (agent path)

```bash
S=/tmp/target-samples/samples
node .claude/skills/run-target-escalations-dashboard/driver.mjs <<EOF
open
load export $S/export_ok.xlsx
load presence $S/presence_ok.xlsx
load tracker $S/tracker_ok.xlsx
load parts $S/parts_ok.xlsx
load data $S/data_ok.xlsx
loaded 5
closeimports
search 9160000006
shot 02-escalation
theme dark
tab vis
shot 04-visits-dark
tab ns
shot 05-noshow-dark
load export $S/export_bad_status.xlsx
errors
EOF
```

Whole run takes about 10 seconds. Each command echoes `> command` and its result; the exit code is 1 if any command failed. Screenshots go to `/tmp/target-shots/<name>.png` (set `SHOTS=dir` to change). Read them with the Read tool and look at them.

Commands:

| Command | Does |
|---|---|
| `open` | Load `index.html`, wait for SheetJS and the Store Visits frame |
| `load <slot> <file>` | Pick a file for `export`, `data`, `presence`, `tracker` or `parts`; prints the toast and the file pill |
| `loaded <n>` | Wait until the pill says `n of 5` |
| `closeimports` | Close the import popup |
| `tab esc\|ns\|vis\|dir` | Escalations, No Show, Store Visits, Directory |
| `search <text>` | Type into the Escalations search box, print the `1 of 29` count |
| `click <sel>` / `fill <sel> <text>` / `text <sel>` | Playwright selector actions on the top page |
| `eval <js>` / `veval <js>` | Evaluate in the top page / in the `visits.html` frame |
| `theme light\|dark` | Switch `prefers-color-scheme` |
| `shot <name>` / `toast` / `wait <ms>` / `errors` | Screenshot, read the pending toast, pause, list page errors and console errors |

Lines starting with `#` are skipped. For the full regression checks of earlier rounds, see the project's shared folder at `/mnt/project-files/.notes/roundN-tests/`. Those scripts expect the repo on `:8765` and use `require('playwright')`, so run them with `NODE_PATH=$(npm root -g)`.

## Run (human path)

```bash
python3 -m http.server 8765
```

Then open http://127.0.0.1:8765/index.html. Open `index.html` itself, not `/`: the Store Visits and Directory tabs are `visits.html` in an iframe, and `index.html` loads SheetJS from cdnjs.

## Gotchas

- **The import popup is modal.** It opens on load and stays up until you close it, so `tab`/`click` time out (`page.click: Timeout 30000ms exceeded`) until you run `closeimports`.
- **A good `data` file shows no toast.** The other four imports toast "✓ Loaded …"; `data` only turns its square green, so `load data` prints `toast: (none)`. A broken `data` file does toast (`data: Missing columns: …`). Check the pill or `loaded 5` instead.
- **The pill's wording changes at five:** `4 of 5 Files Loaded`, then `5 of 5 Files Imported`.
- **`search` follows the store across tabs.** After `search 9160000006`, Store Visits opens on store 581 with a "Following Store 581" chip.
- **A failed import keeps the old data.** `export_bad_status` toasts "Nothing was loaded" and the pill stays at 5 of 5.
- **Time zone matters.** Power BI dates are read as serial numbers and Maximo dates carry their own zone; the driver pins `America/New_York` (K's zone) so week boundaries match `expected.json`.
- **SheetJS comes from cdnjs** in `index.html`; the driver serves the repo's `xlsx.full.min.js` for that URL so runs work without the network. `visits.html` already loads the local copy.

## Troubleshooting

- `Run from the repo root (no index.html in …)`: the driver serves the current directory; `cd` to the repo root.
- `ModuleNotFoundError: No module named 'openpyxl'`: run the pip line above.
- `ERROR unknown command`: the message lists the known commands.
