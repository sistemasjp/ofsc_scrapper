# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-file Playwright scraper (`ofsc_scraper.py`) that automates Oracle Field
Service Cloud's "Consola de despacho" (dispatch console) for Claro Ecuador
(`claro-ec.fs.ocs.oraclecloud.com`). For a given date and list of crews
(`crews`), it walks each crew's activities, opens the ones with state
"finalizada", and records the materials/equipment consumed (with SAP code,
quantity, serial, city) into an Excel workbook.

The core trick: instead of scraping the rendered HTML inventory table, it
intercepts internal JSON responses the page itself already receives. **This
is an undocumented internal mechanism, not the public OFSC REST API** — it
was reverse-engineered live against this specific instance/version and can
break silently if Oracle ships a UI update. If that happens, expect
`page.expect_response(...)` or a locator `.wait_for(...)` to start timing
out; re-derive the real endpoint/selector with DevTools (F12 → Network /
Elements) rather than assuming the old one still applies — this codebase's
history is full of "the endpoint changed under us" surprises (see below).

## Running it

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

Manual login (recommended — supports SSO/MFA): opens a visible browser, you
log in yourself, then press ENTER in the terminal to continue.

```bash
python ofsc_scraper.py --date 2026-09-08 --output consumo_2026-09-08.xlsx
```

Unattended (only works for simple username+password logins, no SSO/MFA).
Validate once with `--auto-login` and a visible window before going headless:

```bash
cp .env.example .env   # fill in OFSC_USERNAME / OFSC_PASSWORD; loaded automatically via python-dotenv
python ofsc_scraper.py --date 2026-09-08 --auto-login
python ofsc_scraper.py --date 2026-09-08 --auto-login --headless
```

`.env` is optional — exporting the env vars yourself, or leaving them unset
(getpass prompt), still works. `.env` is git-ignored; never commit it.

Short command for a non-technical run (used for handoff/cron): `./reporte.sh
2026-09-08` — equivalent to `--auto-login --headless --output
consumo_2026-09-08.xlsx`, and refuses to run with a clear message if `.env`
doesn't exist yet.

Resume after an interruption (skips crews already present in the "Resumen
Actividades" sheet of the target Excel file):

```bash
python ofsc_scraper.py --date 2026-09-08 --output consumo_2026-09-08.xlsx --resume
```

There is no test suite, linter, or build step in this repo — it's a single
script exercised by running it against the live OFSC instance. `python3 -m
py_compile ofsc_scraper.py` is the only available sanity check (syntax
only — it won't catch a broken selector or a changed endpoint).

## Configuration (`config.json`)

- `crews` — crew codes exactly as they appear in the left tree of the
  console.
- `skip_activity_types` — administrative activity types (lunch, warehouse
  start/end, slot block) that never have consumption; logged to stdout but
  never opened, and **not written to the Excel at all** (neither sheet) —
  same for any activity whose state isn't in `estados_con_posible_consumo`.
  This was a deliberate product decision (see git history / conversation),
  not an oversight: only activities the script actually opened and checked
  appear in "Resumen Actividades", so a row there is proof the check
  happened, even when it found 0 items (e.g. a "Mantenimiento" that
  genuinely installed nothing).
- `estados_con_posible_consumo` — activity states worth opening (default:
  only `"finalizada"`); others are skipped entirely (see above), not
  recorded with 0 items.
- `output_file`, `screenshot_on_error`, `slow_mo_ms`, `nav_timeout_ms` — run
  defaults, overridable via CLI flags where applicable. Note: `headless` in
  config.json is currently **not wired up** — only the `--headless` CLI flag
  controls it; don't assume the config key does anything.

## Architecture notes

- Single file, sections in order: LOGIN, NAVEGACIÓN, EXCEL, MAIN. Keep new
  functionality in the matching section rather than splitting into modules.
- Credentials: manual login never touches disk. `--auto-login` reads
  `OFSC_USERNAME`/`OFSC_PASSWORD` from the environment (populated by `.env`
  via `python-dotenv`, or exported manually) or prompts via `getpass`.
- Login form field detection (`LOGIN_USER_SELECTORS` / `LOGIN_PASS_SELECTORS`
  / `LOGIN_SUBMIT_SELECTORS`, near the top of the file) is instance-specific.
  Confirmed working values for this instance: `#username`, `#password`,
  `#sign-in` — these are listed first, with older generic guesses kept as
  fallbacks. `do_login()` waits for `networkidle` before searching for
  fields because the login form is an Oracle JET app that renders via JS
  *after* `domcontentloaded` — searching too early finds nothing.

### The real data-fetching mechanism (important, non-obvious)

The docstring's original design assumed two dedicated endpoints
(`m=Grid&a=get` for a crew's activity list, `m=Hint&a=activity` for an
activity's inventory). Live investigation on this instance found:

- `m=Grid&a=get&itype=manage&output=ajax` **does** fire when a crew is
  clicked, and `activitiesRows` does have the expected fields (`aid`,
  `aworktype`, `appt_number`, `cname`, `astatus`, and `aworkzone` — the
  latter parsed by `_parse_city_from_zone()` for the CIUDAD column, format
  `"TECNOLOGIA.REGION.PROVINCIA/CIUDAD/PARROQUIA"`).
- `m=Hint&a=activity` **does not fire at all** in this instance/version.
  Inventory data instead arrives through the app's generic offline-sync
  channel, `m=sync&a=write` — the same endpoint used for many unrelated
  things (heartbeats, other screens). Its response's `delta.Inventory`
  carries items from potentially *any* activity currently "in focus" on the
  client, each tagged with `inv_aid`. Critically, a given activity's
  inventory may already have arrived in an *earlier* `m=sync&a=write`
  response (e.g. bundled with a previous activity's load) — waiting for a
  *new* response specifically while opening that activity can see nothing
  and wrongly report 0 items.
- Fix: `InventoryCache` (class in the file) is a **global, session-lifetime**
  listener registered once on `page` (`page.on("response", inventory_cache.listener)`).
  It accumulates every `m=sync&a=write` response's `delta.Inventory` items
  (filtered to `invpool == "install"`) indexed by `inv_aid`, for the whole
  run. `open_activity_and_get_inventory()` just queries this cache after
  opening an activity (and after forcing the "Inventario" tab if still
  empty) instead of listening locally per-call. **Do not revert to a
  per-call local listener** — it silently loses data for activities whose
  inventory synced earlier than expected.
- Item field mapping inside `delta.Inventory` (unchanged from the original
  design, still correct): `"186"` = SAP code (materials only — almost never
  present for equipment, see business rule below), `"187"` = model, `"200"`
  = description, `invtype == 1` → "Equipos" else "Materiales", `invsn` =
  serial, `quantity`, `invid`/key = inventory id, `inv_aid` = activity id,
  `invpool` = pool (`"install"` is what's wanted; there's also a "Cliente"/
  customer pool and a "Recurso"/resource-truck-stock pool that must stay
  excluded).
- **Identification business rule (`_sap_cell_value()`)**: for "Equipos" the
  serial number alone is sufficient identification — OFSC almost never
  exposes a SAP code for equipment in the intercepted data, and that's
  expected, not an error to flag. Only warn/highlight an equipment row if
  the serial is *also* missing. For "Materiales" the SAP code is still
  required (materials rarely have their own serial to fall back on). There
  used to be a manual `equipment_sap_map` config table to backfill
  equipment SAP codes by model — it was removed deliberately (see
  conversation): maintaining it was busywork for something not actually
  required. Don't reintroduce it without being asked.

### UI quirks that broke things repeatedly (all now handled, don't remove the fixes)

- **View mode**: a session with no saved preference (a fresh Playwright
  context) starts the console in "Vista de tiempo" (Gantt), not "Vista de
  lista" which the rest of the script assumes (`div.oj-datagrid-cell`
  selectors, etc.). `ensure_list_view()` forces list view once at startup by
  clicking `button[aria-label="Vista de lista"]` if it isn't already
  selected (checked via the `radio-selected` CSS class).
- **Crew tree starts collapsed**: a fresh session has the whole left crew
  tree collapsed (`button.edt-open.ptplus`); crew nodes inside a collapsed
  group may not even exist in the DOM yet. `expand_crew_tree()` clicks every
  `ptplus` button (re-querying the first match each time, in a loop) until
  none remain, then waits for `networkidle` — expanding it can trigger a
  burst of `m=Provider&a=opentree` requests that measurably slow down the
  *next* `m=Grid&a=get` (~20s cold, observed). `select_crew()` also
  re-expands on demand if the target crew isn't visible (the tree can
  re-collapse — see next point), and filters `button.edt-label:visible` to
  skip hidden/virtualized duplicate nodes.
- **`go_to_console()` must use the local breadcrumb, not the global nav
  link.** The global side-menu link
  (`a.global-navigation-item--activities`) works fine from one level deep
  ("Detalles de actividad") but from two levels deep ("Lista de
  inventarios" — reached via the Inventario-tab fallback) it triggers a
  **full SPA reload** instead of a soft transition. That reload is slow and
  resets the view mode + crew tree expansion, which then breaks *every
  subsequent crew* in the run (cascading failure — this was a real
  multi-hour debugging session). The fix: `go_to_console()` walks back one
  level at a time using the local breadcrumb buttons (`button:has-text("Consola
  de despacho")`, `button:has-text("Detalles de actividad")` — distinct
  elements from the global nav `<a>` with the same visible text), falling
  back to the global link only if no breadcrumb is found.
- **`#plugin-overlay-window`**: a global loading overlay (z-index 9999)
  that appears briefly during page transitions and blocks pointer events
  ("intercepts pointer events" in Playwright's error) if the next click
  lands while it's still up. `wait_no_overlay()` waits for it to become
  hidden; called before clicks in `go_to_console()`, `select_crew()`, and
  `open_activity_and_get_inventory()`.
- **Return-to-list must not be able to crash the whole run.** Going back to
  a crew's list after each activity (`return_to_crew_list()`) wraps
  `go_to_console()` + `select_crew()` in try/except and returns
  `False`/breaks out of that crew's activity loop on failure instead of
  propagating — an earlier version let this raise uncaught and it took down
  the entire multi-crew run on one bad transition.
- Knockout/Oracle-JET rendering can finish *after* `networkidle` (network
  goes quiet before the DOM finishes painting) — this bit both the login
  form and the date-header text. Where it matters, code polls/waits for the
  actual attribute/text rather than trusting a single instantaneous check
  right after a network-idle wait.

- `set_date()` navigates via `div.toolbar-datepicker-wrapper` (buttons
  `aria-label="Anterior"`/`"Siguiente"`, current date read from
  `button.toolbar-date-picker-button`'s `aria-label`/`title`, e.g. "Viernes
  11 Septiembre 2026") — not by counting header buttons positionally
  (fragile, was the original approach and broke).
- The Excel workbook is written incrementally: `autosize_and_save()` runs
  after every crew, not just at the end, so a mid-run crash doesn't lose
  completed crews (combine with `--resume` to pick up where it left off).
  `already_done_crews()` determines what's already done by reading column A
  of the "Resumen Actividades" sheet.
- On a per-activity error, the script screenshots (if
  `screenshot_on_error`), then calls `return_to_crew_list()` before
  continuing to the next activity — don't remove this recovery step.

## Future direction (from README, not yet implemented)

If automatic forwarding of this consumption data to SAP or another backend
is ever needed, check whether the OFSC account has access to the official
Oracle Field Service REST API (Core/Metadata, OAuth) instead of extending
this internal-endpoint scraper.
