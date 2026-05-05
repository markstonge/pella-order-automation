# Pella Order Automation Plan

## Goal

Build an easy desktop application for non-technical users that takes:

- one CSV purchase order
- one XLSX work order

and generates a completed Pella workbook matching the provided completed-sheet examples. The app should show warnings in the app when extraction is uncertain, but it should still generate the workbook unless the input file cannot be read.

Ideal user flow:

1. Open the app.
2. Select the purchase order CSV.
3. Select the work order XLSX.
4. Click Generate.
5. Save the completed workbook.
6. Review the generated workbook.

## Decisions Made

- Version 1 is offline-first.
- Cloud AI is allowed later as an optional fallback for difficult descriptions, but it is not required for the initial build.
- Version 1 supports CSV purchase orders only.
- Version 1 supports XLSX work orders only.
- Legacy `.xls` purchase orders and work orders are out of scope for version 1.
- Example 3 is out of scope for version 1 because both source files are legacy `.xls`.
- The generated workbook does not need macros.
- The output can be `.xlsx` if it matches the completed-sheet format and values.
- `Schedule` and `Date` can be blank for now.
- Users do not need a mandatory pre-generation field review screen.
- Warnings should be displayed in the app, not written to a separate logbook.
- The generated workbook is the main audit artifact.

## Recommended Direction

Use a packaged desktop app with a Python extraction/generation engine.

Current prototype path:

- Python core package for parsing, extraction, validation, and workbook generation.
- Simple Python/Tkinter GUI as an immediate usable prototype.
- Tauri desktop app with a modern React UI, shadcn-style components, Radix primitives, Tailwind CSS, and an offline local generation engine.
- Double-click development launchers for macOS and Windows.
- The Python engine can be bundled as a Tauri sidecar binary, or migrated behind Tauri commands if Rust becomes worthwhile later.

Why this direction:

- Python is strong for CSV/XLSX parsing and workbook writing.
- A local parser keeps the app offline and avoids sending customer/order files to the cloud.
- A desktop wrapper gives users a normal file-picker workflow.
- The extraction engine can be reused behind a polished Tauri UI.

## Version 1 Scope

Inputs:

- Purchase order: `.csv`, using the format shown in Purchase Order Examples 1 and 2.
- Work order: `.xlsx`, using the format shown in Work Order Examples 1 and 2.

Output:

- A generated workbook with `Jamb-BOM` and `Lineal-BOM` sheets.
- Macro preservation is not required.
- `.xlsx` is acceptable unless a downstream requirement appears for `.xlsm`.

Warnings:

- Unknown or low-confidence description parsing.
- Unmatched PO/work-order lines.
- Suspicious duplicate or missing dimensions.
- Complex arch/curved units.
- Unsupported file formats.

Out of scope for version 1:

- Legacy `.xls` purchase orders.
- Legacy `.xls` work orders.
- PDF input.
- Cloud AI parsing.
- Batch processing.
- App signing/notarized installers.
- Perfect visual parity with the provided completed workbooks.

## Extraction Strategy

Use a hybrid deterministic parser rather than one fragile regex.

Stable document parsing:

- CSV PO header and line fields.
- Work order header cells.
- Work order item rows.
- PO xref to work-order item matching.

Flexible description parsing:

- Jamb depth, including mixed fractions like `3-9/16"` and `4-5/8"`.
- Lumber type, such as Poplar, Pine, Red Oak, Maple.
- Finish text and finish code mapping, such as `Primed -> PRIME`, `Unfinished -> NONE`, `Paint Pella White -> PWHITE`.
- Jamb type, including 4-sided windows and 3-sided patio doors.
- Piece count, such as `4 pcs`, `6 pcs`, `16 pcs`.
- Special cases, especially arch and curved units.

Validation:

- PO xref prefix should match work order order number.
- PO xref suffix should match work order `Item#`.
- Piece count should produce plausible generated row counts.
- Jamb dimensions should be present and plausible.
- Unknown finishes/materials should generate warnings.

## Data Model

Order header:

- `order_number`
- `pella_po_number`
- `customer_name`
- `ship_address`
- `ship_city`
- `ship_state`
- `ship_zip`
- `request_date`
- `schedule`
- `date_code`

Source line:

- `source_row`
- `po_line_number`
- `pella_xref`
- `xref_suffix`
- `work_order_item_number`
- `raw_description`
- `raw_comment`
- `raw_location`
- `raw_quantity`
- `raw_rough_opening`

Extracted jamb line:

- `type`
- `jamb_depth`
- `jamb_width`
- `jamb_height`
- `quantity`
- `piece_count`
- `jamb_style`
- `lumber_type`
- `finish_text`
- `finish_code`
- `calculated_jamb_height`
- `calculated_footage`
- `warnings`

Output row:

- `date`
- `schedule`
- `batch`
- `bin`
- `part`
- `loc`
- `width`
- `finish`
- `quantity`
- `length`
- `info`
- `order_number`
- `assembled`
- `part_no`
- `customer`
- `ship_city`
- `ship_state`
- `po_number`
- `customer_ref`
- `ship_address`
- `ship_zip`
- `request_date`

## Work Completed

Created the initial application structure:

- `README.md`
- `pyproject.toml`
- `.gitignore`
- `src/pella_order_automation/`
- `tests/`

Implemented:

- CSV PO parser.
- XLSX work-order parser.
- Mixed-fraction and dimension parsing utilities.
- Extraction model that matches PO lines to work-order items.
- Description parsing for jamb depth, piece count, lumber type, finish, jamb type, and dimensions.
- Warning generation for uncertain/complex lines.
- Workbook writer that creates `Jamb-BOM` and `Lineal-BOM`.
- CLI entry point: `python -m pella_order_automation.cli generate ...`
- Simple desktop prototype: `python -m pella_order_automation.gui`
- Polished Tkinter GUI layout with clearer panels, file selectors, primary action button, status/progress area, and tabbed warnings/errors.
- GUI summary panel showing output path, input files, order number, PO number, customer, request dates, generated counts, and warning count.
- GUI metric cards for jamb line count, generated BOM row count, and warning count.
- GUI warnings panel with deduplicated warning messages.
- GUI error panel showing any failure message and diagnostic details if generation fails.
- Thread-safe GUI result handling so background generation reports back through the main window instead of touching Tk directly.
- Material UI web app:
  - `frontend/src/main.jsx`
  - `package.json`
  - `vite.config.js`
  - `src/pella_order_automation/web_server.py`
- Replaced the Material UI dependency with a Tauri-oriented React UI using Tailwind, Radix primitives, shadcn-style local components, and lucide icons.
- Added Tauri scaffold:
  - `src-tauri/Cargo.toml`
  - `src-tauri/build.rs`
  - `src-tauri/tauri.conf.json`
  - `src-tauri/src/main.rs`
  - `src-tauri/capabilities/default.json`
  - `src-tauri/icons/icon.png`
- Added native desktop file selection and save flow through the Tauri dialog plugin.
- Added a Tauri command that invokes the local Python generator without a browser API server.
- Added `--summary-json` CLI output for machine-readable desktop generation results.
- Added Tauri development launchers:
  - `launchers/Open Tauri Pella Order Automation.command`
  - `launchers/Open Tauri Pella Order Automation.bat`
- Added `scripts/tauri-dev.mjs` so `npm run desktop:dev` finds a Python runtime with `openpyxl` and passes it to the Tauri command.
- Added yellow workbook highlighting for generated `Jamb-BOM` input and output cells associated with warning-bearing lines, so review areas are visible inside the completed workbook.
- Added client packaging path:
  - `scripts/pella_generator_entry.py`
  - `scripts/build-generator.mjs`
  - `npm run sidecar:build`
  - `npm run desktop:package`
  - `npm run desktop:package:mac`
- Configured Tauri `bundle.externalBin` for `binaries/pella-generator`.
- Added `tauri-plugin-shell` and production sidecar execution so packaged clients do not need Python installed.
- Built and verified the macOS ARM sidecar:
  - `src-tauri/binaries/pella-generator-aarch64-apple-darwin`
- Built a macOS DMG for local delivery testing:
  - `src-tauri/target/release/bundle/dmg/Pella Order Automation_0.1.0_aarch64.dmg`
- Double-click launchers:
  - `launchers/Open Pella Order Automation.command`
  - `launchers/Open Pella Order Automation.bat`
  - `launchers/Open Modern Pella Order Automation.command`
  - `launchers/Open Modern Pella Order Automation.bat`
- Smoke test script for Examples 1 and 2.

Current generated files are ignored under `outputs/`.

## Verification So Far

Compile check:

- `python -m compileall -q src tests` passed.

Smoke test:

- Example 1 generated successfully.
- Example 1 produced 4 jamb lines and 0 warnings.
- Example 2 generated successfully.
- Example 2 produced 14 jamb lines and 8 warnings.
- Example 2 warnings are currently for arch/curved units that need review.

`pytest` was not available in the bundled runtime, so verification used compile checks and the smoke script.

Tauri desktop build note:

- `npm run build` passes for the React/Tailwind frontend.
- `npm run tauri -- --version` reports Tauri CLI 2.11.0.
- Rust/Cargo is now available in the local environment.
- The Python generator is now packaged as a PyInstaller sidecar for macOS ARM.
- `npm run tauri -- build --bundles app` completed successfully.
- `npm run tauri -- build --bundles dmg` completed successfully.
- Windows delivery still needs a Windows build machine or CI runner to generate the Windows sidecar and installer.

## Windows Packaged App

The step-by-step Windows packaging plan now lives in `windows_package.plan`.

## Current Limitations

- Workbook output is a first-pass `.xlsx`, not yet visually matched to the completed examples.
- The output writer currently generates static values rather than preserving the original workbook formulas/macros.
- Example 2 arch/curved logic is intentionally conservative and warning-heavy.
- The simple GUI is a prototype, not the final polished cross-platform desktop app.
- No packaged installer exists yet.
- The double-click launchers are for local development/testing; a packaged app should replace them for non-technical users.
- No cloud AI fallback exists yet.
- No user-editable correction workflow exists yet.

## Next Steps

1. Compare generated Example 1 workbook against Completed Sheet Example 1 cell-by-cell for key ranges.
2. Compare generated Example 2 workbook against Completed Sheet Example 2 and decide how arch/curved rows should be represented.
3. Tighten workbook formatting to more closely match completed examples.
4. Improve warning deduplication so repeated arch/curved warnings are easier to read.
5. Add parser regression tests that do not depend on `pytest`, or add a proper test environment.
6. Decide final output extension: `.xlsx` versus `.xlsm`.
7. Decide whether visual parity or downstream data correctness matters more.
8. Build a more polished desktop UI after the extraction rules stabilize.
9. Package the app for Windows and macOS once the core workflow is reliable, so users can launch it from a normal app icon/shortcut without Python setup.

## Remaining Questions

1. Will users always provide exactly one PO and one work order per job?
2. Can users reliably export work orders as XLSX?
3. Is `.xlsx` fully acceptable as the final output if the data and layout match what is needed?
4. Which sheets are actually used downstream: `Jamb-BOM`, `Lineal-BOM`, both, or only exported rows?
5. For arch and curved units, what should the final generated rows look like when dimensions are ambiguous?
6. Should future versions allow users to correct extracted fields before generation as an advanced option?
7. What level of warning frequency is acceptable for daily use?
8. If cloud AI is added later, should it be opt-in per file, automatic for low-confidence lines, or admin-controlled?
