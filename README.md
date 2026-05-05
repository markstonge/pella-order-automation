# Pella Order Automation

Offline-first prototype for generating completed Pella order workbooks from:

- CSV purchase orders
- XLSX work orders

Version 1 intentionally defers legacy `.xls` support and macro preservation.

## Prototype CLI

```bash
python -m pella_order_automation.cli generate \
  --po "/path/to/Purchase Order.csv" \
  --work-order "/path/to/Work Order.xlsx" \
  --output "/path/to/output.xlsx"
```

The CLI prints detected warnings and generates a workbook with `Jamb-BOM` and `Lineal-BOM` sheets.

## Client Folder Script

For a simple no-GUI workflow, use `create_output.py`.

Place the script in the same folder as these two files:

- `purchase_order.csv`
- `work_order.xlsx`

Then double-click the script or run:

```bash
python create_output.py
```

It creates `output.xlsx` in that same folder. The script prints simple terminal errors if an input file is missing, the work order cannot be read, or Python is missing the required `openpyxl` package.

## Prototype Desktop App

```bash
python -m pella_order_automation.gui
```

This opens a simple file-picker UI for selecting the CSV purchase order, XLSX work order, and output workbook path.

The GUI displays:

- polished file selector panels for the purchase order, work order, and output file
- status and progress indicator
- generated order summary
- metric cards for jamb lines, generated BOM rows, and warnings
- order number, PO number, customer, request dates
- generated line/row counts
- tabbed warnings and errors panels
- unsupported/low-confidence messages
- diagnostic details if generation fails

## Tauri Desktop App

The production direction is a Tauri desktop application. The UI is built with React, shadcn-style components, Radix primitives, Tailwind CSS, and lucide icons. Users will open a normal desktop app, not a browser tab or local web server.

Current desktop behavior:

- native file picker for the purchase order CSV
- native file picker for the work order XLSX
- native save dialog for the completed workbook
- in-app status, warnings, errors, and summary
- local generation through a Tauri command

Development prerequisites:

- Node.js and npm
- Rust/Cargo for running or building Tauri
- Python for the current generator bridge

Frontend build:

```bash
npm run build
```

Desktop development:

```bash
npm run desktop:dev
```

This command automatically looks for a Python runtime that has `openpyxl` installed and passes it to Tauri as `PELLA_PYTHON`.

Development launchers:

- macOS: double-click `launchers/Open Tauri Pella Order Automation.command`
- Windows: double-click `launchers/Open Tauri Pella Order Automation.bat`

The current Tauri command invokes the Python generator module directly. For final distribution, the Python generator should be bundled as a sidecar executable so end users do not need Python installed.

Final packaged app:

```bash
npm run desktop:package
```

macOS DMG package:

```bash
npm run desktop:package:mac
```

Windows portable package:

```bash
npm run desktop:package:windows
```

The packaging command first builds the Python generator into a standalone Tauri sidecar binary under `src-tauri/binaries/`, then runs the Tauri bundle build. After packaging, users should open the generated `.app`, `.dmg`, `.exe`, or installer directly. No CLI command, Python install, Rust install, npm install, browser, or local server is needed for normal use.

Current macOS output location:

```text
src-tauri/target/release/bundle/dmg/Pella Order Automation_0.1.0_aarch64.dmg
```

Windows package through GitHub Actions:

1. Push the project to GitHub.
2. Open the repository's Actions tab.
3. Select `Build Windows Portable Package`.
4. Click `Run workflow`.
5. Download the `Pella Order Automation Windows Portable` artifact.

The artifact contains `Pella Order Automation Windows Portable.zip`, which users can unzip and run by opening `Pella Order Automation.exe`.

Windows packages are built on a hosted GitHub Actions Windows runner. A local Windows machine is still useful for final Windows 10 and Windows 11 launch testing.

## Legacy Launchers

The earlier Tkinter and local browser launchers remain in `launchers/` for reference while the Tauri app is being built out.
