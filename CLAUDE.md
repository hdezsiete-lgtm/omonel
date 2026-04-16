# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

**Omonel Dispersión** is a desktop GUI application (Tkinter) that processes two Excel files — a payroll file ("People") and an employee card accounts file ("Cuenta Vales") — and generates a formatted Excel layout for Omonel bank dispersal. The output is a styled `.xlsx` with employee data, CLABE numbers, vale account numbers, and amounts, plus a summary sheet.

## Running the app

```bash
pip3 install -r requirements.txt
python3 app.py
```

## Building the standalone executable

```bash
# Using the spec file (recommended — preserves build config)
pyinstaller OmonelDispersion.spec

# Or from scratch
pyinstaller --onefile --windowed --name OmonelDispersion app.py
```

The executable is written to `dist/OmonelDispersion` (macOS: `dist/OmonelDispersion.app`).

## Architecture

The entire application lives in `app.py` (~635 lines). There are no modules, packages, or tests.

**UI layer** — three reusable Tkinter widgets:
- `TagInput` — chip-style multi-value input for concept key codes
- `FilePickerCard` — labeled file selector card with status indicator
- `LogPanel` — scrollable timestamped log with color-coded severity levels

**Business logic** — two pure functions (no UI dependencies):
- `process_files(path_people, path_vales, conceptos, log_fn) → DataFrame | None` — reads both Excel files, normalizes column names to `snake_case`, filters by concept keys if provided, and inner-joins on `clave_empleado` to produce a normalized output DataFrame
- `generate_omonel_layout(result_df, output_path) → str` — writes the formatted Omonel bank layout to xlsx using openpyxl directly (not pandas to_excel), with two sheets: "Dispersión Omonel" and "Resumen"

**Main window** — `OmonelApp(tk.Tk)` wires the widgets together. The `_run()` method is the only place that connects UI state to the two business-logic functions.

## Key column expectations

`process_files` expects these column names (after normalization to lowercase snake_case) in the input files:

| File | Required | Auto-detected alternatives |
|------|----------|---------------------------|
| People | `clave_empleado` | — |
| People | `concepto` or `clave_concepto` | falls back to all rows if missing |
| People | `importe`, `monto`, or `cantidad` | defaults to 0 if none found |
| People | `rfc`, `nombre` / `nombre_empleado`, `clabe` / `clabe_interbancaria`, `banco` | empty string if missing |
| Cuenta Vales | `clave_empleado` | — |
| Cuenta Vales | `cuenta_vale` | falls back to last column if missing |
