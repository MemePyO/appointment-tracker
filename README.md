# Appointment Tracker

A phone-friendly web app (PWA) for tracking patient appointments, built on the **same Python backend as the desktop app** — it runs unmodified in the browser via [Pyodide](https://pyodide.org) together with `openpyxl`, so the generated Excel files are identical on every platform.

**Live app:** https://memepyo.github.io/appointment-tracker/

## Legend

| Cell | Meaning |
|------|---------|
| `X` | Appointment |
| `L` | Legal free day |
| `C` | Free day |
| *(empty)* | Not there |

`X` always wins over `L`/`C`.

## Features

- **Add / edit / remove patients** with appointment days (`Zile sedinte`: `1, 5, 12-14`)
- **Autocomplete** patient names — remembered on your device as you type
- **Pattern for all patients** — the same `C` and `L` free days applied to everyone
- **Counts** (`X`) per patient, shown live in the table
- **Multi-sheet workbook** — sheets named `MM.YYYY` (needs the Year field, default = current year); "Update existing" asks before replacing a same-named sheet, "Create new" starts fresh
- **Introdu ședință** — mark one appointment by date (e.g. STOICA FLORIN, 30.04.2026): `X` on that day in the matching sheet, patient/sheet auto-created, nothing changes if already marked; the updated workbook is downloaded
- **Load workbook (.xlsx)** — pull a workbook from your device into the browser so "Update existing" and Introdu ședință can edit it
- **Export real `.xlsx`** — title row, per-day columns, `COUNTIF` formulas, `TOTAL` row, color-coded cells (identical layout to the desktop app)
- **Export / import project `.json`** — move work between phone and PC, or between devices
- **Autosave** — everything is stored in your browser (`localStorage`) and survives reloads / offline
- **Offline** — works without internet after the first visit (service worker caches the app and the Python runtime)
- **Installable** — add it to your home screen and use it like a native app

## Install on your phone

1. Open https://memepyo.github.io/appointment-tracker/ in Chrome (Android) / Safari (iOS)
2. Chrome: **⋮ menu → Add to Home screen → Install**
   Safari (iOS): **Share → Add to Home Screen**

## Phone ↔ PC

Use **Export project (.json)** on the phone and open the file with the
[desktop app](https://github.com/MemePyO/appointment-tracker-desktop) (or vice versa) — both use the same JSON format.

## Privacy

- No server, no account, no analytics
- Patient data never leaves your device — it is stored only in the browser's `localStorage`
- The exported `.xlsx` / `.json` files stay wherever you save them

## How it works

| File | Role |
|------|------|
| `backend.py` | Pure-Python core (parsing, validation, Excel export) — shared with the desktop app |
| `workbook.py` | Multi-sheet `MM.YYYY` export — shared with the desktop app |
| `appointment.py` | Introdu ședință logic (mark X by date) — shared with the desktop app |
| `web_bridge.py` | Thin JSON ↔ Python bridge for the browser |
| `app.js` / `index.html` / `style.css` | UI, autocomplete, autosave, export modal |
| `sw.js` / `manifest.webmanifest` | Offline support + installable app |
| `icons/` | App icons |

The browser loads the Python files as text, executes them in Pyodide (CPython compiled to WebAssembly), installs `openpyxl` via `micropip`, and calls the same functions the desktop app uses.
