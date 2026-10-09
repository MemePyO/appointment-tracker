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

- **Add / edit / remove patients** with appointment days (`1, 5, 12-14`)
- **Autocomplete** patient names — remembered on your device as you type
- **Pattern for all patients** — the same `C` and `L` free days applied to everyone
- **Counts** (`X`) per patient, shown live in the table
- **Export real `.xlsx`** — title row, per-day columns, `COUNTIF` formulas, `TOTAL` row, color-coded cells (compatible with the desktop version's layout)
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
| `web_bridge.py` | Thin JSON ↔ Python bridge for the browser |
| `app.js` / `index.html` / `style.css` | UI, autocomplete, autosave |
| `sw.js` / `manifest.webmanifest` | Offline support + installable app |
| `icons/` | App icons |

The browser loads the Python files as text, executes them in Pyodide (CPython compiled to WebAssembly), installs `openpyxl` via `micropip`, and calls the same functions the desktop app uses.
