"use strict";

const PYODIDE_BASE = "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/";
const KEY_PROJECT = "appointment_project";
const KEY_NAMES = "appointment_names";

const $ = (id) => document.getElementById(id);

const ui = {
  title: $("f-title"), month: $("f-month"), days: $("f-days"),
  c: $("f-c"), l: $("f-l"),
  name: $("f-name"), x: $("f-x"), dec: $("f-dec"),
  zile: $("f-zile"), loc: $("f-loc"),
  acList: $("ac-list"),
  formTitle: $("form-title"),
  btnAdd: $("btn-add"), btnUpdate: $("btn-update"),
  btnCancel: $("btn-cancel"), btnRemove: $("btn-remove"),
  tableBody: document.querySelector("#table tbody"),
  empty: $("empty"), badge: $("badge"),
  btnXlsx: $("btn-xlsx"), btnJsonOut: $("btn-json-out"),
  jsonIn: $("f-json-in"), btnClear: $("btn-clear"),
  status: $("status"), pill: $("pill"), toast: $("toast"),
};

const state = {
  ready: false,
  editIndex: null,
  knownNames: [],
  bridge: null,
  project: {
    title: "", month: "", days_in_month: 30,
    c_days: [], l_days: [], patients: [],
  },
};

let toastTimer = null;
let saveTimer = null;
let acItems = [];
let acIndex = -1;

function toast(message, kind) {
  kind = kind || "error";
  ui.toast.textContent = message;
  ui.toast.className = "toast " + kind;
  ui.toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { ui.toast.hidden = true; }, 3400);
}

function setStatus(text, pill) {
  ui.status.textContent = text;
  ui.pill.textContent = pill;
  ui.pill.className = "pill " + pill;
}

function daysValue() {
  const n = parseInt(ui.days.value, 10);
  return Number.isInteger(n) && n >= 1 && n <= 31 ? n : null;
}

function loadStorage() {
  try {
    const names = JSON.parse(localStorage.getItem(KEY_NAMES) || "[]");
    if (Array.isArray(names)) {
      state.knownNames = names.filter((n) => typeof n === "string" && n.trim());
    }
  } catch (e) { /* fresh start */ }
  try {
    const raw = localStorage.getItem(KEY_PROJECT);
    if (!raw) return;
    const p = JSON.parse(raw);
    if (!p || typeof p !== "object") return;
    state.project = {
      title: typeof p.title === "string" ? p.title : "",
      month: typeof p.month === "string" ? p.month : "",
      days_in_month: Number.isInteger(p.days_in_month) ? p.days_in_month : 30,
      c_days: Array.isArray(p.c_days) ? p.c_days : [],
      l_days: Array.isArray(p.l_days) ? p.l_days : [],
      patients: Array.isArray(p.patients) ? p.patients : [],
    };
  } catch (e) { /* fresh start */ }
}

function parseDaysSafe(text, fallback) {
  if (!state.bridge) return fallback;
  const r = JSON.parse(state.bridge.parse(text, daysValue()));
  return r.ok ? r.days : fallback;
}

function collectRaw() {
  const n = daysValue();
  if (n === null) return null;
  return {
    title: ui.title.value.trim(),
    month: ui.month.value.trim(),
    days_in_month: n,
    c_days: parseDaysSafe(ui.c.value, state.project.c_days),
    l_days: parseDaysSafe(ui.l.value, state.project.l_days),
    patients: state.project.patients,
  };
}

function persist() {
  const p = collectRaw();
  if (!p) return;
  state.project = p;
  try { localStorage.setItem(KEY_PROJECT, JSON.stringify(p)); } catch (e) { /* full */ }
}

function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(persist, 400);
}

function saveNames() {
  try { localStorage.setItem(KEY_NAMES, JSON.stringify(state.knownNames)); } catch (e) { /* full */ }
}

function fillHeaderForm() {
  ui.title.value = state.project.title;
  ui.month.value = state.project.month;
  ui.days.value = String(state.project.days_in_month);
  ui.c.value = state.project.c_days.join(", ");
  ui.l.value = state.project.l_days.join(", ");
}

function renderTable() {
  const patients = state.project.patients;
  ui.tableBody.innerHTML = "";
  ui.badge.textContent = String(patients.length);
  ui.empty.hidden = patients.length > 0;
  patients.forEach((p, i) => {
    const tr = document.createElement("tr");
    if (i === state.editIndex) tr.classList.add("selected");
    const days = Array.isArray(p.x_days) ? p.x_days : [];
    const cells = [
      [String(i + 1), "num"],
      [p.name || "", "name"],
      [days.join(", "), "days"],
      [String(days.length), "count"],
      [p.decision_date || "", "meta"],
      [p.nr_zile || "", "meta"],
      [p.localitate || "", "loc"],
    ];
    cells.forEach(([text, cls]) => {
      const td = document.createElement("td");
      td.textContent = text;
      td.className = cls;
      tr.appendChild(td);
    });
    tr.addEventListener("click", () => startEdit(i));
    ui.tableBody.appendChild(tr);
  });
}

function hideAc() {
  ui.acList.hidden = true;
  ui.acList.innerHTML = "";
  acItems = [];
  acIndex = -1;
}

function showAc(items) {
  ui.acList.innerHTML = "";
  acItems = items;
  acIndex = items.length ? 0 : -1;
  if (!items.length) { hideAc(); return; }
  items.forEach((name, i) => {
    const div = document.createElement("div");
    div.className = "ac-item" + (i === acIndex ? " active" : "");
    div.textContent = name;
    div.addEventListener("mousedown", (e) => { e.preventDefault(); pickAc(i); });
    ui.acList.appendChild(div);
  });
  ui.acList.hidden = false;
}

function acFilter() {
  if (!state.ready) return;
  const items = JSON.parse(
    state.bridge.filter(JSON.stringify(state.knownNames), ui.name.value));
  showAc(items);
}

function pickAc(i) {
  if (i < 0 || i >= acItems.length) return;
  ui.name.value = acItems[i];
  hideAc();
  ui.x.focus();
}

function moveAc(step) {
  if (!acItems.length) return;
  acIndex = (acIndex + step + acItems.length) % acItems.length;
  Array.from(ui.acList.children).forEach((el, i) =>
    el.classList.toggle("active", i === acIndex));
  ui.acList.children[acIndex].scrollIntoView({ block: "nearest" });
}

function clearForm() {
  state.editIndex = null;
  [ui.name, ui.x, ui.dec, ui.zile, ui.loc].forEach((el) => { el.value = ""; });
  ui.formTitle.textContent = "Add patient";
  ui.btnAdd.hidden = false;
  ui.btnUpdate.hidden = true;
  ui.btnCancel.hidden = true;
  ui.btnRemove.hidden = true;
  hideAc();
  renderTable();
}

function startEdit(i) {
  const p = state.project.patients[i];
  state.editIndex = i;
  ui.name.value = p.name || "";
  ui.x.value = (Array.isArray(p.x_days) ? p.x_days : []).join(", ");
  ui.dec.value = p.decision_date || "";
  ui.zile.value = p.nr_zile || "";
  ui.loc.value = p.localitate || "";
  ui.formTitle.textContent = "Edit patient";
  ui.btnAdd.hidden = true;
  ui.btnUpdate.hidden = false;
  ui.btnCancel.hidden = false;
  ui.btnRemove.hidden = false;
  hideAc();
  renderTable();
  ui.formTitle.scrollIntoView({ behavior: "smooth", block: "center" });
}

function readForm() {
  const name = ui.name.value.trim();
  if (!name) { toast("Patient name is required."); ui.name.focus(); return null; }
  const r = JSON.parse(state.bridge.parse(ui.x.value, daysValue()));
  if (!r.ok) { toast("X days: " + r.error); ui.x.focus(); return null; }
  return {
    name: name,
    x_days: r.days,
    decision_date: ui.dec.value.trim(),
    nr_zile: ui.zile.value.trim(),
    localitate: ui.loc.value.trim(),
  };
}

function registerName(name) {
  const lower = name.toLowerCase();
  if (!state.knownNames.some((n) => n.toLowerCase() === lower)) {
    state.knownNames.push(name);
    saveNames();
  }
}

function addPatient() {
  const p = readForm();
  if (!p) return;
  state.project.patients.push(p);
  registerName(p.name);
  persist();
  clearForm();
  toast("Patient added.", "success", 1500);
}

function updatePatient() {
  if (state.editIndex === null) {
    toast("Select a patient in the table first.");
    return;
  }
  const p = readForm();
  if (!p) return;
  state.project.patients[state.editIndex] = p;
  registerName(p.name);
  persist();
  clearForm();
  toast("Patient updated.", "success", 1500);
}

function removePatient() {
  if (state.editIndex === null) {
    toast("Select a patient in the table first.");
    return;
  }
  const p = state.project.patients[state.editIndex];
  if (!window.confirm("Remove " + (p.name || "this patient") + "?")) return;
  state.project.patients.splice(state.editIndex, 1);
  persist();
  clearForm();
  toast("Patient removed.", "success", 1500);
}

function download(bytes, filename, mime) {
  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 6000);
}

function b64ToBytes(b64) {
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

function exportXlsx() {
  const raw = collectRaw();
  if (!raw) { toast("Days must be a whole number between 1 and 31."); return; }
  const v = JSON.parse(state.bridge.validate(JSON.stringify(raw)));
  if (!v.ok) { toast(v.error); return; }
  const r = JSON.parse(state.bridge.export(JSON.stringify(v.project)));
  if (!r.ok) { toast(r.error); return; }
  download(b64ToBytes(r.b64), r.filename, r.mime);
  toast("Excel file exported.", "success", 2200);
}

function exportJson() {
  const p = collectRaw();
  if (!p) { toast("Days must be a whole number between 1 and 31."); return; }
  const base = (p.title || "appointments").replace(/[^\w\- ]+/g, "").trim()
    .replace(/\s+/g, "_") || "appointments";
  download(new TextEncoder().encode(JSON.stringify(p, null, 2)),
    base + ".json", "application/json");
  toast("Project file exported.", "success", 2200);
}

async function importJson(file) {
  if (!file) return;
  try {
    const text = await file.text();
    const v = JSON.parse(state.bridge.validate(text));
    if (!v.ok) { toast("Import failed: " + v.error); return; }
    state.project = v.project;
    state.project.patients.forEach((p) => {
      if (p.name) registerName(p.name);
    });
    fillHeaderForm();
    clearForm();
    persist();
    toast("Project imported.", "success", 2200);
  } catch (e) {
    toast("Import failed: " + e.message);
  }
}

function clearAll() {
  if (state.project.patients.length &&
      !window.confirm("Clear header, pattern and all patients?")) return;
  state.project = {
    title: "", month: "", days_in_month: 30,
    c_days: [], l_days: [], patients: [],
  };
  fillHeaderForm();
  clearForm();
  persist();
}

function validateField(input, label) {
  if (!state.ready || !input.value.trim()) return;
  const r = JSON.parse(state.bridge.parse(input.value, daysValue()));
  if (!r.ok) toast(label + ": " + r.error);
  else { persist(); }
}

function setEnabled(on) {
  ["f-title", "f-month", "f-days", "f-c", "f-l", "f-name", "f-x",
   "f-dec", "f-zile", "f-loc", "btn-add", "btn-xlsx", "btn-json-out",
   "btn-clear"].forEach((id) => { $(id).disabled = !on; });
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error("Cannot download " + src));
    document.head.appendChild(s);
  });
}

async function boot() {
  loadStorage();
  fillHeaderForm();
  renderTable();
  setStatus("Loading Python runtime…", "loading");
  try {
    await loadScript(PYODIDE_BASE + "pyodide.js");
    setStatus("Starting Python…", "loading");
    const pyodide = await loadPyodide({ indexURL: PYODIDE_BASE });
    setStatus("Loading Excel engine (openpyxl)…", "loading");
    await pyodide.loadPackage("micropip");
    await pyodide.runPythonAsync(
      'import micropip\nawait micropip.install("openpyxl")');
    setStatus("Loading app logic…", "loading");
    const backendCode = await (await fetch("backend.py")).text();
    const bridgeCode = await (await fetch("web_bridge.py")).text();
    pyodide.runPython(`
import sys, types

def _load_module(name, code):
    mod = types.ModuleType(name)
    mod.__file__ = name + ".py"
    exec(compile(code, name + ".py", "exec"), mod.__dict__)
    sys.modules[name] = mod
    return mod

def _init(backend_code, bridge_code):
    _load_module("backend", backend_code)
    wb = _load_module("web_bridge", bridge_code)
    g = globals()
    for fn in ("parse_days_json", "filter_names_json",
               "validate_project_json", "export_xlsx_b64"):
        g[fn] = getattr(wb, fn)
`);
    const init = pyodide.globals.get("_init");
    init(backendCode, bridgeCode);
    init.destroy();
    state.bridge = {
      parse: pyodide.globals.get("parse_days_json"),
      filter: pyodide.globals.get("filter_names_json"),
      validate: pyodide.globals.get("validate_project_json"),
      export: pyodide.globals.get("export_xlsx_b64"),
    };
    state.ready = true;
    setEnabled(true);
    setStatus(navigator.onLine ? "Ready" : "Ready — offline", "ready");
  } catch (err) {
    console.error(err);
    setStatus("Load failed: " + err.message + " — tap to retry", "error");
    ui.status.style.cursor = "pointer";
    ui.status.onclick = () => location.reload();
  }
}

function wireEvents() {
  ui.btnAdd.addEventListener("click", addPatient);
  ui.btnUpdate.addEventListener("click", updatePatient);
  ui.btnCancel.addEventListener("click", clearForm);
  ui.btnRemove.addEventListener("click", removePatient);
  ui.btnXlsx.addEventListener("click", exportXlsx);
  ui.btnJsonOut.addEventListener("click", exportJson);
  ui.btnClear.addEventListener("click", clearAll);
  ui.jsonIn.addEventListener("change", () => {
    const f = ui.jsonIn.files[0];
    ui.jsonIn.value = "";
    if (!state.ready) { toast("Still loading — try again in a moment."); return; }
    importJson(f);
  });

  [ui.title, ui.month, ui.days, ui.c, ui.l].forEach((el) =>
    el.addEventListener("input", scheduleSave));
  ui.c.addEventListener("blur", () => validateField(ui.c, "C days"));
  ui.l.addEventListener("blur", () => validateField(ui.l, "L days"));
  ui.days.addEventListener("blur", () => {
    if (daysValue() === null) toast("Days must be a whole number between 1 and 31.");
    else persist();
  });

  ui.name.addEventListener("focus", acFilter);
  ui.name.addEventListener("input", acFilter);
  ui.name.addEventListener("blur", () => setTimeout(hideAc, 160));
  ui.name.addEventListener("keydown", (e) => {
    if (!ui.acList.hidden) {
      if (e.key === "ArrowDown") { e.preventDefault(); moveAc(1); return; }
      if (e.key === "ArrowUp") { e.preventDefault(); moveAc(-1); return; }
      if (e.key === "Enter") { e.preventDefault(); pickAc(acIndex); return; }
      if (e.key === "Escape") { hideAc(); return; }
    }
    if (e.key === "Enter") {
      e.preventDefault();
      if (state.editIndex === null) addPatient(); else updatePatient();
    }
  });

  [ui.x, ui.dec, ui.zile, ui.loc].forEach((el) =>
    el.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      if (state.editIndex === null) addPatient(); else updatePatient();
    }));

  window.addEventListener("online", () => {
    if (state.ready) setStatus("Ready", "ready");
  });
  window.addEventListener("offline", () => {
    if (state.ready) setStatus("Ready — offline", "ready");
  });

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js")
      .catch((e) => console.warn("Service worker failed:", e));
  }
}

wireEvents();
setEnabled(false);
boot();
