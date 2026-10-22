/* global Office, Excel */
// Task pane UI logic. Deliberately thin: every real decision (URL parsing,
// query building, formula text, error messages) lives in src/lib/*.js and is
// unit-tested there. This file's only job is DOM wiring and calling Office.js.

import { getClient, getSettings, saveSettings, loadSettings } from "../runtime/session.js";
import { defaultSettings } from "../lib/settings.js";
import { toCellMessage } from "../lib/errors.js";
import { pickerValueToIso, isoToPickerValue, addTags, removeTag } from "../lib/picker.js";
import { buildGuide, guideConflicts } from "./guide.js";
import {
  buildHistoryFormula,
  buildSampleFormula,
  buildDatasetsFormula,
  looksLikeReference,
} from "../lib/formula.js";

const $ = (id) => document.getElementById(id);

// --- tabs --------------------------------------------------------------------

function initTabs() {
  const tabs = document.querySelectorAll(".tb-tab");
  const panels = {
    connect: $("tab-connect"),
    insert: $("tab-insert"),
  };
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((t) => {
        t.classList.toggle("is-active", t === tab);
        t.setAttribute("aria-selected", String(t === tab));
      });
      Object.entries(panels).forEach(([name, el]) => {
        el.hidden = name !== tab.dataset.tab;
      });
      if (tab.dataset.tab === "insert") {
        populateDatasetDropdowns().catch(reportBackground);
      }
    });
  });
}

// --- status helpers ------------------------------------------------------------

function setStatus(el, message, kind) {
  el.textContent = message ?? "";
  el.classList.remove("is-error", "is-ok");
  if (kind) el.classList.add(kind === "error" ? "is-error" : "is-ok");
}

/** Errors from a background populate (e.g. dataset dropdown refresh) that the
 * user did not directly trigger: log rather than interrupt them mid-task. */
function reportBackground(err) {
  // eslint-disable-next-line no-console
  console.warn(toCellMessage(err));
}

// --- Connect tab ---------------------------------------------------------------

function loadSettingsIntoForm() {
  const s = getSettings();
  // Neither the Historian nor Pulse sends CORS headers, so the working setup is
  // the one where this page's own origin serves (or proxies) both. Offer that
  // as the starting point for a workbook that has never been connected.
  const fresh = !s.historianUrl && !s.clientId;
  $("historianUrl").value = fresh ? window.location.origin : s.historianUrl;
  $("authEnabled").checked = s.authEnabled;
  $("pulseUrl").value = fresh ? window.location.origin : s.pulseUrl;
  $("clientId").value = s.clientId;
  $("clientSecret").value = s.clientSecret;
  $("maxRows").value = s.maxRows;
  $("authFields").style.display = s.authEnabled ? "" : "none";
}

function readFormSettings() {
  return {
    ...defaultSettings(),
    historianUrl: $("historianUrl").value,
    authEnabled: $("authEnabled").checked,
    pulseUrl: $("pulseUrl").value,
    clientId: $("clientId").value,
    clientSecret: $("clientSecret").value,
    maxRows: $("maxRows").value,
  };
}

function initConnectTab() {
  loadSettingsIntoForm();

  $("authEnabled").addEventListener("change", () => {
    $("authFields").style.display = $("authEnabled").checked ? "" : "none";
  });

  $("testConnection").addEventListener("click", async () => {
    const statusEl = $("connectStatus");
    setStatus(statusEl, "Connecting…", null);
    try {
      // Save first: getClient() reads from the persisted store, so a test
      // must reflect what "Save" would actually apply.
      const saved = await saveSettings(readFormSettings());
      const client = getClient();
      const datasets = await client.listDatasets();
      setStatus(
        statusEl,
        `Connected. Found ${datasets.length} dataset${datasets.length === 1 ? "" : "s"}` +
          (saved.authEnabled ? " (token obtained)." : " (authentication is off)."),
        "ok",
      );
    } catch (err) {
      setStatus(statusEl, toCellMessage(err), "error");
    }
  });

  $("saveSettings").addEventListener("click", async () => {
    const statusEl = $("connectStatus");
    try {
      await saveSettings(readFormSettings());
      setStatus(statusEl, "Saved.", "ok");
      populateDatasetDropdowns().catch(reportBackground);
    } catch (err) {
      setStatus(statusEl, toCellMessage(err), "error");
    }
  });

  $("securityLink").addEventListener("click", (e) => {
    e.preventDefault();
    window.open(
      "https://github.com/jbknepp98/timebase-historian-excel-add-in/blob/main/docs/security.md",
      "_blank",
    );
  });
}

// --- server-held authentication -------------------------------------------------

/**
 * Ask the server this page came from whether it signs in to Pulse itself
 * (tools/serve.js with a credentials file). If so, Excel must not hold
 * credentials at all: authentication is switched off here, the fields are
 * hidden, and any client secret saved earlier is wiped from storage.
 */
async function applyServerAuth() {
  let info = { enabled: false };
  try {
    const res = await fetch("/__server-auth", { credentials: "omit", cache: "no-store" });
    if (res.ok) info = await res.json();
  } catch {
    /* not served by tools/serve.js: keep client-side authentication */
  }
  if (!info.enabled) return;

  $("serverAuthNote").hidden = false;
  $("serverAuthClient").textContent = info.clientId || "";
  $("authToggle").hidden = true;
  $("authFields").style.display = "none";
  $("authEnabled").checked = false;

  const s = getSettings();
  if (s.authEnabled || s.clientSecret) {
    try {
      await saveSettings({ ...s, historianUrl: s.historianUrl || window.location.origin, authEnabled: false });
    } catch (err) {
      reportBackground(err);
    }
  }
}

// --- no-value rows preference --------------------------------------------------

/** One saved preference; new formulas follow it. */
function noValueMode() {
  return getSettings().noValueRows === "drop" ? "drop" : "flag";
}

function initNoValueOption() {
  const boxes = document.querySelectorAll(".tb-drop-novalue");
  const sync = () => boxes.forEach((b) => (b.checked = noValueMode() === "drop"));
  sync();
  boxes.forEach((box) =>
    box.addEventListener("change", async () => {
      const mode = box.checked ? "drop" : "flag";
      boxes.forEach((b) => (b.checked = box.checked));
      try {
        await saveSettings({ ...getSettings(), noValueRows: mode });
      } catch (err) {
        // Not connected yet: the choice still applies to formulas built now.
        getSettings().noValueRows = mode;
        reportBackground(err);
      }
      updateFormulaPreview();
    }),
  );
}

// --- extra columns preference ---------------------------------------------------

function extraColumns() {
  const cols = getSettings().extraColumns;
  return Array.isArray(cols) ? cols : [];
}

/** The four Extra columns checkboxes, saved as one preference. */
function initExtraColumns() {
  const boxes = document.querySelectorAll(".tb-extra-col");
  const sync = () => boxes.forEach((b) => (b.checked = extraColumns().includes(b.dataset.col)));
  sync();
  boxes.forEach((box) =>
    box.addEventListener("change", async () => {
      const chosen = new Set(extraColumns());
      if (box.checked) chosen.add(box.dataset.col);
      else chosen.delete(box.dataset.col);
      const ordered = ["quality", "status", "samples", "good%"].filter((k) => chosen.has(k));
      try {
        await saveSettings({ ...getSettings(), extraColumns: ordered });
      } catch (err) {
        getSettings().extraColumns = ordered;
        reportBackground(err);
      }
      sync();
      updateFormulaPreview();
    }),
  );
}

// --- getting-started guide ------------------------------------------------------

/**
 * Add the guide's sheets. If they already exist, the first click says so and a
 * second click within ten seconds replaces them -- never silently overwrite
 * sheets someone may have edited.
 */
function initGuideButton() {
  let confirmUntil = 0;
  $("addGuide").addEventListener("click", async () => {
    const statusEl = $("guideStatus");
    try {
      const conflicts = await guideConflicts();
      const replace = conflicts.length > 0 && Date.now() < confirmUntil;
      if (conflicts.length && !replace) {
        confirmUntil = Date.now() + 10000;
        setStatus(statusEl, `This workbook already has guide sheets (${conflicts.join(", ")}). Click again within 10 seconds to replace them.`, "error");
        return;
      }
      setStatus(statusEl, "Adding the guide\u2026", null);
      $("addGuide").disabled = true;
      await buildGuide({ replace });
      confirmUntil = 0;
      setStatus(statusEl, "Guide added. Start on the Welcome sheet; put your own dataset and tags on Setup.", "ok");
      console.info("[guide] added");
    } catch (err) {
      setStatus(statusEl, toCellMessage(err), "error");
      console.error("[guide] failed", err && err.debugInfo ? JSON.stringify(err.debugInfo) : "", err);
    } finally {
      $("addGuide").disabled = false;
    }
  });
}

// --- shared dataset dropdown population -----------------------------------------

let datasetsCache = null;
let datasetsCacheAt = 0;
const DATASET_CACHE_MS = 30000;

async function listDatasetsCached() {
  const now = Date.now();
  if (datasetsCache && now - datasetsCacheAt < DATASET_CACHE_MS) return datasetsCache;
  const client = getClient();
  datasetsCache = await client.listDatasets();
  datasetsCacheAt = now;
  return datasetsCache;
}

async function populateDatasetDropdowns() {
  if (!getSettings().historianUrl) return;
  const datasets = await listDatasetsCached();
  for (const selectId of ["insertDataset"]) {
    const select = $(selectId);
    const previous = select.value;
    select.innerHTML = "";
    for (const d of datasets) {
      const opt = document.createElement("option");
      opt.value = d.name;
      opt.textContent = `${d.name} (${d.tagCount ?? 0} tags)`;
      select.appendChild(opt);
    }
    if (datasets.some((d) => d.name === previous)) select.value = previous;
  }
}

// --- Insert tab --------------------------------------------------------------------

// The Insert tab's tag list: chosen through the tag search, shown as chips.
let insertTagList = [];
const MAX_RESULTS_SHOWN = 200;

function renderTagChips() {
  const box = $("insertTagChips");
  box.innerHTML = "";
  for (const name of insertTagList) {
    const chip = document.createElement("span");
    chip.className = "tb-chip";
    const label = document.createElement("span");
    label.className = "tb-chip-name";
    label.textContent = name;
    label.title = name;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "\u00d7";
    remove.title = `Remove ${name}`;
    remove.setAttribute("aria-label", `Remove ${name}`);
    remove.addEventListener("click", () => {
      insertTagList = removeTag(insertTagList, name);
      renderTagChips();
      refreshPickerAddedState();
      updateFormulaPreview();
    });
    chip.append(label, remove);
    box.appendChild(chip);
  }
}

function initTagPicker() {
  const picker = $("tagPicker");
  $("openTagPicker").addEventListener("click", () => {
    picker.hidden = !picker.hidden;
    if (!picker.hidden) {
      $("tagPickerQuery").focus();
      if (!$("tagPickerResults").children.length) runTagSearch().catch(reportPickerError);
    }
  });
  $("tagPickerClose").addEventListener("click", () => {
    picker.hidden = true;
  });
  $("tagPickerSearch").addEventListener("click", () => runTagSearch().catch(reportPickerError));
  $("tagPickerQuery").addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      runTagSearch().catch(reportPickerError);
    }
  });
  $("tagPickerAdd").addEventListener("click", () => {
    const chosen = Array.from($("tagPickerResults").querySelectorAll("input[type=checkbox]:checked:not(:disabled)")).map(
      (cb) => cb.value,
    );
    insertTagList = addTags(insertTagList, chosen);
    renderTagChips();
    refreshPickerAddedState();
    updateFormulaPreview();
    setStatus($("tagPickerStatus"), `Added ${chosen.length} tag${chosen.length === 1 ? "" : "s"}.`, "ok");
  });
  // A new dataset means a new tag universe: clear stale results.
  $("insertDataset").addEventListener("change", () => {
    $("tagPickerResults").innerHTML = "";
    setStatus($("tagPickerStatus"), "", null);
  });
}

function reportPickerError(err) {
  setStatus($("tagPickerStatus"), toCellMessage(err), "error");
}

/**
 * Search the selected dataset. Plain text is a substring match (the API's
 * `contains`); anything using the Historian's filter language -- a field
 * prefix like name:, a wildcard, AND/OR/NOT -- goes to `query` unchanged.
 */
async function runTagSearch() {
  const dataset = $("insertDataset").value;
  const statusEl = $("tagPickerStatus");
  if (!dataset) {
    setStatus(statusEl, "Choose a dataset first.", "error");
    return;
  }
  const text = $("tagPickerQuery").value.trim();
  const usesFilterLanguage = /[*?:]|\b(AND|OR|NOT)\b/.test(text);
  setStatus(statusEl, "Searching\u2026", null);
  const tags = await getClient().listTags(dataset, usesFilterLanguage ? { query: text } : { contains: text || null });

  const list = $("tagPickerResults");
  list.innerHTML = "";
  for (const tag of tags.slice(0, MAX_RESULTS_SHOWN)) {
    const li = document.createElement("li");
    const label = document.createElement("label");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = tag.name;
    cb.addEventListener("change", updatePickerAddButton);
    const text = document.createElement("span");
    const name = document.createElement("span");
    name.className = "tb-result-name";
    name.textContent = tag.name;
    const desc = document.createElement("span");
    desc.className = "tb-result-desc";
    desc.textContent = [tag.description, tag.dataType ? tag.dataType.replace(/^System\./, "") : ""].filter(Boolean).join(" \u00b7 ");
    text.append(name, desc);
    label.append(cb, text);
    li.appendChild(label);
    list.appendChild(li);
  }
  refreshPickerAddedState();

  const shown = Math.min(tags.length, MAX_RESULTS_SHOWN);
  setStatus(
    statusEl,
    tags.length === 0
      ? "No tags matched."
      : tags.length > shown
        ? `Showing the first ${shown} of ${tags.length}. Narrow the search to see the rest.`
        : `${tags.length} tag${tags.length === 1 ? "" : "s"}.`,
    tags.length ? "ok" : null,
  );
}

/** Tags already in the list are shown ticked and locked, so they cannot be added twice. */
function refreshPickerAddedState() {
  for (const cb of $("tagPickerResults").querySelectorAll("input[type=checkbox]")) {
    const added = insertTagList.includes(cb.value);
    cb.disabled = added;
    if (added) cb.checked = true;
    else if (cb.dataset.wasAdded) cb.checked = false;
    cb.dataset.wasAdded = added ? "1" : "";
    cb.closest("li").classList.toggle("is-added", added);
  }
  updatePickerAddButton();
}

function updatePickerAddButton() {
  $("tagPickerAdd").disabled = !$("tagPickerResults").querySelector("input[type=checkbox]:checked:not(:disabled)");
}

/** Calendar button + datetime-local for one of the Start/End text boxes. */
function initDateTimePicker(which) {
  const text = $(`insert${which}`);
  const toggle = $(`insert${which}Pick`);
  const panel = $(`insert${which}Picker`);
  const input = $(`insert${which}DateTime`);

  const apply = () => {
    text.value = pickerValueToIso(input.value);
    updateFormulaPreview();
  };
  toggle.addEventListener("click", () => {
    panel.hidden = !panel.hidden;
    toggle.setAttribute("aria-expanded", String(!panel.hidden));
    if (!panel.hidden) {
      input.value = isoToPickerValue(text.value);
      input.focus();
    }
  });
  input.addEventListener("change", apply);
  input.addEventListener("input", apply);
  $(`insert${which}Now`).addEventListener("click", () => {
    input.value = isoToPickerValue("");
    apply();
  });
  $(`insert${which}Clear`).addEventListener("click", () => {
    input.value = "";
    text.value = "";
    updateFormulaPreview();
  });
}

function tagSource() {
  return document.querySelector("input[name=tagSource]:checked").value;
}

function initTagSource() {
  for (const radio of document.querySelectorAll("input[name=tagSource]")) {
    radio.addEventListener("change", () => {
      const range = tagSource() === "range";
      $("tagSourceRange").hidden = !range;
      $("tagSourceList").hidden = range;
      updateFormulaPreview();
    });
  }
  $("insertTagRange").addEventListener("input", updateFormulaPreview);
  $("useSelectedRange").addEventListener("click", () =>
    useSelectedRange().catch((err) => setStatus($("insertStatus"), toCellMessage(err), "error")),
  );
}

/** Put the current Excel selection's address (sheet-qualified) in the range box. */
async function useSelectedRange() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    $("insertTagRange").value = range.address;
  });
  updateFormulaPreview();
}

function initInsertTab() {
  initTagSource();
  initTagPicker();
  initDateTimePicker("Start");
  initDateTimePicker("End");
  renderTagChips();

  const fields = [
    "insertDataset",
    "insertStart",
    "insertEnd",
    "insertInterval",
    "insertAggregate",
    "insertQuality",
  ];
  fields.forEach((id) => $(id).addEventListener("input", updateFormulaPreview));
  fields.forEach((id) => $(id).addEventListener("change", updateFormulaPreview));

  // "Last value (step)" is a different function with different defaults, so
  // the form follows it: a 5-minute interval and good-quality-only.
  $("insertAggregate").addEventListener("change", () => {
    const step = $("insertAggregate").value === "step";
    $("stepHint").hidden = !step;
    $("insertInterval").placeholder = step ? "5 (minutes), or 15m, 1h, 1d" : "5m";
    if (step) $("insertQuality").value = "good";
    updateFormulaPreview();
  });

  $("insertFormula").addEventListener("click", async () => {
    const statusEl = $("insertStatus");
    try {
      const formula = currentInsertFormula();
      await insertFormulaAtActiveCell(formula);
      setStatus(statusEl, "Inserted.", "ok");
    } catch (err) {
      setStatus(statusEl, toCellMessage(err), "error");
    }
  });

  updateFormulaPreview();
}

function currentInsertFormula() {
  let tags = insertTagList;
  let tagsAreReference = false;
  if (tagSource() === "range") {
    const address = $("insertTagRange").value.trim();
    if (!address) return buildDatasetsFormula();
    if (!looksLikeReference(address)) {
      throw new RangeError(`"${address}" is not a cell range. Use something like Sheet1!A2:A20, or select the cells and click "Use selected cells".`);
    }
    tags = address;
    tagsAreReference = true;
  } else if (insertTagList.length === 0) {
    // No tags chosen: fall back to TB.DATASETS() so the preview is never blank
    // and there is always something sensible to insert.
    return buildDatasetsFormula();
  }
  if ($("insertAggregate").value === "step") {
    return buildSampleFormula({
      dataset: $("insertDataset").value,
      tags,
      tagsAreReference,
      start: $("insertStart").value,
      end: $("insertEnd").value,
      interval: $("insertInterval").value,
      quality: $("insertQuality").value,
      columns: extraColumns(),
    });
  }
  return buildHistoryFormula({
    dataset: $("insertDataset").value,
    tags,
    tagsAreReference,
    start: $("insertStart").value,
    end: $("insertEnd").value,
    interval: $("insertInterval").value,
    aggregate: $("insertAggregate").value,
    quality: $("insertQuality").value,
    novalue: noValueMode(),
    columns: extraColumns(),
  });
}

function updateFormulaPreview() {
  try {
    $("formulaPreview").textContent = currentInsertFormula();
  } catch (err) {
    $("formulaPreview").textContent = `(${toCellMessage(err)})`;
  }
}

// --- Office.js ---------------------------------------------------------------------

async function insertFormulaAtActiveCell(formula) {
  await Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.formulas = [[formula]];
    await context.sync();
  });
}

// --- boot ----------------------------------------------------------------------

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  try {
    await loadSettings();
  } catch (err) {
    reportBackground(err);
  }

  $("loadMessage").hidden = true;
  $("app").hidden = false;

  initTabs();
  initConnectTab();
  initNoValueOption();
  initExtraColumns();
  initGuideButton();
  await applyServerAuth();
  initInsertTab();

  if (getSettings().historianUrl) {
    populateDatasetDropdowns().catch(reportBackground);
  }
});
