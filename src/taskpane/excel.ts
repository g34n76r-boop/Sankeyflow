/* global document, window, Excel, Office, XMLSerializer, Image, CSS, Event, setTimeout, clearTimeout, ResizeObserver */
/* global HTMLElement, HTMLSelectElement, HTMLInputElement, HTMLButtonElement, HTMLDetailsElement, HTMLTableRowElement, SVGSVGElement */

import { SankeyError, applyNodePositions, computeSankey, getNodePosition } from "./sankey";
import type { Flow, NodePosition, SankeyLayout, SankeyNode } from "./sankey";
import {
  DARK_THEME,
  LIGHT_THEME,
  defaultNodeColor,
  formatValue,
  nodeColor,
  renderSankey,
  themeForBackground,
  titleHeight,
} from "./chart";
import type {
  ChartHandle,
  ChartTheme,
  ChartTitle,
  ColorChoice,
  LinkColorMode,
  NodeGroup,
} from "./chart";
import {
  DEFAULT_LABELS,
  FONT_CHOICES,
  LABEL_GAP,
  LINE_HEIGHT,
  NAME_WEIGHT,
  formatNumber,
  labelBlock,
  resolvePlacement,
  totalFlow,
  valueText,
} from "./labels";
import type { LabelPlacement, LabelSettings, NodeLabelOverride, ValueMode } from "./labels";
import {
  listedTables,
  openTable,
  releaseTable,
  reloadSelectedTable,
  setupTablePicker,
} from "./tables";
import {
  diagramById,
  findBySource,
  forgetDiagram,
  listDiagrams,
  newDiagramId,
  openDiagramSetting,
  renameDiagram,
  renderMemory,
  saveDiagram,
  setOpenDiagram,
  takeLegacyTable,
} from "./memory";
import type {
  BackgroundMode,
  ColorGroup,
  DiagramSettings,
  GroupSet,
  SavedDiagram,
  SavedPicture,
  SavedSource,
} from "./memory";
import type { TableData } from "./tables";
import { computeRoles } from "./finance";
import { INDUSTRY_PRESETS, industryPreset } from "./domains";
import type { IndustryId, Tone } from "./domains";
import { SAMPLES, sampleById } from "./samples";
import type { NodeRole } from "./finance";
import { closePopout, openPopout, popoutOpen, refreshPopout } from "./popout";
import { layoutToModel } from "./popoutModel";
import { applyZoom, stepZoom, updateZoomButtons, zoomKeepingCentre } from "./zoom";
import { attachPreviewResizer } from "./resize";
import { attachSizeMenu } from "./sizeMenu";
import { setupToolbarIconSelects } from "./iconSelect";
import type { SizeMenu } from "./sizeMenu";
import { attachChartEdges } from "./chartEdges";
import type { ChartEdges } from "./chartEdges";
import type {
  ChartModel,
  PopoutAction,
  PopoutContent,
  ToolbarOptionId,
  ToolbarOptions,
} from "./popoutModel";

/** Narrowest the chart is drawn in the pane; narrower panes scroll sideways. */
const MIN_CHART_WIDTH = 640;
/** Width of the inserted picture (drawn at 2x, so 1920px like the reference image). */
const EXPORT_WIDTH = 960;

/** Where the diagram's data came from (saved as is with the diagram). */
type DataSource = SavedSource;

interface LoadedData {
  flows: Flow[];
  headers: [string, string, string];
  source: DataSource;
}

let current: LoadedData | null = null;

// Customizations, keyed by node name so they survive Refresh, resizing and export.
const nodeColors = new Map<string, ColorChoice>();
const nodePositions = new Map<string, NodePosition>();
const labelOverrides = new Map<string, NodeLabelOverride>();
let labelSettings: LabelSettings = { ...DEFAULT_LABELS };
let selectedNode: string | null = null;

type Margin = { top: number; right: number; bottom: number; left: number };

/** Lines in a label: two when the value sits under the name. */
const labelLines = () =>
  labelSettings.valueMode !== "none" && labelSettings.valueLayout === "below" ? 2 : 1;

/**
 * Chart margins: room above/below so labels centred on the top or bottom box
 * (or placed above/below boxes) aren't cut off. Left/right room for labels is
 * added by buildLayout, which measures them.
 */
function chartMargin(side: number, base: number, withTitle = true): Margin {
  const placements = [
    labelSettings.placement,
    ...Array.from(labelOverrides.values(), (o) => o.placement),
  ];
  const vertical = placements.some((p) => p === "above" || p === "below");
  const lineH = labelSettings.fontSize * LINE_HEIGHT;
  let v = Math.max(base, (labelLines() * lineH) / 2 + 4);
  if (vertical) v = Math.max(v, labelLines() * lineH + LABEL_GAP + 4);
  // The title (if any) sits above everything else.
  return {
    top: v + (withTitle ? titleHeight(chartTitle) : 0),
    right: side,
    bottom: v,
    left: side,
  };
}

/* ---------- Title ---------- */

const DEFAULT_TITLE: ChartTitle = {
  text: "",
  subtitle: "",
  size: 20,
  subtitleSize: null, // automatic
  bold: true,
  italic: false,
  fontFamily: null, // same as the labels
  align: "left",
  color: null,
};

/** Value of the "Same as labels" option in the title font dropdown. */
const SAME_AS_LABELS = "";
let chartTitle: ChartTitle = { ...DEFAULT_TITLE };

function setupTitleControls() {
  const text = $<HTMLInputElement>("title-text");
  const sub = $<HTMLInputElement>("subtitle-text");
  const size = $<HTMLInputElement>("title-size");
  const update = (patch: Partial<ChartTitle>) => {
    chartTitle = { ...chartTitle, ...patch };
    syncTitleControls();
    draw();
  };
  text.addEventListener("input", () => update({ text: text.value }));
  sub.addEventListener("input", () => update({ subtitle: sub.value }));
  size.addEventListener("change", () => {
    const v = Math.round(Number(size.value));
    update({ size: isFinite(v) ? Math.min(40, Math.max(12, v)) : DEFAULT_TITLE.size });
  });

  // Font: "Same as labels" (follows the Labels font) or a font of its own.
  const font = $<HTMLSelectElement>("title-font");
  font.appendChild(
    Object.assign(document.createElement("option"), {
      value: SAME_AS_LABELS,
      textContent: "Same as labels",
    })
  );
  for (const f of FONT_CHOICES) {
    const opt = Object.assign(document.createElement("option"), {
      value: f.stack,
      textContent: f.name,
    });
    opt.style.fontFamily = f.stack;
    font.appendChild(opt);
  }
  font.onchange = () => update({ fontFamily: font.value === SAME_AS_LABELS ? null : font.value });

  const subSize = $<HTMLInputElement>("subtitle-size");
  subSize.addEventListener("change", () => {
    const raw = subSize.value.trim();
    const v = Math.round(Number(raw));
    // Blank (or invalid) = automatic.
    update({ subtitleSize: raw === "" || !isFinite(v) ? null : Math.min(32, Math.max(9, v)) });
  });

  $<HTMLInputElement>("title-bold").onchange = (e) =>
    update({ bold: (e.target as HTMLInputElement).checked });
  $<HTMLInputElement>("title-italic").onchange = (e) =>
    update({ italic: (e.target as HTMLInputElement).checked });
  $<HTMLSelectElement>("title-align").onchange = (e) =>
    update({ align: (e.target as HTMLSelectElement).value as ChartTitle["align"] });
  $<HTMLInputElement>("title-color").addEventListener("input", (e) =>
    update({ color: (e.target as HTMLInputElement).value })
  );
  $("title-color-auto").onclick = () => update({ color: null });
  $("remove-title-btn").onclick = () => update({ ...DEFAULT_TITLE });
  syncTitleControls();
}

/** Writes the title settings into the controls (skipping the one being typed in). */
function syncTitleControls() {
  const t = chartTitle;
  const set = (id: string, value: string) => {
    const el = $<HTMLInputElement | HTMLSelectElement>(id);
    if (document.activeElement !== el && el.value !== value) el.value = value;
  };
  set("title-text", t.text);
  set("subtitle-text", t.subtitle);
  set("title-size", String(t.size));
  set("title-font", t.fontFamily ?? SAME_AS_LABELS);
  set("subtitle-size", t.subtitleSize === null ? "" : String(t.subtitleSize));
  $<HTMLInputElement>("title-bold").checked = t.bold;
  $<HTMLInputElement>("title-italic").checked = t.italic;
  set("title-align", t.align);
  set("title-color", toHex6(t.color ?? chartTheme(false).textPrimary));
  $<HTMLButtonElement>("title-color-auto").disabled = t.color === null;
  const any = t.text.trim() !== "" || t.subtitle.trim() !== "";
  $<HTMLButtonElement>("remove-title-btn").disabled =
    JSON.stringify(t) === JSON.stringify(DEFAULT_TITLE);
  // The section heading shows the title, so it's visible while collapsed.
  $("title-note").textContent = any ? t.text.trim() || t.subtitle.trim() : "None";
}

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;
  setupThemeToggle();
  setupPopout();
  setupTopScale();
  setupZoom();
  setupSizeControls();
  $("sideload-msg").hidden = true;
  $("app-body").hidden = false;

  $("load-btn").onclick = () => run(() => loadFromSelection());
  setupClearConfirm();
  $("refresh-btn").onclick = () =>
    run(async () => {
      if (current?.source.kind === "range") await loadFromAddress(current.source.address);
      else if (current?.source.kind === "table") await reloadSelectedTable();
    });
  $("sample-btn").onclick = () => run(insertSampleData);
  $("insert-btn").onclick = () => run(insertChartIntoSheet);
  $<HTMLSelectElement>("link-color").onchange = () => draw();
  $<HTMLSelectElement>("end-align").onchange = () => draw();
  $<HTMLSelectElement>("node-order").onchange = () => draw();

  $("reset-positions-btn").onclick = () => {
    nodePositions.clear();
    // Dragged label offsets are part of the layout too.
    for (const [name, o] of Array.from(labelOverrides))
      setOverride(name, { ...o, dx: undefined, dy: undefined });
    draw();
  };
  $("reset-colors-btn").onclick = () => {
    nodeColors.clear();
    draw();
  };
  $("reset-node-color").onclick = () => {
    if (selectedNode) setNodeColor(selectedNode, null);
  };
  $("reset-node-position").onclick = () => {
    if (!selectedNode) return;
    nodePositions.delete(selectedNode);
    const o = labelOverrides.get(selectedNode);
    if (o) setOverride(selectedNode, { ...o, dx: undefined, dy: undefined });
    draw();
  };
  $<HTMLInputElement>("custom-color").addEventListener("input", (e) => {
    if (selectedNode) setNodeColor(selectedNode, (e.target as HTMLInputElement).value);
  });

  // Per-node label
  $<HTMLInputElement>("node-label-text").addEventListener("input", (e) => {
    if (!selectedNode) return;
    const text = (e.target as HTMLInputElement).value;
    setOverride(selectedNode, {
      ...labelOverrides.get(selectedNode),
      text: text.trim() ? text : undefined,
    });
    draw();
  });
  $<HTMLSelectElement>("node-label-placement").onchange = (e) => {
    if (!selectedNode) return;
    const v = (e.target as HTMLSelectElement).value as LabelPlacement | "";
    // A new position starts from its own anchor, so drop any old drag offset.
    setOverride(selectedNode, {
      ...labelOverrides.get(selectedNode),
      placement: v || undefined,
      dx: undefined,
      dy: undefined,
    });
    draw();
  };
  $("reset-node-label").onclick = () => {
    if (!selectedNode) return;
    labelOverrides.delete(selectedNode);
    draw();
  };

  setupLabelControls();
  setupBackgroundControls();
  setupToolbarIconSelects();
  setupTabs();
  setupImbalanceMessage();
  setupSummaryToggle();
  setupTitleControls();
  setupSampleControls();
  setupGroupControls();
  setupSaveControls();
  // Every setting as it is before anything is loaded: what Clear goes back to.
  defaultSettings = captureSettings();
  setupTablePicker({
    run,
    onTableData: showTableData,
    onError: (message) => setStatus(message, "error"),
    onNotice: (message) => setStatus(message),
    onWatchEnded: (id) => {
      // If the diagram still shows that table, say it has stopped updating.
      // (When switching tables, the new table's data replaces this right away.)
      if (current?.source.kind === "table" && current.source.id === id) {
        $("source-address").textContent =
          `Table "${current.source.name}" (${current.source.sheet}) · not updating`;
      }
    },
  }).then(() => run(restoreOnStartup));

  // Redraw when the chart's box changes width — from resizing the pane, but also
  // when the pane's scrollbar appears or disappears as content changes.
  let resizeTimer: ReturnType<typeof setTimeout> | undefined;
  let lastWidth = 0;
  const scheduleRedraw = () => {
    const w = $("chart").clientWidth;
    if (w === lastWidth) return;
    lastWidth = w;
    if (resizeTimer) clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => draw(), 120);
  };
  if (typeof ResizeObserver !== "undefined") new ResizeObserver(scheduleRedraw).observe($("chart"));
  else window.addEventListener("resize", scheduleRedraw);
  // The example preview: redraw when its box changes width (pane resized, or
  // shown again after switching back to the Chart tab).
  let previewTimer: ReturnType<typeof setTimeout> | undefined;
  let previewWidth = 0;
  const schedulePreview = () => {
    const w = $("example-preview").clientWidth;
    if (w === previewWidth) return;
    previewWidth = w;
    if (previewTimer) clearTimeout(previewTimer);
    previewTimer = setTimeout(renderExamplePreview, 80);
  };
  if (typeof ResizeObserver !== "undefined") {
    new ResizeObserver(schedulePreview).observe($("example-preview"));
  } else window.addEventListener("resize", schedulePreview);
  // The preview sits in the collapsible "Getting started" section: draw it on opening.
  $("getting-started").addEventListener("toggle", () => {
    if (($("getting-started") as HTMLDetailsElement).open) renderExamplePreview();
  });
  renderExamplePreview();
});

/* ---------- Label controls ---------- */

function setOverride(name: string, o: NodeLabelOverride) {
  const clean: NodeLabelOverride = {};
  if (o.text) clean.text = o.text;
  if (o.placement) clean.placement = o.placement;
  if (o.dx) clean.dx = o.dx;
  if (o.dy) clean.dy = o.dy;
  if (Object.keys(clean).length) labelOverrides.set(name, clean);
  else labelOverrides.delete(name);
}

function updateLabels(patch: Partial<LabelSettings>) {
  labelSettings = { ...labelSettings, ...patch };
  syncLabelControls();
  draw();
}

function setupLabelControls() {
  const font = $<HTMLSelectElement>("label-font");
  for (const f of FONT_CHOICES) {
    const opt = document.createElement("option");
    opt.value = f.stack;
    opt.textContent = f.name;
    opt.style.fontFamily = f.stack;
    font.appendChild(opt);
  }
  font.onchange = () => updateLabels({ fontFamily: font.value });

  const size = $<HTMLInputElement>("label-size");
  size.addEventListener("change", () => {
    const v = Math.round(Number(size.value));
    updateLabels({
      fontSize: isFinite(v) ? Math.min(28, Math.max(8, v)) : DEFAULT_LABELS.fontSize,
    });
  });

  $<HTMLInputElement>("label-bold").onchange = (e) =>
    updateLabels({ boldNames: (e.target as HTMLInputElement).checked });
  $<HTMLInputElement>("label-color").addEventListener("input", (e) =>
    updateLabels({ color: (e.target as HTMLInputElement).value })
  );
  $("label-color-auto").onclick = () => updateLabels({ color: null });
  $<HTMLSelectElement>("label-placement").onchange = (e) =>
    updateLabels({ placement: (e.target as HTMLSelectElement).value as LabelPlacement });
  $<HTMLSelectElement>("value-mode").onchange = (e) =>
    updateLabels({ valueMode: (e.target as HTMLSelectElement).value as ValueMode });
  $<HTMLSelectElement>("value-layout").onchange = (e) =>
    updateLabels({
      valueLayout: (e.target as HTMLSelectElement).value as LabelSettings["valueLayout"],
    });
  $<HTMLSelectElement>("value-decimals").onchange = (e) => {
    const v = (e.target as HTMLSelectElement).value;
    updateLabels({ decimals: v === "" ? null : Number(v) });
  };
  $<HTMLInputElement>("value-prefix").addEventListener("input", (e) =>
    updateLabels({ prefix: (e.target as HTMLInputElement).value })
  );
  $<HTMLInputElement>("value-suffix").addEventListener("input", (e) =>
    updateLabels({ suffix: (e.target as HTMLInputElement).value })
  );
  document.querySelectorAll<HTMLButtonElement>("#symbol-chips .chip").forEach((chip) => {
    chip.onclick = () => {
      const patch: Partial<LabelSettings> = {};
      if (chip.dataset.prefix !== undefined) patch.prefix = chip.dataset.prefix;
      if (chip.dataset.suffix !== undefined) patch.suffix = chip.dataset.suffix;
      // A currency symbol shows numbers, so switch away from percent-only.
      if (labelSettings.valueMode === "percent" || labelSettings.valueMode === "none") {
        if (patch.prefix || patch.suffix) patch.valueMode = "number";
      }
      updateLabels(patch);
    };
  });

  $("reset-labels-btn").onclick = () => {
    labelSettings = { ...DEFAULT_LABELS };
    labelOverrides.clear();
    syncLabelControls();
    draw();
  };

  syncLabelControls();
}

/** Writes the current settings into the controls (skipping the one being typed in). */
function syncLabelControls() {
  const s = labelSettings;
  const set = (id: string, value: string) => {
    const el = $<HTMLInputElement | HTMLSelectElement>(id);
    if (document.activeElement !== el && el.value !== value) el.value = value;
  };
  set("label-font", s.fontFamily);
  set("label-size", String(s.fontSize));
  $<HTMLInputElement>("label-bold").checked = s.boldNames;
  set("label-placement", s.placement);
  set("value-mode", s.valueMode);
  set("value-layout", s.valueLayout);
  $<HTMLSelectElement>("value-layout").disabled = s.valueMode === "none";
  set("value-decimals", s.decimals === null ? "" : String(s.decimals));
  set("value-prefix", s.prefix);
  set("value-suffix", s.suffix);
  const theme = chartTheme(false);
  set("label-color", toHex6(s.color ?? theme.textPrimary));
  $<HTMLButtonElement>("label-color-auto").disabled = s.color === null;
  $<HTMLSelectElement>("value-decimals").disabled = s.valueMode === "none";
}

function renderLabelPreview(layout: SankeyLayout, theme: ChartTheme) {
  const preview = $("label-preview");
  preview.innerHTML = "";
  // Preview the biggest box's label.
  const n = layout.nodes.reduce((a, b) => (b.value > a.value ? b : a), layout.nodes[0]);
  if (!n) return;
  const s = labelSettings;
  preview.style.fontFamily = s.fontFamily;
  preview.style.fontSize = `${s.fontSize}px`;
  preview.style.color = s.color ?? theme.textPrimary;
  const name = document.createElement("span");
  name.style.fontWeight = s.boldNames ? String(NAME_WEIGHT) : "400";
  name.textContent = labelOverrides.get(n.name)?.text || n.name;
  preview.appendChild(name);
  const value = valueText(n, totalFlow(layout), s);
  if (value) {
    const v = document.createElement("span");
    v.style.marginLeft = "0.4em";
    v.style.opacity = s.color ? "0.75" : "1";
    if (!s.color) v.style.color = theme.textSecondary;
    v.textContent = value;
    preview.appendChild(v);
  }
}

/* ---------- Tabs (Chart / Data / Memory) ---------- */

/** Shows a tab ("chart", "data" or "memory"). Set up by setupTabs. */
let showTab: (id: "chart" | "data" | "memory") => void = () => undefined;

function setupTabs() {
  const tabs = [
    $<HTMLButtonElement>("tab-chart"),
    $<HTMLButtonElement>("tab-data"),
    $<HTMLButtonElement>("tab-memory"),
  ];
  const select = (tab: HTMLButtonElement, focus: boolean) => {
    for (const t of tabs) {
      const on = t === tab;
      t.setAttribute("aria-selected", String(on));
      t.tabIndex = on ? 0 : -1;
      $(t.getAttribute("aria-controls")!).hidden = !on;
    }
    // The list of saved diagrams is read fresh each time it's shown.
    if (tab.id === "tab-memory") renderMemoryTab();
    if (focus) tab.focus();
  };
  showTab = (id) => select($<HTMLButtonElement>(`tab-${id}`), false);
  // "…under the Data tab." link in the Chart tab's empty state.
  $("goto-data-tab").onclick = () => select(tabs[1], true);
  tabs.forEach((tab) => {
    tab.onclick = () => select(tab, false);
    // Arrow keys move between the visible tabs, as in other tab strips.
    tab.addEventListener("keydown", (e) => {
      const shown = tabs.filter((t) => !t.hidden);
      const i = shown.indexOf(tab);
      const next =
        e.key === "ArrowRight"
          ? shown[(i + 1) % shown.length]
          : e.key === "ArrowLeft"
            ? shown[(i - 1 + shown.length) % shown.length]
            : e.key === "Home"
              ? shown[0]
              : e.key === "End"
                ? shown[shown.length - 1]
                : null;
      if (next) {
        e.preventDefault();
        select(next, true);
      }
    });
  });
}

/* ---------- Background ---------- */

const backgroundMode = () => $<HTMLSelectElement>("background-mode").value as BackgroundMode;

/**
 * The chart's theme for the chosen background. "Auto" follows the pane's
 * light/dark mode in the pane and is white in the inserted picture. "Transparent" lays out as if
 * on white (the usual worksheet color) but draws no background.
 */
function chartTheme(forExport: boolean): ChartTheme {
  return baseTheme(forExport);
}

function baseTheme(forExport: boolean): ChartTheme {
  switch (backgroundMode()) {
    case "white":
    case "transparent":
      return themeForBackground("#ffffff");
    case "dark":
      // The dark-mode look, in the pane and in the inserted picture.
      return DARK_THEME;
    case "custom":
      return themeForBackground($<HTMLInputElement>("background-color").value);
    default:
      return forExport || !isDarkTheme() ? LIGHT_THEME : DARK_THEME;
  }
}

function setupBackgroundControls() {
  const mode = $<HTMLSelectElement>("background-mode");
  const color = $<HTMLInputElement>("background-color");
  mode.onchange = () => {
    color.hidden = mode.value !== "custom";
    // Label color "Auto" adapts to the background; refresh the swatch in Labels.
    syncLabelControls();
    draw();
  };
  color.addEventListener("input", () => {
    syncLabelControls();
    draw();
  });
}

/* ---------- UI helpers ---------- */

/* ---------- Light / dark toggle ---------- */

const THEME_KEY = "sankeyflow.paneTheme";

/** Your light/dark choice; null until you use the toggle (then the pane is light). */
let themeChoice: "light" | "dark" | null = (() => {
  try {
    const v = window.localStorage.getItem(THEME_KEY);
    return v === "light" || v === "dark" ? v : null;
  } catch {
    return null; // storage can be unavailable; default to light
  }
})();

function setupThemeToggle() {
  applyPaneTheme();
  $("theme-toggle").onclick = () => {
    themeChoice = isDarkTheme() ? "light" : "dark";
    try {
      window.localStorage.setItem(THEME_KEY, themeChoice);
    } catch {
      // Not remembered next time, but still switches now.
    }
    applyPaneTheme();
    // The Labels color swatch and the chart preview follow the theme.
    syncLabelControls();
    draw();
  };
}

/** Sun: a solid disc with eight short, separate rays. */
const SUN_ICON =
  `<svg class="sun-icon" viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" focusable="false">` +
  `<circle cx="12" cy="12" r="4.5" fill="currentColor"/>` +
  `<g stroke="currentColor" stroke-width="2" stroke-linecap="round">` +
  `<line x1="12" y1="1.5" x2="12" y2="4"/><line x1="12" y1="20" x2="12" y2="22.5"/>` +
  `<line x1="1.5" y1="12" x2="4" y2="12"/><line x1="20" y1="12" x2="22.5" y2="12"/>` +
  `<line x1="4.6" y1="4.6" x2="6.3" y2="6.3"/><line x1="17.7" y1="17.7" x2="19.4" y2="19.4"/>` +
  `<line x1="4.6" y1="19.4" x2="6.3" y2="17.7"/><line x1="17.7" y1="6.3" x2="19.4" y2="4.6"/>` +
  `</g></svg>`;

/** Sets the pane's theme and the toggle's icon and label. */
function applyPaneTheme() {
  const dark = isDarkTheme();
  document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
  const btn = $("theme-toggle");
  // The icon shows the current mode: moon in dark mode, sun in light mode.
  if (dark) {
    btn.textContent = "☾";
  } else {
    // Drawn rather than the ☀ glyph, which some fonts render as a starburst.
    btn.innerHTML = SUN_ICON;
  }
  const label = dark ? "Switch to light mode" : "Switch to dark mode";
  btn.title = label;
  btn.setAttribute("aria-label", label);
  btn.setAttribute("aria-pressed", String(dark));
}

/** Dark mode: your toggle choice if you've made one, otherwise light. */
function isDarkTheme(): boolean {
  return themeChoice === "dark";
}

function setStatus(message: string | null, kind: "info" | "error" = "info") {
  const s = $("status");
  s.hidden = !message;
  $("status-text").textContent = message ?? "";
  s.classList.toggle("error", kind === "error");
  // The × hides the message. (Not a .btn, so run() leaves it enabled.)
  $("status-dismiss").onclick = () => setStatus(null);
}

async function run(action: () => Promise<void>) {
  // Only the main action buttons; the customize controls manage their own state.
  const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>("button.btn"));
  const wasDisabled = buttons.map((b) => b.disabled);
  buttons.forEach((b) => (b.disabled = true));
  try {
    await action();
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    setStatus(msg, "error");
  } finally {
    buttons.forEach((b, i) => (b.disabled = wasDisabled[i]));
    $<HTMLButtonElement>("refresh-btn").disabled = !current;
    $<HTMLButtonElement>("insert-btn").disabled = !current;
    $<HTMLButtonElement>("clear-btn").disabled = !current;
    updateSaveState();
  }
}

/* ---------- Reading data from Excel ---------- */

async function loadFromSelection() {
  await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    range.load("rowCount,columnCount");
    await context.sync();
    // A single selected cell: use the table-like block around it.
    if (
      range.rowCount === 1 &&
      range.columnCount === 1 &&
      Office.context.requirements.isSetSupported("ExcelApi", "1.7")
    ) {
      range = range.getSurroundingRegion();
    }
    range.load("values,address");
    await context.sync();
    applyValues(range.values, range.address);
  });
  // The diagram now comes from the selection, so stop following any table.
  await releaseTable();
}

async function loadFromAddress(fullAddress: string) {
  await Excel.run(async (context) => {
    const { sheet, address } = splitAddress(fullAddress);
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    range.load("values,address");
    await context.sync();
    applyValues(range.values, range.address);
  });
}

function splitAddress(full: string): { sheet: string; address: string } {
  const i = full.lastIndexOf("!");
  let sheet = full.slice(0, i);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  return { sheet, address: full.slice(i + 1) };
}

function toNumber(v: unknown): number {
  if (typeof v === "number") return v;
  if (typeof v === "string" && v.trim() !== "") return Number(v.replace(/[,\s]/g, ""));
  return NaN;
}

function applyValues(values: unknown[][], address: string) {
  if (!values.length || values[0].length < 3) {
    throw new Error("Select at least three columns: Source, Target and Value.");
  }

  let headers: [string, string, string] = ["Source", "Target", "Value"];
  const flows: Flow[] = [];
  let skipped = 0;

  values.forEach((row, i) => {
    const source = String(row[0] ?? "").trim();
    const target = String(row[1] ?? "").trim();
    const value = toNumber(row[2]);
    if (!source && !target && (row[2] === "" || row[2] == null)) return; // blank row
    if (i === 0 && !isFinite(value)) {
      headers = [source || "Source", target || "Target", String(row[2] ?? "") || "Value"];
      return;
    }
    if (!source || !target || !isFinite(value) || value <= 0) {
      skipped++;
      return;
    }
    flows.push({ source, target, value });
  });

  if (!flows.length) {
    throw new Error(
      "No usable rows found. Each row needs a source, a target and a positive number."
    );
  }

  const notes: string[] = [];
  if (values[0].length > 3) notes.push("Only the first three columns are used.");
  if (skipped)
    notes.push(
      `Skipped ${skipped} row${skipped === 1 ? "" : "s"} with a missing name or a value that isn't a positive number.`
    );
  showFlows(flows, headers, { kind: "range", address }, notes);
}

/**
 * Diagram data from a watched table (when picked, and after each change).
 * Errors are shown in the status area instead of thrown, and the last good
 * diagram stays on screen, so a half-typed edit doesn't wipe it out.
 */
function showTableData(data: TableData, reason: "picked" | "changed") {
  const source: DataSource = { kind: "table", id: data.id, name: data.name, sheet: data.sheet };
  const label = `Table "${data.name}"`;
  if (data.columns < 3) {
    setStatus(`${label} needs at least three columns: Source, Target and Value.`, "error");
    return;
  }

  // The list shows every numeric row; the diagram also needs names and positive values.
  const flows: Flow[] = [];
  let unusable = 0;
  for (const r of data.rows) {
    if (r.source && r.target && r.value > 0) flows.push({ ...r });
    else unusable++;
  }
  const notes: string[] = [];
  if (data.columns > 3) notes.push("Only the first three columns are used.");
  if (unusable)
    notes.push(
      `${unusable} row${unusable === 1 ? " isn't" : "s aren't"} in the diagram (missing name, or zero/negative value).`
    );

  if (!flows.length) {
    setStatus(
      `${label} has no rows the diagram can use yet. Each row needs a source, a target and a positive number.` +
        (reason === "changed" ? " The diagram will update when it does." : ""),
      "error"
    );
    return;
  }
  try {
    showFlows(flows, data.headers, source, notes);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    setStatus(
      `${label}: ${msg}` + (reason === "changed" ? " Showing the last valid diagram." : ""),
      "error"
    );
  }
}

function sourceLabel(s: DataSource): string {
  if (s.kind === "range") return s.address;
  const time = new Date().toLocaleTimeString(undefined, {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  return `Table "${s.name}" (${s.sheet}) · live, updated ${time}`;
}

/** Shows a validated set of flows as the diagram. Throws (leaving the screen as it was) on loops. */
function showFlows(
  flows: Flow[],
  headers: [string, string, string],
  source: DataSource,
  notes: string[]
) {
  // Validate the shape before replacing what's on screen (throws on loops).
  computeSankey(flows, { width: 300, height: 300, iterations: 0 });

  // Data that has a saved diagram (a table or range you saved before) opens
  // with that diagram's settings.
  if (!openDiagramId) {
    const saved = findBySource(source);
    if (saved) {
      applySettings(saved.settings);
      markOpen(saved.id);
      setOpenDiagram(saved.id).catch(() => undefined);
      notes.push(`Opened with the settings saved as "${saved.name}".`);
    }
  }

  current = { flows, headers, source };
  setLoadedMode(true);
  setStatus(notes.length ? notes.join(" ") : null);
  $<HTMLButtonElement>("refresh-btn").disabled = false;
  $<HTMLButtonElement>("insert-btn").disabled = false;
  $<HTMLButtonElement>("clear-btn").disabled = false;

  $("empty-state").hidden = true;
  $("chart-section").hidden = false;
  $("source-address").textContent = sourceLabel(source);
  $("chart-meta").hidden = false; // in the Data tab, above the Flows table
  // Totals are shown in the summary under the top buttons.
  $("totals").textContent = `${flows.length} flow${flows.length === 1 ? "" : "s"}`;
  fillTable();
  draw();
}

/* ---------- Clearing the loaded data ---------- */

/**
 * Unloads the data: stops following a table, removes the chart, summary and
 * flows table, shows the data choices and empty state again, and puts every
 * setting back to its default (save first to keep them). With `persist`, the
 * workbook also stops reopening this diagram next time.
 */
async function clearData(opts: { persist?: boolean } = {}) {
  await releaseTable();
  current = null;
  selectedNode = null;
  chartHandle = null;
  imbalanceHover = false;
  openDiagramId = null;
  savedSnapshot = null;
  pendingPictures = [];
  applySettings(defaultSettings!);
  setLoadedMode(false);
  if (opts.persist !== false) {
    try {
      await setOpenDiagram(null);
    } catch {
      // Not fatal: it would just reopen next time.
    }
  }

  // Chart tab back to the "your data should look like this" state.
  const svg = document.getElementById("sankey-svg")!;
  while (svg.firstChild) svg.removeChild(svg.firstChild);
  $("tooltip").hidden = true;
  $("chart-section").hidden = true;
  $("empty-state").hidden = false;
  renderExamplePreview();
  refreshPopout();

  // Summary table and alert (its space stays reserved, just invisible).
  $("stage-summary").hidden = true;
  // No summary to make room for: the text and load controls can use the full width.
  document.querySelector(".top-bar")!.classList.add("no-data");
  showImbalanceMessage(false);

  // Data tab.
  $("chart-meta").hidden = true;
  $("flow-table").innerHTML = "";
  $("flows-wrap").hidden = true;
  $("flows-empty").hidden = false;
  $("sample-section").hidden = false;
  $("data-gaps").hidden = true;
  dataGaps = new Map();
  currentImbalance = tableImbalance = "";
  dismissedImbalance = null;

  setStatus(null);
  updateSaveState();
  renderMemoryTab();
}

/* ---------- Saved diagrams (Save settings, the Memory tab) ---------- */

/** What Clear goes back to: every setting as it was when the pane opened. */
let defaultSettings: DiagramSettings | null = null;
/** The saved diagram shown in the pane (null = not saved yet, or nothing loaded). */
let openDiagramId: string | null = null;
/** The open diagram's settings as last saved (JSON), to tell if there are unsaved changes. */
let savedSnapshot: string | null = null;
/** Pictures inserted before the diagram was first saved; saved with it. */
let pendingPictures: SavedPicture[] = [];

const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v));

/** Every setting of the diagram in the pane, as saved with it. */
function captureSettings(): DiagramSettings {
  return clone({
    nodeColors: Array.from(nodeColors),
    nodePositions: Array.from(nodePositions),
    labelOverrides: Array.from(labelOverrides),
    labelSettings,
    chartTitle,
    groupSet,
    groups,
    genericTheme,
    addedGroupCount,
    background: { mode: backgroundMode(), color: $<HTMLInputElement>("background-color").value },
    linkColor: $<HTMLSelectElement>("link-color").value as LinkColorMode,
    endAlign: $<HTMLSelectElement>("end-align").value as "left" | "justify",
    nodeOrder: $<HTMLSelectElement>("node-order").value as "data" | "color",
    chartSize,
    summaryCollapsed,
  });
}

/**
 * Puts saved settings into the pane (state and controls). Anything missing
 * (e.g. saved by an older version) falls back to the default.
 */
function applySettings(saved: Partial<DiagramSettings>) {
  const s: DiagramSettings = { ...defaultSettings!, ...clone(saved) };
  const refill = <V>(map: Map<string, V>, entries: [string, V][]) => {
    map.clear();
    for (const [k, v] of entries ?? []) map.set(k, v);
  };
  refill(nodeColors, s.nodeColors);
  refill(nodePositions, s.nodePositions);
  refill(labelOverrides, s.labelOverrides);
  labelSettings = { ...DEFAULT_LABELS, ...s.labelSettings };
  chartTitle = { ...DEFAULT_TITLE, ...s.chartTitle };
  groupSet = s.groupSet;
  groups = s.groups;
  genericTheme = s.genericTheme;
  addedGroupCount = s.addedGroupCount;
  chartSize = { width: s.chartSize?.width ?? null, height: s.chartSize?.height ?? null };
  summaryCollapsed = !!s.summaryCollapsed;
  selectedNode = null;

  const mode = $<HTMLSelectElement>("background-mode");
  const color = $<HTMLInputElement>("background-color");
  mode.value = s.background.mode;
  color.value = s.background.color;
  color.hidden = mode.value !== "custom";
  $<HTMLSelectElement>("link-color").value = s.linkColor;
  $<HTMLSelectElement>("end-align").value = s.endAlign;
  $<HTMLSelectElement>("node-order").value = s.nodeOrder;

  // The group cards are rebuilt for the new groups.
  $("group-rows").dataset.sig = "";
  syncTitleControls();
  syncLabelControls();
  syncColourSelect();
  syncSizeControls();
}

/** Marks a saved diagram as the one in the pane, with its settings as just applied. */
function markOpen(id: string) {
  openDiagramId = id;
  savedSnapshot = JSON.stringify(captureSettings());
  pendingPictures = [];
}

/** The loaded chart has changes that aren't saved (or was never saved). */
function hasUnsavedChanges(): boolean {
  if (!current) return false;
  return !diagramById(openDiagramId) || JSON.stringify(captureSettings()) !== savedSnapshot;
}

/** "Save" when there's something to save, "Saved ✓" when not; and the diagram's name. */
function updateSaveState() {
  const btn = $<HTMLButtonElement>("save-btn");
  const dirty = hasUnsavedChanges();
  const saved = !!current && !dirty;
  btn.disabled = !dirty;
  // Only the label changes; the floppy-disk icon beside it stays.
  $("save-label").textContent = saved ? "Saved ✓" : "Save";
  btn.classList.toggle("is-saved", saved);
  btn.title = saved
    ? "This chart's data link and settings are saved in the workbook"
    : "Save this chart's data link and settings in the workbook";

  // Nothing loaded but the workbook has saved charts: the Memory tab turns
  // green to point you to them.
  const memoryTab = $("tab-memory");
  const count = current ? 0 : listDiagrams().length;
  memoryTab.classList.toggle("has-saved", count > 0);
  memoryTab.title = count
    ? `${count} saved chart${count === 1 ? "" : "s"} in this workbook. Open one from here.`
    : "";

  const info = $("loaded-info");
  info.textContent = "";
  if (!current) return;
  const d = diagramById(openDiagramId);
  const name = document.createElement("b");
  name.textContent = d ? d.name : "Unsaved chart";
  info.append(
    name,
    d
      ? dirty
        ? " · unsaved changes"
        : " · saved in this workbook"
      : " · Save it to keep its look for next time."
  );
}

/**
 * Data chosen: the data choices give way to [Close] [Save],
 * and the Memory tab goes away: close the chart before opening another.
 */
function setLoadedMode(on: boolean) {
  $("select-actions").hidden = on;
  $("select-hint").hidden = on;
  $("loaded-actions").hidden = !on;
  $("loaded-info").hidden = !on;
  $("insert-btn").hidden = !on;
  const memoryTab = $("tab-memory");
  if (on && memoryTab.getAttribute("aria-selected") === "true") showTab("chart");
  memoryTab.hidden = on;
}

/** A name for a newly saved diagram: its table's name or its cells, made unique. */
function defaultDiagramName(source: DataSource): string {
  const base = source.kind === "table" ? source.name : source.address.replace(/'/g, "");
  const taken = new Set(listDiagrams().map((d) => d.name));
  let name = base;
  for (let i = 2; taken.has(name); i++) name = `${base} (${i})`;
  return name;
}

/** Saves the loaded chart (its data link and settings) in the workbook. */
async function saveCurrent() {
  if (!current) return;
  const existing = diagramById(openDiagramId);
  const d: SavedDiagram = {
    id: existing?.id ?? newDiagramId(),
    name: existing?.name ?? defaultDiagramName(current.source),
    source: clone(current.source),
    settings: captureSettings(),
    pictures: [...(existing?.pictures ?? []), ...pendingPictures],
    savedAt: Date.now(),
  };
  await saveDiagram(d);
  openDiagramId = d.id;
  savedSnapshot = JSON.stringify(d.settings);
  pendingPictures = [];
  setStatus(`Saved "${d.name}" in this workbook. It's listed in the Memory tab.`);
  updateSaveState();
  renderMemoryTab();
}

/**
 * Opens a saved diagram in place of whatever is loaded: its settings, then its
 * data (the table, or the cells at its address, read again now).
 */
async function openDiagram(d: SavedDiagram) {
  await clearData({ persist: false });
  applySettings(d.settings);
  markOpen(d.id);
  let problem: string | null = null;
  if (d.source.kind === "table") {
    if (!(await openTable(d.source))) {
      problem = `the table "${d.source.name}" is no longer in this workbook`;
    }
  } else {
    try {
      await loadFromAddress(d.source.address);
    } catch (e) {
      problem = `couldn't read ${d.source.address} (${e instanceof Error ? e.message : String(e)})`;
    }
  }
  if (problem) {
    // Back to a clean pane; the diagram stays in Memory, to fix or forget.
    await clearData();
    setStatus(
      `Couldn't open "${d.name}": ${problem}. You can forget it in the Memory tab.`,
      "error"
    );
    return;
  }
  try {
    await setOpenDiagram(d.id);
  } catch {
    // Not fatal: it just won't reopen by itself next time.
  }
  // The table may have no usable rows yet (the status says so); it opens when it
  // does. Once it has opened, no message: the chart itself shows it worked.
  if (current) setStatus(null);
  updateSaveState();
  renderMemoryTab();
}

/**
 * Forgets a saved diagram: removes everything saved for it and, if it's the
 * one in the pane, unloads it. Its pictures stay in the sheet, renamed so
 * they're no longer marked as SankeyFlow's.
 */
async function forget(d: SavedDiagram) {
  if (d.pictures.length && Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    try {
      await Excel.run(async (context) => {
        // Searched on every sheet: the picture may have been moved, or its sheet renamed.
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        const shapes = sheets.items.map((ws) => ws.shapes.load("items/id"));
        await context.sync();
        const ids = new Set(d.pictures.map((p) => p.shapeId));
        for (const coll of shapes) {
          for (const shape of coll.items) if (ids.has(shape.id)) shape.name = UNLINKED_PICTURE_NAME;
        }
        await context.sync();
      });
    } catch {
      // The pictures may be gone already; forgetting goes ahead regardless.
    }
  }
  await forgetDiagram(d.id);
  if (d.id === openDiagramId) await clearData({ persist: false });
  setStatus(`Forgot "${d.name}".`);
  renderMemoryTab();
}

/** A picture's name in the sheet while it belongs to a diagram, and after it's forgotten. */
const pictureName = (diagram: string) => `SankeyFlow: ${diagram}`.slice(0, 250);
const UNLINKED_PICTURE_NAME = "Sankey chart";

/** The Memory tab's list of saved diagrams. */
function renderMemoryTab() {
  renderMemory({
    openId: openDiagramId,
    tables: listedTables(),
    themeName: (d) => `Style: ${themeLabel(d.settings)}`,
    onOpen: requestOpen,
    onForget: (d) => run(() => forget(d)),
    onRename: (d, name) =>
      run(async () => {
        await renameDiagram(d.id, name);
        updateSaveState();
        renderMemoryTab();
      }),
  });
}

/** A saved diagram's colour theme, as named in the Style dropdown. */
function themeLabel(s: DiagramSettings): string {
  if (s.groupSet === "money") return "Cashflow";
  if (s.groupSet === "generic") {
    return GROUP_THEMES.find((t) => t.id === s.genericTheme)?.name ?? "Generic";
  }
  return industryPreset(s.groupSet)?.name ?? "Custom";
}

/**
 * Open (Memory tab). The tab is only there while no chart is loaded, so
 * there's nothing unsaved to lose.
 */
function requestOpen(d: SavedDiagram) {
  run(async () => {
    await openDiagram(d);
    showTab("chart");
  });
}

/** Save settings. */
function setupSaveControls() {
  $("save-btn").onclick = () => run(saveCurrent);
}

/** When the pane opens: the diagram that was open last time, if any. */
async function restoreOnStartup() {
  const d = diagramById(openDiagramSetting());
  if (d) {
    await openDiagram(d);
  } else {
    // Workbooks from before saved diagrams remember just a table.
    const legacy = await takeLegacyTable();
    if (legacy && !(await openTable(legacy))) {
      setStatus(
        `The table "${legacy.name}" this workbook used last time no longer exists. Pick another table.`
      );
    }
  }
  renderMemoryTab();
}

/* ---------- Drawing ---------- */

/** Vertical room each box needs: its label's height plus some breathing space. */
const rowSpace = () => labelLines() * labelSettings.fontSize * LINE_HEIGHT;

/** Layout options shared by every layout of the current data. */
function sankeyOptions(width: number, height: number) {
  return {
    width,
    height,
    align: $<HTMLSelectElement>("end-align").value as "left" | "justify",
    // Wide, solid boxes (about 3% of the chart width, as in the reference style).
    nodeWidth: Math.round(Math.min(22, Math.max(8, width * 0.031))),
    // Enough space between boxes in a column for their labels.
    nodePadding: Math.max(14, rowSpace() + 6),
    maxGapShare: GAP_SHARE,
    groupKeys: $<HTMLSelectElement>("node-order").value === "color" ? colorGroupKeys : undefined,
  };
}

/** Share of the chart height the gaps between boxes may use (the rest is for boxes). */
const GAP_SHARE = 0.6;

function chartHeight(flows: Flow[], width: number): number {
  // Size the chart to the busiest column so small nodes and their labels stay readable.
  const probe = computeSankey(flows, { ...sankeyOptions(width, 400), iterations: 0 });
  const perColumn = new Array(probe.columnCount).fill(0);
  probe.nodes.forEach((n) => perColumn[n.column]++);
  const rows = Math.max(...perColumn);
  // Tall enough that the gaps in the busiest column fit a label each, without
  // the gaps using more than GAP_SHARE of the height.
  const forLabels = rows > 1 ? ((rows - 1) * (rowSpace() + 6)) / (GAP_SHARE - 0.05) : 0;
  return Math.max(220, Math.min(900, Math.max(rows * 44, forLabels)));
}

/**
 * Computes the layout for a given outer width, leaving room at the sides for
 * labels that sit outside the boxes (starting boxes are labelled on the left,
 * end boxes on the right), and applies the user's moved positions.
 */
function buildLayout(
  width: number,
  base: Margin,
  // The example preview passes its own flows, without your moved boxes or label edits.
  flows: Flow[] = current!.flows,
  personal = true,
  // A chosen overall height (the Size setting); otherwise sized to the busiest column.
  totalHeight: number | null = null
): { layout: SankeyLayout; margin: Margin } {
  const positions = personal ? nodePositions : new Map<string, NodePosition>();
  const overrides = personal ? labelOverrides : new Map<string, NodeLabelOverride>();
  let innerW = Math.max(120, width - base.left - base.right);
  const height = totalHeight
    ? Math.max(80, totalHeight - base.top - base.bottom)
    : chartHeight(flows, innerW);

  // A quick first layout tells us how far labels stick out past each edge.
  const probe = computeSankey(flows, { ...sankeyOptions(innerW, height), iterations: 0 });
  applyNodePositions(probe, positions);
  const total = totalFlow(probe);
  let padL = 0;
  let padR = 0;
  for (const n of probe.nodes) {
    const o = overrides.get(n.name);
    const p = resolvePlacement(n, probe, labelSettings, o);
    if (p !== "left" && p !== "right") continue;
    const w = labelBlock(n, total, labelSettings, o).width + LABEL_GAP + 6;
    if (p === "left") padL = Math.max(padL, w - n.x0);
    else padR = Math.max(padR, n.x1 + w - innerW);
  }
  // Never let labels take more than about a third of the width on either side.
  const cap = width * 0.35;
  padL = Math.min(cap, Math.max(0, Math.ceil(padL)));
  padR = Math.min(cap, Math.max(0, Math.ceil(padR)));
  innerW = Math.max(120, innerW - padL - padR);

  const layout = computeSankey(flows, sankeyOptions(innerW, height));
  applyNodePositions(layout, positions);
  return { layout, margin: { ...base, left: base.left + padL, right: base.right + padR } };
}

/* ---------- Example preview (Chart tab, before any data is loaded) ---------- */

/** The example table above the preview (keep the two in step). */
const EXAMPLE_FLOWS: Flow[] = [
  { source: "Salary", target: "Budget", value: 4200 },
  { source: "Budget", target: "Housing", value: 2000 },
  { source: "Budget", target: "Savings", value: 2200 },
];

/** Draws the example data with the current style (theme, labels, background…). */
function renderExamplePreview() {
  const box = $("example-preview");
  const svg = document.getElementById("example-svg") as unknown as SVGSVGElement;
  if ($("empty-state").hidden || !box.clientWidth) return;
  try {
    const theme = chartTheme(false);
    const width = Math.max(box.clientWidth, 300);
    // The title isn't part of the example.
    const { layout, margin } = buildLayout(width, chartMargin(4, 10, false), EXAMPLE_FLOWS, false);
    renderSankey(svg, layout, {
      theme,
      linkColor: $<HTMLSelectElement>("link-color").value as LinkColorMode,
      labels: labelSettings,
      labelOverrides: new Map(),
      colors: new Map(),
      groups: groupsFor(layout.nodes, theme),
      margin,
      transparent: backgroundMode() === "transparent",
    });
    box.classList.toggle("transparent-bg", backgroundMode() === "transparent");
  } catch {
    // A preview isn't worth an error message.
  }
}

function draw(opts: { focusSelected?: boolean } = {}) {
  // Whatever changed, the pop-out window (if open) follows, and Save settings
  // shows whether there's something to save.
  refreshPopout();
  updateSaveState();
  if (!current) {
    // Nothing loaded: keep the example preview in the current style instead.
    renderExamplePreview();
    return;
  }
  const container = $("chart");
  const svg = document.getElementById("sankey-svg") as unknown as SVGSVGElement;
  const tooltip = $("tooltip");
  const theme = chartTheme(false);
  const transparent = backgroundMode() === "transparent";
  // A checkerboard behind a transparent chart shows that it has no background.
  container.classList.toggle("transparent-bg", transparent);
  tooltip.hidden = true;
  try {
    // Labels sit beside the boxes, which needs room: below MIN_CHART_WIDTH the
    // chart keeps that width and is shown scaled down instead of squashing.
    // A chosen size (the size menu) wins over fitting the pane.
    const width =
      chartSize.width ?? Math.max(container.clientWidth - CHART_EDGE_GAP, MIN_CHART_WIDTH);
    const { layout, margin } = buildLayout(
      width,
      chartMargin(4, 10),
      undefined,
      true,
      chartSize.height
    );
    const gm = groupsFor(layout.nodes, theme);
    if (selectedNode && !layout.nodes.some((n) => n.name === selectedNode)) selectedNode = null;
    chartHandle = renderSankey(svg, layout, {
      theme,
      linkColor: $<HTMLSelectElement>("link-color").value as LinkColorMode,
      labels: labelSettings,
      labelOverrides,
      colors: chartColors(),
      groups: gm,
      margin,
      transparent,
      title: chartTitle,
      interactive: {
        selected: selectedNode,
        focusSelected: opts.focusSelected,
        onNodeSelect: (n) => selectNode(n.name, true),
        // Clicking empty background deselects (removes the outline, closes the box editor).
        onBackgroundClick: () => {
          if (!selectedNode) return;
          selectedNode = null;
          draw();
        },
        onNodeMoved: (n, l) => {
          nodePositions.set(n.name, getNodePosition(n, l));
          updateResetButtons();
          markMoved(n.name);
          refreshPopout();
          updateSaveState();
        },
        onLabelMoved: (n, dx, dy) => {
          setOverride(n.name, { ...labelOverrides.get(n.name), dx, dy });
          updateResetButtons();
          markMoved(n.name);
          refreshPopout();
          updateSaveState();
        },
        onHover: (h) => {
          if (!h) {
            tooltip.hidden = true;
            return;
          }
          tooltip.innerHTML = h.html;
          tooltip.hidden = false;
          // Keep the tooltip inside the chart box.
          const pad = 12;
          const w = tooltip.offsetWidth;
          const ht = tooltip.offsetHeight;
          // Keep it within the visible (possibly scrolled) part of the chart box.
          let x = h.x + pad;
          if (x + w > container.scrollLeft + container.clientWidth) x = h.x - w - pad;
          let y = h.y + pad;
          if (y + ht > container.clientHeight) y = h.y - ht - pad;
          tooltip.style.left = `${Math.max(0, x)}px`;
          tooltip.style.top = `${Math.max(0, y)}px`;
        },
      },
    });
    // Shown scaled down if need be so the whole chart fits the preview's width
    // (and its height follows), then magnified by the + / − buttons.
    const drawnWidth = Number(svg.getAttribute("width")) || width;
    paneFit = Math.min(1, (container.clientWidth - CHART_EDGE_GAP) / drawnWidth);
    applyZoom(svg, paneZoom * paneFit);
    placeResizeHandle?.();
    syncSizeControls();
    renderCustomize(layout, theme, gm);
    renderSchemeHint(theme, layout, gm);
    renderTotals(layout);
  } catch (e) {
    setStatus(e instanceof SankeyError ? e.message : String(e), "error");
  }
}

/* ---------- Totals summary ---------- */

/**
 * Total in (everything leaving the starting boxes), total out (everything
 * reaching the end boxes) and the difference, formatted with the label symbols.
 */
function renderTotals(layout: SankeyLayout) {
  const totalIn = totalFlow(layout);
  const totalOut = layout.nodes
    .filter((n) => !n.sourceLinks.length)
    .reduce((s, n) => s + n.inValue, 0);

  // One row per account (every box with money both in and out), left to right
  // through the chart then top to bottom, and the overall total at the bottom.
  const accounts = layout.nodes
    .filter((n) => n.targetLinks.length && n.sourceLinks.length)
    .sort((a, b) => a.column - b.column || a.y0 - b.y0);
  const body = $("stage-rows");
  body.innerHTML = "";
  for (const n of accounts) {
    const name = labelOverrides.get(n.name)?.text || n.name;
    body.appendChild(summaryRow(name, n.inValue, n.outValue));
  }
  const foot = $("stage-total");
  foot.innerHTML = "";
  foot.appendChild(summaryRow("Total", totalIn, totalOut));

  $("stage-summary").hidden = false;
  // Keep the right-hand column for the summary.
  document.querySelector(".top-bar")!.classList.remove("no-data");
  applySummaryCollapse();
  updateImbalanceMessage(accounts, totalIn - totalOut);
}

/* ---------- "Value is unaccounted for." alert ---------- */

/** The chart just drawn, so the message can highlight gaps on it. */
let chartHandle: ChartHandle | null = null;
/** Which set of differences was dismissed; the message returns if they change. */
let dismissedImbalance: string | null = null;
let currentImbalance = "";
/** The pointer (or keyboard focus) is on the message, so the chart shows the gaps. */
let imbalanceHover = false;
/** Boxes whose in and out differ (by their name in the data) → in minus out. */
let dataGaps = new Map<string, number>();
/** The gaps the Data tab's flows table was last drawn with. */
let tableImbalance = "";

function updateImbalanceMessage(accounts: SankeyNode[], totalDiff: number) {
  const nearZero = (d: number, scale: number) => Math.abs(d) < Math.max(scale, 1) * 1e-9;
  const off = accounts
    .filter((n) => !nearZero(n.inValue - n.outValue, Math.max(n.inValue, n.outValue)))
    .map((n) => `${n.name}:${n.inValue - n.outValue}`);
  if (!nearZero(totalDiff, Math.abs(totalDiff))) off.push(`total:${totalDiff}`);
  currentImbalance = off.join("|");

  // The Data tab marks the same boxes in amber (rebuilt only when the gaps change).
  dataGaps = new Map(
    accounts
      .filter((n) => !nearZero(n.inValue - n.outValue, Math.max(n.inValue, n.outValue)))
      .map((n) => [n.name, n.inValue - n.outValue])
  );
  if (current && currentImbalance !== tableImbalance) fillTable();

  const anyGap = currentImbalance !== "";
  // The ⓘ stays while anything is unaccounted for; the amber alert until dismissed.
  showImbalanceMessage(anyGap, anyGap && currentImbalance !== dismissedImbalance);
  if (!anyGap) imbalanceHover = false;
  // A redraw replaced the chart; mark the gaps on it again.
  syncImbalanceHighlight();
}

/** The amber alert is showing (not dismissed). */
let imbalanceAlertShown = false;

/**
 * The gaps are marked on the chart for as long as the alert is showing, and
 * after it's dismissed, while the pointer (or focus) is on the ⓘ.
 */
function syncImbalanceHighlight() {
  chartHandle?.highlightImbalance(imbalanceAlertShown || imbalanceHover);
}

/**
 * Shows or hides the ⓘ icon and the amber alert without changing the layout:
 * their space is always kept, so the tabs and chart don't move.
 */
function showImbalanceMessage(icon: boolean, alert = icon) {
  const row = $("imbalance-row");
  row.classList.toggle("is-invisible", !icon);
  row.setAttribute("aria-hidden", String(!icon));
  const msg = $("imbalance-msg");
  msg.classList.toggle("is-invisible", !alert);
  msg.setAttribute("aria-hidden", String(!alert));
  imbalanceAlertShown = icon && alert;
  // Once the alert is dismissed, the ⓘ is the way to see the gaps.
  const info = row.querySelector<HTMLElement>(".info-icon")!;
  const tip = alert
    ? "Some boxes have a difference: money comes in that isn't passed on, or more goes out than comes in. They're marked in amber on the chart until you dismiss this message."
    : "Value is unaccounted for. Hover here to see where on the chart.";
  info.title = tip;
  info.setAttribute("aria-label", tip);
}

/**
 * The chart's size in px when you've chosen one (the size menu); null = automatic (fit the pane or window; height from the data).
 * Used for the pane, the pop-out window and the inserted picture alike.
 */
let chartSize: { width: number | null; height: number | null } = { width: null, height: null };
/** Keeps the preview's height bar on its bottom edge after a redraw or zoom. */
let placeResizeHandle: (() => void) | null = null;

/** The size menu from the chart's dimensions icon. */
let sizeMenu: SizeMenu | null = null;
/** The chart's dotted right and bottom edges, which show and change its size. */
let chartEdges: ChartEdges | null = null;
/** Gap (px) right of and below the pane's chart, so its dotted edges stand clear
 *  of the preview's border (--chart-edge-gap in taskpane.css). */
const CHART_EDGE_GAP = 10;

function setupSizeControls() {
  sizeMenu = attachSizeMenu($<HTMLButtonElement>("size-btn"), (size) => {
    chartSize = size;
    draw();
  });

  // The dotted right and bottom edges: hover for the width or height, click to change it.
  chartEdges = attachChartEdges(
    document.getElementById("sankey-svg") as unknown as SVGSVGElement,
    (axis, value) => {
      chartSize = { ...chartSize, [axis]: value };
      draw();
    }
  );

  // The bar on the preview's bottom edge changes the preview's height (not the chart's).
  placeResizeHandle = attachPreviewResizer(
    document.querySelector<HTMLElement>(".chart-frame")!,
    $("chart")
  );
}

/** The size menu shows a chosen size, or the current automatic size as a hint. */
function syncSizeControls() {
  const svg = document.getElementById("sankey-svg") as unknown as SVGSVGElement;
  const drawn = {
    width: Number(svg.getAttribute("width")) || 0,
    height: Number(svg.getAttribute("height")) || 0,
  };
  sizeMenu?.sync(chartSize, drawn);
  chartEdges?.sync(chartSize, drawn);
}

/**
 * How magnified the pane's chart is (the + / − buttons in the toolbar above it), on top
 * of `paneFit`: at 100% the whole chart fits the preview without scrolling.
 */
let paneZoom = 1;
/** Scale (≤ 1) that fits the whole chart into the preview's width. */
let paneFit = 1;

function setupZoom() {
  const zin = $<HTMLButtonElement>("zoom-in-btn");
  const zout = $<HTMLButtonElement>("zoom-out-btn");
  const level = document.getElementById("zoom-level");
  const svg = document.getElementById("sankey-svg") as unknown as SVGSVGElement;
  const zoom = (dir: 1 | -1) => {
    // Keep the middle of what you're looking at in the middle.
    zoomKeepingCentre($("chart"), () => {
      paneZoom = stepZoom(paneZoom, dir);
      applyZoom(svg, paneZoom * paneFit);
    });
    placeResizeHandle?.();
    updateZoomButtons(zin, zout, paneZoom, level);
  };
  zin.onclick = () => zoom(1);
  zout.onclick = () => zoom(-1);
  updateZoomButtons(zin, zout, paneZoom, level);
}

/** Pane widths over which the top area shrinks, and how small it gets. */
const TOP_FULL_SIZE_WIDTH = 800;
const TOP_SMALLEST_WIDTH = 500;
const TOP_SMALLEST_SCALE = 0.75;

/**
 * Shrinks the top area (title, load buttons, summary) gradually as the pane
 * narrows from 800px to 500px, so its two columns keep fitting side by side.
 */
function setupTopScale() {
  const update = () => {
    const w = document.documentElement.clientWidth;
    const t = Math.min(
      1,
      Math.max(0, (w - TOP_SMALLEST_WIDTH) / (TOP_FULL_SIZE_WIDTH - TOP_SMALLEST_WIDTH))
    );
    const scale = TOP_SMALLEST_SCALE + (1 - TOP_SMALLEST_SCALE) * t;
    document.documentElement.style.setProperty("--top-scale", scale.toFixed(3));
  };
  window.addEventListener("resize", update);
  update();
}

/**
 * "Clear loaded data" asks first, in a small box under it. With unsaved
 * changes it suggests saving: Cancel, Save and close, Close without saving.
 * (Office add-ins can't use the browser's confirm() dialog.) Escape, Cancel,
 * or clicking anywhere else closes it without clearing.
 */
function setupClearConfirm() {
  const btn = $<HTMLButtonElement>("clear-btn");
  const box = $("clear-confirm");
  const save = $<HTMLButtonElement>("clear-save");
  const ok = $<HTMLButtonElement>("clear-ok");
  const show = (open: boolean) => {
    if (open) {
      const dirty = hasUnsavedChanges();
      const saved = diagramById(openDiagramId);
      $("clear-confirm-text").textContent = dirty
        ? `Save first! This chart's colours, groups, labels and other settings ${
            saved ? "have unsaved changes, which" : "aren't saved, so they"
          } will be lost when the data is cleared. Your workbook's data isn't changed.`
        : `"${saved?.name}" is saved in this workbook, so you can open it again from the Memory tab. Your workbook's data isn't changed.`;
      save.hidden = !dirty;
      ok.textContent = dirty ? "Close without saving" : "Close";
    }
    box.hidden = !open;
    btn.setAttribute("aria-expanded", String(open));
    // The safe choice first: saving when there's something to save.
    if (open) (save.hidden ? $<HTMLButtonElement>("clear-cancel") : save).focus();
  };
  btn.onclick = () => show(box.hidden);
  $("clear-cancel").onclick = () => {
    show(false);
    btn.focus();
  };
  save.onclick = () => {
    show(false);
    run(async () => {
      await saveCurrent();
      const name = diagramById(openDiagramId)?.name;
      await clearData();
      if (name) setStatus(`Saved "${name}". Open it again any time from the Memory tab.`);
    });
  };
  ok.onclick = () => {
    show(false);
    run(() => clearData());
  };
  box.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      show(false);
      btn.focus();
    }
  });
  document.addEventListener("mousedown", (e) => {
    if (!box.hidden && !box.parentElement!.contains(e.target as HTMLElement)) show(false);
  });
}

function setupSummaryToggle() {
  $("stage-toggle").onclick = () => {
    summaryCollapsed = !summaryCollapsed;
    applySummaryCollapse();
    updateSaveState();
  };
}

function setupImbalanceMessage() {
  const msg = $("imbalance-msg");
  const text = msg.querySelector<HTMLElement>(".imbalance-text")!;
  const info = $("imbalance-row").querySelector<HTMLElement>(".info-icon")!;
  const set = (on: boolean) => {
    imbalanceHover = on;
    syncImbalanceHighlight();
  };
  // Hovering (or focusing) the alert or the ⓘ marks the gaps on the chart.
  for (const el of [msg, info]) {
    el.addEventListener("mouseenter", () => set(true));
    el.addEventListener("mouseleave", () => set(false));
  }
  for (const el of [text, info]) {
    el.addEventListener("focus", () => set(true));
    el.addEventListener("blur", () => set(false));
  }
  $("imbalance-dismiss").onclick = () => {
    // The alert stays hidden until the differences change (e.g. after editing
    // the data); the ⓘ stays so the gaps can still be shown.
    dismissedImbalance = currentImbalance;
    showImbalanceMessage(true, false);
    set(false);
  };
}

/** Chevron icons for the summary collapse bar (drawn in the bar's text colour). */
const chevron = (d: string) =>
  `<svg viewBox="0 0 16 16" aria-hidden="true" focusable="false"><path d="${d}" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
const CHEVRON_UP = chevron("M4 10 L8 6 L12 10");
const CHEVRON_DOWN = chevron("M4 6 L8 10 L12 6");

/** More account rows than this, and the summary table can be collapsed. */
const SUMMARY_COLLAPSE_AFTER = 4;
/** The summary table is collapsed (the default); it stays that way through redraws until you toggle it. */
let summaryCollapsed = true;

/**
 * Shows the collapse button when there are more than 4 account rows, and in
 * the collapsed state hides the balanced account rows (the Total row and any
 * row with a difference stay visible).
 */
function applySummaryCollapse() {
  const rows = Array.from($("stage-rows").querySelectorAll<HTMLTableRowElement>("tr"));
  const toggle = $<HTMLButtonElement>("stage-toggle");
  const canCollapse = rows.length > SUMMARY_COLLAPSE_AFTER;
  const collapsed = canCollapse && summaryCollapsed;
  toggle.hidden = !canCollapse;
  for (const tr of rows) tr.hidden = collapsed && tr.dataset.balanced === "true";
  if (canCollapse) {
    const hidden = rows.filter((tr) => tr.hidden).length;
    // A chevron: down = expand (show all rows), up = collapse. The words are in the tooltip.
    toggle.innerHTML = collapsed ? CHEVRON_DOWN : CHEVRON_UP;
    toggle.title = collapsed
      ? `Show all ${rows.length} rows (${hidden} balanced row${hidden === 1 ? "" : "s"} hidden)`
      : "Collapse: show only the total and any rows with a difference";
    toggle.setAttribute("aria-label", toggle.title);
    toggle.setAttribute("aria-expanded", String(!collapsed));
  }
  // The alert row (beside the tabs) spans the table and its collapse bar, so it
  // lines up with the table on the left and the bar on the right.
  const width = document.querySelector<HTMLElement>(".summary-with-toggle")!.offsetWidth;
  $("imbalance-row").style.minWidth = width ? `${width}px` : "";
}

/** A difference with its sign: "+£150", "−£50" or "£0". */
function signed(diff: number): string {
  return (diff > 0 ? "+" : diff < 0 ? "−" : "") + formatNumber(Math.abs(diff), labelSettings);
}

/** A summary row: name, in, out, and the difference — green when balanced (0), red otherwise. */
function summaryRow(name: string, inValue: number, outValue: number): HTMLTableRowElement {
  let diff = inValue - outValue;
  // Ignore floating-point dust from adding many decimals.
  if (Math.abs(diff) < Math.max(inValue, outValue, 1) * 1e-9) diff = 0;

  const tr = document.createElement("tr");
  // Balanced rows can be hidden when the summary is collapsed.
  tr.dataset.balanced = String(diff === 0);
  const cell = (text: string, className = "") => {
    const td = document.createElement("td");
    td.textContent = text;
    if (className) td.className = className;
    tr.appendChild(td);
  };
  cell(name);
  cell(formatNumber(inValue, labelSettings), "num");
  cell(formatNumber(outValue, labelSettings), "num");
  cell(signed(diff), "num " + (diff === 0 ? "diff-balanced" : "diff-unbalanced"));

  const amount = formatNumber(Math.abs(diff), labelSettings);
  tr.title =
    diff === 0
      ? `${name}: everything that comes in goes out`
      : diff > 0
        ? `${name}: ${amount} more comes in than goes out`
        : `${name}: ${amount} more goes out than comes in`;
  return tr;
}

/* ---------- Customize: colors & positions ---------- */

function selectNode(name: string, fromChart: boolean) {
  selectedNode = name;
  // The box editor lives in the Customize section; open it if it's collapsed.
  ($("customize-section") as HTMLDetailsElement).open = true;
  draw({ focusSelected: fromChart });
  $("color-editor").scrollIntoView({ block: "nearest" });
}

function setNodeColor(name: string, choice: ColorChoice | null) {
  if (choice) nodeColors.set(name, choice);
  else nodeColors.delete(name);
  draw();
}

const hasLabelOffset = (o: NodeLabelOverride | undefined) => !!(o && (o.dx || o.dy));
const hasLabelEdit = (o: NodeLabelOverride | undefined) => !!(o && (o.text || o.placement));

function updateResetButtons() {
  const anyOffset = Array.from(labelOverrides.values()).some(hasLabelOffset);
  $<HTMLButtonElement>("reset-positions-btn").disabled = nodePositions.size === 0 && !anyOffset;
  $<HTMLButtonElement>("reset-colors-btn").disabled = nodeColors.size === 0;
  $<HTMLButtonElement>("reset-labels-btn").disabled =
    labelOverrides.size === 0 && JSON.stringify(labelSettings) === JSON.stringify(DEFAULT_LABELS);
  if (selectedNode) {
    $<HTMLButtonElement>("reset-node-position").disabled =
      !nodePositions.has(selectedNode) && !hasLabelOffset(labelOverrides.get(selectedNode));
    $<HTMLButtonElement>("reset-node-label").disabled = !labelOverrides.has(selectedNode);
  }
}

/** Shows the "moved" tag on a list row without rebuilding the chart mid-interaction. */
function markMoved(name: string) {
  const row = document.querySelector<HTMLElement>(
    `#node-list [data-node="${CSS.escape(name)}"] .tag`
  );
  if (row) row.textContent = tagText(name);
}

function tagText(name: string) {
  const tags: string[] = [];
  const o = labelOverrides.get(name);
  const assigned = manualGroup(name);
  if (assigned) tags.push(`in ${assigned.name}`);
  else if (nodeColors.has(name)) tags.push("custom color");
  if (hasLabelEdit(o)) tags.push("custom label");
  if (nodePositions.has(name) || hasLabelOffset(o)) tags.push("moved");
  return tags.join(" · ");
}

/** Converts any CSS hex color to #rrggbb, which <input type="color"> requires. */
const toHex6 = (c: string) =>
  /^#[0-9a-f]{3}$/i.test(c) ? "#" + [1, 2, 3].map((i) => c[i] + c[i]).join("") : c.toLowerCase();

type GroupMap = Map<SankeyNode, NodeGroup & { id: string }>;

/** A box assigned to a group by hand (via the box editor), if any. */
function manualGroup(name: string): ColorGroup | undefined {
  const choice = nodeColors.get(name);
  if (!choice?.startsWith(GROUP_PREFIX)) return undefined;
  return groups.find((g) => g.id === choice.slice(GROUP_PREFIX.length));
}

/** Does a box name contain a keyword (as a whole word, plural allowed)? */
function keywordMatch(name: string, keyword: string): boolean {
  const k = keyword.trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return !!k && new RegExp(`(^|\\W)${k}s?(\\W|$)`, "i").test(name);
}

/** The group a box falls into automatically: your keywords first, then the built-in rules. */
function autoGroup(
  n: SankeyNode,
  roles: Map<SankeyNode, NodeRole>,
  memo = new Map<SankeyNode, ColorGroup | undefined>()
): ColorGroup | undefined {
  if (memo.has(n)) return memo.get(n);
  const found = findAutoGroup(n, roles, memo);
  memo.set(n, found);
  return found;
}

function findAutoGroup(
  n: SankeyNode,
  roles: Map<SankeyNode, NodeRole>,
  memo: Map<SankeyNode, ColorGroup | undefined>
): ColorGroup | undefined {
  // Keywords are checked from the last group to the first: groups you add (at
  // the end) win, and outcome groups ("Customer returns" → Returns, "Lost
  // leads" → Lost) beat the broader groups before them.
  for (let i = groups.length - 1; i >= 0; i--) {
    if (groups[i].keywords.some((k) => keywordMatch(n.name, k))) return groups[i];
  }
  // Personal / general themes: the name-guessing rules in finance.ts.
  const r = roles.get(n);
  const byRole = groups.find((g) => g.role && g.role === r);
  if (byRole) return byRole;
  // Professional themes: a box whose parents are all in the same outcome group
  // joins it (e.g. "Other" under "Operating expenses" → Costs), as in Default.
  if (n.targetLinks.length) {
    const parents = new Set(n.targetLinks.map((l) => autoGroup(l.source, roles, memo)));
    const only = parents.size === 1 ? [...parents][0] : undefined;
    if (only && only.position !== "start" && only.position !== "middle") return only;
  }
  // Otherwise by position in the chart.
  const pos = !n.targetLinks.length ? "start" : !n.sourceLinks.length ? "end" : "middle";
  return groups.find((g) => g.position === pos);
}

/**
 * Each box's group (undefined only if group colouring is off). Order of
 * precedence: assigned by hand, then your keywords, then the built-in rules.
 * Boxes whose group was removed are left out (drawn as "Other").
 */
function groupsFor(nodes: SankeyNode[], theme: ChartTheme): GroupMap | undefined {
  if (!isMoneyFlow()) return undefined;
  const roles = computeRoles({ nodes });
  const map: GroupMap = new Map();
  for (const n of nodes) {
    const g = manualGroup(n.name) ?? autoGroup(n, roles);
    if (g) map.set(n, { id: g.id, name: g.name, color: groupColor(g, theme) });
  }
  return map;
}

/** Box colors for the chart: your color picks, minus group assignments (those come via groups). */
const chartColors = () =>
  new Map(Array.from(nodeColors).filter(([, c]) => !c.startsWith(GROUP_PREFIX)));

/**
 * For "Group by colour": each box's final color (including your own picks).
 * The theme's shade doesn't matter here, only which boxes share a color.
 */
function colorGroupKeys(nodes: SankeyNode[]): Map<SankeyNode, string> {
  const theme = chartTheme(false);
  const gm = groupsFor(nodes, theme);
  const colors = chartColors();
  return new Map(nodes.map((n) => [n, nodeColor(n, theme, colors, gm).toLowerCase()]));
}

/** The legend + group editor under the Colours dropdown (Money flow only). */
function renderSchemeHint(theme: ChartTheme, layout: SankeyLayout, gm: GroupMap | undefined) {
  const hint = $("scheme-hint");
  syncColourSelect();
  updateGroupWording();
  hint.hidden = !isMoneyFlow();
  if (hint.hidden) return;

  // Legend: one dot + name per group, plus "Other" if some boxes have no group.
  const legend = $("group-legend");
  legend.innerHTML = "";
  const item = (color: string, text: string) => {
    const el = document.createElement("span");
    el.className = "legend-item";
    const dot = document.createElement("span");
    dot.className = "legend-dot";
    dot.style.background = color;
    el.append(dot, document.createTextNode(text));
    legend.appendChild(el);
  };
  for (const g of groups) item(groupColor(g, theme), g.name);
  if (gm && layout.nodes.some((n) => !gm.has(n) && !chartColors().has(n.name))) {
    item(theme.neutral, "Other");
  }
  renderGroupEditor(theme);
}

/* ---------- Group colors & names (Money flow) ---------- */

/** Display order of the groups. */
const ROLE_ORDER: NodeRole[] = ["income", "account", "spending", "savings", "remainder"];

const DEFAULT_GROUP_NAMES: Record<NodeRole, string> = {
  income: "Income",
  account: "Accounts",
  spending: "Spending",
  savings: "Savings",
  remainder: "Remainder",
};

/** Ready-made color sets. "default" (shown as Cashflow) uses the built-in, theme-aware colors. */
const GROUP_THEMES: { id: string; name: string; colors: Partial<Record<NodeRole, string>> }[] = [
  { id: "default", name: "Default", colors: {} },
  {
    id: "traffic",
    name: "Traffic light",
    colors: {
      income: "#2e9e44",
      account: "#7a7a7a",
      spending: "#d64541",
      savings: "#2f6fb7",
      remainder: "#e0a100",
    },
  },
  {
    id: "ocean",
    name: "Ocean",
    colors: {
      income: "#0f7c8c",
      account: "#5b6770",
      spending: "#e4572e",
      savings: "#1d4e89",
      remainder: "#f2a541",
    },
  },
  {
    id: "pastel",
    name: "Pastel",
    colors: {
      income: "#6cc28a",
      account: "#a7a7a7",
      spending: "#ef7f8f",
      savings: "#79a7d3",
      remainder: "#f2c46d",
    },
  },
  {
    id: "mono",
    name: "Monochrome",
    colors: {
      income: "#1f1f1f",
      account: "#8a8a8a",
      spending: "#4d4d4d",
      savings: "#666666",
      remainder: "#b3b3b3",
    },
  },
];

/* ColorGroup and GroupSet are defined in memory.ts (they're saved with each diagram). */
let groupSet: GroupSet = "money";
const industryOf = (): IndustryId | null =>
  groupSet === "money" || groupSet === "generic" ? null : groupSet;

/** A built-in group's default name: "Income"… in Money flow, "Group 1"… in the general themes. */
function defaultName(g: ColorGroup): string {
  if (!g.role) return "Group";
  return groupSet === "generic"
    ? `Group ${ROLE_ORDER.indexOf(g.role) + 1}`
    : DEFAULT_GROUP_NAMES[g.role];
}

function industryGroups(id: IndustryId): ColorGroup[] {
  return industryPreset(id)!.groups.map((g, i) => ({
    id: `${id}-${i}`,
    name: g.name,
    color: null,
    keywords: [...g.keywords],
    tone: g.tone,
    position: g.position,
  }));
}

/** The groups a theme starts with (what "Reset groups" goes back to). */
function defaultGroups(): ColorGroup[] {
  const ind = industryOf();
  return ind ? industryGroups(ind) : builtinGroups(groupSet === "generic");
}

function toneColor(tone: Tone, theme: ChartTheme): string {
  return tone === "blue" ? theme.series[0] : theme.roles[tone];
}

/** Switches to another set of groups (Money flow, general, or an industry). */
function switchGroupSet(next: GroupSet) {
  if (next === groupSet) return;
  groupSet = next;
  groups = defaultGroups();
  // Boxes put in a group by hand go back to automatic if that group is gone.
  for (const [name, c] of Array.from(nodeColors)) {
    if (c.startsWith(GROUP_PREFIX) && !manualGroup(name)) nodeColors.delete(name);
  }
}

/** The five built-in groups: named Income… (Money flow) or Group 1… (general themes). */
const builtinGroups = (numbered = false): ColorGroup[] =>
  ROLE_ORDER.map((r, i) => ({
    id: r,
    name: numbered ? `Group ${i + 1}` : DEFAULT_GROUP_NAMES[r],
    color: null,
    keywords: [],
    role: r,
  }));

let groups: ColorGroup[] = builtinGroups();
let addedGroupCount = 0;

function groupColor(g: ColorGroup, theme: ChartTheme): string {
  if (g.color) return g.color;
  if (g.role) return theme.roles[g.role];
  if (g.tone) return toneColor(g.tone, theme);
  return theme.neutral;
}

/** Color choice that assigns a box to a group by hand: "group:<id>". */
const GROUP_PREFIX = "group:";

const THEME_PREFIX = "theme:";
const INDUSTRY_PREFIX = "industry:";

/** The Generic theme last picked; the dropdown keeps showing it after you edit a colour. */
let genericTheme = "traffic";

/** Boxes are always coloured by group (the one-colour-per-box Palette option was removed). */
const isMoneyFlow = () => true;

/** Applies a theme's colors to the built-in groups (groups you added keep theirs). */
function applyThemeColors(colors: Partial<Record<NodeRole, string>>) {
  for (const g of groups) if (g.role) g.color = colors[g.role] ?? null;
}

/**
 * The single Colour theme dropdown:
 *   Personal     — Cashflow (Income, Accounts, …)
 *   Professional — Finance, Supply chain, Manufacturing, HR, Sales
 *   Generic      — Traffic light, Ocean, Pastel, Monochrome (Group 1…Group 5)
 */
function setupGroupControls() {
  const select = $<HTMLSelectElement>("color-scheme");
  const option = (value: string, text: string) =>
    Object.assign(document.createElement("option"), { value, textContent: text });

  const money = Object.assign(document.createElement("optgroup"), { label: "Personal" });
  money.appendChild(option(THEME_PREFIX + "default", "Cashflow"));

  const other = Object.assign(document.createElement("optgroup"), { label: "Generic" });
  for (const t of GROUP_THEMES) {
    if (t.id !== "default") other.appendChild(option(THEME_PREFIX + t.id, t.name));
  }

  const industries = Object.assign(document.createElement("optgroup"), { label: "Professional" });
  for (const p of INDUSTRY_PRESETS) {
    industries.appendChild(
      Object.assign(document.createElement("option"), {
        value: INDUSTRY_PREFIX + p.id,
        textContent: p.name,
      })
    );
  }
  select.append(money, industries, other);
  select.value = THEME_PREFIX + "default";
  updateColoursNote();

  select.onchange = () => {
    const v = select.value;
    if (v.startsWith(INDUSTRY_PREFIX)) {
      switchGroupSet(v.slice(INDUSTRY_PREFIX.length) as IndustryId);
    } else if (v === THEME_PREFIX + "default") {
      switchGroupSet("money");
    } else if (v.startsWith(THEME_PREFIX)) {
      // A Generic theme: Group 1…Group 5. Moving between Generic themes only
      // recolours, so your group names, keywords and added groups stay.
      switchGroupSet("generic");
      genericTheme = v.slice(THEME_PREFIX.length);
      const t = GROUP_THEMES.find((x) => x.id === genericTheme);
      if (t) applyThemeColors(t.colors);
    }
    // Group colours read best with flows colored by where they go.
    $<HTMLSelectElement>("link-color").value = "target";
    draw();
  };
  $("reset-groups-btn").onclick = () => {
    groups = defaultGroups();
    // In a Generic theme, reset to that theme's colours.
    if (groupSet === "generic") {
      const t = GROUP_THEMES.find((x) => x.id === genericTheme);
      if (t) applyThemeColors(t.colors);
    }
    // Boxes assigned to a group you'd added go back to their automatic group.
    for (const [name, c] of Array.from(nodeColors)) {
      if (c.startsWith(GROUP_PREFIX) && !manualGroup(name)) nodeColors.delete(name);
    }
    draw();
  };
  $("add-group-btn").onclick = () => {
    addedGroupCount++;
    const theme = chartTheme(false);
    const used = new Set(groups.map((g) => groupColor(g, theme).toLowerCase()));
    const color = theme.series.find((c) => !used.has(c.toLowerCase())) ?? theme.neutral;
    const id = `custom-${Date.now().toString(36)}-${addedGroupCount}`;
    // Numbered after the groups already there (Group 6 after Group 1…Group 5).
    groups = [...groups, { id, name: `Group ${groups.length + 1}`, color, keywords: [] }];
    ($("group-editor") as HTMLDetailsElement).open = true;
    draw();
    // Start editing the new group's name.
    $("group-rows").querySelector<HTMLInputElement>(`[data-id="${id}"] input[type=text]`)?.select();
  };
}

function removeGroup(id: string) {
  groups = groups.filter((g) => g.id !== id);
  // Boxes assigned to it by hand go back to their automatic group.
  for (const [name, c] of Array.from(nodeColors)) {
    if (c === GROUP_PREFIX + id) nodeColors.delete(name);
  }
  draw();
}

/** Shows the theme in use in the dropdown (edited colours keep their theme's name). */
function syncColourSelect() {
  const select = $<HTMLSelectElement>("color-scheme");
  const ind = industryOf();
  select.value = ind
    ? INDUSTRY_PREFIX + ind
    : groupSet === "generic"
      ? THEME_PREFIX + genericTheme
      : THEME_PREFIX + "default";
  updateColoursNote();
}

/** The hint under the legend, and the summary table's first heading, for the chosen theme. */
function updateGroupWording() {
  const ind = industryOf();
  const preset = ind ? industryPreset(ind) : undefined;
  $("stage-col-head").textContent = preset
    ? preset.stageWord
    : groupSet === "generic"
      ? "Stage"
      : "Account";
  const hint = $("scheme-hint-text");
  if (groupSet === "generic") {
    // Same rules as Money flow, described with the (numbered) group names.
    const name = (r: NodeRole) => groups.find((g) => g.role === r)?.name;
    const bits = [
      name("income") && `starting boxes go in ${name("income")}`,
      name("account") && `boxes in between in ${name("account")}`,
      name("spending") && `end boxes in ${name("spending")}`,
    ].filter(Boolean);
    const extra = [
      name("savings") && `savings → ${name("savings")}`,
      name("remainder") && `remainder → ${name("remainder")}`,
    ].filter(Boolean);
    hint.textContent =
      (bits.length ? bits.join(", ").replace(/^./, (c) => c.toUpperCase()) + "." : "") +
      (extra.length ? ` Some names are recognised (e.g. ${extra.join(", ")}).` : "") +
      " Click a box to change its colour.";
    return;
  }
  if (!preset) {
    hint.innerHTML =
      "Starting boxes (nothing flowing in) go in the first group; others are guessed from names " +
      "(e.g. <i>bills</i>, <i>rent</i>, <i>savings</i>, <i>fund</i>, <i>remainder</i>, <i>account</i>). " +
      "Click a box to change its color.";
    return;
  }
  // Examples from the current groups, and where unmatched boxes go.
  const examples = groups
    .map((g) => g.keywords[0])
    .filter(Boolean)
    .slice(0, 6);
  const at = (pos: "start" | "middle" | "end") => groups.find((g) => g.position === pos)?.name;
  const parts = [
    at("start") && `starting boxes in ${at("start")}`,
    at("middle") && `boxes in between in ${at("middle")}`,
    at("end") && `end boxes in ${at("end")}`,
  ].filter(Boolean);
  hint.textContent =
    `Boxes are grouped by words in their names (e.g. ${examples.join(", ")})` +
    (parts.length ? `; otherwise ${parts.join(", ")}.` : ".") +
    " Click a box to change its colour.";
}

/** The Style section heading shows the current theme, so it's visible while collapsed. */
function updateColoursNote() {
  const select = $<HTMLSelectElement>("color-scheme");
  const text = select.selectedOptions[0]?.textContent ?? "";
  $("colours-note").textContent = text;
}

/** One card per group: color, name, keywords, and Reset (built-in) or Remove. */
function renderGroupEditor(theme: ChartTheme) {
  $<HTMLButtonElement>("reset-groups-btn").disabled =
    JSON.stringify(groups) === JSON.stringify(builtinGroups());

  const rows = $("group-rows");
  // Rebuild the cards only when groups are added or removed, so the input
  // you're typing in keeps its focus and cursor.
  // Includes the group set, so switching Default ↔ a general theme (same IDs,
  // different names and placeholders) rebuilds the cards too.
  const sig = groupSet + ":" + groups.map((g) => g.id).join("|");
  if (rows.dataset.sig !== sig) {
    rows.dataset.sig = sig;
    rows.innerHTML = "";
    for (const g of groups) rows.appendChild(groupCard(g));
  }
  for (const g of groups) {
    const card = rows.querySelector<HTMLElement>(`[data-id="${g.id}"]`)!;
    const color = card.querySelector<HTMLInputElement>("input[type=color]")!;
    const name = card.querySelector<HTMLInputElement>("input.group-name")!;
    const words = card.querySelector<HTMLInputElement>("input.group-keywords")!;
    if (document.activeElement !== color) color.value = toHex6(groupColor(g, theme));
    if (document.activeElement !== name) {
      name.value = g.role && g.name === defaultName(g) ? "" : g.name;
    }
    if (document.activeElement !== words) words.value = g.keywords.join(", ");
    const reset = card.querySelector<HTMLButtonElement>(".group-reset");
    if (reset && g.role) {
      reset.disabled = g.color === null && g.name === defaultName(g) && g.keywords.length === 0;
    }
  }
}

function groupCard(g: ColorGroup): HTMLElement {
  const label = g.role ? defaultName(g) : "custom";
  const card = document.createElement("div");
  card.className = "group-card";
  card.dataset.id = g.id;
  // Handlers look the group up when used: "Reset groups" or switching themes
  // replaces the group objects, and an edit must go to the current one.
  const id = g.id;
  const cur = () => groups.find((x) => x.id === id);

  const color = Object.assign(document.createElement("input"), { type: "color" });
  color.setAttribute("aria-label", "Group colour");
  color.addEventListener("input", () => {
    const c = cur();
    if (!c) return;
    c.color = color.value;
    draw();
  });

  const name = Object.assign(document.createElement("input"), {
    type: "text",
    className: "text-input group-name",
    maxLength: 30,
    placeholder: g.role ? defaultName(g) : "Group name",
  });
  name.setAttribute("aria-label", g.role ? `Name for the ${label} group` : "Group name");
  name.addEventListener("input", () => {
    const c = cur();
    if (!c) return;
    c.name = name.value.trim() || (c.role ? defaultName(c) : "Group");
    draw();
  });

  const actions = document.createElement("span");
  actions.className = "group-actions";
  if (g.role) {
    const reset = Object.assign(document.createElement("button"), {
      type: "button",
      className: "link-btn group-reset",
      textContent: "Reset",
      title: "Default colour, name and no extra keywords",
    });
    reset.onclick = () => {
      const c = cur();
      if (!c) return;
      c.color = null;
      c.name = defaultName(c);
      c.keywords = [];
      draw();
    };
    actions.appendChild(reset);
  }
  const remove = Object.assign(document.createElement("button"), {
    type: "button",
    className: "icon-link group-remove",
    textContent: "✕",
    title: "Remove this group",
  });
  remove.setAttribute("aria-label", `Remove the ${g.name} group`);
  remove.onclick = () => removeGroup(id);
  actions.appendChild(remove);

  const words = Object.assign(document.createElement("input"), {
    type: "text",
    className: "text-input group-keywords",
    placeholder: g.role ? "Extra keywords, e.g. vet, pet" : "Keywords, e.g. dog, vet, pet",
  });
  words.setAttribute("aria-label", `Keywords for the ${g.name} group`);
  words.addEventListener("input", () => {
    const c = cur();
    if (!c) return;
    c.keywords = words.value
      .split(",")
      .map((k) => k.trim())
      .filter(Boolean);
    draw();
  });

  card.append(color, name, actions, words);
  return card;
}

function renderCustomize(layout: SankeyLayout, theme: ChartTheme, gm: GroupMap | undefined) {
  renderLabelPreview(layout, theme);

  // Node list: one row per node, in order of first appearance (the palette order).
  const list = $("node-list");
  list.innerHTML = "";
  for (const n of layout.nodes) {
    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.dataset.node = n.name;
    btn.setAttribute("aria-current", String(n.name === selectedNode));
    const dot = document.createElement("span");
    dot.className = "dot";
    dot.style.background = nodeColor(n, theme, chartColors(), gm);
    const label = document.createElement("span");
    label.textContent = n.name;
    const tag = document.createElement("span");
    tag.className = "tag";
    tag.textContent = tagText(n.name);
    btn.append(dot, label, tag);
    btn.onclick = () => selectNode(n.name, false);
    li.appendChild(btn);
    list.appendChild(li);
  }

  // Color editor for the selected node.
  const editor = $("color-editor");
  const node = layout.nodes.find((n) => n.name === selectedNode);
  editor.hidden = !node;
  if (node) renderColorEditor(node, layout, theme, gm);
  updateResetButtons();
}

function renderColorEditor(
  node: SankeyNode,
  layout: SankeyLayout,
  theme: ChartTheme,
  gm: GroupMap | undefined
) {
  const color = nodeColor(node, theme, chartColors(), gm);
  $("editor-name").textContent = node.name;
  $("editor-swatch").style.background = color;
  $<HTMLInputElement>("custom-color").value = toHex6(color);
  $<HTMLButtonElement>("reset-node-color").disabled = !nodeColors.has(node.name);

  // Label fields (don't overwrite the text box while the user is typing in it).
  const o = labelOverrides.get(node.name);
  const text = $<HTMLInputElement>("node-label-text");
  text.placeholder = node.name;
  if (document.activeElement !== text) text.value = o?.text ?? "";
  $<HTMLSelectElement>("node-label-placement").value = o?.placement ?? "";

  // Swatches are stored as slots/groups, not hex, so they follow the theme.
  const choice = nodeColors.get(node.name);
  const manual = manualGroup(node.name);
  const auto = gm ? autoGroup(node, computeRoles(layout)) : undefined;
  const defaultColor = defaultNodeColor(node, theme, gm && withoutManual(gm, node, auto, theme));

  // Group colour (Money flow only): picking a group puts the box in it — its
  // color, its name in the tooltip, and its place in "Grouped by colour".
  $("group-colour-part").hidden = !gm;
  if (gm) {
    const current = manual ?? auto; // the group the box is in
    renderSwatches(
      $("group-swatches"),
      groups.map((g) => ({
        choice: GROUP_PREFIX + g.id,
        color: groupColor(g, theme),
        name: g.name,
        checked: !!current && g.id === current.id && (!choice || !!manual),
        // Picking the group it would be in anyway clears the override.
        apply: g.id === auto?.id ? null : GROUP_PREFIX + g.id,
      }))
    );
    const autoName = auto ? auto.name : "no group (Other)";
    $("group-note").textContent = manual
      ? `Put in ${manual.name} by hand. Automatically it would be in ${autoName}.`
      : choice
        ? `In ${autoName}, but showing an individual colour (below).`
        : `Automatically in ${autoName}${auto ? " (from its name or position)" : ""}.`;
  }

  // Individual colour: just this box, whatever its group.
  const individual = [
    ...theme.series.map((c, i) => ({
      choice: `slot:${i}`,
      color: c,
      name:
        ["Blue", "Orange", "Aqua", "Yellow", "Magenta", "Green", "Violet", "Red"][i] ??
        `Color ${i + 1}`,
    })),
    { choice: "neutral", color: theme.neutral, name: "Gray" },
  ];
  renderSwatches(
    $("swatches"),
    individual.map((c) => {
      // Without group colouring a box's default color is one of these; picking it clears the override.
      const isDefault = !gm && c.color === defaultColor;
      return {
        ...c,
        checked: choice ? choice === c.choice : isDefault,
        apply: isDefault ? null : c.choice,
      };
    })
  );
}

function renderSwatches(
  box: HTMLElement,
  items: { choice: string; color: string; name: string; checked: boolean; apply: string | null }[]
) {
  box.innerHTML = "";
  for (const c of items) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "swatch";
    b.style.background = c.color;
    b.title = c.name;
    b.setAttribute("role", "radio");
    b.setAttribute("aria-checked", String(c.checked));
    b.setAttribute("aria-label", c.name);
    // The editor always edits the selected box.
    b.onclick = () => {
      if (selectedNode) setNodeColor(selectedNode, c.apply);
    };
    box.appendChild(b);
  }
}

/** The group map as it would be without this box's hand assignment. */
function withoutManual(
  gm: GroupMap,
  node: SankeyNode,
  auto: ColorGroup | undefined,
  theme: ChartTheme
): GroupMap {
  const copy: GroupMap = new Map(gm);
  if (auto) copy.set(node, { id: auto.id, name: auto.name, color: groupColor(auto, theme) });
  else copy.delete(node);
  return copy;
}

function fillTable() {
  $("flows-empty").hidden = true;
  $("flows-wrap").hidden = false;
  // Sample data is for getting started; it goes away once there's real data.
  $("sample-section").hidden = true;
  const table = $("flow-table");
  const [h1, h2, h3] = current!.headers;
  const esc = (s: string) =>
    s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);
  const flows = current!.flows;
  tableImbalance = currentImbalance;

  // Where value is unaccounted for: a box that passes on less than it receives
  // gets a placeholder flow after its last outgoing flow; one that sends out
  // more than it receives, after its last incoming flow.
  const gapAfter = new Map<number, string[]>();
  for (const [name, diff] of dataGaps) {
    const notPassedOn = diff > 0;
    let last = -1;
    flows.forEach((f, i) => {
      if ((notPassedOn ? f.source : f.target) === name) last = i;
    });
    if (last < 0) continue;
    const tip = esc(gapText(name, diff));
    const cells = notPassedOn
      ? `<td class="gap-name">${esc(name)}</td><td>Unaccounted for</td>`
      : `<td>Unaccounted for</td><td class="gap-name">${esc(name)}</td>`;
    const row = `<tr class="gap-row" title="${tip}">${cells}<td class="num">${formatValue(Math.abs(diff))}</td></tr>`;
    gapAfter.set(last, [...(gapAfter.get(last) ?? []), row]);
  }

  const nameCell = (name: string) =>
    dataGaps.has(name)
      ? `<td class="gap-name" title="${esc(gapText(name, dataGaps.get(name)!))}">${esc(name)}</td>`
      : `<td>${esc(name)}</td>`;
  table.innerHTML =
    `<thead><tr><th>${esc(h1)}</th><th>${esc(h2)}</th><th class="num">${esc(h3)}</th></tr></thead><tbody>` +
    flows
      .map(
        (f, i) =>
          `<tr>${nameCell(f.source)}${nameCell(f.target)}<td class="num">${formatValue(f.value)}</td></tr>` +
          (gapAfter.get(i) ?? []).join("")
      )
      .join("") +
    "</tbody>";

  // The note above the table lists each gap.
  const note = $("data-gaps");
  note.hidden = dataGaps.size === 0;
  note.innerHTML = dataGaps.size
    ? `<strong><span aria-hidden="true">⚠</span> Value is unaccounted for.</strong> Marked in amber below:<ul>` +
      [...dataGaps].map(([name, diff]) => `<li>${esc(gapText(name, diff))}</li>`).join("") +
      "</ul>"
    : "";
}

/** "Account 4: 300 comes in but isn't passed on." or "…: 300 more goes out than comes in." */
function gapText(name: string, diff: number) {
  const amount = formatValue(Math.abs(diff));
  return diff > 0
    ? `${name}: ${amount} comes in but isn't passed on.`
    : `${name}: ${amount} more goes out than comes in.`;
}

/* ---------- Writing to Excel ---------- */

/** Rasterizes an SVG element to a Base64 PNG (no data: prefix), as Excel's addImage expects. */
function svgToPngBase64(
  svg: SVGSVGElement,
  width: number,
  height: number,
  scale: number
): Promise<string> {
  const xml = new XMLSerializer().serializeToString(svg);
  const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(xml);
  return new Promise((resolve, reject) => {
    const fail = () => reject(new Error("Couldn't create the chart image."));
    const img = new Image();
    img.onload = () => {
      // An error thrown in here wouldn't reject the promise by itself (e.g. a
      // canvas the browser treats as "tainted"), leaving Insert stuck for good.
      try {
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(width * scale);
        canvas.height = Math.round(height * scale);
        const ctx = canvas.getContext("2d");
        if (!ctx) return fail();
        ctx.scale(scale, scale);
        ctx.drawImage(img, 0, 0, width, height);
        const png = canvas.toDataURL("image/png").split(",")[1];
        if (!png) return fail();
        resolve(png);
      } catch {
        fail();
      }
    };
    img.onerror = fail;
    img.src = url;
  });
}

/** The sample dropdown: one entry per Style theme subject, defaulting to the current theme's. */
function setupSampleControls() {
  const select = $<HTMLSelectElement>("sample-select");
  for (const s of SAMPLES) {
    select.appendChild(
      Object.assign(document.createElement("option"), { value: s.id, textContent: s.name })
    );
  }
  // Start on the subject that matches the Style theme in use.
  const ind = industryOf();
  select.value = ind && sampleById(ind) ? ind : "cashflow";
}

async function insertSampleData() {
  const sample = sampleById($<HTMLSelectElement>("sample-select").value) ?? SAMPLES[0];
  const data = sample.rows;
  await Excel.run(async (context) => {
    const anchor = context.workbook.getSelectedRange().getCell(0, 0);
    const target = anchor.getResizedRange(data.length - 1, 2);
    target.load("values");
    await context.sync();
    const occupied = target.values.some((row) => row.some((v) => v !== "" && v !== null));
    if (occupied) {
      throw new Error(
        `The ${sample.name} sample needs an empty block of ${data.length} rows × 3 columns. Select an empty cell and try again.`
      );
    }
    // Match the Style theme to the sample before it's drawn.
    const theme = $<HTMLSelectElement>("color-scheme");
    if (theme.value !== sample.theme) {
      theme.value = sample.theme;
      theme.dispatchEvent(new Event("change"));
    }
    target.values = data;
    target.getRow(0).format.font.bold = true;
    target.format.autofitColumns();
    target.select();
    target.load("values,address");
    await context.sync();
    applyValues(target.values, target.address);
  });
  await releaseTable();
}

/**
 * The chart as a picture: drawn off-screen at a fixed width, with the user's
 * background, colours and moved positions ("Auto" background = white). Used
 * for inserting into the sheet and for the pop-out window.
 */
function renderPicture(width = chartSize.width ?? EXPORT_WIDTH): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  const { layout, margin } = buildLayout(
    width,
    chartMargin(16, 16),
    undefined,
    true,
    chartSize.height
  );
  const theme = chartTheme(true);
  renderSankey(svg, layout, {
    theme,
    transparent: backgroundMode() === "transparent",
    groups: groupsFor(layout.nodes, theme),
    title: chartTitle,
    linkColor: $<HTMLSelectElement>("link-color").value as LinkColorMode,
    labels: labelSettings,
    labelOverrides,
    colors: chartColors(),
    margin,
  });
  return svg;
}

/* ---------- Pop-out chart window ---------- */

/**
 * What the pop-out window shows: the chart as it'll be inserted (same colours,
 * background and title), laid out for the window's width (no narrower than in
 * the pane, so labels have room). The window draws it itself, interactively.
 */
function popoutContent(width: number): PopoutContent {
  const theme = chartTheme(true);
  const transparent = backgroundMode() === "transparent";
  let chart: ChartModel | null = null;
  if (current) {
    try {
      // A chosen size wins over fitting the window.
      const { layout, margin } = buildLayout(
        chartSize.width ?? Math.max(Math.round(width), MIN_CHART_WIDTH),
        chartMargin(16, 16),
        undefined,
        true,
        chartSize.height
      );
      const gm = groupsFor(layout.nodes, theme);
      chart = {
        layout: layoutToModel(layout),
        theme,
        linkColor: $<HTMLSelectElement>("link-color").value as LinkColorMode,
        labels: labelSettings,
        labelOverrides: Array.from(labelOverrides),
        colors: Array.from(chartColors()),
        groups: gm
          ? Array.from(gm, ([n, g]): [number, NodeGroup] => [
              n.index,
              { name: g.name, color: g.color },
            ])
          : undefined,
        margin,
        transparent,
        title: chartTitle,
        selected: selectedNode,
        size: chartSize,
      };
    } catch {
      chart = null; // e.g. the data can't be drawn; the pane shows why.
    }
  }
  const options = {} as ToolbarOptions;
  for (const id of TOOLBAR_OPTION_IDS) options[id] = $<HTMLInputElement>(id).value;
  return { chart, background: theme.surface, text: theme.textSecondary, transparent, options };
}

/** The toolbar's options; while popped out they're set from the window's toolbar. */
const TOOLBAR_OPTION_IDS: ToolbarOptionId[] = [
  "background-mode",
  "background-color",
  "node-order",
  "end-align",
  "link-color",
];

/**
 * Something you did to the chart in the pop-out window: applied here exactly
 * as if done on the pane's chart, then redrawn (which updates the window too).
 */
function popoutAction(a: PopoutAction) {
  if (a.action === "close") {
    closePopout();
    return;
  }
  if (a.action === "option") {
    // Set the pane's own control and fire its usual event, so it all runs as
    // if changed here (saving, the Labels swatch, the redraw that updates the window).
    if (!TOOLBAR_OPTION_IDS.includes(a.id)) return;
    const el = $<HTMLInputElement>(a.id);
    if (el.value === a.value) return;
    if (el instanceof HTMLSelectElement && !Array.from(el.options).some((o) => o.value === a.value)) return;
    el.value = a.value;
    el.dispatchEvent(new Event(a.id === "background-color" ? "input" : "change"));
    return;
  }
  if (!current) return;
  switch (a.action) {
    case "nodeMoved":
      nodePositions.set(a.name, a.position);
      markMoved(a.name);
      break;
    case "labelMoved":
      setOverride(a.name, { ...labelOverrides.get(a.name), dx: a.dx, dy: a.dy });
      markMoved(a.name);
      break;
    case "resize":
      // Size set in the window's size menu: the chart size everywhere.
      chartSize = { width: a.width, height: a.height };
      break;
    case "select":
      // Opens that box's settings here in the pane, beside the window.
      if (a.name) return selectNode(a.name, false);
      if (!selectedNode) return;
      selectedNode = null;
      break;
  }
  updateResetButtons();
  draw();
}

function setupPopout() {
  // While the chart is in its own window the pane is settings-only: the chart
  // gives way to a slim note with a button to bring it back. (The pane's chart
  // redraws to fit when it reappears.)
  const changed = (open: boolean) => {
    $("app-body").classList.toggle("popped-out", open);
    $("popout-placeholder").hidden = !open;
    if (!open && current) draw();
  };
  $("popout-btn").onclick = () => {
    if (!popoutOpen()) {
      openPopout(popoutContent, changed, (message) => setStatus(message, "error"), popoutAction);
    }
  };
  $("popout-return").onclick = () => closePopout();
}

async function insertChartIntoSheet() {
  if (!current) return;
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    throw new Error(
      "This version of Excel can't insert pictures from add-ins. Try Excel for Microsoft 365 or Excel on the web."
    );
  }

  const svg = renderPicture();
  const width = Number(svg.getAttribute("width"));
  const height = Number(svg.getAttribute("height"));
  const png = await svgToPngBase64(svg, width, height, 2);

  const source = current.source;
  const diagram = diagramById(openDiagramId);
  const picture = await Excel.run(async (context): Promise<SavedPicture> => {
    // Place the picture beside the data it was drawn from.
    let ws: Excel.Worksheet;
    let sheetName: string;
    let dataRange: Excel.Range;
    if (source.kind === "table") {
      // Look the sheet up fresh, in case the table was moved or its sheet renamed.
      const table = context.workbook.tables.getItem(source.id);
      table.load("worksheet/name");
      await context.sync();
      sheetName = table.worksheet.name;
      ws = context.workbook.worksheets.getItem(sheetName);
      dataRange = table.getRange();
    } else {
      const { sheet, address } = splitAddress(source.address);
      sheetName = sheet;
      ws = context.workbook.worksheets.getItem(sheet);
      dataRange = ws.getRange(address);
    }
    dataRange.load("left,top,width");
    const shape = ws.shapes.addImage(png);
    // Named after its diagram; forgetting the diagram renames it (the picture stays).
    shape.name = pictureName(diagram?.name ?? defaultDiagramName(source));
    shape.load("id");
    await context.sync();
    // Display at 1x size (the bitmap is 2x for sharpness); Excel sizes are in points.
    // Width AND height are both set: Excel doesn't reliably apply the aspect-ratio
    // lock to sizes set from code, so setting only the width could leave the
    // height at the bitmap's 2x size and stretch the picture.
    // Sized from the bitmap's whole-pixel dimensions (2x), so the ratio is exact.
    shape.lockAspectRatio = false;
    shape.width = (Math.round(width * 2) / 2) * 0.75;
    shape.height = (Math.round(height * 2) / 2) * 0.75;
    shape.lockAspectRatio = true; // keeps its shape if you resize it by hand
    // Move with the cells underneath, but never resize with them: by default a
    // picture stretches when those columns or rows change size (e.g. a table
    // auto-fitting its columns as you edit it).
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.10")) {
      shape.placement = Excel.Placement.oneCell;
    }
    shape.left = dataRange.left + dataRange.width + 24;
    shape.top = dataRange.top;
    await context.sync();
    return { sheet: sheetName, shapeId: shape.id };
  });
  // Remember the picture with its diagram: straight away if it's saved
  // (without saving other unsaved changes), otherwise with its first save.
  if (diagram) {
    await saveDiagram({ ...diagram, pictures: [...diagram.pictures, picture] });
    renderMemoryTab();
  } else {
    pendingPictures.push(picture);
  }
  setStatus(
    `Chart inserted next to ${source.kind === "table" ? `table "${source.name}"` : source.address}.`
  );
}
