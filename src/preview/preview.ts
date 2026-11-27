/* global Office, document, window, setTimeout, clearTimeout, SVGSVGElement, HTMLButtonElement, HTMLElement, HTMLInputElement, HTMLSelectElement */

/*
 * The pop-out chart window. The task pane sends the chart (its finished layout
 * and styling) in numbered chunks whenever it changes; this page puts them
 * together and draws it with the same chart code as the pane, so it's fully
 * interactive: hover for details, click a box to open its settings in the pane,
 * and drag boxes and labels (with the same snapping and guides). Moves and
 * selections go back to the pane, which sends the updated chart back here.
 */

import { renderSankey } from "../taskpane/chart";
import { getNodePosition } from "../taskpane/sankey";
import { TOOLBAR_OPTION_IDS, modelToLayout } from "../taskpane/popoutModel";
import { applyZoom, stepZoom, updateZoomButtons, zoomKeepingCentre } from "../taskpane/zoom";
import { attachSizeMenu } from "../taskpane/sizeMenu";
import type { ChartSize, SizeMenu } from "../taskpane/sizeMenu";
import { attachChartEdges } from "../taskpane/chartEdges";
import type { ChartEdges } from "../taskpane/chartEdges";
import { setupToolbarIconSelects } from "../taskpane/iconSelect";
import { showTooltip } from "../taskpane/tooltip";
import type {
  ChartModel,
  PopoutAction,
  PopoutContent,
  ToolbarOptions,
} from "../taskpane/popoutModel";

interface Chunk {
  id: number;
  i: number;
  n: number;
  d: string;
}

let latest = 0;
let parts: string[] = [];

/** How magnified the chart is in this window (its own + / − buttons). */
let zoom = 1;
/** The size menu from the dimensions icon. */
let sizeMenu: SizeMenu | null = null;
/** The chart's dotted right and bottom edges, which show and change its size. */
let chartEdges: ChartEdges | null = null;
/** The chart size chosen in the task pane (null = automatic). */
let currentSize: ChartSize = { width: null, height: null };

/** A pointer is down on the chart (maybe dragging): updates wait until it's released. */
let pressing = false;
let pending: PopoutContent | null = null;

/** Width the chart has in this window (flush left, 16px gap on the right), so
 *  the task pane can lay it out to fit. */
const chartWidth = () => Math.max(1, document.documentElement.clientWidth - 16);

const tell = (a: PopoutAction) => Office.context.ui.messageParent(JSON.stringify(a));

Office.onReady(() => {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    receive(arg.message);
  });
  // Ask the task pane for the chart, laid out to fit this window...
  Office.context.ui.messageParent(JSON.stringify({ ready: true, width: chartWidth() }));
  // ...and again (laid out to the new width) when the window is resized.
  let timer: ReturnType<typeof setTimeout> | undefined;
  let lastWidth = chartWidth();
  window.addEventListener("resize", () => {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      const w = chartWidth();
      if (w === lastWidth) return;
      lastWidth = w;
      Office.context.ui.messageParent(JSON.stringify({ width: w }));
    }, 200);
  });

  // Bring the chart back into the task pane (the pane closes this window).
  document.getElementById("popin-btn")!.onclick = () => tell({ action: "close" });

  // Background / Order / Ends / Links: the same icon dropdowns as in the pane.
  // A change is made in the pane, which sends the redrawn chart back here.
  setupToolbarIconSelects();
  for (const id of TOOLBAR_OPTION_IDS) {
    const el = document.getElementById(id) as HTMLInputElement;
    el.addEventListener(id === "background-color" ? "input" : "change", () => {
      if (id === "background-mode") syncCustomColor();
      tell({ action: "option", id, value: el.value });
    });
  }

  // + / − magnify or shrink the chart, keeping the middle of the view in place.
  const zin = document.getElementById("zoom-in-btn") as HTMLButtonElement;
  const zout = document.getElementById("zoom-out-btn") as HTMLButtonElement;
  const level = document.getElementById("zoom-level");
  const svg = document.getElementById("chart-svg") as unknown as SVGSVGElement;
  const step = (dir: 1 | -1) => {
    zoomKeepingCentre(document.scrollingElement ?? document.documentElement, () => {
      zoom = stepZoom(zoom, dir);
      applyZoom(svg, zoom);
    });
    updateZoomButtons(zin, zout, zoom, level);
  };
  zin.onclick = () => step(1);
  zout.onclick = () => step(-1);
  updateZoomButtons(zin, zout, zoom, level);

  // The size menu: the size is set in the task pane (it applies there and to
  // the inserted picture too), which sends the chart back laid out at that size.
  sizeMenu = attachSizeMenu(document.getElementById("size-btn") as HTMLButtonElement, (s) =>
    tell({ action: "resize", width: s.width, height: s.height })
  );

  // The dotted right and bottom edges: hover for the width or height, click to
  // change just that one (also set in the task pane, like the size menu).
  chartEdges = attachChartEdges(svg, (axis, value) => {
    const s = { ...currentSize, [axis]: value };
    tell({ action: "resize", width: s.width, height: s.height });
  });

  // Don't redraw under the pointer mid-drag; apply any update once it's released
  // (after the chart has reported the move).
  const stage = document.getElementById("stage")!;
  stage.addEventListener("pointerdown", () => (pressing = true), true);
  const release = () => {
    if (!pressing) return;
    pressing = false;
    setTimeout(() => {
      if (pending && !pressing) {
        const p = pending;
        pending = null;
        show(p);
      }
    }, 0);
  };
  window.addEventListener("pointerup", release, true);
  window.addEventListener("pointercancel", release, true);
});

function receive(message: string) {
  let c: Chunk;
  try {
    c = JSON.parse(message);
  } catch {
    return;
  }
  // A newer update replaces one still arriving.
  if (c.id < latest) return;
  if (c.id > latest) {
    latest = c.id;
    parts = new Array(c.n);
  }
  parts[c.i] = c.d;
  if (parts.filter((p) => p !== undefined).length < c.n) return;
  let content: PopoutContent;
  try {
    content = JSON.parse(parts.join(""));
  } catch {
    return; // Incomplete or broken update: keep showing the last one.
  }
  if (pressing) pending = content;
  else show(content);
}

function show(content: PopoutContent) {
  const body = document.body;
  body.style.setProperty("--bg", content.background);
  body.style.setProperty("--text", content.text);
  body.classList.toggle("transparent", content.transparent);
  body.classList.toggle("has-chart", !!content.chart);

  const message = document.getElementById("message")!;
  const frame = document.getElementById("chart-frame")!;
  message.hidden = !!content.chart;
  frame.hidden = !content.chart;
  for (const el of Array.from(document.querySelectorAll<HTMLElement>(".chart-only"))) {
    el.hidden = !content.chart;
  }
  if (content.options) showOptions(content.options);
  if (!content.chart) {
    message.textContent = content.error
      ? `The chart can't be drawn: ${content.error}`
      : "No data loaded. Pick a table or load cells in the task pane to see the chart here.";
    return;
  }
  draw(content.chart);
}

/** Shows the toolbar's options as set in the pane. */
function showOptions(o: ToolbarOptions) {
  for (const id of TOOLBAR_OPTION_IDS) {
    const el = document.getElementById(id) as HTMLInputElement;
    // Don't reset the colour picker while it's in use (its updates lag behind it).
    if (id === "background-color" && document.activeElement === el) continue;
    if (typeof o[id] === "string" && el.value !== o[id]) el.value = o[id];
  }
  syncCustomColor();
}

/** The colour picker shows only with a custom background, as in the pane. */
function syncCustomColor() {
  const mode = document.getElementById("background-mode") as HTMLSelectElement;
  document.getElementById("background-color")!.hidden = mode.value !== "custom";
}

function draw(m: ChartModel) {
  const svg = document.getElementById("chart-svg") as unknown as SVGSVGElement;
  const tooltip = document.getElementById("tooltip")!;
  const frame = document.getElementById("chart-frame")!;
  tooltip.hidden = true;
  const layout = modelToLayout(m.layout);
  const byIndex = new Map(layout.nodes.map((n) => [n.index, n]));
  const groups = m.groups
    ? new Map(m.groups.flatMap(([i, g]) => (byIndex.has(i) ? [[byIndex.get(i)!, g]] : [])))
    : undefined;

  renderSankey(svg, layout, {
    theme: m.theme,
    linkColor: m.linkColor,
    labels: m.labels,
    labelOverrides: new Map(m.labelOverrides),
    colors: new Map(m.colors),
    groups,
    margin: m.margin,
    transparent: m.transparent,
    title: m.title,
    interactive: {
      selected: m.selected,
      onNodeSelect: (n) => tell({ action: "select", name: n.name }),
      onBackgroundClick: () => tell({ action: "select", name: null }),
      onNodeMoved: (n, l) =>
        tell({ action: "nodeMoved", name: n.name, position: getNodePosition(n, l) }),
      onLabelMoved: (n, dx, dy) => tell({ action: "labelMoved", name: n.name, dx, dy }),
      onHover: (h) => {
        if (!h) {
          tooltip.hidden = true;
          return;
        }
        // Keep the tooltip inside the window: the chart frame can be bigger
        // than it (zoomed in) and the page scrolls, so measure what's visible.
        const r = frame.getBoundingClientRect();
        const view = document.documentElement;
        showTooltip(tooltip, h.html, h.x, h.y, {
          left: Math.max(0, -r.left),
          top: Math.max(0, -r.top),
          right: view.clientWidth - r.left,
          bottom: view.clientHeight - r.top,
        });
      },
    },
  });
  applyZoom(svg, zoom);
  currentSize = m.size ?? { width: null, height: null };
  const drawn = {
    width: Number(svg.getAttribute("width")) || 0,
    height: Number(svg.getAttribute("height")) || 0,
  };
  sizeMenu?.sync(currentSize, drawn);
  chartEdges?.sync(currentSize, drawn);
}
