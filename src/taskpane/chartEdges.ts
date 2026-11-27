/*
 * The dotted lines along the chart's right and bottom edges (in the task pane
 * and the pop-out window), which show the chart's size. Hovering a line shows
 * that dimension (width for the right edge, height for the bottom edge); clicking
 * it (or Enter when it has keyboard focus) opens a small box to change just that
 * dimension. A blank box means automatic.
 *
 * The lines are CSS only: they're not part of the chart inserted into the sheet.
 */

/* global document, window, HTMLElement, HTMLInputElement, SVGSVGElement, Node, KeyboardEvent, PointerEvent, MouseEvent */

import { CHART_SIZE_LIMITS as LIMITS, parseChartSize } from "./resize";
import type { ChartSize } from "./sizeMenu";

export type Axis = "width" | "height";

export interface ChartEdges {
  /** Shows the chosen size, with the chart's drawn size for "automatic". */
  sync(size: ChartSize, drawn: { width: number; height: number }): void;
  close(): void;
}
const NAME: Record<Axis, string> = { width: "Width", height: "Height" };

/**
 * Wraps `svg` so the two edge lines sit over its right and bottom edges, at
 * whatever size it's shown. `onChange` gets the axis and its new size (null =
 * automatic) when a value is entered.
 */
export function attachChartEdges(
  svg: SVGSVGElement,
  onChange: (axis: Axis, value: number | null) => void
): ChartEdges {
  const wrap = document.createElement("div");
  wrap.className = "chart-edges";
  svg.replaceWith(wrap);
  wrap.appendChild(svg);

  let size: ChartSize = { width: null, height: null };
  let drawn = { width: 0, height: 0 };

  const describe = (axis: Axis) => {
    const v = size[axis];
    const d = Math.round(drawn[axis]);
    return v === null ? `${NAME[axis]}: ${d || "–"} px (automatic)` : `${NAME[axis]}: ${v} px`;
  };

  // Hover label, following the pointer along the line.
  const label = document.createElement("div");
  label.className = "chart-edge-label";
  label.hidden = true;
  document.body.appendChild(label);

  // The small box for typing a new value.
  const pop = document.createElement("div");
  pop.className = "chart-edge-pop";
  pop.hidden = true;
  pop.setAttribute("role", "dialog");
  pop.innerHTML = `
    <label class="chart-edge-pop-row"><span class="chart-edge-pop-name"></span>
      <input type="number" class="chart-edge-input" step="10" />
      <span class="chart-edge-pop-unit">px</span></label>
    <div class="chart-edge-pop-hint">Blank for automatic · Enter to apply</div>`;
  document.body.appendChild(pop);
  const input = pop.querySelector<HTMLInputElement>(".chart-edge-input")!;
  const popName = pop.querySelector<HTMLElement>(".chart-edge-pop-name")!;

  let openAxis: Axis | null = null;
  let opener: HTMLElement | null = null;

  const place = (el: HTMLElement, x: number, y: number) => {
    // Keep it inside the window.
    const pad = 6;
    const w = el.offsetWidth;
    const h = el.offsetHeight;
    const vw = document.documentElement.clientWidth;
    const vh = document.documentElement.clientHeight;
    el.style.left = `${Math.max(pad, Math.min(x, vw - w - pad))}px`;
    el.style.top = `${Math.max(pad, Math.min(y, vh - h - pad))}px`;
  };

  const close = (apply: boolean, refocus = false) => {
    if (!openAxis) return;
    const axis = openAxis;
    openAxis = null;
    pop.hidden = true;
    opener?.classList.remove("active");
    if (refocus) opener?.focus();
    if (!apply) return;
    const v = parseChartSize(input.value, axis);
    if (v !== size[axis]) onChange(axis, v);
  };

  const open = (axis: Axis, edge: HTMLElement, x: number, y: number) => {
    close(false);
    label.hidden = true;
    openAxis = axis;
    opener = edge;
    edge.classList.add("active");
    const [lo, hi] = LIMITS[axis];
    popName.textContent = NAME[axis];
    input.min = String(lo);
    input.max = String(hi);
    input.value = size[axis]?.toString() ?? "";
    input.placeholder = `auto (${Math.round(drawn[axis]) || "–"})`;
    input.setAttribute("aria-label", `Chart ${axis} in pixels (blank for automatic)`);
    pop.setAttribute("aria-label", `Chart ${axis}`);
    pop.hidden = false;
    // Beside the line, on the chart's side of it.
    if (axis === "width") place(pop, x - pop.offsetWidth - 10, y - pop.offsetHeight / 2);
    else place(pop, x - pop.offsetWidth / 2, y - pop.offsetHeight - 10);
    input.focus();
    input.select();
  };

  const makeEdge = (axis: Axis) => {
    const edge = document.createElement("div");
    edge.className = `chart-edge chart-edge-${axis === "width" ? "right" : "bottom"}`;
    edge.tabIndex = 0;
    edge.setAttribute("role", "button");
    wrap.appendChild(edge);

    edge.addEventListener("pointermove", (e: PointerEvent) => {
      if (openAxis) return;
      label.textContent = describe(axis);
      label.hidden = false;
      place(label, e.clientX + 12, e.clientY + 14);
    });
    edge.addEventListener("pointerleave", () => (label.hidden = true));
    // Don't let the chart under it start a drag or clear the selection.
    edge.addEventListener("pointerdown", (e: PointerEvent) => e.stopPropagation());
    edge.addEventListener("click", (e: MouseEvent) => {
      e.stopPropagation();
      if (openAxis === axis) close(true);
      else open(axis, edge, e.clientX, e.clientY);
    });
    edge.addEventListener("keydown", (e: KeyboardEvent) => {
      if (e.key !== "Enter" && e.key !== " ") return;
      e.preventDefault();
      const r = edge.getBoundingClientRect();
      open(axis, edge, r.left + r.width / 2, r.top + r.height / 2);
    });
    return edge;
  };
  const edges: Record<Axis, HTMLElement> = { width: makeEdge("width"), height: makeEdge("height") };

  input.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Enter") {
      e.preventDefault();
      close(true, true);
    } else if (e.key === "Escape") {
      e.preventDefault();
      close(false, true);
    }
  });
  // A click elsewhere applies what was typed; scrolling or resizing closes the box too.
  document.addEventListener(
    "pointerdown",
    (e) => {
      if (openAxis && !pop.contains(e.target as Node) && !(e.target as HTMLElement).closest?.(".chart-edge"))
        close(true);
    },
    true
  );
  const dismiss = () => {
    label.hidden = true;
    close(true);
  };
  window.addEventListener("scroll", dismiss, true);
  window.addEventListener("resize", dismiss);

  const update = () => {
    for (const axis of ["width", "height"] as Axis[]) {
      edges[axis].setAttribute("aria-label", `Chart ${describe(axis)}. Press Enter to change.`);
    }
  };
  update();

  return {
    sync(s, d) {
      size = s;
      drawn = d;
      update();
    },
    close: () => close(false),
  };
}
