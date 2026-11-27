/*
 * The bar on the bottom edge of the task pane's chart preview, styled like a
 * scroll-bar thumb: drag it up or down to make the preview area shorter or
 * taller (the chart itself keeps its size and scrolls inside a shorter area).
 * Double-click it to fit the preview to the chart again. It also works from
 * the keyboard: focus it and use the arrow keys. When the preview has a
 * horizontal scroll bar, the bar sits just above it.
 *
 * The chart's own size is set only from the size menu (sizeMenu.ts).
 */

/* global document, HTMLElement, PointerEvent, KeyboardEvent */

export const MIN_CHART_W = 320;
export const MAX_CHART_W = 2400;
export const MIN_CHART_H = 160;
export const MAX_CHART_H = 2000;

/** Limits for the preview area's height. */
const MIN_VIEW_H = 120;
const MAX_VIEW_H = 2400;

const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v));

const BAR_LENGTH = 44;
const BAR_THICKNESS = 8;
/** Space between the bar and a horizontal scroll bar under it. */
const SCROLLBAR_GAP = 4;

/**
 * Adds the bar to `host` (a positioned element containing the scrolling
 * preview `box`). Call the returned function after each redraw, zoom or scroll
 * to keep the bar on the box's bottom edge.
 */
export function attachPreviewResizer(host: HTMLElement, box: HTMLElement): () => void {
  const bar = document.createElement("div");
  bar.className = "chart-size-bar chart-size-bar-height";
  bar.tabIndex = 0;
  bar.setAttribute("role", "separator");
  bar.setAttribute("aria-orientation", "horizontal");
  bar.setAttribute("aria-label", "Preview height");
  bar.setAttribute("aria-valuemin", String(MIN_VIEW_H));
  bar.setAttribute("aria-valuemax", String(MAX_VIEW_H));
  bar.title =
    "Drag up or down to change the height of the preview. Double-click to fit it to the chart.";
  host.appendChild(bar);

  /** The chosen preview height (null = fit the chart). */
  let viewHeight: number | null = null;

  const setHeight = (h: number | null) => {
    viewHeight = h === null ? null : Math.round(clamp(h, MIN_VIEW_H, MAX_VIEW_H));
    box.style.boxSizing = "border-box"; // the height includes the border, as measured
    box.style.height = viewHeight === null ? "" : `${viewHeight}px`;
    // A shorter preview scrolls the chart up and down too.
    box.style.overflowY = viewHeight === null ? "" : "auto";
    place();
  };

  const place = () => {
    // No chart drawn yet: no bar. (When the tab is hidden it goes with it.)
    if (!box.querySelector("svg")?.getAttribute("width")) {
      bar.hidden = true;
      return;
    }
    if (box.offsetWidth === 0) return;
    bar.hidden = false;
    const h = host.getBoundingClientRect();
    const b = box.getBoundingClientRect();
    const x = b.left - h.left;
    const y = b.top - h.top;
    bar.style.left = `${x + b.width / 2 - BAR_LENGTH / 2}px`;
    const scrollbar =
      box.scrollWidth > box.clientWidth
        ? Math.max(0, box.offsetHeight - box.clientHeight - box.clientTop * 2)
        : 0;
    if (scrollbar > 0) {
      // Just above the horizontal scroll bar, with a small gap.
      const scrollbarTop = y + box.clientTop + box.clientHeight;
      bar.style.top = `${scrollbarTop - SCROLLBAR_GAP - BAR_THICKNESS}px`;
    } else {
      // Centred on the bottom edge.
      bar.style.top = `${y + b.height - BAR_THICKNESS / 2}px`;
    }
    bar.setAttribute("aria-valuenow", String(Math.round(b.height)));
  };

  let start: { y: number; h: number } | null = null;
  bar.addEventListener("pointerdown", (e: PointerEvent) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    try {
      bar.setPointerCapture(e.pointerId);
    } catch {
      // synthetic pointer
    }
    start = { y: e.clientY, h: box.getBoundingClientRect().height };
    bar.classList.add("dragging");
  });
  const end = (e: PointerEvent) => {
    if (!start) return;
    start = null;
    bar.classList.remove("dragging");
    try {
      if (bar.hasPointerCapture(e.pointerId)) bar.releasePointerCapture(e.pointerId);
    } catch {
      // synthetic pointer
    }
  };
  // The preview follows the pointer only while the button is held down: if the
  // release was missed (e.g. it happened outside the pane), the drag ends at
  // the next move instead of carrying on without a click.
  bar.addEventListener("pointermove", (e: PointerEvent) => {
    if (!start) return;
    if ((e.buttons & 1) === 0) {
      end(e);
      return;
    }
    setHeight(start.h + e.clientY - start.y);
  });
  bar.addEventListener("pointerup", end);
  bar.addEventListener("pointercancel", end);
  bar.addEventListener("lostpointercapture", end);
  document.addEventListener("pointerup", end, true);
  bar.addEventListener("dblclick", () => setHeight(null));

  // Arrow keys: 10px a press (50px with Shift); Home fits the chart again.
  bar.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Home") {
      e.preventDefault();
      setHeight(null);
      return;
    }
    const dir = e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0;
    if (!dir) return;
    e.preventDefault();
    setHeight(box.getBoundingClientRect().height + (e.shiftKey ? 50 : 10) * dir);
  });

  place();
  return place;
}
