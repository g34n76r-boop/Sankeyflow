/*
 * Chart magnification for the + / − buttons, shared by the task pane and the
 * pop-out window. The chart is drawn as usual and then shown larger or smaller
 * (labels included); dragging allows for it (see zoomOf in chart.ts).
 */

/* global SVGSVGElement, HTMLButtonElement, HTMLElement */

const ZOOM_LEVELS = [0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];

/** The next zoom level in (+1) or out (−1) from `z`. */
export function stepZoom(z: number, dir: 1 | -1): number {
  if (dir > 0) return ZOOM_LEVELS.find((l) => l > z + 1e-6) ?? ZOOM_LEVELS[ZOOM_LEVELS.length - 1];
  return [...ZOOM_LEVELS].reverse().find((l) => l < z - 1e-6) ?? ZOOM_LEVELS[0];
}

const canZoom = (z: number, dir: 1 | -1) =>
  dir > 0 ? z < ZOOM_LEVELS[ZOOM_LEVELS.length - 1] - 1e-6 : z > ZOOM_LEVELS[0] + 1e-6;

/** Shows a drawn chart at `z` times its size. */
export function applyZoom(svg: SVGSVGElement, z: number) {
  const w = Number(svg.getAttribute("width"));
  const h = Number(svg.getAttribute("height"));
  if (!w || !h) return;
  svg.style.width = `${w * z}px`;
  svg.style.height = `${h * z}px`;
}

/**
 * Sets a + / − button's tooltip and enabled state for the current zoom, and
 * shows the zoom as a percentage in `level` (to the right of the buttons).
 */
export function updateZoomButtons(
  zoomIn: HTMLButtonElement,
  zoomOut: HTMLButtonElement,
  z: number,
  level?: HTMLElement | null
) {
  const pct = `${Math.round(z * 100)}%`;
  if (level) level.textContent = pct;
  zoomIn.disabled = !canZoom(z, 1);
  zoomOut.disabled = !canZoom(z, -1);
  zoomIn.title = `Zoom in (now ${pct})`;
  zoomOut.title = `Zoom out (now ${pct})`;
  zoomIn.setAttribute("aria-label", `Zoom in, now ${pct}`);
  zoomOut.setAttribute("aria-label", `Zoom out, now ${pct}`);
}

/**
 * Zooms, keeping the middle of what's visible in the middle, then redraws the
 * buttons. `scroller` is the element that scrolls the chart.
 */
export function zoomKeepingCentre(
  scroller: {
    scrollLeft: number;
    scrollTop: number;
    scrollWidth: number;
    scrollHeight: number;
    clientWidth: number;
    clientHeight: number;
  },
  apply: () => void
) {
  const cx = (scroller.scrollLeft + scroller.clientWidth / 2) / Math.max(1, scroller.scrollWidth);
  const cy = (scroller.scrollTop + scroller.clientHeight / 2) / Math.max(1, scroller.scrollHeight);
  apply();
  scroller.scrollLeft = cx * scroller.scrollWidth - scroller.clientWidth / 2;
  scroller.scrollTop = cy * scroller.scrollHeight - scroller.clientHeight / 2;
}
