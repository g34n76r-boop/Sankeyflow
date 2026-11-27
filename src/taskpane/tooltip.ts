/*
 * Places the chart's hover tooltip (task pane and pop-out window): beside the
 * pointer, flipped to the other side when it would run past the visible area.
 */

/* global HTMLElement */

/** The visible part of the tooltip's container, in the same coordinates as the pointer. */
export interface VisibleArea {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

const PAD = 12;

export function showTooltip(
  tooltip: HTMLElement,
  html: string,
  x: number,
  y: number,
  visible: VisibleArea
) {
  tooltip.innerHTML = html;
  tooltip.hidden = false;
  const w = tooltip.offsetWidth;
  const h = tooltip.offsetHeight;
  let left = x + PAD;
  if (left + w > visible.right) left = x - w - PAD;
  let top = y + PAD;
  if (top + h > visible.bottom) top = y - h - PAD;
  tooltip.style.left = `${Math.max(visible.left, left)}px`;
  tooltip.style.top = `${Math.max(visible.top, top)}px`;
}
