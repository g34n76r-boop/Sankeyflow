/*
 * Snapping and alignment guides for moving Sankey nodes.
 *
 * A node can line up with:
 *  - horizontally: the top, middle or bottom of any other node, or the chart's vertical middle;
 *  - vertically: any other node's x position, or one of the chart's original column positions.
 */

import type { SankeyLayout, SankeyNode } from "./sankey";

/** How close (px) a dragged node must get before it snaps. */
export const SNAP_DISTANCE = 6;
/** Tolerance (px) for treating two positions as lined up. */
const EPS = 0.5;

export interface Guide {
  /** "x": a vertical line at x = pos. "y": a horizontal line at y = pos. */
  axis: "x" | "y";
  pos: number;
  from: number;
  to: number;
  /** Snap points drawn on the line: one per aligned node (y for vertical guides, x for horizontal). */
  points: number[];
}

const mid = (n: SankeyNode) => (n.y0 + n.y1) / 2;
const others = (n: SankeyNode, layout: SankeyLayout) => layout.nodes.filter((o) => o !== n);

function columnXs(n: SankeyNode, layout: SankeyLayout): number[] {
  const w = n.x1 - n.x0;
  if (layout.columnCount <= 1) return [(layout.width - w) / 2];
  const step = (layout.width - w) / (layout.columnCount - 1);
  return Array.from({ length: layout.columnCount }, (_, c) => c * step);
}

/** Values of n.x0 that would line n up with something. */
function xTargets(n: SankeyNode, layout: SankeyLayout): number[] {
  return [...columnXs(n, layout), ...others(n, layout).map((o) => o.x0)];
}

/** Values of n.y0 that would line n's top, middle or bottom up with something. */
function yTargets(n: SankeyNode, layout: SankeyLayout): number[] {
  const h = n.y1 - n.y0;
  const out = [layout.height / 2 - h / 2];
  for (const o of others(n, layout)) {
    for (const t of [o.y0, mid(o), o.y1]) {
      out.push(t, t - h / 2, t - h);
    }
  }
  return out;
}

function nearest(value: number, targets: number[]): number {
  let best = value;
  let bestDist = SNAP_DISTANCE + EPS;
  for (const t of targets) {
    const d = Math.abs(t - value);
    if (d < bestDist) {
      best = t;
      bestDist = d;
    }
  }
  return best;
}

/** For dragging: pulls a proposed position onto any target within SNAP_DISTANCE, per axis. */
export function snapPosition(
  n: SankeyNode,
  layout: SankeyLayout,
  x0: number,
  y0: number
): { x0: number; y0: number } {
  return { x0: nearest(x0, xTargets(n, layout)), y0: nearest(y0, yTargets(n, layout)) };
}

/**
 * For arrow keys: moving from `from` toward `to`, stop at the first alignment
 * crossed on the way, so a big step never skips past a line-up.
 */
export function snapStep(
  n: SankeyNode,
  layout: SankeyLayout,
  axis: "x" | "y",
  from: number,
  to: number
): number {
  const targets = axis === "x" ? xTargets(n, layout) : yTargets(n, layout);
  const up = to > from;
  const between = targets.filter((t) => (up ? t > from + EPS && t < to : t < from - EPS && t > to));
  if (!between.length) return to;
  return up ? Math.min(...between) : Math.max(...between);
}

/** Guides for everything n currently lines up with. */
export function findGuides(n: SankeyNode, layout: SankeyLayout): Guide[] {
  const guides: Guide[] = [];
  const rest = others(n, layout);
  const cx = (n.x0 + n.x1) / 2;
  const pad = 8;

  // Vertical guide: same x as other nodes, or on an empty column position.
  const xMatches = rest.filter((o) => Math.abs(o.x0 - n.x0) < EPS);
  if (xMatches.length) {
    const all = [n, ...xMatches];
    guides.push({
      axis: "x",
      pos: cx,
      from: Math.min(...all.map((o) => o.y0)) - pad,
      to: Math.max(...all.map((o) => o.y1)) + pad,
      points: all.map(mid),
    });
  } else if (columnXs(n, layout).some((x) => Math.abs(x - n.x0) < EPS)) {
    guides.push({ axis: "x", pos: cx, from: 0, to: layout.height, points: [mid(n)] });
  }

  // Horizontal guides: n's top, middle and bottom against other nodes' top, middle and bottom.
  for (const y of [n.y0, mid(n), n.y1]) {
    const matched = rest.filter((o) => [o.y0, mid(o), o.y1].some((t) => Math.abs(t - y) < EPS));
    if (!matched.length) continue;
    const all = [n, ...matched];
    guides.push({
      axis: "y",
      pos: y,
      from: Math.min(...all.map((o) => o.x0)) - pad,
      to: Math.max(...all.map((o) => o.x1)) + pad,
      points: all.map((o) => (o.x0 + o.x1) / 2),
    });
  }

  // Centered vertically in the chart.
  const center = layout.height / 2;
  if (
    Math.abs(mid(n) - center) < EPS &&
    !guides.some((g) => g.axis === "y" && Math.abs(g.pos - center) < EPS)
  ) {
    guides.push({ axis: "y", pos: center, from: 0, to: layout.width, points: [cx] });
  }

  return guides;
}
