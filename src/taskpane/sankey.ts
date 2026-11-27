/*
 * Sankey layout: turns a list of flows (source -> target, value) into positioned
 * nodes and links. No dependencies; the algorithm follows the same ideas as
 * d3-sankey (columns by longest path, node heights proportional to value,
 * iterative relaxation to reduce link crossings).
 */

export interface Flow {
  source: string;
  target: string;
  value: number;
}

export interface SankeyNode {
  name: string;
  index: number;
  column: number;
  value: number;
  inValue: number;
  outValue: number;
  x0: number;
  x1: number;
  y0: number;
  y1: number;
  sourceLinks: SankeyLink[]; // links leaving this node
  targetLinks: SankeyLink[]; // links entering this node
}

export interface SankeyLink {
  index: number;
  source: SankeyNode;
  target: SankeyNode;
  value: number;
  width: number;
  y0: number; // vertical center where the link leaves the source
  y1: number; // vertical center where the link enters the target
}

export interface SankeyLayout {
  nodes: SankeyNode[];
  links: SankeyLink[];
  columnCount: number;
  width: number;
  height: number;
}

export interface SankeyOptions {
  width: number;
  height: number;
  nodeWidth?: number;
  nodePadding?: number;
  iterations?: number;
  /** Column for end nodes: "left" (default) keeps them next to their source. */
  align?: "left" | "justify";
  /** Largest share of the height the gaps between boxes may take (default 0.4). */
  maxGapShare?: number;
  /**
   * Keeps boxes with the same key together in each column (e.g. their color).
   * Called once the columns are known.
   */
  groupKeys?: (nodes: SankeyNode[]) => Map<SankeyNode, string>;
}

export class SankeyError extends Error {}

/** Combines duplicate source/target pairs and validates the flows. */
export function aggregateFlows(flows: Flow[]): Flow[] {
  const byKey = new Map<string, Flow>();
  for (const f of flows) {
    if (f.source === f.target) {
      throw new SankeyError(`"${f.source}" flows into itself. A Sankey diagram can't show loops.`);
    }
    const key = f.source + "\u0000" + f.target;
    const existing = byKey.get(key);
    if (existing) existing.value += f.value;
    else byKey.set(key, { source: f.source, target: f.target, value: f.value });
  }
  return Array.from(byKey.values());
}

export function computeSankey(input: Flow[], opts: SankeyOptions): SankeyLayout {
  const { width, height } = opts;
  const nodeWidth = opts.nodeWidth ?? 12;
  const iterations = opts.iterations ?? 32;
  const flows = aggregateFlows(input);

  // Nodes, in order of first appearance (this order also drives color).
  const nodeByName = new Map<string, SankeyNode>();
  const nodes: SankeyNode[] = [];
  const getNode = (name: string): SankeyNode => {
    let n = nodeByName.get(name);
    if (!n) {
      n = {
        name,
        index: nodes.length,
        column: 0,
        value: 0,
        inValue: 0,
        outValue: 0,
        x0: 0,
        x1: 0,
        y0: 0,
        y1: 0,
        sourceLinks: [],
        targetLinks: [],
      };
      nodeByName.set(name, n);
      nodes.push(n);
    }
    return n;
  };

  const links: SankeyLink[] = flows.map((f, i) => {
    const source = getNode(f.source);
    const target = getNode(f.target);
    const link: SankeyLink = { index: i, source, target, value: f.value, width: 0, y0: 0, y1: 0 };
    source.sourceLinks.push(link);
    target.targetLinks.push(link);
    source.outValue += f.value;
    target.inValue += f.value;
    return link;
  });

  for (const n of nodes) n.value = Math.max(n.inValue, n.outValue);

  // Columns: longest path from any start node (topological order). Detects cycles.
  const depth = new Map<SankeyNode, number>();
  const indegree = new Map<SankeyNode, number>();
  for (const n of nodes) {
    indegree.set(n, n.targetLinks.length);
    depth.set(n, 0);
  }
  const queue = nodes.filter((n) => n.targetLinks.length === 0);
  let processed = 0;
  while (queue.length) {
    const n = queue.shift()!;
    processed++;
    for (const l of n.sourceLinks) {
      depth.set(l.target, Math.max(depth.get(l.target)!, depth.get(n)! + 1));
      const d = indegree.get(l.target)! - 1;
      indegree.set(l.target, d);
      if (d === 0) queue.push(l.target);
    }
  }
  if (processed < nodes.length) {
    const stuck = nodes.filter((n) => indegree.get(n)! > 0).map((n) => `"${n.name}"`);
    throw new SankeyError(
      `The flows form a loop (involving ${stuck.slice(0, 3).join(", ")}${stuck.length > 3 ? ", …" : ""}). ` +
        "A Sankey diagram needs flows that only move forward."
    );
  }
  const maxDepth = Math.max(0, ...nodes.map((n) => depth.get(n)!));
  // "left": every node sits one column after its furthest source (end nodes stay
  // next to what feeds them). "justify": end nodes are pushed to the last column.
  const justify = opts.align === "justify";
  for (const n of nodes) {
    n.column = justify && !n.sourceLinks.length ? maxDepth : depth.get(n)!;
  }
  const columnCount = maxDepth + 1;

  const columns: SankeyNode[][] = Array.from({ length: columnCount }, () => []);
  for (const n of nodes) columns[n.column].push(n);

  // Optional grouping: boxes with the same key stay together in each column,
  // groups in order of first appearance, data order within a group. The order
  // is then kept fixed while the layout is relaxed.
  let fixedOrder: ((a: SankeyNode, b: SankeyNode) => number) | null = null;
  if (opts.groupKeys) {
    const keys = opts.groupKeys(nodes);
    const rank = new Map<SankeyNode, number>();
    for (const col of columns) {
      const firstSeen = new Map<string, number>();
      for (const n of col) {
        const k = keys.get(n) ?? "";
        if (!firstSeen.has(k)) firstSeen.set(k, n.index);
        rank.set(n, firstSeen.get(k)!);
      }
    }
    fixedOrder = (a, b) => rank.get(a)! - rank.get(b)! || a.index - b.index;
    for (const col of columns) col.sort(fixedOrder);
  }

  // Horizontal positions.
  const kx = columnCount > 1 ? (width - nodeWidth) / (columnCount - 1) : 0;
  for (const n of nodes) {
    n.x0 = columnCount > 1 ? n.column * kx : (width - nodeWidth) / 2;
    n.x1 = n.x0 + nodeWidth;
  }

  // Vertical scale: the fullest column must fit the height.
  const maxRows = Math.max(...columns.map((c) => c.length));
  // Gaps may use at most this share of the height, so boxes keep the rest.
  const maxGapShare = opts.maxGapShare ?? 0.4;
  const padding = Math.min(
    opts.nodePadding ?? 14,
    maxRows > 1 ? (height * maxGapShare) / (maxRows - 1) : 0
  );
  const ky = Math.min(
    ...columns.map((c) => (height - (c.length - 1) * padding) / c.reduce((s, n) => s + n.value, 0))
  );
  for (const l of links) l.width = l.value * ky;

  // Initial placement: stack each column, centered.
  for (const col of columns) {
    let y = 0;
    for (const n of col) {
      n.y0 = y;
      n.y1 = y + n.value * ky;
      y = n.y1 + padding;
    }
    const offset = (height - (y - padding)) / 2;
    for (const n of col) shift(n, offset);
  }

  // Relaxation: pull each node toward the weighted center of its neighbors.
  for (let i = 0; i < iterations; i++) {
    const alpha = Math.pow(0.99, i) * 0.6;
    for (let c = 1; c < columnCount; c++) {
      for (const n of columns[c]) relaxTo(n, n.targetLinks, (l) => l.source, alpha);
      resolveCollisions(columns[c], height, padding, fixedOrder);
    }
    for (let c = columnCount - 2; c >= 0; c--) {
      for (const n of columns[c]) relaxTo(n, n.sourceLinks, (l) => l.target, alpha);
      resolveCollisions(columns[c], height, padding, fixedOrder);
    }
  }

  updateLinkPositions(nodes);

  return { nodes, links, columnCount, width, height };
}

/** Orders links within each node to minimise crossings, then stacks them. Call after moving nodes. */
export function updateLinkPositions(nodes: SankeyNode[]) {
  for (const n of nodes) {
    n.sourceLinks.sort((a, b) => a.target.y0 - b.target.y0 || a.target.index - b.target.index);
    n.targetLinks.sort((a, b) => a.source.y0 - b.source.y0 || a.source.index - b.source.index);
    let y = n.y0;
    for (const l of n.sourceLinks) {
      l.y0 = y + l.width / 2;
      y += l.width;
    }
    y = n.y0;
    for (const l of n.targetLinks) {
      l.y1 = y + l.width / 2;
      y += l.width;
    }
  }
}

/**
 * A node position the user chose, stored as fractions (0–1) of the space the
 * node can move in, so it survives resizing and the wider exported picture.
 */
export interface NodePosition {
  fx: number;
  fy: number;
}

export function getNodePosition(n: SankeyNode, layout: SankeyLayout): NodePosition {
  const freeX = layout.width - (n.x1 - n.x0);
  const freeY = layout.height - (n.y1 - n.y0);
  return { fx: freeX > 0 ? n.x0 / freeX : 0, fy: freeY > 0 ? n.y0 / freeY : 0 };
}

export function applyNodePositions(layout: SankeyLayout, positions: Map<string, NodePosition>) {
  if (!positions.size) return;
  for (const n of layout.nodes) {
    const p = positions.get(n.name);
    if (!p) continue;
    const w = n.x1 - n.x0;
    const h = n.y1 - n.y0;
    n.x0 = clamp01(p.fx) * Math.max(0, layout.width - w);
    n.x1 = n.x0 + w;
    n.y0 = clamp01(p.fy) * Math.max(0, layout.height - h);
    n.y1 = n.y0 + h;
  }
  updateLinkPositions(layout.nodes);
}

const clamp01 = (v: number) => Math.min(1, Math.max(0, v));

function shift(n: SankeyNode, dy: number) {
  n.y0 += dy;
  n.y1 += dy;
}

function center(n: SankeyNode) {
  return (n.y0 + n.y1) / 2;
}

function relaxTo(
  n: SankeyNode,
  ls: SankeyLink[],
  other: (l: SankeyLink) => SankeyNode,
  alpha: number
) {
  if (!ls.length) return;
  let sum = 0;
  let weight = 0;
  for (const l of ls) {
    sum += center(other(l)) * l.value;
    weight += l.value;
  }
  shift(n, (sum / weight - center(n)) * alpha);
}

function resolveCollisions(
  col: SankeyNode[],
  height: number,
  padding: number,
  fixedOrder: ((a: SankeyNode, b: SankeyNode) => number) | null = null
) {
  // Normally boxes may swap places as they settle; with a fixed order they don't.
  col.sort(fixedOrder ?? ((a, b) => a.y0 - b.y0));
  // Push down overlapping nodes.
  let y = 0;
  for (const n of col) {
    const dy = y - n.y0;
    if (dy > 0) shift(n, dy);
    y = n.y1 + padding;
  }
  // Push back up anything that fell off the bottom.
  y = height;
  for (let i = col.length - 1; i >= 0; i--) {
    const n = col[i];
    const dy = n.y1 - y;
    if (dy > 0) shift(n, -dy);
    y = n.y0 - padding;
  }
}

/** Horizontal cubic Bézier through the middle of a link band. */
export function linkPath(l: SankeyLink): string {
  const x0 = l.source.x1;
  const x1 = l.target.x0;
  const xm = (x0 + x1) / 2;
  return `M${x0},${l.y0}C${xm},${l.y0} ${xm},${l.y1} ${x1},${l.y1}`;
}
