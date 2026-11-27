/*
 * Renders a SankeyLayout as SVG. All colors and fonts are written as SVG
 * attributes (not CSS classes) so the same SVG can be serialized and inserted
 * into the worksheet as a picture.
 *
 * In interactive mode, nodes can be dragged with the mouse/touch, moved with the
 * arrow keys when focused, and clicked to select them.
 */

/* global document, setTimeout, clearTimeout, Element, MouseEvent, PointerEvent, KeyboardEvent, SVGElementTagNameMap, SVGSVGElement, SVGPathElement, SVGRectElement, SVGTextElement, SVGGElement, SVGTSpanElement */

import { linkPath, updateLinkPositions } from "./sankey";
import type { SankeyLayout, SankeyLink, SankeyNode } from "./sankey";
import { findGuides, snapPosition, snapStep } from "./snap";
import type { Guide } from "./snap";
import {
  LABEL_GAP,
  LINE_HEIGHT,
  NAME_WEIGHT,
  formatNumber,
  formatPercent,
  labelBlock,
  resolvePlacement,
  textWidth,
  totalFlow,
} from "./labels";

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}
const overlaps = (a: Box, b: Box) =>
  a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;
import type { LabelBlock, LabelSettings, NodeLabelOverride } from "./labels";
import type { NodeRole } from "./finance";

export type LinkColorMode = "source" | "target" | "neutral";

export interface ChartTheme {
  surface: string;
  textPrimary: string;
  textSecondary: string;
  neutral: string;
  /** Alignment guides shown while moving nodes (matches the pane's accent color). */
  guide: string;
  /** Amber, for marking values that are unaccounted for. */
  warning: string;
  /** "Money flow" scheme colors, sampled from the reference image (Example_flow.png). */
  roles: Record<NodeRole, string>;
  series: string[];
}

// Validated categorical palette (fixed order; never cycled).
export const LIGHT_THEME: ChartTheme = {
  surface: "#ffffff",
  textPrimary: "#000000",
  textSecondary: "#52514e",
  neutral: "#898781",
  guide: "#0f6cbd",
  warning: "#d97706",
  roles: {
    income: "#00a34c",
    spending: "#d1003f",
    savings: "#366b87",
    remainder: "#b88400",
    account: "#696969",
  },
  series: ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"],
};

// Dark mode: the same hues, lifted so they stand out on the dark surface.
export const DARK_THEME: ChartTheme = {
  surface: "#1a1a19",
  textPrimary: "#f0efec",
  textSecondary: "#c3c2b7",
  neutral: "#898781",
  guide: "#479ef5",
  warning: "#f59e0b",
  roles: {
    income: "#1fbf63",
    spending: "#f0466f",
    savings: "#5d9cc0",
    remainder: "#d6a21a",
    account: "#8c8c8c",
  },
  series: ["#3987e5", "#d95926", "#199e70", "#c98500", "#d55181", "#0ca30c", "#9085e9", "#e66767"],
};

/** Relative luminance (0 = black, 1 = white) of a #rrggbb color. */
export function luminance(hex: string): number {
  const h = hex.replace("#", "");
  const full = h.length === 3 ? [...h].map((x) => x + x).join("") : h;
  const [r, g, b] = [0, 2, 4].map((i) => {
    const c = parseInt(full.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * A theme for a chosen background color: dark backgrounds get the dark theme's
 * brighter box colors and light text; light backgrounds the light theme's.
 */
export function themeForBackground(background: string): ChartTheme {
  // Below this, white text reads better than black (the point where contrast is equal).
  const dark = luminance(background) < 0.18;
  return { ...(dark ? DARK_THEME : LIGHT_THEME), surface: background };
}

/** Flows are solid tints of their color: this much of the background is mixed in. */
const LINK_TINT = 0.55;

/** Mixes a #rrggbb (or #rgb) color toward another by `t` (0 = a, 1 = b). */
export function mixColors(a: string, b: string, t: number): string {
  const rgb = (c: string) => {
    const h = c.replace("#", "");
    const full = h.length === 3 ? [...h].map((x) => x + x).join("") : h;
    return [0, 2, 4].map((i) => parseInt(full.slice(i, i + 2), 16));
  };
  const [ca, cb] = [rgb(a), rgb(b)];
  if (ca.some(isNaN) || cb.some(isNaN)) return a;
  return (
    "#" +
    ca
      .map((v, i) => Math.round(v * (1 - t) + cb[i] * t))
      .map((v) => v.toString(16).padStart(2, "0"))
      .join("")
  );
}

export interface InteractionHandlers {
  /** Hover tooltip; null hides it. Coordinates are relative to the SVG. */
  onHover: (info: { html: string; x: number; y: number } | null) => void;
  /** A node was clicked (or focused and Enter pressed). */
  onNodeSelect: (node: SankeyNode) => void;
  /** A node finished moving (drag released or arrow key pressed). */
  onNodeMoved: (node: SankeyNode, layout: SankeyLayout) => void;
  /** A label was dragged; offset is relative to its placement position. */
  onLabelMoved: (node: SankeyNode, dx: number, dy: number) => void;
  /** Empty chart background clicked (used to clear the selection). */
  onBackgroundClick?: () => void;
  /** Name of the node to outline as selected. */
  selected: string | null;
  /** Put keyboard focus back on the selected node after rendering. */
  focusSelected?: boolean;
}

export interface RenderOptions {
  theme: ChartTheme;
  linkColor: LinkColorMode;
  labels: LabelSettings;
  /** Per-node label text, placement and offset, by node name. */
  labelOverrides: Map<string, NodeLabelOverride>;
  /** User-chosen node colors by node name; overrides the palette. */
  colors: Map<string, string>;
  /**
   * Present for the "Money flow" scheme: each node's color group (name and
   * color). A node missing from the map is ungrouped and drawn in neutral gray.
   */
  groups?: Map<SankeyNode, NodeGroup>;
  margin: { top: number; right: number; bottom: number; left: number };
  /** Leave the background out (flows and label pills still use theme.surface). */
  transparent?: boolean;
  /** Title (and subtitle) drawn across the top. The top margin must include titleHeight(). */
  title?: ChartTitle;
  /** Omit for a static render (e.g. the picture inserted into the sheet). */
  interactive?: InteractionHandlers;
}

const SVG_NS = "http://www.w3.org/2000/svg";
/** Labels dragged within this distance of their anchor snap back onto it. */
const LABEL_SNAP = 6;
/** Flows are solid colors; hovering fades the unrelated ones. */
const LINK_OPACITY = 1;
const LINK_OPACITY_HOVER = 1;
const LINK_OPACITY_FADED = 0.25;
/** Translucent background behind labels, so they stay readable over flows. */
const PILL_OPACITY = 0.55;
const PILL_PAD_X = 4;
const PILL_PAD_Y = 2;
const DRAG_THRESHOLD = 3;

export const formatValue = (v: number) =>
  v.toLocaleString(undefined, { maximumFractionDigits: Math.abs(v) < 10 ? 2 : 0 });

/** A color group a node belongs to ("Money flow" scheme). */
export interface NodeGroup {
  name: string;
  color: string;
}

/**
 * Default color for a node. With groups ("Money flow" scheme): its group's
 * color, or neutral gray if it has none. Without: palette slot by order of
 * appearance; past the eighth, neutral.
 */
export function defaultNodeColor(
  n: SankeyNode,
  theme: ChartTheme,
  groups?: Map<SankeyNode, NodeGroup>
): string {
  if (groups) return groups.get(n)?.color ?? theme.neutral;
  return n.index < theme.series.length ? theme.series[n.index] : theme.neutral;
}

/**
 * A user color choice: "slot:N" (palette slot N, follows light/dark theme),
 * "role:<category>" (a Money flow color), "neutral", or a literal "#rrggbb"
 * from the custom picker.
 */
export type ColorChoice = string;

export function resolveColor(choice: ColorChoice, theme: ChartTheme): string {
  if (choice === "neutral") return theme.neutral;
  // "role:spending" etc. — a Money flow category color (follows the theme).
  const role = /^role:(\w+)$/.exec(choice);
  if (role) return theme.roles[role[1] as NodeRole] ?? theme.neutral;
  const slot = /^slot:(\d+)$/.exec(choice);
  if (slot) return theme.series[Number(slot[1])] ?? theme.neutral;
  return choice;
}

export function nodeColor(
  n: SankeyNode,
  theme: ChartTheme,
  colors: Map<string, ColorChoice>,
  groups?: Map<SankeyNode, NodeGroup>
): string {
  const choice = colors.get(n.name);
  return choice ? resolveColor(choice, theme) : defaultNodeColor(n, theme, groups);
}

function el<K extends keyof SVGElementTagNameMap>(
  tag: K,
  attrs: Record<string, string | number>,
  parent?: Element
): SVGElementTagNameMap[K] {
  const e = document.createElementNS(SVG_NS, tag);
  setAttrs(e, attrs);
  if (parent) parent.appendChild(e);
  return e;
}

function setAttrs(e: Element, attrs: Record<string, string | number>) {
  for (const k of Object.keys(attrs)) e.setAttribute(k, String(attrs[k]));
}

const escapeHtml = (s: string) =>
  s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]!);

interface NodeEls {
  rect: SVGRectElement;
  /** Group holding the label's background pill and text. */
  labelGroup: SVGGElement;
  pill: SVGRectElement;
  label: SVGTextElement;
  nameSpan: SVGTSpanElement;
  valueSpan: SVGTSpanElement | null;
  block: LabelBlock;
  /** Where the label's text block was last placed (for hit testing and guides). */
  box?: { x: number; y: number; w: number; h: number };
  hit?: SVGRectElement;
  labelHit?: SVGRectElement;
}

/** A title (and optional subtitle) across the top of the chart. */
export interface ChartTitle {
  text: string;
  subtitle: string;
  /** Title size in px. */
  size: number;
  /** Subtitle size in px; null = automatic (about 60% of the title size). */
  subtitleSize: number | null;
  bold: boolean;
  italic: boolean;
  /** CSS font-family stack; null = the same font as the labels. */
  fontFamily: string | null;
  align: "left" | "center" | "right";
  /** null = theme text color. */
  color: string | null;
}

const TITLE_PAD_TOP = 12;
const TITLE_GAP_BELOW = 10;
const TITLE_SIDE_PAD = 14;
const subtitleSize = (t: ChartTitle) => t.subtitleSize ?? Math.max(11, Math.round(t.size * 0.6));
const hasTitle = (t: ChartTitle | undefined): t is ChartTitle =>
  !!t && (t.text.trim() !== "" || t.subtitle.trim() !== "");

/** Height the title needs at the top of the chart (0 when there's no title). */
export function titleHeight(t: ChartTitle | undefined): number {
  if (!hasTitle(t)) return 0;
  let h = TITLE_PAD_TOP;
  if (t.text.trim()) h += t.size * 1.25;
  if (t.subtitle.trim()) h += subtitleSize(t) * 1.45;
  return Math.ceil(h + TITLE_GAP_BELOW);
}

/** Lets the pane control a drawn chart after rendering. */
export interface ChartHandle {
  /** Marks (or unmarks) every box whose money in and out don't match. */
  highlightImbalance(on: boolean): void;
}

/** Ignores floating-point dust when comparing money in and out. */
function imbalanceOf(n: SankeyNode): number {
  const d = n.inValue - n.outValue;
  return Math.abs(d) < Math.max(n.inValue, n.outValue, 1) * 1e-9 ? 0 : d;
}

export function renderSankey(
  svg: SVGSVGElement,
  layout: SankeyLayout,
  opts: RenderOptions
): ChartHandle {
  const { theme, margin, labels } = opts;
  const total = totalFlow(layout);
  const labelOffset = new Map<SankeyNode, { dx: number; dy: number }>();
  for (const n of layout.nodes) {
    const o = opts.labelOverrides.get(n.name);
    labelOffset.set(n, { dx: o?.dx ?? 0, dy: o?.dy ?? 0 });
  }
  const totalW = layout.width + margin.left + margin.right;
  const totalH = layout.height + margin.top + margin.bottom;

  while (svg.firstChild) svg.removeChild(svg.firstChild);
  setAttrs(svg, {
    xmlns: SVG_NS,
    width: totalW,
    height: totalH,
    viewBox: `0 0 ${totalW} ${totalH}`,
    "font-family": labels.fontFamily,
  });

  const background = opts.transparent
    ? null
    : el("rect", { x: 0, y: 0, width: totalW, height: totalH, fill: theme.surface }, svg);

  // Clicking empty background (not a box, label or flow) clears the selection.
  // Assigned as a property so redraws replace it rather than stacking listeners.
  const onBackground = opts.interactive?.onBackgroundClick;
  svg.onclick = onBackground
    ? (e: MouseEvent) => {
        if (e.target === svg || e.target === background) onBackground();
      }
    : null;
  // Title and subtitle, in the space the top margin keeps for them.
  if (hasTitle(opts.title)) {
    const t = opts.title;
    const x =
      t.align === "left"
        ? TITLE_SIDE_PAD
        : t.align === "right"
          ? totalW - TITLE_SIDE_PAD
          : totalW / 2;
    const anchor = t.align === "left" ? "start" : t.align === "right" ? "end" : "middle";
    let y = TITLE_PAD_TOP;
    if (t.text.trim()) {
      y += t.size * 1.25;
      const title = el(
        "text",
        {
          x,
          y: y - t.size * 0.3, // baseline inside the line box
          "text-anchor": anchor,
          "font-family": t.fontFamily ?? labels.fontFamily,
          "font-size": t.size,
          "font-weight": t.bold ? 700 : 400,
          "font-style": t.italic ? "italic" : "normal",
          fill: t.color ?? theme.textPrimary,
        },
        svg
      );
      title.textContent = t.text.trim();
    }
    if (t.subtitle.trim()) {
      const s = subtitleSize(t);
      y += s * 1.45;
      const sub = el(
        "text",
        {
          x,
          y: y - s * 0.4,
          "text-anchor": anchor,
          "font-family": t.fontFamily ?? labels.fontFamily,
          "font-size": s,
          "font-style": t.italic ? "italic" : "normal",
          fill: t.color ?? theme.textSecondary,
          ...(t.color ? { "fill-opacity": 0.8 } : {}),
        },
        svg
      );
      sub.textContent = t.subtitle.trim();
    }
  }

  const root = el("g", { transform: `translate(${margin.left},${margin.top})` }, svg);

  const colorOf = (n: SankeyNode) => nodeColor(n, theme, opts.colors, opts.groups);
  // Flows are opaque tints (the node color mixed with the background), so
  // overlapping flows don't blend into muddy colors.
  const linkColor = (l: SankeyLink) =>
    mixColors(
      opts.linkColor === "source"
        ? colorOf(l.source)
        : opts.linkColor === "target"
          ? colorOf(l.target)
          : theme.neutral,
      theme.surface,
      LINK_TINT
    );

  // Create elements once; positions are applied by place() so dragging can update them in place.
  const linkLayer = el("g", { fill: "none" }, root);
  const linkEls = new Map<SankeyLink, SVGPathElement>();
  for (const l of layout.links) {
    linkEls.set(
      l,
      el(
        "path",
        {
          stroke: linkColor(l),
          "stroke-opacity": LINK_OPACITY,
          "stroke-width": Math.max(1, l.width),
        },
        linkLayer
      )
    );
  }

  const nodeLayer = el("g", {}, root);
  // Unaccounted-for markers (shown on request), above boxes and below labels.
  const imbalanceLayer = el("g", { "pointer-events": "none" }, root);
  const labelLayer = el(
    "g",
    { "font-size": labels.fontSize, "font-family": labels.fontFamily },
    root
  );
  const nodeEls = new Map<SankeyNode, NodeEls>();
  const selected = opts.interactive?.selected ?? null;
  for (const n of layout.nodes) {
    const isSelected = n.name === selected;
    const rect = el(
      "rect",
      {
        width: n.x1 - n.x0,
        height: Math.max(1, n.y1 - n.y0),
        fill: colorOf(n),
        ...(isSelected ? { stroke: theme.textPrimary, "stroke-width": 2 } : {}),
      },
      nodeLayer
    );
    // Label: a translucent rounded "pill" behind bold name + value.
    const block = labelBlock(n, total, labels, opts.labelOverrides.get(n.name));
    const labelGroup = el("g", { "pointer-events": "none" }, labelLayer);
    const pill = el(
      "rect",
      { rx: 4, fill: theme.surface, "fill-opacity": PILL_OPACITY },
      labelGroup
    );
    const label = el("text", { fill: labels.color ?? theme.textPrimary }, labelGroup);
    const nameSpan = el("tspan", { "font-weight": labels.boldNames ? NAME_WEIGHT : 400 }, label);
    nameSpan.textContent = block.name;
    let valueSpan: SVGTSpanElement | null = null;
    if (block.value) {
      valueSpan = el("tspan", { "font-weight": 400 }, label);
      valueSpan.textContent = block.value;
    }
    nodeEls.set(n, { rect, labelGroup, pill, label, nameSpan, valueSpan, block });
  }

  // The unaccounted-for amounts go above every label, so none of them hide it.
  const imbalanceTextLayer = el("g", { "pointer-events": "none" }, root);

  // Alignment guides (only while moving), then the interactive layer: invisible,
  // larger hit targets on top so thin links/nodes are easy to grab.
  const guideLayer = opts.interactive ? el("g", { "pointer-events": "none" }, root) : null;
  const hitLayer = opts.interactive ? el("g", { fill: "none" }, root) : null;
  const linkHits = new Map<SankeyLink, SVGPathElement>();

  /** Top-left of a label's text block before any drag offset, and its text alignment. */
  const labelAnchor = (n: SankeyNode, b: LabelBlock) => {
    const p = resolvePlacement(n, layout, labels, opts.labelOverrides.get(n.name));
    const cx = (n.x0 + n.x1) / 2;
    const cy = (n.y0 + n.y1) / 2;
    switch (p) {
      case "right":
        return { p, x: n.x1 + LABEL_GAP, y: cy - b.height / 2, align: "start" };
      case "left":
        return { p, x: n.x0 - LABEL_GAP - b.width, y: cy - b.height / 2, align: "end" };
      case "above":
        return { p, x: cx - b.width / 2, y: n.y0 - LABEL_GAP - b.height, align: "middle" };
      case "below":
        return { p, x: cx - b.width / 2, y: n.y1 + LABEL_GAP, align: "middle" };
      default:
        return { p, x: cx - b.width / 2, y: cy - b.height / 2, align: "middle" };
    }
  };

  const placeLabel = (n: SankeyNode, e: NodeEls) => {
    const b = e.block;
    const a = labelAnchor(n, b);
    const off = labelOffset.get(n)!;
    if (a.p === "hidden") {
      e.labelGroup.setAttribute("visibility", "hidden");
      if (e.labelHit) e.labelHit.setAttribute("visibility", "hidden");
      e.box = undefined;
      return;
    }
    e.labelGroup.removeAttribute("visibility");
    const x = a.x + off.dx;
    const y = a.y + off.dy;
    e.box = { x, y, w: b.width, h: b.height };

    // Text lines: name, then value (below it, or after it on the same line).
    const lineH = labels.fontSize * LINE_HEIGHT;
    const tx = a.align === "start" ? x : a.align === "end" ? x + b.width : x + b.width / 2;
    const baseline = y + lineH / 2 + labels.fontSize * 0.35;
    setAttrs(e.label, { "text-anchor": a.align });
    setAttrs(e.nameSpan, { x: tx, y: baseline });
    if (e.valueSpan) {
      if (b.lines === 2) setAttrs(e.valueSpan, { x: tx, y: baseline + lineH, dx: 0 });
      else {
        // Inline: continue after the name (x/y unset so it flows on).
        e.valueSpan.removeAttribute("x");
        e.valueSpan.removeAttribute("y");
        setAttrs(e.valueSpan, { dx: labels.fontSize * 0.4 });
      }
    }
    setAttrs(e.pill, {
      x: x - PILL_PAD_X,
      y: y - PILL_PAD_Y,
      width: b.width + PILL_PAD_X * 2,
      height: b.height + PILL_PAD_Y * 2,
    });
    if (e.labelHit) {
      e.labelHit.removeAttribute("visibility");
      setAttrs(e.labelHit, {
        x: x - PILL_PAD_X,
        y: y - PILL_PAD_Y,
        width: b.width + PILL_PAD_X * 2,
        height: b.height + PILL_PAD_Y * 2,
      });
    }
  };

  /**
   * Positions everything. `markImbalance: false` (while a box is dragged)
   * hides the unaccounted-for markers instead of rebuilding them on every
   * pointer move; they're rebuilt once the drag ends.
   */
  const place = (markImbalance = true) => {
    for (const [l, p] of linkEls) p.setAttribute("d", linkPath(l));
    for (const [l, p] of linkHits) p.setAttribute("d", linkPath(l));
    for (const [n, e] of nodeEls) {
      const h = n.y1 - n.y0;
      setAttrs(e.rect, { x: n.x0, y: n.y0 });
      placeLabel(n, e);
      if (e.hit) {
        setAttrs(e.hit, {
          x: n.x0 - 4,
          y: n.y0 - Math.max(0, (14 - h) / 2),
          width: n.x1 - n.x0 + 8,
          height: Math.max(14, h),
        });
      }
    }
    // Boxes moved: redraw the unaccounted-for markers around their new places.
    if (!imbalanceOn) return;
    if (markImbalance) highlightImbalance(true);
    else {
      imbalanceLayer.setAttribute("visibility", "hidden");
      imbalanceTextLayer.setAttribute("visibility", "hidden");
    }
  };

  /**
   * For each box whose in and out differ: a dashed amber outline, an amber
   * hatched band where the missing value is (right side if money isn't passed
   * on, left side if more goes out than came in), and the amount.
   */
  let imbalanceOn = false;
  const highlightImbalance = (on: boolean) => {
    imbalanceOn = on;
    for (const layer of [imbalanceLayer, imbalanceTextLayer]) {
      while (layer.firstChild) layer.removeChild(layer.firstChild);
      layer.removeAttribute("visibility");
    }
    linkEls.forEach((p) =>
      p.setAttribute("stroke-opacity", String(on ? LINK_OPACITY_FADED * 1.6 : LINK_OPACITY))
    );
    if (!on) return;

    // Where the amounts mustn't go: on a box, a label, or another amount.
    const avoid: Box[] = [];
    for (const [n, e] of nodeEls) {
      avoid.push({ x: n.x0 - 2, y: n.y0 - 2, w: n.x1 - n.x0 + 4, h: n.y1 - n.y0 + 4 });
      if (e.box)
        avoid.push({
          x: e.box.x - PILL_PAD_X,
          y: e.box.y - PILL_PAD_Y,
          w: e.box.w + PILL_PAD_X * 2,
          h: e.box.h + PILL_PAD_Y * 2,
        });
    }

    const amber = theme.warning;
    const defs = el("defs", {}, imbalanceLayer);
    const pattern = el(
      "pattern",
      {
        id: "unaccounted-hatch",
        width: 6,
        height: 6,
        patternUnits: "userSpaceOnUse",
        patternTransform: "rotate(45)",
      },
      defs
    );
    el("rect", { width: 6, height: 6, fill: amber, "fill-opacity": 0.25 }, pattern);
    el("line", { x1: 0, y1: 0, x2: 0, y2: 6, stroke: amber, "stroke-width": 3 }, pattern);

    const stub = 18; // width of the marked gap
    for (const n of layout.nodes) {
      const diff = imbalanceOf(n);
      if (!diff || !n.targetLinks.length || !n.sourceLinks.length) continue;
      // Pixels per unit of value, from the box's own height.
      const ky = n.value ? (n.y1 - n.y0) / n.value : 0;
      const missingH = Math.max(2, Math.abs(diff) * ky);
      // Flows are stacked from the top of the box, so the gap is at the bottom
      // of the side that's short.
      const passedOnRight = diff > 0;
      const y = n.y0 + (passedOnRight ? n.outValue : n.inValue) * ky;
      const x = passedOnRight ? n.x1 : n.x0 - stub;

      el(
        "rect",
        {
          x: n.x0 - 3,
          y: n.y0 - 3,
          width: n.x1 - n.x0 + 6,
          height: n.y1 - n.y0 + 6,
          fill: "none",
          stroke: amber,
          "stroke-width": 2,
          "stroke-dasharray": "4 3",
        },
        imbalanceLayer
      );
      el(
        "rect",
        {
          x,
          y,
          width: stub,
          height: missingH,
          fill: "url(#unaccounted-hatch)",
          stroke: amber,
          "stroke-width": 1.5,
        },
        imbalanceLayer
      );
      // The amount, on an opaque amber-edged pill above every label, placed
      // where it doesn't cover a label (or another amount).
      const amount = formatNumber(Math.abs(diff), labels);
      const str = passedOnRight ? `${amount} not passed on` : `${amount} more out than in`;
      const fs = Math.max(10, labels.fontSize - 1);
      const w = textWidth(str, { ...labels, fontSize: fs, boldNames: true }, true) + 12;
      const h = fs * 1.2 + 6;
      const midY = y + missingH / 2;
      // Outward from the box: right of the gap if it's on the right, else left.
      const sideX = passedOnRight ? x + stub + 4 : x - 4 - w;
      const nearX = passedOnRight ? x : x + stub - w; // lined up with the gap
      const candidates: Box[] = [{ x: sideX, y: midY - h / 2, w, h }];
      for (let k = 1; k <= 6; k++) {
        const step = (h + 4) * k;
        candidates.push(
          { x: sideX, y: midY - h / 2 + step, w, h },
          { x: sideX, y: midY - h / 2 - step, w, h },
          { x: nearX, y: y + missingH + 4 + (k - 1) * (h + 4), w, h },
          { x: nearX, y: y - 4 - h - (k - 1) * (h + 4), w, h }
        );
      }
      const free = (c: Box) =>
        c.x >= -margin.left &&
        c.x + c.w <= layout.width + margin.right &&
        c.y >= -4 && // not up into the title
        c.y + c.h <= layout.height + margin.bottom &&
        !avoid.some((b) => overlaps(c, b));
      const spot = candidates.find(free) ?? candidates[0];
      avoid.push(spot);

      // A thin leader back to the gap when the amount had to move away from it.
      if (spot !== candidates[0]) {
        const gx = passedOnRight ? x + stub : x;
        const px = Math.min(Math.max(gx, spot.x), spot.x + spot.w);
        const py = Math.min(Math.max(midY, spot.y), spot.y + spot.h);
        el(
          "line",
          { x1: gx, y1: midY, x2: px, y2: py, stroke: amber, "stroke-width": 1 },
          imbalanceTextLayer
        );
      }
      el(
        "rect",
        {
          x: spot.x,
          y: spot.y,
          width: spot.w,
          height: spot.h,
          rx: 4,
          fill: theme.surface,
          stroke: amber,
          "stroke-width": 1,
        },
        imbalanceTextLayer
      );
      const text = el(
        "text",
        {
          x: spot.x + spot.w / 2,
          y: spot.y + spot.h / 2,
          dy: "0.35em",
          "text-anchor": "middle",
          "font-size": fs,
          "font-family": labels.fontFamily,
          "font-weight": NAME_WEIGHT,
          fill: amber,
        },
        imbalanceTextLayer
      );
      text.textContent = str;
    }
  };
  const handle: ChartHandle = { highlightImbalance };

  if (!opts.interactive || !hitLayer) {
    place();
    return handle;
  }
  const ui = opts.interactive;

  const setLinkOpacity = (active: (l: SankeyLink) => boolean) =>
    linkEls.forEach((p, l) =>
      p.setAttribute("stroke-opacity", String(active(l) ? LINK_OPACITY_HOVER : LINK_OPACITY_FADED))
    );
  const resetLinks = () =>
    linkEls.forEach((p) => p.setAttribute("stroke-opacity", String(LINK_OPACITY)));
  const pointer = (e: PointerEvent) => {
    const box = svg.getBoundingClientRect();
    return { x: e.clientX - box.left, y: e.clientY - box.top };
  };
  /** How magnified the chart is on screen (the + / − buttons), so drags follow the pointer. */
  const zoomOf = () => {
    const w = svg.getBoundingClientRect().width;
    return w > 0 ? w / totalW : 1;
  };

  let dragging: SankeyNode | null = null;
  const labelName = (n: SankeyNode) => opts.labelOverrides.get(n.name)?.text || n.name;

  for (const l of layout.links) {
    const hit = el(
      "path",
      {
        stroke: "transparent",
        "stroke-width": Math.max(10, l.width),
        "pointer-events": "stroke",
      },
      hitLayer
    );
    linkHits.set(l, hit);
    const ofSource = l.source.outValue ? l.value / l.source.outValue : 0;
    const html =
      `<div class="tt-title">${escapeHtml(labelName(l.source))} → ${escapeHtml(labelName(l.target))}</div>` +
      `<div class="tt-value">${escapeHtml(formatNumber(l.value, labels))}</div>` +
      `<div class="tt-sub">${formatPercent(ofSource, labels)} of ${escapeHtml(labelName(l.source))}</div>` +
      `<div class="tt-sub">${formatPercent(total ? l.value / total : 0, labels)} of total</div>`;
    hit.addEventListener("pointerenter", () => {
      if (!dragging) setLinkOpacity((x) => x === l);
    });
    hit.addEventListener("pointermove", (e) => {
      if (!dragging) ui.onHover({ html, ...pointer(e) });
    });
    hit.addEventListener("pointerleave", () => {
      if (dragging) return;
      resetLinks();
      ui.onHover(null);
    });
  }

  /** Moves a box. `dragging`: mid-drag, so skip the imbalance markers (see place). */
  const moveNodeTo = (n: SankeyNode, x0: number, y0: number, dragging = false) => {
    const w = n.x1 - n.x0;
    const h = n.y1 - n.y0;
    n.x0 = Math.min(Math.max(0, x0), Math.max(0, layout.width - w));
    n.x1 = n.x0 + w;
    n.y0 = Math.min(Math.max(0, y0), Math.max(0, layout.height - h));
    n.y1 = n.y0 + h;
    updateLinkPositions(layout.nodes);
    place(!dragging);
  };

  // Dotted guide lines with a snap point on each aligned node.
  let guideTimer: ReturnType<typeof setTimeout> | undefined;
  const drawGuides = (guides: Guide[]) => {
    if (guideTimer) clearTimeout(guideTimer);
    if (!guideLayer) return;
    while (guideLayer.firstChild) guideLayer.removeChild(guideLayer.firstChild);
    for (const g of guides) {
      const [x1, y1, x2, y2] =
        g.axis === "x" ? [g.pos, g.from, g.pos, g.to] : [g.from, g.pos, g.to, g.pos];
      el(
        "line",
        {
          x1,
          y1,
          x2,
          y2,
          stroke: theme.guide,
          "stroke-width": 1.5,
          "stroke-dasharray": "0.5 4",
          "stroke-linecap": "round",
        },
        guideLayer
      );
      for (const p of g.points) {
        el(
          "circle",
          {
            cx: g.axis === "x" ? g.pos : p,
            cy: g.axis === "x" ? p : g.pos,
            r: 3,
            fill: theme.guide,
            stroke: theme.surface,
            "stroke-width": 1.5,
          },
          guideLayer
        );
      }
    }
  };
  const clearGuides = () => drawGuides([]);
  /** Keyboard moves show guides briefly, since there's no "release" to hide them. */
  const flashGuides = (guides: Guide[]) => {
    drawGuides(guides);
    if (guides.length) guideTimer = setTimeout(clearGuides, 1200);
  };

  for (const n of layout.nodes) {
    const num = (v: number) => escapeHtml(formatNumber(v, labels));
    const parts = [
      `<div class="tt-title">${escapeHtml(labelName(n))}</div>`,
      `<div class="tt-value">${num(n.value)}</div>`,
      `<div class="tt-sub">${formatPercent(total ? n.value / total : 0, labels)} of total</div>`,
    ];
    const group = opts.groups?.get(n);
    if (group && !opts.colors.has(n.name)) {
      parts.push(`<div class="tt-sub">Group: ${escapeHtml(group.name)}</div>`);
    }
    if (n.targetLinks.length) parts.push(`<div class="tt-sub">In: ${num(n.inValue)}</div>`);
    if (n.sourceLinks.length) parts.push(`<div class="tt-sub">Out: ${num(n.outValue)}</div>`);
    parts.push(`<div class="tt-hint">Drag to move · click to edit</div>`);
    const html = parts.join("");

    const hit = el(
      "rect",
      {
        fill: "transparent",
        tabindex: 0,
        role: "button",
        "aria-label": `${labelName(n)}, ${formatNumber(n.value, labels)}. Arrow keys move it; Enter selects it.`,
        style: "cursor: grab; outline: none;",
      },
      hitLayer
    );
    nodeEls.get(n)!.hit = hit;
    const highlight = () => setLinkOpacity((x) => x.source === n || x.target === n);

    let start: { px: number; py: number; x0: number; y0: number; moved: boolean } | null = null;

    hit.addEventListener("pointerenter", () => {
      if (!dragging) highlight();
    });
    hit.addEventListener("pointerleave", () => {
      if (dragging) return;
      resetLinks();
      ui.onHover(null);
    });
    hit.addEventListener("pointerdown", (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      try {
        hit.setPointerCapture(e.pointerId);
      } catch {
        // Not an active pointer (e.g. synthetic event); dragging still works while over the node.
      }
      start = { px: e.clientX, py: e.clientY, x0: n.x0, y0: n.y0, moved: false };
    });
    hit.addEventListener("pointermove", (e) => {
      if (!start) {
        if (!dragging) ui.onHover({ html, ...pointer(e) });
        return;
      }
      // Pointer movement on screen, in chart units (allowing for zoom).
      const z = zoomOf();
      const dx = (e.clientX - start.px) / z;
      const dy = (e.clientY - start.py) / z;
      if (!start.moved && Math.hypot(dx * z, dy * z) < DRAG_THRESHOLD) return;
      if (!start.moved) {
        start.moved = true;
        dragging = n;
        ui.onHover(null);
        hit.style.cursor = "grabbing";
        highlight();
      }
      // Snap to nearby alignments unless Alt (or Ctrl/Cmd) is held for free positioning.
      const free = e.altKey || e.ctrlKey || e.metaKey;
      const pos = free
        ? { x0: start.x0 + dx, y0: start.y0 + dy }
        : snapPosition(n, layout, start.x0 + dx, start.y0 + dy);
      moveNodeTo(n, pos.x0, pos.y0, true);
      drawGuides(findGuides(n, layout));
    });
    const endDrag = (e: PointerEvent, cancelled: boolean) => {
      if (!start) return;
      const moved = start.moved;
      start = null;
      dragging = null;
      hit.style.cursor = "grab";
      clearGuides();
      if (hit.hasPointerCapture(e.pointerId)) hit.releasePointerCapture(e.pointerId);
      // The markers were hidden during the drag: put them back around the new place.
      if (moved && imbalanceOn) highlightImbalance(true);
      if (moved) ui.onNodeMoved(n, layout);
      else if (!cancelled) ui.onNodeSelect(n);
    };
    hit.addEventListener("pointerup", (e) => endDrag(e, false));
    hit.addEventListener("pointercancel", (e) => endDrag(e, true));

    // Keyboard: arrows move (Shift = bigger steps), Enter/Space selects.
    hit.addEventListener("focus", highlight);
    hit.addEventListener("blur", () => {
      resetLinks();
      clearGuides();
    });
    hit.addEventListener("keydown", (e: KeyboardEvent) => {
      const step = e.shiftKey ? 20 : 4;
      const delta: Record<string, [number, number]> = {
        ArrowUp: [0, -step],
        ArrowDown: [0, step],
        ArrowLeft: [-step, 0],
        ArrowRight: [step, 0],
      };
      if (delta[e.key]) {
        e.preventDefault();
        const [dx, dy] = delta[e.key];
        // Stop at any alignment on the way (Alt/Ctrl/Cmd moves freely).
        const free = e.altKey || e.ctrlKey || e.metaKey;
        const x0 = dx && !free ? snapStep(n, layout, "x", n.x0, n.x0 + dx) : n.x0 + dx;
        const y0 = dy && !free ? snapStep(n, layout, "y", n.y0, n.y0 + dy) : n.y0 + dy;
        moveNodeTo(n, x0, y0);
        flashGuides(findGuides(n, layout));
        ui.onNodeMoved(n, layout);
      } else if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        ui.onNodeSelect(n);
      }
    });
  }

  // Label hit targets go on top of everything so labels can be dragged even over links.
  const labelHitLayer = el("g", {}, root);
  for (const n of layout.nodes) {
    const e = nodeEls.get(n)!;
    const hit = el("rect", { fill: "transparent", style: "cursor: move;" }, labelHitLayer);
    e.labelHit = hit;
    let start: { px: number; py: number; dx: number; dy: number; moved: boolean } | null = null;

    hit.addEventListener("pointerdown", (ev) => {
      if (ev.button !== 0) return;
      ev.preventDefault();
      try {
        hit.setPointerCapture(ev.pointerId);
      } catch {
        // synthetic pointer
      }
      const off = labelOffset.get(n)!;
      start = { px: ev.clientX, py: ev.clientY, dx: off.dx, dy: off.dy, moved: false };
      ui.onHover(null);
    });
    hit.addEventListener("pointermove", (ev) => {
      if (!start) return;
      // Pointer movement on screen, in chart units (allowing for zoom).
      const z = zoomOf();
      const mx = (ev.clientX - start.px) / z;
      const my = (ev.clientY - start.py) / z;
      if (!start.moved && Math.hypot(mx * z, my * z) < DRAG_THRESHOLD) return;
      start.moved = true;
      dragging = n;
      let dx = start.dx + mx;
      let dy = start.dy + my;
      // Snap back into line with the node unless Alt/Ctrl/Cmd is held.
      const free = ev.altKey || ev.ctrlKey || ev.metaKey;
      if (!free && Math.abs(dx) <= LABEL_SNAP) dx = 0;
      if (!free && Math.abs(dy) <= LABEL_SNAP) dy = 0;
      labelOffset.set(n, { dx, dy });
      placeLabel(n, e);
      drawGuides(labelGuides(n, e, dx, dy));
    });
    const end = (ev: PointerEvent, cancelled: boolean) => {
      if (!start) return;
      const moved = start.moved;
      start = null;
      dragging = null;
      clearGuides();
      if (hit.hasPointerCapture(ev.pointerId)) hit.releasePointerCapture(ev.pointerId);
      const off = labelOffset.get(n)!;
      if (moved) ui.onLabelMoved(n, off.dx, off.dy);
      else if (!cancelled) ui.onNodeSelect(n);
    };
    hit.addEventListener("pointerup", (ev) => end(ev, false));
    hit.addEventListener("pointercancel", (ev) => end(ev, true));
  }

  /** Dotted guide from the node to its label when the label is lined up with it. */
  function labelGuides(n: SankeyNode, e: NodeEls, dx: number, dy: number): Guide[] {
    const a = labelAnchor(n, e.block);
    if (!e.box) return [];
    const b = { x: e.box.x, y: e.box.y, width: e.box.w, height: e.box.h };
    const cx = (n.x0 + n.x1) / 2;
    const cy = (n.y0 + n.y1) / 2;
    const sideways = a.p === "left" || a.p === "right";
    if (sideways && dy === 0) {
      const lx = a.p === "right" ? b.x : b.x + b.width;
      return [
        { axis: "y", pos: cy, from: Math.min(cx, lx), to: Math.max(cx, lx), points: [cx, lx] },
      ];
    }
    if (!sideways && dx === 0) {
      const ly = a.p === "above" ? b.y + b.height : b.y;
      return [
        { axis: "x", pos: cx, from: Math.min(cy, ly), to: Math.max(cy, ly), points: [cy, ly] },
      ];
    }
    return [];
  }

  place();

  if (ui.focusSelected && selected) {
    for (const [n, e] of nodeEls) {
      if (n.name === selected && e.hit) e.hit.focus({ preventScroll: true });
    }
  }
  return handle;
}
