/*
 * Label settings and value formatting (numbers, symbols, percentages).
 */

/* global document, CanvasRenderingContext2D */

import type { SankeyLayout, SankeyNode } from "./sankey";

export type LabelPlacement = "auto" | "left" | "right" | "above" | "below" | "hidden";
export type ValueMode = "none" | "number" | "percent" | "both";

export interface LabelSettings {
  /** CSS font-family stack. */
  fontFamily: string;
  fontSize: number;
  boldNames: boolean;
  /** null = theme text color. */
  color: string | null;
  valueMode: ValueMode;
  /** null = automatic (fewer decimals for big numbers). */
  decimals: number | null;
  prefix: string;
  suffix: string;
  placement: LabelPlacement;
  /** "below": value on its own line under the name. "inline": after the name. */
  valueLayout: "below" | "inline";
}

/** Per-node label changes, keyed by node name. */
export interface NodeLabelOverride {
  /** Replacement label text (the data isn't changed). */
  text?: string;
  placement?: LabelPlacement;
  /** Fine-tune offset in px from the placement position (set by dragging the label). */
  dx?: number;
  dy?: number;
}

export const FONT_CHOICES: { name: string; stack: string }[] = [
  { name: "Segoe UI", stack: "Segoe UI, system-ui, -apple-system, sans-serif" },
  { name: "Calibri", stack: "Calibri, Carlito, sans-serif" },
  { name: "Arial", stack: "Arial, Helvetica, sans-serif" },
  { name: "Verdana", stack: "Verdana, Geneva, sans-serif" },
  { name: "Georgia", stack: "Georgia, serif" },
  { name: "Times New Roman", stack: "Times New Roman, Times, serif" },
  { name: "Consolas", stack: "Consolas, Menlo, monospace" },
];

export const DEFAULT_LABELS: LabelSettings = {
  fontFamily: FONT_CHOICES[0].stack,
  fontSize: 12,
  boldNames: true,
  color: null,
  valueMode: "number",
  decimals: null,
  prefix: "",
  suffix: "",
  placement: "auto",
  valueLayout: "below",
};

/** Line height of label text, as a multiple of the font size. */
export const LINE_HEIGHT = 1.2;
/** Gap between a box and its label, in px. */
export const LABEL_GAP = 8;
/** Font weight for box names when "Bold names" is on. */
export const NAME_WEIGHT = 700;

/** Total flow through the diagram: everything entering from the start nodes. */
export function totalFlow(layout: SankeyLayout): number {
  return layout.nodes.filter((n) => !n.targetLinks.length).reduce((s, n) => s + n.outValue, 0);
}

function fixed(v: number, decimals: number | null, autoDecimals: number) {
  const d = decimals ?? autoDecimals;
  return v.toLocaleString(undefined, {
    minimumFractionDigits: decimals ?? 0,
    maximumFractionDigits: d,
  });
}

/** A number with the user's symbols, e.g. "$4,200" or "4,200 kg". */
export function formatNumber(v: number, s: LabelSettings): string {
  return s.prefix + fixed(v, s.decimals, Math.abs(v) < 10 ? 2 : 0) + s.suffix;
}

/** A share as a percentage, e.g. "82%" or "8.4%". */
export function formatPercent(share: number, s: LabelSettings): string {
  const pct = share * 100;
  return fixed(pct, s.decimals, Math.abs(pct) < 10 ? 1 : 0) + "%";
}

/** The value part of a node label, per the chosen mode ("" when values are hidden). */
export function valueText(n: SankeyNode, total: number, s: LabelSettings): string {
  const share = total ? n.value / total : 0;
  switch (s.valueMode) {
    case "none":
      return "";
    case "number":
      return formatNumber(n.value, s);
    case "percent":
      return formatPercent(share, s);
    case "both":
      return `${formatNumber(n.value, s)} (${formatPercent(share, s)})`;
  }
}

export function resolvePlacement(
  n: SankeyNode,
  layout: SankeyLayout,
  s: LabelSettings,
  o: NodeLabelOverride | undefined
): Exclude<LabelPlacement, "auto"> {
  const p = o?.placement ?? s.placement;
  if (p !== "auto") return p;
  // Auto: labels sit on the outside — starting boxes on the left, end boxes on
  // the right, middle boxes on the side nearer the chart's edge.
  if (!n.targetLinks.length) return "left";
  if (!n.sourceLinks.length) return "right";
  return (n.x0 + n.x1) / 2 < layout.width / 2 ? "left" : "right";
}

/* ---------- Measuring labels ---------- */

let measureCtx: CanvasRenderingContext2D | null = null;

/** Width of a piece of label text in px, measured with the label font. */
export function textWidth(text: string, s: LabelSettings, bold: boolean): number {
  if (!measureCtx) measureCtx = document.createElement("canvas").getContext("2d");
  if (!measureCtx) return text.length * s.fontSize * 0.6; // rough fallback
  measureCtx.font = `${bold && s.boldNames ? NAME_WEIGHT : 400} ${s.fontSize}px ${s.fontFamily}`;
  return measureCtx.measureText(text).width;
}

export interface LabelBlock {
  name: string;
  value: string;
  /** Size of the text block in px. */
  width: number;
  height: number;
  lines: 1 | 2;
}

/** The text of a node's label and the size of its block. */
export function labelBlock(
  n: SankeyNode,
  total: number,
  s: LabelSettings,
  o: NodeLabelOverride | undefined
): LabelBlock {
  const name = o?.text || n.name;
  const value = valueText(n, total, s);
  const lineH = s.fontSize * LINE_HEIGHT;
  const nameW = textWidth(name, s, true);
  if (!value) return { name, value, width: nameW, height: lineH, lines: 1 };
  const valueW = textWidth(value, s, false);
  if (s.valueLayout === "inline") {
    return { name, value, width: nameW + s.fontSize * 0.4 + valueW, height: lineH, lines: 1 };
  }
  return { name, value, width: Math.max(nameW, valueW), height: lineH * 2, lines: 2 };
}
