/*
 * What the task pane sends the pop-out window so it can draw the chart itself,
 * fully interactive: the finished layout (boxes and flows with their
 * positions), plus everything the chart needs to colour and label it. Plain
 * JSON, as it travels between windows; the window rebuilds the linked layout.
 *
 * And what the window sends back when you interact with the chart there.
 */

import { updateLinkPositions } from "./sankey";
import type { NodePosition, SankeyLayout, SankeyLink, SankeyNode } from "./sankey";
import type { ChartTheme, ChartTitle, LinkColorMode, NodeGroup } from "./chart";
import type { LabelSettings, NodeLabelOverride } from "./labels";

export interface ChartModel {
  layout: {
    width: number;
    height: number;
    columnCount: number;
    nodes: Pick<
      SankeyNode,
      "name" | "index" | "column" | "value" | "inValue" | "outValue" | "x0" | "x1" | "y0" | "y1"
    >[];
    /** source and target are node indexes. */
    links: { index: number; source: number; target: number; value: number; width: number }[];
  };
  theme: ChartTheme;
  linkColor: LinkColorMode;
  labels: LabelSettings;
  labelOverrides: [string, NodeLabelOverride][];
  colors: [string, string][];
  /** Colour group of each box, by node index (absent: not using groups). */
  groups?: [number, NodeGroup][];
  margin: { top: number; right: number; bottom: number; left: number };
  transparent: boolean;
  title: ChartTitle;
  selected: string | null;
  /** The chosen chart size (null = automatic), for the window's size menu. */
  size: { width: number | null; height: number | null };
}

/** Everything the window shows. */
export interface PopoutContent {
  /** The chart, or null when there's no data loaded (or it can't be drawn: see `error`). */
  chart: ChartModel | null;
  /** Why the chart can't be drawn, when data is loaded but drawing it failed. */
  error?: string;
  /** Colour behind the chart (the window's page background). */
  background: string;
  /** Text colour for messages in the window. */
  text: string;
  /** The background is transparent: show a checkerboard. */
  transparent: boolean;
  /** The toolbar's options, so the window's toolbar shows them as set in the pane. */
  options: ToolbarOptions;
}

/** The toolbar's dropdowns (by select id) and the custom background colour. */
export type ToolbarOptionId = "background-mode" | "node-order" | "end-align" | "link-color" | "background-color";
export type ToolbarOptions = Record<ToolbarOptionId, string>;

/** Sent back by the window as you use the chart there. */
export type PopoutAction =
  | { action: "nodeMoved"; name: string; position: NodePosition }
  | { action: "labelMoved"; name: string; dx: number; dy: number }
  | { action: "select"; name: string | null }
  /** A chosen chart size (null = automatic). */
  | { action: "resize"; width: number | null; height: number | null }
  /** A toolbar option changed in the window (Background, Order, Ends, Links, or the custom colour). */
  | { action: "option"; id: ToolbarOptionId; value: string }
  /** Bring the chart back into the task pane (closes the window). */
  | { action: "close" };

export function layoutToModel(layout: SankeyLayout): ChartModel["layout"] {
  return {
    width: layout.width,
    height: layout.height,
    columnCount: layout.columnCount,
    nodes: layout.nodes.map(
      ({ name, index, column, value, inValue, outValue, x0, x1, y0, y1 }) => ({
        name,
        index,
        column,
        value,
        inValue,
        outValue,
        x0,
        x1,
        y0,
        y1,
      })
    ),
    links: layout.links.map((l) => ({
      index: l.index,
      source: l.source.index,
      target: l.target.index,
      value: l.value,
      width: l.width,
    })),
  };
}

/** Rebuilds the linked layout (boxes know their flows, flows their boxes). */
export function modelToLayout(m: ChartModel["layout"]): SankeyLayout {
  const nodes: SankeyNode[] = m.nodes.map((n) => ({ ...n, sourceLinks: [], targetLinks: [] }));
  const byIndex = new Map(nodes.map((n) => [n.index, n]));
  const links: SankeyLink[] = [];
  for (const l of m.links) {
    const source = byIndex.get(l.source);
    const target = byIndex.get(l.target);
    if (!source || !target) continue;
    const link: SankeyLink = {
      index: l.index,
      source,
      target,
      value: l.value,
      width: l.width,
      y0: 0,
      y1: 0,
    };
    source.sourceLinks.push(link);
    target.targetLinks.push(link);
    links.push(link);
  }
  // Orders each box's flows and works out where they join it.
  updateLinkPositions(nodes);
  return { nodes, links, columnCount: m.columnCount, width: m.width, height: m.height };
}
