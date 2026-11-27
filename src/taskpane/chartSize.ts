/*
 * The chart's chosen size (the size menu and the chart's dotted edges, in the
 * task pane and the pop-out window): used for the pane, the window and the
 * picture inserted into the sheet alike.
 */

/** A chosen width and height in px; null = automatic. */
export interface ChartSize {
  width: number | null;
  height: number | null;
}

/** The sizes you can choose for each dimension, in px: [smallest, largest]. */
export const CHART_SIZE_LIMITS: Record<"width" | "height", [number, number]> = {
  width: [320, 2400],
  height: [160, 2000],
};

/**
 * A chart width or height typed into a box: whole px within the limits, or
 * null (automatic) when blank or not a number.
 */
export function parseChartSize(raw: string, axis: "width" | "height"): number | null {
  const v = Number(raw);
  if (raw.trim() === "" || !Number.isFinite(v)) return null;
  const [lo, hi] = CHART_SIZE_LIMITS[axis];
  return Math.round(Math.min(hi, Math.max(lo, v)));
}
