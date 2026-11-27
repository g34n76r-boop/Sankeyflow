/*
 * Reading back a diagram's saved settings safely: whatever a workbook holds
 * (saved by an older version, partly missing, or edited by hand), the pane
 * gets a complete, well-formed set.
 */

import type { ChartTitle, LinkColorMode } from "./chart";
import type { LabelSettings } from "./labels";
import { CASHFLOW } from "./memory";
import type { BackgroundMode, ColorGroup, DiagramSettings, GroupSet } from "./memory";
import { industryPreset } from "./domains";

/** A deep copy of plain JSON data. */
export const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v));

/** A plain object, or {} for anything else (null, an array, a string…). */
const asObject = (v: unknown): Record<string, unknown> =>
  v && typeof v === "object" && !Array.isArray(v) ? (v as Record<string, unknown>) : {};

/** `v` if it's a finite number or null, otherwise `fallback`. */
const sizeOr = (v: unknown, fallback: number | null) =>
  v === null || (typeof v === "number" && isFinite(v)) ? v : fallback;

const isGroupSet = (v: unknown): v is GroupSet =>
  v === CASHFLOW || v === "generic" || (typeof v === "string" && !!industryPreset(v));

/**
 * Saved settings with every gap filled from `defaults`, at every level, and
 * anything of the wrong kind replaced by its default rather than breaking the
 * pane. The result is a fresh copy: nothing in it is shared with `defaults`
 * (editing a group's colour changes it in place, and mustn't change what
 * Clear goes back to).
 */
export function withDefaults(
  saved: Partial<DiagramSettings>,
  defaults: DiagramSettings
): DiagramSettings {
  const d = clone(defaults);
  const s = asObject(clone(saved ?? {}));
  const list = <T>(v: unknown, fallback: T[]): T[] => (Array.isArray(v) ? (v as T[]) : fallback);
  // [name, value] lists: keep only well-formed pairs.
  const pairs = <V>(v: unknown, fallback: [string, V][]): [string, V][] =>
    list<unknown>(v, fallback).filter(
      (e): e is [string, V] => Array.isArray(e) && e.length === 2 && typeof e[0] === "string"
    );
  const background = asObject(s.background);
  const chartSizeSaved = asObject(s.chartSize);
  const color = background.color;
  return {
    nodeColors: pairs(s.nodeColors, d.nodeColors),
    nodePositions: pairs(s.nodePositions, d.nodePositions),
    labelOverrides: pairs(s.labelOverrides, d.labelOverrides),
    labelSettings: { ...d.labelSettings, ...asObject(s.labelSettings) } as LabelSettings,
    chartTitle: { ...d.chartTitle, ...asObject(s.chartTitle) } as ChartTitle,
    groupSet: isGroupSet(s.groupSet) ? s.groupSet : d.groupSet,
    groups: list<ColorGroup>(s.groups, d.groups).filter(
      (g) => !!g && typeof g.id === "string" && typeof g.name === "string"
    ),
    genericTheme: typeof s.genericTheme === "string" ? s.genericTheme : d.genericTheme,
    addedGroupCount:
      typeof s.addedGroupCount === "number" && isFinite(s.addedGroupCount)
        ? s.addedGroupCount
        : d.addedGroupCount,
    background: {
      mode: (background.mode as BackgroundMode) ?? d.background.mode,
      color:
        typeof color === "string" && /^#[0-9a-f]{6}$/i.test(color) ? color : d.background.color,
    },
    linkColor: (s.linkColor as LinkColorMode) ?? d.linkColor,
    endAlign: (s.endAlign as DiagramSettings["endAlign"]) ?? d.endAlign,
    nodeOrder: (s.nodeOrder as DiagramSettings["nodeOrder"]) ?? d.nodeOrder,
    chartSize: {
      width: sizeOr(chartSizeSaved.width, d.chartSize.width),
      height: sizeOr(chartSizeSaved.height, d.chartSize.height),
    },
    summaryCollapsed:
      typeof s.summaryCollapsed === "boolean" ? s.summaryCollapsed : d.summaryCollapsed,
  };
}
