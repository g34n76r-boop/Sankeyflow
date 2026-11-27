/*
 * Colour groups: the ready-made sets of groups (Cashflow, the Generic themes,
 * the industry themes), and the rules that put each box in a group. Pure
 * functions: the pane passes in the groups in use.
 *
 * ColorGroup and GroupSet are defined in memory.ts (they're saved with each diagram).
 */

import type { ChartTheme, ColorChoice, NodeGroup } from "./chart";
import type { SankeyNode } from "./sankey";
import { computeRoles } from "./finance";
import type { NodeRole } from "./finance";
import { industryPreset } from "./domains";
import type { IndustryId, Tone } from "./domains";
import { CASHFLOW } from "./memory";
import type { ColorGroup, GroupSet } from "./memory";

/** Each box's group, with the group's id. */
export type GroupMap = Map<SankeyNode, NodeGroup & { id: string }>;

/** Colour choice that assigns a box to a group by hand: "group:<id>". */
export const GROUP_PREFIX = "group:";

/** Display order of the built-in groups. */
const ROLE_ORDER: NodeRole[] = ["income", "account", "spending", "savings", "remainder"];

const DEFAULT_GROUP_NAMES: Record<NodeRole, string> = {
  income: "Income",
  account: "Accounts",
  spending: "Spending",
  savings: "Savings",
  remainder: "Remainder",
};

/** Ready-made colour sets. "cashflow" uses the built-in, theme-aware colours. */
export const GROUP_THEMES: {
  id: string;
  name: string;
  colors: Partial<Record<NodeRole, string>>;
}[] = [
  { id: "cashflow", name: "Cashflow", colors: {} },
  {
    id: "traffic",
    name: "Traffic light",
    colors: {
      income: "#2e9e44",
      account: "#7a7a7a",
      spending: "#d64541",
      savings: "#2f6fb7",
      remainder: "#e0a100",
    },
  },
  {
    id: "ocean",
    name: "Ocean",
    colors: {
      income: "#0f7c8c",
      account: "#5b6770",
      spending: "#e4572e",
      savings: "#1d4e89",
      remainder: "#f2a541",
    },
  },
  {
    id: "pastel",
    name: "Pastel",
    colors: {
      income: "#6cc28a",
      account: "#a7a7a7",
      spending: "#ef7f8f",
      savings: "#79a7d3",
      remainder: "#f2c46d",
    },
  },
  {
    id: "mono",
    name: "Monochrome",
    colors: {
      income: "#1f1f1f",
      account: "#8a8a8a",
      spending: "#4d4d4d",
      savings: "#666666",
      remainder: "#b3b3b3",
    },
  },
];

/** The industry of a set of groups (null for Cashflow and the Generic themes). */
export const industryOf = (set: GroupSet): IndustryId | null =>
  set === CASHFLOW || set === "generic" ? null : set;

/** A built-in group's default name: "Income"… in Cashflow, "Group 1"… in the Generic themes. */
export function defaultGroupName(g: ColorGroup, set: GroupSet): string {
  if (!g.role) return "Group";
  return set === "generic"
    ? `Group ${ROLE_ORDER.indexOf(g.role) + 1}`
    : DEFAULT_GROUP_NAMES[g.role];
}

/** The five built-in groups: named Income… (Cashflow) or Group 1… (Generic themes). */
export const builtinGroups = (numbered = false): ColorGroup[] =>
  ROLE_ORDER.map((r, i) => ({
    id: r,
    name: numbered ? `Group ${i + 1}` : DEFAULT_GROUP_NAMES[r],
    color: null,
    keywords: [],
    role: r,
  }));

function industryGroups(id: IndustryId): ColorGroup[] {
  return industryPreset(id)!.groups.map((g, i) => ({
    id: `${id}-${i}`,
    name: g.name,
    color: null,
    keywords: [...g.keywords],
    tone: g.tone,
    position: g.position,
  }));
}

/** The groups a set starts with, before colouring (see startingGroups). */
export function defaultGroups(set: GroupSet): ColorGroup[] {
  const ind = industryOf(set);
  return ind ? industryGroups(ind) : builtinGroups(set === "generic");
}

/**
 * What "Reset groups" goes back to: the set's groups, in a Generic theme
 * (`genericTheme`) with that theme's colours.
 */
export function startingGroups(set: GroupSet, genericTheme: string): ColorGroup[] {
  const start = defaultGroups(set);
  if (set === "generic") {
    const t = GROUP_THEMES.find((x) => x.id === genericTheme);
    if (t) applyThemeColors(start, t.colors);
  }
  return start;
}

/** Applies a theme's colours to the built-in groups (groups you added keep theirs). */
export function applyThemeColors(groups: ColorGroup[], colors: Partial<Record<NodeRole, string>>) {
  for (const g of groups) if (g.role) g.color = colors[g.role] ?? null;
}

function toneColor(tone: Tone, theme: ChartTheme): string {
  return tone === "blue" ? theme.series[0] : theme.roles[tone];
}

/** A group's colour: your pick, or its built-in, theme-aware one. */
export function groupColor(g: ColorGroup, theme: ChartTheme): string {
  if (g.color) return g.color;
  if (g.role) return theme.roles[g.role];
  if (g.tone) return toneColor(g.tone, theme);
  return theme.neutral;
}

/** The group a box was put in by hand (its colour choice is "group:<id>"), if it still exists. */
export function manualGroupOf(
  choice: ColorChoice | undefined,
  groups: ColorGroup[]
): ColorGroup | undefined {
  if (!choice?.startsWith(GROUP_PREFIX)) return undefined;
  return groups.find((g) => g.id === choice.slice(GROUP_PREFIX.length));
}

/** Compiled keyword patterns: every box is checked against every keyword on each draw. */
const keywordPatterns = new Map<string, RegExp | null>();

/** Does a box name contain a keyword (as a whole word, plural allowed)? */
function keywordMatch(name: string, keyword: string): boolean {
  let re = keywordPatterns.get(keyword);
  if (re === undefined) {
    const k = keyword.trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    re = k ? new RegExp(`(^|\\W)${k}s?(\\W|$)`, "i") : null;
    // Keywords are typed a letter at a time; don't keep every half-typed one.
    if (keywordPatterns.size > 500) keywordPatterns.clear();
    keywordPatterns.set(keyword, re);
  }
  return !!re && re.test(name);
}

/** The group a box falls into automatically: your keywords first, then the built-in rules. */
export function autoGroup(
  n: SankeyNode,
  groups: ColorGroup[],
  roles: Map<SankeyNode, NodeRole>,
  memo = new Map<SankeyNode, ColorGroup | undefined>()
): ColorGroup | undefined {
  if (memo.has(n)) return memo.get(n);
  const found = findAutoGroup(n, groups, roles, memo);
  memo.set(n, found);
  return found;
}

function findAutoGroup(
  n: SankeyNode,
  groups: ColorGroup[],
  roles: Map<SankeyNode, NodeRole>,
  memo: Map<SankeyNode, ColorGroup | undefined>
): ColorGroup | undefined {
  // Keywords are checked from the last group to the first: groups you add (at
  // the end) win, and outcome groups ("Customer returns" → Returns, "Lost
  // leads" → Lost) beat the broader groups before them.
  for (let i = groups.length - 1; i >= 0; i--) {
    if (groups[i].keywords.some((k) => keywordMatch(n.name, k))) return groups[i];
  }
  // Personal / Generic themes: the name-guessing rules in finance.ts.
  const r = roles.get(n);
  const byRole = groups.find((g) => g.role && g.role === r);
  if (byRole) return byRole;
  // Professional themes: a box whose parents are all in the same outcome group
  // joins it (e.g. "Other" under "Operating expenses" → Costs), as in Cashflow.
  if (n.targetLinks.length) {
    const parents = new Set(n.targetLinks.map((l) => autoGroup(l.source, groups, roles, memo)));
    const only = parents.size === 1 ? [...parents][0] : undefined;
    if (only && only.position !== "start" && only.position !== "middle") return only;
  }
  // Otherwise by position in the chart.
  const pos = !n.targetLinks.length ? "start" : !n.sourceLinks.length ? "end" : "middle";
  return groups.find((g) => g.position === pos);
}

/**
 * Each box's group. Order of precedence: assigned by hand (`nodeColors`),
 * then your keywords, then the built-in rules. Boxes whose group was removed
 * are left out (drawn as "Other").
 */
export function groupMap(
  nodes: SankeyNode[],
  theme: ChartTheme,
  groups: ColorGroup[],
  nodeColors: Map<string, ColorChoice>
): GroupMap {
  const roles = computeRoles({ nodes });
  // One memo for all boxes: a box's automatic group can depend on its parents',
  // which would otherwise be worked out again for every box below them.
  const memo = new Map<SankeyNode, ColorGroup | undefined>();
  const map: GroupMap = new Map();
  for (const n of nodes) {
    const g = manualGroupOf(nodeColors.get(n.name), groups) ?? autoGroup(n, groups, roles, memo);
    if (g) map.set(n, { id: g.id, name: g.name, color: groupColor(g, theme) });
  }
  return map;
}
