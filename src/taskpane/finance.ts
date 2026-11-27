/*
 * "Money flow" color scheme: classifies each box as income, spending, savings,
 * remainder or account (neutral) — the way household budgets and financial
 * statements are usually colored.
 */

import type { SankeyLayout, SankeyNode } from "./sankey";

export type NodeRole = "income" | "spending" | "savings" | "remainder" | "account";

// Checked in this order, so "Income tax" is spending and "Savings account" is savings.
const KEYWORDS: [NodeRole, RegExp][] = [
  [
    "spending",
    /\b(costs?|cogs|expenses?|opex|capex|tax(es)?|loss(es)?|depreciation|amorti[sz]ation|r&d|research|sg&a|selling|marketing|administrative|admin|interest expense|write-?offs?|impairments?|rent|mortgage|bills?|subscriptions?|insurance|utilities|groceries|debts?|loans?)\b/i,
  ],
  [
    "savings",
    /\b(savings?|funds?|invest(ment)?s?|investing|pensions?|isa|retirement|emergency)\b/i,
  ],
  ["remainder", /\b(remainder|remaining|left ?over|leftover|spare|unallocated|buffer)\b/i],
  ["income", /\b(profits?|income|margins?|earnings|ebitd?a|ebit|surplus|gains?|retained)\b/i],
  ["account", /\b(accounts?|cards?|transfers?|wallets?|pots?)\b/i],
];

export function roleFromName(name: string): NodeRole | null {
  for (const [role, re] of KEYWORDS) if (re.test(name)) return role;
  return null;
}

/**
 * Roles for every node:
 *  1. Boxes with nothing flowing in are income.
 *  2. Otherwise a keyword in the name decides.
 *  3. Middle boxes (in and out) without a keyword are accounts.
 *  4. End boxes without a keyword follow their parents when all parents agree
 *     (e.g. a breakdown of "Bills"); otherwise they count as spending.
 */
export function computeRoles(layout: Pick<SankeyLayout, "nodes">): Map<SankeyNode, NodeRole> {
  const roles = new Map<SankeyNode, NodeRole>();
  // Left to right, so parents are classified before their children.
  const ordered = [...layout.nodes].sort((a, b) => a.column - b.column || a.index - b.index);
  for (const n of ordered) {
    if (!n.targetLinks.length) {
      roles.set(n, "income");
      continue;
    }
    const byName = roleFromName(n.name);
    if (byName) {
      roles.set(n, byName);
      continue;
    }
    if (n.sourceLinks.length) {
      roles.set(n, "account");
      continue;
    }
    const parents = new Set(n.targetLinks.map((l) => roles.get(l.source) ?? "account"));
    const only = parents.size === 1 ? [...parents][0] : null;
    roles.set(n, only && only !== "account" && only !== "income" ? only : "spending");
  }
  return roles;
}
