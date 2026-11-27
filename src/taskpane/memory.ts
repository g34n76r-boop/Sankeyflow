/*
 * Saved diagrams: each diagram's data link (a table or a cell range) and all of
 * its settings, stored in the workbook's document settings so they come back
 * when the file is reopened. Also draws the Memory tab, which lists them and
 * lets you open or forget each one.
 */

/* global document, Office, HTMLElement, HTMLElementTagNameMap, HTMLInputElement, HTMLButtonElement */

import type { ChartTitle, ColorChoice, LinkColorMode } from "./chart";
import type { LabelSettings, NodeLabelOverride } from "./labels";
import type { NodePosition } from "./sankey";
import type { ChartSize } from "./chartSize";
import type { NodeRole } from "./finance";
import type { IndustryId, Tone } from "./domains";
import { $ } from "./dom";

/* ---------- Types shared with taskpane.ts ---------- */

export type BackgroundMode = "auto" | "white" | "dark" | "transparent" | "custom";

/**
 * Which set of groups is in use: Cashflow (Income, Accounts, …), "generic"
 * (Traffic light, Ocean, … : Group 1…Group 5), or an industry.
 */
export type GroupSet = typeof CASHFLOW | "generic" | IndustryId;

/**
 * The Cashflow set of groups. Saved charts store it under its original name,
 * "money", so that value stays.
 */
export const CASHFLOW = "money";

/**
 * A color group. The five built-in groups have a `role` (used by the
 * name-guessing rules); groups you add are matched by their keywords only.
 */
export interface ColorGroup {
  id: string;
  name: string;
  /** null = the built-in, theme-aware color (role or tone). */
  color: string | null;
  /** Words that put a box in this group (checked before the built-in rules). */
  keywords: string[];
  /** Cashflow groups: the role the name-guessing rules assign. */
  role?: NodeRole;
  /** Industry groups: theme-aware colour, and the position whose unmatched boxes land here. */
  tone?: Tone;
  position?: "start" | "middle" | "end";
}

/** Where a saved diagram's data comes from. A range is remembered by its address only. */
export type SavedSource =
  { kind: "table"; id: string; name: string; sheet: string } | { kind: "range"; address: string };

/** Everything about a diagram's look that you can change. */
export interface DiagramSettings {
  nodeColors: [string, ColorChoice][];
  nodePositions: [string, NodePosition][];
  labelOverrides: [string, NodeLabelOverride][];
  labelSettings: LabelSettings;
  chartTitle: ChartTitle;
  groupSet: GroupSet;
  groups: ColorGroup[];
  genericTheme: string;
  addedGroupCount: number;
  background: { mode: BackgroundMode; color: string };
  linkColor: LinkColorMode;
  endAlign: "left" | "justify";
  nodeOrder: "data" | "color";
  chartSize: ChartSize;
  summaryCollapsed: boolean;
}

/** A picture this diagram inserted into the sheet. */
export interface SavedPicture {
  sheet: string;
  shapeId: string;
}

export interface SavedDiagram {
  id: string;
  name: string;
  source: SavedSource;
  settings: DiagramSettings;
  pictures: SavedPicture[];
  /** When it was last saved (ms since 1970). */
  savedAt: number;
}

/* ---------- Storage (inside the workbook) ---------- */

const DIAGRAMS_KEY = "sankeyflow.diagrams";
const OPEN_KEY = "sankeyflow.openDiagram";
/** Before saved diagrams: just the table that was picked last. */
const LEGACY_TABLE_KEY = "sankeyflow.selectedTable";

function read<T>(key: string): T | null {
  try {
    return (Office.context.document.settings.get(key) as T) ?? null;
  } catch {
    return null;
  }
}

/**
 * Sets (or, with null, removes) workbook settings and saves them. Only writes
 * when something actually changes, so the workbook isn't marked as modified
 * needlessly. Throws if Excel couldn't save, after putting the settings back
 * as they were (otherwise a retry would see "no change" and never save).
 */
async function persist(changes: Record<string, unknown>) {
  const settings = Office.context.document.settings;
  const previous = new Map<string, unknown>();
  for (const [key, value] of Object.entries(changes)) {
    const before = read(key);
    if (JSON.stringify(before) === JSON.stringify(value ?? null)) continue;
    previous.set(key, before);
    if (value === null || value === undefined) settings.remove(key);
    else settings.set(key, value);
  }
  if (!previous.size) return;
  try {
    await new Promise<void>((resolve, reject) =>
      settings.saveAsync((r) =>
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
      )
    );
  } catch (e) {
    for (const [key, before] of previous) {
      if (before === null) settings.remove(key);
      else settings.set(key, before);
    }
    throw e;
  }
}

const isDiagram = (d: unknown): d is SavedDiagram => {
  const x = d as SavedDiagram;
  return !!x && typeof x.id === "string" && !!x.source && !!x.settings;
};

export function listDiagrams(): SavedDiagram[] {
  const v = read<unknown[]>(DIAGRAMS_KEY);
  return Array.isArray(v)
    ? v.filter(isDiagram).map((d) => ({ ...d, pictures: d.pictures ?? [] }))
    : [];
}

export const diagramById = (id: string | null) =>
  id ? listDiagrams().find((d) => d.id === id) : undefined;

/** The saved diagram for a table (matched by ID, then by name) or a range (by address). */
export function findBySource(src: SavedSource): SavedDiagram | undefined {
  const all = listDiagrams();
  if (src.kind === "range") {
    return all.find((d) => d.source.kind === "range" && d.source.address === src.address);
  }
  const tables = all.filter(
    (d): d is SavedDiagram & { source: { kind: "table" } } => d.source.kind === "table"
  );
  return (
    tables.find((d) => d.source.id === src.id) ?? tables.find((d) => d.source.name === src.name)
  );
}

/**
 * Adds the diagram, or replaces the saved one with the same ID. With `open`,
 * it's also remembered as the open diagram (reopened with the workbook), in
 * the same save.
 */
export async function saveDiagram(d: SavedDiagram, opts: { open?: boolean } = {}) {
  const all = listDiagrams();
  const i = all.findIndex((x) => x.id === d.id);
  if (i >= 0) all[i] = d;
  else all.push(d);
  await persist({ [DIAGRAMS_KEY]: all, ...(opts.open ? { [OPEN_KEY]: d.id } : {}) });
}

export async function renameDiagram(id: string, name: string) {
  const all = listDiagrams();
  const d = all.find((x) => x.id === id);
  if (!d || !name.trim() || d.name === name.trim()) return;
  d.name = name.trim();
  await persist({ [DIAGRAMS_KEY]: all });
}

/** Removes the diagram and everything saved for it. */
export async function forgetDiagram(id: string) {
  const all = listDiagrams().filter((d) => d.id !== id);
  const changes: Record<string, unknown> = { [DIAGRAMS_KEY]: all.length ? all : null };
  if (read(OPEN_KEY) === id) changes[OPEN_KEY] = null;
  await persist(changes);
}

/** Which diagram is open, so it comes back when the workbook is reopened. */
export async function setOpenDiagram(id: string | null) {
  await persist({ [OPEN_KEY]: id });
}

export const openDiagramSetting = () => read<string>(OPEN_KEY);

/** The table the add-in followed before it saved diagrams (read once, then removed). */
export async function takeLegacyTable(): Promise<{ id: string; name: string } | null> {
  const v = read<{ id: string; name: string }>(LEGACY_TABLE_KEY);
  if (v === null) return null;
  try {
    await persist({ [LEGACY_TABLE_KEY]: null });
  } catch {
    // Not fatal: it's read again next time.
  }
  return typeof v?.id === "string" && typeof v?.name === "string" ? v : null;
}

export function newDiagramId(): string {
  return `d-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

export function sourceText(s: SavedSource): string {
  return s.kind === "table" ? `Table "${s.name}" (${s.sheet})` : s.address;
}

/* ---------- Memory tab ---------- */

/**
 * What the Memory tab needs. The tab is only shown while no chart is loaded,
 * so none of the listed diagrams is ever the open one.
 */
export interface MemoryView {
  /** Workbook tables that exist now (to mark saved diagrams whose table is gone). */
  tables: { id: string; name: string }[] | null;
  /** A short description of the diagram's colour theme. */
  themeName: (d: SavedDiagram) => string;
  onOpen: (d: SavedDiagram) => void;
  onForget: (d: SavedDiagram) => void;
  onRename: (d: SavedDiagram, name: string) => void;
}

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className = "",
  text = ""
): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  if (text) e.textContent = text;
  return e;
}

let outsideClickWired = false;

/** The Memory tab: one card per saved diagram, with Open and Forget. */
export function renderMemory(view: MemoryView) {
  const list = $("memory-list");
  if (!outsideClickWired) {
    outsideClickWired = true;
    // Clicking anywhere else closes an open "Forget …?" box without forgetting.
    document.addEventListener("mousedown", (e) => {
      const inside = (e.target as HTMLElement).closest(".memory-card .clear-wrap");
      list.querySelectorAll<HTMLElement>(".memory-card .clear-wrap").forEach((w) => {
        if (w === inside) return;
        const pop = w.querySelector<HTMLElement>(".confirm-pop")!;
        if (pop.hidden) return;
        pop.hidden = true;
        w.querySelector("button")?.setAttribute("aria-expanded", "false");
      });
    });
  }
  const diagrams = listDiagrams();
  $("memory-empty").hidden = diagrams.length > 0;
  list.hidden = diagrams.length === 0;
  // Keep a name being typed in (rebuilding would drop its focus), unless its
  // diagram is gone.
  const active = document.activeElement as HTMLElement | null;
  const editing = active?.classList.contains("memory-name")
    ? active.closest<HTMLElement>("[data-diagram]")
    : null;
  if (
    editing &&
    list.contains(editing) &&
    diagrams.some((d) => d.id === editing.dataset.diagram) &&
    list.children.length === diagrams.length
  ) {
    return;
  }

  list.innerHTML = "";
  for (const d of diagrams) {
    const li = el("li", "memory-card");
    li.dataset.diagram = d.id;

    const head = el("div", "memory-head");
    const name = el("input", "text-input memory-name") as HTMLInputElement;
    name.type = "text";
    name.value = d.name;
    name.maxLength = 60;
    name.setAttribute("aria-label", "Diagram name");
    name.addEventListener("change", () => {
      if (name.value.trim()) view.onRename(d, name.value);
      else name.value = d.name;
    });
    name.addEventListener("keydown", (e) => {
      if (e.key === "Enter") name.blur();
      if (e.key === "Escape") {
        name.value = d.name;
        name.blur();
      }
    });
    head.appendChild(name);
    li.appendChild(head);

    // Where the data comes from (and whether it's still there).
    const missing =
      d.source.kind === "table" &&
      view.tables !== null &&
      !view.tables.some(
        (t) => d.source.kind === "table" && (t.id === d.source.id || t.name === d.source.name)
      );
    const src = el("div", "memory-source", sourceText(d.source));
    if (missing) src.appendChild(el("span", "memory-missing", " · table missing"));
    li.appendChild(src);

    const facts = [view.themeName(d)];
    const title = d.settings.chartTitle?.text?.trim();
    if (title) facts.push(`Title "${title}"`);
    if (d.pictures.length) {
      facts.push(`${d.pictures.length} picture${d.pictures.length === 1 ? "" : "s"} in sheet`);
    }
    facts.push(
      `Saved ${new Date(d.savedAt).toLocaleString(undefined, {
        dateStyle: "medium",
        timeStyle: "short",
      })}`
    );
    li.appendChild(el("div", "memory-facts", facts.join(" · ")));

    // [Open] [Forget]; Forget asks first, in a small box under the button.
    const actions = el("div", "memory-actions");
    const open = el("button", "btn", "Open") as HTMLButtonElement;
    open.type = "button";
    open.title = "Open this diagram (its data and settings)";
    open.onclick = () => view.onOpen(d);

    const wrap = el("span", "clear-wrap");
    const forget = el("button", "btn", "Forget") as HTMLButtonElement;
    forget.type = "button";
    forget.setAttribute("aria-haspopup", "dialog");
    forget.setAttribute("aria-expanded", "false");
    const pop = el("div", "confirm-pop");
    pop.setAttribute("role", "alertdialog");
    pop.hidden = true;
    pop.appendChild(el("p", "confirm-title", `Forget "${d.name}"?`));
    pop.appendChild(
      el(
        "p",
        "confirm-text",
        "Its saved settings and data link are removed from this workbook" +
          ". Your data isn't changed" +
          (d.pictures.length ? ", and pictures already in the sheet are kept." : ".")
      )
    );
    const confirmRow = el("div", "confirm-actions");
    const cancel = el("button", "btn", "Cancel") as HTMLButtonElement;
    cancel.type = "button";
    const ok = el("button", "btn danger", "Forget") as HTMLButtonElement;
    ok.type = "button";
    confirmRow.append(cancel, ok);
    pop.appendChild(confirmRow);
    const show = (on: boolean) => {
      pop.hidden = !on;
      forget.setAttribute("aria-expanded", String(on));
      if (on) cancel.focus();
    };
    forget.onclick = () => show(pop.hidden);
    cancel.onclick = () => {
      show(false);
      forget.focus();
    };
    ok.onclick = () => {
      show(false);
      view.onForget(d);
    };
    pop.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        show(false);
        forget.focus();
      }
    });
    wrap.append(forget, pop);
    actions.append(open, wrap);
    li.appendChild(actions);
    list.appendChild(li);
  }
}
