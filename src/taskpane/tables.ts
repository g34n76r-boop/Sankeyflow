/*
 * "Workbook tables" picker: lists every Excel table in the workbook, shows the
 * chosen table's rows (first three columns = Source, Target, Value), hands them
 * to the diagram, and keeps watching the table so the diagram redraws when it
 * changes (edits, rows added or deleted).
 */

/* global document, Excel, Office, OfficeExtension, HTMLElement, HTMLSelectElement, setTimeout, clearTimeout */

export interface TableRow {
  source: string;
  target: string;
  value: number;
}

export interface TableData {
  id: string;
  name: string;
  sheet: string;
  /** Header names of the first three columns. */
  headers: [string, string, string];
  rows: TableRow[];
  /** Rows skipped because the value was blank or non-numeric. */
  ignored: number;
  /** Total columns in the table (only the first three are used). */
  columns: number;
}

export interface TablePickerOptions {
  /** The pane's wrapper for user actions: disables buttons and shows errors. */
  run: (action: () => Promise<void>) => Promise<void>;
  /** Called with fresh rows when a table is picked and after each change to it. */
  onTableData: (data: TableData, reason: "picked" | "changed") => void;
  /** Shows an error from a background (event-driven) update. */
  onError: (message: string) => void;
  /** The table is no longer being watched (another source, table deleted, …). */
  onWatchEnded: (tableId: string) => void;
  /** Shows an informational message (not an error). */
  onNotice: (message: string) => void;
}

/** Wait this long after the last change before re-reading the table. */
const CHANGE_DEBOUNCE_MS = 500;

/**
 * Keeps rows whose third column is a real number. Rows with a blank or
 * non-numeric value (text, TRUE/FALSE, errors such as #N/A) are ignored.
 */
export function parseTableRows(values: unknown[][]): { rows: TableRow[]; ignored: number } {
  const rows: TableRow[] = [];
  let ignored = 0;
  for (const r of values) {
    const v = r[2];
    if (typeof v !== "number" || !isFinite(v)) {
      ignored++;
      continue;
    }
    rows.push({ source: String(r[0] ?? "").trim(), target: String(r[1] ?? "").trim(), value: v });
  }
  return { rows, ignored };
}

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

let options: TablePickerOptions;

/* ---------- Watching the selected table ---------- */

type Handler = OfficeExtension.EventHandlerResult<
  Excel.TableChangedEventArgs | Excel.WorksheetChangedEventArgs
>;

interface Watch {
  id: string;
  handlers: Handler[];
  /** Snapshot of the table's contents, so unrelated sheet edits don't trigger redraws. */
  signature: string;
}

let watch: Watch | null = null;
let changeTimer: ReturnType<typeof setTimeout> | undefined;
/** Increments on every read; a read whose result is no longer the latest is discarded. */
let readSeq = 0;

const canWatch = () => Office.context.requirements.isSetSupported("ExcelApi", "1.7");

async function startWatching(id: string, signature: string) {
  if (watch?.id === id) {
    watch.signature = signature;
    return;
  }
  await stopWatching();
  if (!canWatch()) return;

  const handlers = await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(id);
    // The table's own event covers edits and rows inserted/deleted inside it.
    // Its worksheet's event is a safety net for the table growing or shrinking
    // (e.g. typing just below it); the signature check stops needless redraws.
    const result: Handler[] = [
      table.onChanged.add(onChanged),
      table.worksheet.onChanged.add(onChanged),
    ];
    await context.sync();
    return result;
  });
  watch = { id, handlers, signature };
}

/** Stops watching the current table (e.g. another source was loaded). Safe to call anytime. */
export async function stopWatching() {
  if (changeTimer) clearTimeout(changeTimer);
  changeTimer = undefined;
  readSeq++; // discard any read still in flight
  const w = watch;
  watch = null;
  if (!w) return;
  options.onWatchEnded(w.id);
  try {
    await Excel.run(w.handlers[0].context, async (context) => {
      w.handlers.forEach((h) => h.remove());
      await context.sync();
    });
  } catch {
    // The table or sheet may already be gone; its handlers went with it.
  }
}

/** Event handler: restart the half-second timer on every change. */
async function onChanged(): Promise<void> {
  if (!watch) return;
  scheduleReload();
}

function scheduleReload() {
  if (changeTimer) clearTimeout(changeTimer);
  changeTimer = setTimeout(reloadWatched, CHANGE_DEBOUNCE_MS);
}

async function reloadWatched() {
  changeTimer = undefined;
  const w = watch;
  if (!w) return;
  const seq = ++readSeq;
  try {
    const data = await readTable(w.id);
    if (seq !== readSeq || watch !== w) return; // superseded
    const sig = signatureOf(data);
    if (sig === w.signature) return; // nothing in the table changed
    w.signature = sig;
    showData(data, "changed");
  } catch (e) {
    if (seq !== readSeq) return;
    const err = e as OfficeExtension.Error;
    // The user is still typing in a cell; try again shortly.
    if (err?.code === "InvalidOperationInCellEditMode") {
      scheduleReload();
      return;
    }
    if (err?.code === "ItemNotFound") {
      await stopWatching();
      options.onError("The selected table was deleted. Pick a table again.");
      await refreshTableList();
      return;
    }
    options.onError(err?.message ?? String(e));
  }
}

/* ---------- Picker UI ---------- */

/** The workbook's tables as last listed (null until the first listing). */
let knownTables: { id: string; name: string; sheet: string }[] | null = null;
export const listedTables = () => knownTables;

/** Fills the list of tables. Resolves once it's filled (the pane opens a saved diagram after). */
export async function setupTablePicker(opts: TablePickerOptions) {
  options = opts;
  const select = $<HTMLSelectElement>("table-select");
  select.onchange = () => opts.run(() => pickTable(select.value));
  $("tables-refresh").onclick = () => opts.run(async () => void (await refreshTableList()));
  await opts.run(async () => void (await refreshTableList()));
}

/**
 * Opens a saved diagram's table: matched by ID first (survives renames), then
 * by name. Returns false if the workbook no longer has it.
 */
export async function openTable(t: { id: string; name: string }): Promise<boolean> {
  const tables = await refreshTableList();
  const match = tables.find((x) => x.id === t.id) ?? tables.find((x) => x.name === t.name);
  if (!match) return false;
  $<HTMLSelectElement>("table-select").value = match.id;
  await pickTable(match.id);
  return true;
}

/** Puts the picker back to "Choose a table…" and stops following the table. */
export async function releaseTable() {
  await stopWatching();
  $<HTMLSelectElement>("table-select").value = "";
  clearRows();
}

/** Re-reads the selected table now (used by the pane's Refresh button). */
export async function reloadSelectedTable() {
  const id = $<HTMLSelectElement>("table-select").value;
  if (id) await pickTable(id);
}

async function refreshTableList(): Promise<{ id: string; name: string; sheet: string }[]> {
  const select = $<HTMLSelectElement>("table-select");
  const previous = select.value;

  const tables = await Excel.run(async (context) => {
    const coll = context.workbook.tables;
    coll.load("items/id,items/name,items/worksheet/name");
    await context.sync();
    return coll.items.map((t) => ({ id: t.id, name: t.name, sheet: t.worksheet.name }));
  });
  knownTables = tables;

  select.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = tables.length ? "Choose a table…" : "No tables in this workbook";
  select.appendChild(placeholder);
  for (const t of tables) {
    const opt = document.createElement("option");
    opt.value = t.id;
    opt.textContent = `${t.name} (${t.sheet})`;
    select.appendChild(opt);
  }
  select.disabled = tables.length === 0;

  // Keep the current choice if that table still exists; otherwise clear the list.
  if (previous && tables.some((t) => t.id === previous)) {
    select.value = previous;
    await pickTable(previous);
  } else {
    select.value = "";
    await stopWatching();
    clearRows();
  }
  return tables;
}

function clearRows() {
  $("table-summary").hidden = true;
}

async function pickTable(id: string) {
  if (!id) {
    await stopWatching();
    clearRows();
    return;
  }
  const select = $<HTMLSelectElement>("table-select");
  select.disabled = true;
  try {
    const seq = ++readSeq;
    const data = await readTable(id);
    if (seq !== readSeq) return;
    showData(data, "picked");
    await startWatching(id, signatureOf(data));
    if (!canWatch()) {
      options.onError(
        "This version of Excel can't report table changes, so the diagram won't update by itself. Use Refresh after editing."
      );
    }
  } finally {
    select.disabled = false;
  }
}

async function readTable(id: string): Promise<TableData> {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(id);
    const body = table.getDataBodyRange();
    const header = table.getHeaderRowRange();
    table.load("name,worksheet/name");
    body.load("values,columnCount");
    header.load("values");
    await context.sync();
    const h = header.values[0] ?? [];
    const headers = [0, 1, 2].map(
      (i) => String(h[i] ?? "") || ["Source", "Target", "Value"][i]
    ) as [string, string, string];
    const { rows, ignored } =
      body.columnCount >= 3 ? parseTableRows(body.values) : { rows: [], ignored: 0 };
    return {
      id,
      name: table.name,
      sheet: table.worksheet.name,
      headers,
      rows,
      ignored,
      columns: body.columnCount,
    };
  });
}

function signatureOf(d: TableData): string {
  return JSON.stringify([d.name, d.headers, d.rows, d.ignored, d.columns]);
}

function showData(data: TableData, reason: "picked" | "changed") {
  const summary = $("table-summary");
  summary.hidden = false;
  if (data.columns < 3) {
    summary.textContent = `"${data.name}" has only ${data.columns} column${data.columns === 1 ? "" : "s"}. It needs at least three: Source, Target and Value.`;
  } else {
    summary.textContent =
      `${data.rows.length} row${data.rows.length === 1 ? "" : "s"} · columns ${data.headers.join(", ")}` +
      (data.ignored ? ` · ${data.ignored} ignored (blank or non-numeric value)` : "");
  }
  options.onTableData(data, reason);
}
