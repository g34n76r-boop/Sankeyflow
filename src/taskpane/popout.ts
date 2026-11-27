/* global Office, setTimeout, clearTimeout, window, URL */

/*
 * "Pop out": shows the chart in a separate window (an Office dialog) while the
 * settings stay in the task pane, so the two can sit side by side. The window
 * draws the chart itself and is fully interactive (hover, select, move boxes
 * and labels); what you do there is sent back here, and the pane stays the one
 * source of truth: every change is sent back to the window as a fresh update.
 *
 * Messages to the window can be size-limited on some platforms, so each update
 * is sent as numbered chunks that the window puts back together.
 */

import type { PopoutAction, PopoutContent } from "./popoutModel";

const CHUNK = 16000;
const UPDATE_DELAY = 150;

let dialog: Office.Dialog | null = null;
let ready = false;
let render: ((width: number) => PopoutContent) | null = null;
/** Width the window has for the chart (it tells us when it loads or resizes). */
let width = 960;
let onChange: ((open: boolean) => void) | null = null;
let onAction: ((a: PopoutAction) => void) | null = null;
let reportError: ((message: string) => void) | null = null;
let timer: ReturnType<typeof setTimeout> | undefined;
let seq = 0;

export const popoutSupported = () => Office.context.requirements.isSetSupported("DialogApi", "1.2");

export const popoutOpen = () => dialog !== null;

/**
 * Opens the window. `content` is called for each update; `changed` hears when
 * the window opens or closes (including when it's closed with its own ✕);
 * `action` hears what you do to the chart in the window.
 */
export function openPopout(
  content: (width: number) => PopoutContent,
  changed: (open: boolean) => void,
  onError: (message: string) => void,
  action: (a: PopoutAction) => void
) {
  onAction = action;
  if (dialog) return;
  if (!popoutSupported()) {
    onError(
      "This version of Excel can't open a separate chart window. Try Excel for Microsoft 365 or Excel on the web."
    );
    return;
  }
  render = content;
  onChange = changed;
  reportError = onError;
  const url = new URL("preview.html", window.location.href).toString();
  Office.context.ui.displayDialogAsync(
    url,
    { height: 70, width: 55, displayInIframe: false, promptBeforeOpen: false },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // 12007: a window is already open; 12009: the user blocked the pop-up.
        onError(
          result.error.code === 12009
            ? "The chart window was blocked. Allow pop-ups for this add-in and try again."
            : `Couldn't open the chart window (${result.error.message}).`
        );
        return;
      }
      dialog = result.value;
      ready = false;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (!("message" in arg)) return;
        // {ready, width} when the window has loaded; {width} when it's resized;
        // {action, ...} when you use the chart there.
        let m: { ready?: boolean; width?: number; action?: string };
        try {
          m = JSON.parse(arg.message);
        } catch {
          return;
        }
        if (m.action) {
          onAction?.(m as PopoutAction);
          return;
        }
        if (typeof m.width === "number" && m.width > 0) width = m.width;
        if (m.ready) ready = true;
        send();
      });
      // Closed with its own ✕ (12006), or it failed to load.
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => closed());
      onChange?.(true);
    }
  );
}

export function closePopout() {
  if (!dialog) return;
  try {
    dialog.close();
  } catch {
    // Already gone.
  }
  closed();
}

function closed() {
  dialog = null;
  ready = false;
  if (timer) clearTimeout(timer);
  onChange?.(false);
}

/** The chart changed: update the window shortly (several quick changes send once). */
export function refreshPopout() {
  if (!dialog || !ready) return;
  if (timer) clearTimeout(timer);
  timer = setTimeout(send, UPDATE_DELAY);
}

function send() {
  // Runs from a timer or a window message, so nothing above would catch an
  // error: every failure is handled (and reported) here.
  timer = undefined;
  if (!dialog || !ready || !render) return;
  let payload: string;
  try {
    payload = JSON.stringify(render(width));
  } catch (e) {
    // Building the update failed (not just "can't draw this data", which the
    // window is told about): keep the last chart there, and say so here.
    const why = e instanceof Error ? e.message : String(e);
    reportError?.(`Couldn't update the chart window (${why}).`);
    return;
  }
  const id = ++seq;
  const n = Math.max(1, Math.ceil(payload.length / CHUNK));
  try {
    for (let i = 0; i < n; i++) {
      dialog.messageChild(
        JSON.stringify({ id, i, n, d: payload.slice(i * CHUNK, (i + 1) * CHUNK) })
      );
    }
  } catch {
    // The window is gone or no longer answering (e.g. it closed while this
    // update was waiting). Close it properly so the pane shows its chart again.
    closePopout();
    reportError?.(
      "The chart window stopped responding, so it was closed. Pop it out again to reopen it."
    );
  }
}
