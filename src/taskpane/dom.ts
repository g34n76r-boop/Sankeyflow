/* Small DOM helpers shared by the task pane's modules. */

/* global document, HTMLElement, HTMLInputElement, HTMLSelectElement */

import { FONT_CHOICES } from "./labels";

/** The element with this id (the pane's markup always has it). */
export const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Sets a control's value, unless it's the one being typed in (or already shows it). */
export function setControl(id: string, value: string) {
  const el = $<HTMLInputElement | HTMLSelectElement>(id);
  if (document.activeElement !== el && el.value !== value) el.value = value;
}

/** One option per font choice, each shown in its own font. */
export function addFontOptions(select: HTMLSelectElement) {
  for (const f of FONT_CHOICES) {
    const opt = Object.assign(document.createElement("option"), {
      value: f.stack,
      textContent: f.name,
    });
    opt.style.fontFamily = f.stack;
    select.appendChild(opt);
  }
}

const HTML_ESCAPES: Record<string, string> = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
};

/** Text made safe to put inside HTML (element content or a double-quoted attribute). */
export const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (c) => HTML_ESCAPES[c]);
