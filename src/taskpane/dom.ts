/* Small DOM helpers shared by the task pane's modules. */

/* global document, HTMLElement */

/** The element with this id (the pane's markup always has it). */
export const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const HTML_ESCAPES: Record<string, string> = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
};

/** Text made safe to put inside HTML (element content or a double-quoted attribute). */
export const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (c) => HTML_ESCAPES[c]);
