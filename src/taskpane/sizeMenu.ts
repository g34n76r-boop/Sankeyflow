/*
 * The chart-size menu, opened from the dimensions icon in the chart's top-left
 * corner (in the task pane and in the pop-out window): type a width and height
 * in px, or go back to the automatic size. Blank boxes mean automatic and show
 * the current automatic size as a hint.
 */

/* global document, HTMLButtonElement, HTMLInputElement, Node, KeyboardEvent */

import { CHART_SIZE_LIMITS, parseChartSize } from "./chartSize";
import type { ChartSize } from "./chartSize";

const [MIN_W, MAX_W] = CHART_SIZE_LIMITS.width;
const [MIN_H, MAX_H] = CHART_SIZE_LIMITS.height;

export interface SizeMenu {
  /** Shows the chosen size, with the chart's drawn size as the "auto" hint. */
  sync(size: ChartSize, drawn: { width: number; height: number }): void;
  close(): void;
}

let menuCount = 0;

/**
 * Turns `button` into the opener of a small size menu placed under it.
 * `onChange` gets the new size (null = automatic) when a box is changed
 * (Enter or leaving it) or "Automatic size" is pressed.
 */
export function attachSizeMenu(
  button: HTMLButtonElement,
  onChange: (size: ChartSize) => void
): SizeMenu {
  const id = `size-menu-${++menuCount}`;
  const wrap = document.createElement("span");
  wrap.className = "size-menu-wrap";
  button.replaceWith(wrap);
  wrap.appendChild(button);

  const menu = document.createElement("div");
  menu.className = "size-menu";
  menu.id = id;
  menu.hidden = true;
  menu.setAttribute("role", "dialog");
  menu.setAttribute("aria-label", "Chart size");
  menu.innerHTML = `
    <div class="size-menu-title">Chart size</div>
    <label class="size-menu-row"><span>Width</span>
      <input type="number" class="size-menu-input" data-axis="width" min="${MIN_W}" max="${MAX_W}" step="10" placeholder="auto" aria-label="Chart width in pixels (blank for automatic)" />
      <span class="size-menu-unit">px</span></label>
    <label class="size-menu-row"><span>Height</span>
      <input type="number" class="size-menu-input" data-axis="height" min="${MIN_H}" max="${MAX_H}" step="10" placeholder="auto" aria-label="Chart height in pixels (blank for automatic)" />
      <span class="size-menu-unit">px</span></label>
    <p class="size-menu-hint">Leave blank for automatic. Also used for the picture inserted into the sheet.</p>
    <button type="button" class="size-menu-auto">Automatic size</button>`;
  wrap.appendChild(menu);

  button.setAttribute("aria-haspopup", "dialog");
  button.setAttribute("aria-controls", id);
  button.setAttribute("aria-expanded", "false");

  const wIn = menu.querySelector<HTMLInputElement>('[data-axis="width"]')!;
  const hIn = menu.querySelector<HTMLInputElement>('[data-axis="height"]')!;
  const autoBtn = menu.querySelector<HTMLButtonElement>(".size-menu-auto")!;

  const apply = () =>
    onChange({
      width: parseChartSize(wIn.value, "width"),
      height: parseChartSize(hIn.value, "height"),
    });
  wIn.addEventListener("change", apply);
  hIn.addEventListener("change", apply);
  autoBtn.onclick = () => {
    wIn.value = hIn.value = "";
    onChange({ width: null, height: null });
  };

  const open = () => {
    menu.hidden = false;
    button.setAttribute("aria-expanded", "true");
    button.classList.add("active");
    wIn.focus();
    wIn.select();
  };
  const close = (refocus = false) => {
    if (menu.hidden) return;
    menu.hidden = true;
    button.setAttribute("aria-expanded", "false");
    button.classList.remove("active");
    if (refocus) button.focus();
  };
  button.addEventListener("click", () => (menu.hidden ? open() : close()));
  // Escape, or a click anywhere outside the menu, closes it.
  menu.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "Escape") {
      e.preventDefault();
      close(true);
    }
  });
  document.addEventListener(
    "pointerdown",
    (e) => {
      if (!wrap.contains(e.target as Node)) close();
    },
    true
  );

  return {
    sync(size, drawn) {
      if (document.activeElement !== wIn) wIn.value = size.width?.toString() ?? "";
      if (document.activeElement !== hIn) hIn.value = size.height?.toString() ?? "";
      const px = (v: number) => (v ? String(Math.round(v)) : "–");
      wIn.placeholder = `auto (${px(drawn.width)})`;
      hIn.placeholder = `auto (${px(drawn.height)})`;
      autoBtn.disabled = size.width === null && size.height === null;
      button.classList.toggle("set", size.width !== null || size.height !== null);
    },
    close: () => close(),
  };
}
