/*
 * Icon dropdowns for the chart toolbar: a small button showing the chosen
 * option's icon, opening a menu of the options, each an icon with its name
 * beside it. The real <select> stays in the
 * page, hidden, and remains the source of truth: picking an icon sets its value
 * and fires "change", and setting its value from code updates the icon.
 */

/* global document, HTMLSelectElement, HTMLButtonElement, HTMLLabelElement, Node, KeyboardEvent, Event */

/*
 * The icons are drawn for small sizes: a 20 × 20 grid shown at exactly 20px (no
 * scaling, so nothing lands between pixels), built from whole-pixel blocks and
 * 1px outlines centred on pixels, with no thin curves or tiny text.
 * Straight-edged shapes also get crispEdges, so they stay sharp when Windows
 * display scaling (125%, 150%) isn't a whole number.
 */
const ICON = (body: string) =>
  `<svg viewBox="0 0 20 20" width="20" height="20" aria-hidden="true" focusable="false">${body}</svg>`;

/** A solid block covering pixels x…x+w, y…y+h. */
const F = (x: number, y: number, w: number, h: number, opacity = 1) =>
  `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="currentColor"${
    opacity < 1 ? ` opacity="${opacity}"` : ""
  } shape-rendering="crispEdges"/>`;

/** A 1px outline whose outer edge covers pixels x…x+w, y…y+h. */
const O = (x: number, y: number, w: number, h: number) =>
  `<rect x="${x + 0.5}" y="${y + 0.5}" width="${w - 1}" height="${h - 1}" fill="none" stroke="currentColor" stroke-width="1" shape-rendering="crispEdges"/>`;

/** The chart's right edge: a 1px dashed line. */
const EDGE = `<path d="M18.5 2v16" fill="none" stroke="currentColor" stroke-width="1" stroke-dasharray="2 2" opacity="0.6" shape-rendering="crispEdges"/>`;

/** Icons for each toolbar dropdown's options, by select id, then option value. */
export const TOOLBAR_ICONS: Record<string, Record<string, string>> = {
  "background-mode": {
    // Half empty, half solid: follows the theme.
    auto: ICON(O(3, 3, 14, 14) + F(10, 3, 7, 14)),
    white: ICON(O(3, 3, 14, 14)),
    dark: ICON(F(3, 3, 14, 14)),
    // A checkerboard of 4px squares.
    transparent: ICON(
      O(3, 3, 14, 14) + F(4, 4, 4, 4, 0.55) + F(12, 4, 4, 4, 0.55) + F(8, 8, 4, 4, 0.55) +
        F(4, 12, 4, 4, 0.55) + F(12, 12, 4, 4, 0.55)
    ),
    // A solid colour drop.
    custom: ICON(`<path d="M10 2.5C13 6.5 15 9 15 12a5 5 0 0 1-10 0c0-3 2-5.5 5-9.5Z" fill="currentColor"/>`),
  },
  "node-order": {
    // Rows as they come in the data: the colours mixed.
    data: ICON(F(3, 3, 14, 3) + F(3, 7, 14, 3, 0.4) + F(3, 11, 14, 3) + F(3, 15, 14, 3, 0.4)),
    // Rows of the same colour together: two solid, a gap, two faint.
    color: ICON(F(3, 2, 14, 3) + F(3, 6, 14, 3) + F(3, 11, 14, 3, 0.4) + F(3, 15, 14, 3, 0.4)),
  },
  "end-align": {
    // A source column on the left; its end box just after it…
    left: ICON(F(2, 3, 3, 14) + F(9, 6, 3, 8) + EDGE),
    // …or pushed against the chart's right edge.
    justify: ICON(F(2, 3, 3, 14) + F(15, 6, 3, 8) + EDGE),
  },
  "link-color": {
    // Two boxes joined by a link; the link matches the solid box.
    source: ICON(F(2, 3, 3, 14) + O(15, 3, 3, 14) + F(5, 8, 10, 4)),
    target: ICON(O(2, 3, 3, 14) + F(15, 3, 3, 14) + F(5, 8, 10, 4)),
    // Neither: a faint grey link.
    neutral: ICON(O(2, 3, 3, 14) + O(15, 3, 3, 14) + F(5, 8, 10, 4, 0.4)),
  },
};

/** A small solid down-pointing triangle on a 10px grid (a fill stays sharper than a thin chevron). */
const CARET = `<svg class="icon-select-caret" viewBox="0 0 10 10" width="10" height="10" aria-hidden="true" focusable="false"><path d="M2 4h6L5 7Z" fill="currentColor"/></svg>`;

const valueProp = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!;

/** Puts an icon dropdown in place of `select` (which stays, hidden). */
export function attachIconSelect(select: HTMLSelectElement, icons: Record<string, string>) {
  const wrap = document.createElement("span");
  wrap.className = "icon-select";
  select.after(wrap);
  select.hidden = true;
  select.tabIndex = -1;

  const label = document.querySelector<HTMLLabelElement>(`label[for="${select.id}"]`);
  const name = label?.textContent?.trim() || select.title || select.id;

  const button = document.createElement("button");
  button.type = "button";
  button.id = `${select.id}-btn`;
  button.className = "icon-select-btn";
  button.setAttribute("aria-haspopup", "listbox");
  button.setAttribute("aria-expanded", "false");
  if (label) label.htmlFor = button.id; // clicking the text opens the menu

  const menu = document.createElement("div");
  menu.className = "icon-select-menu";
  menu.id = `${select.id}-menu`;
  menu.hidden = true;
  menu.setAttribute("role", "listbox");
  menu.setAttribute("aria-label", name);
  button.setAttribute("aria-controls", menu.id);

  const items = Array.from(select.options).map((o) => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "icon-select-item";
    item.setAttribute("role", "option");
    item.dataset.value = o.value;
    // The icon, with the option's name beside it to explain it.
    item.innerHTML = icons[o.value] ?? "";
    const text = document.createElement("span");
    text.className = "icon-select-text";
    text.textContent = o.text;
    item.appendChild(text);
    item.addEventListener("click", () => {
      choose(o.value);
      close(true);
    });
    menu.appendChild(item);
    return item;
  });

  wrap.append(button, menu);

  const render = () => {
    const opt = select.selectedOptions[0];
    const v = opt?.value ?? "";
    button.innerHTML = (icons[v] ?? "") + CARET;
    button.title = `${name}: ${opt?.text ?? ""}`;
    button.setAttribute("aria-label", `${name}: ${opt?.text ?? ""}`);
    for (const item of items) {
      const on = item.dataset.value === v;
      item.classList.toggle("selected", on);
      item.setAttribute("aria-selected", String(on));
    }
  };

  // Setting the select's value from code (e.g. opening saved settings) updates the icon.
  Object.defineProperty(select, "value", {
    configurable: true,
    get() {
      return valueProp.get!.call(this);
    },
    set(v: string) {
      valueProp.set!.call(this, v);
      render();
    },
  });
  select.addEventListener("change", render);

  const choose = (v: string) => {
    if (v === select.value) return;
    select.value = v;
    select.dispatchEvent(new Event("change"));
  };

  const open = () => {
    menu.hidden = false;
    // Open to the left instead if the menu would run past the pane's right edge.
    menu.classList.remove("align-right");
    if (menu.getBoundingClientRect().right > document.documentElement.clientWidth - 4) {
      menu.classList.add("align-right");
    }
    button.setAttribute("aria-expanded", "true");
    button.classList.add("active");
    (items.find((i) => i.classList.contains("selected")) ?? items[0])?.focus();
  };
  const close = (refocus = false) => {
    if (menu.hidden) return;
    menu.hidden = true;
    button.setAttribute("aria-expanded", "false");
    button.classList.remove("active");
    if (refocus) button.focus();
  };

  button.addEventListener("click", () => (menu.hidden ? open() : close()));
  button.addEventListener("keydown", (e: KeyboardEvent) => {
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      open();
    }
  });
  // Arrow keys move between the icons; Escape closes; Tab leaves.
  menu.addEventListener("keydown", (e: KeyboardEvent) => {
    const i = items.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === "ArrowDown" || e.key === "ArrowRight") {
      e.preventDefault();
      items[(i + 1) % items.length].focus();
    } else if (e.key === "ArrowUp" || e.key === "ArrowLeft") {
      e.preventDefault();
      items[(i - 1 + items.length) % items.length].focus();
    } else if (e.key === "Escape") {
      e.preventDefault();
      close(true);
    } else if (e.key === "Tab") {
      close();
    }
  });
  document.addEventListener(
    "pointerdown",
    (e) => {
      if (!wrap.contains(e.target as Node) && !(label && label.contains(e.target as Node))) close();
    },
    true
  );

  render();
}

/** Turns the toolbar's four dropdowns into icon dropdowns. */
export function setupToolbarIconSelects() {
  for (const [id, icons] of Object.entries(TOOLBAR_ICONS)) {
    const el = document.getElementById(id);
    if (el instanceof HTMLSelectElement) attachIconSelect(el, icons as Record<string, string>);
  }
}
