/*
 * The task pane's light / dark mode: the toggle button in its top right, and
 * the choice remembered for next time (in this browser's storage).
 */

/* global document, window */

import { $ } from "./dom";

const THEME_KEY = "sankeyflow.paneTheme";

/** Your light/dark choice; null until you use the toggle (then the pane is light). */
let themeChoice: "light" | "dark" | null = (() => {
  try {
    const v = window.localStorage.getItem(THEME_KEY);
    return v === "light" || v === "dark" ? v : null;
  } catch {
    return null; // storage can be unavailable; default to light
  }
})();

/** Dark mode: your toggle choice if you've made one, otherwise light. */
export function isDarkTheme(): boolean {
  return themeChoice === "dark";
}

/** Sets up the toggle. `onChange` runs after the pane switches between light and dark. */
export function setupThemeToggle(onChange: () => void) {
  applyPaneTheme();
  $("theme-toggle").onclick = () => {
    themeChoice = isDarkTheme() ? "light" : "dark";
    try {
      window.localStorage.setItem(THEME_KEY, themeChoice);
    } catch {
      // Not remembered next time, but still switches now.
    }
    applyPaneTheme();
    onChange();
  };
}

/** Sun: a solid disc with eight short, separate rays. */
const SUN_ICON =
  `<svg class="sun-icon" viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" focusable="false">` +
  `<circle cx="12" cy="12" r="4.5" fill="currentColor"/>` +
  `<g stroke="currentColor" stroke-width="2" stroke-linecap="round">` +
  `<line x1="12" y1="1.5" x2="12" y2="4"/><line x1="12" y1="20" x2="12" y2="22.5"/>` +
  `<line x1="1.5" y1="12" x2="4" y2="12"/><line x1="20" y1="12" x2="22.5" y2="12"/>` +
  `<line x1="4.6" y1="4.6" x2="6.3" y2="6.3"/><line x1="17.7" y1="17.7" x2="19.4" y2="19.4"/>` +
  `<line x1="4.6" y1="19.4" x2="6.3" y2="17.7"/><line x1="17.7" y1="6.3" x2="19.4" y2="4.6"/>` +
  `</g></svg>`;

/** Sets the pane's theme and the toggle's icon and label. */
function applyPaneTheme() {
  const dark = isDarkTheme();
  document.documentElement.setAttribute("data-theme", dark ? "dark" : "light");
  const btn = $("theme-toggle");
  // The icon shows the current mode: moon in dark mode, sun in light mode.
  if (dark) {
    btn.textContent = "☾";
  } else {
    // Drawn rather than the ☀ glyph, which some fonts render as a starburst.
    btn.innerHTML = SUN_ICON;
  }
  const label = dark ? "Switch to light mode" : "Switch to dark mode";
  btn.title = label;
  btn.setAttribute("aria-label", label);
  btn.setAttribute("aria-pressed", String(dark));
}
