/*
 * The chart's title and subtitle: their settings, and the Title section's
 * controls in the task pane. Owns the title settings; the pane reads them with
 * getChartTitle() and puts saved ones back with setChartTitle().
 */

/* global document, HTMLInputElement, HTMLSelectElement, HTMLButtonElement */

import type { ChartTitle } from "./chart";
import { toHex6 } from "./chart";
import { $, addFontOptions, setControl } from "./dom";

export const DEFAULT_TITLE: ChartTitle = {
  text: "",
  subtitle: "",
  size: 20,
  subtitleSize: null, // automatic
  bold: true,
  italic: false,
  fontFamily: null, // same as the labels
  align: "left",
  color: null,
};

/** Value of the "Same as labels" option in the title font dropdown. */
const SAME_AS_LABELS = "";

let chartTitle: ChartTitle = { ...DEFAULT_TITLE };
/** The title colour "Auto" stands for (the chart's text colour), for the colour picker. */
let autoColor: () => string = () => "#000000";

/** The chart's title settings. */
export const getChartTitle = (): ChartTitle => chartTitle;

/** Puts title settings back (e.g. a saved diagram's), with any missing ones from the default. */
export function setChartTitle(t: Partial<ChartTitle>) {
  chartTitle = { ...DEFAULT_TITLE, ...t };
  syncTitleControls();
}

/**
 * Sets up the Title section. `onChange` runs after every change (the chart
 * redraws); `autoColor` gives the colour that "Auto" currently stands for.
 */
export function setupTitleControls(opts: { onChange: () => void; autoColor: () => string }) {
  autoColor = opts.autoColor;
  const text = $<HTMLInputElement>("title-text");
  const sub = $<HTMLInputElement>("subtitle-text");
  const size = $<HTMLInputElement>("title-size");
  const update = (patch: Partial<ChartTitle>) => {
    chartTitle = { ...chartTitle, ...patch };
    syncTitleControls();
    opts.onChange();
  };
  text.addEventListener("input", () => update({ text: text.value }));
  sub.addEventListener("input", () => update({ subtitle: sub.value }));
  size.addEventListener("change", () => {
    const v = Math.round(Number(size.value));
    update({ size: isFinite(v) ? Math.min(40, Math.max(12, v)) : DEFAULT_TITLE.size });
  });

  // Font: "Same as labels" (follows the Labels font) or a font of its own.
  const font = $<HTMLSelectElement>("title-font");
  font.appendChild(
    Object.assign(document.createElement("option"), {
      value: SAME_AS_LABELS,
      textContent: "Same as labels",
    })
  );
  addFontOptions(font);
  font.onchange = () => update({ fontFamily: font.value === SAME_AS_LABELS ? null : font.value });

  const subSize = $<HTMLInputElement>("subtitle-size");
  subSize.addEventListener("change", () => {
    const raw = subSize.value.trim();
    const v = Math.round(Number(raw));
    // Blank (or invalid) = automatic.
    update({ subtitleSize: raw === "" || !isFinite(v) ? null : Math.min(32, Math.max(9, v)) });
  });

  $<HTMLInputElement>("title-bold").onchange = (e) =>
    update({ bold: (e.target as HTMLInputElement).checked });
  $<HTMLInputElement>("title-italic").onchange = (e) =>
    update({ italic: (e.target as HTMLInputElement).checked });
  $<HTMLSelectElement>("title-align").onchange = (e) =>
    update({ align: (e.target as HTMLSelectElement).value as ChartTitle["align"] });
  $<HTMLInputElement>("title-color").addEventListener("input", (e) =>
    update({ color: (e.target as HTMLInputElement).value })
  );
  $("title-color-auto").onclick = () => update({ color: null });
  $("remove-title-btn").onclick = () => update({ ...DEFAULT_TITLE });
  syncTitleControls();
}

/** Writes the title settings into the controls (skipping the one being typed in). */
export function syncTitleControls() {
  const t = chartTitle;
  const set = setControl;
  set("title-text", t.text);
  set("subtitle-text", t.subtitle);
  set("title-size", String(t.size));
  set("title-font", t.fontFamily ?? SAME_AS_LABELS);
  set("subtitle-size", t.subtitleSize === null ? "" : String(t.subtitleSize));
  $<HTMLInputElement>("title-bold").checked = t.bold;
  $<HTMLInputElement>("title-italic").checked = t.italic;
  set("title-align", t.align);
  set("title-color", toHex6(t.color ?? autoColor()));
  $<HTMLButtonElement>("title-color-auto").disabled = t.color === null;
  const any = t.text.trim() !== "" || t.subtitle.trim() !== "";
  $<HTMLButtonElement>("remove-title-btn").disabled =
    JSON.stringify(t) === JSON.stringify(DEFAULT_TITLE);
  // The section heading shows the title, so it's visible while collapsed.
  $("title-note").textContent = any ? t.text.trim() || t.subtitle.trim() : "None";
}
