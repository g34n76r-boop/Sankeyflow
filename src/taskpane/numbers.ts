/*
 * Reading a flow's value from a cell. Excel hands numeric cells over as
 * numbers; this also accepts numbers typed or pasted as text, so a table and
 * a selected range treat the same cells the same way.
 */

/** Currency symbols allowed before or after a number typed as text. */
const CURRENCY = /[£$€¥₹]/g;

/**
 * A cell's value as a number, or NaN if it isn't one. Text is accepted with
 * spaces, a currency symbol, a sign (or brackets for negative) and either
 * thousands style: "1,234.56" and "1.234,56" are both 1234.56. A lone comma
 * followed by exactly three digits ("1,234") is read as thousands, otherwise
 * as a decimal comma ("12,5" is 12.5).
 */
export function parseNumber(v: unknown): number {
  if (typeof v === "number") return isFinite(v) ? v : NaN;
  if (typeof v !== "string") return NaN;

  let s = v.replace(/[\s\u00a0\u202f']/g, "").replace(CURRENCY, "");
  if (!s) return NaN;
  let sign = 1;
  const bracketed = /^\((.*)\)$/.exec(s);
  if (bracketed) {
    sign = -1;
    s = bracketed[1];
  }
  if (/^[-−]/.test(s)) {
    sign = -sign;
    s = s.slice(1);
  } else if (s.startsWith("+")) {
    s = s.slice(1);
  }
  if (!/^[\d.,]+$/.test(s)) return NaN;

  const comma = s.lastIndexOf(",");
  const dot = s.lastIndexOf(".");
  if (comma >= 0 && dot >= 0) {
    // Both used: whichever comes last is the decimal point.
    if (comma > dot) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (comma >= 0) {
    if (/^\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, "");
    else if (s.indexOf(",") === comma) s = s.replace(",", ".");
    else return NaN;
  } else if (dot >= 0 && s.indexOf(".") !== dot) {
    // Several dots: only valid as thousands separators ("1.234.567").
    if (!/^\d{1,3}(\.\d{3})+$/.test(s)) return NaN;
    s = s.replace(/\./g, "");
  }
  if (!/^(\d+(\.\d*)?|\.\d+)$/.test(s)) return NaN;
  return sign * Number(s);
}
