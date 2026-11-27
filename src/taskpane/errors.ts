/*
 * Turns errors from Excel into messages that say what happened and what to do,
 * instead of Excel's generic wording ("The requested resource doesn't exist.").
 */

const MESSAGES: Record<string, string> = {
  InvalidOperationInCellEditMode:
    "Excel is still editing a cell. Press Enter or Esc to finish editing, then try again.",
  ItemNotFound:
    "Couldn't find the sheet, table or cells this chart uses. They may have been renamed, moved or deleted.",
  RequestPayloadSizeLimitExceeded:
    "That's too much data to move between Excel and the add-in at once. Try a smaller chart or fewer cells.",
  ResponsePayloadSizeLimitExceeded:
    "That's too much data to load at once. Select a smaller block of cells.",
  AccessDenied:
    "Excel didn't allow the change. The sheet or workbook may be protected or open read-only.",
};

/** The Excel error code, if `e` is an error from the Excel API. */
export const errorCode = (e: unknown): string | undefined => {
  const code = (e as { code?: unknown } | null)?.code;
  return typeof code === "string" ? code : undefined;
};

/** A message for showing to the user. */
export function friendlyError(e: unknown): string {
  const code = errorCode(e);
  if (code && MESSAGES[code]) return MESSAGES[code];
  return e instanceof Error ? e.message : String(e);
}
