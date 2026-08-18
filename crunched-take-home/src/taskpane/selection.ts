import { getSelectionMeta } from "../excel/gateway";

export function formatSelectionChip(payload: unknown): string | null {
  if (!payload || typeof payload !== "object") {
    return null;
  }
  const record = payload as Record<string, unknown>;
  const sheet = typeof record.sheet === "string" ? record.sheet : "";
  const address = typeof record.address === "string" ? record.address : "";
  if (!sheet && !address) {
    return null;
  }
  const loc = address.includes("!") ? address : sheet && address ? `${sheet}!${address}` : address || sheet;
  const rows = typeof record.rows === "number" ? record.rows : null;
  const columns = typeof record.columns === "number" ? record.columns : null;
  if (rows !== null && columns !== null) {
    return `${loc} · ${rows}×${columns}`;
  }
  return loc;
}

/** Refreshes on pane/window focus, not on every Excel click. */
export async function refreshSelectionChip(): Promise<string | null> {
  try {
    return formatSelectionChip(await getSelectionMeta());
  } catch {
    return null;
  }
}
