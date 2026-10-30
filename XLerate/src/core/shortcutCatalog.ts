/**
 * Every keyboard-shortcut action the add-in registers, in display order.
 * Must stay in sync with shortcuts.json (enforced by tests/formatTweaks.spec.ts).
 */

export type ShortcutEntry = {
  id: string;
  name: string;
  key: string;
  group: "Cycles" | "Audit" | "Autocolor" | "Formulas" | "Other";
};

export const SHORTCUT_CATALOG: readonly ShortcutEntry[] = [
  { id: "CycleNumberFormat", name: "Number cycle", key: "Ctrl+Shift+1", group: "Cycles" },
  { id: "CycleDateFormat", name: "Date cycle", key: "Ctrl+Shift+2", group: "Cycles" },
  { id: "CycleCellFormat", name: "Cell format cycle", key: "Ctrl+Shift+3", group: "Cycles" },
  { id: "CycleTextStyle", name: "Text style cycle", key: "Ctrl+Shift+4", group: "Cycles" },
  { id: "CyclePercentFormat", name: "% cycle", key: "Ctrl+Shift+5", group: "Cycles" },
  { id: "CycleMultipleFormat", name: "Multiple cycle", key: "Ctrl+Shift+8", group: "Cycles" },
  { id: "TracePrecedents", name: "Trace Precedents", key: "Ctrl+Shift+P", group: "Audit" },
  { id: "TraceDependents", name: "Trace Dependents", key: "Ctrl+Shift+D", group: "Audit" },
  { id: "HorizontalCheck", name: "Horizontal check", key: "Ctrl+Alt+Shift+H", group: "Audit" },
  { id: "AutoColorSelection", name: "Autocolor selection", key: "Ctrl+Alt+A", group: "Autocolor" },
  { id: "InsertCagr", name: "Insert CAGR", key: "Ctrl+Alt+Shift+C", group: "Formulas" },
  { id: "SmartFillRight", name: "Smart Fill Right", key: "Ctrl+Shift+R", group: "Formulas" },
  { id: "BorderTop", name: "Border cycle (top)", key: "Ctrl+Alt+Shift+Up", group: "Other" },
  { id: "BorderBottom", name: "Border cycle (bottom)", key: "Ctrl+Alt+Shift+Down", group: "Other" },
  { id: "BorderLeft", name: "Border cycle (left)", key: "Ctrl+Alt+Shift+Left", group: "Other" },
  { id: "BorderRight", name: "Border cycle (right)", key: "Ctrl+Alt+Shift+Right", group: "Other" },
  { id: "SwitchSign", name: "Switch sign", key: "Ctrl+Shift+N", group: "Other" },
  { id: "ShortcutManager", name: "Shortcut manager", key: "Ctrl+Alt+Shift+M", group: "Other" },
  { id: "PrintShortcuts", name: "Print all shortcuts", key: "Ctrl+Alt+Shift+P", group: "Other" },
  { id: "DecreaseDecimals", name: "Decrease decimals", key: "Ctrl+Alt+Shift+9", group: "Other" },
  { id: "IncreaseDecimals", name: "Increase decimals", key: "Ctrl+Alt+Shift+0", group: "Other" },
  { id: "ResetFormatSettings", name: "Reset format settings", key: "Ctrl+Shift+0", group: "Other" },
];
