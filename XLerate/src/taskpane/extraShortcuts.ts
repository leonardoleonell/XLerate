/* global Excel, Office, document */

/**
 * Keyboard-shortcut actions added on top of the ribbon commands: percent /
 * multiple cycles, border cycle, decimals, and the shortcut manager / print
 * list. Keyboard handlers receive no ribbon event, so unlike the
 * `finish()`-wrapped ribbon handlers they must not call event.completed().
 */

import { ExcelPortLive } from "../adapters/excelPortLive";
import {
  adjustDecimals,
  MULTIPLE_FORMATS,
  nextBorderWeight,
  nextFormatInCycle,
  PERCENT_FORMATS,
} from "../core/formatTweaks";
import { SHORTCUT_CATALOG } from "../core/shortcutCatalog";
import { runAutoColor as runAutoColorService } from "../services/autoColor.service";
import { runInsertCagr as runInsertCagrService } from "../services/cagr.service";
import { runSwitchSign as runSwitchSignService } from "../services/switchSign.service";
import { readWorkbookFormatSettings } from "./formatSettingsStore";
import { openTraceDialog } from "./traceDialogLauncher";
import { applyFormulaConsistencyAction } from "./workbookActions";

// Above this many cells, reading every number format is too slow; the
// decimals shortcut then follows the active cell like the cycles do.
const PER_CELL_LIMIT = 20000;

type BorderSide = "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight";

function logged(work: () => Promise<unknown>): () => Promise<void> {
  return async () => {
    try {
      await work();
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error("[XLerate shortcut]", error);
    }
  };
}

async function cycleSelectionFormat(cycle: readonly string[]): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const first = range.getCell(0, 0);
    first.load("numberFormat");
    await context.sync();
    const current = String(first.numberFormat[0][0] ?? "General");
    range.numberFormat = nextFormatInCycle(current, cycle) as unknown as string[][];
    await context.sync();
  });
}

async function changeDecimals(delta: 1 | -1): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("cellCount");
    await context.sync();

    if (range.cellCount > PER_CELL_LIMIT) {
      const first = range.getCell(0, 0);
      first.load("numberFormat");
      await context.sync();
      const next = adjustDecimals(String(first.numberFormat[0][0] ?? "General"), delta);
      range.numberFormat = next as unknown as string[][];
    } else {
      range.load("numberFormat");
      await context.sync();
      range.numberFormat = range.numberFormat.map((row) =>
        row.map((format) => adjustDecimals(String(format ?? "General"), delta))
      );
    }
    await context.sync();
  });
}

async function cycleBorder(side: BorderSide): Promise<void> {
  await Excel.run(async (context) => {
    const border = context.workbook.getSelectedRange().format.borders.getItem(side);
    border.load(["style", "weight"]);
    await context.sync();

    const next = nextBorderWeight(border.style as string | null, border.weight as string | null);
    if (next === "None") {
      border.style = "None";
    } else {
      // Style first, then weight/colour — see the border gotcha in CLAUDE.md.
      border.style = "Continuous";
      border.weight = next;
      border.color = "#000000";
    }
    await context.sync();
  });
}

async function currentShortcuts(): Promise<Record<string, string | null>> {
  const defaults: Record<string, string | null> = {};
  for (const item of SHORTCUT_CATALOG) defaults[item.id] = item.key;
  if (!Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1")) return defaults;
  try {
    return { ...defaults, ...(await Office.actions.getShortcuts()) };
  } catch {
    return defaults;
  }
}

async function printShortcuts(): Promise<void> {
  const keys = await currentShortcuts();
  const rows: string[][] = [["Group", "Action", "Shortcut"]];
  for (const item of SHORTCUT_CATALOG) {
    rows.push([item.group, item.name, keys[item.id] ?? "(disabled)"]);
  }

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const existing = sheets.getItemOrNullObject("XLerate Shortcuts");
    await context.sync();
    if (!existing.isNullObject) existing.delete();
    const sheet = sheets.add("XLerate Shortcuts");
    const range = sheet.getRangeByIndexes(0, 0, rows.length, 3);
    range.values = rows;
    sheet.getRange("A1:C1").format.font.bold = true;
    range.format.autofitColumns();
    sheet.activate();
    await context.sync();
  });
}

// ---- Shortcut manager (taskpane section) ----------------------------------

function setManagerStatus(message: string): void {
  const status = document.getElementById("shortcut-manager-status");
  if (status) status.textContent = message;
}

export async function renderShortcutManager(): Promise<void> {
  const body = document.getElementById("shortcut-manager-rows");
  if (!body) return;
  const keys = await currentShortcuts();
  body.innerHTML = "";
  for (const item of SHORTCUT_CATALOG) {
    const row = document.createElement("tr");
    const name = document.createElement("td");
    name.textContent = item.name;
    const keyCell = document.createElement("td");
    const input = document.createElement("input");
    input.type = "text";
    input.value = keys[item.id] ?? "";
    input.placeholder = item.key;
    input.dataset.actionId = item.id;
    keyCell.appendChild(input);
    row.append(name, keyCell);
    body.appendChild(row);
  }
}

async function saveShortcutManager(): Promise<void> {
  if (!Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1")) {
    setManagerStatus("This Excel version cannot change shortcuts at runtime.");
    return;
  }
  const changes: Record<string, string | null> = {};
  document
    .querySelectorAll<HTMLInputElement>("#shortcut-manager-rows input[data-action-id]")
    .forEach((input) => {
      const value = input.value.trim();
      changes[input.dataset.actionId as string] = value === "" ? null : value;
    });
  try {
    await Office.actions.replaceShortcuts(changes);
    setManagerStatus("Shortcuts saved.");
  } catch (error) {
    setManagerStatus(`Could not save: ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function resetShortcutManager(): Promise<void> {
  const reset: Record<string, string | null> = {};
  for (const item of SHORTCUT_CATALOG) reset[item.id] = null;
  try {
    await Office.actions.replaceShortcuts(reset);
    await renderShortcutManager();
    setManagerStatus("Shortcuts reset to defaults.");
  } catch (error) {
    setManagerStatus(`Could not reset: ${error instanceof Error ? error.message : String(error)}`);
  }
}

export function wireShortcutManager(): void {
  document.getElementById("shortcut-manager-save")?.addEventListener("click", () => {
    void saveShortcutManager();
  });
  document.getElementById("shortcut-manager-reset")?.addEventListener("click", () => {
    void resetShortcutManager();
  });
  void renderShortcutManager();
}

export async function openShortcutManager(): Promise<void> {
  await Office.addin.showAsTaskpane();
  await renderShortcutManager();
  document.getElementById("shortcut-manager")?.scrollIntoView();
}

export { printShortcuts };

// ---- Registrations ---------------------------------------------------------

Office.actions.associate("CyclePercentFormat", logged(() => cycleSelectionFormat(PERCENT_FORMATS)));
Office.actions.associate("CycleMultipleFormat", logged(() => cycleSelectionFormat(MULTIPLE_FORMATS)));
Office.actions.associate(
  "TracePrecedents",
  logged(() => openTraceDialog("precedents", readWorkbookFormatSettings().trace))
);
Office.actions.associate(
  "TraceDependents",
  logged(() => openTraceDialog("dependents", readWorkbookFormatSettings().trace))
);
Office.actions.associate("HorizontalCheck", logged(() => applyFormulaConsistencyAction()));
Office.actions.associate(
  "AutoColorSelection",
  logged(() =>
    runAutoColorService(new ExcelPortLive(), readWorkbookFormatSettings().autoColorPalette)
  )
);
Office.actions.associate("InsertCagr", logged(() => runInsertCagrService(new ExcelPortLive())));
Office.actions.associate("SwitchSign", logged(() => runSwitchSignService(new ExcelPortLive())));
Office.actions.associate("BorderTop", logged(() => cycleBorder("EdgeTop")));
Office.actions.associate("BorderBottom", logged(() => cycleBorder("EdgeBottom")));
Office.actions.associate("BorderLeft", logged(() => cycleBorder("EdgeLeft")));
Office.actions.associate("BorderRight", logged(() => cycleBorder("EdgeRight")));
Office.actions.associate("DecreaseDecimals", logged(() => changeDecimals(-1)));
Office.actions.associate("IncreaseDecimals", logged(() => changeDecimals(1)));
Office.actions.associate("ShortcutManager", logged(openShortcutManager));
Office.actions.associate("PrintShortcuts", logged(printShortcuts));
