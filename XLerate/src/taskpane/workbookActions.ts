/* global Excel */

import { computeSmartFillRight, type SmartFillRow } from "../core/smartFillRight";

/**
 * Shared Office.js-using action helpers for features that live on more
 * than one surface (taskpane buttons + ribbon ExecuteFunction handlers).
 * Each exported `apply*Action()` returns a structured result; callers
 * wrap their own status reporting (setStatus in the taskpane, logging
 * in the commands runtime) around it. No DOM access here — the commands
 * runtime has no taskpane DOM to touch.
 */

type CellFormula = string | number | boolean;
type CellValue = string | number | boolean | null;

function asFormulaCell(cell: CellFormula): string | null {
  return typeof cell === "string" && cell.startsWith("=") ? cell : null;
}

type MergedArea = { rowIndex: number; columnIndex: number; rowCount: number; columnCount: number };

function toSmartFillRows(
  values: CellValue[][],
  formulas: CellFormula[][],
  isMerged: (r: number, c: number) => boolean
): SmartFillRow[] {
  return values.map((rowValues, r) =>
    rowValues.map((value, c) => {
      const formula = asFormulaCell(formulas[r][c]);
      const isEmpty = formula === null && (value === null || value === "");
      return { isEmpty, isMerged: isMerged(r, c) };
    })
  );
}

// ---- Smart Fill Right ----

export type SmartFillRightResult =
  | { ok: false; reason: "no_formula" | "merged" | "no_boundary"; address: string }
  | { ok: true; boundaryColumn1Based: number };

/**
 * Fill the active cell's formula to the right up to the boundary found
 * in the 3 rows above (or the used range's right edge, whichever is
 * closer). Skipped with a structured reason when the active cell has no
 * formula, is merged, or no boundary exists.
 */
export async function applySmartFillRightAction(): Promise<SmartFillRightResult> {
  let out: SmartFillRightResult = {
    ok: false,
    reason: "no_boundary",
    address: "",
  };
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheet = workbook.worksheets.getActiveWorksheet();
    const activeCell = workbook.getActiveCell();
    const usedRange = worksheet.getUsedRangeOrNullObject();

    activeCell.load(["rowIndex", "columnIndex", "formulas", "address"]);
    usedRange.load(["isNullObject", "columnIndex", "columnCount"]);
    await context.sync();

    const activeFormula = asFormulaCell(activeCell.formulas[0][0] as CellFormula) ?? "";
    const startRowIndex = Math.max(0, activeCell.rowIndex - 3);
    const sampleRowCount = activeCell.rowIndex - startRowIndex + 1;
    const usedLastColExclusive = usedRange.isNullObject
      ? activeCell.columnIndex + 1
      : usedRange.columnIndex + usedRange.columnCount;
    const sampleColCount = Math.max(
      1,
      Math.min(2000, usedLastColExclusive - activeCell.columnIndex)
    );

    const sample = worksheet.getRangeByIndexes(
      startRowIndex,
      activeCell.columnIndex,
      sampleRowCount,
      sampleColCount
    );
    sample.load(["values", "formulas"]);
    const mergedInSample = sample.getMergedAreasOrNullObject();
    mergedInSample.load("isNullObject");
    await context.sync();

    // The VBA skipped header rows containing merged cells (e.g. a merged
    // "Projections" banner) and refused a merged active cell.
    let mergedAreas: MergedArea[] = [];
    if (!mergedInSample.isNullObject) {
      mergedInSample.areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
      await context.sync();
      mergedAreas = mergedInSample.areas.items.map((area) => ({
        rowIndex: area.rowIndex,
        columnIndex: area.columnIndex,
        rowCount: area.rowCount,
        columnCount: area.columnCount,
      }));
    }
    const isMerged = (r: number, c: number): boolean => {
      const row = startRowIndex + r;
      const col = activeCell.columnIndex + c;
      return mergedAreas.some(
        (area) =>
          row >= area.rowIndex &&
          row < area.rowIndex + area.rowCount &&
          col >= area.columnIndex &&
          col < area.columnIndex + area.columnCount
      );
    };

    const sampleRows = toSmartFillRows(
      sample.values as CellValue[][],
      sample.formulas as CellFormula[][],
      isMerged
    );
    const result = computeSmartFillRight(sampleRows, {
      row: sampleRowCount,
      col: 1,
      formula: activeFormula,
      isMerged: isMerged(sampleRowCount - 1, 0),
    });

    if (!result.ok) {
      const reason: SmartFillRightResult extends { ok: false; reason: infer R }
        ? R
        : never =
        result.reason === "active_cell_must_have_formula"
          ? "no_formula"
          : result.reason === "active_cell_is_merged"
            ? "merged"
            : "no_boundary";
      out = { ok: false, reason, address: activeCell.address };
      return;
    }

    const boundaryAbsCol = activeCell.columnIndex + (result.boundaryCol - 1);
    const destination = worksheet.getRangeByIndexes(
      activeCell.rowIndex,
      activeCell.columnIndex,
      1,
      boundaryAbsCol - activeCell.columnIndex + 1
    );

    // AutoFill (as in the VBA) carries the source cell's formatting along
    // with the formula, matching Excel's own fill right.
    activeCell.autoFill(destination, Excel.AutoFillType.fillDefault);
    await context.sync();
    out = { ok: true, boundaryColumn1Based: boundaryAbsCol + 1 };
  });
  return out;
}
