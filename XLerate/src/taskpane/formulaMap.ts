/* global Excel */

import {
  classifyFormulaMap,
  isStillPainted,
  planFillRuns,
  planRestoreFill,
  type FillRun,
  type FormulaMapCell,
  type OriginalFill,
  type RestoreFill,
} from "../core/formulaMap";

/**
 * Formula Map toggle (the Horizontal Check shortcut). First press paints
 * the map, remembering each painted cell's original fill in memory; the
 * next press on the same sheet puts those fills back. The memory lives
 * only while Excel is open — never in document.settings (CLAUDE.md:
 * saveAsync breaks the undo chain).
 *
 * One cell selected maps the sheet's used range; a range maps just it.
 */

const MAX_ROW = 1048575;
const MAX_COL = 16383;
const READ_CHUNK_CELLS = 40000;
const OPS_PER_SYNC = 1500;

type SheetMap = {
  runs: FillRun[]; // absolute sheet coordinates
  originals: Map<string, OriginalFill>; // "row,col" -> fill before painting
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
};

const activeMaps = new Map<string, SheetMap>();

export type FormulaMapResult =
  | { mode: "on"; address: string; paintedCells: number }
  | { mode: "off"; restoredCells: number; keptCells: number }
  | { mode: "nothing" };

type Rect = { rowIndex: number; columnIndex: number; rowCount: number; columnCount: number };

const key = (row: number, col: number): string => `${row},${col}`;

function toMapCell(a1: unknown, r1c1: unknown): FormulaMapCell {
  if (typeof a1 === "string" && a1.startsWith("=")) {
    return { kind: "formula", a1, r1c1: typeof r1c1 === "string" ? r1c1 : a1 };
  }
  if (typeof a1 === "number") return { kind: "number" };
  return { kind: "other" };
}

/** Run `work` over row chunks of `rect`, one sync per chunk. */
async function forEachChunk(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  rect: Rect,
  load: (range: Excel.Range) => () => void
): Promise<void> {
  const rowsPerChunk = Math.max(1, Math.floor(READ_CHUNK_CELLS / rect.columnCount));
  for (let r = 0; r < rect.rowCount; r += rowsPerChunk) {
    const rows = Math.min(rowsPerChunk, rect.rowCount - r);
    const range = sheet.getRangeByIndexes(
      rect.rowIndex + r,
      rect.columnIndex,
      rows,
      rect.columnCount
    );
    const after = load(range);
    await context.sync();
    after();
  }
}

async function readFills(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  rect: Rect
): Promise<OriginalFill[][]> {
  const fills: OriginalFill[][] = [];
  await forEachChunk(context, sheet, rect, (range) => {
    const props = range.getCellProperties({
      format: { fill: { color: true, pattern: true, patternColor: true } },
    });
    return () => {
      for (const row of props.value) {
        fills.push(
          row.map((cell) => ({
            color: cell.format?.fill?.color ?? null,
            pattern: (cell.format?.fill?.pattern as string | undefined) ?? null,
            patternColor: cell.format?.fill?.patternColor ?? null,
          }))
        );
      }
    };
  });
  return fills;
}

async function flushEvery(context: Excel.RequestContext, counter: { ops: number }): Promise<void> {
  counter.ops += 1;
  if (counter.ops % OPS_PER_SYNC === 0) await context.sync();
}

async function paintMap(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  sheetId: string,
  target: Excel.Range
): Promise<FormulaMapResult> {
  const t: Rect = {
    rowIndex: target.rowIndex,
    columnIndex: target.columnIndex,
    rowCount: target.rowCount,
    columnCount: target.columnCount,
  };

  // Read one extra cell around the target so edge cells still see their
  // neighbours (a formula copied from just outside the selection).
  const ctxRow = Math.max(0, t.rowIndex - 1);
  const ctxCol = Math.max(0, t.columnIndex - 1);
  const ctx: Rect = {
    rowIndex: ctxRow,
    columnIndex: ctxCol,
    rowCount: Math.min(MAX_ROW, t.rowIndex + t.rowCount) - ctxRow + 1,
    columnCount: Math.min(MAX_COL, t.columnIndex + t.columnCount) - ctxCol + 1,
  };

  const grid: FormulaMapCell[][] = [];
  await forEachChunk(context, sheet, ctx, (range) => {
    range.load(["formulas", "formulasR1C1"]);
    return () => {
      range.formulas.forEach((row, r) => {
        grid.push(row.map((a1, c) => toMapCell(a1, range.formulasR1C1[r][c])));
      });
    };
  });

  const categories = classifyFormulaMap(grid);
  const localRuns = planFillRuns(categories, {
    row: t.rowIndex - ctx.rowIndex,
    col: t.columnIndex - ctx.columnIndex,
    rowCount: t.rowCount,
    columnCount: t.columnCount,
  });
  if (localRuns.length === 0) return { mode: "nothing" };

  const runs = localRuns.map((run) => ({
    ...run,
    row: run.row + ctx.rowIndex,
    col: run.col + ctx.columnIndex,
  }));

  const before = await readFills(context, sheet, t);
  const originals = new Map<string, OriginalFill>();
  for (const run of runs) {
    for (let c = run.col; c < run.col + run.length; c += 1) {
      originals.set(key(run.row, c), before[run.row - t.rowIndex][c - t.columnIndex]);
    }
  }

  const counter = { ops: 0 };
  for (const run of runs) {
    const range = sheet.getRangeByIndexes(run.row, run.col, 1, run.length);
    // Pattern before colour, or the fill may not render (CLAUDE.md).
    range.format.fill.pattern = "Solid";
    range.format.fill.color = run.color;
    await flushEvery(context, counter);
  }
  await context.sync();

  activeMaps.set(sheetId, { runs, originals, ...t });
  return { mode: "on", address: target.address, paintedCells: originals.size };
}

function sameRestore(a: RestoreFill, b: RestoreFill): boolean {
  if (a.kind !== b.kind) return false;
  if (a.kind === "clear" || b.kind === "clear") return true;
  return (
    a.fill.pattern === b.fill.pattern &&
    a.fill.color === b.fill.color &&
    a.fill.patternColor === b.fill.patternColor
  );
}

async function restoreMap(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  sheetId: string,
  map: SheetMap
): Promise<FormulaMapResult> {
  const current = await readFills(context, sheet, map);
  const counter = { ops: 0 };
  let restored = 0;
  let kept = 0;

  for (const run of map.runs) {
    let start = -1;
    let pending: RestoreFill | null = null;
    const flush = async (end: number): Promise<void> => {
      if (pending === null || start < 0) return;
      const range = sheet.getRangeByIndexes(run.row, start, 1, end - start);
      if (pending.kind === "clear") {
        range.format.fill.clear();
      } else {
        range.format.fill.pattern = pending.fill.pattern as Excel.FillPattern;
        range.format.fill.color = pending.fill.color;
        if (pending.fill.patternColor) range.format.fill.patternColor = pending.fill.patternColor;
      }
      await flushEvery(context, counter);
    };

    for (let c = run.col; c < run.col + run.length; c += 1) {
      const now = current[run.row - map.rowIndex]?.[c - map.columnIndex];
      const original = map.originals.get(key(run.row, c));
      // A cell the user recoloured while the map was on keeps its new colour.
      const plan =
        now && original && isStillPainted(now, run.color) ? planRestoreFill(original) : null;
      if (plan === null) kept += 1;
      else restored += 1;

      if (pending !== null && plan !== null && sameRestore(pending, plan)) continue;
      await flush(c);
      pending = plan;
      start = c;
    }
    await flush(run.col + run.length);
  }
  await context.sync();

  activeMaps.delete(sheetId);
  return { mode: "off", restoredCells: restored, keptCells: kept };
}

export async function toggleFormulaMapAction(): Promise<FormulaMapResult> {
  let out: FormulaMapResult = { mode: "nothing" };
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selected = context.workbook.getSelectedRange();
    const used = sheet.getUsedRangeOrNullObject(true);
    sheet.load("id");
    selected.load(["cellCount", "rowIndex", "columnIndex", "rowCount", "columnCount", "address"]);
    used.load(["isNullObject", "rowIndex", "columnIndex", "rowCount", "columnCount", "address"]);
    await context.sync();

    const existing = activeMaps.get(sheet.id);
    if (existing) {
      out = await restoreMap(context, sheet, sheet.id, existing);
      return;
    }

    if (selected.cellCount === 1 && used.isNullObject) return;
    const target = selected.cellCount === 1 ? used : selected;
    out = await paintMap(context, sheet, sheet.id, target);
  });
  return out;
}
