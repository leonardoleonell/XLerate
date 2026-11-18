/**
 * Formula Map (Arixcel-style): classify every cell of a grid by how its
 * formula relates to its neighbours, then plan the fills that paint the
 * map and the fills that restore the original colours afterwards.
 *
 * Pure logic: the live Excel plumbing lives in taskpane/formulaMap.ts.
 */

export type FormulaMapCell =
  | { kind: "formula"; r1c1: string; a1: string }
  | { kind: "number" }
  | { kind: "other" };

export type FormulaMapCategory =
  | "across"
  | "down"
  | "acrossDown"
  | "unique"
  | "uniqueExternal"
  | "uniqueSheet"
  | "uniqueNoRefs"
  | "number"
  | "none";

export const FORMULA_MAP_COLORS: Readonly<Record<Exclude<FormulaMapCategory, "none">, string>> = {
  across: "#B8CCE4",
  down: "#C4D79B",
  acrossDown: "#C4BD97",
  unique: "#FFC000",
  uniqueExternal: "#00FF80",
  uniqueSheet: "#F79646",
  uniqueNoRefs: "#00FFFF",
  number: "#FFFF66",
};

export type FormulaReferenceKind = "external" | "sheet" | "local" | "none";

/** Drop the contents of "string literals" so they can't look like references. */
function stripStringLiterals(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/g, '""');
}

/**
 * What a formula (A1 notation, en-US as Office.js returns it) points at:
 * another workbook, another sheet, cells/names on this sheet, or nothing.
 */
export function classifyFormulaReferences(formulaA1: string): FormulaReferenceKind {
  const body = stripStringLiterals(formulaA1.startsWith("=") ? formulaA1.slice(1) : formulaA1).replace(
    /#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)/gi,
    "0"
  );

  // [Book.xlsx]Sheet!A1, 'C:\dir\[Book.xlsx]Sheet'!A1 or [1]Sheet!A1.
  if (/'[^']*\[[^\]]+\][^']*'!|\[[^\]]+\][A-Za-z0-9_.\u00C0-\uFFFF]+!/.test(body)) return "external";
  // Sheet!A1 or 'My sheet'!A1 (the quoted name may itself contain "!").
  if (/'(?:[^']|'')+'!|[A-Za-z0-9_.\u00C0-\uFFFF]+!/.test(body)) return "sheet";

  let i = 0;
  while (i < body.length) {
    const ch = body[i];
    if (/[0-9.]/.test(ch)) {
      const num = /^[0-9]*\.?[0-9]+(?:[eE][+-]?[0-9]+)?/.exec(body.slice(i));
      const len = num ? num[0].length : 1;
      // 1:5 is a whole-row reference.
      if (body[i + len] === ":") return "local";
      i += len;
      continue;
    }
    if (ch === "#") {
      // Error literal such as #N/A or #DIV/0!.
      const err = /^#[A-Za-z0-9/]+[!?]?/.exec(body.slice(i));
      i += err ? err[0].length : 1;
      continue;
    }
    if (/[A-Za-z_$\\\u00C0-\uFFFF]/.test(ch)) {
      const ident = /^[A-Za-z_$\\\u00C0-\uFFFF][A-Za-z0-9_.$\\\u00C0-\uFFFF]*/.exec(body.slice(i));
      const word = ident ? ident[0] : ch;
      let j = i + word.length;
      while (body[j] === " ") j += 1;
      const isFunction = body[j] === "(";
      const upper = word.toUpperCase();
      if (!isFunction && upper !== "TRUE" && upper !== "FALSE") return "local";
      i += word.length;
      continue;
    }
    if (ch === "[") return "local"; // structured table reference
    i += 1;
  }
  return "none";
}

function sameFormula(a: FormulaMapCell | undefined, r1c1: string): boolean {
  return a?.kind === "formula" && a.r1c1 === r1c1;
}

/** Classify every cell, looking at its four neighbours inside the grid. */
export function classifyFormulaMap(grid: FormulaMapCell[][]): FormulaMapCategory[][] {
  return grid.map((row, r) =>
    row.map((cell, c): FormulaMapCategory => {
      if (cell.kind === "number") return "number";
      if (cell.kind !== "formula") return "none";

      const across = sameFormula(row[c - 1], cell.r1c1) || sameFormula(row[c + 1], cell.r1c1);
      const down = sameFormula(grid[r - 1]?.[c], cell.r1c1) || sameFormula(grid[r + 1]?.[c], cell.r1c1);
      if (across && down) return "acrossDown";
      if (across) return "across";
      if (down) return "down";

      const refs = classifyFormulaReferences(cell.a1);
      if (refs === "external") return "uniqueExternal";
      if (refs === "sheet") return "uniqueSheet";
      if (refs === "none") return "uniqueNoRefs";
      return "unique";
    })
  );
}

export type FillRun = { row: number; col: number; length: number; color: string };

/**
 * Horizontal runs of equal colour, so a row of 120 "copied across" cells
 * becomes one fill operation instead of 120. Coordinates are relative to
 * `categories`; `bounds` limits painting to a sub-rectangle (the selection
 * inside a grid read with one extra cell of context on each side).
 */
export function planFillRuns(
  categories: FormulaMapCategory[][],
  bounds: { row: number; col: number; rowCount: number; columnCount: number }
): FillRun[] {
  const runs: FillRun[] = [];
  for (let r = bounds.row; r < bounds.row + bounds.rowCount; r += 1) {
    let current: FillRun | null = null;
    for (let c = bounds.col; c < bounds.col + bounds.columnCount; c += 1) {
      const cat = categories[r]?.[c] ?? "none";
      const color = cat === "none" ? null : FORMULA_MAP_COLORS[cat];
      if (current && color === current.color && current.col + current.length === c) {
        current.length += 1;
        continue;
      }
      if (current) runs.push(current);
      current = color ? { row: r, col: c, length: 1, color } : null;
    }
    if (current) runs.push(current);
  }
  return runs;
}

export type OriginalFill = {
  color?: string | null;
  pattern?: string | null;
  patternColor?: string | null;
};

export type RestoreFill = { kind: "clear" } | { kind: "set"; fill: { pattern: string; color: string; patternColor?: string } };

function normColor(color: string | null | undefined): string {
  return (color ?? "").toUpperCase();
}

/**
 * How to put one painted cell back. Excel Desktop reads a solid fill's
 * pattern back as null (see CLAUDE.md), so a null pattern with a non-white
 * colour is treated as Solid; "None" or plain white means "no fill".
 */
export function planRestoreFill(original: OriginalFill): RestoreFill {
  const pattern = original.pattern ?? null;
  const color = normColor(original.color);
  if (pattern === "None" || color === "" || ((pattern === null || pattern === "Solid") && color === "#FFFFFF")) {
    return { kind: "clear" };
  }
  const fill: { pattern: string; color: string; patternColor?: string } = {
    pattern: pattern ?? "Solid",
    color,
  };
  if (pattern !== null && pattern !== "Solid" && original.patternColor) {
    fill.patternColor = original.patternColor;
  }
  return { kind: "set", fill };
}

/** True while the cell still shows the colour the map painted on it. */
export function isStillPainted(current: OriginalFill, painted: string): boolean {
  return normColor(current.color) === normColor(painted);
}
