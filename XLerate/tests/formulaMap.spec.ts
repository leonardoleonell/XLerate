import { describe, expect, it } from "vitest";
import {
  classifyFormulaMap,
  classifyFormulaReferences,
  FORMULA_MAP_COLORS,
  isStillPainted,
  planFillRuns,
  planRestoreFill,
  type FormulaMapCell,
} from "../src/core/formulaMap";

const f = (r1c1: string, a1 = "=A1"): FormulaMapCell => ({ kind: "formula", r1c1, a1 });
const n: FormulaMapCell = { kind: "number" };
const o: FormulaMapCell = { kind: "other" };

describe("classifyFormulaReferences", () => {
  it.each([
    ["=100*1.05", "none"],
    ["=SUM(1,2)+ROUND(3.456, 1)", "none"],
    ['=IF(TRUE,"Sheet2!A1",1E+5)', "none"],
    ["=IFERROR(1/0,#N/A)", "none"],
    ["=A1*2", "local"],
    ["=$B$3+C4", "local"],
    ["=SUM(1:1)", "local"],
    ["=TaxRate*2", "local"],
    ["=SUM(Table1[Revenue])", "local"],
    ["=IFERROR(A1,#REF!)", "local"],
    ["=Inputs!B3*2", "sheet"],
    ["='Balance Sheet'!B3", "sheet"],
    ["=SUM(Table1[Col])+Calcs!A1", "sheet"],
    ["=[Book2.xlsx]Sheet1!A1", "external"],
    ["='C:\\dir\\[Book 2.xlsx]My Sheet'!A1", "external"],
    ["=[1]Inputs!B3", "external"],
  ])("%s -> %s", (formula, expected) => {
    expect(classifyFormulaReferences(formula)).toBe(expected);
  });
});

describe("classifyFormulaMap", () => {
  it("separates copied across, down, both and unique formulas", () => {
    const grid: FormulaMapCell[][] = [
      [f("=R[-1]C"), f("=R[-1]C"), f("=X")],
      [f("=R[-1]C"), f("=Y"), f("=Z")],
    ];
    expect(classifyFormulaMap(grid)).toEqual([
      ["acrossDown", "across", "unique"],
      ["down", "unique", "unique"],
    ]);
  });

  it("splits unique formulas by what they reference", () => {
    const grid: FormulaMapCell[][] = [
      [f("=a", "=[B.xlsx]S!A1"), o, f("=b", "=Inputs!A1"), o, f("=c", "=5*2"), o, f("=d", "=A1")],
    ];
    expect(classifyFormulaMap(grid)[0]).toEqual([
      "uniqueExternal",
      "none",
      "uniqueSheet",
      "none",
      "uniqueNoRefs",
      "none",
      "unique",
    ]);
  });

  it("marks numbers and leaves text/empty untouched", () => {
    expect(classifyFormulaMap([[n, o]])).toEqual([["number", "none"]]);
  });

  it("flags the odd formula in a row copied across", () => {
    const grid: FormulaMapCell[][] = [
      [f("=RC[-1]*2"), f("=RC[-1]*2"), f("=RC[-1]*3"), f("=RC[-1]*2"), f("=RC[-1]*2")],
    ];
    expect(classifyFormulaMap(grid)[0]).toEqual(["across", "across", "unique", "across", "across"]);
  });
});

describe("planFillRuns", () => {
  it("merges equal neighbours into one run and skips untouched cells", () => {
    const cats = classifyFormulaMap([[n, f("=a"), f("=a"), f("=a"), o, n]]);
    expect(planFillRuns(cats, { row: 0, col: 0, rowCount: 1, columnCount: 6 })).toEqual([
      { row: 0, col: 0, length: 1, color: FORMULA_MAP_COLORS.number },
      { row: 0, col: 1, length: 3, color: FORMULA_MAP_COLORS.across },
      { row: 0, col: 5, length: 1, color: FORMULA_MAP_COLORS.number },
    ]);
  });

  it("only paints inside the bounds", () => {
    const cats = classifyFormulaMap([
      [n, n, n],
      [n, f("=a"), n],
      [n, n, n],
    ]);
    expect(planFillRuns(cats, { row: 1, col: 1, rowCount: 1, columnCount: 1 })).toEqual([
      { row: 1, col: 1, length: 1, color: FORMULA_MAP_COLORS.unique },
    ]);
  });
});

describe("planRestoreFill", () => {
  it("clears cells that had no fill", () => {
    expect(planRestoreFill({ pattern: "None", color: "#FFFFFF" })).toEqual({ kind: "clear" });
    expect(planRestoreFill({ pattern: null, color: "#FFFFFF" })).toEqual({ kind: "clear" });
  });

  it("restores solid fills even when the pattern reads back as null", () => {
    expect(planRestoreFill({ pattern: null, color: "#ffff00" })).toEqual({
      kind: "set",
      fill: { pattern: "Solid", color: "#FFFF00" },
    });
  });

  it("keeps pattern fills with their pattern colour", () => {
    expect(
      planRestoreFill({ pattern: "Gray25", color: "#DDEEFF", patternColor: "#000000" })
    ).toEqual({
      kind: "set",
      fill: { pattern: "Gray25", color: "#DDEEFF", patternColor: "#000000" },
    });
  });

  it("detects cells the user recoloured while the map was on", () => {
    expect(isStillPainted({ color: "#b8cce4" }, "#B8CCE4")).toBe(true);
    expect(isStillPainted({ color: "#FF0000" }, "#B8CCE4")).toBe(false);
  });
});
