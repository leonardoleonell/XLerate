import { describe, expect, it } from "vitest";
import {
  columnLettersToIndex,
  matchesCriteria,
  parseFormula,
  type RefValue,
  type Scalar,
} from "../src/core/formulaEval";
import {
  areasAddress,
  hasDynamicFunctions,
  mergeCells,
  resolveDynamicReferences,
  type CellRect,
  type DynamicDataProvider,
} from "../src/core/traceDynamic";

/** In-memory workbook: { Sheet: { "A1": value } }. */
function fakeProvider(
  sheets: Record<string, Record<string, unknown>>,
  names: Record<string, RefValue | Scalar> = {}
): DynamicDataProvider & { reads: CellRect[] } {
  const parse = (a1: string): { r: number; c: number } => {
    const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
    return { r: Number(m[2]) - 1, c: columnLettersToIndex(m[1]) };
  };
  const grid = new Map<string, Map<string, unknown>>();
  for (const [sheet, cells] of Object.entries(sheets)) {
    const m = new Map<string, unknown>();
    for (const [a1, v] of Object.entries(cells)) {
      const { r, c } = parse(a1);
      m.set(`${r},${c}`, v);
    }
    grid.set(sheet.toLowerCase(), m);
  }
  const reads: CellRect[] = [];
  return {
    reads,
    async readRects(rects) {
      reads.push(...rects);
      return rects.map((rect) => {
        const m = grid.get(rect.sheet.toLowerCase()) ?? new Map();
        const rows: unknown[][] = [];
        for (let r = rect.r1; r <= rect.r2; r += 1) {
          const row: unknown[] = [];
          for (let c = rect.c1; c <= rect.c2; c += 1) row.push(m.get(`${r},${c}`) ?? "");
          rows.push(row);
        }
        return rows;
      });
    },
    async readUsedRanges(list) {
      return list.map((sheet) => {
        const m = grid.get(sheet.toLowerCase());
        if (!m || m.size === 0) return null;
        const keys = [...m.keys()].map((k) => k.split(",").map(Number));
        return {
          r1: Math.min(...keys.map((k) => k[0])),
          r2: Math.max(...keys.map((k) => k[0])),
          c1: Math.min(...keys.map((k) => k[1])),
          c2: Math.max(...keys.map((k) => k[1])),
        };
      });
    },
    async resolveNames(list) {
      return list.map((n) => names[n.name] ?? { error: "#NAME?" });
    },
  };
}

const at = (sheet: string, row: number, col: number) => ({ sheet, row, col });

describe("formula parser", () => {
  it("parses the Arixcel OFFSET example", () => {
    expect(() =>
      parseFormula(
        "=IF(I$3>4,OFFSET(I$140,0,-4)*OFFSET($H$127,0,4),0)+IF(I$3>1,OFFSET(I$140,0,-1)*OFFSET($H$127,0,1),0)+I$140*$H$127"
      )
    ).not.toThrow();
  });

  it("parses sheet-qualified, whole-column and quoted references", () => {
    expect(() =>
      parseFormula('=SUMIFS(Data!C:C,\'My Sheet\'!A:A,"x",Data!B1:B9,">="&EOMONTH(A1,0))')
    ).not.toThrow();
  });
});

describe("hasDynamicFunctions", () => {
  it("only fires for the functions we resolve", () => {
    expect(hasDynamicFunctions("=SUM(A1:A3)")).toBe(false);
    expect(hasDynamicFunctions("=OFFSET(A1,1,0)")).toBe(true);
    expect(hasDynamicFunctions("=_xlfn.XLOOKUP(A1,B:B,C:C)")).toBe(true);
    expect(hasDynamicFunctions("=sumifs(C1:C3,A1:A3,1)")).toBe(true);
  });
});

describe("resolveDynamicReferences", () => {
  it("resolves OFFSET to the cell actually used (Arixcel print 1)", async () => {
    // I153 = IF(I$3>1, OFFSET(I$140,0,-1)*OFFSET($H$127,0,1), 0) ; I3 = 2
    const provider = fakeProvider({
      Calcs: { I3: 2, H140: 5294.12, I140: 7467, H127: 0.3, I127: 0.4 },
    });
    const items = await resolveDynamicReferences(
      "=IF(I$3>1,OFFSET(I$140,0,-1)*OFFSET($H$127,0,1),0)",
      at("Calcs", 152, 8),
      provider
    );
    expect(items.map((i) => [i.label, areasAddress(i.sheet, i.areas)])).toEqual([
      ["OFFSET(I$140,0,-1)", "Calcs!H140"],
      ["OFFSET($H$127,0,1)", "Calcs!I127"],
    ]);
  });

  it("filters SUMIF to the cells that match the criteria (Arixcel print 5)", async () => {
    const provider = fakeProvider({
      Calcs: {
        G16: "SU",
        G17: "MU",
        G18: "SU",
        G19: "MU",
        G20: "SU",
        G21: "MU",
        I16: 213,
        I17: 2400,
        I18: 160,
        I19: 2400,
        I20: 160,
        I21: 1867,
        I27: 1,
        I28: 2,
        I29: 1,
        I30: 2.1,
        I31: 1,
        I32: 2.1,
        G34: "MU",
      },
    });
    const items = await resolveDynamicReferences(
      "=SUMPRODUCT(--($G$16:$G$21=$G34),I$16:I$21,I$27:I$32)/SUMIF($G$16:$G$21,$G34,I$16:I$21)",
      at("Calcs", 33, 8),
      provider
    );
    expect(
      items.map((i) => [
        i.functionName,
        areasAddress(i.sheet, i.areas),
        i.cellCount,
        i.candidateCount,
      ])
    ).toEqual([
      ["SUMPRODUCT", "Calcs!I17,I19,I21", 3, 6],
      ["SUMIF", "Calcs!I17,I19,I21", 3, 6],
    ]);
  });

  it("resolves INDEX/MATCH and VLOOKUP/XLOOKUP to the returned cell", async () => {
    const provider = fakeProvider({
      S: { A1: "a", A2: "b", A3: "c", B1: 10, B2: 20, B3: 30, D1: "b" },
    });
    const items = await resolveDynamicReferences(
      '=INDEX(B1:B3,MATCH(D1,A1:A3,0))+VLOOKUP("c",A1:B3,2,FALSE)+XLOOKUP(D1,A:A,B:B)',
      at("S", 0, 5),
      provider
    );
    expect(items.map((i) => [i.functionName, areasAddress(i.sheet, i.areas)])).toEqual([
      ["INDEX", "S!B2"],
      ["VLOOKUP", "S!B3"],
      ["XLOOKUP", "S!B2"],
    ]);
  });

  it("handles SUMIFS with date criteria and other sheets", async () => {
    const provider = fakeProvider({
      Data: { A1: 45000, A2: 45031, A3: 45061, B1: "x", B2: "x", B3: "y", C1: 1, C2: 2, C3: 3 },
      Calc: { A1: 45020 },
    });
    const items = await resolveDynamicReferences(
      '=SUMIFS(Data!C:C,Data!A:A,">="&Calc!A1,Data!B:B,"x")',
      at("Calc", 0, 1),
      provider
    );
    expect(items.map((i) => areasAddress(i.sheet, i.areas))).toEqual(["Data!C2"]);
  });

  it("resolves INDIRECT and named ranges", async () => {
    const provider = fakeProvider(
      { S: { A1: "B5", B5: 7, C1: 1, C2: 2 } },
      { Rates: { kind: "ref", sheet: "S", r1: 0, c1: 2, r2: 1, c2: 2 } }
    );
    const items = await resolveDynamicReferences(
      "=INDIRECT(A1)+INDEX(Rates,2)",
      at("S", 0, 5),
      provider
    );
    expect(items.map((i) => areasAddress(i.sheet, i.areas))).toEqual(["S!B5", "S!C2"]);
  });

  it("skips what it cannot evaluate without failing the rest", async () => {
    const provider = fakeProvider({ S: { A1: 1, B1: 2 } });
    const items = await resolveDynamicReferences(
      "=OFFSET(A1,FOOBAR(1),0)+OFFSET(A1,0,1)",
      at("S", 5, 5),
      provider
    );
    expect(items.map((i) => areasAddress(i.sheet, i.areas))).toEqual(["S!B1"]);
  });

  it("does nothing for formulas without dynamic functions", async () => {
    const provider = fakeProvider({ S: { A1: 1 } });
    expect(await resolveDynamicReferences("=A1*2", at("S", 1, 1), provider)).toEqual([]);
    expect(provider.reads).toEqual([]);
  });
});

describe("matchesCriteria", () => {
  it.each<[Scalar, Scalar, boolean]>([
    ["MU", "MU", true],
    ["mu", "MU", true],
    ["SU", "MU", false],
    [5, 5, true],
    ["5", 5, true],
    [5, ">=5", true],
    [4, ">=5", false],
    ["abc", ">=5", false],
    [0, "<>0", false],
    [null, "<>0", true],
    ["North Eng", "North*", true],
    [null, "", true],
    [3, "<>", true],
    [null, "<>", false],
  ])("%j vs %j -> %s", (cell, criteria, expected) => {
    expect(matchesCriteria(cell, criteria)).toBe(expected);
  });
});

describe("mergeCells", () => {
  it("merges runs and stacks equal runs", () => {
    expect(
      mergeCells([
        { r: 0, c: 0 },
        { r: 0, c: 1 },
        { r: 1, c: 0 },
        { r: 1, c: 1 },
        { r: 3, c: 0 },
      ])
    ).toEqual([
      { r1: 0, c1: 0, r2: 1, c2: 1 },
      { r1: 3, c1: 0, r2: 3, c2: 0 },
    ]);
  });
});
