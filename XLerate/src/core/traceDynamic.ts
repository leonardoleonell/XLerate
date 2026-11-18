/**
 * Resolve dynamic references inside a formula for the trace dialog:
 * OFFSET / INDEX / INDIRECT / lookups -> the cell(s) actually used, and
 * SUMIF-style functions / SUMPRODUCT masks -> the cells that pass the
 * criteria. Pure: Excel data comes through a `DynamicDataProvider`.
 */

import {
  type Ast,
  EvalData,
  Evaluator,
  isArray,
  isError,
  isRef,
  matchesCriteria,
  NeedData,
  parseFormula,
  type RefValue,
  type Scalar,
  truthy,
  type UsedBounds,
  Unsupported,
  type Value,
  rectA1,
  quoteSheet,
  toScalar,
} from "./formulaEval";

export type CellRect = { sheet: string; r1: number; c1: number; r2: number; c2: number };

export type DynamicDataProvider = {
  /** Values (Office.js `values`) for each rectangle, same order. */
  readRects(rects: CellRect[]): Promise<unknown[][][]>;
  /** Used range per sheet (null when the sheet is empty). */
  readUsedRanges(sheets: string[]): Promise<UsedBounds[]>;
  /** Named ranges -> reference (or a constant / null when not a range). */
  resolveNames(
    names: Array<{ name: string; sheet?: string }>,
    formulaSheet: string
  ): Promise<Array<RefValue | Scalar>>;
};

export type DynamicKind = "reference" | "filter";

export type DynamicItem = {
  kind: DynamicKind;
  functionName: string;
  /** Source text of the call, e.g. OFFSET(I$140,0,-1). */
  label: string;
  sheet: string;
  /** Areas the item points at, merged into rectangles. */
  areas: Array<{ r1: number; c1: number; r2: number; c2: number }>;
  /** Number of cells in `areas`. */
  cellCount: number;
  /** Cells in the range that was filtered (filters only). */
  candidateCount?: number;
};

const REFERENCE_FUNCTIONS = new Set([
  "OFFSET",
  "INDEX",
  "INDIRECT",
  "VLOOKUP",
  "HLOOKUP",
  "XLOOKUP",
]);
const FILTER_FUNCTIONS = new Set([
  "SUMIF",
  "SUMIFS",
  "COUNTIF",
  "COUNTIFS",
  "AVERAGEIF",
  "AVERAGEIFS",
  "MAXIFS",
  "MINIFS",
  "SUMPRODUCT",
]);
const QUICK_CHECK =
  /\b(?:_xlfn\.)?(OFFSET|INDEX|INDIRECT|VLOOKUP|HLOOKUP|XLOOKUP|SUMIFS?|COUNTIFS?|AVERAGEIFS?|MAXIFS|MINIFS|SUMPRODUCT)\s*\(/i;

const MAX_ROUNDS = 6;
const MAX_LABEL = 70;

/** Cheap pre-check so formulas without dynamic functions cost nothing. */
export function hasDynamicFunctions(formula: string): boolean {
  return QUICK_CHECK.test(formula);
}

type CallNode = Extract<Ast, { k: "call" }>;

function collectCalls(node: Ast, out: CallNode[]): void {
  switch (node.k) {
    case "call":
      if (REFERENCE_FUNCTIONS.has(node.name) || FILTER_FUNCTIONS.has(node.name)) out.push(node);
      node.args.forEach((a) => collectCalls(a, out));
      return;
    case "un":
    case "pct":
      collectCalls(node.x, out);
      return;
    case "bin":
      collectCalls(node.l, out);
      collectCalls(node.r, out);
      return;
    case "array":
      node.rows.forEach((row) => row.forEach((x) => collectCalls(x, out)));
      return;
    default:
  }
}

function collectStatic(
  node: Ast,
  refs: Array<Extract<Ast, { k: "ref" }>>,
  names: Array<{ name: string; sheet?: string }>
): void {
  switch (node.k) {
    case "ref":
      refs.push(node);
      return;
    case "name":
      names.push({ name: node.name, sheet: node.sheet });
      return;
    case "call":
      node.args.forEach((a) => collectStatic(a, refs, names));
      return;
    case "un":
    case "pct":
      collectStatic(node.x, refs, names);
      return;
    case "bin":
      collectStatic(node.l, refs, names);
      collectStatic(node.r, refs, names);
      return;
    case "array":
      node.rows.forEach((row) => row.forEach((x) => collectStatic(x, refs, names)));
      return;
    default:
  }
}

function makeLabel(formula: string, call: CallNode): string {
  const text = formula.slice(call.s, call.e);
  return text.length > MAX_LABEL ? `${text.slice(0, MAX_LABEL - 1)}â€¦` : text;
}

/** Merge a set of cells into rectangles (row runs, then stacked equal runs). */
export function mergeCells(
  cells: Array<{ r: number; c: number }>
): Array<{ r1: number; c1: number; r2: number; c2: number }> {
  const sorted = [...cells].sort((a, b) => a.r - b.r || a.c - b.c);
  const runs: Array<{ r1: number; c1: number; r2: number; c2: number }> = [];
  for (const cell of sorted) {
    const last = runs[runs.length - 1];
    if (last && last.r1 === cell.r && last.c2 + 1 === cell.c) last.c2 = cell.c;
    else if (!(last && last.r1 === cell.r && last.c2 >= cell.c))
      runs.push({ r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c });
  }
  const merged: typeof runs = [];
  for (const run of runs) {
    const above = merged.find((m) => m.c1 === run.c1 && m.c2 === run.c2 && m.r2 + 1 === run.r1);
    if (above) above.r2 = run.r2;
    else merged.push({ ...run });
  }
  return merged;
}

export function areasAddress(sheet: string, areas: DynamicItem["areas"]): string {
  return `${quoteSheet(sheet)}!${areas.map(rectA1).join(",")}`;
}

class Resolver {
  constructor(
    private readonly ev: Evaluator,
    private readonly formula: string
  ) {}

  item(call: CallNode): DynamicItem | null {
    if (REFERENCE_FUNCTIONS.has(call.name)) return this.referenceItem(call);
    if (call.name === "SUMPRODUCT") return this.sumproductItem(call);
    return this.criteriaItem(call);
  }

  private referenceItem(call: CallNode): DynamicItem | null {
    const v: Value = this.ev.call(call);
    if (!isRef(v) || v.r2 < v.r1 || v.c2 < v.c1) return null;
    return {
      kind: "reference",
      functionName: call.name,
      label: makeLabel(this.formula, call),
      sheet: v.sheet,
      areas: [{ r1: v.r1, c1: v.c1, r2: v.r2, c2: v.c2 }],
      cellCount: (v.r2 - v.r1 + 1) * (v.c2 - v.c1 + 1),
    };
  }

  private refArg(call: CallNode, i: number): RefValue {
    const arg = call.args[i];
    if (!arg || arg.k === "missing") throw new Unsupported("missing range");
    const v = this.ev.eval(arg);
    if (!isRef(v)) throw new Unsupported("criteria range must be a reference");
    return v;
  }

  private criteriaItem(call: CallNode): DynamicItem | null {
    const name = call.name;
    const pairs: Array<{ range: RefValue; criteria: Scalar }> = [];
    let target: RefValue;
    if (name === "SUMIF" || name === "AVERAGEIF" || name === "COUNTIF") {
      const range = this.refArg(call, 0);
      pairs.push({ range, criteria: this.ev.scalar(call.args[1] ?? { k: "missing" }) });
      if (name !== "COUNTIF" && call.args[2] && call.args[2].k !== "missing") {
        const sum = this.refArg(call, 2);
        // Excel sizes sum_range like range, anchored at its top-left cell.
        target = { ...sum, r2: sum.r1 + (range.r2 - range.r1), c2: sum.c1 + (range.c2 - range.c1) };
      } else {
        target = range;
      }
    } else {
      const first = name === "COUNTIFS" ? 0 : 1;
      for (let i = first; i + 1 < call.args.length; i += 2) {
        pairs.push({ range: this.refArg(call, i), criteria: this.ev.scalar(call.args[i + 1]) });
      }
      if (pairs.length === 0) return null;
      target = name === "COUNTIFS" ? pairs[0].range : this.refArg(call, 0);
    }

    const height = target.r2 - target.r1 + 1;
    const width = target.c2 - target.c1 + 1;
    const grids = pairs.map((p) => {
      if (p.range.r2 - p.range.r1 + 1 !== height || p.range.c2 - p.range.c1 + 1 !== width) {
        throw new Unsupported("criteria ranges differ in size");
      }
      return this.ev.deref(p.range);
    });

    const cells: Array<{ r: number; c: number }> = [];
    for (let i = 0; i < height; i += 1) {
      for (let j = 0; j < width; j += 1) {
        if (pairs.every((p, k) => matchesCriteria(grids[k][i]?.[j] ?? null, p.criteria))) {
          cells.push({ r: target.r1 + i, c: target.c1 + j });
        }
      }
    }
    return this.filterItem(call, target.sheet, cells, height * width);
  }

  private sumproductItem(call: CallNode): DynamicItem | null {
    const factors: Ast[] = [];
    const flatten = (node: Ast): void => {
      if (node.k === "bin" && node.op === "*") {
        flatten(node.l);
        flatten(node.r);
      } else factors.push(node);
    };
    call.args.forEach((a) => flatten(a));

    const data: RefValue[] = [];
    const masks: Scalar[][][] = [];
    for (const factor of factors) {
      const v = this.ev.eval(factor);
      if (isRef(v)) {
        if (v.r1 !== v.r2 || v.c1 !== v.c2) data.push(v);
        continue;
      }
      const val = this.ev.valueOf(v);
      if (isArray(val)) masks.push(val.rows);
    }
    if (masks.length === 0 || data.length === 0) return null;

    const target = data[0];
    const height = target.r2 - target.r1 + 1;
    const width = target.c2 - target.c1 + 1;
    for (const mask of masks) {
      if (mask.length !== height || (mask[0]?.length ?? 0) !== width)
        throw new Unsupported("mask size");
    }
    const cells: Array<{ r: number; c: number }> = [];
    for (let i = 0; i < height; i += 1) {
      for (let j = 0; j < width; j += 1) {
        const pass = masks.every((m) => {
          const v = m[i][j];
          if (isError(v)) return false;
          try {
            return truthy(v);
          } catch {
            return false;
          }
        });
        if (pass) cells.push({ r: target.r1 + i, c: target.c1 + j });
      }
    }
    return this.filterItem(call, target.sheet, cells, height * width);
  }

  private filterItem(
    call: CallNode,
    sheet: string,
    cells: Array<{ r: number; c: number }>,
    candidates: number
  ): DynamicItem {
    return {
      kind: "filter",
      functionName: call.name,
      label: makeLabel(this.formula, call),
      sheet,
      areas: mergeCells(cells),
      cellCount: cells.length,
      candidateCount: candidates,
    };
  }
}

/**
 * Find and resolve every dynamic call in `formula`. Items that cannot be
 * resolved (unsupported functions, external links, errors) are skipped.
 */
export async function resolveDynamicReferences(
  formula: string,
  ctx: { sheet: string; row: number; col: number },
  provider: DynamicDataProvider
): Promise<DynamicItem[]> {
  if (!hasDynamicFunctions(formula)) return [];
  let ast: Ast;
  try {
    ast = parseFormula(formula);
  } catch {
    return [];
  }
  const calls: CallNode[] = [];
  collectCalls(ast, calls);
  if (calls.length === 0) return [];

  const data = new EvalData();
  const ev = new Evaluator(data, ctx);

  // Round 0: load everything the formula references literally.
  const staticRefs: Array<Extract<Ast, { k: "ref" }>> = [];
  const staticNames: Array<{ name: string; sheet?: string }> = [];
  collectStatic(ast, staticRefs, staticNames);
  const sheetsNeedingUsed = new Set<string>();
  for (const r of staticRefs) {
    if (r.v.r1 === null || r.v.c1 === null) sheetsNeedingUsed.add(r.v.sheet ?? ctx.sheet);
  }
  try {
    await loadUsed(provider, data, [...sheetsNeedingUsed]);
    await loadNames(provider, data, staticNames, ctx.sheet);
  } catch {
    // Loaded lazily per call below instead.
  }
  const initial: RefValue[] = [];
  for (const r of staticRefs) {
    try {
      initial.push(ev.refFromLiteral(r.v));
    } catch {
      // Clipped away or unsupported; the evaluator will ask again if needed.
    }
  }
  for (const n of staticNames) {
    const v = data.getName(n.name, n.sheet);
    if (isRef(v)) initial.push(v);
  }
  try {
    await loadRefs(provider, data, initial);
  } catch {
    // One bad reference (missing sheet) would sink the batch; go lazy.
  }

  const resolver = new Resolver(ev, formula);
  const items: DynamicItem[] = [];
  for (const call of calls) {
    for (let round = 0; round < MAX_ROUNDS; round += 1) {
      try {
        const item = resolver.item(call);
        if (item) items.push(item);
        break;
      } catch (error) {
        if (!(error instanceof NeedData)) break;
        try {
          if (error.usedSheets.length) await loadUsed(provider, data, error.usedSheets);
          if (error.names.length) await loadNames(provider, data, error.names, ctx.sheet);
          if (error.refs.length) await loadRefs(provider, data, error.refs);
        } catch {
          // e.g. INDIRECT to a sheet that does not exist: skip this call only.
          break;
        }
      }
    }
  }
  return items;
}

async function loadUsed(
  provider: DynamicDataProvider,
  data: EvalData,
  sheets: string[]
): Promise<void> {
  const missing = [...new Set(sheets)].filter((s) => !data.hasUsed(s));
  if (missing.length === 0) return;
  const bounds = await provider.readUsedRanges(missing);
  missing.forEach((s, i) => data.setUsed(s, bounds[i] ?? null));
}

async function loadNames(
  provider: DynamicDataProvider,
  data: EvalData,
  names: Array<{ name: string; sheet?: string }>,
  formulaSheet: string
): Promise<void> {
  const missing = names.filter((n) => data.getName(n.name, n.sheet) === undefined);
  if (missing.length === 0) return;
  const values = await provider.resolveNames(missing, formulaSheet);
  missing.forEach((n, i) => data.setName(n.name, n.sheet, values[i] ?? { error: "#NAME?" }));
}

const MAX_LOAD_CELLS = 200000;

async function loadRefs(
  provider: DynamicDataProvider,
  data: EvalData,
  refs: RefValue[]
): Promise<void> {
  const todo = refs.filter(
    (r) =>
      r.r2 >= r.r1 &&
      r.c2 >= r.c1 &&
      !data.covers(r) &&
      (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) <= MAX_LOAD_CELLS
  );
  if (todo.length === 0) return;
  const rects = todo.map((r) => ({ sheet: r.sheet, r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 }));
  const values = await provider.readRects(rects);
  rects.forEach((rect, i) => {
    data.addBlock({
      ...rect,
      values: (values[i] ?? []).map((row) => row.map((v) => toScalar(v))),
    });
  });
}
