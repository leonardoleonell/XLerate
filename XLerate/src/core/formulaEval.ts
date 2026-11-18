/**
 * A small Excel formula parser and evaluator, just big enough to resolve
 * where dynamic references point (OFFSET, INDEX, INDIRECT, lookups and the
 * criteria of SUMIF-style functions). It is NOT a calculation engine:
 * anything it does not understand throws `Unsupported`, and the caller
 * simply skips that item.
 *
 * Formulas are A1, en-US (what Office.js `formulas` returns).
 * Coordinates are 0-based sheet indexes; r2/c2 are inclusive.
 */

// ---------------------------------------------------------------- values

export type ErrorValue = { error: string };
export type Scalar = number | string | boolean | null | ErrorValue;
export type RefValue = {
  kind: "ref";
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
};
export type ArrayValue = { kind: "array"; rows: Scalar[][] };
export type Value = Scalar | RefValue | ArrayValue;

export class Unsupported extends Error {}

/** Thrown when the evaluator needs cells or names it has not been given yet. */
export class NeedData extends Error {
  constructor(
    readonly refs: RefValue[],
    readonly names: Array<{ name: string; sheet?: string }>,
    readonly usedSheets: string[] = []
  ) {
    super("need data");
  }
}

export const MAX_ROW = 1048575;
export const MAX_COL = 16383;

const ERROR_LITERALS = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#SPILL!",
  "#CALC!",
  "#GETTING_DATA",
];

export function isError(v: unknown): v is ErrorValue {
  return typeof v === "object" && v !== null && "error" in v;
}
export function isRef(v: unknown): v is RefValue {
  return typeof v === "object" && v !== null && (v as RefValue).kind === "ref";
}
export function isArray(v: unknown): v is ArrayValue {
  return typeof v === "object" && v !== null && (v as ArrayValue).kind === "array";
}

/** Normalise a raw Office.js `values` entry. */
export function toScalar(raw: unknown): Scalar {
  if (raw === "" || raw === null || raw === undefined) return null;
  if (typeof raw === "number" || typeof raw === "boolean") return raw;
  if (typeof raw === "string") {
    if (ERROR_LITERALS.includes(raw.toUpperCase())) return { error: raw.toUpperCase() };
    return raw;
  }
  if (isError(raw)) return raw;
  return String(raw);
}

// ---------------------------------------------------------------- A1 helpers

export function columnLettersToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function columnIndexToLetters(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function cellA1(row: number, col: number): string {
  return `${columnIndexToLetters(col)}${row + 1}`;
}

export function rectA1(r: { r1: number; c1: number; r2: number; c2: number }): string {
  const a = cellA1(r.r1, r.c1);
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${cellA1(r.r2, r.c2)}`;
}

export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

// ---------------------------------------------------------------- tokenizer

type Token =
  | { t: "num"; v: number; s: number; e: number }
  | { t: "str"; v: string; s: number; e: number }
  | { t: "bool"; v: boolean; s: number; e: number }
  | { t: "err"; v: string; s: number; e: number }
  | { t: "ref"; v: RefLiteral; s: number; e: number }
  | { t: "name"; v: string; sheet?: string; s: number; e: number }
  | { t: "func"; v: string; s: number; e: number }
  | { t: "op"; v: string; s: number; e: number }
  | { t: "("; s: number; e: number }
  | { t: ")"; s: number; e: number }
  | { t: ","; s: number; e: number }
  | { t: ";"; s: number; e: number }
  | { t: "{"; s: number; e: number }
  | { t: "}"; s: number; e: number };

type RefLiteral = {
  sheet?: string;
  r1: number | null; // null = whole column
  c1: number | null; // null = whole row
  r2: number | null;
  c2: number | null;
};

const SHEET = "(?:'(?:[^']|'')+'|[A-Za-z_\\u00C0-\\uFFFF][A-Za-z0-9_.\\u00C0-\\uFFFF]*)";
const CELL = "\\$?([A-Za-z]{1,3})\\$?([0-9]{1,7})";
const RE_CELL_RANGE = new RegExp(`^(?:(${SHEET})!)?${CELL}(?::${CELL})?`);
const RE_COL_RANGE = new RegExp(
  `^(?:(${SHEET})!)?\\$?([A-Za-z]{1,3}):\\$?([A-Za-z]{1,3})(?![A-Za-z0-9_(])`
);
const RE_ROW_RANGE = new RegExp(`^(?:(${SHEET})!)?\\$?([0-9]{1,7}):\\$?([0-9]{1,7})(?![0-9])`);
const RE_IDENT = /^[A-Za-z_\\À-￿][A-Za-z0-9_.À-￿]*/;
const RE_SHEET_NAME = new RegExp(`^(${SHEET})!([A-Za-z_\\\\][A-Za-z0-9_.]*)`);

function unquote(sheet: string | undefined): string | undefined {
  if (sheet === undefined) return undefined;
  return sheet.startsWith("'") ? sheet.slice(1, -1).replace(/''/g, "'") : sheet;
}

function isIdentChar(ch: string | undefined): boolean {
  return ch !== undefined && /[A-Za-z0-9_.À-￿]/.test(ch);
}

function validCell(col: string, row: string): boolean {
  const c = columnLettersToIndex(col);
  const r = Number(row) - 1;
  return c >= 0 && c <= MAX_COL && r >= 0 && r <= MAX_ROW;
}

export function tokenize(formula: string): Token[] {
  const src = formula.startsWith("=") ? formula.slice(1) : formula;
  const off = formula.length - src.length;
  const out: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    const s = i + off;
    if (ch === " " || ch === "\n" || ch === "\r" || ch === "\t") {
      i += 1;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let text = "";
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            text += '"';
            j += 2;
            continue;
          }
          break;
        }
        text += src[j];
        j += 1;
      }
      out.push({ t: "str", v: text, s, e: j + 1 + off });
      i = j + 1;
      continue;
    }
    if (ch === "#") {
      const lit = ERROR_LITERALS.find((e) => src.slice(i, i + e.length).toUpperCase() === e);
      if (!lit) throw new Unsupported("error literal");
      out.push({ t: "err", v: lit, s, e: s + lit.length });
      i += lit.length;
      continue;
    }
    if (ch === "[") throw new Unsupported("external or structured reference");
    if ("(),;{}".includes(ch)) {
      out.push({ t: ch as "(", s, e: s + 1 });
      i += 1;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === "<=" || two === ">=" || two === "<>") {
      out.push({ t: "op", v: two, s, e: s + 2 });
      i += 2;
      continue;
    }
    if ("+-*/^&=<>%:".includes(ch)) {
      out.push({ t: "op", v: ch, s, e: s + 1 });
      i += 1;
      continue;
    }

    const rest = src.slice(i);
    const prev = src[i - 1];
    if (!isIdentChar(prev)) {
      // Row range 1:5 must be tried before numbers.
      const rowRange = RE_ROW_RANGE.exec(rest);
      if (rowRange) {
        out.push({
          t: "ref",
          v: {
            sheet: unquote(rowRange[1]),
            r1: Number(rowRange[2]) - 1,
            r2: Number(rowRange[3]) - 1,
            c1: null,
            c2: null,
          },
          s,
          e: s + rowRange[0].length,
        });
        i += rowRange[0].length;
        continue;
      }
    }
    if (/[0-9.]/.test(ch)) {
      const m = /^[0-9]*\.?[0-9]+(?:[eE][+-]?[0-9]+)?/.exec(rest);
      if (!m) throw new Unsupported("number");
      out.push({ t: "num", v: Number(m[0]), s, e: s + m[0].length });
      i += m[0].length;
      continue;
    }

    const cellRange = RE_CELL_RANGE.exec(rest);
    if (cellRange) {
      const next = rest[cellRange[0].length];
      const ok =
        !isIdentChar(next) &&
        next !== "(" &&
        validCell(cellRange[2], cellRange[3]) &&
        (cellRange[4] === undefined || validCell(cellRange[4], cellRange[5]));
      if (ok) {
        const r1 = Number(cellRange[3]) - 1;
        const c1 = columnLettersToIndex(cellRange[2]);
        const r2 = cellRange[5] ? Number(cellRange[5]) - 1 : r1;
        const c2 = cellRange[4] ? columnLettersToIndex(cellRange[4]) : c1;
        out.push({
          t: "ref",
          v: { sheet: unquote(cellRange[1]), r1, c1, r2, c2 },
          s,
          e: s + cellRange[0].length,
        });
        i += cellRange[0].length;
        continue;
      }
    }
    const colRange = RE_COL_RANGE.exec(rest);
    if (colRange) {
      out.push({
        t: "ref",
        v: {
          sheet: unquote(colRange[1]),
          r1: null,
          r2: null,
          c1: columnLettersToIndex(colRange[2]),
          c2: columnLettersToIndex(colRange[3]),
        },
        s,
        e: s + colRange[0].length,
      });
      i += colRange[0].length;
      continue;
    }
    const sheetName = RE_SHEET_NAME.exec(rest);
    if (sheetName) {
      out.push({
        t: "name",
        v: sheetName[2],
        sheet: unquote(sheetName[1]),
        s,
        e: s + sheetName[0].length,
      });
      i += sheetName[0].length;
      continue;
    }
    const ident = RE_IDENT.exec(rest);
    if (ident) {
      const word = ident[0];
      let j = i + word.length;
      while (src[j] === " ") j += 1;
      if (src[j] === "(") {
        out.push({
          t: "func",
          v: word.toUpperCase().replace(/^_XLFN\.|^_XLWS\./, ""),
          s,
          e: s + word.length,
        });
      } else if (word.toUpperCase() === "TRUE" || word.toUpperCase() === "FALSE") {
        out.push({ t: "bool", v: word.toUpperCase() === "TRUE", s, e: s + word.length });
      } else {
        out.push({ t: "name", v: word, s, e: s + word.length });
      }
      i += word.length;
      continue;
    }
    throw new Unsupported(`character ${ch}`);
  }
  return out;
}

// ---------------------------------------------------------------- parser

export type Ast =
  | { k: "num"; v: number }
  | { k: "str"; v: string }
  | { k: "bool"; v: boolean }
  | { k: "err"; v: string }
  | { k: "missing" }
  | { k: "ref"; v: RefLiteral }
  | { k: "name"; name: string; sheet?: string }
  | { k: "call"; name: string; args: Ast[]; s: number; e: number }
  | { k: "un"; op: string; x: Ast }
  | { k: "pct"; x: Ast }
  | { k: "bin"; op: string; l: Ast; r: Ast }
  | { k: "array"; rows: Ast[][] };

const BINARY_PRECEDENCE: Record<string, number> = {
  "=": 1,
  "<>": 1,
  "<": 1,
  ">": 1,
  "<=": 1,
  ">=": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5,
  ":": 8,
};

export function parseFormula(formula: string): Ast {
  const tokens = tokenize(formula);
  let p = 0;
  const peek = (): Token | undefined => tokens[p];

  const parseExpr = (minPrec: number): Ast => {
    let left = parseUnary();
    for (;;) {
      const tok = peek();
      if (!tok || tok.t !== "op") break;
      if (tok.v === "%") {
        p += 1;
        left = { k: "pct", x: left };
        continue;
      }
      const prec = BINARY_PRECEDENCE[tok.v];
      if (prec === undefined || prec < minPrec) break;
      p += 1;
      // ^ is left-associative in Excel too.
      const right = parseExpr(prec + 1);
      left = { k: "bin", op: tok.v, l: left, r: right };
    }
    return left;
  };

  const parseUnary = (): Ast => {
    const tok = peek();
    if (tok?.t === "op" && (tok.v === "-" || tok.v === "+")) {
      p += 1;
      // Unary minus binds tighter than ^ in Excel (=-2^2 is 4).
      const x = parseUnary();
      return tok.v === "-" ? { k: "un", op: "-", x } : x;
    }
    return parsePrimary();
  };

  const parsePrimary = (): Ast => {
    const tok = tokens[p];
    if (!tok) throw new Unsupported("unexpected end");
    p += 1;
    switch (tok.t) {
      case "num":
        return { k: "num", v: tok.v };
      case "str":
        return { k: "str", v: tok.v };
      case "bool":
        return { k: "bool", v: tok.v };
      case "err":
        return { k: "err", v: tok.v };
      case "ref":
        return { k: "ref", v: tok.v };
      case "name":
        return { k: "name", name: tok.v, sheet: tok.sheet };
      case "(": {
        const inner = parseExpr(0);
        if (tokens[p]?.t !== ")") throw new Unsupported("missing )");
        p += 1;
        return inner;
      }
      case "{": {
        const rows: Ast[][] = [[]];
        for (;;) {
          rows[rows.length - 1].push(parseUnary());
          const sep = tokens[p];
          p += 1;
          if (sep?.t === "}") break;
          if (sep?.t === ",") continue;
          if (sep?.t === ";") {
            rows.push([]);
            continue;
          }
          throw new Unsupported("array constant");
        }
        return { k: "array", rows };
      }
      case "func": {
        if (tokens[p]?.t !== "(") throw new Unsupported("function without (");
        p += 1;
        const args: Ast[] = [];
        if (tokens[p]?.t === ")") {
          p += 1;
          return { k: "call", name: tok.v, args, s: tok.s, e: tokens[p - 1].e };
        }
        for (;;) {
          const next = tokens[p];
          if (next?.t === "," || next?.t === ")") args.push({ k: "missing" });
          else args.push(parseExpr(0));
          const sep = tokens[p];
          p += 1;
          if (sep?.t === ",") {
            if (tokens[p]?.t === ")") {
              args.push({ k: "missing" });
              p += 1;
              break;
            }
            continue;
          }
          if (sep?.t === ")") break;
          throw new Unsupported("bad argument list");
        }
        return { k: "call", name: tok.v, args, s: tok.s, e: tokens[p - 1].e };
      }
      default:
        throw new Unsupported(`unexpected ${tok.t}`);
    }
  };

  const ast = parseExpr(0);
  if (p !== tokens.length) throw new Unsupported("trailing tokens");
  return ast;
}

// ---------------------------------------------------------------- data access

export type UsedBounds = { r1: number; c1: number; r2: number; c2: number } | null;

export type EvalContext = {
  sheet: string; // sheet holding the formula
  row: number;
  col: number;
};

type Block = { sheet: string; r1: number; c1: number; r2: number; c2: number; values: Scalar[][] };

/**
 * Cells, names and used ranges loaded so far. The evaluator asks for
 * what it lacks by throwing NeedData; the driver loads it and retries.
 */
export class EvalData {
  private blocks: Block[] = [];
  private names = new Map<string, RefValue | Scalar>();
  private used = new Map<string, UsedBounds>();

  addBlock(block: Block): void {
    this.blocks.push(block);
  }
  setName(name: string, sheet: string | undefined, value: RefValue | Scalar): void {
    this.names.set(nameKey(name, sheet), value);
  }
  setUsed(sheet: string, bounds: UsedBounds): void {
    this.used.set(sheet.toLowerCase(), bounds);
  }
  hasUsed(sheet: string): boolean {
    return this.used.has(sheet.toLowerCase());
  }
  getUsed(sheet: string): UsedBounds | undefined {
    return this.used.get(sheet.toLowerCase());
  }
  getName(name: string, sheet: string | undefined): RefValue | Scalar | undefined {
    return this.names.get(nameKey(name, sheet));
  }
  cell(sheet: string, r: number, c: number): Scalar | undefined {
    const lower = sheet.toLowerCase();
    for (const b of this.blocks) {
      if (b.sheet.toLowerCase() === lower && r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2) {
        return b.values[r - b.r1]?.[c - b.c1] ?? null;
      }
    }
    return undefined;
  }
  covers(ref: RefValue): boolean {
    const lower = ref.sheet.toLowerCase();
    return this.blocks.some(
      (b) =>
        b.sheet.toLowerCase() === lower &&
        ref.r1 >= b.r1 &&
        ref.r2 <= b.r2 &&
        ref.c1 >= b.c1 &&
        ref.c2 <= b.c2
    );
  }
}

function nameKey(name: string, sheet: string | undefined): string {
  return `${(sheet ?? "").toLowerCase()}!${name.toLowerCase()}`;
}

// ---------------------------------------------------------------- evaluator

const MAX_DEREF_CELLS = 200000;

export class Evaluator {
  constructor(
    readonly data: EvalData,
    readonly ctx: EvalContext
  ) {}

  /** Resolve a literal reference, clipping whole rows/columns to the used range. */
  refFromLiteral(lit: RefLiteral): RefValue {
    const sheet = lit.sheet ?? this.ctx.sheet;
    if (lit.r1 === null || lit.c1 === null) {
      if (!this.data.hasUsed(sheet)) throw new NeedData([], [], [sheet]);
      const used = this.data.getUsed(sheet);
      if (!used) return { kind: "ref", sheet, r1: 0, c1: 0, r2: -1, c2: -1 };
      if (lit.r1 === null) {
        return { kind: "ref", sheet, r1: used.r1, r2: used.r2, c1: lit.c1 ?? 0, c2: lit.c2 ?? 0 };
      }
      return { kind: "ref", sheet, r1: lit.r1, r2: lit.r2 ?? lit.r1, c1: used.c1, c2: used.c2 };
    }
    return {
      kind: "ref",
      sheet,
      r1: Math.min(lit.r1, lit.r2 ?? lit.r1),
      r2: Math.max(lit.r1, lit.r2 ?? lit.r1),
      c1: Math.min(lit.c1, lit.c2 ?? lit.c1),
      c2: Math.max(lit.c1, lit.c2 ?? lit.c1),
    };
  }

  deref(ref: RefValue): Scalar[][] {
    if (ref.r2 < ref.r1 || ref.c2 < ref.c1) return [];
    const count = (ref.r2 - ref.r1 + 1) * (ref.c2 - ref.c1 + 1);
    if (count > MAX_DEREF_CELLS) throw new Unsupported("range too large");
    if (!this.data.covers(ref)) throw new NeedData([ref], []);
    const rows: Scalar[][] = [];
    for (let r = ref.r1; r <= ref.r2; r += 1) {
      const row: Scalar[] = [];
      for (let c = ref.c1; c <= ref.c2; c += 1) row.push(this.data.cell(ref.sheet, r, c) ?? null);
      rows.push(row);
    }
    return rows;
  }

  /** Evaluate to a value that may still be a reference. */
  eval(node: Ast): Value {
    switch (node.k) {
      case "num":
        return node.v;
      case "str":
        return node.v;
      case "bool":
        return node.v;
      case "err":
        return { error: node.v };
      case "missing":
        return null;
      case "ref":
        return this.refFromLiteral(node.v);
      case "name": {
        const v = this.data.getName(node.name, node.sheet);
        if (v === undefined) throw new NeedData([], [{ name: node.name, sheet: node.sheet }]);
        return v;
      }
      case "array":
        return { kind: "array", rows: node.rows.map((row) => row.map((x) => this.scalar(x))) };
      case "un":
        return this.map1(this.eval(node.x), (v) => {
          const n = toNumber(v);
          return isError(n) ? n : -n;
        });
      case "pct":
        return this.map1(this.eval(node.x), (v) => {
          const n = toNumber(v);
          return isError(n) ? n : n / 100;
        });
      case "bin":
        return this.binary(node);
      case "call":
        return this.call(node);
    }
  }

  /** Evaluate to a single scalar (top-left of a range or array). */
  scalar(node: Ast): Scalar {
    return firstScalar(this.valueOf(this.eval(node)));
  }

  /** Dereference refs; keep arrays. */
  valueOf(v: Value): Scalar | ArrayValue {
    if (isRef(v)) {
      const rows = this.deref(v);
      if (rows.length === 1 && rows[0].length === 1) return rows[0][0];
      return { kind: "array", rows };
    }
    return v;
  }

  toRows(v: Value): Scalar[][] {
    const val = this.valueOf(v);
    return isArray(val) ? val.rows : [[val]];
  }

  private map1(v: Value, f: (s: Scalar) => Scalar): Value {
    const val = this.valueOf(v);
    if (isArray(val)) return { kind: "array", rows: val.rows.map((row) => row.map(f)) };
    return f(val);
  }

  private binary(node: Extract<Ast, { k: "bin" }>): Value {
    if (node.op === ":") {
      const l = this.eval(node.l);
      const r = this.eval(node.r);
      if (!isRef(l) || !isRef(r) || l.sheet.toLowerCase() !== r.sheet.toLowerCase())
        throw new Unsupported("range op");
      return {
        kind: "ref",
        sheet: l.sheet,
        r1: Math.min(l.r1, r.r1),
        c1: Math.min(l.c1, r.c1),
        r2: Math.max(l.r2, r.r2),
        c2: Math.max(l.c2, r.c2),
      };
    }
    const l = this.valueOf(this.eval(node.l));
    const r = this.valueOf(this.eval(node.r));
    return broadcast(l, r, (a, b) => applyOp(node.op, a, b));
  }

  private arg(args: Ast[], i: number): Ast {
    return args[i] ?? { k: "missing" };
  }

  private has(args: Ast[], i: number): boolean {
    return args[i] !== undefined && args[i].k !== "missing";
  }

  private num(args: Ast[], i: number, fallback?: number): number {
    if (!this.has(args, i)) {
      if (fallback === undefined) throw new Unsupported("missing argument");
      return fallback;
    }
    const n = toNumber(this.scalar(args[i]));
    if (isError(n)) throw new Unsupported("error argument");
    return n;
  }

  private refArg(args: Ast[], i: number): RefValue {
    const v = this.eval(this.arg(args, i));
    if (!isRef(v)) throw new Unsupported("expected a reference");
    return v;
  }

  private flat(args: Ast[]): Scalar[] {
    const out: Scalar[] = [];
    for (const a of args) {
      if (a.k === "missing") continue;
      for (const row of this.toRows(this.eval(a))) out.push(...row);
    }
    return out;
  }

  call(node: Extract<Ast, { k: "call" }>): Value {
    const a = node.args;
    switch (node.name) {
      case "OFFSET": {
        const base = this.refArg(a, 0);
        const r1 = base.r1 + Math.trunc(this.num(a, 1, 0));
        const c1 = base.c1 + Math.trunc(this.num(a, 2, 0));
        const h = this.has(a, 3) ? Math.trunc(this.num(a, 3)) : base.r2 - base.r1 + 1;
        const w = this.has(a, 4) ? Math.trunc(this.num(a, 4)) : base.c2 - base.c1 + 1;
        if (h < 1 || w < 1) throw new Unsupported("negative OFFSET size");
        const ref: RefValue = {
          kind: "ref",
          sheet: base.sheet,
          r1,
          c1,
          r2: r1 + h - 1,
          c2: c1 + w - 1,
        };
        if (ref.r1 < 0 || ref.c1 < 0 || ref.r2 > MAX_ROW || ref.c2 > MAX_COL)
          return { error: "#REF!" };
        return ref;
      }
      case "INDEX": {
        const first = this.eval(this.arg(a, 0));
        const rowN = Math.trunc(this.num(a, 1, 0));
        const colN = this.has(a, 2) ? Math.trunc(this.num(a, 2)) : undefined;
        if (this.has(a, 3)) throw new Unsupported("INDEX area");
        if (isRef(first)) return indexRef(first, rowN, colN);
        const rows = this.toRows(first);
        const height = rows.length;
        const width = rows[0]?.length ?? 0;
        let r = rowN;
        let c = colN ?? 1;
        if (colN === undefined && height === 1) {
          c = rowN;
          r = 1;
        }
        if (r < 1 || c < 1 || r > height || c > width) return { error: "#REF!" };
        return rows[r - 1][c - 1];
      }
      case "INDIRECT": {
        if (this.has(a, 1)) {
          const a1 = this.scalar(a[1]);
          if (a1 === false || a1 === 0) throw new Unsupported("INDIRECT R1C1");
        }
        const text = this.scalar(this.arg(a, 0));
        if (typeof text !== "string") throw new Unsupported("INDIRECT text");
        const parsed = parseFormula(`=${text}`);
        if (parsed.k !== "ref" && !(parsed.k === "name")) return { error: "#REF!" };
        return this.eval(parsed);
      }
      case "MATCH":
      case "XMATCH": {
        const lookup = this.scalar(this.arg(a, 0));
        const list = this.toRows(this.eval(this.arg(a, 1)));
        const flatList = list.length === 1 ? list[0] : list.map((row) => row[0]);
        const mode =
          node.name === "MATCH"
            ? this.has(a, 2)
              ? this.num(a, 2)
              : 1
            : this.has(a, 2)
              ? this.num(a, 2)
              : 0;
        const pos =
          node.name === "MATCH"
            ? matchPosition(lookup, flatList, mode)
            : xmatchPosition(lookup, flatList, mode, this.has(a, 3) ? this.num(a, 3) : 1);
        return pos < 0 ? { error: "#N/A" } : pos + 1;
      }
      case "VLOOKUP":
      case "HLOOKUP": {
        const lookup = this.scalar(this.arg(a, 0));
        const table = this.refArg(a, 1);
        const idx = Math.trunc(this.num(a, 2));
        const approx = this.has(a, 3) ? truthy(this.scalar(a[3])) : true;
        const vertical = node.name === "VLOOKUP";
        const keys = vertical
          ? this.deref({ ...table, c2: table.c1 }).map((row) => row[0])
          : (this.deref({ ...table, r2: table.r1 })[0] ?? []);
        const pos = matchPosition(lookup, keys, approx ? 1 : 0);
        if (pos < 0) return { error: "#N/A" };
        if (vertical) {
          if (idx < 1 || table.c1 + idx - 1 > table.c2) return { error: "#REF!" };
          return {
            kind: "ref",
            sheet: table.sheet,
            r1: table.r1 + pos,
            r2: table.r1 + pos,
            c1: table.c1 + idx - 1,
            c2: table.c1 + idx - 1,
          };
        }
        if (idx < 1 || table.r1 + idx - 1 > table.r2) return { error: "#REF!" };
        return {
          kind: "ref",
          sheet: table.sheet,
          r1: table.r1 + idx - 1,
          r2: table.r1 + idx - 1,
          c1: table.c1 + pos,
          c2: table.c1 + pos,
        };
      }
      case "XLOOKUP": {
        const lookup = this.scalar(this.arg(a, 0));
        const lookupRef = this.refArg(a, 1);
        const ret = this.refArg(a, 2);
        const keysRows = this.deref(lookupRef);
        const vertical = keysRows.length > 1 || (keysRows[0]?.length ?? 0) === 1;
        const keys = vertical ? keysRows.map((row) => row[0]) : keysRows[0];
        const mode = this.has(a, 4) ? this.num(a, 4) : 0;
        const search = this.has(a, 5) ? this.num(a, 5) : 1;
        const pos = xmatchPosition(lookup, keys, mode, search);
        if (pos < 0) {
          if (this.has(a, 3)) return this.eval(a[3]);
          return { error: "#N/A" };
        }
        return vertical
          ? {
              kind: "ref",
              sheet: ret.sheet,
              r1: ret.r1 + pos,
              r2: ret.r1 + pos,
              c1: ret.c1,
              c2: ret.c2,
            }
          : {
              kind: "ref",
              sheet: ret.sheet,
              r1: ret.r1,
              r2: ret.r2,
              c1: ret.c1 + pos,
              c2: ret.c1 + pos,
            };
      }
      case "CHOOSE": {
        const n = Math.trunc(this.num(a, 0));
        if (n < 1 || n >= a.length) return { error: "#VALUE!" };
        return this.eval(a[n]);
      }
      case "IF": {
        const cond = this.scalar(this.arg(a, 0));
        if (isError(cond)) return cond;
        if (truthy(cond)) return this.has(a, 1) ? this.eval(a[1]) : 0;
        return a.length > 2 ? (this.has(a, 2) ? this.eval(a[2]) : 0) : false;
      }
      case "IFERROR":
      case "IFNA": {
        const v = this.eval(this.arg(a, 0));
        const s = isRef(v) ? firstScalar(this.valueOf(v)) : isArray(v) ? firstScalar(v) : v;
        const bad = isError(s) && (node.name === "IFERROR" || s.error === "#N/A");
        return bad ? this.eval(this.arg(a, 1)) : v;
      }
      case "AND":
      case "OR": {
        const vals = this.flat(a).filter((v) => v !== null && typeof v !== "string");
        const err = vals.find(isError);
        if (err) return err;
        return node.name === "AND" ? vals.every(truthy) : vals.some(truthy);
      }
      case "NOT":
        return !truthy(this.scalar(this.arg(a, 0)));
      case "ROW":
      case "COLUMN": {
        if (!this.has(a, 0)) return node.name === "ROW" ? this.ctx.row + 1 : this.ctx.col + 1;
        const ref = this.refArg(a, 0);
        return node.name === "ROW" ? ref.r1 + 1 : ref.c1 + 1;
      }
      case "ROWS":
      case "COLUMNS": {
        const v = this.eval(this.arg(a, 0));
        if (isRef(v)) return node.name === "ROWS" ? v.r2 - v.r1 + 1 : v.c2 - v.c1 + 1;
        const rows = this.toRows(v);
        return node.name === "ROWS" ? rows.length : (rows[0]?.length ?? 0);
      }
      case "SUM":
      case "MIN":
      case "MAX":
      case "AVERAGE":
      case "COUNT":
      case "COUNTA":
      case "PRODUCT": {
        const vals = this.flat(a);
        if (node.name === "COUNTA") return vals.filter((v) => v !== null).length;
        const err = vals.find(isError);
        if (err) return err;
        const nums = vals.filter((v): v is number => typeof v === "number");
        if (node.name === "COUNT") return nums.length;
        if (node.name === "SUM") return nums.reduce((x, y) => x + y, 0);
        if (node.name === "PRODUCT") return nums.reduce((x, y) => x * y, 1);
        if (node.name === "AVERAGE")
          return nums.length ? nums.reduce((x, y) => x + y, 0) / nums.length : { error: "#DIV/0!" };
        if (nums.length === 0) return 0;
        return node.name === "MIN" ? Math.min(...nums) : Math.max(...nums);
      }
      case "ABS":
        return Math.abs(this.num(a, 0));
      case "INT":
        return Math.floor(this.num(a, 0));
      case "MOD": {
        const n = this.num(a, 0);
        const d = this.num(a, 1);
        return d === 0 ? { error: "#DIV/0!" } : n - d * Math.floor(n / d);
      }
      case "ROUND":
      case "ROUNDUP":
      case "ROUNDDOWN": {
        const n = this.num(a, 0);
        const f = 10 ** this.num(a, 1, 0);
        const scaled = n * f;
        if (node.name === "ROUND")
          return (Math.sign(scaled) * Math.round(Math.abs(scaled) + 1e-9)) / f;
        if (node.name === "ROUNDUP")
          return (Math.sign(scaled) * Math.ceil(Math.abs(scaled) - 1e-9)) / f;
        return (Math.sign(scaled) * Math.floor(Math.abs(scaled) + 1e-9)) / f;
      }
      case "DATE":
        return dateSerial(this.num(a, 0), this.num(a, 1), this.num(a, 2));
      case "YEAR":
      case "MONTH":
      case "DAY": {
        const d = serialToDate(this.num(a, 0));
        return node.name === "YEAR" ? d.y : node.name === "MONTH" ? d.m : d.d;
      }
      case "EOMONTH":
      case "EDATE": {
        const d = serialToDate(this.num(a, 0));
        const months = Math.trunc(this.num(a, 1));
        if (node.name === "EOMONTH") return dateSerial(d.y, d.m + months + 1, 0);
        const last = serialToDate(dateSerial(d.y, d.m + months + 1, 0)).d;
        return dateSerial(d.y, d.m + months, Math.min(d.d, last));
      }
      case "LEFT":
      case "RIGHT": {
        const s = toText(this.scalar(this.arg(a, 0)));
        const n = this.has(a, 1) ? this.num(a, 1) : 1;
        return node.name === "LEFT" ? s.slice(0, n) : s.slice(Math.max(0, s.length - n));
      }
      case "MID": {
        const s = toText(this.scalar(this.arg(a, 0)));
        return s.substr(this.num(a, 1) - 1, this.num(a, 2));
      }
      case "LEN":
        return toText(this.scalar(this.arg(a, 0))).length;
      case "UPPER":
        return toText(this.scalar(this.arg(a, 0))).toUpperCase();
      case "LOWER":
        return toText(this.scalar(this.arg(a, 0))).toLowerCase();
      case "TRIM":
        return toText(this.scalar(this.arg(a, 0)))
          .trim()
          .replace(/ +/g, " ");
      case "CONCATENATE":
      case "CONCAT":
        return this.flat(a).map(toText).join("");
      case "VALUE": {
        const n = toNumber(this.scalar(this.arg(a, 0)));
        return n;
      }
      default:
        throw new Unsupported(`function ${node.name}`);
    }
  }
}

function indexRef(ref: RefValue, rowN: number, colN: number | undefined): Value {
  const height = ref.r2 - ref.r1 + 1;
  const width = ref.c2 - ref.c1 + 1;
  let r = rowN;
  let c = colN ?? 1;
  if (colN === undefined && height === 1 && width > 1) {
    c = rowN;
    r = 1;
  } else if (colN === undefined && width > 1) {
    // INDEX(rect, n) with no column on a 2-D range returns the whole row.
    if (r < 0 || r > height) return { error: "#REF!" };
    return r === 0 ? ref : { ...ref, r1: ref.r1 + r - 1, r2: ref.r1 + r - 1 };
  }
  if (r < 0 || c < 0 || r > height || c > width) return { error: "#REF!" };
  return {
    kind: "ref",
    sheet: ref.sheet,
    r1: r === 0 ? ref.r1 : ref.r1 + r - 1,
    r2: r === 0 ? ref.r2 : ref.r1 + r - 1,
    c1: c === 0 ? ref.c1 : ref.c1 + c - 1,
    c2: c === 0 ? ref.c2 : ref.c1 + c - 1,
  };
}

// ---------------------------------------------------------------- scalar ops

function firstScalar(v: Scalar | ArrayValue): Scalar {
  return isArray(v) ? (v.rows[0]?.[0] ?? null) : v;
}

export function toNumber(v: Scalar): number | ErrorValue {
  if (v === null) return 0;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (isError(v)) return v;
  const trimmed = v.trim();
  if (trimmed !== "" && !Number.isNaN(Number(trimmed))) return Number(trimmed);
  return { error: "#VALUE!" };
}

function toText(v: Scalar): string {
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (isError(v)) return v.error;
  return String(v);
}

export function truthy(v: Scalar): boolean {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v === null) return false;
  if (isError(v)) return false;
  const upper = v.toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  throw new Unsupported("text as boolean");
}

function typeRank(v: Scalar): number {
  if (typeof v === "number" || v === null) return 0;
  if (typeof v === "string") return 1;
  if (typeof v === "boolean") return 2;
  return 3;
}

/** Excel ordering: numbers < text < booleans; text is case-insensitive. */
export function compareScalars(a: Scalar, b: Scalar): number {
  const aa = a === null ? (typeof b === "string" ? "" : typeof b === "boolean" ? false : 0) : a;
  const bb = b === null ? (typeof a === "string" ? "" : typeof a === "boolean" ? false : 0) : b;
  const ra = typeRank(aa);
  const rb = typeRank(bb);
  if (ra !== rb) return ra - rb;
  if (typeof aa === "number" && typeof bb === "number") return aa === bb ? 0 : aa < bb ? -1 : 1;
  if (typeof aa === "string" && typeof bb === "string") {
    const x = aa.toLowerCase();
    const y = bb.toLowerCase();
    return x === y ? 0 : x < y ? -1 : 1;
  }
  if (typeof aa === "boolean" && typeof bb === "boolean") return Number(aa) - Number(bb);
  return 0;
}

function applyOp(op: string, a: Scalar, b: Scalar): Scalar {
  if (isError(a)) return a;
  if (isError(b)) return b;
  switch (op) {
    case "&":
      return toText(a) + toText(b);
    case "=":
      return compareScalars(a, b) === 0;
    case "<>":
      return compareScalars(a, b) !== 0;
    case "<":
      return compareScalars(a, b) < 0;
    case ">":
      return compareScalars(a, b) > 0;
    case "<=":
      return compareScalars(a, b) <= 0;
    case ">=":
      return compareScalars(a, b) >= 0;
    default: {
      const x = toNumber(a);
      const y = toNumber(b);
      if (isError(x)) return x;
      if (isError(y)) return y;
      if (op === "+") return x + y;
      if (op === "-") return x - y;
      if (op === "*") return x * y;
      if (op === "/") return y === 0 ? { error: "#DIV/0!" } : x / y;
      if (op === "^") return x ** y;
      throw new Unsupported(`operator ${op}`);
    }
  }
}

function broadcast(
  l: Scalar | ArrayValue,
  r: Scalar | ArrayValue,
  f: (a: Scalar, b: Scalar) => Scalar
): Scalar | ArrayValue {
  if (!isArray(l) && !isArray(r)) return f(l, r);
  const lr = isArray(l) ? l.rows : [[l]];
  const rr = isArray(r) ? r.rows : [[r]];
  const height = Math.max(lr.length, rr.length);
  const width = Math.max(lr[0]?.length ?? 0, rr[0]?.length ?? 0);
  const pick = (rows: Scalar[][], i: number, j: number): Scalar => {
    const row = rows.length === 1 ? rows[0] : rows[i];
    if (!row) return { error: "#N/A" };
    const v = row.length === 1 ? row[0] : row[j];
    return v === undefined ? { error: "#N/A" } : v;
  };
  const rows: Scalar[][] = [];
  for (let i = 0; i < height; i += 1) {
    const row: Scalar[] = [];
    for (let j = 0; j < width; j += 1) row.push(f(pick(lr, i, j), pick(rr, i, j)));
    rows.push(row);
  }
  return { kind: "array", rows };
}

function wildcardRegex(pattern: string): RegExp {
  let re = "";
  for (let i = 0; i < pattern.length; i += 1) {
    const ch = pattern[i];
    if (ch === "~" && i + 1 < pattern.length) {
      re += pattern[i + 1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      i += 1;
    } else if (ch === "*") re += ".*";
    else if (ch === "?") re += ".";
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`, "i");
}

function equalsLookup(lookup: Scalar, v: Scalar, wildcards: boolean): boolean {
  if (wildcards && typeof lookup === "string" && typeof v === "string" && /[*?~]/.test(lookup)) {
    return wildcardRegex(lookup).test(v);
  }
  if (v === null) return false;
  return typeRank(lookup) === typeRank(v) && compareScalars(lookup, v) === 0;
}

/** MATCH semantics: 0 exact (wildcards), 1 largest <= lookup, -1 smallest >= lookup. */
export function matchPosition(lookup: Scalar, list: Scalar[], mode: number): number {
  if (mode === 0) return list.findIndex((v) => equalsLookup(lookup, v, true));
  let best = -1;
  for (let i = 0; i < list.length; i += 1) {
    const v = list[i];
    if (v === null || typeRank(v) !== typeRank(lookup)) continue;
    const cmp = compareScalars(v, lookup);
    if (mode > 0) {
      if (cmp <= 0) best = i;
      else break;
    } else if (cmp >= 0) best = i;
    else break;
  }
  return best;
}

/** XMATCH/XLOOKUP: mode 0 exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcard. */
export function xmatchPosition(
  lookup: Scalar,
  list: Scalar[],
  mode: number,
  search: number
): number {
  const order = list.map((_, i) => i);
  if (search < 0) order.reverse();
  if (Math.abs(search) === 2) throw new Unsupported("binary search mode");
  for (const i of order) if (equalsLookup(lookup, list[i], mode === 2)) return i;
  if (mode === 0 || mode === 2) return -1;
  let best = -1;
  for (const i of order) {
    const v = list[i];
    if (v === null || typeRank(v) !== typeRank(lookup)) continue;
    const cmp = compareScalars(v, lookup);
    if (mode < 0 && cmp < 0 && (best < 0 || compareScalars(v, list[best]) > 0)) best = i;
    if (mode > 0 && cmp > 0 && (best < 0 || compareScalars(v, list[best]) < 0)) best = i;
  }
  return best;
}

/** SUMIF-style criteria: "MU", 5, ">=5", "<>0", "a*", "", "=", "<>". */
export function matchesCriteria(cell: Scalar, criteria: Scalar): boolean {
  if (isError(criteria)) return false;
  if (typeof criteria === "number" || typeof criteria === "boolean") {
    if (typeof cell === "string") {
      const n = Number(cell.trim());
      return typeof criteria === "number" && cell.trim() !== "" && n === criteria;
    }
    return (
      cell !== null && typeRank(cell) === typeRank(criteria) && compareScalars(cell, criteria) === 0
    );
  }
  const text = criteria ?? "";
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(text);
  const op = m?.[1] ?? "=";
  const operand = m?.[2] ?? "";
  const hasOp = m?.[1] !== undefined;

  if (operand === "") {
    if (op === "=") return hasOp ? cell === null : cell === null || cell === "";
    if (op === "<>") return cell !== null && cell !== "";
    return false;
  }
  const opNum = Number(operand);
  const numeric = operand.trim() !== "" && !Number.isNaN(opNum);
  const upper = operand.toUpperCase();
  const opValue: Scalar = numeric
    ? opNum
    : upper === "TRUE"
      ? true
      : upper === "FALSE"
        ? false
        : operand;

  if (op === "=" || op === "<>") {
    let eq: boolean;
    if (typeof opValue === "string") {
      eq = typeof cell === "string" && wildcardRegex(opValue).test(cell);
    } else if (typeof opValue === "number") {
      eq =
        (typeof cell === "number" && cell === opValue) ||
        (typeof cell === "string" && cell.trim() !== "" && Number(cell) === opValue);
    } else {
      eq = cell === opValue;
    }
    return op === "=" ? eq : !eq;
  }
  if (cell === null || isError(cell)) return false;
  if (typeRank(cell) !== typeRank(opValue)) return false;
  const cmp = compareScalars(cell, opValue);
  if (op === "<") return cmp < 0;
  if (op === ">") return cmp > 0;
  if (op === "<=") return cmp <= 0;
  return cmp >= 0;
}

// ---------------------------------------------------------------- dates

const EPOCH = Date.UTC(1899, 11, 30);

export function dateSerial(y: number, m: number, d: number): number {
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
}

function serialToDate(serial: number): { y: number; m: number; d: number } {
  const dt = new Date(EPOCH + Math.floor(serial) * 86400000);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}
