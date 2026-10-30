export type CellInput = {
  isFormula: boolean;
  isArrayFormula?: boolean;
  formula?: string;
  value?: unknown;
};

/**
 * Returns X when `body` is exactly `-(X)` (the shape Switch Sign writes), so a
 * second press toggles back instead of nesting `-(-(...))`. Returns null when
 * the leading `-(` does not close at the very end, e.g. `-(A1)+(B1)`.
 */
function unwrapNegation(body: string): string | null {
  const trimmed = body.trim();
  if (!trimmed.startsWith("-(") || !trimmed.endsWith(")")) {
    return null;
  }

  let depth = 0;
  let inString = false;
  for (let i = 1; i < trimmed.length; i++) {
    const ch = trimmed[i];
    if (ch === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;
    if (ch === "(") depth++;
    else if (ch === ")") {
      depth--;
      if (depth === 0 && i !== trimmed.length - 1) {
        return null;
      }
    }
  }

  return depth === 0 ? trimmed.slice(2, -1) : null;
}

export function switchSignCell(input: CellInput): CellInput {
  if (input.isFormula) {
    const isArrayFormula = Boolean(input.isArrayFormula);
    let formulaText = input.formula ?? "";

    if (isArrayFormula && formulaText.startsWith("{") && formulaText.endsWith("}")) {
      formulaText = formulaText.slice(1, -1);
    }

    const hasEquals = formulaText.startsWith("=");
    const body = hasEquals ? formulaText.slice(1) : formulaText;
    const unwrapped = unwrapNegation(body);
    const switched = unwrapped ?? `-(${body})`;
    formulaText = hasEquals ? `=${switched}` : switched;

    if (isArrayFormula) {
      return {
        isFormula: true,
        isArrayFormula: true,
        formula: `{${formulaText}}`
      };
    }

    return {
      isFormula: true,
      formula: formulaText
    };
  }

  const value = input.value;
  const isNumericValue = typeof value === "number" && Number.isFinite(value);

  if (isNumericValue) {
    const negated = -(value as number);
    return {
      isFormula: false,
      value: Object.is(negated, -0) ? 0 : negated
    };
  }

  return {
    isFormula: false,
    value
  };
}
