/**
 * Pure number-format helpers behind the percent / multiple cycles and the
 * increase / decrease decimals shortcuts. No Office.js here.
 */

export const PERCENT_FORMATS: readonly string[] = ["0%", "0.0%", "0.00%"];
export const MULTIPLE_FORMATS: readonly string[] = ['0.0"x"', '0.00"x"', '0"x"'];

// Excel may hand back `0.0"x"` as `0.0\x` or `0.0x`; compare without the
// literal-text quoting so the cycle still recognises its own output.
function normalizeFormat(format: string): string {
  return format.replace(/["\\]/g, "").trim().toLowerCase();
}

export function nextFormatInCycle(current: string, cycle: readonly string[]): string {
  const normalized = normalizeFormat(current);
  const index = cycle.findIndex((item) => normalizeFormat(item) === normalized);
  return cycle[(index + 1) % cycle.length];
}

/**
 * Marks which characters of a format code are format code (true) versus
 * literal text: quoted strings, `\x` escapes, `_x` / `*x` padding and
 * `[...]` colour / condition blocks.
 */
function codeMask(format: string): boolean[] {
  const mask = new Array<boolean>(format.length).fill(false);
  let i = 0;
  while (i < format.length) {
    const ch = format[i];
    if (ch === '"') {
      const end = format.indexOf('"', i + 1);
      i = end === -1 ? format.length : end + 1;
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      i += 2;
    } else if (ch === "[") {
      const end = format.indexOf("]", i + 1);
      i = end === -1 ? format.length : end + 1;
    } else {
      mask[i] = true;
      i += 1;
    }
  }
  return mask;
}

function adjustSection(section: string, mask: boolean[], delta: 1 | -1): string {
  const isPlaceholder = (i: number): boolean => mask[i] && "0#?".includes(section[i]);
  const placeholders: number[] = [];
  let dot = -1;
  for (let i = 0; i < section.length; i += 1) {
    if (isPlaceholder(i)) placeholders.push(i);
    else if (dot === -1 && mask[i] && section[i] === ".") dot = i;
  }
  // Text sections ("@") and dates have no digit placeholders: leave them.
  if (placeholders.length === 0) return section;

  if (dot >= 0) {
    let end = dot + 1;
    while (end < section.length && isPlaceholder(end)) end += 1;
    const decimals = end - dot - 1;
    if (delta > 0) return section.slice(0, end) + "0" + section.slice(end);
    if (decimals >= 2) return section.slice(0, end - 1) + section.slice(end);
    return section.slice(0, dot) + section.slice(end);
  }

  if (delta < 0) return section;
  const last = placeholders[placeholders.length - 1];
  return section.slice(0, last + 1) + ".0" + section.slice(last + 1);
}

export function adjustDecimals(format: string, delta: 1 | -1): string {
  if (format.trim() === "" || format.trim().toLowerCase() === "general") {
    return delta > 0 ? "0.0" : "0";
  }
  const mask = codeMask(format);
  const sections: string[] = [];
  let start = 0;
  for (let i = 0; i <= format.length; i += 1) {
    if (i === format.length || (mask[i] && format[i] === ";")) {
      sections.push(adjustSection(format.slice(start, i), mask.slice(start, i), delta));
      start = i + 1;
    }
  }
  return sections.join(";");
}

export type BorderWeight = "None" | "Thin" | "Medium" | "Thick";

const BORDER_CYCLE: readonly BorderWeight[] = ["None", "Thin", "Medium", "Thick"];

export function nextBorderWeight(style: string | null, weight: string | null): BorderWeight {
  const current: BorderWeight =
    !style || style === "None" ? "None" : (BORDER_CYCLE.find((w) => w === weight) ?? "None");
  return BORDER_CYCLE[(BORDER_CYCLE.indexOf(current) + 1) % BORDER_CYCLE.length];
}
