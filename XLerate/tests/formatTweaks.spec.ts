import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  adjustDecimals,
  MULTIPLE_FORMATS,
  nextBorderWeight,
  nextFormatInCycle,
  PERCENT_FORMATS,
} from "../src/core/formatTweaks";
import { SHORTCUT_CATALOG } from "../src/core/shortcutCatalog";

describe("nextFormatInCycle", () => {
  it("starts the percent cycle from an unknown format", () => {
    expect(nextFormatInCycle("General", PERCENT_FORMATS)).toBe("0%");
  });

  it("advances and wraps the percent cycle", () => {
    expect(nextFormatInCycle("0%", PERCENT_FORMATS)).toBe("0.0%");
    expect(nextFormatInCycle("0.00%", PERCENT_FORMATS)).toBe("0%");
  });

  it("recognises multiples however Excel quotes the x", () => {
    expect(nextFormatInCycle('0.0"x"', MULTIPLE_FORMATS)).toBe('0.00"x"');
    expect(nextFormatInCycle("0.0\\x", MULTIPLE_FORMATS)).toBe('0.00"x"');
    expect(nextFormatInCycle("0x", MULTIPLE_FORMATS)).toBe('0.0"x"');
  });
});

describe("adjustDecimals", () => {
  it("treats General as an integer format", () => {
    expect(adjustDecimals("General", 1)).toBe("0.0");
    expect(adjustDecimals("General", -1)).toBe("0");
  });

  it("adds and removes decimals on plain numbers", () => {
    expect(adjustDecimals("#,##0", 1)).toBe("#,##0.0");
    expect(adjustDecimals("#,##0.0", 1)).toBe("#,##0.00");
    expect(adjustDecimals("#,##0.00", -1)).toBe("#,##0.0");
    expect(adjustDecimals("#,##0.0", -1)).toBe("#,##0");
    expect(adjustDecimals("#,##0", -1)).toBe("#,##0");
  });

  it("keeps percent signs, multiples and scaling commas in place", () => {
    expect(adjustDecimals("0%", 1)).toBe("0.0%");
    expect(adjustDecimals('0.0"x"', -1)).toBe('0"x"');
    expect(adjustDecimals('#,##0,,"mm"', 1)).toBe('#,##0.0,,"mm"');
  });

  it("changes every numeric section and ignores literal text", () => {
    expect(adjustDecimals('#,##0.0_);(#,##0.0);"-"_)', 1)).toBe(
      '#,##0.00_);(#,##0.00);"-"_)'
    );
    expect(adjustDecimals('[Red]0.0;"0.0 text"', -1)).toBe('[Red]0;"0.0 text"');
  });

  it("leaves text and date formats alone", () => {
    expect(adjustDecimals("@", 1)).toBe("@");
    expect(adjustDecimals("dd.mm.yyyy", 1)).toBe("dd.mm.yyyy");
  });
});

describe("nextBorderWeight", () => {
  it("cycles none -> thin -> medium -> thick -> none", () => {
    expect(nextBorderWeight("None", "Thin")).toBe("Thin");
    expect(nextBorderWeight(null, null)).toBe("Thin");
    expect(nextBorderWeight("Continuous", "Thin")).toBe("Medium");
    expect(nextBorderWeight("Continuous", "Medium")).toBe("Thick");
    expect(nextBorderWeight("Continuous", "Thick")).toBe("None");
  });
});

describe("shortcut catalog", () => {
  it("matches shortcuts.json exactly", () => {
    const json = JSON.parse(readFileSync("shortcuts.json", "utf8")) as {
      actions: { id: string; name: string }[];
      shortcuts: { action: string; key: { default: string } }[];
    };
    const fromJson = json.actions.map((action) => ({
      id: action.id,
      name: action.name,
      key: json.shortcuts.find((s) => s.action === action.id)?.key.default,
    }));
    expect(fromJson).toEqual(SHORTCUT_CATALOG.map(({ id, name, key }) => ({ id, name, key })));
  });

  it("only uses keys Office accepts and never repeats one", () => {
    const keys = SHORTCUT_CATALOG.map((item) => item.key);
    for (const key of keys) expect(key).toMatch(/^[A-Za-z0-9-_+]+$/);
    expect(new Set(keys.map((k) => k.toUpperCase())).size).toBe(keys.length);
  });
});
