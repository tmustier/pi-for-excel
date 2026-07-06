import assert from "node:assert/strict";
import { describe, test } from "node:test";

import {
  diffFromDefaults,
  getStoredConventions,
  mergeStoredConventions,
  removeCustomPresets,
  resolveConventions,
  setStoredConventions,
} from "../src/conventions/store.ts";
import {
  DEFAULT_COLOR_CONVENTIONS,
  DEFAULT_HEADER_STYLE,
  DEFAULT_PRESET_FORMATS,
  DEFAULT_VISUAL_DEFAULTS,
} from "../src/conventions/defaults.ts";
import type { StoredConventions } from "../src/conventions/types.ts";

function createFakeStore(): {
  get: (key: string) => Promise<DynamicValue>;
  set: (key: string, value: DynamicValue) => Promise<void>;
  data: Map<string, DynamicValue>;
} {
  const data = new Map<string, DynamicValue>();

  return {
    data,
    get: (key: string) => Promise.resolve(data.get(key)),
    set: (key: string, value: DynamicValue) => {
      data.set(key, value);
      return Promise.resolve();
    },
  };
}

void describe("getStoredConventions", () => {
  void test("returns empty object when nothing stored", async () => {
    const store = createFakeStore();
    const result = await getStoredConventions(store);
    assert.deepEqual(result, {});
  });

  void test("validates nested sections and normalizes colors", async () => {
    const store = createFakeStore();
    store.data.set("conventions.v1", {
      presetFormats: {
        number: { format: "#,##0.000" },
      },
      customPresets: {
        bps: {
          format: '#,##0 "bps"',
          description: "Basis points",
        },
      },
      visualDefaults: {
        fontName: "Calibri",
        fontSize: 11,
      },
      colorConventions: {
        hardcodedValueColor: "rgb(0,0,255)",
        crossSheetLinkColor: "#008000",
      },
      headerStyle: {
        fillColor: "#002060",
        fontColor: "#fff",
        bold: true,
        wrapText: false,
      },
    });

    const result = await getStoredConventions(store);
    assert.equal(result.presetFormats?.number?.format, "#,##0.000");
    assert.equal(result.customPresets?.bps?.description, "Basis points");
    assert.equal(result.visualDefaults?.fontName, "Calibri");
    assert.equal(result.colorConventions?.hardcodedValueColor, "#0000FF");
    assert.equal(result.colorConventions?.crossSheetLinkColor, "#008000");
    assert.equal(result.headerStyle?.fontColor, "#FFFFFF");
  });

  void test("drops invalid nested values", async () => {
    const store = createFakeStore();
    store.data.set("conventions.v1", {
      presetFormats: {
        number: { format: "" },
      },
      customPresets: {
        "": { format: "0.00" },
      },
      visualDefaults: {
        fontSize: 1000,
      },
      colorConventions: {
        hardcodedValueColor: "blue",
      },
    });

    const result = await getStoredConventions(store);
    assert.equal(result.presetFormats, undefined);
    assert.equal(result.customPresets, undefined);
    assert.equal(result.visualDefaults, undefined);
    assert.equal(result.colorConventions, undefined);
  });
});

void describe("setStoredConventions", () => {
  void test("persists validated shape", async () => {
    const store = createFakeStore();
    await setStoredConventions(store, {
      visualDefaults: { fontName: "Georgia" },
      colorConventions: { hardcodedValueColor: "rgb(255,0,0)" },
    });

    const raw = store.data.get("conventions.v1");
    assert.equal(typeof raw, "object");
    assert.ok(raw !== null);

    if (!raw || typeof raw !== "object") {
      assert.fail("Expected conventions payload to be an object");
    }

    const visualDefaults = "visualDefaults" in raw ? raw.visualDefaults : undefined;
    const colorConventions = "colorConventions" in raw ? raw.colorConventions : undefined;

    const fontName = (
      visualDefaults
      && typeof visualDefaults === "object"
      && "fontName" in visualDefaults
      && typeof visualDefaults.fontName === "string"
    )
      ? visualDefaults.fontName
      : undefined;

    const hardcodedValueColor = (
      colorConventions
      && typeof colorConventions === "object"
      && "hardcodedValueColor" in colorConventions
      && typeof colorConventions.hardcodedValueColor === "string"
    )
      ? colorConventions.hardcodedValueColor
      : undefined;

    assert.equal(fontName, "Georgia");
    assert.equal(hardcodedValueColor, "#FF0000");
  });
});

void describe("resolveConventions", () => {
  void test("returns defaults for empty input", () => {
    const resolved = resolveConventions({});
    assert.deepEqual(resolved.presetFormats, DEFAULT_PRESET_FORMATS);
    assert.deepEqual(resolved.visualDefaults, DEFAULT_VISUAL_DEFAULTS);
    assert.deepEqual(resolved.colorConventions, DEFAULT_COLOR_CONVENTIONS);
    assert.deepEqual(resolved.headerStyle, DEFAULT_HEADER_STYLE);
    assert.deepEqual(resolved.customPresets, {});
  });

  void test("returns deep-cloned preset defaults", () => {
    const first = resolveConventions({});
    const second = resolveConventions({});

    const firstCurrencyBuilder = first.presetFormats.currency.builderParams;
    const secondCurrencyBuilder = second.presetFormats.currency.builderParams;

    assert.ok(firstCurrencyBuilder);
    assert.ok(secondCurrencyBuilder);

    if (!firstCurrencyBuilder || !secondCurrencyBuilder) {
      assert.fail("Expected currency builder params to be defined");
    }

    firstCurrencyBuilder.dp = 4;
    first.presetFormats.currency.format = "custom";

    assert.equal(secondCurrencyBuilder.dp, DEFAULT_PRESET_FORMATS.currency.builderParams?.dp);
    assert.equal(second.presetFormats.currency.format, DEFAULT_PRESET_FORMATS.currency.format);
  });

  void test("merges overrides over defaults", () => {
    const resolved = resolveConventions({
      presetFormats: {
        currency: { format: "£#,##0.00" },
      },
      visualDefaults: {
        fontName: "Times New Roman",
      },
    });

    assert.equal(resolved.presetFormats.currency.format, "£#,##0.00");
    assert.equal(resolved.presetFormats.number.format, DEFAULT_PRESET_FORMATS.number.format);
    assert.equal(resolved.visualDefaults.fontName, "Times New Roman");
    assert.equal(resolved.visualDefaults.fontSize, DEFAULT_VISUAL_DEFAULTS.fontSize);
  });
});

void describe("mergeStoredConventions", () => {
  void test("merges nested sections additively", () => {
    const current: StoredConventions = {
      visualDefaults: { fontName: "Calibri" },
      customPresets: { bps: { format: '#,##0 "bps"' } },
    };

    const updates: StoredConventions = {
      visualDefaults: { fontSize: 12 },
      customPresets: { date: { format: "dd-mmm-yyyy" } },
    };

    const merged = mergeStoredConventions(current, updates);
    assert.equal(merged.visualDefaults?.fontName, "Calibri");
    assert.equal(merged.visualDefaults?.fontSize, 12);
    assert.ok(merged.customPresets?.bps);
    assert.ok(merged.customPresets?.date);
  });

  void test("removeCustomPresets removes requested preset names", () => {
    const current: StoredConventions = {
      customPresets: {
        bps: { format: '#,##0 "bps"' },
        date: { format: "dd-mmm-yyyy" },
      },
    };

    const next = removeCustomPresets(current, ["bps"]);
    assert.equal(next.customPresets?.bps, undefined);
    assert.ok(next.customPresets?.date);
  });
});

void describe("diffFromDefaults", () => {
  void test("returns empty for default config", () => {
    const diffs = diffFromDefaults(resolveConventions({}));
    assert.equal(diffs.length, 0);
  });

  void test("includes custom presets and visual overrides", () => {
    const diffs = diffFromDefaults(resolveConventions({
      customPresets: {
        bps: {
          format: '#,##0 "bps"',
          description: "Basis points",
        },
      },
      visualDefaults: {
        fontName: "Calibri",
      },
    }));

    const fields = diffs.map((diff) => diff.field);
    assert.ok(fields.includes("customPresets.bps"));
    assert.ok(fields.includes("visualDefaults.fontName"));
  });
});
