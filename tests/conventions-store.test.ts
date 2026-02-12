import assert from "node:assert/strict";
import { test, describe } from "node:test";

import {
  getStoredConventions,
  setStoredConventions,
  resolveConventions,
  mergeStoredConventions,
  diffFromDefaults,
} from "../src/conventions/store.ts";
import { DEFAULT_CONVENTIONS, DEFAULT_CURRENCY_SYMBOL, PRESET_DEFAULT_DP } from "../src/conventions/defaults.ts";
import type { StoredConventions } from "../src/conventions/types.ts";

// ── Fake SettingsStore ───────────────────────────────────────────────

function createFakeStore(): { get: (k: string) => Promise<unknown>; set: (k: string, v: unknown) => Promise<void>; data: Map<string, unknown> } {
  const data = new Map<string, unknown>();
  return {
    data,
    get: (key: string) => Promise.resolve(data.get(key)),
    set: (key: string, value: unknown) => { data.set(key, value); return Promise.resolve(); },
  };
}

// ── getStoredConventions ─────────────────────────────────────────────

void describe("getStoredConventions", () => {
  void test("returns empty object when nothing stored", async () => {
    const store = createFakeStore();
    const result = await getStoredConventions(store);
    assert.deepEqual(result, {});
  });

  void test("returns validated stored values", async () => {
    const store = createFakeStore();
    store.data.set("conventions.v1", {
      currencySymbol: "£",
      negativeStyle: "minus",
      zeroStyle: "blank",
    });
    const result = await getStoredConventions(store);
    assert.equal(result.currencySymbol, "£");
    assert.equal(result.negativeStyle, "minus");
    assert.equal(result.zeroStyle, "blank");
  });

  void test("ignores invalid values", async () => {
    const store = createFakeStore();
    store.data.set("conventions.v1", {
      currencySymbol: 42, // not a string
      negativeStyle: "invalid", // not parens/minus
      zeroStyle: "dash", // valid
      thousandsSeparator: "yes", // not a boolean
    });
    const result = await getStoredConventions(store);
    assert.equal(result.currencySymbol, undefined);
    assert.equal(result.negativeStyle, undefined);
    assert.equal(result.zeroStyle, "dash");
    assert.equal(result.thousandsSeparator, undefined);
  });

  void test("validates presetDp values", async () => {
    const store = createFakeStore();
    store.data.set("conventions.v1", {
      presetDp: {
        number: 3, // valid
        currency: -1, // invalid (negative)
        percent: 11, // invalid (> 10)
        ratio: 1.5, // invalid (not integer)
        text: 2, // ignored (text not in overridable list)
      },
    });
    const result = await getStoredConventions(store);
    assert.deepEqual(result.presetDp, { number: 3 });
  });
});

// ── setStoredConventions ─────────────────────────────────────────────

void describe("setStoredConventions", () => {
  void test("persists to store", async () => {
    const store = createFakeStore();
    await setStoredConventions(store, { currencySymbol: "€" });
    const raw = store.data.get("conventions.v1") as StoredConventions;
    assert.equal(raw.currencySymbol, "€");
  });
});

// ── resolveConventions ───────────────────────────────────────────────

void describe("resolveConventions", () => {
  void test("returns all defaults for empty stored", () => {
    const resolved = resolveConventions({});
    assert.deepEqual(resolved.conventions, DEFAULT_CONVENTIONS);
    assert.equal(resolved.currencySymbol, DEFAULT_CURRENCY_SYMBOL);
    assert.deepEqual(resolved.presetDp, PRESET_DEFAULT_DP);
  });

  void test("merges stored values over defaults", () => {
    const resolved = resolveConventions({
      currencySymbol: "£",
      negativeStyle: "minus",
      presetDp: { currency: 0, number: 1 },
    });

    assert.equal(resolved.currencySymbol, "£");
    assert.equal(resolved.conventions.negativeStyle, "minus");
    // Overridden
    assert.equal(resolved.presetDp.currency, 0);
    assert.equal(resolved.presetDp.number, 1);
    // Kept default
    assert.equal(resolved.conventions.thousandsSeparator, DEFAULT_CONVENTIONS.thousandsSeparator);
    assert.equal(resolved.presetDp.percent, PRESET_DEFAULT_DP.percent);
  });
});

// ── mergeStoredConventions ───────────────────────────────────────────

void describe("mergeStoredConventions", () => {
  void test("merges partial updates into existing", () => {
    const current: StoredConventions = { currencySymbol: "£", negativeStyle: "parens" };
    const updates: StoredConventions = { negativeStyle: "minus", zeroStyle: "zero" };
    const result = mergeStoredConventions(current, updates);

    assert.equal(result.currencySymbol, "£"); // kept
    assert.equal(result.negativeStyle, "minus"); // updated
    assert.equal(result.zeroStyle, "zero"); // added
  });

  void test("merges presetDp additively", () => {
    const current: StoredConventions = { presetDp: { number: 1 } };
    const updates: StoredConventions = { presetDp: { currency: 0 } };
    const result = mergeStoredConventions(current, updates);

    assert.deepEqual(result.presetDp, { number: 1, currency: 0 });
  });
});

// ── diffFromDefaults ─────────────────────────────────────────────────

void describe("diffFromDefaults", () => {
  void test("returns empty for all defaults", () => {
    const resolved = resolveConventions({});
    const diffs = diffFromDefaults(resolved);
    assert.equal(diffs.length, 0);
  });

  void test("detects currency symbol change", () => {
    const resolved = resolveConventions({ currencySymbol: "€" });
    const diffs = diffFromDefaults(resolved);
    assert.equal(diffs.length, 1);
    assert.equal(diffs[0].field, "currencySymbol");
    assert.equal(diffs[0].value, "€");
  });

  void test("detects multiple changes", () => {
    const resolved = resolveConventions({
      currencySymbol: "£",
      negativeStyle: "minus",
      presetDp: { currency: 0 },
    });
    const diffs = diffFromDefaults(resolved);
    const fields = diffs.map((d) => d.field);
    assert.ok(fields.includes("currencySymbol"));
    assert.ok(fields.includes("negativeStyle"));
    assert.ok(fields.includes("presetDp.currency"));
    assert.equal(diffs.length, 3);
  });
});
