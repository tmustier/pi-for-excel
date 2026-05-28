/**
 * Lightweight translation function — zero dependencies.
 * 
 * Usage:
 *   import { t } from "../language/index.js";
 *   element.textContent = t("welcome.subtitle");
 *   element.textContent = t("welcome.connected", { label: "DeepSeek" });
 */

import en from "./locales/en.json";
import zhCN from "./locales/zh-CN.json";

const translations: Record<string, Record<string, string>> = {
  en,
  "zh-CN": zhCN,
};

let currentLang = "en";

export function initLanguage(lang: string): void {
  if (translations[lang]) {
    currentLang = lang;
  }
}

export function t(key: string, vars?: Record<string, string>): string {
  const dict = translations[currentLang] ?? translations.en;
  let value = dict[key] ?? translations.en[key] ?? key;
  if (vars) {
    for (const [k, v] of Object.entries(vars)) {
      value = value.replace(`{${k}}`, v);
    }
  }
  return value;
}

export function getLanguage(): string {
  return currentLang;
}
