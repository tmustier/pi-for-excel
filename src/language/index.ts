import en from "./locales/en.json" with { type: "json" };
import zhCN from "./locales/zh-CN.json" with { type: "json" };

export const SUPPORTED_LANGUAGES = ["en", "zh-CN"] as const;
export type SupportedLanguage = (typeof SUPPORTED_LANGUAGES)[number];

const translations: Record<SupportedLanguage, Record<string, string>> = {
  en,
  "zh-CN": zhCN,
};

let currentLang: SupportedLanguage = "en";

export function isSupportedLanguage(lang: string): lang is SupportedLanguage {
  return (SUPPORTED_LANGUAGES as readonly string[]).includes(lang);
}

export function initLanguage(lang: string): void {
  if (isSupportedLanguage(lang)) {
    currentLang = lang;
    if (typeof document !== "undefined") document.documentElement.lang = lang;
  }
}

export function t(key: string, vars?: Record<string, string | number>): string {
  const dict = translations[currentLang] ?? translations.en;
  let value = dict[key] ?? translations.en[key] ?? key;
  if (vars) {
    for (const [k, v] of Object.entries(vars)) {
      value = value.split(`{${k}}`).join(String(v));
    }
  }
  return value;
}

export function getLanguage(): SupportedLanguage {
  return currentLang;
}
