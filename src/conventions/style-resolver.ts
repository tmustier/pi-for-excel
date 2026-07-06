/**
 * style-resolver — compose named styles + overrides into a flat ResolvedCellStyle.
 */

import type { CellStyle, NumberFormatConventions, ResolvedCellStyle, ResolvedConventions } from "./types.js";
import { DEFAULT_CONVENTION_CONFIG, DEFAULT_CONVENTIONS, getBuiltinStyles } from "./defaults.js";
import { buildFormatString } from "./format-builder.js";
import { getPresetFormat, isBuiltinPresetName, isPresetName } from "./store.js";

function resolveBuilderConventions(params: {
  negativeStyle?: NumberFormatConventions["negativeStyle"];
  zeroStyle?: NumberFormatConventions["zeroStyle"];
  thousandsSeparator?: boolean;
}): NumberFormatConventions {
  return {
    negativeStyle: params.negativeStyle ?? DEFAULT_CONVENTIONS.negativeStyle,
    zeroStyle: params.zeroStyle ?? DEFAULT_CONVENTIONS.zeroStyle,
    thousandsSeparator: params.thousandsSeparator ?? DEFAULT_CONVENTIONS.thousandsSeparator,
    accountingPadding: DEFAULT_CONVENTIONS.accountingPadding,
  };
}

export function resolveStyles(
  styles: string | string[] | undefined,
  overrides?: Partial<CellStyle>,
  config: ResolvedConventions = DEFAULT_CONVENTION_CONFIG,
): ResolvedCellStyle {
  const warnings: string[] = [];
  const merged: CellStyle = {};

  if (styles) {
    const names = Array.isArray(styles) ? styles : [styles];
    const builtinStyles = getBuiltinStyles(config);

    for (const styleName of names) {
      const style = builtinStyles.get(styleName);
      if (style) {
        mergeCellStyle(merged, style.properties);
        continue;
      }

      if (Object.prototype.hasOwnProperty.call(config.customPresets, styleName)) {
        mergeCellStyle(merged, { numberFormat: styleName });
        continue;
      }

      warnings.push(`Unknown style "${styleName}" — ignored.`);
    }
  }

  if (overrides) {
    mergeCellStyle(merged, overrides);
  }

  let excelNumberFormat: string | undefined;

  if (merged.numberFormat) {
    const presetName = merged.numberFormat;

    if (isPresetName(presetName, config)) {
      const preset = getPresetFormat(config, presetName);
      if (!preset) {
        warnings.push(`Unknown number format preset "${presetName}".`);
      } else {
        const hasDpOverride = merged.numberFormatDp !== undefined;
        const hasCurrencyOverride = merged.currencySymbol !== undefined;

        if (!hasDpOverride && !hasCurrencyOverride) {
          excelNumberFormat = preset.format;
        } else if (isBuiltinPresetName(presetName)) {
          const builderParams = preset.builderParams;
          const built = buildFormatString(
            presetName,
            merged.numberFormatDp ?? builderParams?.dp,
            merged.currencySymbol ?? builderParams?.currencySymbol,
            resolveBuilderConventions({
              ...(builderParams?.negativeStyle !== undefined ? { negativeStyle: builderParams.negativeStyle } : {}),
              ...(builderParams?.zeroStyle !== undefined ? { zeroStyle: builderParams.zeroStyle } : {}),
              ...(builderParams?.thousandsSeparator !== undefined
                ? { thousandsSeparator: builderParams.thousandsSeparator }
                : {}),
            }),
          );

          excelNumberFormat = built.format;
          warnings.push(...built.warnings);
        } else {
          excelNumberFormat = preset.format;
          if (hasDpOverride) {
            warnings.push("number_format_dp ignored for custom presets unless quick-toggle metadata exists.");
          }
          if (hasCurrencyOverride) {
            warnings.push("currency_symbol ignored for non-currency custom presets.");
          }
        }
      }
    } else {
      excelNumberFormat = presetName;

      if (merged.numberFormatDp !== undefined) {
        warnings.push("number_format_dp ignored — only applies to preset names, not raw format strings.");
      }
      if (merged.currencySymbol) {
        warnings.push("currency_symbol ignored — only applies to currency presets, not raw format strings.");
      }
    }
  } else {
    if (merged.numberFormatDp !== undefined) {
      warnings.push("number_format_dp ignored — no number format or style specified.");
    }
    if (merged.currencySymbol) {
      warnings.push("currency_symbol ignored — no currency format or style specified.");
    }
  }

  const resolved: ResolvedCellStyle = { properties: merged, warnings };
  if (excelNumberFormat !== undefined) {
    resolved.excelNumberFormat = excelNumberFormat;
  }

  return resolved;
}

function mergeCellStyle(target: CellStyle, source: Partial<CellStyle>): void {
  if (source.numberFormat !== undefined) target.numberFormat = source.numberFormat;
  if (source.numberFormatDp !== undefined) target.numberFormatDp = source.numberFormatDp;
  if (source.currencySymbol !== undefined) target.currencySymbol = source.currencySymbol;

  if (source.bold !== undefined) target.bold = source.bold;
  if (source.italic !== undefined) target.italic = source.italic;
  if (source.underline !== undefined) target.underline = source.underline;
  if (source.fontColor !== undefined) target.fontColor = source.fontColor;
  if (source.fontSize !== undefined) target.fontSize = source.fontSize;
  if (source.fontName !== undefined) target.fontName = source.fontName;

  if (source.fillColor !== undefined) target.fillColor = source.fillColor;

  if (source.borderTop !== undefined) target.borderTop = source.borderTop;
  if (source.borderBottom !== undefined) target.borderBottom = source.borderBottom;
  if (source.borderLeft !== undefined) target.borderLeft = source.borderLeft;
  if (source.borderRight !== undefined) target.borderRight = source.borderRight;

  if (source.horizontalAlignment !== undefined) target.horizontalAlignment = source.horizontalAlignment;
  if (source.verticalAlignment !== undefined) target.verticalAlignment = source.verticalAlignment;
  if (source.wrapText !== undefined) target.wrapText = source.wrapText;
}
