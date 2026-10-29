export type BorderWeight = "thin" | "medium" | "thick" | "none";

export type NumberPreset = "number" | "integer" | "currency" | "percent" | "ratio" | "text";

export type NegativeStyle = "parens" | "minus";
export type ZeroStyle = "dash" | "single-dash" | "zero" | "blank";

export interface NumberFormatConventions {
  negativeStyle: NegativeStyle;
  zeroStyle: ZeroStyle;
  thousandsSeparator: boolean;
  accountingPadding: boolean;
}

export interface FormatBuilderParams {
  dp?: number;
  negativeStyle?: NegativeStyle;
  zeroStyle?: ZeroStyle;
  thousandsSeparator?: boolean;
  currencySymbol?: string;
}

export interface StoredFormatPreset {
  format: string;
  builderParams?: FormatBuilderParams;
}

export interface StoredCustomPreset extends StoredFormatPreset {
  description?: string;
}

export interface StoredVisualDefaults {
  fontName?: string;
  fontSize?: number;
}

export interface StoredColorConventions {
  hardcodedValueColor?: string;
  crossSheetLinkColor?: string;
}

export interface StoredHeaderStyle {
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
  wrapText?: boolean;
}

export interface StoredConventions {
  presetFormats?: Partial<Record<NumberPreset, StoredFormatPreset>>;
  customPresets?: Record<string, StoredCustomPreset>;
  visualDefaults?: StoredVisualDefaults;
  colorConventions?: StoredColorConventions;
  headerStyle?: StoredHeaderStyle;
}

export interface ResolvedVisualDefaults {
  fontName: string;
  fontSize: number;
}

export interface ResolvedColorConventions {
  hardcodedValueColor: string;
  crossSheetLinkColor: string;
}

export interface ResolvedHeaderStyle {
  fillColor: string;
  fontColor: string;
  bold: boolean;
  wrapText: boolean;
}

export interface ResolvedConventions {
  presetFormats: Record<NumberPreset, StoredFormatPreset>;
  customPresets: Record<string, StoredCustomPreset>;
  visualDefaults: ResolvedVisualDefaults;
  colorConventions: ResolvedColorConventions;
  headerStyle: ResolvedHeaderStyle;
}

export interface CellStyle {
  numberFormat?: string;
  numberFormatDp?: number;
  currencySymbol?: string;

  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  fontSize?: number;
  fontName?: string;

  fillColor?: string;

  borderTop?: BorderWeight;
  borderBottom?: BorderWeight;
  borderLeft?: BorderWeight;
  borderRight?: BorderWeight;

  horizontalAlignment?: "Left" | "Center" | "Right" | "General";
  verticalAlignment?: "Top" | "Center" | "Bottom";
  wrapText?: boolean;
}

export interface NamedStyle {
  name: string;
  description: string;
  properties: CellStyle;
  builtIn: true;
}

export interface ResolvedCellStyle {
  properties: CellStyle;
  excelNumberFormat?: string;
  warnings: string[];
}
