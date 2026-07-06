/**
 * Whimsical rotating messages for the "Working…" indicator.
 *
 * Inspired by mitsuhiko/agent-stuff whimsical.ts, but tailored
 * for a spreadsheet / finance audience instead of a coding agent.
 */

import { t } from "../language/index.js";

// Locale keys — resolved lazily via t() so the active language (set at boot,
// after module import) is respected. Do NOT call t() at module scope.
const messageKeys: string[] = [
  // ── Short — universally charming verbs ──────────────────

  "whimsical.schlepping",
  "whimsical.combobulating",
  "whimsical.vibing",
  "whimsical.concocting",
  "whimsical.transmuting",
  "whimsical.pontificating",
  "whimsical.cogitating",
  "whimsical.noodling",
  "whimsical.percolating",
  "whimsical.ruminating",
  "whimsical.simmering",
  "whimsical.marinating",
  "whimsical.fermenting",
  "whimsical.brewing",
  "whimsical.steeping",
  "whimsical.contemplating",
  "whimsical.musing",
  "whimsical.pondering",
  "whimsical.mulling",
  "whimsical.daydreaming",
  "whimsical.tinkering",
  "whimsical.finagling",
  "whimsical.wrangling",
  "whimsical.meandering",
  "whimsical.moseying",
  "whimsical.pottering",
  "whimsical.bumbling",
  "whimsical.futzing",
  "whimsical.kerfuffling",
  "whimsical.bamboozling",
  "whimsical.discombobulating",
  "whimsical.recombobulating",
  "whimsical.confabulating",
  "whimsical.flummoxing",
  "whimsical.befuddling",
  "whimsical.effervescing",
  "whimsical.fizzing",
  "whimsical.bubbling",
  "whimsical.scintillating",
  "whimsical.improvising",
  "whimsical.frolicking",

  // ── Short — spreadsheet / finance flavored ──────────────

  "whimsical.calculating",
  "whimsical.recalculating",
  "whimsical.pivoting",
  "whimsical.subtotaling",
  "whimsical.autofilling",
  "whimsical.tabulating",
  "whimsical.auditing",
  "whimsical.reconciling",
  "whimsical.amortizing",
  "whimsical.compounding",
  "whimsical.accruing",
  "whimsical.depreciating",
  "whimsical.forecasting",
  "whimsical.extrapolating",
  "whimsical.interpolating",

  // ── Long — universally fun ──────────────────────────────

  "whimsical.consulting_the_void",
  "whimsical.asking_the_electrons",
  "whimsical.negotiating_with_entropy",
  "whimsical.waxing_philosophical",
  "whimsical.reading_tea_leaves",
  "whimsical.shaking_the_magic_8_ball",
  "whimsical.warming_up_the_hamsters",
  "whimsical.having_a_little_think",
  "whimsical.stroking_chin_thoughtfully",
  "whimsical.squinting_at_the_problem",
  "whimsical.staring_into_the_abyss",
  "whimsical.abyss_staring_back",
  "whimsical.achieving_enlightenment",
  "whimsical.consulting_the_oracle",
  "whimsical.reticulating_splines",
  "whimsical.calibrating_the_flux_capacitor",
  "whimsical.hoping_for_the_best",
  "whimsical.manifesting_solutions",
  "whimsical.willing_it_into_existence",
  "whimsical.believing_really_hard",
  "whimsical.reading_the_room",
  "whimsical.kicking_the_tires",
  "whimsical.dusting_off_the_neurons",
  "whimsical.rearranging_deck_chairs",

  // ── Long — spreadsheet & Excel themed ───────────────────

  "whimsical.appeasing_the_circular_reference",
  "whimsical.bribing_the_formula_bar",
  "whimsical.reasoning_with_rounding_errors",
  "whimsical.pleading_with_the_print_preview",
  "whimsical.herding_cells_into_alignment",
  "whimsical.wrestling_with_array_formulas",
  "whimsical.taming_wild_ref_errors",
  "whimsical.hunting_for_the_missing_penny",
  "whimsical.consulting_the_spreadsheet_gods",
  "whimsical.reticulating_spreadsheets",
  "whimsical.massaging_the_margins",
  "whimsical.having_words_with_merged_cells",
  "whimsical.flirting_with_conditional_formatting",
  "whimsical.negotiating_with_column_widths",
  "whimsical.asking_index_match_nicely",
  "whimsical.befriending_the_ribbon",
  "whimsical.tiptoeing_past_the_macros",
  "whimsical.convincing_the_cells_to_cooperate",
  "whimsical.feeding_the_data_validation",
  "whimsical.warming_up_the_what_if_analysis",
  "whimsical.cross_referencing_the_worksheets",
  "whimsical.auditing_the_formula_trail",
  "whimsical.tracing_the_precedents",
  "whimsical.evaluating_the_dependents",
  "whimsical.freezing_the_panes_thoughtfully",
  "whimsical.persuading_offset_to_cooperate",
  "whimsical.checking_under_the_hood_of_indirect",

  // ── Long — finance & modeling themed ────────────────────

  "whimsical.balancing_the_books",
  "whimsical.crunching_the_numbers",
  "whimsical.counting_beans",
  "whimsical.discounting_future_cash_flows",
  "whimsical.adjusting_for_seasonality",
  "whimsical.running_the_monte_carlo",
  "whimsical.stress_testing_the_model",
  "whimsical.sanity_checking_the_totals",
  "whimsical.reconciling_to_the_penny",
  "whimsical.marking_to_market",
  "whimsical.rolling_forward_the_forecast",
  "whimsical.building_the_bridge",
  "whimsical.waterfalling_the_revenue",
  "whimsical.sensitizing_the_assumptions",
  "whimsical.triangulating_the_valuation",
  "whimsical.normalizing_the_ebitda",
  "whimsical.checking_the_foot",
  "whimsical.tying_out_the_balance_sheet",
  "whimsical.hardcoding_the_overrides",
  "whimsical.forgetting_the_mid_year_convention",
];

/** Pick a random message, avoiding the one currently shown. */
export function pickWhimsicalMessage(current?: string): string {
  if (messageKeys.length <= 1) {
    return messageKeys[0] ? t(messageKeys[0]) : t("working.default");
  }
  let msg: string;
  do {
    const key = messageKeys[Math.floor(Math.random() * messageKeys.length)] ?? messageKeys[0] ?? "working.default";
    msg = t(key);
  } while (msg === current && messageKeys.length > 1);
  return msg;
}
