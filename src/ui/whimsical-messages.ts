/**
 * Whimsical rotating messages for the "Working…" indicator.
 *
 * Inspired by mitsuhiko/agent-stuff whimsical.ts, but tailored
 * for a spreadsheet / finance audience instead of a coding agent.
 */

import { t } from "../language/index.js";

const messages: string[] = [
  // ── Short — universally charming verbs ──────────────────

  t("whimsical.schlepping"),
  t("whimsical.combobulating"),
  t("whimsical.vibing"),
  t("whimsical.concocting"),
  t("whimsical.transmuting"),
  t("whimsical.pontificating"),
  t("whimsical.cogitating"),
  t("whimsical.noodling"),
  t("whimsical.percolating"),
  t("whimsical.ruminating"),
  t("whimsical.simmering"),
  t("whimsical.marinating"),
  t("whimsical.fermenting"),
  t("whimsical.brewing"),
  t("whimsical.steeping"),
  t("whimsical.contemplating"),
  t("whimsical.musing"),
  t("whimsical.pondering"),
  t("whimsical.mulling"),
  t("whimsical.daydreaming"),
  t("whimsical.tinkering"),
  t("whimsical.finagling"),
  t("whimsical.wrangling"),
  t("whimsical.meandering"),
  t("whimsical.moseying"),
  t("whimsical.pottering"),
  t("whimsical.bumbling"),
  t("whimsical.futzing"),
  t("whimsical.kerfuffling"),
  t("whimsical.bamboozling"),
  t("whimsical.discombobulating"),
  t("whimsical.recombobulating"),
  t("whimsical.confabulating"),
  t("whimsical.flummoxing"),
  t("whimsical.befuddling"),
  t("whimsical.effervescing"),
  t("whimsical.fizzing"),
  t("whimsical.bubbling"),
  t("whimsical.scintillating"),
  t("whimsical.improvising"),
  t("whimsical.frolicking"),

  // ── Short — spreadsheet / finance flavored ──────────────

  t("whimsical.calculating"),
  t("whimsical.recalculating"),
  t("whimsical.pivoting"),
  t("whimsical.subtotaling"),
  t("whimsical.autofilling"),
  t("whimsical.tabulating"),
  t("whimsical.auditing"),
  t("whimsical.reconciling"),
  t("whimsical.amortizing"),
  t("whimsical.compounding"),
  t("whimsical.accruing"),
  t("whimsical.depreciating"),
  t("whimsical.forecasting"),
  t("whimsical.extrapolating"),
  t("whimsical.interpolating"),

  // ── Long — universally fun ──────────────────────────────

  t("whimsical.consulting_the_void"),
  t("whimsical.asking_the_electrons"),
  t("whimsical.negotiating_with_entropy"),
  t("whimsical.waxing_philosophical"),
  t("whimsical.reading_tea_leaves"),
  t("whimsical.shaking_the_magic_8_ball"),
  t("whimsical.warming_up_the_hamsters"),
  t("whimsical.having_a_little_think"),
  t("whimsical.stroking_chin_thoughtfully"),
  t("whimsical.squinting_at_the_problem"),
  t("whimsical.staring_into_the_abyss"),
  t("whimsical.abyss_staring_back"),
  t("whimsical.achieving_enlightenment"),
  t("whimsical.consulting_the_oracle"),
  t("whimsical.reticulating_splines"),
  t("whimsical.calibrating_the_flux_capacitor"),
  t("whimsical.hoping_for_the_best"),
  t("whimsical.manifesting_solutions"),
  t("whimsical.willing_it_into_existence"),
  t("whimsical.believing_really_hard"),
  t("whimsical.reading_the_room"),
  t("whimsical.kicking_the_tires"),
  t("whimsical.dusting_off_the_neurons"),
  t("whimsical.rearranging_deck_chairs"),

  // ── Long — spreadsheet & Excel themed ───────────────────

  t("whimsical.appeasing_the_circular_reference"),
  t("whimsical.bribing_the_formula_bar"),
  t("whimsical.reasoning_with_rounding_errors"),
  t("whimsical.pleading_with_the_print_preview"),
  t("whimsical.herding_cells_into_alignment"),
  t("whimsical.wrestling_with_array_formulas"),
  t("whimsical.taming_wild_ref_errors"),
  t("whimsical.hunting_for_the_missing_penny"),
  t("whimsical.consulting_the_spreadsheet_gods"),
  t("whimsical.reticulating_spreadsheets"),
  t("whimsical.massaging_the_margins"),
  t("whimsical.having_words_with_merged_cells"),
  t("whimsical.flirting_with_conditional_formatting"),
  t("whimsical.negotiating_with_column_widths"),
  t("whimsical.asking_index_match_nicely"),
  t("whimsical.befriending_the_ribbon"),
  t("whimsical.tiptoeing_past_the_macros"),
  t("whimsical.convincing_the_cells_to_cooperate"),
  t("whimsical.feeding_the_data_validation"),
  t("whimsical.warming_up_the_what_if_analysis"),
  t("whimsical.cross_referencing_the_worksheets"),
  t("whimsical.auditing_the_formula_trail"),
  t("whimsical.tracing_the_precedents"),
  t("whimsical.evaluating_the_dependents"),
  t("whimsical.freezing_the_panes_thoughtfully"),
  t("whimsical.persuading_offset_to_cooperate"),
  t("whimsical.checking_under_the_hood_of_indirect"),

  // ── Long — finance & modeling themed ────────────────────

  t("whimsical.balancing_the_books"),
  t("whimsical.crunching_the_numbers"),
  t("whimsical.counting_beans"),
  t("whimsical.discounting_future_cash_flows"),
  t("whimsical.adjusting_for_seasonality"),
  t("whimsical.running_the_monte_carlo"),
  t("whimsical.stress_testing_the_model"),
  t("whimsical.sanity_checking_the_totals"),
  t("whimsical.reconciling_to_the_penny"),
  t("whimsical.marking_to_market"),
  t("whimsical.rolling_forward_the_forecast"),
  t("whimsical.building_the_bridge"),
  t("whimsical.waterfalling_the_revenue"),
  t("whimsical.sensitizing_the_assumptions"),
  t("whimsical.triangulating_the_valuation"),
  t("whimsical.normalizing_the_ebitda"),
  t("whimsical.checking_the_foot"),
  t("whimsical.tying_out_the_balance_sheet"),
  t("whimsical.hardcoding_the_overrides"),
  t("whimsical.forgetting_the_mid_year_convention"),
];

/** Pick a random message, avoiding the one currently shown. */
export function pickWhimsicalMessage(current?: string): string {
  if (messages.length <= 1) return messages[0] ?? t("working.default");
  let msg: string;
  do {
    msg = messages[Math.floor(Math.random() * messages.length)];
  } while (msg === current && messages.length > 1);
  return msg;
}
