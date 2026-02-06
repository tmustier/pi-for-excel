/**
 * Whimsical rotating messages for the "Working…" indicator.
 *
 * Inspired by mitsuhiko/agent-stuff whimsical.ts, but tailored
 * for a spreadsheet / finance audience instead of a coding agent.
 */

const messages: string[] = [
  // ── Short — universally charming verbs ──────────────────

  "Schlepping…",
  "Combobulating…",
  "Vibing…",
  "Concocting…",
  "Transmuting…",
  "Pontificating…",
  "Cogitating…",
  "Noodling…",
  "Percolating…",
  "Ruminating…",
  "Simmering…",
  "Marinating…",
  "Fermenting…",
  "Brewing…",
  "Steeping…",
  "Contemplating…",
  "Musing…",
  "Pondering…",
  "Mulling…",
  "Daydreaming…",
  "Tinkering…",
  "Finagling…",
  "Wrangling…",
  "Meandering…",
  "Moseying…",
  "Pottering…",
  "Bumbling…",
  "Futzing…",
  "Kerfuffling…",
  "Bamboozling…",
  "Discombobulating…",
  "Recombobulating…",
  "Confabulating…",
  "Flummoxing…",
  "Befuddling…",
  "Effervescing…",
  "Fizzing…",
  "Bubbling…",
  "Scintillating…",
  "Improvising…",
  "Frolicking…",

  // ── Short — spreadsheet / finance flavored ──────────────

  "Calculating…",
  "Recalculating…",
  "Pivoting…",
  "Subtotaling…",
  "Autofilling…",
  "Tabulating…",
  "Auditing…",
  "Reconciling…",
  "Amortizing…",
  "Compounding…",
  "Accruing…",
  "Depreciating…",
  "Forecasting…",
  "Extrapolating…",
  "Interpolating…",

  // ── Long — universally fun ──────────────────────────────

  "Consulting the void…",
  "Asking the electrons…",
  "Negotiating with entropy…",
  "Waxing philosophical…",
  "Reading tea leaves…",
  "Shaking the magic 8-ball…",
  "Warming up the hamsters…",
  "Having a little think…",
  "Stroking chin thoughtfully…",
  "Squinting at the problem…",
  "Staring into the abyss…",
  "Abyss staring back…",
  "Achieving enlightenment…",
  "Consulting the oracle…",
  "Reticulating splines…",
  "Calibrating the flux capacitor…",
  "Hoping for the best…",
  "Manifesting solutions…",
  "Willing it into existence…",
  "Believing really hard…",
  "Reading the room…",
  "Kicking the tires…",
  "Dusting off the neurons…",
  "Rearranging deck chairs…",

  // ── Long — spreadsheet & Excel themed ───────────────────

  "Appeasing the circular reference…",
  "Bribing the formula bar…",
  "Reasoning with rounding errors…",
  "Pleading with the print preview…",
  "Herding cells into alignment…",
  "Wrestling with array formulas…",
  "Taming wild #REF! errors…",
  "Hunting for the missing penny…",
  "Consulting the spreadsheet gods…",
  "Reticulating spreadsheets…",
  "Massaging the margins…",
  "Having words with merged cells…",
  "Flirting with conditional formatting…",
  "Negotiating with column widths…",
  "Asking INDEX MATCH nicely…",
  "Befriending the Ribbon…",
  "Tiptoeing past the macros…",
  "Convincing the cells to cooperate…",
  "Feeding the data validation…",
  "Warming up the what-if analysis…",
  "Cross-referencing the worksheets…",
  "Auditing the formula trail…",
  "Tracing the precedents…",
  "Evaluating the dependents…",
  "Freezing the panes thoughtfully…",
  "Persuading OFFSET to cooperate…",
  "Checking under the hood of INDIRECT…",

  // ── Long — finance & modeling themed ────────────────────

  "Balancing the books…",
  "Crunching the numbers…",
  "Counting beans…",
  "Discounting future cash flows…",
  "Adjusting for seasonality…",
  "Running the Monte Carlo…",
  "Stress-testing the model…",
  "Sanity-checking the totals…",
  "Reconciling to the penny…",
  "Marking to market…",
  "Rolling forward the forecast…",
  "Building the bridge…",
  "Waterfalling the revenue…",
  "Sensitizing the assumptions…",
  "Triangulating the valuation…",
  "Normalizing the EBITDA…",
  "Checking the foot…",
  "Tying out the balance sheet…",
  "Hardcoding the overrides…",
  "Forgetting the mid-year convention…",
];

/** Pick a random message, avoiding the one currently shown. */
export function pickWhimsicalMessage(current?: string): string {
  if (messages.length <= 1) return messages[0] ?? "Working…";
  let msg: string;
  do {
    msg = messages[Math.floor(Math.random() * messages.length)];
  } while (msg === current && messages.length > 1);
  return msg;
}
