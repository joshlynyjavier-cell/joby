# Model Audit – Joby_Financial_Model.xlsx

**Recalculation engine:** LibreOffice (headless). Microsoft Excel was **not** used, so the
results below show that the formulas evaluate correctly in LibreOffice. Excel's own
recalculation, including the "recalculate on open" flag, has not been tested.

## Method
For each test case, `audit_model.py` copies the workbook, changes the inputs, recalculates the
copy in LibreOffice and then:
1. Recomputes every monthly and annual line in Python from the raw inputs. It compares the
   results cell by cell with the Financial Model, the Scenario Comparison engine (all three
   cases) and the Scenario Comparison summary table (936 cells per case, plus 193 Sensitivity Analysis cells and 36 Executive Dashboard cells since those tabs were added).
2. Checks these identities on the workbook's own values, every month:
   revenue = passengers × fare;
   passengers = completed flights × seats per flight × load factor;
   contribution = revenue − direct flight costs − passenger-variable costs;
   route operating profit = contribution − fixed route costs;
   annual totals = sum of the months;
   annual revenue = Σ(monthly passengers × monthly fare).
3. Breakeven test: at the breakeven load factor, it rebuilds profit with passenger-variable
   costs scaled to that load factor. It then confirms contribution is zero (flight-level
   breakeven) and route operating profit is zero (route breakeven, fixed costs included).
4. Confirms that no formula returns an error and that all 19 integrity checks pass.

## Finding and fix
| # | Finding | Severity | Fix |
|---|---|---|---|
| 1 | When a ratio's denominator was zero, the workbook returned **0** instead of "not meaningful". For example, with zero load factor the contribution margin showed 0.0% while contribution was −$3.06M. With zero flights, annual cost per flight, fare and load factor showed 0. There were 141 cell mismatches across the zero-flight, zero-completion and zero-load-factor cases. All of them were this issue; there were no arithmetic errors. | Medium (misleading display) | All 18 ratio guards now return `"n.m."`. The two checks that used these ratios were updated to handle text: revenue is now reconciled month by month, and contribution per flight is guarded for zero flights. |
| 2 | When demand growth pushed the load factor past 100%, the model capped it silently. The "monthly load factor within 0–100%" check can never fail because of that cap. | Low | Added an information-only count of capped months below the check block. It is highlighted when above zero. |

## Results after the fix
| Case | Revenue | Contribution | Route op. profit | Contrib./flight | BE LF flight | BE LF route | Cells compared | Identities | Mismatches | Status | Sensitivity status | Dashboard status |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Base (as delivered) | $4,662,285 | $1,278,266 | $198,266 | $125.14 | 46.0% | 62.2% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Upside (selector) | $6,111,313 | $2,918,336 | $1,912,245 | $279.82 | 37.8% | 51.0% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Downside (selector) | $3,219,996 | $-282,303 | $-1,515,859 | $-29.17 | 58.6% | 81.4% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Zero flights (0 flights/day) | $0 | $0 | $-1,080,000 | n.m. | n.m. | n.m. | 1165 | 393 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (5 pass, 5 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Zero completion rate | $0 | $0 | $-1,080,000 | n.m. | n.m. | n.m. | 1165 | 393 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (5 pass, 5 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Completion rate 50% | $2,453,834 | $672,771 | $-407,229 | $125.14 | 46.0% | 76.8% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Completion rate 100% | $4,907,669 | $1,345,543 | $265,543 | $125.14 | 46.0% | 61.4% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Zero load factor | $0 | $-3,064,320 | $-4,144,320 | $-300.00 | 46.0% | 62.2% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (9 pass, 1 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Full load factor, no growth | $7,150,080 | $3,595,469 | $2,515,469 | $352.00 | 46.0% | 62.2% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Full load factor + demand growth (cap) | $7,150,080 | $3,595,469 | $2,515,469 | $352.00 | 46.0% | 62.2% | 1165 | 418 | 0 | ALL 21 CHECKS PASS; 11 capped month(s) | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Higher flight op. cost ($600/flight) | $4,662,285 | $-1,786,054 | $-2,866,054 | $-174.86 | 92.0% | 108.2% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Fare below cost per passenger ($10) | $266,416 | $-3,117,603 | $-4,197,603 | $-305.22 | n.m. | n.m. | 1165 | 392 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (4 pass, 6 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Sensitivity: $5 fare step, LF 50%-85% by 2.5pt | $4,662,285 | $1,278,266 | $198,266 | $125.14 | 46.0% | 62.2% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Sensitivity: wider $40 fare step, $0 band | $4,662,285 | $1,278,266 | $198,266 | $125.14 | 46.0% | 62.2% | 1165 | 418 | 0 | ALL 21 CHECKS PASS | ALL SENSITIVITY CHECKS PASS (10 pass, 0 n/a) | ALL 11 DASHBOARD CHECKS PASS |
| Sensitivity: LF axis past 100% (should FAIL) | $4,662,285 | $1,278,266 | $198,266 | $125.14 | 46.0% | 62.2% | 1165 | 418 | 0 | REVIEW: 1 CHECK(S) FAILING | REVIEW: 1 FAILING | ALL 11 DASHBOARD CHECKS PASS |

Dashboard chart references (read from the chart XML): "Passenger revenue", "Route operating profit / (loss)", 'Financial Model'!$C$29:$N$29, 'Financial Model'!$C$37:$N$37, 'Financial Model'!$C$8:$N$8 - linked to Financial Model revenue, route operating profit and period rows.

The breakeven load factor does not change with load factor (it depends only on fare, unit costs,
seats and fixed costs), so it is the same in the zero-load-factor and full-load-factor cases. That is expected.
Lowering the completion rate leaves the flight-level breakeven unchanged and raises the route
breakeven, because the same fixed costs are spread over fewer flights. A $600 cost per flight gives a route breakeven of
108.2% (not reachable). A fare below the variable cost per passenger gives "n.m.".

## Sensitivity Analysis tab (added later)
The tab shows contribution profit per completed flight = seats × load factor × (fare −
variable cost per passenger) − direct flight cost. It uses the Base scenario's Month 1 inputs,
all linked to the Assumptions tab (Base column). It has 9 fare columns centred on the Base fare
and 15 load factor rows, with conditional formatting for loss-making, breakeven (within ± a
$ band) and profitable combinations.

- **Breakeven definitions.** Flight-level breakeven load factor = direct flight cost ÷ (seats ×
  (fare − variable cost per passenger)). Route-level breakeven load factor = (direct flight
  cost + fixed route cost per completed flight) ÷ the same denominator. Both are shown at the
  Base fare and in every fare column.
- **What the audit checks.** For every run, the audit recomputes the fare and load factor
  axes, all 135 matrix cells, both breakeven rows, the legend counts and the linked inputs in
  Python. It also confirms that contribution is $0 at each column's flight-level breakeven,
  that Month 1 route operating profit is $0 at the route-level breakeven, and that the
  insight text contains no errors.
- **The tab's own checks.** It has 10 checks, including reconciliation of the Base cell and
  both breakeven values to the Month 1 Base calculation on Scenario Comparison. Its status
  feeds the Financial Model's overall status as check #20.
- **Extra runs.** Three runs edit only the matrix settings: a $5 fare step with a finer load
  factor axis; a $40 fare step with a $0 band; and a load factor axis that runs past 100%.
  The last was correctly flagged FAIL by both the sheet and the overall status.
- **Existing tabs preserved.** I compared every cell with the previous version. Scenario
  Comparison is unchanged. Assumptions and Read Me only point to the moved status cell, and
  Read Me's workbook map describes the new tab. Financial Model gained one check row, which
  moves the two rows below it down by one. No financial value changed.

## Executive Dashboard tab (added later)
This is a one-screen view sized for a PowerPoint screenshot (print area A1:N30, one landscape
page). It has six KPI cards, a monthly revenue and route operating profit chart, a Base /
Upside / Downside table and three Key Decisions boxes. Every number is a formula link to the
Financial Model tab (KPI cards, chart, decision evidence) or the Scenario Comparison tab
(scenario table).

- **Audit coverage.** In every run, each KPI card is compared with the Python result for the
  scenario shown. All 30 table cells are compared with the Python results for Base, Upside and
  Downside. The audit also confirms that the three decision texts contain no errors, and that
  the subtitle names the scenario shown.
- **The chart.** The audit reads the chart's references from the file's chart XML. They point
  at the Financial Model's passenger revenue, route operating profit and period rows (C:N).
- **The tab's own checks.** Nine reconciliation checks sit below the print area: each KPI card
  vs. the Financial Model, the cards vs. the Scenario Comparison active column, every table
  cell vs. Scenario Comparison, the chart months vs. the KPI totals, and the decision text has
  no errors. They feed the Financial Model's overall status as check #21.
- **Issue found and fixed during this audit.** With zero completed flights, or a fare below the
  cost per passenger, the Key Decisions evidence boxes collapsed to "n.m.". Each box now
  explains in plain language why its figures are not meaningful.
- **Existing tabs preserved.** I compared every cell with the previous version. Scenario
  Comparison and Sensitivity Analysis are unchanged. The Financial Model gained check #21. The
  status count went from 20 to 21 on the Assumptions and Read Me tabs. The Read Me workbook map
  was renumbered for the new tab order, with the dashboard second.

## Visual restyle to match the case-study deck (latest change)
The look was restyled to match `Joby_Case_Study.pdf`. The palette was sampled from the
rendered slides:

| Use | Colour |
|---|---|
| Background | `#0D131B` |
| Slate cards | `#2D363F` |
| Off-white cards | `#F9F9F7` |
| Rules | `#5B6770` |
| Muted text | `#ADB7BE` |

The typeface is Roboto Condensed, read from the PDF's embedded fonts.

- **Executive Dashboard:** rebuilt as a dark slide in the deck's layout, with the deck's
  header strip, off-white KPI cards and slate decision cards.
- **Other tabs:** a light version of the same palette.
- **No numbers changed.** Every value outside the dashboard is identical to the previous
  version. The only formula change is the Financial Model check that points at the
  dashboard's status cell, which moved when the layout changed. The full audit above was
  re-run after the restyle.

## Limitations
- **Excel not tested.** All recalculation was done in LibreOffice. The workbook uses only
  functions that work in both (SUM, IF, INDEX, MATCH, MIN, MAX, SUMPRODUCT, COUNTIF, EDATE,
  ISNUMBER), but Excel results and the "recalculate on open" behaviour were not tested.
- **The reference calculation is not fully independent.** The Python reference was written by
  the same author as the workbook, from the same definitions. It catches formula, reference and
  range errors, but it would not catch a mistake in a definition both share (for example, if
  cost growth should not apply to fixed costs).
- **Weighted vs. simple averages.** Seats are the same every month, so the weighted and simple
  annual load factors are identical in every case. The weighting is only visible where
  volumes and rates both change, e.g. the Upside weighted fare of $192.75 vs. a simple average
  of $192.63.
- **Annual breakeven assumes one load factor for the whole year.** It is calculated from annual
  totals, so it is the single uniform load factor at which the year breaks even.
- **The sensitivity matrix is Month 1 and Base only.** It does not include fare, demand or
  cost growth, and it does not follow the active-scenario selector.
- **The fare axis is floored at $0.** When the Base fare is smaller than 4 fare steps, the
  columns are no longer centred on it, and the Base-cell check reports N/A.
- **The Base-cell highlight may not show.** The bold box and underline come from conditional
  formatting, which LibreOffice's PDF export did not show. The text callout beside the matrix
  shows the same value either way.
- **The dashboard was only rendered in LibreOffice.** The layout was checked from a LibreOffice
  PDF export, for both the Base and Upside scenarios. Excel's fonts, chart rendering and column
  widths can differ slightly, so check the screenshot in Excel before it goes into a deck.
- **Roboto Condensed must be installed.** It is free from Google Fonts. The workbook names
  the font but does not embed it, so on a machine without it Excel substitutes another font
  and the dashboard spacing will look different. The previews here were rendered with
  Roboto Condensed installed.
- **Decision statements are fixed text.** The bold decision lines in Key Decisions are written
  judgements. Only the evidence text under them is formula-driven and updates with the inputs.
- **Scope of the edge cases.** Each was applied to the Base case, or to all cases for the shared
  route inputs. Negative and above-100% inputs are tested separately in `validate_model.py`.
