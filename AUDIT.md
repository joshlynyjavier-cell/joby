# Model Audit – Joby_Financial_Model.xlsx

**Recalculation engine:** LibreOffice (headless). Microsoft Excel was **not** used, so the
results below show that the formulas evaluate correctly in LibreOffice. Excel's own
recalculation, including the "recalculate on open" flag, has not been tested.

## Method
For each test case, `audit_model.py` copies the workbook, changes the inputs, recalculates the
copy in LibreOffice and then:
1. Recomputes every monthly and annual line in Python from the raw inputs. It compares the
   results cell by cell with the Financial Model, the Scenario Comparison engine (all three
   cases) and the Scenario Comparison summary table (936 cells per case).
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
| Case | Revenue | Contribution | Route op. profit | Contrib./flight | BE LF flight | BE LF route | Cells compared | Identities | Mismatches | Status |
|---|---|---|---|---|---|---|---|---|---|---|
| Base (as delivered) | $4,662,285 | $1,278,266 | $198,266 | $125.14 | 46.0% | 62.2% | 936 | 264 | 0 | ALL 19 CHECKS PASS |
| Upside (selector) | $6,111,313 | $2,918,336 | $1,912,245 | $279.82 | 37.8% | 51.0% | 936 | 264 | 0 | ALL 19 CHECKS PASS |
| Downside (selector) | $3,219,996 | $-282,303 | $-1,515,859 | $-29.17 | 58.6% | 81.4% | 936 | 264 | 0 | ALL 19 CHECKS PASS |
| Zero flights (0 flights/day) | $0 | $0 | $-1,080,000 | n.m. | n.m. | n.m. | 936 | 240 | 0 | ALL 19 CHECKS PASS |
| Zero completion rate | $0 | $0 | $-1,080,000 | n.m. | n.m. | n.m. | 936 | 240 | 0 | ALL 19 CHECKS PASS |
| Completion rate 50% | $2,453,834 | $672,771 | $-407,229 | $125.14 | 46.0% | 76.8% | 936 | 264 | 0 | ALL 19 CHECKS PASS |
| Completion rate 100% | $4,907,669 | $1,345,543 | $265,543 | $125.14 | 46.0% | 61.4% | 936 | 264 | 0 | ALL 19 CHECKS PASS |
| Zero load factor | $0 | $-3,064,320 | $-4,144,320 | $-300.00 | 46.0% | 62.2% | 936 | 264 | 0 | ALL 19 CHECKS PASS |
| Full load factor, no growth | $7,150,080 | $3,595,469 | $2,515,469 | $352.00 | 46.0% | 62.2% | 936 | 264 | 0 | ALL 19 CHECKS PASS |
| Full load factor + demand growth (cap) | $7,150,080 | $3,595,469 | $2,515,469 | $352.00 | 46.0% | 62.2% | 936 | 264 | 0 | ALL 19 CHECKS PASS; 11 capped month(s) |
| Higher flight op. cost ($600/flight) | $4,662,285 | $-1,786,054 | $-2,866,054 | $-174.86 | 92.0% | 108.2% | 936 | 264 | 0 | ALL 19 CHECKS PASS |
| Fare below cost per passenger ($10) | $266,416 | $-3,117,603 | $-4,197,603 | $-305.22 | n.m. | n.m. | 936 | 240 | 0 | ALL 19 CHECKS PASS |

The breakeven load factor does not change with load factor (it depends only on fare, unit costs,
seats and fixed costs), so it is the same in the zero-load-factor and full-load-factor cases. That is expected.
Lowering the completion rate leaves the flight-level breakeven unchanged and raises the route
breakeven, because the same fixed costs are spread over fewer flights. A $600 cost per flight gives a route breakeven of
108.2% (not reachable). A fare below the variable cost per passenger gives "n.m.".

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
- **Scope of the edge cases.** Each was applied to the Base case, or to all cases for the shared
  route inputs. Negative and above-100% inputs are tested separately in `validate_model.py`.
