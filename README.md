# Joby Aviation – Illustrative Air Taxi Route Financial Model

Case-study prototype for a Financial Systems Analyst role. It links air taxi operating drivers
(flights, completion rate, seats, load factor, fare, unit costs) to route-level financial
performance over 12 months.

> **Illustrative only.** Every input is a hypothetical assumption. Nothing here is a Joby
> Aviation forecast, guidance, or company-reported figure.

## Files

| File | Purpose |
|---|---|
| `Joby_Financial_Model.xlsx` | The workbook: Read Me, Executive Dashboard, Assumptions, Financial Model, Scenario Comparison, Sensitivity Analysis |
| `build_joby_model.py` | Generates the workbook with openpyxl (all calculations are Excel formulas) |
| `audit_model.py` | Cell-by-cell audit against an independent Python recalculation across 12 scenario and edge-case runs (results in `AUDIT.md`) |
| `validate_model.py` | Recalculates in headless LibreOffice, checks for formula errors, compares all scenarios to an independent Python calculation, tests the scenario selector, and runs negative tests on the integrity checks |

## Rebuild

```bash
python build_joby_model.py                 # writes Joby_Financial_Model.xlsx
python validate_model.py --write-cached    # recalculates, validates, stores cached values
python audit_model.py                      # full audit (LibreOffice recalculation, not Excel)
```

Requires Python 3 with `openpyxl`, plus LibreOffice (`soffice`) for validation.
