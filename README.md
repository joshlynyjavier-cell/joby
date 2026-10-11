# Joby Aviation – Illustrative Air Taxi Route Financial Model

Case-study prototype for a Financial Systems Analyst role. It links air taxi operating drivers
(flights, completion rate, seats, load factor, fare, unit costs) to route-level financial
performance over 12 months.

> **Illustrative only.** Every input is a hypothetical assumption. Nothing here is a Joby
> Aviation forecast, guidance, or company-reported figure.

## Files

| File | Purpose |
|---|---|
| `Joby_Financial_Model.xlsx` | The workbook: Read Me, Assumptions, Financial Model, Scenario Comparison |
| `build_joby_model.py` | Generates the workbook with openpyxl (all calculations are Excel formulas) |
| `validate_model.py` | Recalculates in headless LibreOffice, checks for formula errors, compares all scenarios to an independent Python calculation, tests the scenario selector, and runs negative tests on the integrity checks |

## Rebuild

```bash
python build_joby_model.py                 # writes Joby_Financial_Model.xlsx
python validate_model.py --write-cached    # recalculates, validates, stores cached values
```

Requires Python 3 with `openpyxl`, plus LibreOffice (`soffice`) for validation.
