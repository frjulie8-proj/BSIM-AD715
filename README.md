# BSIM Excel Export

Generates a formatted Excel file from BSIM webapp export data. Two sheets: Goal Seek (break even analysis) and Data Table (three sensitivity tables: Price x Quantity, Cost x Quantity, Price x Cost).

## Warning

Do not run `bsim_export.py` directly. Do not edit it either. The only file you touch is `run.py`. Open it, update the 4 input values, and call `export_bsim_excel` — that is it.

## How to run locally

Edit the 4 input values in `run.py`, then:

```
python run.py
```

Output file is saved to the same folder as `BSIM_Export_Test.xlsx`.

## File structure

| File | Purpose |
|---|---|
| `run.py` | Only file you edit — input values go here |
| `bsim_export.py` | Core logic — do not edit |
| `test_bsim.py` | Run `python test_bsim.py` to verify everything works |

## Webapp integration

When integrating with the BSIM webapp, call this one function directly:

```python
from bsim_export import export_bsim_excel

export_bsim_excel(
    fixed_cost=...,    # Total Fixed Cost FY1
    var_cost=...,      # Variable Cost per Pint
    avg_price=...,     # Average Revenue per Pint (Total Revenue FY1 / Total Quantity FY1)
    qty_actual=...,    # Total Quantity Sold FY1
    output_path=...    # Full path for the output .xlsx file
)
```

All 4 values come from the student's 36 Month Pro Forma FY1 export. No other changes to `bsim_export.py` are needed.

## Dependencies

```
pip install openpyxl
```
