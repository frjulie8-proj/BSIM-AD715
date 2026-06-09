from bsim_export import export_bsim_excel

import datetime

# ── Edit these values from your BSIM 36 Month Pro Forma export ───────────────
# fixed_cost : Total Fixed Cost FY1
# var_cost   : Variable Cost per Pint (material + labor + other)
# avg_price  : Total Revenue FY1 / Total Quantity FY1
# qty_actual : Total Quantity Sold FY1

export_bsim_excel(
    fixed_cost=111111,
    var_cost=0.92,
    avg_price=3.62,
    qty_actual=2344334,
    output_path=f"data/BSIM_Export_Test_{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
)
