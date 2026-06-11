from openpyxl import Workbook

from scenario_manager import (
    create_scenario_manager_sheet
)

wb = Workbook()

wb.remove(wb.active)

create_scenario_manager_sheet(wb)

wb.save("Scenario_Manager.xlsx")

print("Scenario_Manager.xlsx created")