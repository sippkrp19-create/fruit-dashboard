from pathlib import Path

import openpyxl


src = Path(r"C:\Users\sippk\OneDrive\เอกสาร\Open Design\dashboard\สวนผลไม้_บัญชีและแดชบอร์ด_ใหม่.xlsx")
out = Path(r"C:\Users\sippk\OneDrive\เอกสาร\Open Design\dashboard\สวนผลไม้_บัญชีและแดชบอร์ด_แก้ไขได้.xlsx")

wb = openpyxl.load_workbook(src)
summary = wb["สรุป"]

summary["B6"] = '=SUMIFS(\'ฐานข้อมูล\'!$K:$K,\'ฐานข้อมูล\'!$P:$P,"ปกติ")+SUMIFS(\'ฐานข้อมูล\'!$L:$L,\'ฐานข้อมูล\'!$P:$P,"ปกติ")'
summary["B7"] = "=B5-B6"

for sheet in wb.worksheets:
    sheet.protection.sheet = False
    sheet.protection.enable = False

wb.calculation.fullCalcOnLoad = True
wb.calculation.forceFullCalc = True
wb.calculation.calcMode = "auto"

wb.save(out)
print(out)
