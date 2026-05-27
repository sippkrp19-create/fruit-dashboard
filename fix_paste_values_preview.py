from pathlib import Path

import openpyxl
from openpyxl.styles import Font


src = Path(r"C:\Users\sippk\OneDrive\เอกสาร\Open Design\dashboard\สวนผลไม้_บัญชี_ไม่วนสูตร.xlsx")
out = Path(r"C:\Users\sippk\OneDrive\เอกสาร\Open Design\dashboard\สวนผลไม้_บัญชี_วางค่าได้.xlsx")

wb = openpyxl.load_workbook(src)
form = wb["เพิ่มข้อมูล"]

# Make the preview row paste-values friendly:
# - ID uses ROWS over the real database range, avoiding COUNTA quirks.
# - Dates are formatted as Thai-friendly text, so Paste Values keeps them readable.
form["A18"] = '=TEXT(ROWS(\'ฐานข้อมูล\'!$A$5:INDEX(\'ฐานข้อมูล\'!$A:$A,MATCH("zzz",\'ฐานข้อมูล\'!$A:$A)))+1,"TXN-0000")'
form["B18"] = '=TEXT(\'เพิ่มข้อมูล\'!$C$4,"d/m/yyyy")'
form["Q18"] = '=TEXT(TODAY(),"d/m/yyyy")'
form["B18"].number_format = "@"
form["Q18"].number_format = "@"

form["A20"] = "วิธีใช้: กรอกข้อมูลให้ครบ แล้วคัดลอกแถว 18 ไปวางท้ายฐานข้อมูลด้วย Paste Values / วางค่าเท่านั้น"
form["A21"] = "วันที่ในแถวตัวอย่างถูกทำเป็นข้อความ d/m/yyyy เพื่อให้วางค่าแล้วไม่กลายเป็นเลข serial"
form["A20"].font = Font(italic=True, bold=True, color="8A3B35")
form["A21"].font = Font(italic=True, color="60706A")

for sheet in wb.worksheets:
    sheet.protection.sheet = False
    sheet.protection.enable = False

wb.calculation.fullCalcOnLoad = True
wb.calculation.forceFullCalc = True
wb.calculation.calcMode = "auto"
wb.save(out)
print(out)
