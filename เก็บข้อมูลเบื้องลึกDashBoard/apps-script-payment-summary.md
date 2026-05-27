# Apps Script: เพิ่มสรุปแยกวิธีรับ/จ่ายเงิน

วางโค้ดนี้ต่อท้าย Apps Script เดิม แล้วเลือก Run ฟังก์ชัน `setupPaymentSummary`

```javascript
function setupPaymentSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName("สรุป");
  if (!summary) throw new Error("ไม่พบชีต สรุป");

  summary.getRange("L3:O8").clearContent().clearFormat();

  summary.getRange("L3").setValue("แยกตามวิธีรับ/จ่ายเงิน");
  summary.getRange("L3:O3").merge();
  summary.getRange("L3").setFontWeight("bold").setFontColor("#1F4E3D");

  summary.getRange("L4:O4").setValues([[
    "วิธีรับ/จ่ายเงิน",
    "รายรับ",
    "รายจ่าย",
    "สุทธิ"
  ]]);

  summary.getRange("L5:L8").setValues([
    ["เงินสด"],
    ["เงินโอน"],
    ["ค้างรับ"],
    ["ค้างจ่าย"]
  ]);

  summary.getRange("M5").setFormula(paymentFormula_("I", "L5"));
  summary.getRange("N5").setFormula(paymentFormula_("J", "L5", true));
  summary.getRange("O5").setFormula("=M5-N5");

  summary.getRange("M6").setFormula(paymentFormulaTransfer_("I"));
  summary.getRange("N6").setFormula(paymentFormulaTransfer_("J", true));
  summary.getRange("O6").setFormula("=M6-N6");

  summary.getRange("M7").setFormula(paymentFormula_("I", "L7"));
  summary.getRange("N7").setFormula(paymentFormula_("J", "L7", true));
  summary.getRange("O7").setFormula("=M7-N7");

  summary.getRange("M8").setFormula(paymentFormula_("I", "L8"));
  summary.getRange("N8").setFormula(paymentFormula_("J", "L8", true));
  summary.getRange("O8").setFormula("=M8-N8");

  summary.getRange("L4:O4")
    .setBackground("#1F4E3D")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  summary.getRange("L5:O8").setBackground("#E9F4EC");
  summary.getRange("M5:O8").setNumberFormat("#,##0.00");
  summary.autoResizeColumns(12, 4);

  addPaymentDropdownOption_("เงินโอน");

  SpreadsheetApp.getActive().toast("เพิ่มสรุปเงินสด/เงินโอน/ค้างรับ/ค้างจ่ายแล้ว", "สำเร็จ", 3);
}

function paymentFormula_(valueCol, methodCell, includeExtraExpense) {
  const extra = includeExtraExpense
    ? `+IF($E$3="ทั้งหมด",SUMIFS(ฐานข้อมูล!$K$5:$K$2000,ฐานข้อมูล!$O$5:$O$2000,"ปกติ",ฐานข้อมูล!$B$5:$B$2000,$B$3,ฐานข้อมูล!$N$5:$N$2000,${methodCell}),SUMIFS(ฐานข้อมูล!$K$5:$K$2000,ฐานข้อมูล!$O$5:$O$2000,"ปกติ",ฐานข้อมูล!$B$5:$B$2000,$B$3,ฐานข้อมูล!$N$5:$N$2000,${methodCell},ฐานข้อมูล!$A$5:$A$2000,">="&DATE($B$3-543,$E$3,1),ฐานข้อมูล!$A$5:$A$2000,"<"&EDATE(DATE($B$3-543,$E$3,1),1)))+IF($E$3="ทั้งหมด",SUMIFS(ฐานข้อมูล!$L$5:$L$2000,ฐานข้อมูล!$O$5:$O$2000,"ปกติ",ฐานข้อมูล!$B$5:$B$2000,$B$3,ฐานข้อมูล!$N$5:$N$2000,${methodCell}),SUMIFS(ฐานข้อมูล!$L$5:$L$2000,ฐานข้อมูล!$O$5:$O$2000,"ปกติ",ฐานข้อมูล!$B$5:$B$2000,$B$3,ฐานข้อมูล!$N$5:$N$2000,${methodCell},ฐานข้อมูล!$A$5:$A$2000,">="&DATE($B$3-543,$E$3,1),ฐานข้อมูล!$A$5:$A$2000,"<"&EDATE(DATE($B$3-543,$E$3,1),1)))`
    : "";

  return `=IF($E$3="ทั้งหมด",SUMIFS(ฐานข้อมูล!$${valueCol}$5:$${valueCol}$2000,ฐานข้อมูล!$O$5:$O$2000,"ปกติ",ฐานข้อมูล!$B$5:$B$2000,$B$3,ฐานข้อมูล!$N$5:$N$2000,${methodCell}),SUMIFS(ฐานข้อมูล!$${valueCol}$5:$${valueCol}$2000,ฐานข้อมูล!$O$5:$O$2000,"ปกติ",ฐานข้อมูล!$B$5:$B$2000,$B$3,ฐานข้อมูล!$N$5:$N$2000,${methodCell},ฐานข้อมูล!$A$5:$A$2000,">="&DATE($B$3-543,$E$3,1),ฐานข้อมูล!$A$5:$A$2000,"<"&EDATE(DATE($B$3-543,$E$3,1),1)))${extra}`;
}

function paymentFormulaTransfer_(valueCol, includeExtraExpense) {
  const f1 = paymentFormula_(valueCol, '"โอน"', includeExtraExpense).slice(1);
  const f2 = paymentFormula_(valueCol, '"เงินโอน"', includeExtraExpense).slice(1);
  return `=${f1}+${f2}`;
}

function addPaymentDropdownOption_(value) {
  const listSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("รายการดรอปดาวน์");
  if (!listSheet) return;

  const values = listSheet.getRange("F5:F104").getValues().flat();
  if (values.includes(value)) return;

  const blankIndex = values.findIndex(v => v === "");
  if (blankIndex >= 0) {
    listSheet.getRange(5 + blankIndex, 6).setValue(value);
  }
}
```

