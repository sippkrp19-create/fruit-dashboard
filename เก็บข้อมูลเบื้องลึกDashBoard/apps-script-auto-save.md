# Apps Script: บันทึกข้อมูลจากฟอร์มอัตโนมัติ

ใช้ช่อง `เพิ่มข้อมูล!F14` เป็นตัวสั่งบันทึก

เมื่อเลือก `เสร็จเรียบร้อย` จะเพิ่มแถว `A18:R18` เข้า `ฐานข้อมูล` แล้วล้างฟอร์ม

```javascript
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== "เพิ่มข้อมูล") return;

  if (e.range.getA1Notation() !== "F14") return;

  const status = e.range.getDisplayValue();
  if (status !== "เสร็จเรียบร้อย") return;

  appendFormToDatabase_();
}

function appendFormToDatabase_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const form = ss.getSheetByName("เพิ่มข้อมูล");
    const db = ss.getSheetByName("ฐานข้อมูล");

    const status = form.getRange("F14").getDisplayValue();
    if (status !== "เสร็จเรียบร้อย") return;

    const values = form.getRange("A18:R18").getValues()[0];
    if (values.every(v => v === "")) return;

    const nextRow = Math.max(db.getLastRow() + 1, 5);
    db.getRange(nextRow, 1, 1, values.length).setValues([values]);

    form.getRange("C4").setValue(new Date());
    form.getRange("C5").setValue(2569);
    form.getRange("C6:C11").clearContent();

    form.getRange("F4").setValue(0);
    form.getRange("F5").setValue(0);
    form.getRange("F6").clearContent();
    form.getRange("F7").setValue("ปกติ");
    form.getRange("F8").clearContent();

    form.getRange("F14").setValue("ยังไม่เสร็จ");

    ss.toast("บันทึกเข้าฐานข้อมูลแล้ว และล้างฟอร์มเรียบร้อย", "สำเร็จ", 3);
  } finally {
    lock.releaseLock();
  }
}
```

