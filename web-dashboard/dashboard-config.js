/*
  Edit this file when Google Sheets has names that should be grouped differently.
  Examples:
  - If a row shows grade as "ไม่ระบุ" but you know it is "เกรดดำ",
    add it under gradeAliasesByProduct for that fruit.
  - If a product or grade is typed with a different spelling, add an alias here.
*/
window.DASHBOARD_CONFIG = {
  products: ["ทุเรียน", "ลองกอง", "มังคุด"],

  grades: [
    "เกรด1",
    "เกรด2",
    "เกรด3",
    "เกรด4",
    "เกรดพุ่ม",
    "เกรดปุก",
    "เกรดดำ",
    "เกรดร่วง",
  ],

  productAliases: {
    // "ทุเรียนหมอนทอง": "ทุเรียน",
    // "มังคุดคละ": "มังคุด",
    // "ลองกองเบอร์รวม": "ลองกอง",
  },

  gradeAliases: {
    // "g1": "เกรด1",
    // "เกรด 1": "เกรด1",
    // "ดำ": "เกรดดำ",
  },

  gradeAliasesByProduct: {
    "ทุเรียน": {
      // "ไม่ระบุ": "เกรด1",
    },
    "ลองกอง": {
      // "ไม่ระบุ": "เกรด1",
    },
    "มังคุด": {
      // "ไม่ระบุ": "เกรดดำ",
    },
  },

  /*
    Auto-detect grade from text in "ชื่อ/รายการ" when the grade column is empty.
    Add words that normally appear in your sheet.
  */
  gradeRules: [
    { product: "ทุเรียน", contains: ["เกรด1", "เกรด 1", "g1"], grade: "เกรด1" },
    { product: "ทุเรียน", contains: ["เกรด2", "เกรด 2", "g2"], grade: "เกรด2" },
    { product: "ทุเรียน", contains: ["เกรด3", "เกรด 3", "g3"], grade: "เกรด3" },
    { product: "มังคุด", contains: ["ดำ", "เกรดดำ"], grade: "เกรดดำ" },
    { product: "มังคุด", contains: ["ร่วง", "เกรดร่วง"], grade: "เกรดร่วง" },
    { product: "ลองกอง", contains: ["พุ่ม", "เกรดพุ่ม"], grade: "เกรดพุ่ม" },
  ],
};
