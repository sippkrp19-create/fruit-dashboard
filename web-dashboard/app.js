const SHEET_ID = "1YEsURhIz4oyM9sVebBrXfHId9TRn1BKOVEQVIqvVBHA";
const SHEET_NAME = "ฐานข้อมูล";
const LIST_SHEET_NAME = "รายการดรอปดาวน์";
const GVIZ_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?sheet=${encodeURIComponent(SHEET_NAME)}&tqx=out:json;responseHandler:__dashboardData`;
const LIST_GVIZ_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?sheet=${encodeURIComponent(LIST_SHEET_NAME)}&tqx=out:json;responseHandler:__dashboardListData`;

const MONTHS = [
  "ม.ค.",
  "ก.พ.",
  "มี.ค.",
  "เม.ย.",
  "พ.ค.",
  "มิ.ย.",
  "ก.ค.",
  "ส.ค.",
  "ก.ย.",
  "ต.ค.",
  "พ.ย.",
  "ธ.ค.",
];

const PRODUCT_TYPES = ["ทุเรียน", "ลองกอง", "มังคุด"];
const GRADE_TYPES = ["เกรด1", "เกรด2", "เกรด3", "เกรด4", "เกรดพุ่ม", "เกรดปุก", "เกรดดำ", "เกรดร่วง"];
const COLORS = ["#d9a441", "#3f8a5a", "#4c7fb8", "#e08a3e", "#8e6bbf", "#c9604b", "#6aa6a1", "#9bbf4a"];

const state = {
  rows: [],
  filters: defaultFilters(),
  donutMetric: "revenue",
  search: "",
  dropdowns: {
    years: [],
    products: PRODUCT_TYPES,
    grades: GRADE_TYPES,
    payments: [],
    statuses: [],
  },
};

function defaultFilters() {
  return {
    year: "all",
    month: "all",
    product: "all",
    grade: "all",
    payment: "all",
    status: "ปกติ",
  };
}

const $ = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat("th-TH", { maximumFractionDigits: 2 });
const money = (value) => fmt.format(Math.round((value + Number.EPSILON) * 100) / 100);

function toNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (!value) return 0;
  const cleaned = String(value).replace(/,/g, "").trim();
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "string") {
    const gviz = value.match(/Date\((\d+),(\d+),(\d+)/);
    if (gviz) return new Date(Number(gviz[1]), Number(gviz[2]), Number(gviz[3]));
    const parts = value.match(/(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
    if (parts) {
      let year = Number(parts[3]);
      if (year < 100) year += 2500;
      if (year > 2400) year -= 543;
      return new Date(year, Number(parts[2]) - 1, Number(parts[1]));
    }
  }
  return null;
}

function cellValue(cell) {
  if (!cell) return "";
  if (cell.v !== undefined) return cell.v;
  if (cell.f !== undefined) return cell.f;
  return "";
}

function normalizePayment(value) {
  const text = String(value || "").trim();
  if (text === "โอน") return "เงินโอน";
  return text || "ไม่ระบุ";
}

function cleanPaymentOptions(values) {
  return unique(values.map(normalizePayment).filter((value) => value && value !== "ไม่ระบุ"));
}

function normalizeRows(table) {
  const headers = table.cols.map((col) => String(col.label || "").trim());
  const index = Object.fromEntries(headers.map((header, i) => [header, i]));
  return table.rows
    .map((row) => {
      const c = row.c || [];
      const date = parseDate(cellValue(c[index["วันที่"]]));
      const season = toNumber(cellValue(c[index["รอบ/ฤดูกาล"]])) || (date ? date.getFullYear() + 543 : "");
      const revenue = toNumber(cellValue(c[index["รายรับ"]]));
      const expense = toNumber(cellValue(c[index["รายจ่าย"]])) + toNumber(cellValue(c[index["ค่าตัด"]])) + toNumber(cellValue(c[index["เบิก"]]));
      const netCell = toNumber(cellValue(c[index["สุทธิ"]]));
      return {
        date,
        dateText: date ? `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}` : "",
        year: String(season || ""),
        month: date ? date.getMonth() + 1 : "",
        product: String(cellValue(c[index["สินค้า"]]) || "ไม่ระบุ").trim(),
        type: String(cellValue(c[index["ประเภทรายการ"]]) || "").trim(),
        grade: String(cellValue(c[index["สายพันธุ์/เกรด"]]) || "").trim(),
        name: String(cellValue(c[index["ชื่อ/รายการ"]]) || "").trim(),
        price: toNumber(cellValue(c[index["ราคาต่อกก."]])),
        qty: toNumber(cellValue(c[index["ปริมาณกก."]])),
        revenue,
        expense,
        net: netCell || revenue - expense,
        payment: normalizePayment(cellValue(c[index["วิธีรับ/จ่ายเงิน"]])),
        status: String(cellValue(c[index["สถานะรายการ"]]) || "").trim() || "ปกติ",
        note: String(cellValue(c[index["หมายเหตุ"]]) || "").trim(),
      };
    })
    .filter((row) => row.date || row.product !== "ไม่ระบุ" || row.revenue || row.expense);
}

function loadData() {
  $("errorBox").hidden = true;
  $("insightText").textContent = "กำลังดึงข้อมูลจาก Google Sheets...";
  document.querySelectorAll("script[data-gviz]").forEach((script) => script.remove());
  window.__dashboardData = (response) => {
    try {
      if (!response || !response.table) throw new Error("อ่านข้อมูลจาก Google Sheets ไม่ได้");
      state.rows = normalizeRows(response.table);
      loadDropdownData();
    } catch (error) {
      showError(error.message);
    }
  };
  const script = document.createElement("script");
  script.src = `${GVIZ_URL}&cacheBust=${Date.now()}`;
  script.dataset.gviz = "true";
  script.onerror = () => showError("ดึงข้อมูลไม่ได้ กรุณาตั้งค่า Google Sheets เป็น Anyone with the link can view");
  document.body.appendChild(script);
}

function loadDropdownData() {
  let settled = false;
  window.__dashboardListData = (response) => {
    settled = true;
    if (response?.table) state.dropdowns = parseDropdowns(response.table);
    finishLoad();
  };
  const script = document.createElement("script");
  script.src = `${LIST_GVIZ_URL}&cacheBust=${Date.now()}`;
  script.dataset.gviz = "true";
  script.onerror = () => {
    settled = true;
    finishLoad();
  };
  document.body.appendChild(script);
  setTimeout(() => {
    if (!settled) finishLoad();
  }, 2500);
}

function finishLoad() {
  setupFilters();
  render();
  $("lastUpdated").textContent = `อัปเดตล่าสุด ${new Date().toLocaleString("th-TH")}`;
}

function parseDropdowns(table) {
  const cols = table.cols.map((col) => String(col.label || "").trim());
  const valuesByCol = {};
  cols.forEach((label, colIndex) => {
    valuesByCol[label] = table.rows
      .map((row) => cellValue((row.c || [])[colIndex]))
      .filter((value) => value !== null && value !== undefined && value !== "")
      .map(String);
  });
  return {
    years: valuesByCol["ปี/ฤดูกาล"] || [],
    products: mergePreferred(PRODUCT_TYPES, valuesByCol["สินค้า"] || []),
    grades: mergePreferred(GRADE_TYPES, valuesByCol["สายพันธุ์/เกรด"] || []),
    payments: cleanPaymentOptions(valuesByCol["วิธีรับ/จ่ายเงิน"] || []),
    statuses: valuesByCol["สถานะรายการ"] || [],
  };
}

function showError(message) {
  $("insightText").textContent = "ยังไม่สามารถโหลดข้อมูลได้";
  $("errorBox").hidden = false;
  $("errorBox").innerHTML = `<strong>โหลดข้อมูลไม่สำเร็จ</strong><br>${message}<br>ตรวจว่าไฟล์ Google Sheets แชร์แบบ Anyone with the link can view และมีชีตชื่อ ฐานข้อมูล`;
}

function unique(values) {
  return [...new Set(values.filter(Boolean).map(String))].sort((a, b) => a.localeCompare(b, "th"));
}

function setOptions(id, values, allLabel, labeler = (value) => value, selectedValue = null) {
  const select = $(id);
  const current = selectedValue ?? select.value ?? "all";
  select.innerHTML = `<option value="all">${allLabel}</option>${values.map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(labeler(v))}</option>`).join("")}`;
  select.value = values.includes(current) ? current : "all";
}

function setupFilters() {
  setOptions("yearFilter", mergePreferred(state.dropdowns.years, unique(state.rows.map((r) => r.year)).sort()), "ทุกปี", (value) => value, state.filters.year);
  setOptions("monthFilter", MONTHS.map((m, i) => `${i + 1}|${m}`), "ทุกเดือน", (value) => String(value).split("|")[1], state.filters.month);
  setOptions("productFilter", mergePreferred(state.dropdowns.products, unique(state.rows.map((r) => r.product))), "ทุกสินค้า", (value) => value, state.filters.product);
  setOptions("gradeFilter", gradeOptionsForCurrentProduct(), "ทุกสายพันธุ์/เกรด", (value) => value, state.filters.grade);
  setOptions("paymentFilter", mergePreferred(state.dropdowns.payments, cleanPaymentOptions(state.rows.map((r) => r.payment))), "ทุกวิธี", (value) => value, state.filters.payment);
  setOptions("statusFilter", mergePreferred(state.dropdowns.statuses, unique(state.rows.map((r) => r.status))), "ทุกสถานะ", (value) => value, state.filters.status);
}

function mergePreferred(preferred, values) {
  return [...preferred, ...values.filter((value) => !preferred.includes(value))];
}

function gradeOptionsForCurrentProduct() {
  const fromRows = state.rows
    .filter((row) => state.filters.product === "all" || row.product === state.filters.product)
    .map((row) => row.grade);
  return mergePreferred(state.dropdowns.grades, unique(fromRows));
}

function filteredRows(custom = {}) {
  const filters = { ...state.filters, ...custom };
  return state.rows.filter((row) => {
    if (filters.year !== "all" && row.year !== filters.year) return false;
    if (filters.month !== "all" && String(row.month) !== String(filters.month).split("|")[0]) return false;
    if (filters.product !== "all" && row.product !== filters.product) return false;
    if (filters.grade !== "all" && row.grade !== filters.grade) return false;
    if (filters.payment !== "all" && row.payment !== filters.payment) return false;
    if (filters.status !== "all" && row.status !== filters.status) return false;
    return true;
  });
}

function totals(rows) {
  return rows.reduce(
    (acc, row) => {
      acc.revenue += row.revenue;
      acc.expense += row.expense;
      acc.net += row.net;
      acc.qty += row.qty;
      if (row.payment === "ค้างรับ") acc.pendingIn += row.revenue || Math.max(row.net, 0);
      if (row.payment === "ค้างจ่าย") acc.pendingOut += row.expense || Math.max(-row.net, 0);
      return acc;
    },
    { revenue: 0, expense: 0, net: 0, qty: 0, pendingIn: 0, pendingOut: 0 }
  );
}

function comparisonFilters() {
  const f = { ...state.filters };
  if (f.month !== "all") {
    const month = Number(String(f.month).split("|")[0]);
    const year = Number(f.year);
    if (month > 1) f.month = `${month - 1}|${MONTHS[month - 2]}`;
    else if (Number.isFinite(year)) {
      f.year = String(year - 1);
      f.month = `12|${MONTHS[11]}`;
    }
    return f;
  }
  if (f.year !== "all") f.year = String(Number(f.year) - 1);
  return f;
}

function trend(current, previous) {
  if (!previous) return "ไม่มีข้อมูลเปรียบเทียบ";
  const pct = ((current - previous) / Math.abs(previous)) * 100;
  const arrow = pct >= 0 ? "▲" : "▼";
  return `${arrow} ${Math.abs(pct).toFixed(1)}% จากช่วงก่อน`;
}

function renderKpis(rows) {
  const t = totals(rows);
  const p = totals(filteredRows(comparisonFilters()));
  $("kpiRevenue").textContent = money(t.revenue);
  $("kpiExpense").textContent = money(t.expense);
  $("kpiNet").textContent = money(t.net);
  $("kpiQty").textContent = money(t.qty);
  $("kpiAvg").textContent = t.qty ? money(t.revenue / t.qty) : "0";
  $("kpiPending").textContent = money(t.pendingIn + t.pendingOut);
  $("pendingSplit").textContent = `ค้างรับ ${money(t.pendingIn)} | ค้างจ่าย ${money(t.pendingOut)}`;
  setTrend("trendRevenue", trend(t.revenue, p.revenue), t.revenue - p.revenue);
  setTrend("trendExpense", trend(t.expense, p.expense), -(t.expense - p.expense));
  setTrend("trendNet", trend(t.net, p.net), t.net - p.net);
}

function setTrend(id, text, direction) {
  const node = $(id);
  node.textContent = text;
  node.className = direction >= 0 ? "trend-up" : "trend-down";
}

function groupBy(rows, key, valueKey) {
  const map = new Map();
  rows.forEach((row) => {
    const name = row[key] || "ไม่ระบุ";
    map.set(name, (map.get(name) || 0) + row[valueKey]);
  });
  return [...map.entries()]
    .map(([name, value]) => ({ name, value }))
    .filter((d) => Math.abs(d.value) > 0.0001)
    .sort((a, b) => Math.abs(b.value) - Math.abs(a.value));
}

function renderDonut(rows) {
  const metric = state.donutMetric;
  const products = mergePreferred(PRODUCT_TYPES, unique(rows.map((row) => row.product))).slice(0, 3);
  const data = products.map((name) => ({
    name,
    value: rows.filter((row) => row.product === name).reduce((sum, row) => sum + row[metric], 0),
  }));
  const total = data.reduce((sum, d) => sum + Math.abs(d.value), 0);
  if (!total) {
    $("donutChart").innerHTML = `<div class="empty-state">ไม่มีข้อมูลสำหรับกราฟนี้</div>`;
    $("donutLegend").innerHTML = "";
    return;
  }
  const positiveData = data.filter((d) => Math.abs(d.value) > 0);
  let angle = -90;
  const cx = 120;
  const cy = 120;
  const r = 86;
  const width = 240;
  const paths = positiveData.length === 1
    ? `<circle class="donut-slice" data-product="${escapeHtml(positiveData[0].name)}" cx="${cx}" cy="${cy}" r="70" fill="none" stroke="${COLORS[data.findIndex((d) => d.name === positiveData[0].name) % COLORS.length]}" stroke-width="68"><title>${escapeHtml(positiveData[0].name)} ${money(positiveData[0].value)}</title></circle>`
    : positiveData.map((d) => {
      const i = data.findIndex((item) => item.name === d.name);
      const slice = (Math.abs(d.value) / total) * 360;
      const path = describeArc(cx, cy, r, angle, angle + slice);
      angle += slice;
      return `<path class="donut-slice" data-product="${escapeHtml(d.name)}" d="${path}" fill="${COLORS[i % COLORS.length]}" stroke="#fbfff8" stroke-width="4"><title>${escapeHtml(d.name)} ${money(d.value)}</title></path>`;
    })
    .join("");
  $("donutChart").innerHTML = `
    <svg viewBox="0 0 ${width} ${width}">
      ${paths}
      <circle cx="${cx}" cy="${cy}" r="52" fill="#fbfff8"></circle>
      <text x="${cx}" y="${cy - 4}" text-anchor="middle" font-size="13" fill="#607267">รวม</text>
      <text x="${cx}" y="${cy + 20}" text-anchor="middle" font-size="20" font-weight="700" fill="#1f5b41">${money(total)}</text>
    </svg>`;
  $("donutLegend").innerHTML = data
    .map((d, i) => `<button class="legend-item chipless" type="button" data-product="${escapeHtml(d.name)}"><span class="swatch" style="background:${COLORS[i % COLORS.length]}"></span><span>${escapeHtml(d.name)}</span><strong>${money(d.value)}</strong></button>`)
    .join("");
}

function renderFruitGradeDonuts(rows) {
  const metric = state.donutMetric;
  const products = mergePreferred(PRODUCT_TYPES, unique(rows.map((row) => row.product))).slice(0, 3);
  $("fruitDonutGrid").innerHTML = products.map((product, productIndex) => {
    const productRows = rows.filter((row) => row.product === product);
    const gradeData = groupBy(productRows, "grade", metric);
    const total = gradeData.reduce((sum, d) => sum + Math.abs(d.value), 0);
    const baseColor = COLORS[productIndex % COLORS.length];
    const chart = total
      ? renderSingleFruitDonut(product, gradeData, total, baseColor)
      : `<div class="empty-state">ยังไม่มีข้อมูล</div>`;
    const legend = gradeData.length
      ? gradeData.slice(0, 5).map((grade, gradeIndex) => {
        const color = shadeColor(baseColor, 8 + gradeIndex * 11);
        return `<button type="button" data-product="${escapeHtml(product)}" data-grade="${escapeHtml(grade.name)}"><span style="background:${color}"></span><small>${escapeHtml(grade.name || "ไม่ระบุ")}</small><strong>${money(grade.value)}</strong></button>`;
      }).join("")
      : `<small>ไม่มีสายพันธุ์/เกรดในตัวกรองนี้</small>`;
    return `<article class="fruit-donut-card">
      <h3><span>${escapeHtml(product)}</span><strong>${money(total)}</strong></h3>
      ${chart}
      <div class="mini-legend">${legend}</div>
    </article>`;
  }).join("");
}

function renderSingleFruitDonut(product, gradeData, total, baseColor) {
  const cx = 90;
  const cy = 90;
  let angle = -90;
  const positiveData = gradeData.filter((d) => Math.abs(d.value) > 0);
  const slices = positiveData.length === 1
    ? `<circle class="donut-slice" data-product="${escapeHtml(product)}" data-grade="${escapeHtml(positiveData[0].name)}" cx="${cx}" cy="${cy}" r="55" fill="none" stroke="${shadeColor(baseColor, 12)}" stroke-width="34"><title>${escapeHtml(product)} / ${escapeHtml(positiveData[0].name)} ${money(positiveData[0].value)}</title></circle>`
    : positiveData.map((grade, gradeIndex) => {
      const slice = (Math.abs(grade.value) / total) * 360;
      const start = angle;
      const end = angle + slice;
      angle = end;
      return `<path class="donut-slice" data-product="${escapeHtml(product)}" data-grade="${escapeHtml(grade.name)}" d="${describeDonutArc(cx, cy, 38, 72, start, end)}" fill="${shadeColor(baseColor, 8 + gradeIndex * 11)}" stroke="#fbfff8" stroke-width="2"><title>${escapeHtml(product)} / ${escapeHtml(grade.name)} ${money(grade.value)}</title></path>`;
    }).join("");
  return `<svg viewBox="0 0 180 180">
    ${slices}
    <circle cx="${cx}" cy="${cy}" r="33" fill="#fbfff8"></circle>
    <text x="${cx}" y="${cy - 4}" text-anchor="middle" font-size="11" fill="#607267">รวม</text>
    <text x="${cx}" y="${cy + 18}" text-anchor="middle" font-size="18" font-weight="700" fill="#1f5b41">${money(total)}</text>
  </svg>`;
}

function describeArc(cx, cy, r, startAngle, endAngle) {
  const start = polarToCartesian(cx, cy, r, endAngle);
  const end = polarToCartesian(cx, cy, r, startAngle);
  const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
  return ["M", start.x, start.y, "A", r, r, 0, largeArcFlag, 0, end.x, end.y, "L", cx, cy, "Z"].join(" ");
}

function describeDonutArc(cx, cy, innerR, outerR, startAngle, endAngle) {
  if (Math.abs(endAngle - startAngle) >= 359.99) {
    endAngle = startAngle + 359.99;
  }
  const outerStart = polarToCartesian(cx, cy, outerR, endAngle);
  const outerEnd = polarToCartesian(cx, cy, outerR, startAngle);
  const innerStart = polarToCartesian(cx, cy, innerR, startAngle);
  const innerEnd = polarToCartesian(cx, cy, innerR, endAngle);
  const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
  return [
    "M", outerStart.x, outerStart.y,
    "A", outerR, outerR, 0, largeArcFlag, 0, outerEnd.x, outerEnd.y,
    "L", innerStart.x, innerStart.y,
    "A", innerR, innerR, 0, largeArcFlag, 1, innerEnd.x, innerEnd.y,
    "Z",
  ].join(" ");
}

function shadeColor(hex, percent) {
  const clean = hex.replace("#", "");
  const num = parseInt(clean, 16);
  const amt = Math.round(2.55 * percent);
  const r = Math.max(0, Math.min(255, (num >> 16) + amt));
  const g = Math.max(0, Math.min(255, ((num >> 8) & 0x00ff) + amt));
  const b = Math.max(0, Math.min(255, (num & 0x0000ff) + amt));
  return `#${(0x1000000 + r * 0x10000 + g * 0x100 + b).toString(16).slice(1)}`;
}

function polarToCartesian(cx, cy, r, angle) {
  const rad = ((angle - 90) * Math.PI) / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}

function renderMonthly(rows) {
  const data = Array.from({ length: 12 }, (_, i) => ({ month: i + 1, label: MONTHS[i], revenue: 0, expense: 0, net: 0 }));
  rows.forEach((row) => {
    if (!row.month) return;
    const d = data[row.month - 1];
    d.revenue += row.revenue;
    d.expense += row.expense;
    d.net += row.net;
  });
  const max = Math.max(1, ...data.flatMap((d) => [d.revenue, d.expense, Math.abs(d.net)]));
  const w = 900;
  const h = 260;
  const left = 46;
  const bottom = 38;
  const plotH = h - 64;
  const barW = 13;
  const gap = (w - left - 24) / 12;
  const bars = data
    .map((d, i) => {
      const x = left + i * gap + 8;
      const revH = (d.revenue / max) * plotH;
      const expH = (d.expense / max) * plotH;
      const netH = (Math.abs(d.net) / max) * plotH;
      const base = h - bottom;
      return `
        <g class="bar-click" data-month="${d.month}|${d.label}">
          <rect x="${x}" y="${base - revH}" width="${barW}" height="${revH}" rx="3" fill="#3f8a5a"><title>${d.label} รายรับ ${money(d.revenue)}</title></rect>
          <rect x="${x + barW + 3}" y="${base - expH}" width="${barW}" height="${expH}" rx="3" fill="#c9604b"><title>${d.label} รายจ่าย ${money(d.expense)}</title></rect>
          <rect x="${x + (barW + 3) * 2}" y="${base - netH}" width="${barW}" height="${netH}" rx="3" fill="#d9a441"><title>${d.label} สุทธิ ${money(d.net)}</title></rect>
          <text x="${x + 21}" y="${h - 14}" text-anchor="middle" font-size="12" fill="#607267">${d.label}</text>
        </g>`;
    })
    .join("");
  $("monthlyChart").innerHTML = `
    <svg viewBox="0 0 ${w} ${h}">
      <line x1="${left}" y1="${h - bottom}" x2="${w - 10}" y2="${h - bottom}" stroke="#d6e2d1"/>
      <text x="${left}" y="18" font-size="12" fill="#607267">เขียว=รายรับ แดง=รายจ่าย ทอง=สุทธิ</text>
      ${bars}
    </svg>`;
}

function renderPaymentCards(rows) {
  const methods = ["เงินสด", "เงินโอน", "ค้างรับ", "ค้างจ่าย"];
  $("paymentCards").innerHTML = methods
    .map((method) => {
      const r = rows.filter((row) => row.payment === method || (method === "เงินโอน" && row.payment === "โอน"));
      const t = totals(r);
      const value = method === "ค้างรับ" ? t.pendingIn : method === "ค้างจ่าย" ? t.pendingOut : t.revenue - t.expense;
      return `<article class="payment-card" data-payment="${method}">
        <span>${method}</span>
        <strong>${money(value)}</strong>
        <small>รายรับ ${money(t.revenue)} | รายจ่าย ${money(t.expense)}</small>
      </article>`;
    })
    .join("");
}

function renderRanks(rows) {
  renderRank("productRank", groupBy(rows, "product", "revenue").slice(0, 6), "product");
  renderRank("expenseRank", groupBy(rows.filter((r) => r.expense > 0), "name", "expense").slice(0, 6), "name");
  renderRank("gradeRank", groupBy(rows, "grade", "revenue").slice(0, 6), "grade");
}

function renderRank(id, data, filterKey) {
  const max = Math.max(1, ...data.map((d) => Math.abs(d.value)));
  $(id).innerHTML = data.length
    ? data
        .map((d) => `<div class="rank-row" data-${filterKey}="${escapeHtml(d.name)}" style="--w:${Math.max(8, (Math.abs(d.value) / max) * 100)}%"><span>${escapeHtml(d.name || "ไม่ระบุ")}</span><strong>${money(d.value)}</strong></div>`)
        .join("")
    : `<div class="empty-state">ไม่มีข้อมูล</div>`;
}

function renderFollow(rows) {
  const items = rows
    .filter((r) => r.payment === "ค้างรับ" || r.payment === "ค้างจ่าย" || r.status !== "ปกติ" || r.net < 0)
    .sort((a, b) => Math.abs(b.net) - Math.abs(a.net))
    .slice(0, 12);
  $("followList").innerHTML = items.length
    ? items
        .map((r) => `<article class="follow-item">
          <strong>${escapeHtml(r.product)} ${escapeHtml(r.name || r.grade || "")}</strong>
          <small>${escapeHtml(r.dateText)} | ${escapeHtml(r.payment)} | ${escapeHtml(r.status)} | สุทธิ ${money(r.net)}</small>
        </article>`)
        .join("")
    : `<div class="empty-state">ไม่มีรายการต้องติดตามในตัวกรองนี้</div>`;
}

function renderTable(rows) {
  const query = state.search.trim().toLowerCase();
  const tableRows = rows
    .filter((r) => !query || Object.values(r).join(" ").toLowerCase().includes(query))
    .sort((a, b) => (b.date?.getTime() || 0) - (a.date?.getTime() || 0))
    .slice(0, 60);
  $("rowsTable").innerHTML = tableRows
    .map((r) => `<tr>
      <td>${escapeHtml(r.dateText)}</td>
      <td>${escapeHtml(r.product)}</td>
      <td>${escapeHtml(r.type)}</td>
      <td>${escapeHtml(r.grade)}</td>
      <td>${escapeHtml(r.name)}</td>
      <td class="num">${money(r.revenue)}</td>
      <td class="num">${money(r.expense)}</td>
      <td class="num">${money(r.net)}</td>
      <td>${escapeHtml(r.payment)}</td>
    </tr>`)
    .join("");
}

function renderChips() {
  const labels = {
    year: "ปี",
    month: "เดือน",
    product: "สินค้า",
    grade: "เกรด",
    payment: "วิธี",
    status: "สถานะ",
  };
  const chips = Object.entries(state.filters)
    .filter(([, value]) => value !== "all")
    .map(([key, value]) => {
      const text = key === "month" ? String(value).split("|")[1] : value;
      return `<button class="chip" data-clear="${key}" type="button">${labels[key]}: ${escapeHtml(text)} ×</button>`;
    });
  $("filterChips").innerHTML = chips.join("");
}

function renderInsight(rows) {
  const t = totals(rows);
  const byProduct = groupBy(rows, "product", "revenue")[0];
  const byMonth = Array.from({ length: 12 }, (_, i) => ({ label: MONTHS[i], revenue: 0 }));
  rows.forEach((r) => {
    if (r.month) byMonth[r.month - 1].revenue += r.revenue;
  });
  const topMonth = byMonth.sort((a, b) => b.revenue - a.revenue)[0];
  const productShare = byProduct && t.revenue ? Math.round((byProduct.value / t.revenue) * 100) : 0;
  $("insightText").textContent = rows.length
    ? `ช่วงที่เลือกมีรายรับ ${money(t.revenue)} บาท สุทธิ ${money(t.net)} บาท รายรับหลักมาจาก${byProduct?.name || "ไม่ระบุ"} ${productShare}% เดือนที่รายรับสูงสุดคือ ${topMonth?.label || "-"} และมีเงินค้างรวม ${money(t.pendingIn + t.pendingOut)} บาท`
    : "ไม่มีข้อมูลในตัวกรองนี้";
}

function render() {
  const rows = filteredRows();
  renderKpis(rows);
  renderDonut(rows);
  renderFruitGradeDonuts(rows);
  renderMonthly(rows);
  renderPaymentCards(rows);
  renderRanks(rows);
  renderFollow(rows);
  renderTable(rows);
  renderChips();
  renderInsight(rows);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function bindEvents() {
  $("refreshBtn").addEventListener("click", loadData);
  $("clearBtn").addEventListener("click", () => {
    state.filters = defaultFilters();
    state.donutMetric = "revenue";
    state.search = "";
    $("searchBox").value = "";
    document.querySelectorAll(".segmented button").forEach((button) => {
      button.classList.toggle("active", button.dataset.metric === "revenue");
    });
    setupFilters();
    render();
  });
  [
    ["yearFilter", "year"],
    ["monthFilter", "month"],
    ["productFilter", "product"],
    ["gradeFilter", "grade"],
    ["paymentFilter", "payment"],
    ["statusFilter", "status"],
  ].forEach(([id, key]) => {
    $(id).addEventListener("change", (event) => {
      state.filters[key] = event.target.value;
      if (key === "product") {
        const availableGrades = gradeOptionsForCurrentProduct();
        if (state.filters.grade !== "all" && !availableGrades.includes(state.filters.grade)) {
          state.filters.grade = "all";
        }
        setOptions("gradeFilter", availableGrades, "ทุกสายพันธุ์/เกรด");
        $("gradeFilter").value = state.filters.grade;
      }
      render();
    });
  });
  $("searchBox").addEventListener("input", (event) => {
    state.search = event.target.value;
    renderTable(filteredRows());
  });
  document.querySelectorAll(".segmented button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".segmented button").forEach((b) => b.classList.remove("active"));
      button.classList.add("active");
      state.donutMetric = button.dataset.metric;
      renderDonut(filteredRows());
    });
  });
  document.addEventListener("click", (event) => {
    const product = event.target.closest("[data-product]")?.dataset.product;
    const grade = event.target.closest("[data-grade]")?.dataset.grade;
    const payment = event.target.closest("[data-payment]")?.dataset.payment;
    const month = event.target.closest("[data-month]")?.dataset.month;
    const clear = event.target.closest("[data-clear]")?.dataset.clear;
    if (product) state.filters.product = product;
    if (grade) state.filters.grade = grade;
    if (payment) state.filters.payment = payment;
    if (month) state.filters.month = month;
    if (clear) state.filters[clear] = clear === "status" ? "all" : "all";
    if (product || grade || payment || month || clear) {
      setupFilters();
      Object.entries(state.filters).forEach(([key, value]) => {
        const id = `${key}Filter`;
        if ($(id)) $(id).value = value;
      });
      render();
    }
  });
}

bindEvents();
loadData();
