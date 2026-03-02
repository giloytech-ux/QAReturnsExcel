const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

// ── Reference Data ──────────────────────────────────────────────────────────

const ROUTES = ["City", "Out of Town", "Bacolod", "Cebu", "Davao"];

const AREAS_BY_ROUTE = {
  City: ["DCP", "LGMV", "LJP", "LMP1", "LMP2"],
  "Out of Town": [
    "Aklan", "Altavas", "Antique", "Boracay", "Cabatuan",
    "Calinog", "Capiz", "Dumangas", "Estancia", "Guimaras",
    "Janiuay", "Leon", "Miag-ao", "Passi", "Roxas",
    "Sara", "Tapaz",
  ],
  Bacolod: ["Bacolod"],
  Cebu: ["Cebu"],
  Davao: ["Davao"],
};

const PRODUCTS_BY_CATEGORY = {
  Loaf: ["AL", "JW", "ML", "NSWB", "PL", "SL", "SP", "TB", "WB"],
  Assorted: [
    "AB", "BCR", "BT", "CB", "CR", "EB", "EC",
    "HB", "HR", "MB", "MBS", "ME", "MHB", "PDL", "PDR",
  ],
};

const CATEGORIES = ["Loaf", "Assorted"];
const USERS = ["Marcus Chen", "Aisha Patel", "Jake Morrison", "Sofia Reyes"];
const DESCRIPTIONS = [
  "Good Condition", "Molds", "Normal Smell", "Soft & Moist",
  "Deformed (Gupi)", "Damaged Wrapper", "Dry Crumb", "Others",
];

const SEED_PRODUCT_RETURNS = [
  { route: "City", area: "LMP1", category: "Loaf", product: "ML", qty: 5, batch: "B2026-0220", prodDate: new Date(2026, 1, 20), expiryDate: new Date(2026, 1, 27), dateReturned: new Date(2026, 1, 26), description: "Molds", otherDesc: "", notes: "Visible green mold on crust. Full batch affected.", inspector: "Marcus Chen" },
  { route: "Out of Town", area: "Boracay", category: "Loaf", product: "JW", qty: 12, batch: "B2026-0222", prodDate: new Date(2026, 1, 22), expiryDate: new Date(2026, 2, 1), dateReturned: new Date(2026, 1, 26), description: "Good Condition", otherDesc: "", notes: "Overstock return. Product in good condition.", inspector: "Aisha Patel" },
  { route: "Davao", area: "Davao", category: "Assorted", product: "BCR", qty: 8, batch: "", prodDate: new Date(2026, 1, 18), expiryDate: new Date(2026, 1, 25), dateReturned: new Date(2026, 1, 25), description: "Dry Crumb", otherDesc: "", notes: "Past expiry by 1 day at time of return.", inspector: "Marcus Chen" },
  { route: "City", area: "LGMV", category: "Loaf", product: "AL", qty: 3, batch: "B2026-0224", prodDate: new Date(2026, 1, 24), expiryDate: new Date(2026, 2, 3), dateReturned: new Date(2026, 1, 27), description: "Damaged Wrapper", otherDesc: "", notes: "Packaging torn. Loaves crushed during transit.", inspector: "Aisha Patel" },
  { route: "Out of Town", area: "Guimaras", category: "Loaf", product: "SP", qty: 6, batch: "B2026-0221", prodDate: new Date(2026, 1, 21), expiryDate: new Date(2026, 1, 28), dateReturned: new Date(2026, 1, 27), description: "Good Condition", otherDesc: "", notes: "Area overstocked. All units in sellable condition.", inspector: "Marcus Chen" },
];

const DATA_ROWS = 500;
const DATE_FMT = "DD MMM YYYY";

// ── Mobile-friendly sizing ──────────────────────────────────────────────────

const ROW_HEIGHT = 30;         // taller rows for touch
const FONT_SIZE = 12;          // larger base font
const HEADER_HEIGHT = 34;

// ── Styling ─────────────────────────────────────────────────────────────────

const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
const HEADER_FONT = { bold: true, color: { argb: "FFFFFFFF" }, size: FONT_SIZE };
const HEADER_ALIGNMENT = { vertical: "middle", horizontal: "center", wrapText: true };
const HEADER_BORDER = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "medium" }, right: { style: "thin" } };
const CELL_BORDER = {
  top: { style: "thin", color: { argb: "FFD0D0D0" } },
  left: { style: "thin", color: { argb: "FFD0D0D0" } },
  bottom: { style: "thin", color: { argb: "FFD0D0D0" } },
  right: { style: "thin", color: { argb: "FFD0D0D0" } },
};
const INPUT_BORDER = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
const INPUT_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF8E1" } };
const LABEL_FONT = { bold: true, size: 13 };
const INPUT_FONT = { size: 13 };

// ── Helpers ─────────────────────────────────────────────────────────────────

function rangeName(str) { return str.replace(/[^A-Za-z0-9_]/g, "_"); }

function styleHeaderRow(sheet, row) {
  row.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = HEADER_ALIGNMENT;
    cell.border = HEADER_BORDER;
  });
  row.height = HEADER_HEIGHT;
}

function styleDataRows(sheet, startRow, endRow, colCount) {
  for (let r = startRow; r <= endRow; r++) {
    const row = sheet.getRow(r);
    for (let c = 1; c <= colCount; c++) row.getCell(c).border = CELL_BORDER;
  }
}

function addListValidation(sheet, col, startRow, endRow, formula) {
  for (let r = startRow; r <= endRow; r++) {
    sheet.getCell(r, col).dataValidation = {
      type: "list", allowBlank: true, formulae: [formula],
      showErrorMessage: true, errorTitle: "Invalid", error: "Please select from the list.",
    };
  }
}

function addNumberValidation(sheet, col, startRow, endRow, min) {
  for (let r = startRow; r <= endRow; r++) {
    sheet.getCell(r, col).dataValidation = {
      type: "whole", operator: "greaterThanOrEqual", allowBlank: true, formulae: [min],
      showErrorMessage: true, errorTitle: "Invalid", error: `Please enter a whole number >= ${min}.`,
    };
  }
}

function columnLetter(num) {
  let s = "";
  while (num > 0) { num--; s = String.fromCharCode(65 + (num % 26)) + s; num = Math.floor(num / 26); }
  return s;
}

// ── Lookups Sheet ───────────────────────────────────────────────────────────

function buildLookupsSheet(wb) {
  const ws = wb.addWorksheet("Lookups", { state: "hidden" });

  ws.getCell("A1").value = "Routes";
  ROUTES.forEach((r, i) => ws.getCell(i + 2, 1).value = r);
  wb.definedNames.add(`Lookups!$A$2:$A$${ROUTES.length + 1}`, "Routes");

  ws.getCell("B1").value = "Categories";
  CATEGORIES.forEach((c, i) => ws.getCell(i + 2, 2).value = c);
  wb.definedNames.add(`Lookups!$B$2:$B$${CATEGORIES.length + 1}`, "Categories");

  ws.getCell("C1").value = "Descriptions";
  DESCRIPTIONS.forEach((d, i) => ws.getCell(i + 2, 3).value = d);
  wb.definedNames.add(`Lookups!$C$2:$C$${DESCRIPTIONS.length + 1}`, "Descriptions");

  ws.getCell("D1").value = "Users";
  USERS.forEach((u, i) => ws.getCell(i + 2, 4).value = u);
  wb.definedNames.add(`Lookups!$D$2:$D$${USERS.length + 1}`, "Users");

  let col = 5;
  for (const route of ROUTES) {
    const name = `Route_${rangeName(route)}`;
    ws.getCell(1, col).value = `${route} Areas`;
    const areas = AREAS_BY_ROUTE[route];
    areas.forEach((a, i) => ws.getCell(i + 2, col).value = a);
    const cl = columnLetter(col);
    wb.definedNames.add(`Lookups!$${cl}$2:$${cl}$${areas.length + 1}`, name);
    col++;
  }

  for (const cat of CATEGORIES) {
    const name = `Cat_${rangeName(cat)}`;
    ws.getCell(1, col).value = `${cat} Products`;
    const products = PRODUCTS_BY_CATEGORY[cat];
    products.forEach((p, i) => ws.getCell(i + 2, col).value = p);
    const cl = columnLetter(col);
    wb.definedNames.add(`Lookups!$${cl}$2:$${cl}$${products.length + 1}`, name);
    col++;
  }

  return ws;
}

// ── Data Entry Form Sheet (mobile-friendly vertical layout) ─────────────────

function buildDataEntrySheet(wb) {
  const ws = wb.addWorksheet("Data Entry");

  // Two-column layout: A = labels, B = inputs — wide for mobile
  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 36;

  // Title
  ws.mergeCells("A1:B1");
  ws.getCell("A1").value = "Product Returns Entry";
  ws.getCell("A1").font = { bold: true, size: 18, color: { argb: "FF1F4E79" } };
  ws.getCell("A1").alignment = { vertical: "middle" };
  ws.getRow(1).height = 42;

  // Subtitle
  ws.mergeCells("A2:B2");
  ws.getCell("A2").value = "Fill in the fields below, then run the Submit macro (Alt+F8).";
  ws.getCell("A2").font = { size: 11, italic: true, color: { argb: "FF666666" } };
  ws.getRow(2).height = 22;

  // Form fields — each row is a label + input pair
  const fields = [
    { label: "Route", row: 4, validation: { type: "list", formulae: ["Routes"] } },
    { label: "Area", row: 5, validation: { type: "list", formulae: ['INDIRECT("Route_"&SUBSTITUTE(B4," ","_"))'] } },
    { label: "Category", row: 6, validation: { type: "list", formulae: ["Categories"] } },
    { label: "Product", row: 7, validation: { type: "list", formulae: ['INDIRECT("Cat_"&SUBSTITUTE(B6," ","_"))'] } },
    { label: "Qty (Packs)", row: 8, validation: { type: "whole", operator: "greaterThanOrEqual", formulae: [1] } },
    { label: "Batch #", row: 9 },
    { label: "Production Date", row: 10, isDate: true },
    { label: "Expiry Date", row: 11, isDate: true },
    { label: "Date Returned", row: 12, isDate: true },
    { label: "Description", row: 13, validation: { type: "list", formulae: ["Descriptions"] } },
    { label: "Other Description", row: 14 },
    { label: "Notes", row: 15 },
    { label: "Inspector", row: 16, validation: { type: "list", formulae: ["Users"] } },
  ];

  for (const f of fields) {
    const r = f.row;
    ws.getRow(r).height = ROW_HEIGHT;

    // Label cell
    ws.getCell(r, 1).value = f.label;
    ws.getCell(r, 1).font = LABEL_FONT;
    ws.getCell(r, 1).alignment = { vertical: "middle" };

    // Input cell
    const inputCell = ws.getCell(r, 2);
    inputCell.border = INPUT_BORDER;
    inputCell.fill = INPUT_FILL;
    inputCell.font = INPUT_FONT;
    inputCell.alignment = { vertical: "middle" };

    if (f.isDate) {
      inputCell.numFmt = DATE_FMT;
    }

    if (f.validation) {
      inputCell.dataValidation = {
        ...f.validation,
        allowBlank: true,
        showErrorMessage: true,
        errorTitle: "Invalid",
        error: `Please enter a valid ${f.label}.`,
      };
    }
  }

  // Spacer row
  ws.getRow(17).height = 10;

  // Submit instruction row (styled like a button)
  ws.mergeCells("A18:B18");
  ws.getCell("A18").value = 'Run macro "SubmitProductReturn" to add this entry (Alt+F8)';
  ws.getCell("A18").font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  ws.getCell("A18").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E7D32" } };
  ws.getCell("A18").alignment = { vertical: "middle", horizontal: "center" };
  ws.getCell("A18").border = INPUT_BORDER;
  // Also style B18 since it's merged
  ws.getCell("B18").border = INPUT_BORDER;
  ws.getRow(18).height = 40;

  // Status display
  ws.mergeCells("A20:B20");
  ws.getCell("A20").value = "";
  ws.getCell("A20").font = { size: 12, bold: true, color: { argb: "FF2E7D32" } };
  ws.getCell("A20").alignment = { horizontal: "center" };

  return ws;
}

// ── Tray Entry Form Sheet (mobile-friendly vertical layout) ─────────────────

function buildTrayEntrySheet(wb) {
  const ws = wb.addWorksheet("Tray Entry");

  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 36;

  // Title
  ws.mergeCells("A1:B1");
  ws.getCell("A1").value = "Tray Returns Entry";
  ws.getCell("A1").font = { bold: true, size: 18, color: { argb: "FF1F4E79" } };
  ws.getCell("A1").alignment = { vertical: "middle" };
  ws.getRow(1).height = 42;

  // Subtitle
  ws.mergeCells("A2:B2");
  ws.getCell("A2").value = "Fill in the fields below, then run the Submit macro (Alt+F8).";
  ws.getCell("A2").font = { size: 11, italic: true, color: { argb: "FF666666" } };
  ws.getRow(2).height = 22;

  const fields = [
    { label: "Route", row: 4, validation: { type: "list", formulae: ["Routes"] } },
    { label: "Area", row: 5, validation: { type: "list", formulae: ['INDIRECT("Route_"&SUBSTITUTE(B4," ","_"))'] } },
    { label: "Tray Count", row: 6, validation: { type: "whole", operator: "greaterThanOrEqual", formulae: [0] } },
    { label: "Date Returned", row: 7, isDate: true },
    { label: "Inspector", row: 8, validation: { type: "list", formulae: ["Users"] } },
  ];

  for (const f of fields) {
    const r = f.row;
    ws.getRow(r).height = ROW_HEIGHT;

    ws.getCell(r, 1).value = f.label;
    ws.getCell(r, 1).font = LABEL_FONT;
    ws.getCell(r, 1).alignment = { vertical: "middle" };

    const inputCell = ws.getCell(r, 2);
    inputCell.border = INPUT_BORDER;
    inputCell.fill = INPUT_FILL;
    inputCell.font = INPUT_FONT;
    inputCell.alignment = { vertical: "middle" };

    if (f.isDate) inputCell.numFmt = DATE_FMT;

    if (f.validation) {
      inputCell.dataValidation = {
        ...f.validation,
        allowBlank: true,
        showErrorMessage: true,
        errorTitle: "Invalid",
        error: `Please enter a valid ${f.label}.`,
      };
    }
  }

  // Spacer row
  ws.getRow(9).height = 10;

  // Submit instruction row
  ws.mergeCells("A10:B10");
  ws.getCell("A10").value = 'Run macro "SubmitTrayReturn" to add this entry (Alt+F8)';
  ws.getCell("A10").font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  ws.getCell("A10").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2E7D32" } };
  ws.getCell("A10").alignment = { vertical: "middle", horizontal: "center" };
  ws.getCell("A10").border = INPUT_BORDER;
  ws.getCell("B10").border = INPUT_BORDER;
  ws.getRow(10).height = 40;

  // Status display
  ws.mergeCells("A12:B12");
  ws.getCell("A12").value = "";
  ws.getCell("A12").font = { size: 12, bold: true, color: { argb: "FF2E7D32" } };
  ws.getCell("A12").alignment = { horizontal: "center" };

  return ws;
}

// ── Product Returns Sheet ───────────────────────────────────────────────────

function buildProductReturnsSheet(wb) {
  const ws = wb.addWorksheet("Product Returns");

  const headers = [
    { header: "Record #", key: "record", width: 10 },
    { header: "Date Inspected", key: "dateInspected", width: 16 },
    { header: "Route", key: "route", width: 15 },
    { header: "Area", key: "area", width: 15 },
    { header: "Category", key: "category", width: 13 },
    { header: "Product", key: "product", width: 11 },
    { header: "Qty (Packs)", key: "qty", width: 13 },
    { header: "Batch #", key: "batch", width: 15 },
    { header: "Production Date", key: "prodDate", width: 17 },
    { header: "Expiry Date", key: "expiryDate", width: 17 },
    { header: "Date Returned", key: "dateReturned", width: 17 },
    { header: "Description", key: "description", width: 20 },
    { header: "Other Description", key: "otherDesc", width: 20 },
    { header: "Notes", key: "notes", width: 26 },
    { header: "Inspector", key: "inspector", width: 20 },
  ];

  ws.columns = headers;
  const headerRow = ws.getRow(1);
  styleHeaderRow(ws, headerRow);

  const firstDataRow = 2;
  const lastDataRow = firstDataRow + DATA_ROWS - 1;

  // Insert sample data
  SEED_PRODUCT_RETURNS.forEach((rec, i) => {
    const r = firstDataRow + i;
    ws.getCell(r, 1).value = { formula: `IF(C${r}<>"",ROW()-1,"")` };
    ws.getCell(r, 1).alignment = { horizontal: "center" };
    ws.getCell(r, 2).value = { formula: `IF(C${r}<>"",TODAY(),"")` };
    ws.getCell(r, 2).numFmt = DATE_FMT;
    ws.getCell(r, 3).value = rec.route;
    ws.getCell(r, 4).value = rec.area;
    ws.getCell(r, 5).value = rec.category;
    ws.getCell(r, 6).value = rec.product;
    ws.getCell(r, 7).value = rec.qty;
    ws.getCell(r, 8).value = rec.batch;
    ws.getCell(r, 9).value = rec.prodDate; ws.getCell(r, 9).numFmt = DATE_FMT;
    ws.getCell(r, 10).value = rec.expiryDate; ws.getCell(r, 10).numFmt = DATE_FMT;
    ws.getCell(r, 11).value = rec.dateReturned; ws.getCell(r, 11).numFmt = DATE_FMT;
    ws.getCell(r, 12).value = rec.description;
    ws.getCell(r, 13).value = rec.otherDesc;
    ws.getCell(r, 14).value = rec.notes;
    ws.getCell(r, 15).value = rec.inspector;
  });

  // Empty rows with formulas
  const sampleEnd = firstDataRow + SEED_PRODUCT_RETURNS.length;
  for (let r = sampleEnd; r <= lastDataRow; r++) {
    ws.getCell(r, 1).value = { formula: `IF(C${r}<>"",ROW()-1,"")` };
    ws.getCell(r, 1).alignment = { horizontal: "center" };
    ws.getCell(r, 2).value = { formula: `IF(C${r}<>"",TODAY(),"")` };
    ws.getCell(r, 2).numFmt = DATE_FMT;
  }

  // Date formatting
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    ws.getCell(r, 9).numFmt = DATE_FMT;
    ws.getCell(r, 10).numFmt = DATE_FMT;
    ws.getCell(r, 11).numFmt = DATE_FMT;
    ws.getRow(r).height = ROW_HEIGHT;
  }

  // Dropdowns
  addListValidation(ws, 3, firstDataRow, lastDataRow, "Routes");
  addListValidation(ws, 5, firstDataRow, lastDataRow, "Categories");
  addListValidation(ws, 12, firstDataRow, lastDataRow, "Descriptions");
  addListValidation(ws, 15, firstDataRow, lastDataRow, "Users");

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    ws.getCell(r, 4).dataValidation = { type: "list", allowBlank: true, formulae: [`INDIRECT("Route_"&SUBSTITUTE(C${r}," ","_"))`], showErrorMessage: true, errorTitle: "Invalid", error: "Select a route first." };
    ws.getCell(r, 6).dataValidation = { type: "list", allowBlank: true, formulae: [`INDIRECT("Cat_"&SUBSTITUTE(E${r}," ","_"))`], showErrorMessage: true, errorTitle: "Invalid", error: "Select a category first." };
  }

  addNumberValidation(ws, 7, firstDataRow, lastDataRow, 1);

  // Conditional formatting: Expiry Date red if past today
  ws.addConditionalFormatting({
    ref: `J${firstDataRow}:J${lastDataRow}`,
    rules: [{
      type: "expression",
      formulae: [`AND(J${firstDataRow}<>"",J${firstDataRow}<TODAY())`],
      style: {
        font: { color: { argb: "FF9C0006" } },
        fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FFFFC7CE" } },
      },
      priority: 1,
    }],
  });

  styleDataRows(ws, firstDataRow, lastDataRow, headers.length);
  ws.views = [{ state: "frozen", ySplit: 1 }];

  // Alternate row shading
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    if (r % 2 === 0) {
      for (let c = 1; c <= headers.length; c++) {
        const cell = ws.getCell(r, c);
        if (!cell.fill || !cell.fill.fgColor) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F7FB" } };
        }
      }
    }
  }

  return ws;
}

// ── Tray Returns Sheet ──────────────────────────────────────────────────────

function buildTrayReturnsSheet(wb) {
  const ws = wb.addWorksheet("Tray Returns");

  const headers = [
    { header: "Record #", key: "record", width: 10 },
    { header: "Date Inspected", key: "dateInspected", width: 16 },
    { header: "Route", key: "route", width: 15 },
    { header: "Area", key: "area", width: 15 },
    { header: "Tray Count", key: "trayCount", width: 13 },
    { header: "Date Returned", key: "dateReturned", width: 17 },
    { header: "Inspector", key: "inspector", width: 20 },
  ];

  ws.columns = headers;
  const headerRow = ws.getRow(1);
  styleHeaderRow(ws, headerRow);

  const firstDataRow = 2;
  const lastDataRow = firstDataRow + DATA_ROWS - 1;

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    ws.getCell(r, 1).value = { formula: `IF(C${r}<>"",ROW()-1,"")` };
    ws.getCell(r, 1).alignment = { horizontal: "center" };
    ws.getCell(r, 2).value = { formula: `IF(C${r}<>"",TODAY(),"")` };
    ws.getCell(r, 2).numFmt = DATE_FMT;
    ws.getCell(r, 6).numFmt = DATE_FMT;
    ws.getRow(r).height = ROW_HEIGHT;
  }

  addListValidation(ws, 3, firstDataRow, lastDataRow, "Routes");
  addListValidation(ws, 7, firstDataRow, lastDataRow, "Users");
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    ws.getCell(r, 4).dataValidation = { type: "list", allowBlank: true, formulae: [`INDIRECT("Route_"&SUBSTITUTE(C${r}," ","_"))`], showErrorMessage: true, errorTitle: "Invalid", error: "Select a route first." };
  }

  addNumberValidation(ws, 5, firstDataRow, lastDataRow, 0);
  styleDataRows(ws, firstDataRow, lastDataRow, headers.length);
  ws.views = [{ state: "frozen", ySplit: 1 }];

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    if (r % 2 === 0) {
      for (let c = 1; c <= headers.length; c++) {
        ws.getCell(r, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F7FB" } };
      }
    }
  }

  return ws;
}

// ── Summary Sheet ───────────────────────────────────────────────────────────

function buildSummarySheet(wb) {
  const ws = wb.addWorksheet("Summary");

  const prRange = `'Product Returns'!C2:C${DATA_ROWS + 1}`;
  const qtyRange = `'Product Returns'!G2:G${DATA_ROWS + 1}`;
  const descRange = `'Product Returns'!L2:L${DATA_ROWS + 1}`;

  ws.getCell("A1").value = "QA Returns Summary";
  ws.getCell("A1").font = { bold: true, size: 18, color: { argb: "FF1F4E79" } };
  ws.mergeCells("A1:D1");
  ws.getRow(1).height = 36;

  const metricsStart = 3;
  const metricLabels = ["Total Returns", "Total Quantity (Packs)", "Issue Records"];
  const metricFormulas = [
    `COUNTA(${prRange})-COUNTBLANK(${prRange})`,
    `SUMPRODUCT((${qtyRange}<>"")*1*(${qtyRange}))`,
    `COUNTIF(${descRange},"<>"&"")-COUNTIF(${descRange},"Good Condition")`,
  ];

  ws.getCell("A3").value = "Overview";
  ws.getCell("A3").font = { bold: true, size: 14, color: { argb: "FF1F4E79" } };

  metricLabels.forEach((label, i) => {
    const row = metricsStart + 1 + i;
    ws.getCell(row, 1).value = label;
    ws.getCell(row, 1).font = { bold: true, size: FONT_SIZE };
    ws.getCell(row, 2).value = { formula: metricFormulas[i] };
    ws.getCell(row, 2).numFmt = "#,##0";
    ws.getCell(row, 2).font = { bold: true, size: 14 };
    ws.getRow(row).height = ROW_HEIGHT;
  });

  const routeStart = metricsStart + metricLabels.length + 3;
  ws.getCell(routeStart, 1).value = "Returns by Route";
  ws.getCell(routeStart, 1).font = { bold: true, size: 14, color: { argb: "FF1F4E79" } };

  const routeHeaderRow = routeStart + 1;
  ws.getCell(routeHeaderRow, 1).value = "Route";
  ws.getCell(routeHeaderRow, 2).value = "Count";
  ws.getCell(routeHeaderRow, 3).value = "Qty (Packs)";
  const rhr = ws.getRow(routeHeaderRow);
  rhr.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: FONT_SIZE };
    cell.fill = HEADER_FILL; cell.border = HEADER_BORDER;
    cell.alignment = { horizontal: "center" };
  });
  rhr.height = ROW_HEIGHT;

  ROUTES.forEach((route, i) => {
    const row = routeHeaderRow + 1 + i;
    ws.getCell(row, 1).value = route;
    ws.getCell(row, 1).font = { bold: true, size: FONT_SIZE };
    ws.getCell(row, 2).value = { formula: `COUNTIF(${prRange},"${route}")` };
    ws.getCell(row, 2).numFmt = "#,##0";
    ws.getCell(row, 3).value = { formula: `SUMPRODUCT(('Product Returns'!C2:C${DATA_ROWS + 1}="${route}")*('Product Returns'!G2:G${DATA_ROWS + 1}))` };
    ws.getCell(row, 3).numFmt = "#,##0";
    for (let c = 1; c <= 3; c++) ws.getCell(row, c).border = CELL_BORDER;
    ws.getRow(row).height = ROW_HEIGHT;
  });

  const descStart = routeHeaderRow + ROUTES.length + 3;
  ws.getCell(descStart, 1).value = "Returns by Description";
  ws.getCell(descStart, 1).font = { bold: true, size: 14, color: { argb: "FF1F4E79" } };

  const descHeaderRow = descStart + 1;
  ws.getCell(descHeaderRow, 1).value = "Description";
  ws.getCell(descHeaderRow, 2).value = "Count";
  const dhr = ws.getRow(descHeaderRow);
  dhr.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: FONT_SIZE };
    cell.fill = HEADER_FILL; cell.border = HEADER_BORDER;
    cell.alignment = { horizontal: "center" };
  });
  dhr.height = ROW_HEIGHT;

  DESCRIPTIONS.forEach((desc, i) => {
    const row = descHeaderRow + 1 + i;
    ws.getCell(row, 1).value = desc;
    ws.getCell(row, 1).font = { bold: true, size: FONT_SIZE };
    ws.getCell(row, 2).value = { formula: `COUNTIF(${descRange},"${desc}")` };
    ws.getCell(row, 2).numFmt = "#,##0";
    for (let c = 1; c <= 2; c++) ws.getCell(row, c).border = CELL_BORDER;
    ws.getRow(row).height = ROW_HEIGHT;
  });

  ws.getColumn(1).width = 26;
  ws.getColumn(2).width = 16;
  ws.getColumn(3).width = 16;
  ws.getColumn(4).width = 16;
  ws.views = [{ state: "frozen", ySplit: 1 }];

  return ws;
}

// ── Report Generator Sheet ──────────────────────────────────────────────────

function buildReportGeneratorSheet(wb) {
  const ws = wb.addWorksheet("Report Generator");

  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 28;
  ws.getColumn(3).width = 50;

  ws.mergeCells("A1:C1");
  ws.getCell("A1").value = "Sensory Evaluation Report";
  ws.getCell("A1").font = { bold: true, size: 18, color: { argb: "FF1F4E79" } };
  ws.getRow(1).height = 42;

  // Route
  ws.getCell("A3").value = "Route:";
  ws.getCell("A3").font = LABEL_FONT;
  ws.getCell("B3").value = "City";
  ws.getCell("B3").border = INPUT_BORDER;
  ws.getCell("B3").fill = INPUT_FILL;
  ws.getCell("B3").font = INPUT_FONT;
  ws.getCell("B3").dataValidation = { type: "list", allowBlank: false, formulae: ["Routes"], showErrorMessage: true, errorTitle: "Invalid", error: "Select a route." };
  ws.getRow(3).height = ROW_HEIGHT;

  // Start Date
  ws.getCell("A5").value = "Start Date:";
  ws.getCell("A5").font = LABEL_FONT;
  ws.getCell("B5").value = new Date();
  ws.getCell("B5").numFmt = DATE_FMT;
  ws.getCell("B5").border = INPUT_BORDER;
  ws.getCell("B5").fill = INPUT_FILL;
  ws.getCell("B5").font = INPUT_FONT;
  ws.getRow(5).height = ROW_HEIGHT;

  // Start Time
  ws.getCell("A6").value = "Start Time:";
  ws.getCell("A6").font = LABEL_FONT;
  ws.getCell("B6").value = 0;
  ws.getCell("B6").numFmt = "HH:MM";
  ws.getCell("B6").border = INPUT_BORDER;
  ws.getCell("B6").fill = INPUT_FILL;
  ws.getCell("B6").font = INPUT_FONT;
  ws.getRow(6).height = ROW_HEIGHT;

  // End Date
  ws.getCell("A8").value = "End Date:";
  ws.getCell("A8").font = LABEL_FONT;
  ws.getCell("B8").value = new Date();
  ws.getCell("B8").numFmt = DATE_FMT;
  ws.getCell("B8").border = INPUT_BORDER;
  ws.getCell("B8").fill = INPUT_FILL;
  ws.getCell("B8").font = INPUT_FONT;
  ws.getRow(8).height = ROW_HEIGHT;

  // End Time
  ws.getCell("A9").value = "End Time:";
  ws.getCell("A9").font = LABEL_FONT;
  ws.getCell("B9").value = 0.9999;
  ws.getCell("B9").numFmt = "HH:MM";
  ws.getCell("B9").border = INPUT_BORDER;
  ws.getCell("B9").fill = INPUT_FILL;
  ws.getCell("B9").font = INPUT_FONT;
  ws.getRow(9).height = ROW_HEIGHT;

  // Generate instruction
  ws.mergeCells("A11:C11");
  ws.getCell("A11").value = 'Run macro "GenerateSensoryReport" to generate (Alt+F8)';
  ws.getCell("A11").font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  ws.getCell("A11").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
  ws.getCell("A11").alignment = { vertical: "middle", horizontal: "center" };
  ws.getCell("A11").border = INPUT_BORDER;
  ws.getRow(11).height = 40;

  // Instructions
  ws.getCell("A13").value = "Setup (one-time):";
  ws.getCell("A13").font = { bold: true, size: 13, color: { argb: "FF1F4E79" } };

  const steps = [
    '1. Go to the "VBA Code" sheet and copy ALL the code',
    "2. Press Alt+F11 to open the VBA Editor",
    "3. Go to Insert > Module, then paste the code",
    "4. Close the VBA Editor",
    '5. Save as .xlsm so you only do this once',
  ];

  steps.forEach((s, i) => {
    ws.getCell(14 + i, 1).value = s;
    ws.getCell(14 + i, 1).font = { size: 11 };
    ws.mergeCells(14 + i, 1, 14 + i, 3);
  });

  return ws;
}

// ── VBA Code Sheet ──────────────────────────────────────────────────────────

function buildVbaCodeSheet(wb) {
  const ws = wb.addWorksheet("VBA Code");
  ws.getColumn(1).width = 130;

  ws.getCell("A1").value = 'Copy ALL the code below → Alt+F11 → Insert → Module → Paste → Close → Save as .xlsm';
  ws.getCell("A1").font = { bold: true, size: 12, color: { argb: "FF1F4E79" } };
  ws.getRow(1).height = 24;

  const vbaCode = getVbaCode();
  const lines = vbaCode.split("\n");
  lines.forEach((line, i) => {
    ws.getCell(i + 3, 1).value = line;
    ws.getCell(i + 3, 1).font = { name: "Consolas", size: 10 };
  });

  ws.views = [{ state: "frozen", ySplit: 1 }];
  return ws;
}

function getVbaCode() {
  return `' ============================================================
' QA Returns — VBA Macros
' Contains three macros:
'   1) SubmitProductReturn  — adds Data Entry form to Product Returns
'   2) SubmitTrayReturn     — adds Tray Entry form to Tray Returns
'   3) GenerateSensoryReport — generates Sensory Evaluation report
' ============================================================

Sub SubmitProductReturn()
    Dim wsForm As Worksheet
    Dim wsPR As Worksheet

    Set wsForm = ThisWorkbook.Sheets("Data Entry")
    Set wsPR = ThisWorkbook.Sheets("Product Returns")

    ' ── Read form values ──
    Dim fRoute As String:       fRoute = Trim(CStr(wsForm.Range("B4").Value))
    Dim fArea As String:        fArea = Trim(CStr(wsForm.Range("B5").Value))
    Dim fCategory As String:    fCategory = Trim(CStr(wsForm.Range("B6").Value))
    Dim fProduct As String:     fProduct = Trim(CStr(wsForm.Range("B7").Value))
    Dim fQty As Variant:        fQty = wsForm.Range("B8").Value
    Dim fBatch As String:       fBatch = Trim(CStr(wsForm.Range("B9").Value))
    Dim fProdDate As Variant:   fProdDate = wsForm.Range("B10").Value
    Dim fExpiry As Variant:     fExpiry = wsForm.Range("B11").Value
    Dim fReturned As Variant:   fReturned = wsForm.Range("B12").Value
    Dim fDesc As String:        fDesc = Trim(CStr(wsForm.Range("B13").Value))
    Dim fOtherDesc As String:   fOtherDesc = Trim(CStr(wsForm.Range("B14").Value))
    Dim fNotes As String:       fNotes = Trim(CStr(wsForm.Range("B15").Value))
    Dim fInspector As String:   fInspector = Trim(CStr(wsForm.Range("B16").Value))

    ' ── Validate required fields ──
    If fRoute = "" Then
        MsgBox "Route is required.", vbExclamation
        wsForm.Range("B4").Select
        Exit Sub
    End If
    If fArea = "" Then
        MsgBox "Area is required.", vbExclamation
        wsForm.Range("B5").Select
        Exit Sub
    End If
    If fCategory = "" Then
        MsgBox "Category is required.", vbExclamation
        wsForm.Range("B6").Select
        Exit Sub
    End If
    If fProduct = "" Then
        MsgBox "Product is required.", vbExclamation
        wsForm.Range("B7").Select
        Exit Sub
    End If
    If Not IsNumeric(fQty) Or CLng(fQty) < 1 Then
        MsgBox "Qty (Packs) must be a number >= 1.", vbExclamation
        wsForm.Range("B8").Select
        Exit Sub
    End If
    If fDesc = "" Then
        MsgBox "Description is required.", vbExclamation
        wsForm.Range("B13").Select
        Exit Sub
    End If
    If fInspector = "" Then
        MsgBox "Inspector is required.", vbExclamation
        wsForm.Range("B16").Select
        Exit Sub
    End If

    ' ── Find next empty row in Product Returns (check col C = Route) ──
    Dim nextRow As Long
    nextRow = 2
    Dim i As Long
    For i = 2 To 501
        If wsPR.Cells(i, 3).Value <> "" Then
            nextRow = i + 1
        End If
    Next i

    If nextRow > 501 Then
        MsgBox "Product Returns sheet is full (500 rows).", vbExclamation
        Exit Sub
    End If

    ' ── Write data to Product Returns ──
    ' Col A (Record #) and Col B (Date Inspected) have formulas already
    wsPR.Cells(nextRow, 3).Value = fRoute
    wsPR.Cells(nextRow, 4).Value = fArea
    wsPR.Cells(nextRow, 5).Value = fCategory
    wsPR.Cells(nextRow, 6).Value = fProduct
    wsPR.Cells(nextRow, 7).Value = CLng(fQty)
    wsPR.Cells(nextRow, 8).Value = fBatch
    If IsDate(fProdDate) Then wsPR.Cells(nextRow, 9).Value = CDate(fProdDate)
    If IsDate(fExpiry) Then wsPR.Cells(nextRow, 10).Value = CDate(fExpiry)
    If IsDate(fReturned) Then wsPR.Cells(nextRow, 11).Value = CDate(fReturned)
    wsPR.Cells(nextRow, 12).Value = fDesc
    wsPR.Cells(nextRow, 13).Value = fOtherDesc
    wsPR.Cells(nextRow, 14).Value = fNotes
    wsPR.Cells(nextRow, 15).Value = fInspector

    ' ── Clear form ──
    wsForm.Range("B4").Value = ""
    wsForm.Range("B5").Value = ""
    wsForm.Range("B6").Value = ""
    wsForm.Range("B7").Value = ""
    wsForm.Range("B8").Value = ""
    wsForm.Range("B9").Value = ""
    wsForm.Range("B10").Value = ""
    wsForm.Range("B11").Value = ""
    wsForm.Range("B12").Value = ""
    wsForm.Range("B13").Value = ""
    wsForm.Range("B14").Value = ""
    wsForm.Range("B15").Value = ""
    wsForm.Range("B16").Value = ""

    ' ── Show confirmation ──
    wsForm.Range("A20").Value = "Record #" & (nextRow - 1) & " added successfully!"
    wsForm.Range("B4").Select

    MsgBox "Record added as row " & nextRow & " in Product Returns!", vbInformation
End Sub

' ============================================================

Sub SubmitTrayReturn()
    Dim wsForm As Worksheet
    Dim wsTR As Worksheet

    Set wsForm = ThisWorkbook.Sheets("Tray Entry")
    Set wsTR = ThisWorkbook.Sheets("Tray Returns")

    ' ── Read form values ──
    Dim fRoute As String:       fRoute = Trim(CStr(wsForm.Range("B4").Value))
    Dim fArea As String:        fArea = Trim(CStr(wsForm.Range("B5").Value))
    Dim fTrayCount As Variant:  fTrayCount = wsForm.Range("B6").Value
    Dim fReturned As Variant:   fReturned = wsForm.Range("B7").Value
    Dim fInspector As String:   fInspector = Trim(CStr(wsForm.Range("B8").Value))

    ' ── Validate required fields ──
    If fRoute = "" Then
        MsgBox "Route is required.", vbExclamation
        wsForm.Range("B4").Select
        Exit Sub
    End If
    If fArea = "" Then
        MsgBox "Area is required.", vbExclamation
        wsForm.Range("B5").Select
        Exit Sub
    End If
    If Not IsNumeric(fTrayCount) Or CLng(fTrayCount) < 0 Then
        MsgBox "Tray Count must be a number >= 0.", vbExclamation
        wsForm.Range("B6").Select
        Exit Sub
    End If
    If fInspector = "" Then
        MsgBox "Inspector is required.", vbExclamation
        wsForm.Range("B8").Select
        Exit Sub
    End If

    ' ── Find next empty row in Tray Returns (check col C = Route) ──
    Dim nextRow As Long
    nextRow = 2
    Dim i As Long
    For i = 2 To 501
        If wsTR.Cells(i, 3).Value <> "" Then
            nextRow = i + 1
        End If
    Next i

    If nextRow > 501 Then
        MsgBox "Tray Returns sheet is full (500 rows).", vbExclamation
        Exit Sub
    End If

    ' ── Write data to Tray Returns ──
    ' Col A (Record #) and Col B (Date Inspected) have formulas already
    wsTR.Cells(nextRow, 3).Value = fRoute
    wsTR.Cells(nextRow, 4).Value = fArea
    wsTR.Cells(nextRow, 5).Value = CLng(fTrayCount)
    If IsDate(fReturned) Then wsTR.Cells(nextRow, 6).Value = CDate(fReturned)
    wsTR.Cells(nextRow, 7).Value = fInspector

    ' ── Clear form ──
    wsForm.Range("B4").Value = ""
    wsForm.Range("B5").Value = ""
    wsForm.Range("B6").Value = ""
    wsForm.Range("B7").Value = ""
    wsForm.Range("B8").Value = ""

    ' ── Show confirmation ──
    wsForm.Range("A12").Value = "Record #" & (nextRow - 1) & " added successfully!"
    wsForm.Range("B4").Select

    MsgBox "Tray record added as row " & nextRow & " in Tray Returns!", vbInformation
End Sub

' ============================================================

Sub GenerateSensoryReport()
    Application.ScreenUpdating = False

    Dim wsGen As Worksheet
    Dim wsPR As Worksheet
    Dim wsTR As Worksheet
    Dim wsOut As Worksheet

    Set wsGen = ThisWorkbook.Sheets("Report Generator")
    Set wsPR = ThisWorkbook.Sheets("Product Returns")
    Set wsTR = ThisWorkbook.Sheets("Tray Returns")

    ' ── Read inputs ──
    Dim selRoute As String
    selRoute = Trim(CStr(wsGen.Range("B3").Value))
    If selRoute = "" Then
        MsgBox "Please select a Route on the Report Generator sheet.", vbExclamation
        Exit Sub
    End If

    Dim startDT As Date
    Dim endDT As Date
    startDT = CDate(wsGen.Range("B5").Value) + CDbl(wsGen.Range("B6").Value)
    endDT = CDate(wsGen.Range("B8").Value) + CDbl(wsGen.Range("B9").Value)

    If endDT < startDT Then
        MsgBox "End date/time must be after start date/time.", vbExclamation
        Exit Sub
    End If

    ' ── Delete existing sheet ──
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Sensory Evaluation").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsOut.Name = "Sensory Evaluation"

    ' ── Find last data rows ──
    Dim lastRowPR As Long
    Dim lastRowTR As Long
    lastRowPR = 1
    Dim i As Long
    For i = 2 To 501
        If wsPR.Cells(i, 3).Value <> "" Then lastRowPR = i
    Next i
    lastRowTR = 1
    For i = 2 To 501
        If wsTR.Cells(i, 3).Value <> "" Then lastRowTR = i
    Next i

    ' ── Filter Product Returns by route & date range ──
    Dim filteredRows() As Long
    Dim filtCount As Long
    filtCount = 0

    For i = 2 To lastRowPR
        If CStr(wsPR.Cells(i, 3).Value) = selRoute Then
            Dim inspDate As Date
            If IsDate(wsPR.Cells(i, 2).Value) Then
                inspDate = CDate(wsPR.Cells(i, 2).Value)
                If inspDate >= Int(startDT) And inspDate <= Int(endDT) Then
                    filtCount = filtCount + 1
                End If
            End If
        End If
    Next i

    If filtCount = 0 Then
        MsgBox "No records found for """ & selRoute & """ in the selected date range.", vbInformation
        Application.DisplayAlerts = False
        wsOut.Delete
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ReDim filteredRows(1 To filtCount)
    Dim fi As Long
    fi = 0
    For i = 2 To lastRowPR
        If CStr(wsPR.Cells(i, 3).Value) = selRoute Then
            If IsDate(wsPR.Cells(i, 2).Value) Then
                inspDate = CDate(wsPR.Cells(i, 2).Value)
                If inspDate >= Int(startDT) And inspDate <= Int(endDT) Then
                    fi = fi + 1
                    filteredRows(fi) = i
                End If
            End If
        End If
    Next i

    ' ── Build tray count lookup ──
    Dim trayDict As Object
    Set trayDict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRowTR
        If CStr(wsTR.Cells(i, 3).Value) = selRoute Then
            Dim trayInspDate As Date
            If IsDate(wsTR.Cells(i, 2).Value) Then
                trayInspDate = CDate(wsTR.Cells(i, 2).Value)
                If trayInspDate >= Int(startDT) And trayInspDate <= Int(endDT) Then
                    Dim tArea As String
                    tArea = CStr(wsTR.Cells(i, 4).Value)
                    If tArea <> "" Then
                        Dim tv As Long
                        tv = 0
                        If IsNumeric(wsTR.Cells(i, 5).Value) Then tv = CLng(wsTR.Cells(i, 5).Value)
                        If trayDict.Exists(tArea) Then
                            trayDict(tArea) = trayDict(tArea) + tv
                        Else
                            trayDict.Add tArea, tv
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ' ── Collect & sort unique areas ──
    Dim areaDict As Object
    Set areaDict = CreateObject("Scripting.Dictionary")
    For fi = 1 To filtCount
        Dim aName As String
        aName = CStr(wsPR.Cells(filteredRows(fi), 4).Value)
        If aName <> "" And Not areaDict.Exists(aName) Then areaDict.Add aName, True
    Next fi

    Dim areaKeys() As Variant
    areaKeys = areaDict.Keys
    Dim j As Long
    Dim tmpStr As String
    For i = LBound(areaKeys) To UBound(areaKeys) - 1
        For j = i + 1 To UBound(areaKeys)
            If areaKeys(i) > areaKeys(j) Then
                tmpStr = areaKeys(i)
                areaKeys(i) = areaKeys(j)
                areaKeys(j) = tmpStr
            End If
        Next j
    Next i

    ' ── Page setup ──
    wsOut.PageSetup.Orientation = xlPortrait
    wsOut.PageSetup.PaperSize = xlPaperA4
    wsOut.PageSetup.LeftMargin = Application.InchesToPoints(0.79)
    wsOut.PageSetup.RightMargin = Application.InchesToPoints(0.79)
    wsOut.PageSetup.TopMargin = Application.InchesToPoints(0.79)
    wsOut.PageSetup.BottomMargin = Application.InchesToPoints(0.79)
    wsOut.Columns(1).ColumnWidth = 90

    Dim outRow As Long
    outRow = 1

    ' ── Header ──
    Dim startStr As String
    Dim endStr As String
    startStr = UCase(Format(startDT, "DD MMM YYYY")) & " " & Format(startDT, "h:mm AM/PM")
    endStr = UCase(Format(endDT, "DD MMM YYYY")) & " " & Format(endDT, "h:mm AM/PM")

    wsOut.Cells(outRow, 1).Value = "Cutoff: " & startStr & " - " & endStr
    wsOut.Cells(outRow, 1).Font.Size = 11
    wsOut.Cells(outRow, 1).Font.Color = RGB(148, 163, 184)
    wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
    outRow = outRow + 1

    wsOut.Cells(outRow, 1).Value = "Generated: " & UCase(Format(Now, "DD MMM YYYY, h:mm AM/PM"))
    wsOut.Cells(outRow, 1).Font.Size = 11
    wsOut.Cells(outRow, 1).Font.Color = RGB(148, 163, 184)
    wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
    outRow = outRow + 1

    With wsOut.Cells(outRow, 1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(203, 213, 225)
        .Weight = xlThin
    End With
    outRow = outRow + 2

    ' ── Process each area ──
    Dim aIdx As Long
    For aIdx = LBound(areaKeys) To UBound(areaKeys)
        Dim curArea As String
        curArea = CStr(areaKeys(aIdx))

        Dim areaTrayCount As Long
        areaTrayCount = 0
        If trayDict.Exists(curArea) Then areaTrayCount = trayDict(curArea)

        If aIdx > LBound(areaKeys) Then
            wsOut.HPageBreaks.Add Before:=wsOut.Cells(outRow, 1)
        End If

        wsOut.Cells(outRow, 1).Value = "Sensory Evaluation of Returns from " & UCase(selRoute) & "-" & UCase(curArea) & " (" & areaTrayCount & " Tray/s)"
        wsOut.Cells(outRow, 1).Font.Bold = True
        wsOut.Cells(outRow, 1).Font.Size = 14
        wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
        wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
        outRow = outRow + 1

        ' Inspectors
        Dim inspDict As Object
        Set inspDict = CreateObject("Scripting.Dictionary")
        For fi = 1 To filtCount
            If CStr(wsPR.Cells(filteredRows(fi), 4).Value) = curArea Then
                Dim iName As String
                iName = CStr(wsPR.Cells(filteredRows(fi), 15).Value)
                If iName <> "" And Not inspDict.Exists(iName) Then inspDict.Add iName, True
            End If
        Next fi
        For i = 2 To lastRowTR
            If CStr(wsTR.Cells(i, 3).Value) = selRoute And CStr(wsTR.Cells(i, 4).Value) = curArea Then
                If IsDate(wsTR.Cells(i, 2).Value) Then
                    trayInspDate = CDate(wsTR.Cells(i, 2).Value)
                    If trayInspDate >= Int(startDT) And trayInspDate <= Int(endDT) Then
                        Dim tiName As String
                        tiName = CStr(wsTR.Cells(i, 7).Value)
                        If tiName <> "" And Not inspDict.Exists(tiName) Then inspDict.Add tiName, True
                    End If
                End If
            End If
        Next i

        Dim inspList As String
        inspList = ""
        Dim ik As Variant
        For Each ik In inspDict.Keys
            If inspList <> "" Then inspList = inspList & ", "
            inspList = inspList & UCase(CStr(ik))
        Next ik

        wsOut.Cells(outRow, 1).Value = "Inspected By: " & inspList
        wsOut.Cells(outRow, 1).Font.Size = 12
        wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
        wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
        outRow = outRow + 1

        Dim aQty As Long
        aQty = 0
        For fi = 1 To filtCount
            If CStr(wsPR.Cells(filteredRows(fi), 4).Value) = curArea Then
                If IsNumeric(wsPR.Cells(filteredRows(fi), 7).Value) Then aQty = aQty + CLng(wsPR.Cells(filteredRows(fi), 7).Value)
            End If
        Next fi

        wsOut.Cells(outRow, 1).Value = "Total Items Returned: " & aQty
        wsOut.Cells(outRow, 1).Font.Bold = True
        wsOut.Cells(outRow, 1).Font.Size = 12
        wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
        wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
        outRow = outRow + 2

        ' Records sorted by product
        Dim recRows() As Long
        Dim rc As Long
        rc = 0
        For fi = 1 To filtCount
            If CStr(wsPR.Cells(filteredRows(fi), 4).Value) = curArea Then rc = rc + 1
        Next fi

        If rc > 0 Then
            ReDim recRows(1 To rc)
            Dim ri As Long
            ri = 0
            For fi = 1 To filtCount
                If CStr(wsPR.Cells(filteredRows(fi), 4).Value) = curArea Then
                    ri = ri + 1
                    recRows(ri) = filteredRows(fi)
                End If
            Next fi

            Dim tmpRow As Long
            For i = 1 To rc - 1
                For j = i + 1 To rc
                    If CStr(wsPR.Cells(recRows(i), 6).Value) > CStr(wsPR.Cells(recRows(j), 6).Value) Then
                        tmpRow = recRows(i)
                        recRows(i) = recRows(j)
                        recRows(j) = tmpRow
                    End If
                Next j
            Next i

            For ri = 1 To rc
                Dim sr As Long
                sr = recRows(ri)

                wsOut.Cells(outRow, 1).Value = "Item: " & CStr(wsPR.Cells(sr, 6).Value)
                wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
                wsOut.Cells(outRow, 1).Font.Size = 12
                wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
                outRow = outRow + 1

                wsOut.Cells(outRow, 1).Value = "Quantity: " & CStr(wsPR.Cells(sr, 7).Value)
                wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
                wsOut.Cells(outRow, 1).Font.Size = 12
                wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
                outRow = outRow + 1

                wsOut.Cells(outRow, 1).Value = "Date Returned: " & UCase(Format(wsPR.Cells(sr, 11).Value, "DD MMM YYYY"))
                wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
                wsOut.Cells(outRow, 1).Font.Size = 12
                wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
                outRow = outRow + 1

                wsOut.Cells(outRow, 1).Value = "Date Checked: " & UCase(Format(wsPR.Cells(sr, 2).Value, "DD MMM YYYY"))
                wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
                wsOut.Cells(outRow, 1).Font.Size = 12
                wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
                outRow = outRow + 1

                Dim pdStr As String
                pdStr = "Prod. Date: " & UCase(Format(wsPR.Cells(sr, 9).Value, "DD MMM YYYY"))
                If CStr(wsPR.Cells(sr, 8).Value) <> "" Then pdStr = pdStr & " " & CStr(wsPR.Cells(sr, 8).Value)
                wsOut.Cells(outRow, 1).Value = pdStr
                wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
                wsOut.Cells(outRow, 1).Font.Size = 12
                wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
                outRow = outRow + 1

                wsOut.Cells(outRow, 1).Value = "Expiry Date: " & UCase(Format(wsPR.Cells(sr, 10).Value, "DD MMM YYYY"))
                wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
                wsOut.Cells(outRow, 1).Font.Size = 12
                wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
                outRow = outRow + 1

                wsOut.Cells(outRow, 1).Value = "Description: " & CStr(wsPR.Cells(sr, 12).Value)
                wsOut.Cells(outRow, 1).Font.Name = "Times New Roman"
                wsOut.Cells(outRow, 1).Font.Size = 12
                wsOut.Cells(outRow, 1).Font.Color = RGB(30, 41, 59)
                outRow = outRow + 1

                outRow = outRow + 1
            Next ri
        End If
    Next aIdx

    wsOut.Activate
    wsOut.Cells(1, 1).Select
    Application.ScreenUpdating = True
    MsgBox "Sensory Evaluation report generated for " & selRoute & "!", vbInformation
End Sub`;
}

// ── Main ────────────────────────────────────────────────────────────────────

async function main() {
  console.log("Generating QA Returns workbook...");
  const wb = new ExcelJS.Workbook();
  wb.creator = "QA Returns Generator";
  wb.created = new Date();

  buildLookupsSheet(wb);
  buildDataEntrySheet(wb);    // ← Product Returns form
  buildTrayEntrySheet(wb);    // ← Tray Returns form
  buildProductReturnsSheet(wb);
  buildTrayReturnsSheet(wb);
  buildSummarySheet(wb);
  buildReportGeneratorSheet(wb);
  buildVbaCodeSheet(wb);

  const outputDir = path.join(__dirname, "output");
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const outputPath = path.join(outputDir, "QAReturns.xlsx");
  await wb.xlsx.writeFile(outputPath);
  console.log(`Workbook written to ${outputPath}`);
}

main().catch((err) => {
  console.error("Error generating workbook:", err);
  process.exit(1);
});
