// "Export Excel": a styled, live copy of the payment request.
//
// Amounts are plain numbers; every %, remaining balance and total is an Excel
// formula. Typing a new amount in Excel updates its %, the balances after it
// and the totals straight away. Unit price and discount are live too: the
// price after discount (and the final price) are formulas on them.
//
// Layout follows the sales team's own sheet: Unit Information on top, then
// Terms of Payment grouped by year, in the PDF's forest + gold palette.
import ExcelJS from "exceljs";
import type { Plan } from "./projects";
import type { ScheduleResult } from "./schedule";
import { formatDate, getCurrency, monthsToYears, percent } from "./format";

export interface ExcelInfo {
  companyName: string;
  projectName: string;
  plan: Plan;
  unitType: string;
  unitArea: number; // 0 = not entered
  outdoorArea: number; // 0 = not entered
  floor: string; // "" = not entered
  contractDate: Date;
}

const COLOR = {
  forest: "FF1D3B32",
  forestSoft: "FFE4EDE8",
  forestTint: "FFF3F7F5",
  gold: "FFA9813C",
  goldSoft: "FFF6EFE1",
  goldTint: "FFFCF8F1",
  ink: "FF16181D",
  muted: "FF717783",
  line: "FFD5D9DE",
  red: "FFC00000",
  white: "FFFFFFFF",
};

// A=Year/labels, B=Payment/values, C=Amount, D=Due date, E=%, F=Remaining
const COLS = ["", "A", "B", "C", "D", "E", "F"];
const LAST_COL = 6;
const ref = (col: number, row: number) => `${COLS[col]}${row}`;
const abs = (col: number, row: number) => `$${COLS[col]}$${row}`;

const PCT = "0.00%";
const NUM = "#,##0";
const NUM_RED = "#,##0;[Red]-#,##0";
const DATE = "dd/mm/yyyy";
const AREA = 'General" m²"';

type Style = {
  font?: Partial<ExcelJS.Font>;
  fill?: ExcelJS.Fill;
  border?: Partial<ExcelJS.Borders>;
  alignment?: Partial<ExcelJS.Alignment>;
  numFmt?: string;
};

const solid = (argb: string): ExcelJS.Fill => ({
  type: "pattern",
  pattern: "solid",
  fgColor: { argb },
});
const font = (o: Partial<ExcelJS.Font> = {}): Partial<ExcelJS.Font> => ({
  name: "Calibri",
  size: 11,
  color: { argb: COLOR.ink },
  ...o,
});
const line = (
  style: ExcelJS.BorderStyle = "thin",
  argb = COLOR.line,
): Partial<ExcelJS.Border> => ({ style, color: { argb } });
const box = (argb = COLOR.line): Partial<ExcelJS.Borders> => ({
  top: line("thin", argb),
  left: line("thin", argb),
  bottom: line("thin", argb),
  right: line("thin", argb),
});

// Excel stores dates without a timezone and ExcelJS writes them as UTC, so a
// local midnight east of Greenwich would land on the previous day.
const excelDate = (d: Date) =>
  new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));

// Number format with the currency in front: "EGP 1,234,567".
function moneyFormat(): string {
  const cur = getCurrency().replace(/"/g, "");
  return cur ? `"${cur} "${NUM}` : NUM;
}

function paint(cell: ExcelJS.Cell, s: Style) {
  if (s.font) cell.font = { ...s.font };
  if (s.fill) cell.fill = s.fill;
  if (s.border) cell.border = { ...s.border };
  if (s.alignment) cell.alignment = { ...s.alignment };
  if (s.numFmt) cell.numFmt = s.numFmt;
}

// Styles every cell of a block (merged cells need their borders on each cell)
// and merges it when it is wider/taller than one cell. Returns the top-left.
function block(
  ws: ExcelJS.Worksheet,
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  s: Style,
  value?: ExcelJS.CellValue,
): ExcelJS.Cell {
  for (let r = r1; r <= r2; r++)
    for (let c = c1; c <= c2; c++) paint(ws.getCell(r, c), s);
  if (r1 !== r2 || c1 !== c2) ws.mergeCells(r1, c1, r2, c2);
  const cell = ws.getCell(r1, c1);
  if (value !== undefined) cell.value = value;
  return cell;
}

const formula = (f: string, result: number): ExcelJS.CellFormulaValue => ({
  formula: f,
  result,
});

function band(ws: ExcelJS.Worksheet, row: number, text: string) {
  ws.getRow(row).height = 22;
  block(
    ws,
    row,
    1,
    row,
    LAST_COL,
    {
      font: font({ bold: true, size: 11, color: { argb: COLOR.white } }),
      fill: solid(COLOR.forest),
      alignment: { vertical: "middle", horizontal: "left", indent: 1 },
    },
    text.toUpperCase(),
  );
}

function headerRow(ws: ExcelJS.Worksheet, row: number, titles: string[]) {
  ws.getRow(row).height = 20;
  titles.forEach((t, i) =>
    block(
      ws,
      row,
      i + 1,
      row,
      i + 1,
      {
        font: font({ bold: true, size: 10.5, color: { argb: COLOR.forest } }),
        fill: solid(COLOR.goldSoft),
        border: { ...box(), bottom: line("medium", COLOR.gold) },
        alignment: { vertical: "middle", horizontal: "center" },
      },
      t,
    ),
  );
}

function totalRow(
  ws: ExcelJS.Worksheet,
  row: number,
  label: string,
  amount: ExcelJS.CellFormulaValue,
  pct: ExcelJS.CellFormulaValue,
) {
  ws.getRow(row).height = 20;
  const s: Style = {
    font: font({ bold: true, color: { argb: COLOR.white } }),
    fill: solid(COLOR.forest),
    border: box(COLOR.forest),
    alignment: { vertical: "middle", horizontal: "center" },
  };
  block(ws, row, 1, row, 2, s, label);
  block(ws, row, 3, row, 3, {
    ...s,
    alignment: { vertical: "middle", horizontal: "right", indent: 1 },
    numFmt: NUM,
  }, amount);
  block(ws, row, 4, row, 4, s);
  block(ws, row, 5, row, 5, { ...s, numFmt: PCT }, pct);
  block(ws, row, 6, row, 6, s);
}

// The total % turns red whenever the amounts stop adding up to the target, so
// an edited amount that over- or under-collects is spotted at once.
function flagTotal(ws: ExcelJS.Worksheet, cell: string, target: string) {
  ws.addConditionalFormatting({
    ref: cell,
    rules: [
      {
        type: "expression",
        priority: 1,
        formulae: [`ROUND(${cell}-${target},4)<>0`],
        style: {
          font: { bold: true, color: { argb: "FF9C0006" } },
          fill: {
            type: "pattern",
            pattern: "solid",
            bgColor: { argb: "FFFFC7CE" },
          },
        },
      },
    ],
  });
}

// A plain floor number goes in as a number, so Excel doesn't flag it as
// "number stored as text"; anything else ("Ground", "B1") stays text.
function floorValue(floor: string): ExcelJS.CellValue | undefined {
  const s = floor.trim();
  if (!s) return undefined;
  return /^-?(0|[1-9]\d{0,2})$/.test(s) ? Number(s) : s;
}

type InfoItem = {
  k: string;
  v?: ExcelJS.CellValue;
  numFmt?: string;
  bold?: boolean;
  color?: string;
  fill?: string;
  big?: boolean;
};

export function buildWorkbook(
  result: ScheduleResult,
  info: ExcelInfo,
  logoBase64?: string,
): ExcelJS.Workbook {
  const { plan } = info;
  const company = info.companyName || "Payment Request";
  const MONEY = moneyFormat();

  const wb = new ExcelJS.Workbook();
  wb.creator = company;
  wb.created = new Date();
  // Excel recalculates on open; the cached results below are for previewers
  // (phones, mail apps) that only show what is stored.
  wb.calcProperties.fullCalcOnLoad = true;

  const ws = wb.addWorksheet("Payment Request", {
    views: [{ showGridLines: false }],
    pageSetup: {
      paperSize: 9, // A4
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      margins: {
        left: 0.4,
        right: 0.4,
        top: 0.5,
        bottom: 0.6,
        header: 0.2,
        footer: 0.25,
      },
    },
  });
  ws.columns = [
    { width: 17 },
    { width: 26 },
    { width: 17 },
    { width: 14 },
    { width: 11 },
    { width: 17 },
  ];
  const escapedCompany = company.replace(/&/g, "&&");
  ws.headerFooter.oddFooter = `&L&8&K717783${escapedCompany}&R&8&K717783Page &P of &N`;

  // ---------------- letterhead ----------------
  ws.getRow(1).height = 30;
  ws.getRow(2).height = 26;
  if (logoBase64) {
    const id = wb.addImage({ base64: logoBase64, extension: "png" });
    ws.addImage(id, {
      tl: { col: 0.1, row: 0.12 },
      ext: { width: 132, height: 66 },
      editAs: "oneCell",
    });
  } else {
    block(
      ws,
      1,
      1,
      2,
      2,
      {
        font: font({ bold: true, size: 14, color: { argb: COLOR.forest } }),
        alignment: { vertical: "middle", horizontal: "left" },
      },
      company,
    );
  }
  block(
    ws,
    1,
    3,
    1,
    LAST_COL,
    {
      font: font({ bold: true, size: 20, color: { argb: COLOR.forest } }),
      alignment: { vertical: "bottom", horizontal: "right" },
    },
    "PAYMENT REQUEST",
  );
  block(
    ws,
    2,
    3,
    2,
    LAST_COL,
    {
      font: font({ size: 10, color: { argb: COLOR.muted } }),
      alignment: { vertical: "top", horizontal: "right" },
    },
    `${company} · Issued ${formatDate(info.contractDate)}`,
  );
  for (let c = 1; c <= LAST_COL; c++)
    ws.getCell(2, c).border = { bottom: line("medium", COLOR.gold) };
  ws.getRow(3).height = 8;

  // ---------------- reservation date ----------------
  ws.getRow(4).height = 20;
  block(
    ws,
    4,
    1,
    4,
    1,
    {
      font: font({ bold: true, color: { argb: COLOR.forest } }),
      alignment: { vertical: "middle" },
    },
    "Reservation Date:",
  );
  block(
    ws,
    4,
    2,
    4,
    2,
    {
      font: font({ bold: true }),
      alignment: { vertical: "middle", horizontal: "left" },
      numFmt: DATE,
    },
    excelDate(info.contractDate),
  );
  ws.getRow(5).height = 8;

  // ---------------- unit information ----------------
  band(ws, 6, "Unit Information");
  const top = 7;

  const left: InfoItem[] = [
    { k: "Project", v: info.projectName, bold: true },
    { k: "Unit Type", v: info.unitType || undefined },
    {
      k: "Unit Area",
      v: info.unitArea > 0 ? info.unitArea : undefined,
      numFmt: AREA,
    },
    {
      k: "Outdoor Area",
      v: info.outdoorArea > 0 ? info.outdoorArea : undefined,
      numFmt: AREA,
    },
    { k: "Floor", v: floorValue(info.floor) },
  ];
  const areaRow = top + 2; // "Unit Area" above — the finishing cost uses it
  if (plan.type) left.push({ k: "Payment Status", v: plan.type });
  if (plan.phase) left.push({ k: "Phase", v: plan.phase });
  if (plan.deliveryMonths > 0)
    left.push({ k: "Delivery", v: monthsToYears(plan.deliveryMonths) });

  // Right column: unit price → discount → after discount (→ finishing) →
  // final price. Rows are fixed so the formulas can point at them.
  const priceRow = top;
  const discRow = top + 1;
  const netRow = top + 2;
  const finishing = result.finishingCost > 0;
  const rateRow = top + 3;
  const finishRow = top + 4;
  const finalRow = finishing ? top + 5 : netRow;
  const priceCell = abs(5, priceRow);
  const discCell = abs(5, discRow);
  const finalCell = abs(5, finalRow);

  const premium = plan.discountPct < 0;
  const discPct = Math.abs(plan.discountPct);
  const right: InfoItem[] = [
    {
      k: "Unit Price",
      v: result.originalPrice,
      numFmt: MONEY,
      bold: true,
      fill: COLOR.goldTint,
    },
    {
      k: premium ? "Premium" : "Discount",
      v: discPct,
      numFmt: PCT,
      bold: true,
      color: COLOR.gold,
      fill: COLOR.goldTint,
    },
    {
      k: "Unit Price After Discount",
      v: formula(
        `${priceCell}*(1${premium ? "+" : "-"}${discCell})`,
        result.netPrice,
      ),
      numFmt: MONEY,
      bold: true,
    },
  ];
  if (finishing)
    right.push(
      {
        k: "Finishing Price / m²",
        v: plan.buaRate,
        numFmt: MONEY,
        fill: COLOR.goldTint,
      },
      {
        k: "Finishing Cost",
        v: formula(`${abs(2, areaRow)}*${abs(5, rateRow)}`, result.finishingCost),
        numFmt: MONEY,
      },
      {
        k: "Final Unit Price",
        v: formula(`${abs(5, netRow)}+${abs(5, finishRow)}`, result.finalPrice),
        numFmt: MONEY,
        bold: true,
      },
    );
  right[right.length - 1].big = true; // the price the % are taken on

  const infoRows = Math.max(left.length, right.length);
  for (let i = 0; i < infoRows; i++) {
    const row = top + i;
    ws.getRow(row).height = 19;
    const dotted: Partial<ExcelJS.Borders> = { bottom: line("hair", "FFB9BFC6") };
    const l = left[i];
    if (l) {
      block(
        ws,
        row,
        1,
        row,
        1,
        {
          font: font({ bold: true, size: 10.5, color: { argb: COLOR.forest } }),
          border: dotted,
          alignment: { vertical: "middle" },
        },
        l.k,
      );
      block(
        ws,
        row,
        2,
        row,
        2,
        {
          font: font({ bold: l.bold }),
          border: dotted,
          alignment: { vertical: "middle", horizontal: "left", shrinkToFit: true },
          numFmt: l.numFmt,
        },
        l.v ?? null,
      );
    }
    const r = right[i];
    if (r) {
      const hl = r.big; // final price: highlighted box
      block(
        ws,
        row,
        3,
        row,
        4,
        {
          font: font({ bold: true, size: 10.5, color: { argb: COLOR.forest } }),
          fill: hl ? solid(COLOR.forestSoft) : undefined,
          border: hl ? box(COLOR.forest) : dotted,
          alignment: { vertical: "middle", horizontal: "left", indent: 1 },
        },
        r.k,
      );
      block(
        ws,
        row,
        5,
        row,
        6,
        {
          font: font({
            bold: r.bold,
            size: hl ? 12 : 11,
            color: { argb: hl ? COLOR.forest : (r.color ?? COLOR.ink) },
          }),
          fill: hl
            ? solid(COLOR.forestSoft)
            : r.fill
              ? solid(r.fill)
              : undefined,
          border: hl ? box(COLOR.forest) : dotted,
          alignment: { vertical: "middle", horizontal: "right", indent: 1 },
          numFmt: r.numFmt,
        },
        r.v,
      );
    }
  }

  // Payment plan — on its own full-width line, it is often long.
  let row = top + infoRows;
  ws.getRow(row).height = 21;
  block(
    ws,
    row,
    1,
    row,
    1,
    {
      font: font({ bold: true, size: 10.5, color: { argb: COLOR.forest } }),
      border: { bottom: line("hair", "FFB9BFC6") },
      alignment: { vertical: "middle" },
    },
    "Payment Plan",
  );
  block(
    ws,
    row,
    2,
    row,
    LAST_COL,
    {
      font: font({ bold: true, color: { argb: COLOR.red } }),
      border: { bottom: line("hair", "FFB9BFC6") },
      alignment: { vertical: "middle", horizontal: "left" },
    },
    plan.label,
  );
  row += 2;

  // ---------------- terms of payment ----------------
  band(ws, row, "Terms of Payment");
  row++;
  headerRow(ws, row, ["Year", "Payment", "Amount", "Due Date", "%", "Remaining"]);
  row++;

  const first = row;
  const rows = result.rows;
  const last = first + rows.length - 1;
  rows.forEach((p, i) => {
    const r = first + i;
    const isDp = p.year === 0;
    const bg = isDp
      ? COLOR.goldTint
      : p.year % 2 === 0
        ? COLOR.forestTint
        : COLOR.white;
    const base: Style = {
      font: font({ bold: isDp }),
      fill: solid(bg),
      border: box(),
      alignment: { vertical: "middle", horizontal: "center" },
    };
    ws.getRow(r).height = 18;
    block(ws, r, 2, r, 2, {
      ...base,
      alignment: { vertical: "middle", horizontal: "left", indent: 1 },
    }, p.label);
    block(ws, r, 3, r, 3, {
      ...base,
      alignment: { vertical: "middle", horizontal: "right", indent: 1 },
      numFmt: NUM,
    }, p.amount);
    block(ws, r, 4, r, 4, {
      ...base,
      font: font({ bold: isDp, italic: true }),
      numFmt: DATE,
    }, excelDate(p.dueDate));
    block(ws, r, 5, r, 5, {
      ...base,
      font: font({ bold: true, color: { argb: COLOR.forest } }),
      numFmt: PCT,
    }, formula(`IFERROR(${ref(3, r)}/${finalCell},0)`, p.percent));
    block(ws, r, 6, r, 6, {
      ...base,
      font: font({ color: { argb: COLOR.muted } }),
      alignment: { vertical: "middle", horizontal: "right", indent: 1 },
      numFmt: NUM_RED,
    }, formula(
      i === 0
        ? `ROUND(${finalCell}-${ref(3, r)},2)`
        : `ROUND(${ref(6, r - 1)}-${ref(3, r)},2)`,
      p.balance,
    ));
  });

  // Year column: one merged cell per year (and one for the down payments).
  for (let i = 0; i < rows.length; ) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1].year === rows[i].year) j++;
    const y = rows[i].year;
    block(
      ws,
      first + i,
      1,
      first + j,
      1,
      {
        font: font({ bold: true, color: { argb: COLOR.forest } }),
        fill: solid(y === 0 ? COLOR.goldSoft : COLOR.forestSoft),
        border: box(),
        alignment: { vertical: "middle", horizontal: "center", wrapText: true },
      },
      y === 0 ? (j > i ? "Down Payments" : "Down Payment") : `Year ${y}`,
    );
    i = j + 1;
  }

  row = last + 1;
  totalRow(
    ws,
    row,
    "Total",
    formula(`SUM(${ref(3, first)}:${ref(3, last)})`, result.scheduleTotal),
    formula(
      `SUM(${ref(5, first)}:${ref(5, last)})`,
      result.finalPrice ? result.scheduleTotal / result.finalPrice : 0,
    ),
  );
  flagTotal(ws, ref(5, row), "1");
  // Maintenance + totals print on a page of their own, as in the PDF, so a
  // section heading never ends up alone at the foot of a page.
  ws.getRow(row + 1).addPageBreak();
  row += 2;

  // ---------------- maintenance ----------------
  let maintTotalCell: string | null = null;
  if (result.maintenanceTotal > 0 && result.maintenanceRows.length > 0) {
    const basisCell = plan.maintenanceBasis === "original" ? priceCell : finalCell;
    const basisPrice =
      plan.maintenanceBasis === "original" ? result.originalPrice : result.finalPrice;
    band(
      ws,
      row,
      `Maintenance — ${percent(plan.maintenancePct)} of ${
        plan.maintenanceBasis === "original" ? "unit price" : "final price"
      }`,
    );
    row++;
    headerRow(ws, row, ["No.", "Payment", "Amount", "Due Date", "%", "Remaining"]);
    row++;
    const mFirst = row;
    const mRows = result.maintenanceRows;
    const mLast = mFirst + mRows.length - 1;
    const mTotal = mLast + 1;
    let left = result.maintenanceTotal;
    mRows.forEach((m, i) => {
      const r = mFirst + i;
      left -= m.amount;
      const base: Style = {
        font: font(),
        fill: solid(i % 2 ? COLOR.forestTint : COLOR.white),
        border: box(),
        alignment: { vertical: "middle", horizontal: "center" },
      };
      ws.getRow(r).height = 18;
      block(ws, r, 1, r, 1, {
        ...base,
        font: font({ bold: true, color: { argb: COLOR.forest } }),
      }, i + 1);
      block(ws, r, 2, r, 2, {
        ...base,
        alignment: { vertical: "middle", horizontal: "left", indent: 1 },
      }, m.label);
      block(ws, r, 3, r, 3, {
        ...base,
        alignment: { vertical: "middle", horizontal: "right", indent: 1 },
        numFmt: NUM,
      }, m.amount);
      block(ws, r, 4, r, 4, {
        ...base,
        font: font({ italic: true }),
        numFmt: DATE,
      }, excelDate(m.dueDate));
      block(ws, r, 5, r, 5, {
        ...base,
        font: font({ bold: true, color: { argb: COLOR.forest } }),
        numFmt: PCT,
      }, formula(
        `IFERROR(${ref(3, r)}/${basisCell},0)`,
        basisPrice ? m.amount / basisPrice : 0,
      ));
      block(ws, r, 6, r, 6, {
        ...base,
        font: font({ color: { argb: COLOR.muted } }),
        alignment: { vertical: "middle", horizontal: "right", indent: 1 },
        numFmt: NUM_RED,
      }, formula(
        i === 0
          ? `ROUND(${abs(3, mTotal)}-${ref(3, r)},2)`
          : `ROUND(${ref(6, r - 1)}-${ref(3, r)},2)`,
        Math.abs(left) < 0.5 ? 0 : left,
      ));
    });
    totalRow(
      ws,
      mTotal,
      "Maintenance Total",
      formula(`SUM(${ref(3, mFirst)}:${ref(3, mLast)})`, result.maintenanceTotal),
      formula(
        `SUM(${ref(5, mFirst)}:${ref(5, mLast)})`,
        basisPrice ? result.maintenanceTotal / basisPrice : 0,
      ),
    );
    maintTotalCell = abs(3, mTotal);
    row = mTotal + 2;
  }

  // ---------------- totals ----------------
  band(ws, row, "Totals");
  row++;
  const totals: [string, ExcelJS.CellFormulaValue | number][] = [
    ["Final Unit Price", formula(finalCell, result.finalPrice)],
    [
      "Maintenance",
      maintTotalCell ? formula(maintTotalCell, result.maintenanceTotal) : 0,
    ],
  ];
  const totFirst = row;
  for (const [k, v] of totals) {
    ws.getRow(row).height = 19;
    const s: Style = {
      font: font({ bold: true, size: 10.5, color: { argb: COLOR.forest } }),
      border: { bottom: line("hair", "FFB9BFC6") },
      alignment: { vertical: "middle", horizontal: "left", indent: 1 },
    };
    block(ws, row, 1, row, 4, s, k);
    block(ws, row, 5, row, 6, {
      ...s,
      font: font(),
      alignment: { vertical: "middle", horizontal: "right", indent: 1 },
      numFmt: MONEY,
    }, v);
    row++;
  }
  ws.getRow(row).height = 24;
  const grand: Style = {
    font: font({ bold: true, size: 12, color: { argb: COLOR.white } }),
    fill: solid(COLOR.forest),
    border: box(COLOR.forest),
    alignment: { vertical: "middle", horizontal: "left", indent: 1 },
  };
  block(ws, row, 1, row, 4, grand, "Grand Total");
  block(ws, row, 5, row, 6, {
    ...grand,
    alignment: { vertical: "middle", horizontal: "right", indent: 1 },
    numFmt: MONEY,
  }, formula(`${ref(5, totFirst)}+${ref(5, totFirst + 1)}`, result.grandTotal));
  row += 2;

  // ---------------- footer ----------------
  ws.getRow(row).height = 30;
  block(
    ws,
    row,
    1,
    row,
    LAST_COL,
    {
      font: font({ italic: true, size: 9, color: { argb: COLOR.muted } }),
      alignment: { vertical: "top", horizontal: "left", wrapText: true },
    },
    `${company} — This Payment Request is a price proposal and does not constitute a contract. ` +
      "Instalment due dates are calculated from the reservation date and are subject to the signed contract.",
  );

  ws.pageSetup.printArea = `A1:F${row}`;
  return wb;
}

function toBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000)
    s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

export function excelFileName(info: ExcelInfo): string {
  const d = info.contractDate;
  const date = `${String(d.getDate()).padStart(2, "0")}-${String(
    d.getMonth() + 1,
  ).padStart(2, "0")}-${d.getFullYear()}`;
  const name = `Payment Request - ${info.projectName} - ${date}`;
  return `${name.replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim()}.xlsx`;
}

// Builds the workbook and hands it to the browser as a download. The logo is
// optional: if it can't be fetched the sheet just shows the company name.
export async function exportToExcel(
  result: ScheduleResult,
  info: ExcelInfo,
  logoUrl: string | null,
): Promise<void> {
  let logo: string | undefined;
  if (logoUrl) {
    try {
      const res = await fetch(logoUrl);
      if (res.ok) logo = toBase64(await res.arrayBuffer());
    } catch {
      logo = undefined;
    }
  }
  const wb = buildWorkbook(result, info, logo);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = excelFileName(info);
  document.body.appendChild(a);
  a.click();
  a.remove();
  // give the browser time to start the download before freeing the file
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}
