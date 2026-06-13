import type { ScheduleResult } from "./schedule";
import { formatDate, rawNumber, getCurrency } from "./format";

export interface ExportMeta {
  projectName: string;
  planLabel: string;
  unitStatus: string | null;
  phase: string | null;
  unitType: string; // "" = not entered
  unitArea: number; // 0 = finishing not used
  outdoorArea: number; // 0 = not entered
  contractDate: Date;
}

export async function exportToExcel(result: ScheduleResult, meta: ExportMeta) {
  const XLSX = await import("xlsx");

  const currency = getCurrency();
  const cur = currency ? ` (${currency})` : "";

  const rows: (string | number)[][] = [];
  rows.push(["Payment Request"]);
  rows.push([]);
  rows.push(["Unit Information"]);
  rows.push(["Project", meta.projectName]);
  if (meta.unitType) rows.push(["Unit Type", meta.unitType]);
  if (meta.unitArea > 0) rows.push(["Unit Area (m²)", meta.unitArea]);
  if (meta.outdoorArea > 0) rows.push(["Outdoor Area (m²)", meta.outdoorArea]);
  if (meta.unitStatus) rows.push(["Unit Status", meta.unitStatus]);
  if (meta.phase) rows.push(["Phase", meta.phase]);
  rows.push(["Payment Plan", meta.planLabel]);
  rows.push(["Contract Date", formatDate(meta.contractDate)]);
  rows.push([]);
  rows.push([`Original Unit Price${cur}`, rawNumber(result.originalPrice)]);
  rows.push([`Discount${cur}`, rawNumber(result.discount)]);
  if (result.finishingCost > 0) {
    rows.push([
      `Net Unit Price (Without Finishing)${cur}`,
      rawNumber(result.netPrice),
    ]);
    rows.push([`Finishing Cost${cur}`, rawNumber(result.finishingCost)]);
  }
  rows.push([`Final Unit Price${cur}`, rawNumber(result.finalPrice)]);
  rows.push([]);

  rows.push(["Payment Schedule"]);
  rows.push([
    "Code",
    "Description",
    "Due Date",
    "%",
    `Amount${cur}`,
    `Balance${cur}`,
  ]);
  for (const r of result.rows) {
    rows.push([
      r.code,
      r.label,
      formatDate(r.dueDate),
      `${parseFloat((r.percent * 100).toFixed(2))}%`,
      rawNumber(r.amount),
      rawNumber(r.balance),
    ]);
  }
  rows.push(["", "Schedule Total", "", "", rawNumber(result.scheduleTotal), ""]);

  if (result.maintenanceRows.length > 0) {
    rows.push([]);
    rows.push(["Maintenance Schedule"]);
    rows.push(["Code", "Description", "Due Date", `Amount${cur}`]);
    for (const m of result.maintenanceRows) {
      rows.push([m.code, m.label, formatDate(m.dueDate), rawNumber(m.amount)]);
    }
    rows.push(["", "Maintenance Total", "", rawNumber(result.maintenanceTotal)]);
  }

  rows.push([]);
  rows.push([`Grand Total${cur}`, rawNumber(result.grandTotal)]);

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [
    { wch: 30 },
    { wch: 28 },
    { wch: 14 },
    { wch: 9 },
    { wch: 16 },
    { wch: 16 },
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Payment Request");

  // Keep Arabic letters so Arabic project names don't collapse to "--".
  const safe = (s: string) =>
    s.replace(/[^a-z0-9؀-ۿ]+/gi, "-").replace(/^-+|-+$/g, "") || "order";
  XLSX.writeFile(
    wb,
    `payment-request-${safe(meta.projectName)}-${safe(meta.planLabel)}.xlsx`,
  );
}
