import type { Plan } from "./projects";
import { addMonths } from "./format";

export interface ScheduleInputs {
  originalPrice: number;
  unitArea: number; // m², only used when plan.buaRate > 0
  contractDate: Date;
  netPriceOverride?: number; // optional negotiated net price (without finishing)
}

export interface ScheduleRow {
  code: string; // "DP-1", "DP-2", "INS-1", ...
  label: string;
  dueDate: Date;
  percent: number; // fraction of the final unit price
  amount: number;
  balance: number; // remaining final-price balance after this payment
  year: number; // 0 = down payments, 1..N for installment years
}

export interface MaintenanceRow {
  code: string; // "MNT-1", ...
  label: string;
  dueDate: Date;
  amount: number;
}

export interface ScheduleResult {
  originalPrice: number;
  discount: number;
  netPrice: number; // original − discount (or the override)
  finishingCost: number; // unitArea × buaRate
  finalPrice: number; // netPrice + finishingCost

  rows: ScheduleRow[]; // down payments + installments
  scheduleTotal: number; // sum of rows
  unscheduled: number; // finalPrice − scheduleTotal (0 when the plan adds up)

  maintenanceTotal: number;
  maintenanceRows: MaintenanceRow[];

  grandTotal: number; // finalPrice + maintenanceTotal
}

// Maintenance due months: first, first+every, …, always ending exactly at end.
function maintenanceMonths(plan: Plan): number[] {
  const { maintenanceFirstDueMonths: first, maintenanceEndDueMonths: end } =
    plan;
  const every = plan.maintenanceEveryMonths;
  if (end <= first || every <= 0) return [first];
  const months: number[] = [];
  for (let m = first; m < end; m += every) months.push(m);
  months.push(end);
  return months;
}

export function computeSchedule(
  plan: Plan,
  inputs: ScheduleInputs,
): ScheduleResult {
  const original = inputs.originalPrice;
  const override = inputs.netPriceOverride;
  const netPrice =
    override != null && !Number.isNaN(override)
      ? override
      : original - original * plan.discountPct;
  const discount = original - netPrice;
  const finishingCost = plan.buaRate > 0 ? inputs.unitArea * plan.buaRate : 0;
  const finalPrice = netPrice + finishingCost;

  const rows: ScheduleRow[] = [];
  let balance = finalPrice;

  const pushRow = (
    code: string,
    label: string,
    dueDate: Date,
    amount: number,
    year: number,
  ) => {
    balance -= amount;
    if (Math.abs(balance) < 0.5) balance = 0; // tidy rounding dust
    rows.push({
      code,
      label,
      dueDate,
      percent: finalPrice ? amount / finalPrice : 0,
      amount,
      balance,
      year,
    });
  };

  // Down payments — % of the plan's chosen base (original or final price).
  const dpBase = plan.dpBasis === "original" ? original : finalPrice;
  const dp1 = Math.round(dpBase * plan.downPayment1Pct);
  if (dp1 > 0) {
    pushRow("DP-1", "Down Payment 1", inputs.contractDate, dp1, 0);
  }
  const dp2 = Math.round(dpBase * plan.downPayment2Pct);
  if (dp2 > 0) {
    pushRow(
      "DP-2",
      "Down Payment 2",
      addMonths(inputs.contractDate, plan.monthsToDownPayment2),
      dp2,
      0,
    );
  }

  // Periodic installments — the rest of the FINAL price (so finishing cost is
  // spread across the installments), equal split, last absorbs rounding.
  const remaining = finalPrice - dp1 - dp2;
  const count = plan.installmentsCount;
  if (count > 0 && remaining > 0) {
    const per = Math.round(remaining / count);
    let allocated = 0;
    for (let i = 1; i <= count; i++) {
      const isLast = i === count;
      const amount = isLast ? remaining - allocated : per;
      allocated += amount;
      const monthsOut =
        plan.monthsToFirstInstallment + (i - 1) * plan.installmentEveryMonths;
      const dueDate = addMonths(inputs.contractDate, monthsOut);
      const year =
        Math.floor((monthsOut - plan.monthsToFirstInstallment) / 12) + 1;
      pushRow(`INS-${i}`, `Installment ${i}`, dueDate, amount, year);
    }
  }

  // Maintenance — its own schedule, billed on top of the unit price.
  const maintBase = plan.maintenanceBasis === "original" ? original : finalPrice;
  const maintenanceTotal =
    plan.maintenancePct > 0 ? Math.round(maintBase * plan.maintenancePct) : 0;
  const maintenanceRows: MaintenanceRow[] = [];
  if (maintenanceTotal > 0) {
    const months = maintenanceMonths(plan);
    const per = Math.round(maintenanceTotal / months.length);
    let allocated = 0;
    months.forEach((m, idx) => {
      const isLast = idx === months.length - 1;
      const amount = isLast ? maintenanceTotal - allocated : per;
      allocated += amount;
      maintenanceRows.push({
        code: `MNT-${idx + 1}`,
        label: `Maintenance ${idx + 1}`,
        dueDate: addMonths(inputs.contractDate, m),
        amount,
      });
    });
  }

  const scheduleTotal = rows.reduce((s, r) => s + r.amount, 0);

  return {
    originalPrice: original,
    discount,
    netPrice,
    finishingCost,
    finalPrice,
    rows,
    scheduleTotal,
    unscheduled: Math.round(finalPrice - scheduleTotal),
    maintenanceTotal,
    maintenanceRows,
    grandTotal: finalPrice + maintenanceTotal,
  };
}
