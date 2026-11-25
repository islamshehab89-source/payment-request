import type { Plan } from "./projects";
import { addMonths } from "./format";

export interface ScheduleInputs {
  originalPrice: number;
  unitArea: number; // m², only used when plan.buaRate > 0
  contractDate: Date;
  netPriceOverride?: number; // optional negotiated net price (without finishing)
}

export interface ScheduleRow {
  code: string; // "DP-1", "DP-2", "INS-1", "BUL-1", ...
  label: string;
  dueDate: Date;
  percent: number; // fraction of the final unit price
  amount: number;
  balance: number; // remaining final-price balance after this payment
  // 0 = down payments, 1..N for the years after them. Counted from the first
  // installment, plus the whole grace years before it (an installment due at
  // month 12 opens Year 2), and never past the plan's Years (an installment
  // due on the last anniversary stays in the last year). Plans with bullets
  // or a yearly increase are defined per contract year and so are grouped by
  // contract year.
  year: number;
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
  // Years with no payment before the first installment year (Year 1 …
  // graceYears), e.g. 1 when the installments start a year after the contract.
  graceYears: number;
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

  // Bullets — extra lump sums at the end of contract years 1, 2, … (months 12,
  // 24, …), a % of the same price the down payments are taken on.
  const bullets: { month: number; amount: number; year: number }[] = [];
  (plan.bulletsPct ?? []).forEach((pct, k) => {
    const amount = Math.round(dpBase * pct);
    if (amount > 0) bullets.push({ month: 12 * (k + 1), amount, year: k + 1 });
  });
  const bulletsTotal = bullets.reduce((s, b) => s + b.amount, 0);

  // Contract year of a payment due `m` months after the contract: months 1–12
  // are year 1, 13–24 year 2, … (month 0 counts as year 1).
  const contractYear = (m: number) => Math.max(1, Math.ceil(m / 12));
  const step = plan.yearlyIncreasePct ?? 0;
  const byContractYear = step > 0 || bullets.length > 0;
  // Whole years before the first installment: a plan that starts its
  // installments at month 12 has one grace year, so they start in Year 2.
  const graceYears = Math.floor(plan.monthsToFirstInstallment / 12);
  // No installment is shown past the plan's length: one due on the last
  // anniversary (month 96 of an 8-year plan) closes Year 8, not opens Year 9.
  const lastPlanYear = plan.years > 0 ? Math.ceil(plan.years) : Infinity;

  // Periodic installments — the rest of the FINAL price (so finishing cost is
  // spread across the installments). Equal split, or in a backloaded plan
  // weighted 1, 1 + step, 1 + 2·step… by contract year; the last absorbs
  // rounding.
  const installments: { month: number; amount: number; year: number }[] = [];
  const remaining = finalPrice - dp1 - dp2 - bulletsTotal;
  const count = plan.installmentsCount;
  if (count > 0 && remaining > 0) {
    const months = Array.from(
      { length: count },
      (_, i) => plan.monthsToFirstInstallment + i * plan.installmentEveryMonths,
    );
    const weights = months.map((m) => 1 + step * (contractYear(m) - 1));
    const totalWeight = weights.reduce((s, w) => s + w, 0);
    const per = Math.round(remaining / count);
    let allocated = 0;
    months.forEach((m, i) => {
      const isLast = i === count - 1;
      const amount = isLast
        ? remaining - allocated
        : step > 0
          ? Math.round((remaining * weights[i]) / totalWeight)
          : per;
      allocated += amount;
      const year = Math.max(
        1,
        Math.min(
          lastPlanYear,
          byContractYear
            ? contractYear(m)
            : Math.floor((m - plan.monthsToFirstInstallment) / 12) + 1 + graceYears,
        ),
      );
      installments.push({ month: m, amount, year });
    });
  }

  // Installments and bullets in due-date order; on the same date the
  // installment comes first, then the bullet.
  const timeline = [
    ...installments.map((x, i) => ({
      ...x,
      tie: 0,
      code: `INS-${i + 1}`,
      label: `Installment ${i + 1}`,
    })),
    ...bullets.map((b) => ({
      ...b,
      tie: 1,
      code: `BUL-${b.year}`,
      label: `Bullet Y${b.year}`,
    })),
  ].sort((a, b) => a.month - b.month || a.tie - b.tie);
  for (const x of timeline) {
    pushRow(
      x.code,
      x.label,
      addMonths(inputs.contractDate, x.month),
      x.amount,
      x.year,
    );
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
  const firstYear = rows.find((r) => r.year > 0)?.year ?? 1;

  return {
    originalPrice: original,
    discount,
    netPrice,
    finishingCost,
    finalPrice,
    rows,
    graceYears: firstYear - 1,
    scheduleTotal,
    unscheduled: Math.round(finalPrice - scheduleTotal),
    maintenanceTotal,
    maintenanceRows,
    grandTotal: finalPrice + maintenanceTotal,
  };
}
