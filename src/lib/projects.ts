// =============================================================================
//  PROJECTS & PAYMENT PLANS — FALLBACK SAMPLES + TYPES
// =============================================================================
//  Real data lives in public/projects.xlsx (edited in Excel, parsed by
//  src/lib/loadProjects.ts). The PROJECTS below are only shown when that file
//  is missing or unreadable. Percentages here are fractions: 0.10 = 10%.
//
//  How a plan turns into a schedule (see src/lib/schedule.ts):
//   netPrice      = originalPrice − (originalPrice × discountPct)   [or override]
//   finishingCost = unitArea × buaRate                              [0 when buaRate = 0]
//   finalPrice    = netPrice + finishingCost
//   downPayment1  = dpBase × downPayment1Pct     (dpBase = original or final price)
//   downPayment2  = dpBase × downPayment2Pct     (due monthsToDownPayment2 later)
//   installments  = (finalPrice − dp1 − dp2) / installmentsCount, equal, spaced
//   maintenance   = maintBase × maintenancePct, split into equal payments from
//                   maintenanceFirstDueMonths to maintenanceEndDueMonths every
//                   maintenanceEveryMonths months
// =============================================================================

// What a percentage is taken on: the entered original unit price, or the final
// unit price (after discount, including finishing). In the Excel sheet these
// are written as "Original Price" / "Selling price".
export type PriceBasis = "original" | "selling";

export interface Plan {
  id: string;
  label: string; // shown in the Payment Plan dropdown
  type: string | null; // unit status (Primary / RTM / Finished…), null = not used
  phase: string | null; // project phase, null = not used
  buaRate: number; // finishing price per m²; 0 = no finishing for this plan

  years: number; // informational
  discountPct: number; // discount off the original price (negative = premium)

  dpBasis: PriceBasis; // what the down-payment %s are taken on
  downPayment1Pct: number; // due at contract date
  downPayment2Pct: number; // 0 if there is no second down payment
  monthsToDownPayment2: number;

  installmentsCount: number;
  installmentEveryMonths: number; // 3 = quarterly, 1 = monthly
  monthsToFirstInstallment: number;

  maintenancePct: number;
  maintenanceBasis: PriceBasis;
  maintenanceFirstDueMonths: number; // first maintenance payment (0 = at contract)
  maintenanceEndDueMonths: number; // last maintenance payment
  maintenanceEveryMonths: number; // spacing of maintenance payments

  deliveryMonths: number; // months from contract to delivery (informational)
}

export interface Project {
  id: string;
  name: string;
  plans: Plan[];
}

// -----------------------------------------------------------------------------
//  SAMPLE DATA — fallback only; the real data comes from public/projects.xlsx.
// -----------------------------------------------------------------------------

export const PROJECTS: Project[] = [
  {
    id: "sample-project",
    name: "Sample Project",
    plans: [
      {
        id: "finished-7y",
        label: "5% & 5% & 7 years",
        type: "Finished",
        phase: null,
        buaRate: 20000,
        years: 7,
        discountPct: 0.03,
        dpBasis: "original",
        downPayment1Pct: 0.05,
        downPayment2Pct: 0.05,
        monthsToDownPayment2: 3,
        installmentsCount: 31,
        installmentEveryMonths: 3,
        monthsToFirstInstallment: 3,
        maintenancePct: 0.1,
        maintenanceBasis: "selling",
        maintenanceFirstDueMonths: 15,
        maintenanceEndDueMonths: 39,
        maintenanceEveryMonths: 6,
        deliveryMonths: 42,
      },
      {
        id: "core-8y",
        label: "5% & 5% & 8 years",
        type: "Core & Shell",
        phase: null,
        buaRate: 0,
        years: 8,
        discountPct: 0,
        dpBasis: "original",
        downPayment1Pct: 0.05,
        downPayment2Pct: 0.05,
        monthsToDownPayment2: 3,
        installmentsCount: 31,
        installmentEveryMonths: 3,
        monthsToFirstInstallment: 3,
        maintenancePct: 0.1,
        maintenanceBasis: "original",
        maintenanceFirstDueMonths: 15,
        maintenanceEndDueMonths: 39,
        maintenanceEveryMonths: 6,
        deliveryMonths: 42,
      },
    ],
  },
];

// Currency shown across the app. Change to "" to hide, or "USD", "SAR", etc.
export const CURRENCY = "EGP";

// Company / developer name shown in the header and on the printout.
export const COMPANY_NAME = "Your Company";
