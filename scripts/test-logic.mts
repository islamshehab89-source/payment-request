// Logic tests against the REAL public/projects.xlsx through the real loader.
// These assert INVARIANTS that must hold no matter how the sheet is edited,
// plus the reference example from the user's screenshot. Run: pnpm test
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const buf = fs.readFileSync(path.join(root, "public", "projects.xlsx"));
(globalThis as { fetch: unknown }).fetch = async () =>
  new Response(buf, { status: 200 });

const { loadProjects } = await import("../src/lib/loadProjects.ts");
const { computeSchedule } = await import("../src/lib/schedule.ts");

let failures = 0;
function check(name: string, actual: unknown, expected: unknown) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) failures++;
  console.log(
    `${ok ? "PASS" : "FAIL"}  ${name}  ${ok ? "" : `actual=${JSON.stringify(actual)} expected=${JSON.stringify(expected)}`}`,
  );
}

const data = await loadProjects();
console.log(
  "source:",
  data.source,
  "| company:",
  data.companyName,
  "| currency:",
  data.currency,
);
if (data.errors.length) {
  console.log("loader messages:");
  data.errors.forEach((e) => console.log("  -", e));
}

check("loaded from excel", data.source, "excel");
check("has projects", data.projects.length > 0, true);

// -------- INVARIANT: every plan produces a schedule that balances --------
const PRICE = 1_000_000;
const AREA = 20;
const contract = new Date(2026, 5, 12);
let planCount = 0;
for (const p of data.projects) {
  for (const pl of p.plans) {
    planCount++;
    const area = pl.buaRate > 0 ? AREA : 0;
    const r = computeSchedule(pl, {
      originalPrice: PRICE,
      unitArea: area,
      contractDate: contract,
    });
    const label = `${p.name} / ${pl.type ?? "-"} / ${pl.label}`;

    // final price built correctly
    const expFinal =
      PRICE - PRICE * pl.discountPct + (pl.buaRate > 0 ? area * pl.buaRate : 0);
    check(`[${label}] final price`, Math.round(r.finalPrice), Math.round(expFinal));

    // down payments on the right basis
    const dpBasePrice = pl.dpBasis === "original" ? PRICE : r.finalPrice;
    const dpRows = r.rows.filter((x) => x.code.startsWith("DP"));
    const expDpSum = Math.round(
      dpBasePrice * (pl.downPayment1Pct + pl.downPayment2Pct),
    );
    const dpSum = Math.round(dpRows.reduce((s, x) => s + x.amount, 0));
    check(`[${label}] down-payment sum`, dpSum, expDpSum);

    // schedule covers the full final price, balance closes at 0
    check(`[${label}] schedule total = final`, Math.round(r.scheduleTotal), Math.round(r.finalPrice));
    check(`[${label}] unscheduled = 0`, r.unscheduled, 0);
    if (r.rows.length) {
      check(`[${label}] last balance = 0`, Math.round(r.rows[r.rows.length - 1].balance), 0);
    }
    check(`[${label}] no negative amount`, r.rows.every((x) => x.amount >= 0), true);

    // maintenance on the right basis, rows sum to total
    const mBasePrice = pl.maintenanceBasis === "original" ? PRICE : r.finalPrice;
    check(
      `[${label}] maintenance total`,
      r.maintenanceTotal,
      Math.round(mBasePrice * pl.maintenancePct),
    );
    const mSum = Math.round(r.maintenanceRows.reduce((s, x) => s + x.amount, 0));
    check(`[${label}] maintenance rows sum`, mSum, r.maintenanceTotal);

    check(
      `[${label}] grand total`,
      Math.round(r.grandTotal),
      Math.round(r.finalPrice + r.maintenanceTotal),
    );
  }
}
console.log(`\nChecked ${planCount} plans for balance invariants.`);

// -------- Reference example from the user's screenshot --------
const mornings = data.projects.find((p) => p.name === "The Mornings_Everyday");
const finished7y = mornings?.plans.find(
  (p) => p.type === "Finished" && p.label.includes("7 years"),
);
if (!finished7y) {
  console.log("FAIL  reference plan (Finished 7 years) not found");
  failures++;
} else {
  const r = computeSchedule(finished7y, {
    originalPrice: 1_000_000,
    unitArea: 20,
    contractDate: contract,
  });
  check("ref: BUA rate", finished7y.buaRate, 20000);
  check("ref: discount 3%", finished7y.discountPct, 0.03);
  check("ref: net price", r.netPrice, 970_000);
  check("ref: finishing cost", r.finishingCost, 400_000);
  check("ref: final price", r.finalPrice, 1_370_000);
  check("ref: 31 installments", r.rows.filter((x) => x.code.startsWith("INS")).length, 31);
  // first maintenance Jun 2026 + 15 months = Sep 2027
  check(
    "ref: maintenance first month",
    r.maintenanceRows.length
      ? `${r.maintenanceRows[0].dueDate.getFullYear()}-${r.maintenanceRows[0].dueDate.getMonth() + 1}`
      : "none",
    "2027-9",
  );
}

console.log(
  failures === 0 ? "\nALL CHECKS PASSED" : `\n${failures} CHECK(S) FAILED`,
);
process.exit(failures === 0 ? 0 : 1);
