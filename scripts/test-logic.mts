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
  // mirrors the sheet's Installments Count for this plan (7 years × 4 − 1);
  // update it here when that column changes in projects.xlsx
  check("ref: 27 installments", r.rows.filter((x) => x.code.startsWith("INS")).length, 27);
  // first maintenance Jun 2026 + 15 months = Sep 2027
  check(
    "ref: maintenance first month",
    r.maintenanceRows.length
      ? `${r.maintenanceRows[0].dueDate.getFullYear()}-${r.maintenanceRows[0].dueDate.getMonth() + 1}`
      : "none",
    "2027-9",
  );
}

// -------- Status column (Active / Inactive), on synthetic sheets --------
// Each sheet is written to a real xlsx and read back through the loader, the
// same path the live site takes after "publish" exports the Google Sheet.
console.log("\nStatus column:");
const XLSX = await import("xlsx");
const HEAD = [
  "Project", "Type", "BUA", "Phase", "Plan", "Years", "Discount %",
  "Down Payment Type", "Down Payment 1 %", "Down Payment 2 %", "DP2 Due (months)",
  "Installments Count", "Installment Every (months)", "First Installment Due (months)",
  "Maintenance %", "Delivery (months)", "First Installment Maintenance Due (months)",
  "Ending Installment Maintenance Due (months)", "Maintenance Every (months)",
  "Maintenance Basis", "Status",
];
// A valid cash plan; `dp1` lets a test break the row (over 100%).
const row = (
  project: string, type: string, label: string, status: unknown,
  { dp1 = 100, discount = 0 } = {},
): unknown[] => [
  project, type, "", "", label, 0, discount, "Original Price", dp1, 0, 0,
  0, 0, 0, 0, 0, 0, 0, 0, "Original Price", status,
];
async function loadSheet(
  aoa: unknown[][],
  edit?: (ws: import("xlsx").WorkSheet) => void,
  patchXml?: (xml: string) => string, // edit the written sheet XML itself
) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  edit?.(ws);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Projects");
  let out: ArrayBuffer | Buffer = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  if (patchXml) {
    // CFB (SheetJS' zip reader) is only reachable through the CJS default export.
    const { CFB } = (XLSX as unknown as { default: typeof XLSX }).default;
    const zip = CFB.read(Buffer.from(out), { type: "buffer" });
    const i = zip.FullPaths.findIndex((p) => p.endsWith("worksheets/sheet1.xml"));
    const file = zip.FileIndex[i];
    file.content = Buffer.from(patchXml(Buffer.from(file.content).toString()));
    file.size = file.content.length;
    out = CFB.write(zip, { fileType: "zip", type: "buffer" }) as Buffer;
  }
  (globalThis as { fetch: unknown }).fetch = async () =>
    new Response(out, { status: 200 });
  return loadProjects();
}
const plansOf = (d: Awaited<ReturnType<typeof loadProjects>>, name: string) =>
  d.projects.find((p) => p.name === name)?.plans.map((p) => p.label) ?? null;

{
  // An Inactive first row still names the project for the blank rows below,
  // and a broken Inactive row raises no warning.
  const d = await loadSheet([
    HEAD,
    row("P1", "", "Cash A", "Inactive"),
    row("", "", "Plan B", "Active"),
    row("", "", "Plan C", ""),
    row("", "", "Broken but off", "Inactive", { dp1: 150 }),
    row("P2", "", "Gone 1", "inactive"),
    row("", "", "Gone 2", "INACTIVE"),
  ]);
  check("status: inactive first row passes its project name down", plansOf(d, "P1"), ["Plan B", "Plan C"]);
  check("status: project with only inactive plans disappears", plansOf(d, "P2"), null);
  check("status: inactive rows raise no warnings", d.errors, []);
}
{
  const cases: [unknown, boolean][] = [
    ["Active", true], [" active ", true], ["ACTIVE", true], [true, true], [1, true],
    ["نشط", true], ["Yes", true],
    ["Inactive", false], ["In-active", false], ["Not Active", false], [false, false],
    [0, false], ["غير نشط", false], ["No", false],
  ];
  const d = await loadSheet([
    HEAD,
    ...cases.map(([s], i) => row(i === 0 ? "V" : "", "", `plan ${i}`, s)),
  ]);
  const shown = new Set(plansOf(d, "V") ?? []);
  for (const [i, [s, want]] of cases.entries()) {
    check(`status: ${JSON.stringify(s)} → ${want ? "shown" : "hidden"}`, shown.has(`plan ${i}`), want);
  }
  check("status: known spellings raise no warnings", d.errors, []);
}
{
  const d = await loadSheet([
    HEAD,
    row("U", "", "Typo", "Actve"),
    row("", "", "Fine", "Active"),
  ]);
  check("status: unknown value hides the plan", plansOf(d, "U"), ["Fine"]);
  check(
    "status: unknown value is reported",
    d.errors.length === 1 && d.errors[0].startsWith('Row 2: "Status"'),
    true,
  );
}
{
  const d = await loadSheet(
    [HEAD, row("E", "", "Err", "x"), row("", "", "Ok", "")],
    (ws) => { ws["U2"] = { t: "e", v: 0x2a, w: "#N/A" }; },
  );
  check("status: error cell hides the plan", plansOf(d, "E"), ["Ok"]);
  check("status: error cell is reported", d.errors.length, 1);
}
{
  // Status filled down past the last plan must not create "Plan is empty" rows.
  const blank = Array(HEAD.length - 1).fill("");
  const d = await loadSheet([
    HEAD,
    row("F", "", "Only", "Active"),
    [...blank, "Active"],
    [...blank, "Inactive"],
  ]);
  check("status: status-only rows are ignored", [plansOf(d, "F"), d.errors], [["Only"], []]);
}
{
  // The retired copy of a plan doesn't collide with its live replacement.
  const d = await loadSheet([
    HEAD,
    row("D", "RTM", "Same", "Inactive", { discount: 5 }),
    row("", "RTM", "Same", "Active", { discount: 10 }),
  ]);
  const p = d.projects.find((x) => x.name === "D")?.plans ?? [];
  check("status: inactive twin is not a duplicate", [p.length, p[0]?.discountPct, d.errors], [1, 0.1, []]);
}
{
  // A blank-Type row that is Inactive can't trigger the unreachable-Type warning.
  const d = await loadSheet([
    HEAD,
    row("T", "RTM", "Typed", "Active"),
    row("", "", "Untyped", "Inactive"),
  ]);
  check("status: inactive blank-Type row raises no warning", d.errors, []);
}
{
  const d = await loadSheet(
    [HEAD, row("M", "", "One", "Inactive"), row("", "", "Two", ""), row("", "", "Three", "")],
    (ws) => { (ws["!merges"] ??= []).push(XLSX.utils.decode_range("U2:U3")); },
  );
  check("status: merged Inactive cell hides every row it spans", plansOf(d, "M"), ["Three"]);
}
{
  const head = [...HEAD];
  head[head.length - 1] = "Plan Status";
  const d = await loadSheet([head, row("H", "", "Off", "Inactive"), row("", "", "On", "Active")]);
  check('status: "Plan Status" header works too', [plansOf(d, "H"), d.errors], [["On"], []]);
}
{
  // Whatever the header says, a column of Active / Inactive words is the
  // Status column — Inactive plans must not stay up over a header wording.
  for (const header of ["Payment Status", "Active/Inactive", "Status (Active/Inactive)", "الحالة", "", null]) {
    const head: unknown[] = [...HEAD];
    head[head.length - 1] = header;
    const d = await loadSheet(
      [head, row("W", "", "Off", "Inactive"), row("", "", "On", "Active"), row("", "", "Blank", "")],
      (ws) => { if (header === null) delete ws["U1"]; },
    );
    check(`status: header ${JSON.stringify(header)} still hides Inactive`, [plansOf(d, "W"), d.errors], [["On", "Blank"], []]);
  }
}
{
  // ...but a column that merely mentions status, or holds 1/0, is left alone.
  const head = [...HEAD];
  head[head.length - 1] = "Unit Status";
  const d = await loadSheet([head, row("K", "", "A", "RTM"), row("", "", "B", "Primary")]);
  check('status: "Unit Status" of RTM/Primary is not a Status column', plansOf(d, "K"), ["A", "B"]);
  head[head.length - 1] = "Is Featured";
  const d2 = await loadSheet([head, row("K", "", "A", 1), row("", "", "B", 0)]);
  check("status: a 1/0 column under another header is not a Status column", plansOf(d2, "K"), ["A", "B"]);
}
{
  const cases: [string, boolean][] = [
    ["نشطة", true], ["نشطه", true], ["مفعّلة", true], ["نعم", true],
    ["غير نشطة", false], ["غير مفعلة", false], ["ملغى", false], ["لا", false],
  ];
  const d = await loadSheet([HEAD, ...cases.map(([s], i) => row(i === 0 ? "AR" : "", "", `p${i}`, s))]);
  const shown = new Set(plansOf(d, "AR") ?? []);
  for (const [i, [s, want]] of cases.entries()) {
    check(`status: ${s} → ${want ? "shown" : "hidden"}`, shown.has(`p${i}`), want);
  }
  check("status: Arabic spellings raise no warnings", d.errors, []);
}
{
  // Google's own #ERROR! isn't an Excel error code, so SheetJS reads it as an
  // error cell with no value — it must still hide the plan.
  const d = await loadSheet(
    [HEAD, row("X", "", "Err", "x"), row("", "", "Ok", "")],
    (ws) => { ws["U2"] = { t: "e", v: 0x2a, w: "#N/A" }; },
    (xml) => xml.replace("<v>#N/A</v>", "<v>#ERROR!</v>"),
  );
  check("status: undecoded error cell hides the plan", plansOf(d, "X"), ["Ok"]);
}
{
  const d = await loadSheet([HEAD.slice(0, -1), row("N", "", "Kept", "").slice(0, -1)]);
  check("status: no Status column → every plan shown", [plansOf(d, "N"), d.errors], [["Kept"], []]);
}
{
  const d = await loadSheet([HEAD, row("A", "", "x", "Inactive"), row("", "", "y", "Inactive")]);
  check(
    "status: all inactive → empty list, not the sample projects",
    [d.source, d.projects.length, d.notice !== null, d.errors],
    ["excel", 0, true, []],
  );
}

console.log(
  failures === 0 ? "\nALL CHECKS PASSED" : `\n${failures} CHECK(S) FAILED`,
);
process.exit(failures === 0 ? 0 : 1);
