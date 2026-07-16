// Loads projects & payment plans from public/projects.xlsx (sheet "Projects",
// optional sheet "Settings"). Falls back to the samples in projects.ts when the
// file is missing or unusable. All % columns are entered in percent units
// (10 = 10%); cells formatted as % in Excel are detected and handled too.

import type { Plan, PriceBasis, Project } from "./projects";
import {
  PROJECTS as SAMPLE_PROJECTS,
  CURRENCY as DEFAULT_CURRENCY,
  COMPANY_NAME as DEFAULT_COMPANY,
} from "./projects";

export interface LoadResult {
  projects: Project[];
  companyName: string;
  currency: string;
  source: "excel" | "samples";
  notice: string | null; // why we fell back to samples, if we did
  errors: string[]; // problems in the Excel (bad rows are skipped, bad headers ignored)
}

type PlanNumberField =
  | "buaRate"
  | "years"
  | "discountPct"
  | "downPayment1Pct"
  | "downPayment2Pct"
  | "monthsToDownPayment2"
  | "installmentsCount"
  | "installmentEveryMonths"
  | "monthsToFirstInstallment"
  | "maintenancePct"
  | "maintenanceFirstDueMonths"
  | "maintenanceEndDueMonths"
  | "maintenanceEveryMonths"
  | "deliveryMonths";

type PlanTextField = "name" | "label" | "type" | "phase" | "dpBasis" | "maintenanceBasis";

interface ColumnSpec {
  title: string; // canonical header, used in messages
  field: PlanNumberField | PlanTextField;
  kind: "text" | "number";
  percent?: boolean; // entered in percent units (10 = 10%), stored as fraction
  integer?: boolean; // counts and month offsets must be whole numbers
  min?: number;
  max?: number;
}

// Keyed by normalized header text — see normalizeHeader().
const COLUMNS: Record<string, ColumnSpec> = {
  "project": { title: "Project", field: "name", kind: "text" },
  "type": { title: "Type", field: "type", kind: "text" },
  "bua": { title: "BUA", field: "buaRate", kind: "number", min: 0, max: 10_000_000 },
  "phase": { title: "Phase", field: "phase", kind: "text" },
  "plan": { title: "Plan", field: "label", kind: "text" },
  "years": { title: "Years", field: "years", kind: "number", min: 0, max: 50 },
  "discount %": { title: "Discount %", field: "discountPct", kind: "number", percent: true, min: -100, max: 100 },
  "down payment type": { title: "Down Payment Type", field: "dpBasis", kind: "text" },
  "down payment 1 %": { title: "Down Payment 1 %", field: "downPayment1Pct", kind: "number", percent: true, min: 0, max: 100 },
  "down payment 2 %": { title: "Down Payment 2 %", field: "downPayment2Pct", kind: "number", percent: true, min: 0, max: 100 },
  "dp2 due months": { title: "DP2 Due (months)", field: "monthsToDownPayment2", kind: "number", integer: true, min: 0, max: 600 },
  "installments count": { title: "Installments Count", field: "installmentsCount", kind: "number", integer: true, min: 0, max: 1000 },
  "installment every months": { title: "Installment Every (months)", field: "installmentEveryMonths", kind: "number", integer: true, min: 0, max: 120 },
  "first installment due months": { title: "First Installment Due (months)", field: "monthsToFirstInstallment", kind: "number", integer: true, min: 0, max: 600 },
  "maintenance %": { title: "Maintenance %", field: "maintenancePct", kind: "number", percent: true, min: 0, max: 100 },
  "delivery months": { title: "Delivery (months)", field: "deliveryMonths", kind: "number", integer: true, min: 0, max: 600 },
  "first installment maintenance due months": { title: "First Installment Maintenance Due (months)", field: "maintenanceFirstDueMonths", kind: "number", integer: true, min: 0, max: 600 },
  "ending installment maintenance due months": { title: "Ending Installment Maintenance Due (months)", field: "maintenanceEndDueMonths", kind: "number", integer: true, min: 0, max: 600 },
  "maintenance every months": { title: "Maintenance Every (months)", field: "maintenanceEveryMonths", kind: "number", integer: true, min: 0, max: 120 },
  "maintenance basis": { title: "Maintenance Basis", field: "maintenanceBasis", kind: "text" },
};

// The sheet may reuse the header "Installment Every (months)" for the
// maintenance spacing column (the second occurrence, next to the maintenance
// columns). This spec is assigned to that second occurrence.
const MAINT_EVERY_SPEC = COLUMNS["maintenance every months"];

const EPS = 1e-9;

function normalizeHeader(h: string): string {
  return h.toLowerCase().replace(/[^a-z0-9%]+/g, " ").trim();
}

function slug(s: string): string {
  return (
    s
      .toLowerCase()
      .replace(/[^a-z0-9؀-ۿ]+/g, "-")
      .replace(/^-+|-+$/g, "") || "x"
  );
}

interface SheetCell {
  t?: string; // cell type — "e" marks an Excel error cell (#N/A, #REF!, ...)
  v?: unknown;
  w?: string; // formatted text, carries "#DIV/0!" etc. for error cells
  z?: string; // number format (needs cellNF: true when reading)
}

// True only for genuine Excel percent formats. A '%' inside quotes ("0.00\"%\"")
// or after a backslash is a display-only literal and does NOT scale the value.
function hasPercentFormat(z: string): boolean {
  return z.replace(/"[^"]*"/g, "").replace(/\\./g, "").includes("%");
}

// Reads one numeric cell. Empty cells default to 0. Returns null (and records
// an error) when the value is not a usable number or is out of range.
function readNumber(
  cell: SheetCell | undefined,
  spec: ColumnSpec,
  rowNo: number,
  errors: string[],
): number | null {
  if (cell != null && cell.t === "e") {
    errors.push(
      `Row ${rowNo}: "${spec.title}" contains an Excel error (${typeof cell.w === "string" ? cell.w : "#N/A"}) — fix the formula`,
    );
    return null;
  }
  if (cell == null || cell.v == null || cell.v === "") return 0;
  let n: number;
  const v = cell.v;
  if (typeof v === "number") {
    n = v;
    // A cell formatted as % in Excel stores 0.1 for "10%" — convert to percent
    // units so the shared /100 below lands on the right fraction.
    if (spec.percent && typeof cell.z === "string" && hasPercentFormat(cell.z)) {
      n = n * 100;
    }
  } else if (typeof v === "string") {
    let s = v.trim().replace(/%$/, "").trim();
    if (s === "") return 0;
    if (s.includes(",")) {
      // Only strip commas that are unambiguous thousands separators; "7,5"
      // (decimal comma) must NOT silently become 75.
      if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, "");
      else {
        errors.push(
          `Row ${rowNo}: "${spec.title}" contains a comma ("${v}") — use a dot for decimals`,
        );
        return null;
      }
    }
    n = Number(s);
    if (!Number.isFinite(n)) {
      errors.push(`Row ${rowNo}: "${spec.title}" is not a number (got "${v}")`);
      return null;
    }
  } else {
    errors.push(`Row ${rowNo}: "${spec.title}" has an unsupported value`);
    return null;
  }
  if (
    (spec.min != null && n < spec.min) ||
    (spec.max != null && n > spec.max)
  ) {
    errors.push(
      `Row ${rowNo}: "${spec.title}" must be between ${spec.min} and ${spec.max} (got ${n})`,
    );
    return null;
  }
  if (spec.integer && !Number.isInteger(n)) {
    errors.push(
      `Row ${rowNo}: "${spec.title}" must be a whole number (got ${n})`,
    );
    return null;
  }
  return n;
}

// "Original Price" / "Selling price" → internal basis. Empty → default.
function parseBasis(
  raw: string,
  spec: ColumnSpec,
  rowNo: number,
  errors: string[],
): PriceBasis | null {
  if (!raw) return "original";
  if (/original/i.test(raw)) return "original";
  if (/selling|final/i.test(raw)) return "selling";
  errors.push(
    `Row ${rowNo}: "${spec.title}" must be "Original Price" or "Selling price" (got "${raw}")`,
  );
  return null;
}

// Last-resort result for callers when even loadProjects() itself rejects
// (e.g. the xlsx code chunk failed to download).
export function loadFailedResult(): LoadResult {
  return {
    projects: SAMPLE_PROJECTS,
    companyName: DEFAULT_COMPANY,
    currency: DEFAULT_CURRENCY,
    source: "samples",
    notice:
      "Could not load the Excel reader — showing the built-in sample projects. Refresh to retry.",
    errors: [],
  };
}

export async function loadProjects(): Promise<LoadResult> {
  const fallback = (notice: string, errors: string[] = []): LoadResult => ({
    projects: SAMPLE_PROJECTS,
    companyName: DEFAULT_COMPANY,
    currency: DEFAULT_CURRENCY,
    source: "samples",
    notice,
    errors,
  });

  let buf: ArrayBuffer;
  try {
    // Prefix with the GitHub Pages sub-path when deployed (empty locally).
    const base = process.env.NEXT_PUBLIC_BASE_PATH || "";
    const res = await fetch(`${base}/projects.xlsx`, { cache: "no-store" });
    if (!res.ok) {
      return fallback(
        "public/projects.xlsx not found — showing the built-in sample projects.",
      );
    }
    buf = await res.arrayBuffer();
  } catch {
    return fallback(
      "Could not fetch public/projects.xlsx — showing the built-in sample projects.",
    );
  }

  // Nothing past this point may take the page down: any unexpected parsing
  // failure degrades to the sample data instead.
  try {
    return await parseWorkbook(buf, fallback);
  } catch {
    return fallback(
      "Unexpected error while reading projects.xlsx — showing the built-in sample projects.",
    );
  }
}

async function parseWorkbook(
  buf: ArrayBuffer,
  fallback: (notice: string, errors?: string[]) => LoadResult,
): Promise<LoadResult> {
  const XLSX = await import("xlsx");
  let wb: import("xlsx").WorkBook;
  try {
    wb = XLSX.read(buf, { type: "array", cellNF: true });
  } catch {
    return fallback(
      "public/projects.xlsx could not be read as an Excel file — showing the built-in sample projects.",
    );
  }

  const ws = wb.Sheets["Projects"];
  if (!ws || !ws["!ref"]) {
    return fallback(
      'projects.xlsx has no "Projects" sheet — showing the built-in sample projects.',
    );
  }

  const errors: string[] = [];

  // Vertically merged cells only store a value in the top-left anchor; map
  // every covered address to its anchor so merged Project/Type cells apply to
  // all rows they span.
  const mergeAnchor = new Map<string, string>();
  for (const m of ws["!merges"] ?? []) {
    const anchor = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
    for (let r = m.s.r; r <= m.e.r; r++) {
      for (let c = m.s.c; c <= m.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (addr !== anchor) mergeAnchor.set(addr, anchor);
      }
    }
  }
  const directCell = (r: number, c: number): SheetCell | undefined =>
    ws[XLSX.utils.encode_cell({ r, c })] as SheetCell | undefined;
  const getCell = (r: number, c: number): SheetCell | undefined => {
    const direct = directCell(r, c);
    if (direct && direct.v != null && direct.v !== "") return direct;
    const a = mergeAnchor.get(XLSX.utils.encode_cell({ r, c }));
    return a ? (ws[a] as SheetCell | undefined) : direct;
  };

  // Map header row -> known columns; surface anything we don't recognize so a
  // typo'd header can't silently zero a whole column. A second occurrence of
  // "Installment Every (months)" is the maintenance spacing column.
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const colSpec = new Map<number, ColumnSpec>();
  const unknownHeaders: string[] = [];
  const seenFields = new Set<string>();
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = directCell(range.s.r, c);
    if (cell == null || cell.v == null || String(cell.v).trim() === "") continue;
    const norm = normalizeHeader(String(cell.v));
    let spec = COLUMNS[norm];
    if (spec && seenFields.has(spec.field)) {
      if (spec.field === "installmentEveryMonths" && !seenFields.has("maintenanceEveryMonths")) {
        spec = MAINT_EVERY_SPEC;
      } else {
        errors.push(
          `Duplicate column "${String(cell.v).trim()}" ignored — remove or rename one of them`,
        );
        continue;
      }
    }
    if (spec) {
      colSpec.set(c, spec);
      seenFields.add(spec.field);
    } else {
      unknownHeaders.push(String(cell.v).trim());
    }
  }
  if (!seenFields.has("name") || !seenFields.has("label")) {
    return fallback(
      'The "Projects" sheet is missing the "Project" or "Plan" column — showing the built-in sample projects.',
    );
  }
  if (unknownHeaders.length > 0) {
    errors.push(
      `Unrecognized column(s) ignored: ${unknownHeaders.map((h) => `"${h}"`).join(", ")} — check the spelling against the template headers`,
    );
  }
  const missingColumns = Object.values(COLUMNS).filter(
    (s) => s.kind === "number" && !seenFields.has(s.field),
  );
  if (missingColumns.length > 0) {
    errors.push(
      `Column(s) not found (treated as 0 for all rows): ${missingColumns.map((s) => `"${s.title}"`).join(", ")}`,
    );
  }

  const projectsByName = new Map<string, Project>();
  let lastProjectName = "";
  // Tracked for the reachability warnings after the loop.
  const planRows: {
    project: string;
    type: string | null;
    phase: string | null;
    rowNo: number;
  }[] = [];

  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const rowNo = r + 1; // human/Excel row number

    // A row counts as present only if it has DIRECT content; merge-inherited
    // values alone must not conjure phantom rows.
    let hasAny = false;
    for (const c of colSpec.keys()) {
      const cell = directCell(r, c);
      if (cell && cell.v != null && cell.v !== "") {
        hasAny = true;
        break;
      }
    }
    if (!hasAny) continue;

    const texts: Partial<Record<PlanTextField, string>> = {};
    const nums: Partial<Record<PlanNumberField, number>> = {};
    let rowBroken = false;

    for (const [c, spec] of colSpec) {
      const cell = getCell(r, c);
      if (spec.kind === "text") {
        if (cell != null && cell.t === "e") {
          errors.push(
            `Row ${rowNo}: "${spec.title}" contains an Excel error (${typeof cell.w === "string" ? cell.w : "#N/A"}) — fix the formula`,
          );
          rowBroken = true;
          continue;
        }
        texts[spec.field as PlanTextField] =
          cell?.v != null ? String(cell.v).trim() : "";
      } else {
        const n = readNumber(cell, spec, rowNo, errors);
        if (n === null) rowBroken = true;
        else nums[spec.field as PlanNumberField] = n;
      }
    }

    // Blank Project cell = same project as the row above (handles merged cells too).
    let name = texts.name ?? "";
    if (!name) name = lastProjectName;
    if (!name) {
      errors.push(
        `Row ${rowNo}: "Project" is empty and there is no project above to inherit`,
      );
      continue;
    }
    lastProjectName = name;

    const label = texts.label ?? "";
    if (!label) {
      errors.push(`Row ${rowNo}: "Plan" is empty`);
      continue;
    }

    const dpBasis = parseBasis(
      texts.dpBasis ?? "",
      COLUMNS["down payment type"],
      rowNo,
      errors,
    );
    const maintenanceBasis = parseBasis(
      texts.maintenanceBasis ?? "",
      COLUMNS["maintenance basis"],
      rowNo,
      errors,
    );
    if (dpBasis === null || maintenanceBasis === null || rowBroken) {
      continue; // cell-level errors already recorded
    }

    const dp1 = nums.downPayment1Pct ?? 0;
    const dp2 = nums.downPayment2Pct ?? 0;
    const count = nums.installmentsCount ?? 0;
    const every = nums.installmentEveryMonths ?? 0;

    if (dp1 + dp2 > 100 + EPS) {
      errors.push(
        `Row ${rowNo}: down payments add up to ${dp1 + dp2}% — must not exceed 100%`,
      );
      continue;
    }
    if (count === 0 && dp1 + dp2 < 100 - EPS) {
      errors.push(
        `Row ${rowNo}: down payments cover only ${dp1 + dp2}% and there are no installments — the remaining ${100 - dp1 - dp2}% has no due date`,
      );
      continue;
    }
    if (count > 0 && every < 1) {
      errors.push(
        `Row ${rowNo}: "Installment Every (months)" must be at least 1 when there are installments`,
      );
      continue;
    }

    // Maintenance schedule sanity.
    const maintPct = nums.maintenancePct ?? 0;
    const mFirst = nums.maintenanceFirstDueMonths ?? 0;
    const mEnd = nums.maintenanceEndDueMonths ?? 0;
    const mEvery = nums.maintenanceEveryMonths ?? 0;
    if (maintPct > 0) {
      if (mEnd < mFirst) {
        errors.push(
          `Row ${rowNo}: "Ending Installment Maintenance Due" (${mEnd}) is before "First Installment Maintenance Due" (${mFirst})`,
        );
        continue;
      }
      if (mEnd > mFirst && mEvery < 1) {
        errors.push(
          `Row ${rowNo}: maintenance runs from month ${mFirst} to ${mEnd} but its "Every (months)" is 0`,
        );
        continue;
      }
      // A final gap shorter than "Every (months)" is intentional and harmless:
      // the schedule always lands the last payment exactly on ${mEnd} (e.g. GAIA
      // runs 6 → 24 → 30, an 18-month cadence ending with a 6-month gap). This
      // can never signal a real misconfiguration, so we don't warn about it.
    }

    let project = projectsByName.get(name);
    if (!project) {
      project = { id: slug(name), name, plans: [] };
      projectsByName.set(name, project);
    }
    const type = texts.type || null;
    // A Phase cell may list several phases ("1,2,3,4" or "2,3") — expand each
    // into its own dropdown option so the agent picks a single phase. Splits on
    // comma (Latin or Arabic) and semicolon.
    const phaseRaw = texts.phase || "";
    const phaseList: (string | null)[] = phaseRaw
      ? [
          ...new Set(
            phaseRaw
              .split(/[,،;]/)
              .map((s) => s.trim())
              .filter(Boolean),
          ),
        ]
      : [null];

    for (const phase of phaseList) {
      if (
        project.plans.some(
          (pl) => pl.label === label && pl.type === type && pl.phase === phase,
        )
      ) {
        errors.push(
          `Row ${rowNo}: duplicate plan "${label}" (Type ${type ?? "-"}, Phase ${phase ?? "-"}) in project "${name}"`,
        );
        continue;
      }

      const plan: Plan = {
        id: slug([type ?? "", phase ?? "", label].filter(Boolean).join("-")),
        label,
        type,
        phase,
        buaRate: nums.buaRate ?? 0,
        years: nums.years ?? 0,
        discountPct: (nums.discountPct ?? 0) / 100,
        dpBasis,
        downPayment1Pct: dp1 / 100,
        downPayment2Pct: dp2 / 100,
        monthsToDownPayment2: nums.monthsToDownPayment2 ?? 0,
        installmentsCount: count,
        installmentEveryMonths: every || 3,
        monthsToFirstInstallment: nums.monthsToFirstInstallment ?? 0,
        maintenancePct: maintPct / 100,
        maintenanceBasis,
        maintenanceFirstDueMonths: mFirst,
        maintenanceEndDueMonths: mEnd,
        maintenanceEveryMonths: mEvery,
        deliveryMonths: nums.deliveryMonths ?? 0,
      };
      project.plans.push(plan);
      planRows.push({ project: name, type, phase, rowNo });
    }
  }

  // Reachability: Unit Status (Type) is a hard filter, so a blank Type among
  // filled siblings can never be selected — warn about that. A blank Phase is
  // fine: the cascade now picks the plan first, and a plan with no phase simply
  // skips the Phase step, so no warning is needed.
  const byProject = new Map<string, typeof planRows>();
  for (const pr of planRows) {
    const list = byProject.get(pr.project) ?? [];
    list.push(pr);
    byProject.set(pr.project, list);
  }
  for (const [projName, list] of byProject) {
    const blankType = list.filter((x) => x.type === null);
    if (blankType.length > 0 && blankType.length < list.length) {
      errors.push(
        `Project "${projName}": row(s) ${blankType.map((x) => x.rowNo).join(", ")} have no "Type" while other rows do — they won't appear in the page until you fill the Type`,
      );
    }
  }

  const projects = [...projectsByName.values()].filter(
    (p) => p.plans.length > 0,
  );
  if (projects.length === 0) {
    return fallback(
      "projects.xlsx contains no valid plan rows — showing the built-in sample projects.",
      errors,
    );
  }

  // De-duplicate ids that different names happen to slugify into.
  const seenIds = new Set<string>();
  for (const p of projects) {
    let id = p.id;
    for (let i = 2; seenIds.has(id); i++) id = `${p.id}-${i}`;
    p.id = id;
    seenIds.add(id);
    const seenPlanIds = new Set<string>();
    for (const pl of p.plans) {
      let pid = pl.id;
      for (let i = 2; seenPlanIds.has(pid); i++) pid = `${pl.id}-${i}`;
      pl.id = pid;
      seenPlanIds.add(pid);
    }
  }

  // Optional Settings sheet: "Company Name" / "Currency" key-value rows.
  let companyName = DEFAULT_COMPANY;
  let currency = DEFAULT_CURRENCY;
  const settings = wb.Sheets["Settings"];
  if (settings && settings["!ref"]) {
    const sRange = XLSX.utils.decode_range(settings["!ref"]);
    for (let r = sRange.s.r; r <= sRange.e.r; r++) {
      const keyCell = settings[XLSX.utils.encode_cell({ r, c: 0 })] as
        | SheetCell
        | undefined;
      const valCell = settings[XLSX.utils.encode_cell({ r, c: 1 })] as
        | SheetCell
        | undefined;
      if (keyCell?.v == null || valCell?.v == null) continue;
      const key = String(keyCell.v).toLowerCase();
      const val = String(valCell.v).trim();
      if (!val) continue;
      if (key.includes("company")) companyName = val;
      else if (key.includes("currenc")) currency = val;
    }
  }

  return {
    projects,
    companyName,
    currency,
    source: "excel",
    notice: null,
    errors,
  };
}
