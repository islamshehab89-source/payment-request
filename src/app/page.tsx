"use client";

import { useEffect, useMemo, useState } from "react";
import {
  loadProjects,
  loadFailedResult,
  type LoadResult,
} from "@/lib/loadProjects";
import { computeSchedule, type ScheduleResult } from "@/lib/schedule";
import {
  money,
  percent,
  formatDate,
  setCurrency,
  monthsToYears,
} from "@/lib/format";

// Static assets sit under the GitHub Pages sub-path when deployed (empty locally).
const ASSET_BASE = process.env.NEXT_PUBLIC_BASE_PATH || "";
const LOGO_SRC = `${ASSET_BASE}/logo.webp`;
// transparent black wordmark — used faint (watermark) and small (running header)
const LOGO_DARK_SRC = `${ASSET_BASE}/logo-dark.png`;

function distinct(values: (string | null)[]): string[] {
  const out: string[] = [];
  for (const v of values) if (v && !out.includes(v)) out.push(v);
  return out;
}

export default function Page() {
  const [data, setData] = useState<LoadResult | null>(null);
  const [logoOk, setLogoOk] = useState(true);

  // Contract date is always today and cannot be changed by the user.
  const [today] = useState(() => new Date());

  const [projectId, setProjectId] = useState("");
  const [unitStatus, setUnitStatus] = useState("");
  const [planLabel, setPlanLabel] = useState(""); // chosen payment-plan label
  const [phase, setPhase] = useState(""); // chosen phase (when the plan has phases)

  const [originalPrice, setOriginalPrice] = useState("");
  const [unitType, setUnitType] = useState(""); // e.g. "3BR" — informational
  const [unitArea, setUnitArea] = useState("");
  const [outdoorArea, setOutdoorArea] = useState(""); // informational

  useEffect(() => {
    let alive = true;
    loadProjects()
      .then((r) => {
        if (!alive) return;
        setCurrency(r.currency); // must happen before any money() renders
        setData(r);
      })
      .catch(() => {
        if (alive) setData(loadFailedResult());
      });
    return () => {
      alive = false;
    };
  }, []);

  const projects = data?.projects ?? [];
  const project = projects.find((p) => p.id === projectId);

  // Cascade: Project → Unit Status (Type) → Phase → Plan. A level only shows
  // when the sheet has values for it; blank cells never appear as options.
  const typeOptions = project ? distinct(project.plans.map((p) => p.type)) : [];
  const typedPlans = project
    ? project.plans.filter(
        (p) => typeOptions.length === 0 || p.type === unitStatus,
      )
    : [];
  // Cascade order: Project → Unit Status → Payment Plan → Phase. The Payment
  // Plan dropdown lists each distinct plan label once; the Phase dropdown then
  // offers the phases that label is available in.
  const labelOptions = (() => {
    const seen = new Set<string>();
    const out: typeof typedPlans = [];
    for (const p of typedPlans)
      if (!seen.has(p.label)) {
        seen.add(p.label);
        out.push(p);
      }
    return out;
  })();
  const labelPlans = typedPlans.filter((p) => p.label === planLabel);
  // Phases sorted ascending (numerically when they're numbers, e.g. 1,2,3,4).
  const phaseOptions = distinct(labelPlans.map((p) => p.phase)).sort((a, b) => {
    const na = Number(a);
    const nb = Number(b);
    if (!Number.isNaN(na) && !Number.isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  });
  const plan =
    phaseOptions.length === 0
      ? labelPlans[0]
      : labelPlans.find((p) => p.phase === phase);

  const price = parseFloat(originalPrice);
  const area = parseFloat(unitArea) || 0;
  const outdoor = parseFloat(outdoorArea) || 0;
  const needArea = !!plan && plan.buaRate > 0; // finished units require the area

  const result: ScheduleResult | null = useMemo(() => {
    if (!plan || !price || price <= 0) return null;
    if (plan.buaRate > 0 && !(area > 0)) return null; // finishing needs the area
    return computeSchedule(plan, {
      originalPrice: price,
      unitArea: area,
      contractDate: today,
    });
  }, [plan, price, area, today]);

  // The schedule "fills" a printed page once it has enough rows (installments +
  // down payments + one divider per year). When it does, no logo is added to
  // that page; short schedules (e.g. cash) leave room and get the logo.
  const scheduleFillsPage = (() => {
    if (!result) return true;
    const years = new Set(
      result.rows.filter((r) => r.year > 0).map((r) => r.year),
    ).size;
    const printRows = result.rows.length + years; // body rows + year dividers
    return printRows > 30;
  })();

  // Switching project = a different unit: clear the unit-specific inputs too.
  // Switching type/phase/plan keeps them (same unit, comparing plans).
  function onProjectChange(id: string) {
    setProjectId(id);
    setUnitStatus("");
    setPlanLabel("");
    setPhase("");
    setOriginalPrice("");
    setUnitType("");
    setUnitArea("");
    setOutdoorArea("");
  }
  function onStatusChange(v: string) {
    setUnitStatus(v);
    setPlanLabel("");
    setPhase("");
  }
  function onPlanLabelChange(v: string) {
    setPlanLabel(v);
    setPhase(""); // re-pick the phase for the new plan
  }
  function reset() {
    setProjectId("");
    setUnitStatus("");
    setPlanLabel("");
    setPhase("");
    setOriginalPrice("");
    setUnitType("");
    setUnitArea("");
    setOutdoorArea("");
  }

  const companyName = data?.companyName ?? "";

  return (
    <div className="page">
      <header className="app-header">
        <div className="brand">
          {logoOk && (
            <img
              src={LOGO_SRC}
              alt={companyName || "Logo"}
              className="brand-logo"
              onError={() => setLogoOk(false)}
            />
          )}
          <div className="brand-text">
            <div className="company-name">{companyName || "Company"}</div>
            <div className="doc-type">Payment Request</div>
          </div>
        </div>
      </header>

      {!data && (
        <section className="card">
          <p className="empty">Loading projects…</p>
        </section>
      )}

      {data && (
        <>
          {data.notice && (
            <div className="banner banner-info no-print">{data.notice}</div>
          )}
          {data.errors.length > 0 && (
            <div className="banner banner-warn no-print">
              <strong>
                projects.xlsx has issues (the affected rows/columns were
                skipped) — fix them and refresh:
              </strong>
              <ul>
                {data.errors.map((e, i) => (
                  <li key={i}>{e}</li>
                ))}
              </ul>
            </div>
          )}

          {/* ---------------- form ---------------- */}
          <section className="card no-print">
            <h2>Order details</h2>
            <p className="form-legend">
              Fields marked <span className="req">*</span> are required.
            </p>
            <div className="form-grid">
              <div className="field">
                <label htmlFor="project">
                  Project <span className="req">*</span>
                </label>
                <select
                  id="project"
                  value={projectId}
                  onChange={(e) => onProjectChange(e.target.value)}
                >
                  <option value="">Select a project…</option>
                  {projects.map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.name}
                    </option>
                  ))}
                </select>
              </div>

              {typeOptions.length > 0 && (
                <div className="field">
                  <label htmlFor="status">
                    Payment Status <span className="req">*</span>
                  </label>
                  <select
                    id="status"
                    value={unitStatus}
                    onChange={(e) => onStatusChange(e.target.value)}
                  >
                    <option value="">Select…</option>
                    {typeOptions.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              <div className="field">
                <label htmlFor="plan">
                  Payment Plan <span className="req">*</span>
                </label>
                <select
                  id="plan"
                  value={planLabel}
                  onChange={(e) => onPlanLabelChange(e.target.value)}
                  disabled={labelOptions.length === 0}
                >
                  <option value="">
                    {labelOptions.length === 0
                      ? "Select the options above first"
                      : "Select a plan…"}
                  </option>
                  {labelOptions.map((pl) => (
                    <option key={pl.label} value={pl.label}>
                      {pl.label}
                      {pl.discountPct > 0
                        ? `  ·  ${percent(pl.discountPct)} off`
                        : pl.discountPct < 0
                          ? `  ·  ${percent(-pl.discountPct)} premium`
                          : ""}
                    </option>
                  ))}
                </select>
              </div>

              {phaseOptions.length > 0 && (
                <div className="field">
                  <label htmlFor="phase">
                    Phase <span className="req">*</span>
                  </label>
                  <select
                    id="phase"
                    value={phase}
                    onChange={(e) => setPhase(e.target.value)}
                  >
                    <option value="">Select…</option>
                    {phaseOptions.map((ph) => (
                      <option key={ph} value={ph}>
                        {ph}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              <div className="field">
                <label htmlFor="price">
                  Original Unit Price <span className="req">*</span>
                </label>
                <input
                  id="price"
                  type="number"
                  inputMode="numeric"
                  placeholder="e.g. 1000000"
                  value={originalPrice}
                  onChange={(e) => setOriginalPrice(e.target.value)}
                />
              </div>

              <div className="field">
                <label htmlFor="area">
                  Unit Area (m²)
                  {needArea ? (
                    <span className="req">*</span>
                  ) : (
                    " — optional"
                  )}
                </label>
                <input
                  id="area"
                  type="number"
                  inputMode="decimal"
                  placeholder={needArea ? "e.g. 20" : "optional"}
                  value={unitArea}
                  onChange={(e) => setUnitArea(e.target.value)}
                />
                {needArea && (
                  <span className="hint">
                    Finishing = area × {money(plan!.buaRate)} per m²
                  </span>
                )}
              </div>

              <div className="field">
                <label htmlFor="utype">Unit Type (optional)</label>
                <input
                  id="utype"
                  type="text"
                  placeholder="e.g. 3BR"
                  value={unitType}
                  onChange={(e) => setUnitType(e.target.value)}
                />
              </div>

              <div className="field">
                <label htmlFor="outdoor">Outdoor Area (m², optional)</label>
                <input
                  id="outdoor"
                  type="number"
                  inputMode="decimal"
                  placeholder="0"
                  value={outdoorArea}
                  onChange={(e) => setOutdoorArea(e.target.value)}
                />
              </div>
            </div>
          </section>

          {!result && (
            <section className="card">
              <p className="empty">
                {!plan
                  ? "Choose a project and a payment plan, then enter the original unit price to generate the schedule."
                  : !(price > 0)
                    ? "Enter the original unit price to generate the schedule."
                    : needArea && !(area > 0)
                      ? "Enter the unit area to compute the finishing cost and generate the schedule."
                      : ""}
              </p>
            </section>
          )}

          {result && project && plan && (
            <>
              {result.unscheduled !== 0 && (
                <div className="banner banner-warn no-print">
                  {result.unscheduled > 0
                    ? `This plan leaves ${money(result.unscheduled)} of the final price without a due date`
                    : `This plan's payments exceed the final price by ${money(-result.unscheduled)}`}
                  {" — check the plan's Down Payment Type and Installments Count in the Excel sheet."}
                </div>
              )}

              {/* ---------------- actions ---------------- */}
              <div className="actions no-print">
                <button
                  className="btn btn-primary"
                  onClick={() => window.print()}
                >
                  <span className="only-desktop">Print / Save PDF</span>
                  <span className="only-mobile">Save as PDF</span>
                </button>
                <button className="btn btn-ghost" onClick={reset}>
                  Reset
                </button>
              </div>

              {/* ===== PRINT PAGE 1: Unit Information ===== */}
              <div className="print-page">
              {/* letterhead — PDF only */}
              <header className="letterhead print-only">
                <div className="lh-top">
                  {logoOk && (
                    <img src={LOGO_SRC} alt="" className="brand-logo" />
                  )}
                </div>
                <div className="lh-title">
                  <h1>Payment Request</h1>
                  <div className="lh-issued">Issued {formatDate(today)}</div>
                </div>
              </header>

              {/* ---------------- unit information ---------------- */}
              <section className="card unit-card">
                <h2>Unit Information</h2>
                <div className="info-cols">
                  <div className="info-col">
                    <InfoRow k="Project" v={project.name} />
                    {unitType && <InfoRow k="Unit Type" v={unitType} />}
                    {area > 0 && <InfoRow k="Unit Area" v={`${area} m²`} />}
                    {outdoor > 0 && (
                      <InfoRow k="Outdoor Area" v={`${outdoor} m²`} />
                    )}
                    {plan.type && <InfoRow k="Payment Status" v={plan.type} />}
                    {plan.phase && <InfoRow k="Phase" v={plan.phase} />}
                    <InfoRow k="Payment Plan" v={plan.label} />
                    <InfoRow k="Date" v={formatDate(today)} />
                    {plan.deliveryMonths > 0 && (
                      <InfoRow
                        k="Delivery"
                        v={monthsToYears(plan.deliveryMonths)}
                      />
                    )}
                  </div>
                  <div className="info-col">
                    <InfoRow
                      k="Original Unit Price"
                      v={money(result.originalPrice)}
                    />
                    <InfoRow
                      k={
                        plan.discountPct < 0
                          ? `Premium (${percent(-plan.discountPct)})`
                          : `Discount (${percent(plan.discountPct)})`
                      }
                      v={
                        result.discount > 0
                          ? `− ${money(result.discount)}`
                          : result.discount < 0
                            ? `+ ${money(-result.discount)}`
                            : "—"
                      }
                    />
                    {result.finishingCost > 0 && (
                      <>
                        <InfoRow
                          k="Net Unit Price (Without Finishing)"
                          v={money(result.netPrice)}
                        />
                        <InfoRow
                          k={`Finishing Cost (${area} × ${money(plan.buaRate)})`}
                          v={money(result.finishingCost)}
                        />
                      </>
                    )}
                    <InfoRow
                      k="Final Unit Price"
                      v={money(result.finalPrice)}
                      strong
                    />
                  </div>
                </div>
              </section>
              <PageLogo show={logoOk} />
              <footer className="doc-footer print-only">
                <span className="ft-co">
                  {companyName || "Al Ahly Sabbour Developments"}
                </span>
                <span className="ft-disc">
                  — This Payment Request is a price proposal and does not
                  constitute a contract.
                </span>
                <span className="ft-page">Page 1 of 3</span>
              </footer>
              </div>

              {/* ===== PRINT PAGE 2: Payment Schedule ===== */}
              <div className="print-page">
              {/* running header — PDF only */}
              <header className="runhead print-only">
                {logoOk && <img src={LOGO_DARK_SRC} alt="" className="rh-mark" />}
                <span className="rh-doc">Payment Request · {project.name}</span>
              </header>
              {/* ---------------- schedule ---------------- */}
              <section className="card schedule-card mobile-hide">
                <h2>Payment Schedule</h2>
                <div className="table-wrap">
                  <ScheduleTable result={result} />
                </div>
              </section>
              <PageLogo show={logoOk && !scheduleFillsPage} />
              <footer className="doc-footer print-only">
                <span className="ft-co">
                  {companyName || "Al Ahly Sabbour Developments"}
                </span>
                <span className="ft-disc">
                  — Instalment due dates are calculated from the reservation
                  date and are subject to the signed contract.
                </span>
                <span className="ft-page">Page 2 of 3</span>
              </footer>
              </div>

              {/* ===== PRINT PAGE 3: Maintenance + Totals ===== */}
              <div className="print-page print-page-last">
              {/* running header — PDF only */}
              <header className="runhead print-only">
                {logoOk && <img src={LOGO_DARK_SRC} alt="" className="rh-mark" />}
                <span className="rh-doc">Payment Request · {project.name}</span>
              </header>
              {/* ---------------- maintenance ---------------- */}
              {result.maintenanceTotal > 0 && (
                <section className="card mobile-hide">
                  <h2>
                    Maintenance — {percent(plan.maintenancePct)} of{" "}
                    {plan.maintenanceBasis === "original"
                      ? "original price"
                      : "final price"}
                  </h2>
                  <div className="table-wrap">
                    <table className="schedule extras-table">
                      <thead>
                        <tr>
                          <th>Code</th>
                          <th>Description</th>
                          <th>Due date</th>
                          <th>Amount</th>
                        </tr>
                      </thead>
                      <tbody>
                        {result.maintenanceRows.map((m) => (
                          <tr key={m.code}>
                            <td>
                              <span className="code-chip">{m.code}</span>
                            </td>
                            <td>{m.label}</td>
                            <td>{formatDate(m.dueDate)}</td>
                            <td>{money(m.amount)}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr>
                          <td colSpan={3}>Maintenance total</td>
                          <td>{money(result.maintenanceTotal)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                </section>
              )}

              {/* ---------------- totals ---------------- */}
              <section className="card">
                <h2>Totals</h2>
                {/* screen keeps the existing summary grid (unchanged) */}
                <div className="summary-grid no-print">
                  <Cell k="Final Unit Price" v={money(result.finalPrice)} />
                  <Cell k="Maintenance" v={money(result.maintenanceTotal)} />
                  <Cell k="Grand Total" v={money(result.grandTotal)} total />
                </div>
                {/* PDF gets a dotted-leader ledger */}
                <div className="totals print-only">
                  <div className="tot-row">
                    <span className="k">Final Unit Price</span>
                    <span className="lead" />
                    <span className="v">{money(result.finalPrice)}</span>
                  </div>
                  <div className="tot-row">
                    <span className="k">Maintenance</span>
                    <span className="lead" />
                    <span className="v">{money(result.maintenanceTotal)}</span>
                  </div>
                  <div className="tot-grand">
                    <span className="k">Grand Total</span>
                    <span className="v">{money(result.grandTotal)}</span>
                  </div>
                </div>
              </section>
              <PageLogo show={logoOk} />
              <footer className="doc-footer print-only">
                <span className="ft-co">
                  {companyName || "Al Ahly Sabbour Developments"}
                </span>
                <span className="ft-disc">
                  — Thank you for choosing Al Ahly Sabbour. Figures are valid as
                  of the issue date above.
                </span>
                <span className="ft-page">Page 3 of 3</span>
              </footer>
              </div>
            </>
          )}
        </>
      )}
    </div>
  );
}

// A faint logo that fills the empty bottom space of a printed page. It shrinks
// to fit whatever space is left, so on a full page it collapses to nothing.
// Print-only (hidden on screen via CSS).
function PageLogo({ show }: { show: boolean }) {
  if (!show) return null;
  return (
    <div className="page-filler" aria-hidden="true">
      <img src={LOGO_DARK_SRC} alt="" className="page-logo" />
    </div>
  );
}

function InfoRow({
  k,
  v,
  strong,
}: {
  k: string;
  v: string;
  strong?: boolean;
}) {
  return (
    <div className="info-row">
      <span className="k">{k}</span>
      <span className={strong ? "v strong" : "v"}>{v}</span>
    </div>
  );
}

function Cell({
  k,
  v,
  total,
  accent,
}: {
  k: string;
  v: string;
  total?: boolean;
  accent?: boolean;
}) {
  return (
    <div
      className={`summary-cell${total ? " total" : ""}${accent ? " accent" : ""}`}
    >
      <div className="k">{k}</div>
      <div className="v">{v}</div>
    </div>
  );
}

function ScheduleTable({ result }: { result: ScheduleResult }) {
  const rows = result.rows;
  let lastYear = -1;

  return (
    <table className="schedule">
      <thead>
        <tr>
          <th>Code</th>
          <th>Description</th>
          <th>Due date</th>
          <th>%</th>
          <th>Amount</th>
          <th>Balance</th>
        </tr>
      </thead>
      <tbody>
        {rows.map((r) => {
          const showYearDivider = r.year > 0 && r.year !== lastYear;
          lastYear = r.year;
          const isDp = r.code.startsWith("DP");
          return (
            <Fragmentish key={r.code}>
              {showYearDivider && (
                <tr className="year-divider">
                  <td colSpan={6}>Year {r.year}</td>
                </tr>
              )}
              <tr className={isDp ? "row-dp" : undefined}>
                <td>
                  <span className="code-chip">{r.code}</span>
                </td>
                <td>{r.label}</td>
                <td>{formatDate(r.dueDate)}</td>
                <td>{percent(r.percent)}</td>
                <td>{money(r.amount)}</td>
                <td>{money(r.balance)}</td>
              </tr>
            </Fragmentish>
          );
        })}
      </tbody>
      <tfoot>
        <tr>
          <td colSpan={4}>Schedule total</td>
          <td>{money(result.scheduleTotal)}</td>
          <td></td>
        </tr>
      </tfoot>
    </table>
  );
}

// Small helper so we can return a divider row + the row from one map iteration.
function Fragmentish({ children }: { children: React.ReactNode }) {
  return <>{children}</>;
}
