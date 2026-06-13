# Payment Request

A single-page tool that generates a real-estate **payment schedule + payment request**: pick a
project / unit status / phase / payment plan, enter the unit price (and area for finished units),
and get a Unit Information block, a full installment schedule, and a maintenance schedule — with
a printable PDF and an Excel export.

Built with Next.js 16 + React 19. No backend, no database — everything runs in the browser.

## Run it

```bash
pnpm install
pnpm dev -p 3011     # http://localhost:3011
pnpm test            # runs the logic tests against the real public/projects.xlsx
```

## Edit your data — in Excel

All projects and payment plans live in **`public/projects.xlsx`** (sheet `Projects`, one row per
plan + sheet `Settings` for company name and currency). Edit in Excel, save, refresh the page.

Selection cascade in the page: **Project → Unit Status (`Type`) → `Phase` → `Plan`** — a level
only appears when the sheet has values for it; blank cells never show up.

| Column | Meaning |
|---|---|
| `Project` | Project name (blank = same as the row above) |
| `Type` | Unit status dropdown: Primary / RTM / Core & Shell / Finished… (optional) |
| `BUA` | Finishing price per m². When set, the page asks for the unit area and adds `area × BUA` **after the discount** (optional) |
| `Phase` | Phase dropdown (optional) |
| `Plan` | Plan label shown in the dropdown |
| `Years` | Plan length (informational) |
| `Discount %` | Off the original price (negative = premium) |
| `Down Payment Type` | `Original Price` or `Selling price` — what the DP %s are taken on (`Selling price` = final price incl. finishing) |
| `Down Payment 1 %` / `2 %` | Down payments; DP1 at contract, DP2 after `DP2 Due (months)` |
| `Installments Count` / `Installment Every (months)` / `First Installment Due (months)` | The rest of the **final** price split equally |
| `Maintenance %` + `Maintenance Basis` | Maintenance total = % × basis price |
| `First/Ending Installment Maintenance Due (months)` + `Maintenance Every (months)` | Maintenance paid in equal installments across that window |
| `Delivery (months)` | Months to delivery (informational) |

**All `%` columns are plain numbers in percent units: `10` = 10%.** %-formatted cells also work.

### How the numbers are computed

```
Net Unit Price   = Original − Discount %          (or a manual override)
Finishing Cost   = Unit Area × BUA                (only when BUA is set)
Final Unit Price = Net + Finishing
Down payments    = % of the Down Payment Type price
Installments     = (Final − DP1 − DP2) / count    (equal, last absorbs rounding)
Maintenance      = % of the Maintenance Basis price, equal payments from
                   First to Ending every N months
Grand Total      = Final Unit Price + Maintenance
```

Rows with problems are skipped and listed in a banner with their row numbers; the rest still
loads. If a plan's payments don't add up to the final price, the page shows a warning with the
unscheduled amount. Missing/unreadable file → built-in sample data.

`pnpm template` regenerates a sample `public/projects.xlsx` — it **refuses to overwrite an
existing file** unless you pass `--force`. Backups live in `backups/`.

Excel parsing & validation: `src/lib/loadProjects.ts`. Schedule math: `src/lib/schedule.ts`.
Logic tests: `scripts/test-logic.mts`.

## Page behaviour

- **Logo**: shown from `public/logo.webp` (header + top of the printout). Falls back to the
  company name text if the file is missing.
- **Date** is always today and cannot be edited — it shows in Unit Information and drives the
  schedule.
- **Net price** is computed from the discount and is not editable in the form (it shows in Unit
  Information).
- **Unit Area** is optional for every plan (informational); for *finished* units (a `BUA` value)
  it is required and drives the finishing cost.

## Outputs

- **Print / Save PDF** — browser print dialog with a clean print stylesheet. Row highlights are
  preserved in the PDF (`print-color-adjust: exact`) and the schedule prints compactly.
- **Export Excel** — Unit Information + payment schedule + maintenance schedule (SheetJS).
