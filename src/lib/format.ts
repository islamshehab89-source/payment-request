import { CURRENCY as DEFAULT_CURRENCY } from "./projects";

// Currency comes from the Settings sheet of projects.xlsx; set once after load
// (before any schedule is rendered) via setCurrency().
let currentCurrency = DEFAULT_CURRENCY;

export function setCurrency(c: string) {
  currentCurrency = c;
}

export function getCurrency(): string {
  return currentCurrency;
}

// "EGP 1,234,567" — currency prefix, no decimals, grouped thousands.
export function money(n: number): string {
  const rounded = Math.round(n);
  const s = rounded.toLocaleString("en-US");
  return currentCurrency ? `${currentCurrency} ${s}` : s;
}

// Plain number for spreadsheet cells (no currency suffix, no grouping).
export function rawNumber(n: number): number {
  return Math.round(n);
}

export function percent(fraction: number): string {
  const p = fraction * 100;
  // up to 2 decimals, but trim trailing zeros: 0.0099 -> "0.99%", 0.1 -> "10%"
  return `${parseFloat(p.toFixed(2))}%`;
}

export function formatDate(d: Date): string {
  return d.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}

// Returns YYYY-MM-DD for <input type="date"> from a Date.
export function toInputDate(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

// Parses an <input type="date"> value as a LOCAL date (new Date("YYYY-MM-DD")
// would be UTC midnight and can shift a day in some timezones). Returns null
// for empty/invalid input so callers can guard.
export function parseInputDate(s: string): Date | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return Number.isNaN(d.getTime()) ? null : d;
}

// 48 → "4 Years", 12 → "1 Year", 42 → "3.5 Years", 30 → "2.5 Years".
export function monthsToYears(months: number): string {
  const y = Math.round((months / 12) * 10) / 10;
  const label = Number.isInteger(y) ? String(y) : y.toFixed(1);
  return `${label} ${y === 1 ? "Year" : "Years"}`;
}

export function addMonths(date: Date, months: number): Date {
  const d = new Date(date.getTime());
  const day = d.getDate();
  // Jump via the 1st so a 31st never overflows into the following month
  // (Jan 31 + 1 month must be Feb 28/29, not Mar 2/3).
  d.setDate(1);
  d.setMonth(d.getMonth() + months);
  const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  d.setDate(Math.min(day, lastDay));
  return d;
}
