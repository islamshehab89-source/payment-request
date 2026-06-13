// Regenerates public/projects.xlsx with sample data and a ReadMe sheet.
// Run with: node scripts/generate-template.mjs --force
// REFUSES to overwrite an existing file unless --force is passed, because the
// real project data lives in that file.
import XLSX from "xlsx";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
fs.mkdirSync(path.join(root, "public"), { recursive: true });
const out = path.join(root, "public", "projects.xlsx");

if (fs.existsSync(out) && !process.argv.includes("--force")) {
  console.error(
    "public/projects.xlsx already exists — it holds your real data.\n" +
      "Pass --force to overwrite it with the sample template (this DELETES your data).",
  );
  process.exit(1);
}

const HEAD = [
  "Project",
  "Type",
  "BUA",
  "Phase",
  "Plan",
  "Years",
  "Discount %",
  "Down Payment Type",
  "Down Payment 1 %",
  "Down Payment 2 %",
  "DP2 Due (months)",
  "Installments Count",
  "Installment Every (months)",
  "First Installment Due (months)",
  "Maintenance %",
  "Delivery (months)",
  "First Installment Maintenance Due (months)",
  "Ending Installment Maintenance Due (months)",
  "Maintenance Every (months)",
  "Maintenance Basis",
];

// All % columns are plain numbers in percent units: 10 = 10%, 0.99 = 0.99%.
const ROWS = [
  ["Sample Project", "Core & Shell", null, null, "Cash", 0, 25, "Selling price", 100, 0, 0, 0, 0, 0, 10, 42, 15, 39, 6, "Original Price"],
  [null, "Core & Shell", null, null, "5% & 5% & 8 years", 8, 0, "Original Price", 5, 5, 3, 31, 3, 3, 10, 42, 15, 39, 6, "Original Price"],
  [null, "Finished", 20000, null, "5% & 5% & 7 years", 7, 3, "Original Price", 5, 5, 3, 31, 3, 3, 10, 42, 15, 39, 6, "Selling price"],
];

const projectsWs = XLSX.utils.aoa_to_sheet([HEAD, ...ROWS]);
projectsWs["!cols"] = HEAD.map((h) => ({ wch: Math.max(10, h.length + 2) }));

const settingsWs = XLSX.utils.aoa_to_sheet([
  ["Setting", "Value"],
  ["Company Name", "Your Company"],
  ["Currency", "EGP"],
]);
settingsWs["!cols"] = [{ wch: 16 }, { wch: 24 }];

const readmeWs = XLSX.utils.aoa_to_sheet([
  ["How to fill the Projects sheet", ""],
  ["", ""],
  ["Each row is ONE payment plan. Rows with a blank Project cell belong to the project above them.", "كل صف = نظام سداد واحد. لو خلية المشروع فاضية، الصف بيتبع المشروع اللي فوقه."],
  ["All % columns are plain numbers in percent units: 10 means 10%.", "أعمدة النسب أرقام عادية: 10 يعني 10%."],
  ["Empty Type / Phase / BUA cells simply don't appear in the page.", "خلايا Type / Phase / BUA الفاضية مش بتظهر في الصفحة."],
  ["", ""],
  ["Column", "Meaning / المعنى"],
  ["Project", "Project name. Leave blank to repeat the project above. / اسم المشروع"],
  ["Type", "Unit status shown as a dropdown (Primary, RTM, Finished…). Optional. / حالة الوحدة"],
  ["BUA", "Finishing price per m². The page asks for the unit area and adds area × BUA after the discount. Optional. / سعر متر التشطيب"],
  ["Phase", "Project phase shown as a dropdown. Optional. / مرحلة المشروع"],
  ["Plan", "Plan label shown in the dropdown / اسم نظام السداد"],
  ["Years", "Plan length in years (informational) / مدة النظام بالسنين"],
  ["Discount %", "Discount off the original price; negative = premium / نسبة الخصم"],
  ["Down Payment Type", "\"Original Price\" or \"Selling price\" — what the DP %s are taken on (Selling = final price incl. finishing) / أساس حساب المقدمات"],
  ["Down Payment 1 %", "First down payment, due at contract date / المقدم الأول"],
  ["Down Payment 2 %", "Second down payment, 0 if none / المقدم الثاني"],
  ["DP2 Due (months)", "When DP2 is due, months after contract / استحقاق المقدم الثاني"],
  ["Installments Count", "Number of equal installments. The rest of the FINAL price is split across them. / عدد الأقساط"],
  ["Installment Every (months)", "3 = quarterly, 1 = monthly / الفترة بين الأقساط"],
  ["First Installment Due (months)", "When installment 1 is due / استحقاق أول قسط"],
  ["Maintenance %", "Maintenance, % of the Maintenance Basis price / نسبة الصيانة"],
  ["Delivery (months)", "Months from contract to delivery (informational) / الاستلام بعد كام شهر"],
  ["First Installment Maintenance Due (months)", "First maintenance payment (0 = at contract) / أول قسط صيانة"],
  ["Ending Installment Maintenance Due (months)", "Last maintenance payment / آخر قسط صيانة"],
  ["Maintenance Every (months)", "Spacing of maintenance payments / الفترة بين أقساط الصيانة"],
  ["Maintenance Basis", "\"Original Price\" or \"Selling price\" (= final price incl. finishing) / أساس حساب الصيانة"],
  ["", ""],
  ["After editing: save the file, then refresh the page in the browser.", "بعد التعديل: احفظ الملف واعمل refresh للصفحة."],
]);
readmeWs["!cols"] = [{ wch: 60 }, { wch: 70 }];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, projectsWs, "Projects");
XLSX.utils.book_append_sheet(wb, settingsWs, "Settings");
XLSX.utils.book_append_sheet(wb, readmeWs, "ReadMe");
XLSX.writeFile(wb, out);
console.log("Wrote", out);
