// Pakai: npm run check -- path/ke/file.xlsx
// Parse Excel, cocokkan hitungan dashboard dengan sheet Summary, lalu tulis public/data.json.
import fs from "node:fs";
import path from "node:path";
import { parseWorkbook, checkAgainstSummary } from "../lib/parse.js";

const file = process.argv[2];
if (!file) { console.error("Pakai: npm run check -- file.xlsx"); process.exit(1); }
const data = parseWorkbook(fs.readFileSync(file), path.basename(file));
const diffs = checkAgainstSummary(data);
delete data._wb;
console.log(`${data.indicators.length} indikator, ${data.levels.length} wilayah, bulan berjalan ${data.meta.defaultMonth}`);
if (diffs.length) { console.error("TIDAK cocok dengan Summary:\n" + diffs.join("\n")); process.exit(1); }
console.log("OK: klasifikasi dashboard == sheet Summary di semua wilayah & bulan.");
fs.writeFileSync("public/data.json", JSON.stringify(data));
console.log("public/data.json ditulis.");
