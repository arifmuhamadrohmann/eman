// Excel ITT CGPP -> JSON yang dibaca dashboard (window.__CGPP2__).
import * as XLSX from "xlsx";

export const LEVELS = [
  { k: "ID", type: "national", name: "CGPP Indonesia (Nasional)", prov: null, sheet: "ITT_CGPP Indonesia" },
  { k: "PPAP", type: "province", name: "Provinsi Papua", prov: "Provinsi Papua", sheet: "ITT_Provinsi Papua" },
  { k: "PPAT", type: "province", name: "Provinsi Papua Tengah", prov: "Provinsi Papua Tengah", sheet: "ITT_Provinsi Papua Tengah" },
  { k: "PJAT", type: "province", name: "Provinsi Jawa Timur", prov: "Provinsi Jawa Timur", sheet: "ITT_Provinsi Jawa Timur" },
  { k: "KJAY", type: "district", name: "Kabupaten Jayapura", prov: "Provinsi Papua" },
  { k: "KOJAY", type: "district", name: "Kota Jayapura", prov: "Provinsi Papua" },
  { k: "KMIM", type: "district", name: "Kabupaten Mimika", prov: "Provinsi Papua Tengah" },
  { k: "KNAB", type: "district", name: "Kabupaten Nabire", prov: "Provinsi Papua Tengah" },
  { k: "KOSBY", type: "district", name: "Kota Surabaya", prov: "Provinsi Jawa Timur" },
  { k: "KMLG", type: "district", name: "Kabupaten Malang", prov: "Provinsi Jawa Timur" },
  { k: "KSDA", type: "district", name: "Kabupaten Sidoarjo", prov: "Provinsi Jawa Timur" },
];

const num = (v) => (v === null || v === undefined || v === "" || isNaN(+v) ? null : +v);
const rnd = (v, d) => (v === null ? null : +v.toFixed(d));
const str = (v) => (v === null || v === undefined ? "" : String(v).trim());

function rows(wb, name) {
  const s = wb.Sheets[name];
  if (!s) throw new Error(`Sheet "${name}" tidak ditemukan di file Excel.`);
  return XLSX.utils.sheet_to_json(s, { header: 1, defval: null, raw: true });
}

// Posisi kolom dibaca dari header (baris 1-4), bukan di-hardcode, supaya tahan geser kolom.
function columns(a) {
  const [r1, r2, r3, r4] = a;
  const ytdStart = r1.findIndex((v) => /year-to-date/i.test(str(v)));
  if (ytdStart < 0) throw new Error('Header "Year-to-Date" tidak ditemukan di baris 1.');
  const months = [], H = {}, Y = {};
  const findSub = (from, to, row, label) => {
    for (let c = from; c < to; c++) if (str(row[c]).toLowerCase() === label) return c;
    return -1;
  };
  const blocks = [];
  r2.forEach((v, c) => { if (str(v)) blocks.push({ c, m: str(v) }); });
  blocks.forEach((b, i) => {
    const end = i + 1 < blocks.length ? blocks[i + 1].c : r3.length;
    const tgt = findSub(b.c, end, r3, "target"), ach = findSub(b.c, end, r3, "achievement");
    const col = { t: tgt, m: findSub(ach, end, r4, "male"), f: findSub(ach, end, r4, "female"), tot: findSub(ach, end, r4, "total"), p: findSub(ach, end, r4, "% achieved") };
    if (b.c < ytdStart) H[b.m] = col;
    else { Y[b.m] = col; if (b.m !== "Annual") months.push(b.m); }
  });
  if (!months.length || !Y.Annual) throw new Error("Blok bulan / Annual pada Year-to-Date tidak terbaca.");
  for (const m of months) if (!H[m]) throw new Error(`Bulan ${m} tidak ada di blok ITT Horizon.`);
  return { months, H, Y };
}

export function parseWorkbook(buf, source = "") {
  const wb = XLSX.read(buf, { type: "buffer" });
  const nat = rows(wb, LEVELS[0].sheet);
  const { months, H, Y } = columns(nat);

  // Daftar indikator dari sheet nasional (baris dengan kode indikator, bukan Nu_/De_).
  const indicators = [], numRows = {};
  let lvl = "", ocode = "", oname = "";
  for (let r = 4; r < nat.length; r++) {
    const row = nat[r] || [];
    const a = str(row[0]), code = str(row[4]);
    if (/^objective/i.test(a)) lvl = a;
    if (str(row[1])) { ocode = str(row[1]); oname = str(row[2]); }
    if (!code) continue;
    const nu = code.match(/^Nu_(.+)$/i);
    if (nu) { numRows[nu[1]] = r; continue; }
    if (/^De_/i.test(code)) continue;
    indicators.push({ row: r + 1, c: code, name: str(row[5]), ocode, oname, lvl, unit: str(row[6]), fmt: str(row[7]), impl: /^yes$/i.test(str(row[3])) ? 1 : 0 });
  }
  if (!indicators.length) throw new Error("Tidak ada indikator yang terbaca.");

  const v = {}, implByLevel = {}, annualByLevel = {}, mxNum = {};
  for (const L of LEVELS) {
    const a = rows(wb, L.sheet || L.name);
    const cell = (r, c) => (c < 0 ? null : num((a[r] || [])[c]));
    v[L.k] = {}; implByLevel[L.k] = []; annualByLevel[L.k] = []; mxNum[L.k] = {};
    indicators.forEach((ind, i) => {
      const r = ind.row - 1;
      if (str((a[r] || [])[4]) !== ind.c) throw new Error(`Sheet "${L.sheet || L.name}" baris ${ind.row}: kode indikator bukan ${ind.c}. Struktur baris harus sama di semua sheet.`);
      const impl = /^yes$/i.test(str(a[r][3])) ? 1 : 0;
      implByLevel[L.k].push(impl);
      const YA = Y.Annual;
      annualByLevel[L.k].push([rnd(cell(r, YA.t), 4), rnd(cell(r, YA.tot), 4), rnd(cell(r, YA.p) ?? 0, 6)]);
      const out = {};
      for (const m of months) {
        const y = Y[m], h = H[m];
        const yt = cell(r, y.t), ytot = cell(r, y.tot), ht = cell(r, h.t), htot = cell(r, h.tot);
        if (!yt && ytot === null && !ht && htot === null) continue;
        const hp = ht ? (htot ?? 0) / ht : null;
        out[m] = [
          rnd(yt, 4), rnd(ytot, 4), rnd(cell(r, y.m), 4), rnd(cell(r, y.f), 4), rnd(cell(r, y.p) ?? (yt ? 0 : null), 6), ytot !== null ? 1 : 0,
          rnd(ht, 4), rnd(htot, 4), rnd(cell(r, h.m), 4), rnd(cell(r, h.f), 4), rnd(hp, 6), htot !== null ? 1 : 0,
        ];
      }
      if (Object.keys(out).length) v[L.k][i] = out;
    });
    for (const code in numRows) {
      const mx = {};
      for (const m of months) { const x = cell(numRows[code], Y[m].tot); if (x !== null) mx[m] = x; }
      mxNum[L.k][code] = mx;
    }
  }

  // Bulan berjalan & FY dari sheet Summary (D3 = Current Period, E3 = Fiscal Year).
  let defaultMonth = months[months.length - 1], fy = "";
  if (wb.Sheets.Summary) {
    const s = rows(wb, "Summary");
    if (months.includes(str(s[2]?.[3]))) defaultMonth = str(s[2][3]);
    if (str(s[2]?.[4])) fy = "FY " + str(s[2][4]);
  }

  return {
    months,
    levels: LEVELS.map(({ sheet, ...l }) => l),
    indicators,
    units: [...new Set(indicators.map((x) => x.unit).filter(Boolean))].sort(),
    v, implByLevel, annualByLevel, mxNum,
    meta: {
      fy, defaultMonth, source, updated: new Date().toISOString(),
      bases: { H: "Bulan Ini (ITT Horizon)", Y: "Year-to-Date", A: "Tahunan (Annual)" },
    },
    _wb: wb,
  };
}

// Hitung ulang klasifikasi YTD persis seperti dashboard, lalu bandingkan dengan sheet Summary.
// Mengembalikan daftar selisih (kosong = dashboard == Excel).
export function checkAgainstSummary(data) {
  const wb = data._wb;
  if (!wb?.Sheets.Summary) return ["Sheet Summary tidak ada, verifikasi dilewati."];
  const s = rows(wb, "Summary"), hdr = s[5] || [], diffs = [];
  const cls = (p) => (p <= 0.5 ? 0 : p <= 0.75 ? 1 : p < 0.95 ? 2 : 3);
  for (const L of data.levels) {
    const name = L.type === "national" ? "CGPP Indonesia" : L.name;
    const c0 = hdr.findIndex((v) => str(v) === name);
    if (c0 < 0) { diffs.push(`Kolom "${name}" tidak ada di Summary.`); continue; }
    for (const m of data.months) {
      const row = s.find((r) => str(r?.[0]) === m);
      if (!row) continue;
      let tg = 0; const b = [0, 0, 0, 0];
      data.indicators.forEach((_, i) => {
        const e = data.v[L.k][i]?.[m];
        if (!data.implByLevel[L.k][i] || !e || !(e[0] > 0)) return;
        tg++; if (e[4] !== null) b[cls(e[4])]++;
      });
      const want = [0, 1, 2, 3, 4].map((k) => num(row[c0 + k]) ?? 0);
      const got = [tg, ...b];
      if (want.join() !== got.join()) diffs.push(`${L.name} ${m}: Excel [${want}] vs dashboard [${got}]`);
    }
  }
  return diffs;
}
