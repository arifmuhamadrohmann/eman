import { put } from "@vercel/blob";
import { createHash, timingSafeEqual } from "node:crypto";
import { parseWorkbook, checkAgainstSummary } from "../lib/parse.js";

export const BLOB_PATH = "cgpp/data.json";
const hash = (s) => createHash("sha256").update(String(s ?? "")).digest();

// POST body = file .xlsx mentah. ?dry=1 hanya validasi tanpa publikasi.
export async function POST(req) {
  const pw = process.env.ADMIN_PASSWORD;
  if (!pw || !timingSafeEqual(hash(req.headers.get("x-admin-password")), hash(pw)))
    return Response.json({ error: "Password salah." }, { status: 401 });

  const name = decodeURIComponent(req.headers.get("x-filename") || "upload.xlsx");
  let data;
  try {
    data = parseWorkbook(Buffer.from(await req.arrayBuffer()), name);
  } catch (e) {
    return Response.json({ error: "Gagal membaca Excel: " + e.message }, { status: 400 });
  }
  const warnings = checkAgainstSummary(data);
  delete data._wb;

  const dry = new URL(req.url).searchParams.has("dry");
  if (!dry) {
    try {
      await put(BLOB_PATH, JSON.stringify(data), {
        access: "public", addRandomSuffix: false, allowOverwrite: true,
        contentType: "application/json", cacheControlMaxAge: 60,
      });
    } catch (e) {
      return Response.json({ error: "Gagal menyimpan ke Vercel Blob: " + e.message }, { status: 500 });
    }
  }

  return Response.json({
    published: !dry, warnings,
    summary: {
      file: name, indicators: data.indicators.length, levels: data.levels.length,
      months: data.months, currentMonth: data.meta.defaultMonth, fy: data.meta.fy,
    },
  });
}
