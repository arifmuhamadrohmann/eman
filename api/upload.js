import { put } from "@vercel/blob";
import { createHash, timingSafeEqual } from "node:crypto";
import { parseWorkbook, checkAgainstSummary } from "../lib/parse.js";

export const BLOB_PATH = "cgpp/data.json";
// Token Blob; nama env bisa berawalan lain kalau prefix diubah saat Connect store di Vercel.
export const blobToken = () =>
  process.env.BLOB_READ_WRITE_TOKEN ||
  Object.entries(process.env).find(([k, v]) => k.endsWith("READ_WRITE_TOKEN") && v)?.[1];
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
    const token = blobToken();
    if (!token) {
      const names = Object.keys(process.env).filter((k) => /BLOB|STORE/i.test(k));
      return Response.json({ error: "Token Vercel Blob tidak ada di server. Env terkait yang terlihat: " + (names.join(", ") || "(tidak ada)") + ". Hubungkan Blob store ke project lalu Redeploy." }, { status: 500 });
    }
    try {
      await put(BLOB_PATH, JSON.stringify(data), {
        token, access: "public", addRandomSuffix: false, allowOverwrite: true,
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
