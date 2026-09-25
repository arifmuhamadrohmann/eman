import { head } from "@vercel/blob";
import { BLOB_PATH } from "./upload.js";

// Data terbaru dari Vercel Blob; kalau belum pernah upload, pakai public/data.json.
export async function GET(req) {
  try {
    const b = await head(BLOB_PATH);
    const r = await fetch(`${b.url}?v=${new Date(b.uploadedAt).getTime()}`);
    if (!r.ok) throw new Error("blob " + r.status);
    return new Response(r.body, {
      headers: { "content-type": "application/json", "cache-control": "public, s-maxage=30, stale-while-revalidate=300" },
    });
  } catch {
    return Response.redirect(new URL("/data.json", req.url), 307);
  }
}
