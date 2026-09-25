import { get } from "@vercel/blob";
import { BLOB_PATH, ACCESS } from "./upload.js";

// Data terbaru dari Vercel Blob; kalau belum pernah upload, pakai public/data.json.
export async function GET(req) {
  for (const access of ACCESS) {
    try {
      const b = await get(BLOB_PATH, { access, useCache: false });
      if (b?.statusCode === 200)
        return new Response(b.stream, {
          headers: { "content-type": "application/json", "cache-control": "public, s-maxage=30, stale-while-revalidate=300" },
        });
    } catch {}
  }
  return Response.redirect(new URL("/data.json", req.url), 307);
}
