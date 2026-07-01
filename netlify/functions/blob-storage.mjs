import { getStore } from "@netlify/blobs";

const STORE_NAME = "sales-upload-state";

function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
    },
  });
}

function requireAuthorized(request) {
  const expectedSecret = Netlify.env.get("SALES_STORAGE_SECRET");
  const actualSecret = request.headers.get("x-sales-storage-secret");
  return Boolean(expectedSecret && actualSecret && expectedSecret === actualSecret);
}

export default async (request) => {
  if (!requireAuthorized(request)) {
    return jsonResponse({ error: "Unauthorized" }, 401);
  }

  const url = new URL(request.url);
  const key = url.searchParams.get("key");
  if (!key || key.includes("..") || key.startsWith("/")) {
    return jsonResponse({ error: "Invalid key" }, 400);
  }

  const store = getStore({ name: STORE_NAME, consistency: "strong" });

  if (request.method === "GET") {
    const data = await store.get(key, { type: "arrayBuffer" });
    if (data === null) {
      return jsonResponse({ error: "Not found" }, 404);
    }

    const metadata = await store.getMetadata(key);
    return new Response(data, {
      status: 200,
      headers: {
        "Content-Type": metadata?.metadata?.contentType || "application/octet-stream",
        "Cache-Control": "no-store",
      },
    });
  }

  if (request.method === "PUT") {
    const data = await request.arrayBuffer();
    const contentType = request.headers.get("content-type") || "application/octet-stream";
    await store.set(key, data, {
      metadata: {
        contentType,
        updatedAt: new Date().toISOString(),
      },
    });
    return jsonResponse({ ok: true, key });
  }

  if (request.method === "DELETE") {
    await store.delete(key);
    return jsonResponse({ ok: true, key });
  }

  return jsonResponse({ error: "Method not allowed" }, 405);
};

export const config = {
  path: "/api/blob-storage",
  method: ["GET", "PUT", "DELETE"],
};
