// Kleine wrapper rond Microsoft Graph voor OneDrive.
import fs from "node:fs";
import { getAccessToken } from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
const SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;
const CHUNK_SIZE = 32 * 320 * 1024; // 10 MiB, moet een veelvoud van 320 KiB zijn

// "/Documenten/a b.txt" -> "/me/drive/root:/Documenten/a%20b.txt:"
export function itemUrl(p = "/") {
  const parts = p.split("/").filter(Boolean);
  if (parts.length === 0) return "/me/drive/root";
  return "/me/drive/root:/" + parts.map(encodeURIComponent).join("/") + ":";
}

export async function graph(urlPath, { method = "GET", body, headers = {}, raw = false } = {}) {
  const token = await getAccessToken();
  const isJson = body !== undefined && !(body instanceof Buffer) && typeof body !== "string";
  const res = await fetch(urlPath.startsWith("http") ? urlPath : GRAPH + urlPath, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(isJson ? { "Content-Type": "application/json" } : {}),
      ...headers,
    },
    body: isJson ? JSON.stringify(body) : body,
  });
  if (!res.ok) {
    let detail = await res.text();
    try {
      detail = JSON.parse(detail).error?.message ?? detail;
    } catch {}
    throw new Error(`Graph ${method} ${urlPath} gaf ${res.status}: ${detail}`);
  }
  if (raw) return res;
  if (res.status === 204) return null;
  return res.json();
}

export function summarize(item) {
  return {
    name: item.name,
    path: item.parentReference?.path
      ? item.parentReference.path.replace(/^\/drive\/root:?/, "") + "/" + item.name
      : undefined,
    type: item.folder ? "folder" : "file",
    size: item.size,
    modified: item.lastModifiedDateTime,
    ...(item.folder ? { childCount: item.folder.childCount } : {}),
    ...(item.file ? { mimeType: item.file.mimeType } : {}),
    webUrl: item.webUrl,
    id: item.id,
  };
}

export async function listAll(url) {
  const items = [];
  let next = url;
  while (next) {
    const page = await graph(next);
    items.push(...page.value);
    next = page["@odata.nextLink"];
  }
  return items;
}

export async function downloadBuffer(p) {
  const res = await graph(itemUrl(p) + "/content", { raw: true });
  return Buffer.from(await res.arrayBuffer());
}

export async function uploadBuffer(p, buf, conflict = "replace") {
  if (buf.length <= SIMPLE_UPLOAD_LIMIT) {
    return graph(`${itemUrl(p)}/content?@microsoft.graph.conflictBehavior=${conflict}`, {
      method: "PUT",
      body: buf,
      headers: { "Content-Type": "application/octet-stream" },
    });
  }
  const session = await graph(itemUrl(p) + "/createUploadSession", {
    method: "POST",
    body: { item: { "@microsoft.graph.conflictBehavior": conflict } },
  });
  let result;
  for (let start = 0; start < buf.length; start += CHUNK_SIZE) {
    const chunk = buf.subarray(start, Math.min(start + CHUNK_SIZE, buf.length));
    // De upload-URL is al geautoriseerd; geen Bearer-token meesturen.
    const res = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${start}-${start + chunk.length - 1}/${buf.length}`,
      },
      body: chunk,
    });
    if (!res.ok && res.status !== 202) {
      throw new Error(`Upload mislukt bij byte ${start}: ${res.status} ${await res.text()}`);
    }
    if (res.status === 200 || res.status === 201) result = await res.json();
  }
  return result;
}

export async function uploadFile(localPath, p, conflict) {
  return uploadBuffer(p, fs.readFileSync(localPath), conflict);
}
