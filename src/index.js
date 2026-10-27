#!/usr/bin/env node
// MCP-server (stdio) voor een persoonlijke OneDrive.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import fs from "node:fs";
import path from "node:path";
import { READ_ONLY } from "./auth.js";
import { graph, itemUrl, summarize, listAll, downloadBuffer, uploadBuffer, uploadFile } from "./graph.js";

const MAX_TEXT_BYTES = 2 * 1024 * 1024;
const MAX_IMAGE_BYTES = 5 * 1024 * 1024;
const TEXT_EXT = /\.(txt|md|csv|tsv|json|xml|html?|css|js|ts|py|log|yaml|yml|ini|cfg|sql|svg)$/i;
const IMAGE_MIME = /^image\/(png|jpeg|gif|webp)$/;

const server = new McpServer({ name: "onedrive-personal", version: "1.0.0" });

const json = (data) => ({ content: [{ type: "text", text: JSON.stringify(data, null, 2) }] });
const pathArg = z.string().describe('Pad in OneDrive, bijv. "/Documenten/notities.txt". "/" is de hoofdmap.');

function tool(name, description, inputSchema, handler, annotations = {}) {
  server.registerTool(name, { description, inputSchema, annotations }, async (args) => {
    try {
      return await handler(args);
    } catch (err) {
      return { isError: true, content: [{ type: "text", text: err.message }] };
    }
  });
}

// ---------- Lezen ----------

tool("onedrive_list", "Toon de inhoud van een map in OneDrive.", { path: pathArg.default("/") },
  async ({ path: p }) => json((await listAll(itemUrl(p) + "/children?$top=200")).map(summarize)),
  { readOnlyHint: true });

tool("onedrive_search", "Zoek bestanden en mappen in OneDrive op naam en inhoud.",
  { query: z.string().describe("Zoekterm") },
  async ({ query }) => {
    const q = encodeURIComponent(query.replace(/'/g, "''"));
    const res = await graph(`/me/drive/root/search(q='${q}')?$top=50`);
    return json(res.value.map(summarize));
  },
  { readOnlyHint: true });

tool("onedrive_get_info", "Haal details op van een bestand of map.", { path: pathArg },
  async ({ path: p }) => json(summarize(await graph(itemUrl(p)))),
  { readOnlyHint: true });

tool("onedrive_read_file",
  "Lees een tekstbestand (txt, md, csv, json, ...) of bekijk een afbeelding. Gebruik onedrive_download voor andere bestanden (Word, Excel, PDF).",
  { path: pathArg },
  async ({ path: p }) => {
    const item = await graph(itemUrl(p));
    if (item.folder) throw new Error("Dit is een map; gebruik onedrive_list.");
    const mime = item.file?.mimeType ?? "";
    if (IMAGE_MIME.test(mime)) {
      if (item.size > MAX_IMAGE_BYTES) throw new Error("Afbeelding is te groot; gebruik onedrive_download.");
      const buf = await downloadBuffer(p);
      return { content: [{ type: "image", data: buf.toString("base64"), mimeType: mime }] };
    }
    if (!(mime.startsWith("text/") || /json|xml|javascript/.test(mime) || TEXT_EXT.test(item.name))) {
      throw new Error(`Geen tekstbestand (${mime || "onbekend type"}). Gebruik onedrive_download om het lokaal op te slaan.`);
    }
    if (item.size > MAX_TEXT_BYTES) throw new Error("Bestand is groter dan 2 MB; gebruik onedrive_download.");
    return { content: [{ type: "text", text: (await downloadBuffer(p)).toString("utf8") }] };
  },
  { readOnlyHint: true });

tool("onedrive_download", "Download een bestand uit OneDrive naar een lokaal pad op deze computer.",
  {
    path: pathArg,
    local_path: z.string().describe("Lokaal doelpad (bestand of bestaande map)."),
  },
  async ({ path: p, local_path }) => {
    let target = path.resolve(local_path);
    if (fs.existsSync(target) && fs.statSync(target).isDirectory()) target = path.join(target, path.basename(p));
    const buf = await downloadBuffer(p);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, buf);
    return json({ saved: target, bytes: buf.length });
  });

// ---------- Schrijven (uit met ONEDRIVE_READ_ONLY=1) ----------

if (!READ_ONLY) {
  tool("onedrive_write_text", "Maak of overschrijf een tekstbestand in OneDrive.",
    { path: pathArg, content: z.string().describe("Inhoud van het bestand") },
    async ({ path: p, content }) => json(summarize(await uploadBuffer(p, Buffer.from(content, "utf8")))));

  tool("onedrive_upload", "Upload een lokaal bestand naar OneDrive (ook grote bestanden).",
    {
      local_path: z.string().describe("Pad van het lokale bestand"),
      path: pathArg.describe('Doelpad in OneDrive inclusief bestandsnaam, bijv. "/Foto\'s/vakantie.jpg"'),
      if_exists: z.enum(["replace", "rename", "fail"]).default("rename").describe("Wat te doen als het bestand al bestaat"),
    },
    async ({ local_path, path: p, if_exists }) => json(summarize(await uploadFile(path.resolve(local_path), p, if_exists))));

  tool("onedrive_create_folder", "Maak een nieuwe map aan.",
    { parent_path: pathArg.default("/"), name: z.string().describe("Naam van de nieuwe map") },
    async ({ parent_path, name }) => json(summarize(await graph(itemUrl(parent_path) + "/children", {
      method: "POST",
      body: { name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" },
    }))));

  tool("onedrive_move", "Verplaats en/of hernoem een bestand of map.",
    {
      path: pathArg,
      new_parent_path: z.string().optional().describe("Nieuwe bovenliggende map (weglaten = blijft waar het staat)"),
      new_name: z.string().optional().describe("Nieuwe naam (weglaten = naam blijft gelijk)"),
    },
    async ({ path: p, new_parent_path, new_name }) => {
      if (!new_parent_path && !new_name) throw new Error("Geef new_parent_path en/of new_name op.");
      const body = {};
      if (new_name) body.name = new_name;
      if (new_parent_path) body.parentReference = { id: (await graph(itemUrl(new_parent_path))).id };
      return json(summarize(await graph(itemUrl(p), { method: "PATCH", body })));
    });

  tool("onedrive_delete", "Verwijder een bestand of map (gaat naar de OneDrive-prullenbak).", { path: pathArg },
    async ({ path: p }) => {
      if (p.split("/").filter(Boolean).length === 0) throw new Error("De hoofdmap kan niet verwijderd worden.");
      await graph(itemUrl(p), { method: "DELETE" });
      return json({ deleted: p });
    },
    { destructiveHint: true });

  tool("onedrive_share_link", "Maak een deellink voor een bestand of map.",
    {
      path: pathArg,
      type: z.enum(["view", "edit"]).default("view").describe("view = alleen lezen, edit = bewerken"),
    },
    async ({ path: p, type }) => {
      const res = await graph(itemUrl(p) + "/createLink", { method: "POST", body: { type, scope: "anonymous" } });
      return json({ url: res.link.webUrl, type });
    });
}

await server.connect(new StdioServerTransport());
