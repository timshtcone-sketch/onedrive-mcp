// Aanmelden bij een persoonlijk Microsoft-account (OneDrive privé) via MSAL.
// Het token wordt lokaal bewaard in ~/.onedrive-mcp/token-cache.json.
import { PublicClientApplication } from "@azure/msal-node";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

export const READ_ONLY = process.env.ONEDRIVE_READ_ONLY === "1";
export const SCOPES = [READ_ONLY ? "Files.Read" : "Files.ReadWrite", "offline_access"];

const CONFIG_DIR = process.env.ONEDRIVE_MCP_DIR || path.join(os.homedir(), ".onedrive-mcp");
const CACHE_FILE = path.join(CONFIG_DIR, "token-cache.json");

function clientId() {
  const id = process.env.ONEDRIVE_CLIENT_ID;
  if (!id) {
    throw new Error("ONEDRIVE_CLIENT_ID ontbreekt. Zet de Application (client) ID van je app-registratie in deze omgevingsvariabele.");
  }
  return id;
}

let pca;
function app() {
  if (!pca) {
    pca = new PublicClientApplication({
      auth: {
        clientId: clientId(),
        // "consumers" = alleen persoonlijke Microsoft-accounts (outlook.com, hotmail, live).
        authority: "https://login.microsoftonline.com/consumers",
      },
    });
    if (fs.existsSync(CACHE_FILE)) {
      pca.getTokenCache().deserialize(fs.readFileSync(CACHE_FILE, "utf8"));
    }
  }
  return pca;
}

function saveCache() {
  fs.mkdirSync(CONFIG_DIR, { recursive: true, mode: 0o700 });
  fs.writeFileSync(CACHE_FILE, app().getTokenCache().serialize(), { mode: 0o600 });
}

// Eenmalige login in de terminal (npm run login).
export async function loginWithDeviceCode() {
  const result = await app().acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: (r) => console.log("\n" + r.message + "\n"),
  });
  saveCache();
  return result.account;
}

// Gebruikt door de MCP-server: haalt stil een (vernieuwd) token op.
export async function getAccessToken() {
  const accounts = await app().getTokenCache().getAllAccounts();
  if (accounts.length === 0) {
    throw new Error("Nog niet aangemeld. Voer eerst `npm run login` uit in de map van onedrive-mcp.");
  }
  try {
    const result = await app().acquireTokenSilent({ account: accounts[0], scopes: SCOPES });
    saveCache();
    return result.accessToken;
  } catch (err) {
    throw new Error(`Token verlopen of ingetrokken (${err.message}). Voer opnieuw \`npm run login\` uit.`);
  }
}
