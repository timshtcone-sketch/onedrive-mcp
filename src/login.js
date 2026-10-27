#!/usr/bin/env node
import { loginWithDeviceCode, SCOPES } from "./auth.js";

try {
  console.log(`Aanmelden bij OneDrive (rechten: ${SCOPES.join(", ")})...`);
  const account = await loginWithDeviceCode();
  console.log(`Aangemeld als ${account?.username ?? "onbekend"}. Je kunt de MCP-server nu gebruiken.`);
} catch (err) {
  console.error("Aanmelden mislukt:", err.message);
  process.exit(1);
}
