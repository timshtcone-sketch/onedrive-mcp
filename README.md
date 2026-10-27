# onedrive-mcp

Lokale MCP-server waarmee Claude (Desktop of Claude Code) bij je **persoonlijke OneDrive** kan.
Werkt met privé-Microsoft-accounts (outlook.com, hotmail.com, live.com). Je token blijft op je eigen computer.

## Tools

| Tool | Wat het doet |
|---|---|
| `onedrive_list` | Inhoud van een map tonen |
| `onedrive_search` | Zoeken op naam en inhoud |
| `onedrive_get_info` | Details van een bestand of map |
| `onedrive_read_file` | Tekstbestand lezen of afbeelding bekijken |
| `onedrive_download` | Bestand naar je computer downloaden (Word, Excel, PDF, ...) |
| `onedrive_write_text` | Tekstbestand maken of overschrijven |
| `onedrive_upload` | Lokaal bestand uploaden (ook grote bestanden) |
| `onedrive_create_folder` | Map aanmaken |
| `onedrive_move` | Verplaatsen en/of hernoemen |
| `onedrive_delete` | Verwijderen (gaat naar de OneDrive-prullenbak) |
| `onedrive_share_link` | Deellink maken (bekijken of bewerken) |

Met `ONEDRIVE_READ_ONLY=1` zijn alleen de eerste vijf beschikbaar en vraagt de app ook alleen leesrechten.

## 1. Vereisten

- Node.js 20 of nieuwer (`node -v`)

## 2. App registreren bij Microsoft (eenmalig, gratis)

1. Ga naar <https://entra.microsoft.com> (of <https://portal.azure.com>) en log in met je Microsoft-account.
   - Krijg je een melding dat je geen directory/tenant hebt, maak dan eerst een gratis Azure-account aan
     (<https://azure.microsoft.com/free>). Daarbij kan om verificatie gevraagd worden; app-registraties zelf kosten niets.
2. **App registrations** → **New registration**
   - Name: `OneDrive MCP`
   - Supported account types: **Personal Microsoft accounts only**
   - Redirect URI: leeg laten
   - **Register**
3. Kopieer de **Application (client) ID** van de overzichtspagina.
4. **Authentication** → onderaan **Allow public client flows** op **Yes** → **Save**.
5. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** →
   vink `Files.ReadWrite` (of alleen `Files.Read`) en `offline_access` aan → **Add permissions**.

## 3. Installeren en aanmelden

```bash
cd onedrive-mcp
npm install
```

macOS / Linux:
```bash
ONEDRIVE_CLIENT_ID=<jouw-client-id> npm run login
```

Windows (PowerShell):
```powershell
$env:ONEDRIVE_CLIENT_ID="<jouw-client-id>"; npm run login
```

Je krijgt een link en een code. Open de link, vul de code in en geef toestemming.
Het token wordt bewaard in `~/.onedrive-mcp/token-cache.json` en automatisch vernieuwd.
Verloopt het toch (bijv. na lange tijd niet gebruiken of na een wachtwoordwijziging), voer `npm run login` dan opnieuw uit.

## 4. Toevoegen aan Claude

Gebruik het **volledige pad** naar `src/index.js`.

### Claude Desktop

Instellingen → Developer → Edit Config (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "onedrive": {
      "command": "node",
      "args": ["/volledig/pad/naar/onedrive-mcp/src/index.js"],
      "env": { "ONEDRIVE_CLIENT_ID": "<jouw-client-id>" }
    }
  }
}
```

Op Windows ziet het pad er zo uit: `"C:\\Users\\Tim\\onedrive-mcp\\src\\index.js"`.
Start Claude Desktop daarna opnieuw.

### Claude Code

```bash
claude mcp add onedrive --scope user -e ONEDRIVE_CLIENT_ID=<jouw-client-id> -- node /volledig/pad/naar/onedrive-mcp/src/index.js
```

Controleer met `claude mcp list`.

## Alleen lezen

Zet `"ONEDRIVE_READ_ONLY": "1"` bij `env` en voer `npm run login` opnieuw uit met dezelfde variabele.

## Afmelden

Verwijder `~/.onedrive-mcp/token-cache.json`. Je kunt de toegang ook intrekken op
<https://account.live.com/consent/Manage>.
