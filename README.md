# m365-pull

A static SPA for downloading Teams chats and other M365 content from a single Microsoft tenant, with optional OneDrive sync for cross-device access.

- **Status:** working — two sources (Teams chats, Teams meeting transcripts), browser + OneDrive destinations, cross-device state via OneDrive
- **Maintainer:** TBD *(set before sharing beyond yourself)*

## What it does

A static SPA that lets a signed-in M365 user pull selected content from their own tenant to local files or to their OneDrive.

- **Teams chats** — list, search/filter/sort, mark for sync, per-chat lookback windows, incremental delta-sync, cumulative archive
- **Teams meeting transcripts** — list calendar meetings, locate the recording via Microsoft Graph search, fetch its transcript via SharePoint REST (works for meetings you attended, not just ones you organized)
- **Destinations** — browser save dialog, or a configured OneDrive folder that syncs across all your devices
- **State** — marks, lastSync timestamps, and user preferences live in `/Apps/m365-pull/state.json` in your OneDrive. Sign in on another device and your view carries over.
- **No backend** — pure SPA. Tokens are MSAL.js-managed in the browser. Graph and SharePoint REST are called directly from the page.

## Prerequisites

- Node.js 20+
- Microsoft Edge (or any Chromium-based browser). Firefox and Safari are explicitly **not** supported by the production design — see the design doc for why.
- An Entra app registration in your tenant (one-time setup below).

## One-time Azure setup

1. Open the [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. **Name:** `m365-pull` (anything descriptive — the user-facing name doesn't have to match).
3. **Supported account types:** *Accounts in this organizational directory only (single tenant)*.
4. **Redirect URI:**
   - Platform: **Single-page application (SPA)**
   - URI: `http://localhost:5173`
5. After creation, go to **Authentication** and (later) add any other URLs you'll run the dev server from (e.g. `http://192.168.x.x:5173` for LAN testing) and your deployed SWA URL as additional SPA redirect URIs.
6. Go to **API permissions** → add **Microsoft Graph → Delegated** permissions for what you intend to use:
   - `User.Read` (sign-in, profile)
   - `Chat.Read`, `Chat.ReadBasic` (Teams chats source)
   - `Calendars.Read`, `OnlineMeetings.Read`, `Files.Read.All`, `Sites.Read.All` (Teams meeting transcripts source)
   - `Files.ReadWrite.AppFolder` (cross-device state in OneDrive)
   - `Files.ReadWrite` (OneDrive destination, if you'll use it)
   - `offline_access` (token refresh)

   Then click **Grant admin consent for [your tenant]**. Some permissions need admin approval — file a request with your tenant admin if you don't have it.
7. From the **Overview** tab, copy:
   - **Application (client) ID**
   - **Directory (tenant) ID**

## Local development

```bash
cd m365-pull
npm install
cp .env.example .env.local       # then edit .env.local with your IDs
npm run dev
```

`.env.local` is gitignored — your tenant-specific IDs stay off GitHub. The file should look like:

```env
VITE_MSAL_CLIENT_ID=<your-application-client-id>
VITE_MSAL_TENANT_ID=<your-tenant-id>
```

Open `http://localhost:5173` in Edge. Click **Sign in**. Complete the Entra redirect. You should land back on the page signed in, ready to load chats or meetings.

If the redirect lands on a page that complains about a redirect URI mismatch, go back to your Entra app registration and add the exact URL you're using (including port) as a SPA redirect URI.

## Production build

```bash
npm run build
```

Produces a `dist/` directory. At build time Vite bakes the `VITE_MSAL_*` values from `.env.local` (or your CI environment) into the bundle. `staticwebapp.config.json` at the project root carries the production CSP headers and SPA navigation fallback rule.

## Deploy

Any static-site host works. The repo includes an `amplifier-online.yaml` for deployment via [amplifier-online](https://github.com/microsoft/amplifier-online) (Azure Static Web Apps with EasyAuth and PR previews); for a deployed environment, set the `VITE_MSAL_*` values in the build's CI environment instead of `.env.local`, and add the deployed URL as an additional SPA redirect URI on your Entra app registration.

## Sharing this repo

The repo is designed to be safely sharable — `.env.local` is gitignored, no tenant-specific values live in tracked source. Anyone forking it must provide their own Entra app registration + admin-consented Graph permissions to make it work in their own tenant.

## Open items

- [ ] Information governance review on browser-download exfiltration risk
- [ ] Named maintainer (replace `TBD` above)
