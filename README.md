# MDU QC — Connecteam replacement

| Page | Who | Where | Sign-in |
|---|---|---|---|
| **MDU Job Entry** | Sub techs (no Microsoft account needed) | `triagepartners.github.io/MDU-QC/` | None (invisible bot check) |
| **MDU QC Review** | Triage supervisors | `triagepartners.github.io/MDU-QC/review/` | Their Triage Microsoft account |

## How it fits together

```
Tech's phone ──► GitHub Pages form ──► Cloudflare Worker (free) ──► SharePoint site "MDU QC"
                                                                       ├─ list  "MDU QC Jobs"   (one row per job)
                                                                       └─ folder "MDU QC Photos" (one folder per job)
Supervisor ──► GitHub Pages review page ──(Microsoft sign-in)──────────┘
Power Automate (standard connectors) ── watches the list → Teams/email alerts, "needs fix" emails to techs
```

- **Cost and cards:** $0, and nothing asks for a credit card.
  - GitHub Pages works the same as for your Field Partners site.
  - Cloudflare's free Workers and Turnstile plans don't ask for a card.
  - SharePoint, Entra ID and standard Power Automate are already included in your Microsoft 365.
- **Where the data lives:** all job data and photos live in your SharePoint, not Cloudflare. The Worker is only a relay: it holds the Microsoft credentials so the public form never sees them.
- **Supervisor access:** supervisors use their own SharePoint permissions. Remove someone from the site and they lose access to the review page.

## Setup

### 1. SharePoint (5 min)
- Create a Team site called **MDU QC**, at the address `/sites/MDUQC`.
- Add the supervisors as **Members**.

### 2. App registration (M365 admin, about 15 min)
Go to **Entra admin center → App registrations → New registration**. Name it `MDU QC` and choose single tenant. Then:

1. **Authentication:**
   - Add a **Single-page application** platform.
   - Set the redirect URI to `https://triagepartners.github.io/MDU-QC/review/`.
2. **API permissions → Microsoft Graph**, then **Grant admin consent**:
   - Delegated: `Sites.ReadWrite.All` (for the review page; supervisors still only see sites they belong to).
   - Application: `Sites.Selected` (for the Worker; gives it nothing until step 3).
3. **Certificates & secrets:**
   - Create a new client secret and copy the **Value**.
   - **Put a calendar reminder for its expiry date.** When it expires, uploads stop until you add a new one.
4. **Give the app access to only the MDU QC site.**
   - Open Graph Explorer as an admin.
   - `GET https://graph.microsoft.com/v1.0/sites/triagepartners.sharepoint.com:/sites/MDUQC` and copy the `id`.
   - `POST https://graph.microsoft.com/v1.0/sites/{id}/permissions` with this body:
     ```json
     { "roles": ["fullcontrol"], "grantedToIdentities": [{ "application": { "id": "<client id>", "displayName": "MDU QC" } }] }
     ```
   - Full control is needed once so the app can create the list, and it covers only this one site.
   - Simpler but broader alternative: give the app Application `Sites.ReadWrite.All` instead of `Sites.Selected`, and skip this step.

### 3. Cloudflare (10 min, free, no card)
1. Sign up at cloudflare.com.
2. Go to **Turnstile → Add widget**:
   - Hostname: `triagepartners.github.io`
   - Mode: Managed
   - Copy the site key and the secret key.
3. Deploy the Worker. Run this from the `worker` folder (needs Node):
   ```
   npm i -g wrangler
   wrangler login
   wrangler secret put MS_TENANT_ID
   wrangler secret put MS_CLIENT_ID
   wrangler secret put MS_CLIENT_SECRET
   wrangler secret put TURNSTILE_SECRET
   wrangler secret put TICKET_SECRET      # any long random string
   wrangler secret put SETUP_KEY          # any random string
   wrangler deploy                        # prints https://triage-mdu-qc.<you>.workers.dev
   ```
4. Open `https://triage-mdu-qc.<you>.workers.dev/admin/setup?key=<SETUP_KEY>` once. It creates the **MDU QC Jobs** list and its columns.

### 4. GitHub Pages (5 min)
1. Create a repo called `MDU-QC` in the triagepartners org and upload this folder.
2. Edit the **SETTINGS** block at the top of each page:
   - `index.html`: `apiBase` (the Worker URL) and `turnstileSiteKey`.
   - `review/index.html`: `msal.clientId` and `msal.tenantId`.
3. Turn on Pages: **Settings → Pages → Deploy from branch → main / root**.

### 5. Power Automate (optional, standard connectors only)
- **New job alert:**
  - Trigger: SharePoint *When an item is created* on MDU QC Jobs.
  - Action: post in Teams or send an email.
  - Link: `https://triagepartners.github.io/MDU-QC/review/#@{triggerOutputs()?['body/ID']}`
- **Needs-fix email to the tech:**
  - Trigger: *When an item or a file is modified*.
  - Condition: `Status` is Needs Fix or Rejected **and** `TechEmail` is not empty.
  - Action: send an email with `ReviewNote`.

## Everyday changes
Each page is one self-contained file, so you can open it straight from Downloads to preview it. Submitting and sign-in only work once it's published and set up.

- **Unit codes:** edit the `MDU_UNITS` list in `index.html`. It was generated from *Tech Billing Codes Rates April 2026.xlsx*. Codes marked `caOnly` only appear when the state is CA.
- **States, photo minimum:** edit the SETTINGS block at the top of `index.html`.
- **Week-ending day:** edit the SETTINGS block at the top of `review/index.html`.
- **Supervisors:** add or remove members of the SharePoint site.

## Before canceling Connecteam
Connecteam's photo links will likely stop working after you cancel. Export every date range and run:
```
pip install openpyxl requests
python scripts/archive_connecteam.py "MDU Job Entry - Midwest.xlsx" connecteam_archive
```
Then upload the `connecteam_archive` folder to the MDU QC SharePoint site.
