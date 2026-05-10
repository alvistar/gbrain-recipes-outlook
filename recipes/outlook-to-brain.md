---
id: outlook-to-brain
name: Outlook-to-Brain
version: 0.1.0
description: >
  Outlook messages flow into brain pages via @softeria/ms-365-mcp-server.
  Deterministic collector pulls emails and directory data, agent analyzes
  and enriches entities.
category: sense
requires: []
secrets:
  - name: MS365_MCP_CLIENT_ID
    description: Azure AD application (client) ID for Microsoft Graph access
    where: https://portal.azure.com > App registrations > your app > Overview
  - name: MS365_MCP_TENANT_ID
    description: Azure AD tenant ID (work/school accounts only)
    where: https://portal.azure.com > App registrations > your app > Overview
health_checks:
  - type: env_exists
    name: MS365_MCP_CLIENT_ID
    label: "Azure AD app client ID"
  - type: env_exists
    name: MS365_MCP_TENANT_ID
    label: "Azure AD tenant ID"
setup_time: 15 min
cost_estimate: "$0 (Microsoft Graph API free tier)"
---

# Outlook-to-Brain: Email + Directory That Update Your Brain

Emails arrive. Brain pages get smarter. The collector pulls your Outlook
inbox, resolves senders against Azure AD (title, department, manager),
and the agent reads the digest to enrich brain pages with real metadata.

## IMPORTANT: Instructions for the Agent

**You are the installer.** Follow these steps precisely.

**The core pattern: code for data, LLMs for judgment.**
Email collection is split into two layers:
1. DETERMINISTIC: the collector script pulls emails, generates Outlook
   links, filters noise, and resolves senders against the org directory.
   This never fails. Links are always correct. Directory data is always
   structured.
2. LATENT: you (the agent) read the collected digests and make judgment
   calls. Who is important? What entities are mentioned? What action
   items exist? When directory data is available, use it to pre-populate
   people pages with title, department, and company.

**Do not try to pull emails yourself.** Use the collector script. It
handles pagination, deduplication, noise filtering, directory caching,
and Outlook link generation. If you try to do this via raw API calls,
you WILL miss emails, break pagination, or generate incorrect links.

## Architecture

```
Microsoft 365 Account(s)
  |  (OAuth via device code flow)
  v
@softeria/ms-365-mcp-server (handles auth, token refresh, Graph API)
  |  (MCP protocol via @modelcontextprotocol/sdk)
  v
Outlook Collector Script (deterministic Node.js)
  |  Calls: list-mail-messages, list-mail-folder-messages, list-users
  |  Handles: pagination, dedup, noise filtering, directory cache
  v  Outputs:
  +-- data/messages/{YYYY-MM-DD}.json     (structured email data)
  +-- data/directory/{hash}.json          (cached directory lookups, 7d TTL)
  +-- data/digests/{YYYY-MM-DD}.md        (markdown digest for agent)
  +-- data/state.json                     (pagination state, known IDs)
  |
  v
Agent reads digest
  |  Judgment calls:
  +-- Entity detection (people, companies mentioned)
  +-- Brain page updates (timeline entries, compiled truth)
  +-- People page creation with directory metadata
  +-- Action item extraction
  +-- Priority classification (urgent / normal / noise)
```

## Prerequisites

1. **GBrain installed and configured** (`gbrain doctor` passes)
2. **Node.js 18+** (for the collector script)
3. **Microsoft 365 account** (work/school or personal Outlook.com)

## Setup Flow

### Step 1: Authenticate with Microsoft 365

Ask the user: "Do you have a work/school Microsoft 365 account, or a personal Outlook.com account?"

#### Option A: Work/School Account (recommended, enables directory lookup)

Tell the user:
"I need Azure AD app registration credentials for Microsoft Graph access.

1. Go to https://portal.azure.com
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Name: 'GBrain Outlook Collector' (anything works)
4. Supported account types: **Accounts in this organizational directory only**
5. Redirect URI: leave blank (we use device code flow)
6. Click **Register**
7. Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page
8. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**
9. Add: `Mail.Read`, `User.Read.All`
10. Click **Grant admin consent** (requires admin privileges)"

Then authenticate the MCP server:
```bash
export MS365_MCP_CLIENT_ID=<client-id>
export MS365_MCP_TENANT_ID=<tenant-id>
npx @softeria/ms-365-mcp-server --login --org-mode
```

**STOP until authentication validates.**

#### Option B: Personal Outlook.com Account

Tell the user:
"Personal accounts work but without directory lookup (no Azure AD).
The collector will still pull emails, filter noise, and generate digests.
People pages won't have automatic title/department metadata."

```bash
npx @softeria/ms-365-mcp-server --login
```

**STOP until authentication validates.**

### Step 2: Install the Collector

```bash
cd /path/to/gbrain-recipes-outlook
npm install
```

### Step 3: Run First Collection

```bash
node outlook-collector/collector.mjs collect
node outlook-collector/collector.mjs digest
```

Verify: `ls data/digests/` should show today's digest file.
Read the digest. Confirm it contains real emails with working Outlook links.

### Step 4: Enrich Brain Pages

This is YOUR job (the agent). Read the digest. For each email:

1. **Detect entities**: who sent it? Who is mentioned? What companies?
2. **Use directory data**: if the digest shows `(VP Engineering @ Acme Corp, Engineering)`
   next to a sender, use that to pre-populate the people page:
   ```markdown
   ---
   type: person
   name: Alice Smith
   email: alice@acme.com
   title: VP of Engineering
   department: Engineering
   company: Acme Corp
   ---
   ```
3. **Check the brain**: `gbrain search "sender name"` to check for existing pages
4. **Update brain pages**: append timeline entries:
   `- YYYY-MM-DD | Email from {sender}: {subject} [Source: Outlook, {date}]`
5. **Create new pages**: if sender is notable and has no page, create one
   (use directory data when available)
6. **Extract action items**: if the email requires a response or action, log it
7. **Sync**: run `gbrain sync --no-pull --no-embed` to index changes

### Step 5: Set Up Cron

The collector should run every 30 minutes:

```bash
*/30 * * * * cd /path/to/gbrain-recipes-outlook && node outlook-collector/collector.mjs collect && node outlook-collector/collector.mjs digest
```

The agent should read the digest on a schedule (e.g., 3x/day) and run
the enrichment flow from Step 4.

### Step 6: Log Setup Completion

```bash
mkdir -p ~/.gbrain/integrations/outlook-to-brain
echo '{"ts":"'$(date -u +%Y-%m-%dT%H:%M:%SZ)'","event":"setup_complete","source_version":"0.1.0","status":"ok"}' >> ~/.gbrain/integrations/outlook-to-brain/heartbeat.jsonl
```

## Noise Filtering (Deterministic)

The collector filters these sender patterns as noise:
- `noreply`, `no-reply`, `notifications@`, `calendar-notification`
- `mailer-daemon`, `postmaster`, `donotreply`

Signature-related emails are flagged separately:
- DocuSign, Dropbox Sign, HelloSign, PandaDoc
- "please sign", "signature needed", "ready for your signature"

## Directory Lookup

For work/school accounts with `--org-mode`, the collector resolves each
unique sender against Azure AD using `$filter=mail eq '{email}'`. Results
are cached for 7 days. The digest includes the sender's title, department,
and company when available.

Use `--refresh-directory` to force re-fetch all cached directory entries.

Personal accounts skip directory lookup entirely (no Azure AD available).

## Troubleshooting

**No emails collected:**
- Run `npx @softeria/ms-365-mcp-server --login` to re-authenticate
- Check that `MS365_MCP_CLIENT_ID` and `MS365_MCP_TENANT_ID` are set
- Check the collector's stderr output for error messages

**Directory lookup not working:**
- Requires `--org-mode` (work/school accounts only)
- Requires `User.Read.All` permission on the Azure AD app
- Check that admin consent was granted for the permission

**Outlook links don't work:**
- Links open in the default browser. The user must be logged into the
  correct Microsoft account.

**Corrupted state (collector keeps re-collecting old emails):**
- Delete `data/state.json` to start fresh. The collector will re-collect
  recent emails, but deduplication on the brain side catches duplicates.

## Cost Estimate

| Component | Monthly Cost |
|-----------|-------------|
| Azure AD app registration | $0 |
| Microsoft Graph API | $0 (within free quota) |
| **Total** | **$0** |
