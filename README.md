# gbrain-recipes-outlook

Outlook-to-Brain ingestion recipe for [gbrain](https://github.com/garrytan/gbrain).

Pulls email from Microsoft 365 / Outlook via
[@softeria/ms-365-mcp-server](https://github.com/softeria/ms-365-mcp-server)
and generates markdown digests for agent enrichment. Azure AD directory lookup
auto-populates people pages with title, department, and company.

## Quick Start

```bash
# 1. Install dependencies
npm install
npm install -g @softeria/ms-365-mcp-server

# 2. Authenticate with Microsoft 365
npx @softeria/ms-365-mcp-server --login --org-mode  # work/school account
# or
npx @softeria/ms-365-mcp-server --login              # personal account

# 3. Set environment variables
export MS365_MCP_CLIENT_ID=<your-azure-ad-app-client-id>
export MS365_MCP_TENANT_ID=<your-azure-ad-tenant-id>

# 4. Run the collector
# Defaults to Inbox + Archive + Sent Items with a 2h overlap window so
# messages replied to and archived between polling cycles are still captured.
node outlook-collector/collector.mjs collect --folders inbox,archive,sentitems --lookback 2h
node outlook-collector/collector.mjs digest

# 5. Tell gbrain where to find the recipe
export GBRAIN_RECIPES_DIR=$(pwd)/recipes
gbrain integrations test recipes/outlook-to-brain.md
```

## Architecture

```
Microsoft 365 Account
  |
  v
@softeria/ms-365-mcp-server (handles OAuth, token refresh, Graph API)
  |  (MCP protocol via @modelcontextprotocol/sdk)
  v
Collector Script (deterministic, no LLM calls)
  |  Outputs:
  +-- data/messages/{YYYY-MM-DD}.json
  +-- data/directory/{hash}.json (cached, 7d TTL)
  +-- data/digests/{YYYY-MM-DD}.md
  +-- data/state.json
  |
  v
Agent reads digest, enriches brain pages
```

## Cron Setup

Poll every 10 minutes with a 2-hour safety lookback across Inbox, Archive, and
Sent Items. The overlap plus message-ID/conversation-ID state prevents missing
messages that were replied to and archived before the next poll.

```bash
*/10 * * * * cd /path/to/gbrain-recipes-outlook && node outlook-collector/collector.mjs collect --folders inbox,archive,sentitems --lookback 2h && node outlook-collector/collector.mjs digest
```

For a daily repair/reconciliation pass, use a wider lookback:

```bash
0 3 * * * cd /path/to/gbrain-recipes-outlook && node outlook-collector/collector.mjs collect --folders inbox,archive,sentitems --lookback 7d && node outlook-collector/collector.mjs digest
```

## License

MIT
