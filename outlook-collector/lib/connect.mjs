import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { execSync } from 'child_process';

export function probeServer() {
  try {
    execSync('npx --no-install @softeria/ms-365-mcp-server --help', {
      stdio: 'pipe',
      timeout: 15000,
    });
  } catch {
    console.error(
      'ERROR: @softeria/ms-365-mcp-server not found.\n' +
      'Install it globally: npm install -g @softeria/ms-365-mcp-server\n' +
      'Or as a local dependency: npm install @softeria/ms-365-mcp-server'
    );
    process.exit(1);
  }
}

export async function createClient(preset = 'mail,users') {
  const args = ['-y', '@softeria/ms-365-mcp-server', '--read-only', '--preset', preset];
  if (process.env.MS365_MCP_TENANT_ID) {
    args.push('--org-mode');
  }

  const transport = new StdioClientTransport({
    command: 'npx',
    args,
  });

  const client = new Client(
    { name: 'outlook-collector', version: '0.1.0' },
    { capabilities: {} }
  );

  await client.connect(transport);
  return { client, transport };
}

export async function disconnect({ client }) {
  try {
    await client.close();
  } catch {}
}

export async function callTool(client, name, args = {}) {
  const maxRetries = 3;
  let lastError;

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      const result = await client.callTool({ name, arguments: args });
      return result;
    } catch (err) {
      lastError = err;
      if (attempt < maxRetries - 1) {
        const delay = 2000 * Math.pow(2, attempt);
        await new Promise(r => setTimeout(r, delay));
      }
    }
  }
  throw lastError;
}
