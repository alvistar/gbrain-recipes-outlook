import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

const MOCK_MESSAGES = [
  {
    id: 'msg-001',
    subject: 'Q3 roadmap review',
    from: { emailAddress: { name: 'Alice Smith', address: 'alice@acme.com' } },
    receivedDateTime: '2026-05-10T09:00:00Z',
    bodyPreview: 'Let me know your thoughts on the Q3 roadmap attached.',
    conversationId: 'conv-001',
    webLink: 'https://outlook.office365.com/mail/inbox/id/msg-001',
  },
  {
    id: 'msg-002',
    subject: 'Your DocuSign envelope is ready',
    from: { emailAddress: { name: 'DocuSign', address: 'docusign@docusign.net' } },
    receivedDateTime: '2026-05-10T08:30:00Z',
    bodyPreview: 'Please sign the NDA.',
    conversationId: 'conv-002',
  },
  {
    id: 'msg-003',
    subject: 'Weekly digest',
    from: { emailAddress: { name: 'Notifications', address: 'noreply@service.com' } },
    receivedDateTime: '2026-05-10T07:00:00Z',
    bodyPreview: 'Here is your weekly summary.',
    conversationId: 'conv-003',
  },
];

const MOCK_SENT = [
  {
    id: 'sent-001',
    subject: 'Re: Q3 roadmap review',
    conversationId: 'conv-001',
    sentDateTime: '2026-05-10T10:00:00Z',
  },
];

const MOCK_USERS = [
  {
    displayName: 'Alice Smith',
    mail: 'alice@acme.com',
    jobTitle: 'VP of Engineering',
    department: 'Engineering',
    officeLocation: 'San Francisco',
    companyName: 'Acme Corp',
  },
];

const TOOLS = {
  'list-mail-messages': (args) => {
    return { value: MOCK_MESSAGES };
  },
  'list-mail-folder-messages': (args) => {
    return { value: MOCK_SENT };
  },
  'list-users': (args) => {
    const filter = args.$filter || '';
    const emailMatch = filter.match(/mail eq '([^']+)'/);
    if (emailMatch) {
      const email = emailMatch[1];
      const user = MOCK_USERS.find(u => u.mail === email);
      return { value: user ? [user] : [] };
    }
    if (args.$top === '1') {
      return { value: MOCK_USERS.slice(0, 1) };
    }
    return { value: MOCK_USERS };
  },
};

export { MOCK_MESSAGES, MOCK_SENT, MOCK_USERS };

export async function startMockServer() {
  const server = new Server(
    { name: 'mock-ms365-mcp', version: '1.0.0' },
    { capabilities: { tools: {} } }
  );

  server.setRequestHandler('tools/list', async () => {
    return {
      tools: Object.keys(TOOLS).map(name => ({
        name,
        description: `Mock ${name}`,
        inputSchema: { type: 'object', properties: {} },
      })),
    };
  });

  server.setRequestHandler('tools/call', async (request) => {
    const { name, arguments: args } = request.params;
    const handler = TOOLS[name];
    if (!handler) {
      return { content: [{ type: 'text', text: JSON.stringify({ error: `Unknown tool: ${name}` }) }] };
    }
    const result = handler(args || {});
    return { content: [{ type: 'text', text: JSON.stringify(result) }] };
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
  return server;
}

if (process.argv[1] && process.argv[1].endsWith('mock-mcp-server.mjs')) {
  startMockServer().catch(err => {
    console.error('Mock server failed:', err);
    process.exit(1);
  });
}
