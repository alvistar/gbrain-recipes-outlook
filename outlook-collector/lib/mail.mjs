import { callTool } from './connect.mjs';
import { parseToolResult } from './mcp-utils.mjs';

export async function collectInbox(client, state) {
  const filter = state.lastCollect
    ? `receivedDateTime ge ${state.lastCollect}`
    : `receivedDateTime ge ${new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString()}`;

  const messages = [];
  let hasMore = true;
  let skip = 0;

  while (hasMore) {
    const result = await callTool(client, 'list-mail-messages', {
      $filter: filter,
      $top: '50',
      $skip: String(skip),
      $orderby: 'receivedDateTime desc',
    });

    const items = parseToolResult(result);
    if (!items || items.length === 0) {
      hasMore = false;
      break;
    }

    for (const msg of items) {
      const id = msg.id;
      if (state.knownIds[id]) continue;
      messages.push(msg);
      state.knownIds[id] = { seenAt: Date.now() };
    }

    skip += items.length;
    if (items.length < 50) hasMore = false;
  }

  return messages;
}

export async function collectSent(client, state) {
  const filter = state.lastCollect
    ? `sentDateTime ge ${state.lastCollect}`
    : `sentDateTime ge ${new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString()}`;

  const sentIds = new Set();

  try {
    const result = await callTool(client, 'list-mail-folder-messages', {
      folderName: 'SentItems',
      $filter: filter,
      $top: '30',
    });

    const items = parseToolResult(result);
    if (items) {
      for (const msg of items) {
        sentIds.add(msg.conversationId || msg.id);
        if (!state.knownIds[msg.id]) {
          state.knownIds[msg.id] = { seenAt: Date.now(), isSent: true };
        }
      }
    }
  } catch (err) {
    console.error(`[outlook-collector] Could not fetch sent mail: ${err.message}`);
  }

  return sentIds;
}

