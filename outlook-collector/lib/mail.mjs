import { callTool } from './connect.mjs';
import { parseToolResult } from './mcp-utils.mjs';

const DEFAULT_LOOKBACK = '2h';
const DEFAULT_FOLDERS = ['inbox', 'archive', 'sentitems'];
const PAGE_SIZE = 50;

const FOLDER_ALIASES = {
  inbox: ['inbox', 'Inbox', 'Posta in arrivo'],
  archive: ['archive', 'Archive', 'Archivio'],
  sentitems: ['sentitems', 'sentItems', 'SentItems', 'sent items', 'Sent Items', 'Posta inviata'],
};

export function parseLookbackMs(value = DEFAULT_LOOKBACK) {
  const match = String(value).trim().match(/^(\d+)([mhd])$/i);
  if (!match) throw new Error(`Invalid lookback '${value}'. Use values like 30m, 2h, or 7d.`);
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) {
    throw new Error(`Invalid lookback '${value}'. Amount must be positive.`);
  }
  const unit = match[2].toLowerCase();
  const multiplier = unit === 'm' ? 60_000 : unit === 'h' ? 3_600_000 : 86_400_000;
  return amount * multiplier;
}

export function buildLookbackFilter({ dateField, lastCollect, lookback = DEFAULT_LOOKBACK, now = new Date() }) {
  const fallbackSince = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const base = lastCollect ? new Date(lastCollect) : fallbackSince;
  const validBase = Number.isNaN(base.getTime()) ? fallbackSince : base;
  const since = new Date(validBase.getTime() - parseLookbackMs(lookback));
  return `${dateField} ge ${since.toISOString()}`;
}

export function parseFolders(value) {
  if (!value) return [...DEFAULT_FOLDERS];
  if (Array.isArray(value)) return value.map(normalizeFolderKey).filter(Boolean);
  return String(value).split(',').map(normalizeFolderKey).filter(Boolean);
}

function normalizeFolderKey(value) {
  const key = String(value || '').trim().toLowerCase().replace(/[-_\s]/g, '');
  if (!key) return null;
  if (key === 'sent' || key === 'sentmail' || key === 'sentitems') return 'sentitems';
  if (key === 'inbox') return 'inbox';
  if (key === 'archive') return 'archive';
  return key;
}

export async function collectInbox(client, state, options = {}) {
  const { messages } = await collectRecentMail(client, state, {
    ...options,
    folders: options.folders || ['inbox'],
  });
  return messages;
}

export async function collectSent(client, state, options = {}) {
  const { sentIds } = await collectRecentMail(client, state, {
    ...options,
    folders: options.folders || ['sentitems'],
  });
  return sentIds;
}

export async function collectRecentMail(client, state, options = {}) {
  ensureStateShape(state);
  const folders = parseFolders(options.folders || DEFAULT_FOLDERS);
  const lookback = options.lookback || DEFAULT_LOOKBACK;
  const now = options.now || new Date();

  const messages = [];
  const sentIds = new Set();
  const emitted = new Set();

  for (const folderKey of folders) {
    const folder = await resolveFolder(client, state, folderKey);
    const isSent = folderKey === 'sentitems';
    const dateField = isSent ? 'sentDateTime' : 'receivedDateTime';
    const filter = buildLookbackFilter({ dateField, lastCollect: state.lastCollect, lookback, now });
    const items = await listFolderMessages(client, folder.id, filter, dateField);

    for (const item of items) {
      const msg = annotateMessage(item, folderKey, folder);
      const id = msg.id;
      if (!id) continue;

      if (isSent) {
        sentIds.add(msg.conversationId || msg.id);
        updateKnownId(state, msg, folderKey, true);
        continue;
      }

      const wasKnown = Boolean(state.knownIds[id]);
      const changed = updateKnownId(state, msg, folderKey, false);
      if (!emitted.has(id) && (!wasKnown || changed)) {
        if (wasKnown) msg._updatedExisting = true;
        messages.push(msg);
        emitted.add(id);
      }
    }
  }

  return { messages, sentIds };
}

async function listFolderMessages(client, mailFolderId, filter, dateField) {
  const messages = [];
  let skip = 0;
  while (true) {
    const result = await callTool(client, 'list-mail-folder-messages', {
      mailFolderId,
      $filter: filter,
      $top: PAGE_SIZE,
      $skip: skip,
      $orderby: [`${dateField} desc`],
    });
    const items = parseToolResult(result) || [];
    if (!items.length) break;
    messages.push(...items);
    skip += items.length;
    if (items.length < PAGE_SIZE) break;
  }
  return messages;
}

async function resolveFolder(client, state, folderKey) {
  ensureStateShape(state);
  const cached = state.folders?.[folderKey];
  if (cached?.id) return cached;

  const aliases = FOLDER_ALIASES[folderKey] || [folderKey];
  for (const alias of aliases) {
    try {
      const result = await callTool(client, 'list-mail-folder-messages', {
        mailFolderId: alias,
        $top: 1,
      });
      parseToolResult(result);
      const folder = { id: alias, key: folderKey, displayName: alias, resolvedBy: 'wellKnownName' };
      state.folders[folderKey] = folder;
      return folder;
    } catch {}
  }

  try {
    const result = await callTool(client, 'list-mail-folders', { $top: 100 });
    const folders = parseToolResult(result) || [];
    const found = folders.find(folder => {
      const display = String(folder.displayName || folder.name || '').toLowerCase();
      const wellKnown = String(folder.wellKnownName || '').toLowerCase();
      return aliases.some(alias => {
        const a = String(alias).toLowerCase();
        return display === a || wellKnown === a || String(folder.id || '').toLowerCase() === a;
      });
    });
    if (found?.id) {
      const folder = {
        id: found.id,
        key: folderKey,
        displayName: found.displayName || found.name || folderKey,
        wellKnownName: found.wellKnownName,
        resolvedBy: 'list-mail-folders',
      };
      state.folders[folderKey] = folder;
      return folder;
    }
  } catch (err) {
    throw new Error(`Could not resolve Outlook folder '${folderKey}': ${err.message}`);
  }

  throw new Error(`Could not resolve Outlook folder '${folderKey}'. Tried aliases: ${aliases.join(', ')}`);
}

function annotateMessage(msg, folderKey, folder) {
  return {
    ...msg,
    _folder: {
      key: folderKey,
      id: folder.id,
      displayName: folder.displayName || folderKey,
    },
  };
}

function ensureStateShape(state) {
  state.knownIds ||= {};
  state.folders ||= {};
}

function updateKnownId(state, msg, folderKey, isSent) {
  const now = Date.now();
  const previous = state.knownIds[msg.id];
  const next = {
    ...(typeof previous === 'object' ? previous : {}),
    seenAt: typeof previous === 'object' && previous.seenAt ? previous.seenAt : now,
    lastSeenAt: now,
    folderKey,
    parentFolderId: msg.parentFolderId || msg._folder?.id,
    conversationId: msg.conversationId,
    internetMessageId: msg.internetMessageId,
    changeKey: msg.changeKey,
    lastModifiedDateTime: msg.lastModifiedDateTime,
    isSent: Boolean(isSent),
  };

  const changed = !previous ||
    typeof previous !== 'object' ||
    previous.folderKey !== next.folderKey ||
    previous.parentFolderId !== next.parentFolderId ||
    previous.changeKey !== next.changeKey ||
    previous.lastModifiedDateTime !== next.lastModifiedDateTime;

  state.knownIds[msg.id] = next;
  return changed;
}
