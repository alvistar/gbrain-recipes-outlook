#!/usr/bin/env node

import { probeServer, createClient, disconnect } from './lib/connect.mjs';
import { loadState, saveState } from './lib/state.mjs';
import { collectRecentMail, parseFolders } from './lib/mail.mjs';
import { classifyMessages } from './lib/filter.mjs';
import { probeOrgMode, enrichWithDirectory } from './lib/directory.mjs';
import { generateDigest, writeDigest } from './lib/digest.mjs';
import { writeCollectionArtifacts } from './lib/artifacts.mjs';
import { readFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = join(SCRIPT_DIR, '..', 'data');
const CACHE_DIR = join(DATA_DIR, 'directory');

function parseCollectOptions(argv = process.argv.slice(3)) {
  const options = { folders: undefined, lookback: undefined };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--folders') {
      options.folders = parseFolders(argv[++i]);
    } else if (arg.startsWith('--folders=')) {
      options.folders = parseFolders(arg.slice('--folders='.length));
    } else if (arg === '--lookback') {
      options.lookback = argv[++i];
    } else if (arg.startsWith('--lookback=')) {
      options.lookback = arg.slice('--lookback='.length);
    } else {
      throw new Error(`Unknown collect option: ${arg}`);
    }
  }
  return options;
}

async function collect() {
  probeServer();

  const options = parseCollectOptions();
  const state = loadState(DATA_DIR);
  const { client } = await createClient();

  try {
    const { messages, sentIds } = await collectRecentMail(client, state, options);

    const orgMode = await probeOrgMode(client);

    let enrichedMessages = messages;
    if (orgMode) {
      enrichedMessages = await enrichWithDirectory(client, messages, CACHE_DIR);
    }

    const now = new Date();
    const artifacts = writeCollectionArtifacts(DATA_DIR, {
      messages: enrichedMessages,
      sentIds,
      date: now,
    });

    const classified = classifyMessages(artifacts.dailyMessages);
    const digestContent = generateDigest(classified, artifacts.dailySentIds, now);
    writeDigest(DATA_DIR, digestContent, now);

    state.lastCollect = now.toISOString();
    saveState(DATA_DIR, state);

    const folderList = parseFolders(options.folders).join(',');
    console.log(`Collected ${messages.length} new/updated messages from ${folderList}. Daily aggregate has ${artifacts.dailyMessages.length} messages. Digest generated.`);
  } finally {
    await disconnect({ client });
  }
}

async function digest() {
  const today = new Date().toISOString().split('T')[0];
  const messagesFile = join(DATA_DIR, 'messages', `${today}.json`);

  let data;
  try {
    data = JSON.parse(readFileSync(messagesFile, 'utf-8'));
  } catch {
    console.error(`[outlook-collector] No collected messages found for ${today}. Run 'collect' first.`);
    process.exit(1);
  }

  const messages = data.messages || [];
  const sentIds = new Set(data.sentIds || []);

  const classified = classifyMessages(messages);
  const digestContent = generateDigest(classified, sentIds, new Date());
  writeDigest(DATA_DIR, digestContent, new Date());

  console.log(`Digest regenerated from ${messages.length} collected messages.`);
}

const command = process.argv[2];

switch (command) {
  case 'collect':
    collect().catch(err => {
      console.error(`[outlook-collector] collect failed: ${err.message}`);
      process.exit(1);
    });
    break;
  case 'digest':
    digest().catch(err => {
      console.error(`[outlook-collector] digest failed: ${err.message}`);
      process.exit(1);
    });
    break;
  default:
    console.error('Usage: collector.mjs <collect|digest> [--folders inbox,archive,sentitems] [--lookback 2h]');
    process.exit(1);
}
