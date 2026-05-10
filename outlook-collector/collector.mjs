#!/usr/bin/env node

import { probeServer, createClient, disconnect } from './lib/connect.mjs';
import { loadState, saveState } from './lib/state.mjs';
import { collectInbox, collectSent } from './lib/mail.mjs';
import { classifyMessages } from './lib/filter.mjs';
import { probeOrgMode, enrichWithDirectory } from './lib/directory.mjs';
import { generateDigest, writeDigest } from './lib/digest.mjs';
import { writeFileSync, mkdirSync } from 'fs';
import { join } from 'path';

const DATA_DIR = join(process.cwd(), 'data');
const CACHE_DIR = join(DATA_DIR, 'directory');

async function collect() {
  probeServer();

  const state = loadState(DATA_DIR);
  const { client, transport } = await createClient();

  try {
    const messages = await collectInbox(client, state);
    const sentIds = await collectSent(client, state);

    const orgMode = await probeOrgMode(client);

    let enrichedMessages = messages;
    if (orgMode) {
      enrichedMessages = await enrichWithDirectory(client, messages, CACHE_DIR);
    }

    mkdirSync(join(DATA_DIR, 'messages'), { recursive: true });
    const today = new Date().toISOString().split('T')[0];
    writeFileSync(
      join(DATA_DIR, 'messages', `${today}.json`),
      JSON.stringify(enrichedMessages, null, 2)
    );

    state.lastCollect = new Date().toISOString();
    saveState(DATA_DIR, state);

    console.log(`Collected ${messages.length} new messages.`);
    return { messages: enrichedMessages, sentIds };
  } finally {
    await disconnect({ client, transport });
  }
}

async function digest() {
  const state = loadState(DATA_DIR);
  const { client, transport } = await createClient();

  try {
    const messages = await collectInbox(client, state);
    const sentIds = await collectSent(client, state);

    const orgMode = await probeOrgMode(client);
    let enrichedMessages = messages;
    if (orgMode) {
      enrichedMessages = await enrichWithDirectory(client, messages, CACHE_DIR);
    }

    const classified = classifyMessages(enrichedMessages);
    const digestContent = generateDigest(classified, sentIds, new Date());
    writeDigest(DATA_DIR, digestContent, new Date());

    state.lastCollect = new Date().toISOString();
    saveState(DATA_DIR, state);

    console.log(`Digest generated with ${messages.length} messages.`);
  } finally {
    await disconnect({ client, transport });
  }
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
    console.error('Usage: collector.mjs <collect|digest>');
    process.exit(1);
}
