#!/usr/bin/env node

import { probeServer, createClient, disconnect } from './lib/connect.mjs';
import { loadState, saveState } from './lib/state.mjs';
import { collectInbox, collectSent } from './lib/mail.mjs';
import { classifyMessages } from './lib/filter.mjs';
import { probeOrgMode, enrichWithDirectory } from './lib/directory.mjs';
import { generateDigest, writeDigest } from './lib/digest.mjs';
import { writeFileSync, readFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = join(SCRIPT_DIR, '..', 'data');
const CACHE_DIR = join(DATA_DIR, 'directory');

async function collect() {
  probeServer();

  const state = loadState(DATA_DIR);
  const { client } = await createClient();

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
    const messagesFile = join(DATA_DIR, 'messages', `${today}.json`);
    writeFileSync(
      messagesFile,
      JSON.stringify({ messages: enrichedMessages, sentIds: [...sentIds] }, null, 2)
    );

    const classified = classifyMessages(enrichedMessages);
    const digestContent = generateDigest(classified, sentIds, new Date());
    writeDigest(DATA_DIR, digestContent, new Date());

    state.lastCollect = new Date().toISOString();
    saveState(DATA_DIR, state);

    console.log(`Collected ${messages.length} new messages. Digest generated.`);
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
    console.error('Usage: collector.mjs <collect|digest>');
    process.exit(1);
}
