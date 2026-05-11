import { mkdirSync, readFileSync, writeFileSync } from 'fs';
import { join } from 'path';

export function dateKey(date = new Date()) {
  return date.toISOString().split('T')[0];
}

export function makeRunId(date = new Date()) {
  return date.toISOString().replace(/:/g, '-').replace(/\./g, '-');
}

export function readDailyMessages(dataDir, date = new Date()) {
  const file = join(dataDir, 'messages', `${dateKey(date)}.json`);
  try {
    const parsed = JSON.parse(readFileSync(file, 'utf-8'));
    return normalizeDaily(parsed);
  } catch (err) {
    if (err.code === 'ENOENT') {
      return { messages: [], sentIds: [], runs: [] };
    }
    throw err;
  }
}

export function writeCollectionArtifacts(dataDir, { messages, sentIds, date = new Date(), runId = makeRunId(date) }) {
  const day = dateKey(date);
  const runDir = join(dataDir, 'runs', day);
  const messagesDir = join(dataDir, 'messages');
  mkdirSync(runDir, { recursive: true });
  mkdirSync(messagesDir, { recursive: true });

  const sentArray = [...(sentIds || [])];
  const runPayload = {
    runId,
    collectedAt: date.toISOString(),
    messages: messages || [],
    sentIds: sentArray,
  };
  const runFile = join(runDir, `${runId}.json`);
  writeFileSync(runFile, JSON.stringify(runPayload, null, 2));

  const existing = readDailyMessages(dataDir, date);
  const mergedMessages = mergeMessages(existing.messages, messages || []);
  const mergedSentIds = [...new Set([...(existing.sentIds || []), ...sentArray])];
  const runs = [
    ...(existing.runs || []).filter(r => r.runId !== runId),
    {
      runId,
      collectedAt: date.toISOString(),
      messageCount: (messages || []).length,
      sentIdCount: sentArray.length,
      file: `data/runs/${day}/${runId}.json`,
    },
  ].sort((a, b) => String(a.collectedAt || a.runId).localeCompare(String(b.collectedAt || b.runId)));

  const dailyPayload = {
    date: day,
    updatedAt: date.toISOString(),
    messages: mergedMessages,
    sentIds: mergedSentIds,
    runs,
  };
  const dailyFile = join(messagesDir, `${day}.json`);
  writeFileSync(dailyFile, JSON.stringify(dailyPayload, null, 2));

  return {
    runFile,
    dailyFile,
    dailyMessages: mergedMessages,
    dailySentIds: new Set(mergedSentIds),
    runs,
  };
}

function normalizeDaily(parsed) {
  if (Array.isArray(parsed)) {
    return { messages: parsed, sentIds: [], runs: [] };
  }
  return {
    messages: parsed.messages || [],
    sentIds: parsed.sentIds || [],
    runs: parsed.runs || [],
  };
}

function mergeMessages(existing, incoming) {
  const byId = new Map();
  for (const msg of existing || []) {
    if (msg?.id) byId.set(msg.id, msg);
  }
  for (const msg of incoming || []) {
    if (!msg?.id) continue;
    byId.set(msg.id, { ...(byId.get(msg.id) || {}), ...msg });
  }
  return [...byId.values()].sort(compareMessages);
}

function compareMessages(a, b) {
  const ad = a.receivedDateTime || a.sentDateTime || a.lastModifiedDateTime || '';
  const bd = b.receivedDateTime || b.sentDateTime || b.lastModifiedDateTime || '';
  return bd.localeCompare(ad) || String(a.id).localeCompare(String(b.id));
}
