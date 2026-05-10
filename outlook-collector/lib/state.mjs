import { readFileSync, writeFileSync, renameSync, mkdirSync } from 'fs';
import { join } from 'path';
import { randomBytes } from 'crypto';

const PRUNE_AGE_MS = 30 * 24 * 60 * 60 * 1000;

export function loadState(dataDir) {
  const statePath = join(dataDir, 'state.json');
  try {
    const raw = readFileSync(statePath, 'utf-8');
    return JSON.parse(raw);
  } catch (err) {
    if (err.code === 'ENOENT') {
      return freshState();
    }
    const corruptPath = statePath + '.corrupt';
    try {
      renameSync(statePath, corruptPath);
    } catch {}
    console.error(`[outlook-collector] state.json corrupted, renamed to state.json.corrupt. Starting fresh.`);
    return freshState();
  }
}

export function saveState(dataDir, state) {
  mkdirSync(dataDir, { recursive: true });
  const statePath = join(dataDir, 'state.json');
  const pruned = pruneOldIds(state);
  const tmpPath = join(dataDir, `.state-${randomBytes(4).toString('hex')}.tmp`);
  writeFileSync(tmpPath, JSON.stringify(pruned, null, 2));
  renameSync(tmpPath, statePath);
}

function freshState() {
  return {
    lastCollect: null,
    knownIds: {},
  };
}

function pruneOldIds(state) {
  const cutoff = Date.now() - PRUNE_AGE_MS;
  const pruned = {};
  for (const [id, entry] of Object.entries(state.knownIds || {})) {
    const ts = typeof entry === 'object' ? entry.seenAt : entry;
    if (typeof ts === 'number' && ts >= cutoff) {
      pruned[id] = entry;
    }
  }
  return { ...state, knownIds: pruned };
}
