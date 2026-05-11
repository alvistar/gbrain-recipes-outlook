import { describe, test, expect, beforeEach, afterEach } from 'bun:test';
import { loadState, saveState } from '../outlook-collector/lib/state.mjs';
import { mkdirSync, rmSync, readFileSync, writeFileSync, existsSync } from 'fs';
import { join } from 'path';

const TEST_DIR = join(import.meta.dir, '.tmp-state-test');

beforeEach(() => {
  mkdirSync(TEST_DIR, { recursive: true });
});

afterEach(() => {
  rmSync(TEST_DIR, { recursive: true, force: true });
});

describe('loadState', () => {
  test('returns fresh state when file is missing (first run)', () => {
    const state = loadState(TEST_DIR);
    expect(state.lastCollect).toBeNull();
    expect(state.knownIds).toEqual({});
    expect(state.folders).toEqual({});
  });

  test('loads valid state from file', () => {
    const existing = { lastCollect: '2026-05-01T00:00:00Z', knownIds: { 'msg-1': { seenAt: Date.now() } } };
    writeFileSync(join(TEST_DIR, 'state.json'), JSON.stringify(existing));
    const state = loadState(TEST_DIR);
    expect(state.lastCollect).toBe('2026-05-01T00:00:00Z');
    expect(state.knownIds['msg-1']).toBeDefined();
  });

  test('renames corrupt file and returns fresh state', () => {
    writeFileSync(join(TEST_DIR, 'state.json'), 'not valid json {{{');
    const state = loadState(TEST_DIR);
    expect(state.lastCollect).toBeNull();
    expect(state.knownIds).toEqual({});
    expect(existsSync(join(TEST_DIR, 'state.json.corrupt'))).toBe(true);
  });
});

describe('saveState', () => {
  test('writes state atomically (file exists after save)', () => {
    const state = { lastCollect: '2026-05-10T00:00:00Z', knownIds: { 'msg-1': { seenAt: Date.now() } } };
    saveState(TEST_DIR, state);
    const loaded = JSON.parse(readFileSync(join(TEST_DIR, 'state.json'), 'utf-8'));
    expect(loaded.lastCollect).toBe('2026-05-10T00:00:00Z');
    expect(loaded.knownIds['msg-1']).toBeDefined();
  });

  test('prunes IDs older than 30 days', () => {
    const oldTimestamp = Date.now() - 31 * 24 * 60 * 60 * 1000;
    const recentTimestamp = Date.now();
    const state = {
      lastCollect: '2026-05-10T00:00:00Z',
      knownIds: {
        'old-msg': { seenAt: oldTimestamp },
        'new-msg': { seenAt: recentTimestamp },
      },
    };
    saveState(TEST_DIR, state);
    const loaded = JSON.parse(readFileSync(join(TEST_DIR, 'state.json'), 'utf-8'));
    expect(loaded.knownIds['old-msg']).toBeUndefined();
    expect(loaded.knownIds['new-msg']).toBeDefined();
  });

  test('creates directory if it does not exist', () => {
    const nested = join(TEST_DIR, 'nested', 'dir');
    const state = { lastCollect: null, knownIds: {} };
    saveState(nested, state);
    expect(existsSync(join(nested, 'state.json'))).toBe(true);
  });
});
