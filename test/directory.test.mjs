import { describe, test, expect, beforeEach, afterEach } from 'bun:test';
import { lookupUser, probeOrgMode } from '../outlook-collector/lib/directory.mjs';
import { mkdirSync, rmSync, writeFileSync, readFileSync, existsSync } from 'fs';
import { join } from 'path';

const TEST_CACHE_DIR = join(import.meta.dir, '.tmp-directory-test');

beforeEach(() => {
  mkdirSync(TEST_CACHE_DIR, { recursive: true });
});

afterEach(() => {
  rmSync(TEST_CACHE_DIR, { recursive: true, force: true });
});

function makeMockClient(responses = {}) {
  return {
    callTool: async (request) => {
      const name = request.name;
      const args = request.arguments || {};
      const handler = responses[name];
      if (!handler) throw new Error(`No mock for ${name}`);
      return handler(args);
    },
  };
}

function makeUserResponse(users) {
  return {
    content: [{ type: 'text', text: JSON.stringify({ value: users }) }],
  };
}

describe('probeOrgMode', () => {
  test('returns true when list-users succeeds', async () => {
    const client = makeMockClient({
      'list-users': () => makeUserResponse([{ displayName: 'Test' }]),
    });
    const result = await probeOrgMode(client);
    expect(result).toBe(true);
  });

  test('returns false and logs warning on permission error', async () => {
    const client = makeMockClient({
      'list-users': () => { throw new Error('403 Insufficient privileges'); },
    });
    const result = await probeOrgMode(client);
    expect(result).toBe(false);
  });
});

describe('lookupUser', () => {
  test('returns user data from API', async () => {
    const client = makeMockClient({
      'list-users': () => makeUserResponse([
        { displayName: 'Alice', jobTitle: 'VP Eng', department: 'Engineering', mail: 'alice@acme.com' },
      ]),
    });
    const result = await lookupUser(client, 'alice@acme.com', TEST_CACHE_DIR);
    expect(result.displayName).toBe('Alice');
    expect(result.jobTitle).toBe('VP Eng');
  });

  test('returns cached data on cache hit (no API call)', async () => {
    const cacheFile = join(TEST_CACHE_DIR, '0a0a58273565a8f3.json');
    writeFileSync(cacheFile, JSON.stringify({
      email: 'alice@acme.com',
      data: { displayName: 'Cached Alice', jobTitle: 'CTO' },
      cachedAt: Date.now(),
    }));

    let apiCalled = false;
    const client = makeMockClient({
      'list-users': () => { apiCalled = true; return makeUserResponse([]); },
    });

    const result = await lookupUser(client, 'alice@acme.com', TEST_CACHE_DIR);
    expect(result.displayName).toBe('Cached Alice');
    expect(apiCalled).toBe(false);
  });

  test('re-fetches on expired cache (7d+ old)', async () => {
    const oldTimestamp = Date.now() - 8 * 24 * 60 * 60 * 1000;
    const cacheFile = join(TEST_CACHE_DIR, '0a0a58273565a8f3.json');
    writeFileSync(cacheFile, JSON.stringify({
      email: 'alice@acme.com',
      data: { displayName: 'Old Alice', jobTitle: 'Engineer' },
      cachedAt: oldTimestamp,
    }));

    const client = makeMockClient({
      'list-users': () => makeUserResponse([
        { displayName: 'Updated Alice', jobTitle: 'VP Eng' },
      ]),
    });

    const result = await lookupUser(client, 'alice@acme.com', TEST_CACHE_DIR);
    expect(result.displayName).toBe('Updated Alice');
  });

  test('returns null for external sender not in directory', async () => {
    const client = makeMockClient({
      'list-users': () => makeUserResponse([]),
    });
    const result = await lookupUser(client, 'external@other.com', TEST_CACHE_DIR);
    expect(result).toBeNull();
  });

  test('caches negative result (no re-lookup on next call)', async () => {
    let callCount = 0;
    const client = makeMockClient({
      'list-users': () => { callCount++; return makeUserResponse([]); },
    });

    await lookupUser(client, 'nobody@nowhere.com', TEST_CACHE_DIR);
    await lookupUser(client, 'nobody@nowhere.com', TEST_CACHE_DIR);
    expect(callCount).toBe(1);
  });
});
