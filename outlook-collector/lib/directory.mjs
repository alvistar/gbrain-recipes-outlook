import { readFileSync, writeFileSync, mkdirSync, readdirSync } from 'fs';
import { join } from 'path';
import { createHash } from 'crypto';
import { parseToolResult } from './mcp-utils.mjs';

const CACHE_TTL_MS = 7 * 24 * 60 * 60 * 1000;

async function directCallTool(client, name, args = {}) {
  return client.callTool({ name, arguments: args });
}

export async function probeOrgMode(client) {
  try {
    await directCallTool(client, 'list-users', { $top: '1' });
    return true;
  } catch (err) {
    console.error(
      `[outlook-collector] Directory lookup disabled (requires --org-mode ` +
      `with a work/school account and User.Read.All permission). Error: ${err.message}`
    );
    return false;
  }
}

export async function lookupUser(client, email, cacheDir) {
  mkdirSync(cacheDir, { recursive: true });
  const cacheFile = join(cacheDir, cacheKey(email) + '.json');

  try {
    const cached = JSON.parse(readFileSync(cacheFile, 'utf-8'));
    if (cached.cachedAt && Date.now() - cached.cachedAt < CACHE_TTL_MS) {
      return cached.data;
    }
  } catch {}

  try {
    const result = await directCallTool(client, 'list-users', {
      $filter: `mail eq '${escapeODataString(email)}'`,
    });

    const users = parseToolResult(result);
    const userData = users && users.length > 0 ? users[0] : null;

    writeFileSync(cacheFile, JSON.stringify({
      email,
      data: userData,
      cachedAt: Date.now(),
    }));

    return userData;
  } catch (err) {
    console.error(`[outlook-collector] Directory lookup failed for ${email}: ${err.message}`);
    return null;
  }
}

export async function refreshCache(client, cacheDir) {
  mkdirSync(cacheDir, { recursive: true });
  let files;
  try {
    files = readdirSync(cacheDir).filter(f => f.endsWith('.json'));
  } catch {
    return;
  }

  for (const file of files) {
    try {
      const cached = JSON.parse(readFileSync(join(cacheDir, file), 'utf-8'));
      if (cached.email) {
        await lookupUserFresh(client, cached.email, cacheDir);
      }
    } catch {}
  }
}

export async function enrichWithDirectory(client, messages, cacheDir) {
  const seen = new Map();
  const enriched = [];

  for (const msg of messages) {
    const email = msg.from?.emailAddress?.address;
    if (!email) {
      enriched.push({ ...msg });
      continue;
    }

    if (!seen.has(email)) {
      const userData = await lookupUser(client, email, cacheDir);
      seen.set(email, userData);
    }

    const userData = seen.get(email);
    const copy = { ...msg };
    if (userData) {
      copy._directory = {
        displayName: userData.displayName,
        jobTitle: userData.jobTitle,
        department: userData.department,
        companyName: userData.companyName,
        officeLocation: userData.officeLocation,
      };
    }
    enriched.push(copy);
  }

  return enriched;
}

function escapeODataString(value) {
  return value.replace(/'/g, "''");
}

function cacheKey(email) {
  return createHash('sha256').update(email.toLowerCase()).digest('hex').substring(0, 16);
}

async function lookupUserFresh(client, email, cacheDir) {
  const cacheFile = join(cacheDir, cacheKey(email) + '.json');
  try {
    const result = await directCallTool(client, 'list-users', {
      $filter: `mail eq '${escapeODataString(email)}'`,
    });
    const users = parseToolResult(result);
    const userData = users && users.length > 0 ? users[0] : null;
    writeFileSync(cacheFile, JSON.stringify({
      email,
      data: userData,
      cachedAt: Date.now(),
    }));
    return userData;
  } catch {
    return null;
  }
}

