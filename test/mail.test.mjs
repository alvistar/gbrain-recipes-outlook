import { describe, test, expect } from 'bun:test';
import {
  collectRecentMail,
  buildLookbackFilter,
  parseLookbackMs,
} from '../outlook-collector/lib/mail.mjs';

function makeClient(responses, calls) {
  return {
    async callTool({ name, arguments: args }) {
      calls.push({ name, args });
      const key = args.mailFolderId || name;
      const value = responses[key] || [];
      return { content: [{ type: 'text', text: JSON.stringify({ value }) }] };
    },
  };
}

function msg(id, folderId, overrides = {}) {
  return {
    id,
    parentFolderId: folderId,
    conversationId: `conv-${id}`,
    receivedDateTime: '2026-05-11T12:00:00Z',
    sentDateTime: '2026-05-11T12:01:00Z',
    lastModifiedDateTime: '2026-05-11T12:02:00Z',
    subject: `Subject ${id}`,
    ...overrides,
  };
}

describe('mail lookback helpers', () => {
  test('parseLookbackMs accepts minutes, hours, and days', () => {
    expect(parseLookbackMs('10m')).toBe(10 * 60 * 1000);
    expect(parseLookbackMs('2h')).toBe(2 * 60 * 60 * 1000);
    expect(parseLookbackMs('7d')).toBe(7 * 24 * 60 * 60 * 1000);
  });

  test('parseLookbackMs rejects invalid lookback values', () => {
    expect(() => parseLookbackMs('soon')).toThrow(/Invalid lookback/);
    expect(() => parseLookbackMs('0m')).toThrow(/Invalid lookback/);
  });

  test('buildLookbackFilter overlaps lastCollect by the requested safety window', () => {
    const filter = buildLookbackFilter({
      dateField: 'receivedDateTime',
      lastCollect: '2026-05-11T12:00:00.000Z',
      lookback: '30m',
      now: new Date('2026-05-11T13:00:00.000Z'),
    });
    expect(filter).toBe('receivedDateTime ge 2026-05-11T11:30:00.000Z');
  });
});

describe('collectRecentMail', () => {
  test('collects inbox, archive, and sent folders by mailFolderId with overlap filters', async () => {
    const calls = [];
    const client = makeClient({
      inbox: [msg('inbox-1', 'inbox')],
      archive: [msg('archive-1', 'archive')],
      sentitems: [msg('sent-1', 'sentitems', { conversationId: 'conv-inbox-1' })],
    }, calls);
    const state = {
      lastCollect: '2026-05-11T12:00:00.000Z',
      knownIds: {},
      folders: {
        inbox: { id: 'inbox' },
        archive: { id: 'archive' },
        sentitems: { id: 'sentitems' },
      },
    };

    const result = await collectRecentMail(client, state, {
      folders: ['inbox', 'archive', 'sentitems'],
      lookback: '30m',
      now: new Date('2026-05-11T13:00:00.000Z'),
    });

    expect(result.messages.map(m => m.id).sort()).toEqual(['archive-1', 'inbox-1']);
    expect([...result.sentIds]).toEqual(['conv-inbox-1']);
    expect(calls.map(c => c.args.mailFolderId)).toEqual(['inbox', 'archive', 'sentitems']);
    expect(calls[0].args.$filter).toBe('receivedDateTime ge 2026-05-11T11:30:00.000Z');
    expect(calls[2].args.$filter).toBe('sentDateTime ge 2026-05-11T11:30:00.000Z');
  });

  test('deduplicates repeated messages while updating folder metadata', async () => {
    const calls = [];
    const client = makeClient({
      inbox: [msg('same-id', 'inbox')],
      archive: [msg('same-id', 'archive', { lastModifiedDateTime: '2026-05-11T12:10:00Z' })],
    }, calls);
    const state = {
      lastCollect: '2026-05-11T12:00:00.000Z',
      knownIds: {},
      folders: {
        inbox: { id: 'inbox' },
        archive: { id: 'archive' },
      },
    };

    const result = await collectRecentMail(client, state, {
      folders: ['inbox', 'archive'],
      lookback: '30m',
      now: new Date('2026-05-11T13:00:00.000Z'),
    });

    expect(result.messages).toHaveLength(1);
    expect(state.knownIds['same-id'].folderKey).toBe('archive');
    expect(state.knownIds['same-id'].parentFolderId).toBe('archive');
    expect(state.knownIds['same-id'].lastModifiedDateTime).toBe('2026-05-11T12:10:00Z');
  });

  test('returns updated known messages moved into archive during lookback', async () => {
    const calls = [];
    const client = makeClient({
      archive: [msg('known-moved', 'archive', { changeKey: 'v2' })],
    }, calls);
    const state = {
      lastCollect: '2026-05-11T12:00:00.000Z',
      knownIds: {
        'known-moved': { seenAt: 1, folderKey: 'inbox', parentFolderId: 'inbox', changeKey: 'v1' },
      },
      folders: { archive: { id: 'archive' } },
    };

    const result = await collectRecentMail(client, state, {
      folders: ['archive'],
      lookback: '30m',
      now: new Date('2026-05-11T13:00:00.000Z'),
    });

    expect(result.messages.map(m => m.id)).toEqual(['known-moved']);
    expect(result.messages[0]._folder.key).toBe('archive');
    expect(result.messages[0]._updatedExisting).toBe(true);
    expect(state.knownIds['known-moved'].changeKey).toBe('v2');
  });
});
