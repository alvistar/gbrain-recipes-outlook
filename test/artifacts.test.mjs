import { describe, test, expect, beforeEach, afterEach } from 'bun:test';
import { mkdtempSync, rmSync, readFileSync, existsSync } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';
import {
  writeCollectionArtifacts,
  readDailyMessages,
} from '../outlook-collector/lib/artifacts.mjs';

let dir;

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'outlook-artifacts-'));
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

function msg(id, folderKey = 'inbox', subject = id) {
  return {
    id,
    subject,
    conversationId: `conv-${id}`,
    receivedDateTime: '2026-05-11T12:00:00Z',
    _folder: { key: folderKey, id: folderKey, displayName: folderKey },
  };
}

describe('collection artifacts', () => {
  test('writes every run separately and maintains cumulative daily aggregate', () => {
    const date = new Date('2026-05-11T12:00:00.000Z');
    const first = writeCollectionArtifacts(dir, {
      messages: [msg('a'), msg('b')],
      sentIds: new Set(['conv-a']),
      date,
      runId: '2026-05-11T12-00-00-000Z',
    });

    expect(existsSync(first.runFile)).toBe(true);
    expect(first.dailyMessages).toHaveLength(2);

    const second = writeCollectionArtifacts(dir, {
      messages: [msg('c', 'archive')],
      sentIds: new Set(['conv-a', 'conv-c']),
      date: new Date('2026-05-11T12:10:00.000Z'),
      runId: '2026-05-11T12-10-00-000Z',
    });

    expect(existsSync(second.runFile)).toBe(true);
    expect(second.dailyMessages.map(m => m.id).sort()).toEqual(['a', 'b', 'c']);
    expect([...second.dailySentIds].sort()).toEqual(['conv-a', 'conv-c']);

    const aggregate = readDailyMessages(dir, date);
    expect(aggregate.messages.map(m => m.id).sort()).toEqual(['a', 'b', 'c']);
    expect(aggregate.runs.map(r => r.runId)).toEqual([
      '2026-05-11T12-00-00-000Z',
      '2026-05-11T12-10-00-000Z',
    ]);
  });

  test('empty later run does not erase earlier daily messages', () => {
    const date = new Date('2026-05-11T12:00:00.000Z');
    writeCollectionArtifacts(dir, {
      messages: [msg('a')],
      sentIds: new Set(['conv-a']),
      date,
      runId: '2026-05-11T12-00-00-000Z',
    });

    writeCollectionArtifacts(dir, {
      messages: [],
      sentIds: new Set(['conv-a', 'conv-reply']),
      date: new Date('2026-05-11T12:10:00.000Z'),
      runId: '2026-05-11T12-10-00-000Z',
    });

    const aggregate = JSON.parse(readFileSync(join(dir, 'messages', '2026-05-11.json'), 'utf-8'));
    expect(aggregate.messages.map(m => m.id)).toEqual(['a']);
    expect(aggregate.sentIds.sort()).toEqual(['conv-a', 'conv-reply']);
    expect(aggregate.runs.at(-1).messageCount).toBe(0);
  });

  test('same message id updates daily aggregate instead of duplicating', () => {
    const date = new Date('2026-05-11T12:00:00.000Z');
    writeCollectionArtifacts(dir, {
      messages: [msg('a', 'inbox', 'old')],
      sentIds: new Set(),
      date,
      runId: '2026-05-11T12-00-00-000Z',
    });

    writeCollectionArtifacts(dir, {
      messages: [msg('a', 'archive', 'updated')],
      sentIds: new Set(),
      date: new Date('2026-05-11T12:10:00.000Z'),
      runId: '2026-05-11T12-10-00-000Z',
    });

    const aggregate = readDailyMessages(dir, date);
    expect(aggregate.messages).toHaveLength(1);
    expect(aggregate.messages[0].subject).toBe('updated');
    expect(aggregate.messages[0]._folder.key).toBe('archive');
  });
});
