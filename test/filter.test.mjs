import { describe, test, expect } from 'bun:test';
import { isNoise, isSignature, classify } from '../outlook-collector/lib/filter.mjs';
import { outlookLink } from '../outlook-collector/lib/links.mjs';

function makeMessage(overrides = {}) {
  return {
    id: 'msg-test',
    subject: 'Test subject',
    from: { emailAddress: { name: 'Test', address: 'test@example.com' } },
    ...overrides,
  };
}

describe('noise filtering', () => {
  test('noreply sender classified as noise', () => {
    const msg = makeMessage({ from: { emailAddress: { name: 'System', address: 'noreply@test.com' } } });
    expect(isNoise(msg)).toBe(true);
    expect(classify(msg)).toBe('noise');
  });

  test('notifications@ sender classified as noise (substring match)', () => {
    const msg = makeMessage({ from: { emailAddress: { name: 'Slack', address: 'notifications@slack.com' } } });
    expect(isNoise(msg)).toBe(true);
  });

  test('normal sender passes through as triage', () => {
    const msg = makeMessage({ from: { emailAddress: { name: 'Alice', address: 'alice@acme.com' } } });
    expect(isNoise(msg)).toBe(false);
    expect(classify(msg)).toBe('triage');
  });
});

describe('signature detection', () => {
  test('DocuSign sender classified as signature', () => {
    const msg = makeMessage({ from: { emailAddress: { name: 'DocuSign', address: 'docusign@docusign.net' } } });
    expect(isSignature(msg)).toBe(true);
    expect(classify(msg)).toBe('signature');
  });

  test('"Please sign" in subject classified as signature', () => {
    const msg = makeMessage({ subject: 'Please sign this document' });
    expect(isSignature(msg)).toBe(true);
  });

  test('email with "docusign" in body but not sender/subject is triage', () => {
    const msg = makeMessage({
      subject: 'Contract update',
      from: { emailAddress: { name: 'Legal', address: 'legal@acme.com' } },
    });
    expect(isSignature(msg)).toBe(false);
    expect(classify(msg)).toBe('triage');
  });
});

describe('outlook link generation', () => {
  test('returns webLink when present', () => {
    const msg = makeMessage({ webLink: 'https://outlook.office365.com/mail/inbox/id/msg-001' });
    expect(outlookLink(msg)).toBe('https://outlook.office365.com/mail/inbox/id/msg-001');
  });

  test('constructs work/school URL when webLink is absent', () => {
    const msg = makeMessage({ id: 'msg-abc-123' });
    const link = outlookLink(msg, 'work');
    expect(link).toContain('outlook.office365.com');
    expect(link).toContain('deeplink/read/');
  });

  test('constructs personal URL for personal account', () => {
    const msg = makeMessage({ id: 'msg-abc-123' });
    const link = outlookLink(msg, 'personal');
    expect(link).toContain('outlook.live.com');
    expect(link).toContain('deeplink/read/');
  });
});
