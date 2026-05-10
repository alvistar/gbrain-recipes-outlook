import { describe, test, expect } from 'bun:test';
import { generateDigest } from '../outlook-collector/lib/digest.mjs';

function makeMessage(overrides = {}) {
  return {
    id: 'msg-test',
    subject: 'Test subject',
    from: { emailAddress: { name: 'Test User', address: 'test@example.com' } },
    receivedDateTime: '2026-05-10T09:00:00Z',
    bodyPreview: 'This is a preview of the email body.',
    conversationId: 'conv-test',
    webLink: 'https://outlook.office365.com/mail/inbox/id/msg-test',
    ...overrides,
  };
}

describe('generateDigest', () => {
  test('full digest with all three sections populated', () => {
    const classified = {
      signature: [makeMessage({ id: 'sig-1', subject: 'DocuSign: Sign NDA', from: { emailAddress: { name: 'DocuSign', address: 'docusign@docusign.net' } } })],
      triage: [makeMessage({ id: 'tri-1', subject: 'Q3 roadmap review' })],
      noise: [makeMessage({ id: 'noi-1', subject: 'Weekly digest', from: { emailAddress: { name: 'System', address: 'noreply@service.com' } } })],
    };
    const digest = generateDigest(classified, new Set(), new Date('2026-05-10'));
    expect(digest).toContain('## Signatures Pending');
    expect(digest).toContain('## Messages to Triage');
    expect(digest).toContain('## Noise');
    expect(digest).toContain('DocuSign: Sign NDA');
    expect(digest).toContain('Q3 roadmap review');
  });

  test('empty collection produces "No new messages"', () => {
    const classified = { signature: [], triage: [], noise: [] };
    const digest = generateDigest(classified, new Set(), new Date('2026-05-10'));
    expect(digest).toContain('No new messages');
    expect(digest).not.toContain('## Signatures Pending');
  });

  test('directory-enriched entry includes title and department', () => {
    const msg = makeMessage({
      _directory: {
        displayName: 'Alice Smith',
        jobTitle: 'VP of Engineering',
        department: 'Engineering',
        companyName: 'Acme Corp',
      },
    });
    const classified = { signature: [], triage: [msg], noise: [] };
    const digest = generateDigest(classified, new Set(), new Date('2026-05-10'));
    expect(digest).toContain('VP of Engineering');
    expect(digest).toContain('Acme Corp');
    expect(digest).toContain('Engineering');
  });

  test('entry without directory data omits metadata', () => {
    const msg = makeMessage();
    const classified = { signature: [], triage: [msg], noise: [] };
    const digest = generateDigest(classified, new Set(), new Date('2026-05-10'));
    expect(digest).not.toContain('VP of Engineering');
    expect(digest).toContain('Test User');
  });

  test('sent mail marks thread as replied in triage', () => {
    const msg = makeMessage({ conversationId: 'conv-replied' });
    const sentIds = new Set(['conv-replied']);
    const classified = { signature: [], triage: [msg], noise: [] };
    const digest = generateDigest(classified, sentIds, new Date('2026-05-10'));
    expect(digest).toContain('[replied]');
  });

  test('signature entries appear in Signatures Pending only', () => {
    const sig = makeMessage({ subject: 'Please sign this' });
    const classified = { signature: [sig], triage: [], noise: [] };
    const digest = generateDigest(classified, new Set(), new Date('2026-05-10'));
    expect(digest).toContain('## Signatures Pending');
    expect(digest).not.toContain('## Messages to Triage');
  });
});
