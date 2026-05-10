import { outlookLink } from './links.mjs';
import { writeFileSync, mkdirSync } from 'fs';
import { join } from 'path';

export function generateDigest(classified, sentIds, date) {
  const dateStr = date.toISOString().split('T')[0];
  const lines = [`# Outlook Digest: ${dateStr}`, ''];

  if (classified.signature.length > 0) {
    lines.push('## Signatures Pending', '');
    for (const msg of classified.signature) {
      lines.push(formatEntry(msg, sentIds));
    }
    lines.push('');
  }

  if (classified.triage.length > 0) {
    lines.push('## Messages to Triage', '');
    for (const msg of classified.triage) {
      lines.push(formatEntry(msg, sentIds));
    }
    lines.push('');
  }

  if (classified.noise.length > 0) {
    lines.push('## Noise', '');
    for (const msg of classified.noise) {
      lines.push(formatEntry(msg, sentIds));
    }
    lines.push('');
  }

  if (classified.signature.length === 0 &&
      classified.triage.length === 0 &&
      classified.noise.length === 0) {
    lines.push('No new messages.', '');
  }

  return lines.join('\n');
}

function formatEntry(msg, sentIds) {
  const sender = msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || 'Unknown';
  const subject = msg.subject || '(no subject)';
  const link = outlookLink(msg);
  const date = msg.receivedDateTime
    ? new Date(msg.receivedDateTime).toISOString().split('T')[0]
    : '';
  const snippet = (msg.bodyPreview || '').substring(0, 200);

  const dirInfo = msg._directory
    ? ` (${msg._directory.jobTitle || ''} @ ${msg._directory.companyName || ''}, ${msg._directory.department || ''})`
    : '';

  const replied = sentIds && msg.conversationId && sentIds.has(msg.conversationId)
    ? ' [replied]'
    : '';

  const lines = [
    `- **${sender}**${dirInfo} — ${subject}${replied}`,
    `  [Open in Outlook](${link}) | ${date}`,
  ];
  if (snippet) {
    lines.push(`  > ${snippet}`);
  }
  return lines.join('\n');
}

export function writeDigest(dataDir, content, date) {
  const digestDir = join(dataDir, 'digests');
  mkdirSync(digestDir, { recursive: true });
  const dateStr = date.toISOString().split('T')[0];
  writeFileSync(join(digestDir, `${dateStr}.md`), content);
}
