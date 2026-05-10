const NOISE_PATTERNS = [
  'noreply', 'no-reply', 'notifications@', 'calendar-notification',
  'mailer-daemon', 'postmaster', 'donotreply',
];

const SIGNATURE_PATTERNS = [
  /docusign/i, /dropbox sign/i, /hellosign/i, /pandadoc/i,
  /please sign/i, /signature needed/i, /ready for your signature/i,
  /everyone has signed/i, /you just signed/i,
];

export function isNoise(message) {
  const from = (message.from?.emailAddress?.address || '').toLowerCase();
  return NOISE_PATTERNS.some(p => from.includes(p));
}

export function isSignature(message) {
  const subject = message.subject || '';
  const from = message.from?.emailAddress?.address || '';
  return SIGNATURE_PATTERNS.some(p => p.test(subject) || p.test(from));
}

export function classify(message) {
  if (isSignature(message)) return 'signature';
  if (isNoise(message)) return 'noise';
  return 'triage';
}

export function classifyMessages(messages) {
  const result = { signature: [], triage: [], noise: [] };
  for (const msg of messages) {
    const category = classify(msg);
    result[category].push(msg);
  }
  return result;
}
