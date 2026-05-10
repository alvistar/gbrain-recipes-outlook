export function outlookLink(message, accountType = 'work') {
  if (message.webLink) {
    return message.webLink;
  }

  const id = encodeURIComponent(message.id);
  if (accountType === 'personal') {
    return `https://outlook.live.com/mail/0/deeplink/read/${id}`;
  }
  return `https://outlook.office365.com/mail/deeplink/read/${id}`;
}
