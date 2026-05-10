export function parseToolResult(result) {
  if (!result || !result.content) return null;
  for (const block of result.content) {
    if (block.type === 'text') {
      try {
        const parsed = JSON.parse(block.text);
        return Array.isArray(parsed) ? parsed : parsed.value || [parsed];
      } catch {
        return null;
      }
    }
  }
  return null;
}
