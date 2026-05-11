export function parseToolResult(result) {
  if (!result || !result.content) return null;
  for (const block of result.content) {
    if (block.type === 'text') {
      let parsed;
      try {
        parsed = JSON.parse(block.text);
      } catch {
        return null;
      }
      if (parsed && typeof parsed === 'object' && parsed.error) {
        throw new Error(parsed.error);
      }
      return Array.isArray(parsed) ? parsed : parsed.value || [parsed];
    }
  }
  return null;
}
