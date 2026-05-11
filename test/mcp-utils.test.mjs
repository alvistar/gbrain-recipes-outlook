import { describe, test, expect } from 'bun:test';
import { parseToolResult } from '../outlook-collector/lib/mcp-utils.mjs';

describe('parseToolResult', () => {
  test('returns array values from MCP text blocks', () => {
    const result = {
      content: [{ type: 'text', text: JSON.stringify({ value: [{ id: 'msg-1' }] }) }],
    };
    expect(parseToolResult(result)).toEqual([{ id: 'msg-1' }]);
  });

  test('throws on MCP error payloads instead of treating them as messages', () => {
    const result = {
      content: [{ type: 'text', text: JSON.stringify({ error: 'No accounts found. Please login first.' }) }],
    };
    expect(() => parseToolResult(result)).toThrow('No accounts found. Please login first.');
  });
});
