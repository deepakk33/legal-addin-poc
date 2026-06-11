// Simple character-window chunker for the map-reduce distillation path on large
// docs. Splits on paragraph boundaries where possible so clauses aren't cut
// mid-sentence. No tokenizer dependency — char windows are good enough for the
// "condense each section" step.
export function chunk(text: string, size = 12000, overlap = 500): string[] {
  if (text.length <= size) return [text];
  const chunks: string[] = [];
  let start = 0;
  while (start < text.length) {
    let end = Math.min(start + size, text.length);
    if (end < text.length) {
      // Prefer to break at a paragraph/sentence boundary within the last 1000 chars.
      const window = text.slice(end - 1000, end);
      const para = window.lastIndexOf("\n\n");
      const sent = window.lastIndexOf(". ");
      const breakAt = para >= 0 ? para : sent;
      if (breakAt >= 0) end = end - 1000 + breakAt + 1;
    }
    chunks.push(text.slice(start, end).trim());
    if (end >= text.length) break;
    start = end - overlap;
  }
  return chunks.filter((c) => c.length > 0);
}
