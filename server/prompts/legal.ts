export type EditMode = "edit" | "draft";

// Build the (system, user) pair for a given mode. Centralized so every provider
// drafts/edits identically and the anti-hallucination guardrails can't drift.
export function buildPrompt(mode: EditMode, text: string, instruction: string) {
  if (mode === "draft") {
    return {
      system: LEGAL_DRAFT_SYSTEM_PROMPT,
      user:
        `Instruction: ${instruction}\n\n` +
        `Draft the text now. Return ONLY the drafted text, no commentary, no markdown.`,
    };
  }
  return {
    system: LEGAL_SYSTEM_PROMPT,
    user:
      `Instruction: ${instruction}\n\n` +
      `Text to edit:\n${text}\n\n` +
      `Return ONLY the edited text, no commentary, no markdown.`,
  };
}

// Models sometimes wrap output in ```fences``` or "quotes" despite the prompt.
// Strip a single enclosing layer so the redline lands as clean text.
export function sanitizeModelOutput(raw: string): string {
  let out = raw.trim();
  // Strip a leading/trailing triple-backtick fence (with optional language tag).
  const fence = out.match(/^```[a-zA-Z]*\n([\s\S]*?)\n?```$/);
  if (fence) out = fence[1].trim();
  // Strip one layer of matching wrapping quotes around the whole output.
  if (out.length >= 2) {
    const first = out[0];
    const last = out[out.length - 1];
    if ((first === '"' && last === '"') || (first === "'" && last === "'")) {
      out = out.slice(1, -1).trim();
    }
  }
  return out;
}

// Legal-tuned system prompt with anti-hallucination guardrails (spec §11).
// Refine this as redline quality is evaluated; keep it short and redline-focused.
export const LEGAL_SYSTEM_PROMPT = `You are a legal drafting assistant. You edit contract and legal text precisely
according to the user's instruction.

Rules:
- Return only the edited text. No preamble, no explanation, no markdown, no quotation marks around the whole output.
- Preserve defined terms, the capitalization of defined terms, and clause/section numbering exactly
  unless the instruction explicitly asks you to change them.
- Never invent clauses, parties, dates, figures, or legal citations. If the instruction asks for
  something the text does not support, make the minimal honest edit and do not fabricate.
- Prefer clear, plain drafting. Do not add boilerplate the instruction did not ask for.
- If the instruction cannot be safely applied to the given text, return the original text unchanged.`;

// Drafting prompt: used when there is no source text to edit (empty doc / no
// selection) and the user asks to author new text. Still anti-hallucination:
// never invent specifics the instruction did not provide -- use neutral
// placeholders like [PARTY A], [DATE], [AMOUNT] instead.
export const LEGAL_DRAFT_SYSTEM_PROMPT = `You are a legal drafting assistant. You draft new contract and legal text
from the user's instruction.

Rules:
- Return only the drafted text. No preamble, no explanation, no markdown fences, no quotation marks around the whole output.
- Never invent party names, dates, figures, or citations the instruction did not give. Use bracketed
  placeholders such as [PARTY A], [EFFECTIVE DATE], [AMOUNT] so a lawyer can fill them in.
- Use clear, conventional legal drafting and standard clause structure. Do not pad with boilerplate
  the instruction did not ask for.`;
