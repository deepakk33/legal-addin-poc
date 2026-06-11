import { ReferenceContext } from "../providers/ModelProvider";

export type EditMode = "edit" | "draft";

// Build the (system, user) pair for a given mode. Centralized so every provider
// drafts/edits identically and the anti-hallucination guardrails can't drift.
// When `reference` is present (the user uploaded a doc and picked a mode), the
// system prompt gains a mode-specific addendum and the user message carries the
// projected reference inside a clearly delimited block.
export function buildPrompt(
  mode: EditMode,
  text: string,
  instruction: string,
  reference?: ReferenceContext
) {
  const system =
    (mode === "draft" ? LEGAL_DRAFT_SYSTEM_PROMPT : LEGAL_SYSTEM_PROMPT) +
    (reference ? "\n\n" + referenceAddendum(reference.mode) : "");

  const refBlock = reference
    ? `=== REFERENCE (uploaded) ===\n${reference.projection}\n=== END REFERENCE ===\n\n`
    : "";

  if (mode === "draft") {
    return {
      system,
      user:
        refBlock +
        `Instruction: ${instruction}\n\n` +
        `Draft the text now. Return ONLY the drafted text, no commentary, no markdown.`,
    };
  }
  return {
    system,
    user:
      refBlock +
      `Instruction: ${instruction}\n\n` +
      `Text to edit:\n${text}\n\n` +
      `Return ONLY the edited text, no commentary, no markdown.`,
  };
}

// Mode-specific guidance appended to the system prompt when a reference is in
// play. Keeps the model focused on the uploaded doc the way the user intended,
// without letting it copy facts it shouldn't.
export function referenceAddendum(mode: ReferenceContext["mode"]): string {
  switch (mode) {
    case "format":
      return `A REFERENCE block is provided. Mirror its STRUCTURE: headings, clause ordering,
numbering scheme, and formatting conventions. Do NOT reuse the reference's specific
facts (parties, dates, figures); draw all content from the user's instruction.`;
    case "inspiration":
      return `A REFERENCE block is provided as a loose stylistic guide. Use its tone and
approach for inspiration only; adapt freely. Stay within the user's instruction and
do not copy the reference's specific facts.`;
    case "exact":
      return `A REFERENCE block is provided with a structure and labeled data slots. Follow the
structure closely and fill the slots with the data in the user's instruction. Use
bracketed placeholders (e.g. [PARTY A]) for any slot the user did not supply. Do not
invent facts beyond the instruction.`;
  }
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

// ---------------------------------------------------------------------------
// Ingestion / grounding-artifact prompts
// ---------------------------------------------------------------------------

// Distillation: turn an uploaded reference document into a compact, reusable
// structured artifact (JSON). Facts are stripped from the skeleton fields and
// captured only as labeled slots, so "format" mode can mirror structure without
// leaking the source doc's parties/figures.
export const DISTILL_SYSTEM_PROMPT = `You analyze a legal/contract document and extract its REUSABLE structure as strict JSON.

Output ONLY a JSON object with exactly these keys:
- "headings": string[]            // section/heading titles in order
- "clauseOrder": string[]         // the clause topics in the order they appear (short labels)
- "numberingScheme": string       // e.g. "1., 1.1, 1.1.1" or "Article I / Section 1.01" or "none"
- "formattingConventions": string // defined-term capitalization, recitals style, signature block, etc.
- "toneSummary": string           // 2-3 sentences on drafting tone and style
- "slots": [{"label": string, "exampleValue": string}]  // the variable facts (parties, dates, amounts)

Rules:
- Strip specific facts (party names, dates, figures, addresses) OUT of headings/clauseOrder/
  numberingScheme/formattingConventions; capture them ONLY inside "slots".
- Be concise. Output valid JSON only — no markdown, no commentary.`;

// For very large docs we map-reduce: condense each chunk to a structural outline
// first, then distill the merged outline.
export const CONDENSE_SYSTEM_PROMPT = `Condense this section of a legal document into a brief STRUCTURAL outline:
list its headings, clause topics (in order), numbering style, and any notable
drafting conventions. Omit specific facts. Output plain text, no commentary.`;

export function buildDistillUser(text: string): string {
  return `Document text:\n\n${text}\n\nExtract the JSON artifact now.`;
}

export function buildCondenseUser(chunkText: string): string {
  return `Section text:\n\n${chunkText}\n\nReturn the structural outline now.`;
}
