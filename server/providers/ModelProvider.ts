// Model-agnostic provider boundary. The client never talks to a model directly;
// it always goes through the backend, which goes through a ModelProvider.
// Swapping Ollama -> Anthropic in prod is a single-file change behind this seam.

// Grounding from an uploaded reference document, built at request time from its
// distilled artifact (see attachments/artifact.ts). How it's used (mirror the
// format, take inspiration, reframe with the user's data, ...) is driven by the
// user's own instruction, not a fixed mode — the system prompt interprets it.
export interface ReferenceContext {
  // Projected reference text (structure + tone + data slots), ready to drop
  // into the prompt.
  projection: string;
}

export interface EditRequest {
  text: string;
  instruction: string;
  // "edit" (default): revise the supplied text. "draft": author new text from
  // the instruction alone (text may be empty).
  mode?: "edit" | "draft";
  // Optional grounding from uploaded reference documents.
  reference?: ReferenceContext;
}

export interface ModelProvider {
  name(): string; // provider name, for the audit log (e.g. "ollama")
  version(): string; // model tag/version, for the audit log
  edit(req: EditRequest): Promise<string>;
  // Generic single-shot completion. Used for ingestion (distilling a reference
  // doc into a grounding artifact), separate from the redline edit() loop.
  complete(system: string, user: string): Promise<string>;
}
