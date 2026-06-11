// Model-agnostic provider boundary. The client never talks to a model directly;
// it always goes through the backend, which goes through a ModelProvider.
// Swapping Ollama -> Anthropic in prod is a single-file change behind this seam.

// How an uploaded reference document should steer the edit/draft. Built from a
// grounding artifact at request time (see attachments/artifact.ts).
export interface ReferenceContext {
  mode: "format" | "inspiration" | "exact";
  // Already-projected reference text (skeleton / tone summary / skeleton+slots),
  // ready to drop into the prompt.
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
