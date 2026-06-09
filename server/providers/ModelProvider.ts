// Model-agnostic provider boundary. The client never talks to a model directly;
// it always goes through the backend, which goes through a ModelProvider.
// Swapping Ollama -> Anthropic in prod is a single-file change behind this seam.

export interface EditRequest {
  text: string;
  instruction: string;
  // "edit" (default): revise the supplied text. "draft": author new text from
  // the instruction alone (text may be empty).
  mode?: "edit" | "draft";
}

export interface ModelProvider {
  name(): string; // provider name, for the audit log (e.g. "ollama")
  version(): string; // model tag/version, for the audit log
  edit(req: EditRequest): Promise<string>;
}
