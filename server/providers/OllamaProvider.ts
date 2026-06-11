import { ModelProvider, EditRequest } from "./ModelProvider";
import { buildPrompt, sanitizeModelOutput } from "../prompts/legal";

const MODEL = process.env.OLLAMA_MODEL ?? "llama3.1:8b";
const HOST = process.env.OLLAMA_HOST ?? "http://localhost:11434";
// First call after `ollama serve` loads the model into memory and can be slow.
const TIMEOUT_MS = Number(process.env.OLLAMA_TIMEOUT_MS ?? 120000);

export class OllamaProvider implements ModelProvider {
  name() {
    return "ollama";
  }
  version() {
    return MODEL;
  }

  async edit({ text, instruction, mode = "edit", reference }: EditRequest): Promise<string> {
    const { system, user } = buildPrompt(mode, text, instruction, reference);
    return this.complete(system, user);
  }

  // Generic completion (also used for ingestion/distillation).
  async complete(system: string, user: string): Promise<string> {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
    try {
      const res = await fetch(`${HOST}/api/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal: controller.signal,
        body: JSON.stringify({
          model: MODEL,
          stream: false,
          // Lower temperature: legal edits should be conservative, not creative.
          options: { temperature: 0.2 },
          messages: [
            { role: "system", content: system },
            { role: "user", content: user },
          ],
        }),
      });
      if (!res.ok) {
        const detail = await res.text();
        throw new Error(`Ollama ${res.status}: ${detail}`);
      }
      const data = (await res.json()) as { message?: { content?: string } };
      const content = sanitizeModelOutput(data.message?.content ?? "");
      if (!content) {
        throw new Error("Ollama returned no message content");
      }
      return content;
    } finally {
      clearTimeout(timer);
    }
  }
}
