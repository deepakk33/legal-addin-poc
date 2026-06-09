import { ModelProvider, EditRequest } from "./ModelProvider";
import { LEGAL_SYSTEM_PROMPT } from "../prompts/legal";

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

  async edit({ text, instruction }: EditRequest): Promise<string> {
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
            { role: "system", content: LEGAL_SYSTEM_PROMPT },
            {
              role: "user",
              content:
                `Instruction: ${instruction}\n\n` +
                `Text to edit:\n${text}\n\n` +
                `Return ONLY the edited text, no commentary, no markdown.`,
            },
          ],
        }),
      });
      if (!res.ok) {
        const detail = await res.text();
        throw new Error(`Ollama ${res.status}: ${detail}`);
      }
      const data = (await res.json()) as { message?: { content?: string } };
      const content = data.message?.content;
      if (!content) {
        throw new Error("Ollama returned no message content");
      }
      return content.trim();
    } finally {
      clearTimeout(timer);
    }
  }
}
