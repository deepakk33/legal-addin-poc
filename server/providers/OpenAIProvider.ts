import OpenAI from "openai";
import { ModelProvider, EditRequest } from "./ModelProvider";
import { buildPrompt, sanitizeModelOutput } from "../prompts/legal";

// OpenAI provider. Key is read from OPENAI_API_KEY (server-side only).
const MODEL = process.env.OPENAI_MODEL ?? "gpt-4o";

export class OpenAIProvider implements ModelProvider {
  private client: OpenAI;

  constructor() {
    if (!process.env.OPENAI_API_KEY) {
      throw new Error(
        "OPENAI_API_KEY is not set. Export it (server-side) to use MODEL_PROVIDER=openai."
      );
    }
    this.client = new OpenAI(); // reads OPENAI_API_KEY from env
  }

  name() {
    return "openai";
  }
  version() {
    return MODEL;
  }

  async edit({ text, instruction, mode = "edit" }: EditRequest): Promise<string> {
    const { system, user } = buildPrompt(mode, text, instruction);
    const res = await this.client.chat.completions.create({
      model: MODEL,
      temperature: 0.2, // conservative: legal edits, not creative writing
      messages: [
        { role: "system", content: system },
        { role: "user", content: user },
      ],
    });

    const out = sanitizeModelOutput(res.choices[0]?.message?.content ?? "");
    if (!out) {
      throw new Error("OpenAI returned no message content");
    }
    return out;
  }
}
