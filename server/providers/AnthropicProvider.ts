import Anthropic from "@anthropic-ai/sdk";
import { ModelProvider, EditRequest } from "./ModelProvider";
import { buildPrompt, sanitizeModelOutput } from "../prompts/legal";

// Claude provider. Key is read from ANTHROPIC_API_KEY (server-side only).
const MODEL = process.env.ANTHROPIC_MODEL ?? "claude-opus-4-8";
const MAX_TOKENS = Number(process.env.ANTHROPIC_MAX_TOKENS ?? 4096);

export class AnthropicProvider implements ModelProvider {
  private client: Anthropic;

  constructor() {
    if (!process.env.ANTHROPIC_API_KEY) {
      throw new Error(
        "ANTHROPIC_API_KEY is not set. Export it (server-side) to use MODEL_PROVIDER=anthropic."
      );
    }
    this.client = new Anthropic(); // reads ANTHROPIC_API_KEY from env
  }

  name() {
    return "anthropic";
  }
  version() {
    return MODEL;
  }

  async edit({ text, instruction, mode = "edit", reference }: EditRequest): Promise<string> {
    const { system, user } = buildPrompt(mode, text, instruction, reference);
    return this.complete(system, user);
  }

  // Generic completion (also used for ingestion/distillation).
  // Note: Opus 4.8 rejects temperature/top_p/budget_tokens. We omit them and
  // rely on the system prompt's "return only the requested text" guardrail.
  async complete(system: string, user: string): Promise<string> {
    const res = await this.client.messages.create({
      model: MODEL,
      max_tokens: MAX_TOKENS,
      system,
      messages: [{ role: "user", content: user }],
    });

    if (res.stop_reason === "refusal") {
      throw new Error("Claude refused this request for safety reasons.");
    }

    // content is a list of blocks; concatenate the text blocks.
    const out = sanitizeModelOutput(
      res.content
        .filter((b): b is Anthropic.TextBlock => b.type === "text")
        .map((b) => b.text)
        .join("")
    );

    if (!out) {
      throw new Error("Claude returned no text content");
    }
    return out;
  }
}
