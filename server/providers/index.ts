import { ModelProvider } from "./ModelProvider";
import { OllamaProvider } from "./OllamaProvider";
import { AnthropicProvider } from "./AnthropicProvider";
import { OpenAIProvider } from "./OpenAIProvider";

export type ProviderName = "ollama" | "anthropic" | "openai";

// Provider selection via env. MODEL_PROVIDER=ollama (default) | anthropic | openai.
// The provider is constructed lazily so a missing API key only errors when that
// provider is actually selected.
export function selectProvider(): ModelProvider {
  const which = (process.env.MODEL_PROVIDER ?? "ollama").toLowerCase();
  switch (which) {
    case "anthropic":
    case "claude":
      return new AnthropicProvider();
    case "openai":
    case "gpt":
      return new OpenAIProvider();
    case "ollama":
      return new OllamaProvider();
    default:
      throw new Error(
        `Unknown MODEL_PROVIDER "${which}". Use one of: ollama | anthropic | openai.`
      );
  }
}

// Shared provider instance, reused by every route (edit + ingestion) so we
// don't reconstruct an SDK client per request.
let shared: ModelProvider | undefined;
export function getProvider(): ModelProvider {
  if (!shared) shared = selectProvider();
  return shared;
}
