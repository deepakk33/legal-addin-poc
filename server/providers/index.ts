import { ModelProvider } from "./ModelProvider";
import { OllamaProvider } from "./OllamaProvider";
import { AnthropicProvider } from "./AnthropicProvider";

// Provider selection via env. MODEL_PROVIDER=ollama (default) picks OllamaProvider.
export function selectProvider(): ModelProvider {
  const which = (process.env.MODEL_PROVIDER ?? "ollama").toLowerCase();
  switch (which) {
    case "anthropic":
      return new AnthropicProvider();
    case "ollama":
    default:
      return new OllamaProvider();
  }
}
