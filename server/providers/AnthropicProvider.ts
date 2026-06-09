import { ModelProvider, EditRequest } from "./ModelProvider";

// Prod swap target. Deliberately a stub for the POC: it makes the seam real
// (the interface and selection path exist) without building the integration.
// In prod this would hold the server-side API key and call the Anthropic API.
export class AnthropicProvider implements ModelProvider {
  name() {
    return "anthropic";
  }
  version() {
    return "stub";
  }

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  async edit(_req: EditRequest): Promise<string> {
    throw new Error(
      "AnthropicProvider is not configured for the POC. Set MODEL_PROVIDER=ollama, " +
        "or implement this provider (single-file change) to swap to the Anthropic API."
    );
  }
}
