import { ModelProvider, ReferenceContext } from "../providers/ModelProvider";
import {
  DISTILL_SYSTEM_PROMPT,
  CONDENSE_SYSTEM_PROMPT,
  buildDistillUser,
  buildCondenseUser,
} from "../prompts/legal";
import { chunk } from "./chunk";
import { Artifact } from "./store";

// Above this many characters we map-reduce: condense each chunk to a structural
// outline, then distill the merged outline. Keeps the distill call's input
// "sized to the active model" without embeddings or a vector store.
const MAP_REDUCE_THRESHOLD = 24000;

// Build a grounding artifact from extracted reference text via one (or, for big
// docs, a few) model completion call(s).
export async function buildArtifact(provider: ModelProvider, text: string): Promise<Artifact> {
  if (!text.trim()) {
    throw new Error("No text could be extracted from this file.");
  }

  let toDistill = text;
  if (text.length > MAP_REDUCE_THRESHOLD) {
    const chunks = chunk(text);
    const outlines: string[] = [];
    for (const c of chunks) {
      outlines.push(await provider.complete(CONDENSE_SYSTEM_PROMPT, buildCondenseUser(c)));
    }
    toDistill = outlines.join("\n\n");
  }

  const raw = await provider.complete(DISTILL_SYSTEM_PROMPT, buildDistillUser(toDistill));
  return parseArtifact(raw);
}

// Robust JSON parse: the model may wrap the object in prose despite the prompt,
// so slice from the first { to the last }. Missing keys are defaulted so a
// partial artifact never crashes downstream projection.
function parseArtifact(raw: string): Artifact {
  const start = raw.indexOf("{");
  const end = raw.lastIndexOf("}");
  if (start === -1 || end === -1 || end < start) {
    throw new Error("Distillation did not return JSON.");
  }
  let obj: any;
  try {
    obj = JSON.parse(raw.slice(start, end + 1));
  } catch {
    throw new Error("Distillation returned invalid JSON.");
  }
  return {
    headings: asStringArray(obj.headings),
    clauseOrder: asStringArray(obj.clauseOrder),
    numberingScheme: typeof obj.numberingScheme === "string" ? obj.numberingScheme : "none",
    formattingConventions:
      typeof obj.formattingConventions === "string" ? obj.formattingConventions : "",
    toneSummary: typeof obj.toneSummary === "string" ? obj.toneSummary : "",
    slots: Array.isArray(obj.slots)
      ? obj.slots
          .filter((s: any) => s && typeof s.label === "string")
          .map((s: any) => ({ label: s.label, exampleValue: String(s.exampleValue ?? "") }))
      : [],
  };
}

function asStringArray(v: any): string[] {
  return Array.isArray(v) ? v.filter((x) => typeof x === "string") : [];
}

// Project an artifact into the reference text injected at edit time, per the
// chosen mode. This is the "Prompt + mode -> Assemble context" step.
export function projectArtifact(artifact: Artifact, mode: ReferenceContext["mode"]): string {
  const skeleton = () =>
    [
      artifact.headings.length ? `Headings:\n- ${artifact.headings.join("\n- ")}` : "",
      artifact.clauseOrder.length ? `Clause order:\n- ${artifact.clauseOrder.join("\n- ")}` : "",
      artifact.numberingScheme ? `Numbering scheme: ${artifact.numberingScheme}` : "",
      artifact.formattingConventions
        ? `Formatting conventions: ${artifact.formattingConventions}`
        : "",
    ]
      .filter(Boolean)
      .join("\n\n");

  switch (mode) {
    case "format":
      return skeleton();
    case "inspiration":
      return artifact.toneSummary || skeleton();
    case "exact": {
      const slots = artifact.slots.length
        ? "Data slots to fill:\n" +
          artifact.slots.map((s) => `- ${s.label} (e.g. ${s.exampleValue})`).join("\n")
        : "Data slots: none identified.";
      return `${skeleton()}\n\n${slots}`;
    }
  }
}
