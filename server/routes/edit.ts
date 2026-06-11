import { Router, Request, Response } from "express";
import { getProvider } from "../providers";
import { ReferenceContext } from "../providers/ModelProvider";
import { recordEdit } from "../db/audit";
import * as store from "../attachments/store";
import { projectArtifact } from "../attachments/artifact";

const router = Router();
const provider = getProvider();

interface EditBody {
  text?: string;
  instruction?: string;
  docName?: string;
  mode?: "edit" | "draft";
  attachmentIds?: string[];
}

// Build a reference context from the selected, ready attachments by projecting
// each grounding artifact. Returns undefined if nothing usable is selected. How
// the model uses it is driven by the user's instruction (REFERENCE_ADDENDUM).
function buildReference(ids: string[] | undefined): ReferenceContext | undefined {
  if (!ids?.length) return undefined;
  const projections = ids
    .map((id) => store.get(id))
    .filter((a) => a?.status === "ready" && a.artifact)
    .map((a) => projectArtifact(a!.artifact!));
  if (!projections.length) return undefined;
  return { projection: projections.join("\n\n---\n\n") };
}

// POST /api/edit  { text, instruction, docName?, mode?, attachmentIds? }
//   -> { text: edited }
router.post("/edit", async (req: Request, res: Response) => {
  const { text, instruction, docName, mode, attachmentIds } = (req.body ?? {}) as EditBody;
  const editMode = mode === "draft" ? "draft" : "edit";

  // Edit mode needs source text; draft mode authors from the instruction alone.
  if (editMode === "edit" && (!text || !text.trim())) {
    return res.status(400).json({ error: "Missing 'text' to edit." });
  }
  if (!instruction || !instruction.trim()) {
    return res.status(400).json({ error: "Missing 'instruction'." });
  }

  const reference = buildReference(attachmentIds);

  try {
    const edited = await provider.edit({
      text: text ?? "",
      instruction,
      mode: editMode,
      reference,
    });

    // Append-only audit row (status 'pending' until accept/reject is wired).
    recordEdit({
      docName: docName ?? null,
      instruction,
      modelName: provider.name(),
      modelVersion: provider.version(),
      originalText: text ?? "",
      editedText: edited,
    });

    return res.json({ text: edited });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    // eslint-disable-next-line no-console
    console.error("[/api/edit] error:", message);
    return res.status(502).json({ error: message });
  }
});

export default router;
