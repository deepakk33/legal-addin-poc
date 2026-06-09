import { Router, Request, Response } from "express";
import { selectProvider } from "../providers";
import { recordEdit } from "../db/audit";

const router = Router();
const provider = selectProvider();

interface EditBody {
  text?: string;
  instruction?: string;
  docName?: string;
}

// POST /api/edit  { text, instruction, docName? } -> { text: edited }
router.post("/edit", async (req: Request, res: Response) => {
  const { text, instruction, docName } = (req.body ?? {}) as EditBody;

  if (!text || !text.trim()) {
    return res.status(400).json({ error: "Missing 'text' to edit." });
  }
  if (!instruction || !instruction.trim()) {
    return res.status(400).json({ error: "Missing 'instruction'." });
  }

  try {
    const edited = await provider.edit({ text, instruction });

    // Append-only audit row (status 'pending' until accept/reject is wired).
    recordEdit({
      docName: docName ?? null,
      instruction,
      modelName: provider.name(),
      modelVersion: provider.version(),
      originalText: text,
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
