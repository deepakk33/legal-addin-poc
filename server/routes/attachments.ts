import { Router, Request, Response } from "express";
import multer from "multer";
import * as path from "path";
import { getProvider } from "../providers";
import { extractText } from "../attachments/extract";
import { buildArtifact } from "../attachments/artifact";
import * as store from "../attachments/store";

const router = Router();

const ALLOWED = new Set([".docx", ".pdf", ".txt"]);

// Memory storage: files are small reference docs, held only long enough to
// extract + distill. ~15MB cap.
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ALLOWED.has(ext)) cb(null, true);
    else cb(new Error(`Unsupported file type "${ext}". Upload a .docx, .pdf, or .txt.`));
  },
});

// POST /api/attachments  (multipart: file, sessionId?)
// Responds immediately with the created card; ingestion runs async and the
// client polls GET /api/attachments/:id for state transitions.
router.post("/attachments", upload.single("file"), (req: Request, res: Response) => {
  const file = req.file;
  if (!file) return res.status(400).json({ error: "No file uploaded (field 'file')." });

  const sessionId = (req.body?.sessionId as string) || "default";
  const att = store.create(sessionId, file.originalname);

  // Fire-and-forget ingestion pipeline. Each step checks the entry is still
  // alive (a DELETE = user cancelled) and bails if not.
  void ingest(att.id, file.buffer, file.originalname);

  return res.status(202).json({ id: att.id, name: att.name, status: att.status });
});

// GET /api/attachments/:id -> current card state
router.get("/attachments/:id", (req: Request, res: Response) => {
  const a = store.get(req.params.id);
  if (!a) return res.status(404).json({ error: "Not found" });
  return res.json({ id: a.id, name: a.name, status: a.status, error: a.error });
});

// DELETE /api/attachments/:id -> remove (backs both "remove" and "cancel")
router.delete("/attachments/:id", (req: Request, res: Response) => {
  store.remove(req.params.id);
  return res.status(204).end();
});

async function ingest(id: string, buf: Buffer, filename: string): Promise<void> {
  try {
    store.setStatus(id, "extracting");
    const { text } = await extractText(buf, filename);
    if (!store.isAlive(id)) return; // cancelled during extraction

    store.setStatus(id, "building");
    const artifact = await buildArtifact(getProvider(), text);
    if (!store.isAlive(id)) return; // cancelled during distillation

    store.setArtifact(id, artifact);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    // eslint-disable-next-line no-console
    console.error(`[/api/attachments ${id}] ingest error:`, message);
    store.setError(id, message);
  }
}

export default router;
