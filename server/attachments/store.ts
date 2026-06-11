import { randomUUID } from "crypto";

// In-memory attachment store. POC-scoped: lives for the server process only,
// no DB, no disk. Each uploaded reference doc is distilled into a grounding
// artifact (see artifact.ts) and held here, keyed by a generated id and
// namespaced by the client's per-pane sessionId.

export type AttachmentStatus =
  | "queued"
  | "extracting"
  | "building"
  | "ready"
  | "error"
  | "cancelled";

export interface ArtifactSlot {
  label: string;
  exampleValue: string;
}

export interface Artifact {
  headings: string[];
  clauseOrder: string[];
  numberingScheme: string;
  formattingConventions: string;
  toneSummary: string;
  slots: ArtifactSlot[];
}

export interface Attachment {
  id: string;
  sessionId: string;
  name: string;
  type: "doc";
  status: AttachmentStatus;
  artifact?: Artifact;
  error?: string;
}

const store = new Map<string, Attachment>();

export function create(sessionId: string, name: string): Attachment {
  const att: Attachment = {
    id: randomUUID(),
    sessionId,
    name,
    type: "doc",
    status: "queued",
  };
  store.set(att.id, att);
  return att;
}

export function get(id: string): Attachment | undefined {
  return store.get(id);
}

// Mutators no-op if the entry was removed (e.g. user cancelled mid-pipeline),
// so the ingestion pipeline can keep running without guarding every call.
export function setStatus(id: string, status: AttachmentStatus): void {
  const a = store.get(id);
  if (a) a.status = status;
}

export function setArtifact(id: string, artifact: Artifact): void {
  const a = store.get(id);
  if (a) {
    a.artifact = artifact;
    a.status = "ready";
  }
}

export function setError(id: string, message: string): void {
  const a = store.get(id);
  if (a) {
    a.error = message;
    a.status = "error";
  }
}

export function remove(id: string): boolean {
  return store.delete(id);
}

// The pipeline calls this between steps: if the entry is gone, the user cancelled
// (DELETE removed it), so the pipeline should bail.
export function isAlive(id: string): boolean {
  return store.has(id);
}
