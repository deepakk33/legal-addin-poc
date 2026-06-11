/* global Word Office fetch Response FormData File AbortSignal */

// All in-document interaction lives here (the Word.run logic).
// POC core loop: read the user's selection -> backend edits it -> land the edit
// as a NATIVE tracked change the lawyer can accept or reject. When nothing is
// selected we fall back to DRAFT mode: author new text and insert it (still as
// a tracked change), so an empty document is usable too.

const BACKEND_URL = "https://localhost:3001/api/edit";
const ATTACH_URL = "https://localhost:3001/api/attachments";

export type EditMode = "edit" | "draft";

// Reference grounding from uploaded attachments, threaded into every edit/draft.
// How the model uses the docs is driven by the user's instruction (the system
// prompt interprets it), so there's no mode here — just which docs to ground on.
export interface ReferenceInput {
  attachmentIds: string[];
}

export interface EditResult {
  mode: EditMode;
  original: string; // "" in draft mode
  edited: string;
  changed: boolean; // false when the model returned the text unchanged
}

export interface DocState {
  selectionText: string; // trimmed selection
  hasSelection: boolean;
  bodyEmpty: boolean;
  // Human label for the edit target: a paragraph range ("¶ 13–15") when text is
  // selected, otherwise the document name (the draft target). Word exposes no
  // line numbers — paragraphs are the closest faithful locator.
  location: string;
  documentName: string;
}

// ---------------------------------------------------------------------------
// Browser preview: when not running inside a Word host, mock the document layer
// so the whole UI flow is exercisable in a plain browser. The real Word host
// short-circuits all of this. The backend is mocked too, unless ?realBackend is
// in the URL (then we hit the live localhost:3001 server).
// ---------------------------------------------------------------------------
function inWord(): boolean {
  try {
    return (
      typeof Word !== "undefined" &&
      typeof Office !== "undefined" &&
      !!Office.context &&
      Office.context.host === Office.HostType.Word
    );
  } catch {
    return false;
  }
}

function useRealBackend(): boolean {
  try {
    return /[?&]realBackend\b/.test(window.location.search);
  } catch {
    return false;
  }
}

// Mutable mock selection for browser preview. Seeded with a sample clause so
// Edit mode has something to work on; cleared via setMockSelection("") to
// exercise Draft mode. Writes update it so a re-run sees the prior edit.
let mockSelection =
  "The Company shall indemnify and hold harmless the Client from any and all " +
  "claims arising out of the Company's gross negligence.";

export function setMockSelection(text: string): void {
  mockSelection = text;
}

// Deterministic stand-in for the backend in browser preview.
function mockEdit(body: { text: string; instruction: string; mode: EditMode }): string {
  if (body.mode === "draft") {
    return (
      `[MOCK DRAFT — instruction: "${body.instruction}"]\n\n` +
      "This Agreement is entered into as of the Effective Date by and between the " +
      "parties. The parties agree to the terms set forth herein, including the " +
      "obligations, representations, and warranties described below."
    );
  }
  // Edit: visibly transform so a redline is produced.
  return `${body.text} [MOCK EDIT — applied: "${body.instruction}"]`;
}

// Inspect the document: what's selected and whether the body is empty.
// Drives the UI (Edit vs Draft) without mutating anything.
export async function getDocState(): Promise<DocState> {
  if (!inWord()) {
    const sel = mockSelection.trim();
    const has = sel.length > 0;
    return {
      selectionText: sel,
      hasSelection: has,
      bodyEmpty: sel.length === 0,
      location: has ? `Selection · ${wordCount(sel)} words` : "Untitled.docx",
      documentName: docName(),
    };
  }
  return Word.run(async (context) => {
    const sel = context.document.getSelection();
    const body = context.document.body;
    sel.load("text");
    body.load("text");
    await context.sync();

    const selectionText = (sel.text || "").trim();
    const hasSelection = selectionText.length > 0;
    const name = docName();

    let location = name;
    if (hasSelection) {
      location = `Selection · ${wordCount(selectionText)} words`;
      // Best-effort paragraph range: count paragraph breaks before the selection
      // start (WordApi 1.3 getRange/expandTo). Falls back to the word count above.
      try {
        const before = body.getRange("Start").expandTo(sel.getRange("Start"));
        before.load("text");
        await context.sync();
        const start = paragraphBreaks(before.text) + 1;
        const span = paragraphBreaks(sel.text);
        location = span > 0 ? `¶ ${start}–${start + span}` : `¶ ${start}`;
      } catch {
        /* keep the word-count label */
      }
    }

    return {
      selectionText,
      hasSelection,
      bodyEmpty: (body.text || "").trim().length === 0,
      location,
      documentName: name,
    };
  });
}

// Word separates paragraphs with \r in Range.text; count them to derive a
// paragraph index/span.
function paragraphBreaks(text: string): number {
  return (text.match(/[\r\n]/g) || []).length;
}

function wordCount(text: string): number {
  return text.split(/\s+/).filter(Boolean).length;
}

// Subscribe to Word selection changes so the pane auto-reflects the current
// target without a manual "read" button. Returns an unsubscribe fn. No-op in
// browser preview (the mock selection is static).
export function subscribeSelection(onChange: () => void): () => void {
  if (!inWord()) return () => undefined;
  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onChange
    );
  } catch {
    return () => undefined;
  }
  return () => {
    try {
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
        handler: onChange,
      });
    } catch {
      /* best-effort */
    }
  };
}

// Read the text of the current selection (no document mutation).
export async function readSelection(): Promise<string> {
  if (!inWord()) return mockSelection;
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    return range.text;
  });
}

// Call the backend. Wraps fetch so a dead/uncertified localhost backend yields a
// human message instead of an opaque "Failed to fetch".
async function callBackend(body: {
  text: string;
  instruction: string;
  mode: EditMode;
  docName: string;
  attachmentIds?: string[];
}): Promise<string> {
  if (!inWord() && !useRealBackend()) {
    return mockEdit(body);
  }
  let res: Response;
  try {
    res = await fetch(BACKEND_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
  } catch {
    throw new Error(
      `Can't reach the backend at ${BACKEND_URL}. Is the server running, and have you trusted its HTTPS cert? ` +
        `(Open ${BACKEND_URL.replace("/api/edit", "")} once in a browser to accept it.)`
    );
  }
  if (!res.ok) {
    let detail = await res.text();
    try {
      detail = (JSON.parse(detail) as { error?: string }).error || detail;
    } catch {
      /* keep raw text */
    }
    throw new Error(`Backend error ${res.status}: ${detail}`);
  }
  const data = (await res.json()) as { text?: string };
  if (!data.text || !data.text.trim()) {
    throw new Error("Backend returned no edited text.");
  }
  return data.text;
}

// Turn track changes on so the write lands as a redline. WordApi 1.4 gate so
// older builds don't throw.
function enableTrackChanges(context: Word.RequestContext): void {
  if (Office.context.requirements.isSetSupported("WordApi", "1.4")) {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  }
}

// Whitespace-preserving replacement: the model returns the trimmed core, but the
// selection may carry leading/trailing spaces or newlines we don't want to eat.
function reclothe(original: string, editedCore: string): string {
  const lead = original.match(/^\s*/)?.[0] ?? "";
  const trail = original.match(/\s*$/)?.[0] ?? "";
  return lead + editedCore + trail;
}

// EDIT mode: revise the current selection in place as a tracked change.
async function runEdit(
  instruction: string,
  onStatus?: (m: string) => void,
  reference?: ReferenceInput
): Promise<EditResult> {
  if (!inWord()) {
    const original = mockSelection;
    if (!original.trim()) {
      throw new Error("Nothing selected. Highlight a clause to edit, or clear the doc to draft.");
    }
    onStatus?.("Asking the model…");
    const editedCore = await callBackend({
      text: original,
      instruction,
      mode: "edit",
      docName: docName(),
      attachmentIds: reference?.attachmentIds,
    });
    if (normalize(editedCore) === normalize(original)) {
      return { mode: "edit", original, edited: original, changed: false };
    }
    onStatus?.("Applying tracked change…");
    mockSelection = reclothe(original, editedCore); // reflect the "write"
    return { mode: "edit", original, edited: editedCore, changed: true };
  }
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    const original = range.text;
    if (!original.trim()) {
      throw new Error("Nothing selected. Highlight a clause to edit, or clear the doc to draft.");
    }

    onStatus?.("Asking the model…");
    const editedCore = await callBackend({
      text: original,
      instruction,
      mode: "edit",
      docName: docName(),
      attachmentIds: reference?.attachmentIds,
    });

    // No-op guard: if the model returned the same text (e.g. the instruction
    // can't be applied), don't write an empty redline -- report it instead.
    if (normalize(editedCore) === normalize(original)) {
      return { mode: "edit", original, edited: original, changed: false };
    }

    onStatus?.("Applying tracked change…");
    enableTrackChanges(context);
    // insertText replace flattens range formatting -- acceptable for the POC.
    range.insertText(reclothe(original, editedCore), Word.InsertLocation.replace);
    await context.sync();

    return { mode: "edit", original, edited: editedCore, changed: true };
  });
}

// DRAFT mode: author new text from the instruction and insert it as a tracked
// change. Used when nothing is selected (typically an empty document).
async function runDraft(
  instruction: string,
  onStatus?: (m: string) => void,
  reference?: ReferenceInput
): Promise<EditResult> {
  if (!inWord()) {
    onStatus?.("Drafting…");
    const draft = await callBackend({
      text: "",
      instruction,
      mode: "draft",
      docName: docName(),
      attachmentIds: reference?.attachmentIds,
    });
    onStatus?.("Inserting as tracked change…");
    mockSelection = draft; // reflect the inserted text
    return { mode: "draft", original: "", edited: draft, changed: true };
  }
  return Word.run(async (context) => {
    onStatus?.("Drafting…");
    const draft = await callBackend({
      text: "",
      instruction,
      mode: "draft",
      docName: docName(),
      attachmentIds: reference?.attachmentIds,
    });

    onStatus?.("Inserting as tracked change…");
    enableTrackChanges(context);

    // Insert at the cursor if there's an insertion point; otherwise append to the
    // end of the body. getSelection() returns the caret even when "empty".
    const range = context.document.getSelection();
    range.insertText(draft, Word.InsertLocation.replace);
    await context.sync();

    return { mode: "draft", original: "", edited: draft, changed: true };
  });
}

// Single entry point the UI calls. Picks edit vs draft from the live selection
// so the user can't fire the wrong mode.
export async function runInstruction(
  instruction: string,
  onStatus?: (m: string) => void,
  reference?: ReferenceInput
): Promise<EditResult> {
  if (!instruction.trim()) {
    throw new Error("Type an instruction first.");
  }
  onStatus?.("Reading document…");
  const state = await getDocState();
  return state.hasSelection
    ? runEdit(instruction, onStatus, reference)
    : runDraft(instruction, onStatus, reference);
}

// Collapse whitespace for no-op comparison so trivial spacing differences from
// the model don't count as a real change.
function normalize(s: string): string {
  return s.replace(/\s+/g, " ").trim();
}

// Best-effort document name for the audit trail.
function docName(): string {
  try {
    const url = (Office.context.document as unknown as { url?: string }).url;
    if (!url) return "untitled";
    return url.split(/[\\/]/).pop() || "untitled";
  } catch {
    return "untitled";
  }
}

// ---------------------------------------------------------------------------
// Attachments: upload a reference doc, poll its ingestion state, delete it.
// ---------------------------------------------------------------------------

export type AttachmentStatus =
  | "queued"
  | "extracting"
  | "building"
  | "ready"
  | "error"
  | "cancelled";

export interface AttachmentState {
  id: string;
  name: string;
  status: AttachmentStatus;
  error?: string;
}

// POST the file as multipart. Returns the created card (status "queued").
// `signal` lets the caller abort the upload itself (cancel before it lands).
export async function uploadAttachment(
  file: File,
  sessionId: string,
  signal?: AbortSignal
): Promise<AttachmentState> {
  if (!inWord() && !useRealBackend()) {
    // Browser preview: fake a card. getAttachment() then ramps it to "ready".
    const id = `mock-${file.name}-${file.size}`;
    mockAttachments.set(id, { id, name: file.name, status: "queued" });
    return mockAttachments.get(id)!;
  }

  const form = new FormData();
  form.append("file", file);
  form.append("sessionId", sessionId);

  let res: Response;
  try {
    res = await fetch(ATTACH_URL, { method: "POST", body: form, signal });
  } catch (e) {
    if ((e as Error).name === "AbortError") throw e;
    throw new Error(`Can't reach the backend at ${ATTACH_URL}. Is the server running?`);
  }
  if (!res.ok) {
    let detail = await res.text();
    try {
      detail = (JSON.parse(detail) as { error?: string }).error || detail;
    } catch {
      /* keep raw text */
    }
    throw new Error(detail || `Upload failed (${res.status}).`);
  }
  return (await res.json()) as AttachmentState;
}

// Fetch the current ingestion state of one attachment.
// Browser-preview attachment store + a tiny state machine so the polling UI
// shows the queued -> extracting -> building -> ready progression.
const mockAttachments = new Map<string, AttachmentState>();
const MOCK_RAMP: Record<string, AttachmentStatus> = {
  queued: "extracting",
  extracting: "building",
  building: "ready",
};

export async function getAttachment(id: string, signal?: AbortSignal): Promise<AttachmentState> {
  if (!inWord() && !useRealBackend()) {
    const cur = mockAttachments.get(id);
    if (!cur) throw new Error(`Attachment ${id} not found.`);
    const next = MOCK_RAMP[cur.status];
    if (next) mockAttachments.set(id, { ...cur, status: next });
    return mockAttachments.get(id)!;
  }
  const res = await fetch(`${ATTACH_URL}/${id}`, { signal });
  if (!res.ok) throw new Error(`Attachment ${id} not found.`);
  return (await res.json()) as AttachmentState;
}

// Remove an attachment server-side (also serves as "cancel" for an in-flight
// ingestion: the pipeline bails when the entry disappears).
export async function deleteAttachment(id: string): Promise<void> {
  if (!inWord() && !useRealBackend()) {
    mockAttachments.delete(id);
    return;
  }
  try {
    await fetch(`${ATTACH_URL}/${id}`, { method: "DELETE" });
  } catch {
    /* best-effort; the card is removed from the UI regardless */
  }
}
