/* global Word Office fetch Response */

// All in-document interaction lives here (the Word.run logic).
// POC core loop: read the user's selection -> backend edits it -> land the edit
// as a NATIVE tracked change the lawyer can accept or reject. When nothing is
// selected we fall back to DRAFT mode: author new text and insert it (still as
// a tracked change), so an empty document is usable too.

const BACKEND_URL = "https://localhost:3001/api/edit";

export type EditMode = "edit" | "draft";

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
}

// Inspect the document: what's selected and whether the body is empty.
// Drives the UI (Edit vs Draft) without mutating anything.
export async function getDocState(): Promise<DocState> {
  return Word.run(async (context) => {
    const sel = context.document.getSelection();
    const body = context.document.body;
    sel.load("text");
    body.load("text");
    await context.sync();
    const selectionText = (sel.text || "").trim();
    return {
      selectionText,
      hasSelection: selectionText.length > 0,
      bodyEmpty: (body.text || "").trim().length === 0,
    };
  });
}

// Read the text of the current selection (no document mutation).
export async function readSelection(): Promise<string> {
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
}): Promise<string> {
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
async function runEdit(instruction: string, onStatus?: (m: string) => void): Promise<EditResult> {
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
async function runDraft(instruction: string, onStatus?: (m: string) => void): Promise<EditResult> {
  return Word.run(async (context) => {
    onStatus?.("Drafting…");
    const draft = await callBackend({
      text: "",
      instruction,
      mode: "draft",
      docName: docName(),
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
  onStatus?: (m: string) => void
): Promise<EditResult> {
  if (!instruction.trim()) {
    throw new Error("Type an instruction first.");
  }
  onStatus?.("Reading document…");
  const state = await getDocState();
  return state.hasSelection ? runEdit(instruction, onStatus) : runDraft(instruction, onStatus);
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
