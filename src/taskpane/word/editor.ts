/* global Word Office */

// All in-document interaction lives here (the Word.run logic).
// POC core loop: read the user's selection -> backend edits it -> land the edit
// as a NATIVE tracked change the lawyer can accept or reject.

const BACKEND_URL = "https://localhost:3001/api/edit";

export interface EditResult {
  original: string;
  edited: string;
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

// Send the selected text + instruction to the backend, then write the edited
// text back into the same range as a tracked change.
export async function editSelection(
  instruction: string,
  onStatus?: (msg: string) => void
): Promise<EditResult> {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    const original = range.text;
    if (!original.trim()) {
      throw new Error("Select some text in the document first.");
    }

    // Backend proxies to the model; the client never calls the model directly.
    onStatus?.("Asking the model…");
    const res = await fetch(BACKEND_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: original, instruction, docName: docName() }),
    });
    if (!res.ok) {
      const detail = await res.text();
      throw new Error(`Backend error ${res.status}: ${detail}`);
    }
    const { text: edited } = (await res.json()) as { text: string };

    // Track changes must be ON before the write, so the edit lands as a redline.
    // changeTrackingMode is WordApi 1.4 -- gate it so older builds don't break.
    onStatus?.("Applying tracked change…");
    if (Office.context.requirements.isSetSupported("WordApi", "1.4")) {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }
    // insertText replace flattens range formatting -- acceptable for the POC.
    range.insertText(edited, Word.InsertLocation.replace);
    await context.sync();

    return { original, edited };
  });
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
