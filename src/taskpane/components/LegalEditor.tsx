import * as React from "react";
import { useState, useRef } from "react";
import {
  Button,
  Field,
  Textarea,
  Text,
  Spinner,
  Card,
  RadioGroup,
  Radio,
  tokens,
  makeStyles,
  mergeClasses,
} from "@fluentui/react-components";
import { DismissRegular, DocumentRegular, CheckmarkCircleRegular, ErrorCircleRegular } from "@fluentui/react-icons";
import {
  runInstruction,
  getDocState,
  readSelection,
  uploadAttachment,
  getAttachment,
  deleteAttachment,
  EditResult,
  ReferenceMode,
  AttachmentStatus,
} from "../word/editor";

/* global HTMLTextAreaElement HTMLInputElement setTimeout clearTimeout crypto */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "16px",
  },
  field: { width: "100%" },
  row: { display: "flex", gap: "8px", alignItems: "center" },
  status: { color: tokens.colorNeutralForeground3 },
  error: { color: tokens.colorPaletteRedForeground1 },
  preview: { display: "flex", flexDirection: "column", gap: "8px" },
  card: {
    padding: "10px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
  label: { fontWeight: tokens.fontWeightSemibold, marginBottom: "4px" },
  mono: { whiteSpace: "pre-wrap", fontSize: tokens.fontSizeBase200 },
  attachList: { display: "flex", flexDirection: "column", gap: "6px" },
  attachRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "6px 8px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  attachRowReady: { backgroundColor: "#d9f7e1" }, // positive-pastel tint when ingested
  attachName: { flexGrow: 1, fontSize: tokens.fontSizeBase200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  attachState: { fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3 },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  bad: { color: tokens.colorPaletteRedForeground1 },
});

// One attachment card in the UI. `localId` keys React rows before the server
// id arrives; `id` is the server id once the upload lands.
interface Card {
  localId: string;
  id?: string;
  name: string;
  status: AttachmentStatus;
  error?: string;
  abort: AbortController;
}

// Human label per ingestion state (the "loader per step" the user sees).
const STATE_LABEL: Record<AttachmentStatus, string> = {
  queued: "Queued…",
  extracting: "Extracting text…",
  building: "Building grounding artifact…",
  ready: "Ready",
  error: "Error",
  cancelled: "Cancelled",
};

const TERMINAL: AttachmentStatus[] = ["ready", "error", "cancelled"];

const LegalEditor: React.FC = () => {
  const styles = useStyles();
  const [instruction, setInstruction] = useState<string>("Make this clause mutual.");
  const [selection, setSelection] = useState<string>("");
  const [hasSelection, setHasSelection] = useState<boolean>(false);
  const [result, setResult] = useState<EditResult | null>(null);
  const [status, setStatus] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);
  const [attachments, setAttachments] = useState<Card[]>([]);
  const [referenceMode, setReferenceMode] = useState<ReferenceMode>("format");

  const fileInputRef = useRef<HTMLInputElement | null>(null);
  // One sessionId per pane load — namespaces this session's attachments.
  const sessionId = useRef<string>(crypto.randomUUID());

  const patchCard = (localId: string, patch: Partial<Card>) =>
    setAttachments((prev) => prev.map((c) => (c.localId === localId ? { ...c, ...patch } : c)));

  const onPickFiles = () => fileInputRef.current?.click();

  const onFilesSelected = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files ?? []);
    e.target.value = ""; // allow re-selecting the same file later
    for (const file of files) {
      void ingestFile(file);
    }
  };

  // Upload one file, then poll its server-side ingestion state until terminal,
  // advancing this card through the state enum.
  const ingestFile = async (file: File) => {
    const localId = crypto.randomUUID();
    const abort = new AbortController();
    setAttachments((prev) => [...prev, { localId, name: file.name, status: "queued", abort }]);

    let id: string;
    try {
      const created = await uploadAttachment(file, sessionId.current, abort.signal);
      id = created.id;
      patchCard(localId, { id, status: created.status });
    } catch (err) {
      if ((err as Error).name === "AbortError") {
        patchCard(localId, { status: "cancelled" });
        return;
      }
      patchCard(localId, { status: "error", error: String((err as Error).message || err) });
      return;
    }

    // Poll until terminal (or aborted via the X button).
    while (true) {
      if (abort.signal.aborted) return;
      await sleep(700, abort.signal);
      if (abort.signal.aborted) return;
      try {
        const state = await getAttachment(id, abort.signal);
        patchCard(localId, { status: state.status, error: state.error });
        if (TERMINAL.includes(state.status)) return;
      } catch {
        return; // attachment gone (cancelled) or backend hiccup
      }
    }
  };

  // X on a card: abort any in-flight work, delete server-side, drop the row.
  const removeCard = (card: Card) => {
    card.abort.abort();
    if (card.id) void deleteAttachment(card.id);
    setAttachments((prev) => prev.filter((c) => c.localId !== card.localId));
  };

  const readyIds = attachments.filter((c) => c.status === "ready" && c.id).map((c) => c.id!);

  const onReadSelection = async () => {
    setError("");
    try {
      const text = await readSelection();
      const trimmed = text.trim();
      setHasSelection(trimmed.length > 0);
      setSelection(trimmed || "(nothing selected — will draft new text)");
    } catch (e) {
      setError(String((e as Error).message || e));
    }
  };

  const onEdit = async () => {
    setError("");
    setResult(null);
    setBusy(true);
    setStatus("Working…");
    try {
      const reference =
        readyIds.length > 0 ? { attachmentIds: readyIds, referenceMode } : undefined;
      const r = await runInstruction(instruction, setStatus, reference);
      setResult(r);
      setHasSelection(r.mode === "edit");
      setSelection(r.mode === "edit" ? r.original : "(drafted new text)");
      if (!r.changed) {
        setStatus("No change made — the model returned the text unchanged for this instruction.");
      } else if (r.mode === "draft") {
        setStatus("Drafted — review the tracked insertion in the document.");
      } else {
        setStatus("Done — review the tracked change in the document.");
      }
    } catch (e) {
      setError(String((e as Error).message || e));
      setStatus("");
    } finally {
      setBusy(false);
    }
  };

  // Decide the action label from the last-known selection state. Refreshed by
  // "Read selection"; defaults to Draft until the user reads a selection.
  React.useEffect(() => {
    getDocState()
      .then((s) => setHasSelection(s.hasSelection))
      .catch(() => undefined);
  }, []);

  return (
    <div className={styles.root}>
      <Text size={400} weight="semibold">
        Silks AI
      </Text>
      <Text className={styles.status}>
        Select a clause and type an instruction to redline it, or select nothing and type an
        instruction to draft new text. Attach a reference document to ground the result in its
        format. Either way the change lands as a tracked change you can accept or reject.
      </Text>

      {/* Attachments */}
      <Field className={styles.field} label="Reference documents">
        <div className={styles.attachList}>
          {attachments.map((c) => (
            <div
              key={c.localId}
              className={mergeClasses(styles.attachRow, c.status === "ready" && styles.attachRowReady)}
            >
              <DocumentRegular />
              <span className={styles.attachName} title={c.name}>
                {c.name}
              </span>
              {!TERMINAL.includes(c.status) && <Spinner size="extra-tiny" />}
              {c.status === "ready" && <CheckmarkCircleRegular className={styles.ok} />}
              {c.status === "error" && <ErrorCircleRegular className={styles.bad} />}
              <span className={styles.attachState} title={c.error}>
                {c.status === "error" && c.error ? c.error : STATE_LABEL[c.status]}
              </span>
              <Button
                appearance="subtle"
                size="small"
                icon={<DismissRegular />}
                aria-label="Remove attachment"
                onClick={() => removeCard(c)}
              />
            </div>
          ))}
          <div>
            <Button appearance="secondary" size="small" onClick={onPickFiles} disabled={busy}>
              Add file
            </Button>
            <input
              ref={fileInputRef}
              type="file"
              multiple
              accept=".pdf,.docx,.txt"
              style={{ display: "none" }}
              onChange={onFilesSelected}
            />
          </div>
        </div>
      </Field>

      {readyIds.length > 0 && (
        <Field label="Use the reference for">
          <RadioGroup
            layout="horizontal"
            value={referenceMode}
            onChange={(_e, data) => setReferenceMode(data.value as ReferenceMode)}
          >
            <Radio value="format" label="Format" />
            <Radio value="inspiration" label="Inspiration" />
            <Radio value="exact" label="Exact reframe" />
          </RadioGroup>
        </Field>
      )}

      <div className={styles.row}>
        <Button appearance="secondary" onClick={onReadSelection} disabled={busy}>
          Read selection
        </Button>
      </div>

      {selection && (
        <Card className={styles.card}>
          <div className={styles.label}>Selected text</div>
          <Text className={styles.mono}>{selection}</Text>
        </Card>
      )}

      <Field className={styles.field} label="Instruction">
        <Textarea
          value={instruction}
          onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => setInstruction(e.target.value)}
          rows={3}
        />
      </Field>

      <div className={styles.row}>
        <Button appearance="primary" onClick={onEdit} disabled={busy || !instruction.trim()}>
          {hasSelection ? "Edit selection" : "Draft into document"}
        </Button>
        {busy && <Spinner size="tiny" />}
      </div>

      {status && <Text className={styles.status}>{status}</Text>}
      {error && <Text className={styles.error}>{error}</Text>}

      {result && result.changed && (
        <div className={styles.preview}>
          <Card className={styles.card}>
            <div className={styles.label}>
              {result.mode === "draft"
                ? "Drafted (inserted as tracked change)"
                : "AI edit (written to document)"}
            </div>
            <Text className={styles.mono}>{result.edited}</Text>
          </Card>
        </div>
      )}
    </div>
  );
};

// Promise sleep that resolves early if the signal aborts (so polling stops
// promptly when the user removes a card).
function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve) => {
    const t = setTimeout(resolve, ms);
    signal?.addEventListener("abort", () => {
      clearTimeout(t);
      resolve();
    });
  });
}

export default LegalEditor;
