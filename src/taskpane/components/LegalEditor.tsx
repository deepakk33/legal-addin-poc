import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { Button, Text, Spinner, Tooltip, tokens, makeStyles, mergeClasses } from "@fluentui/react-components";
import {
  AttachRegular,
  SendRegular,
  DismissRegular,
  DocumentRegular,
  DocumentTextRegular,
  CheckmarkCircleRegular,
  ErrorCircleRegular,
} from "@fluentui/react-icons";
import { brand } from "../theme";
import {
  runInstruction,
  getDocState,
  subscribeSelection,
  uploadAttachment,
  getAttachment,
  deleteAttachment,
  EditResult,
  AttachmentStatus,
} from "../word/editor";

/* global HTMLTextAreaElement HTMLInputElement HTMLDivElement setTimeout clearTimeout crypto */

const useStyles = makeStyles({
  root: {
    flexGrow: 1,
    minHeight: 0,
    display: "flex",
    flexDirection: "column",
    backgroundColor: brand.bgGrey,
  },

  // Transcript
  transcript: {
    flexGrow: 1,
    minHeight: 0,
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "16px",
  },
  empty: {
    margin: "auto",
    maxWidth: "260px",
    textAlign: "center",
    color: tokens.colorNeutralForeground3,
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
  msgRow: { display: "flex" },
  userRow: { justifyContent: "flex-end" },
  assistantRow: { justifyContent: "flex-start" },
  bubble: {
    maxWidth: "88%",
    padding: "8px 11px",
    borderRadius: tokens.borderRadiusLarge,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  userBubble: { backgroundColor: brand.purple, color: tokens.colorNeutralForegroundOnBrand },
  assistantBubble: { backgroundColor: tokens.colorNeutralBackground1, border: `1px solid ${brand.border}` },
  errorBubble: { backgroundColor: "#fdecea", color: "#dc3545" },
  caption: { fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, marginTop: "4px" },
  typingRow: { display: "flex", alignItems: "center", gap: "8px", color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },

  // Input dock
  dock: {
    flexShrink: 0,
    borderTop: `1px solid ${brand.border}`,
    backgroundColor: tokens.colorNeutralBackground1,
    padding: "10px 12px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  contextPill: { display: "flex", alignItems: "center", gap: "6px", fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3 },
  contextText: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flexShrink: 1 },

  // The whole box is the input. Clicking anywhere in it focuses the textarea;
  // the only focus affordance is the border color (no inner border, no jitter).
  inputBox: {
    border: `1px solid ${brand.border}`,
    borderRadius: tokens.borderRadiusLarge,
    padding: "8px",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: "text",
    "&:focus-within": { border: `1px solid ${brand.purple}` },
  },
  chipRow: { display: "flex", flexWrap: "wrap", gap: "6px" },
  chip: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    maxWidth: "170px",
    padding: "2px 4px 2px 8px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: brand.middle,
    fontSize: tokens.fontSizeBase100,
  },
  chipReady: { backgroundColor: brand.positivePastel },
  chipError: { backgroundColor: "#fdecea" },
  chipName: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  chipBtn: { minWidth: "20px", height: "20px", padding: 0 },

  // Plain textarea — no border, no resize grabber, inherits the brand font.
  textarea: {
    width: "100%",
    border: "none",
    outlineStyle: "none",
    resize: "none",
    backgroundColor: "transparent",
    padding: "2px 4px",
    margin: 0,
    minHeight: "40px",
    maxHeight: "140px",
    overflowY: "auto",
    fontFamily: "inherit",
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
    color: tokens.colorNeutralForeground1,
    "&::placeholder": { color: tokens.colorNeutralForeground4 },
  },

  toolbar: { display: "flex", alignItems: "center", justifyContent: "space-between" },
  sendBtn: { borderRadius: tokens.borderRadiusCircular, minWidth: "32px", height: "32px", padding: 0 },
});

interface Chip {
  localId: string;
  id?: string;
  name: string;
  status: AttachmentStatus;
  error?: string;
  abort: AbortController;
}

interface Message {
  id: string;
  role: "user" | "assistant";
  text: string;
  kind?: "result" | "error" | "info";
  caption?: string;
}

// The current edit target, read from the document and kept in one object.
interface Target {
  hasSelection: boolean;
  preview: string;
  location: string;
  documentName: string;
}

const STATE_LABEL: Record<AttachmentStatus, string> = {
  queued: "Queued…",
  extracting: "Extracting…",
  building: "Reading…",
  ready: "Ready",
  error: "Error",
  cancelled: "Cancelled",
};

const TERMINAL: AttachmentStatus[] = ["ready", "error", "cancelled"];
const EMPTY_TARGET: Target = { hasSelection: false, preview: "", location: "", documentName: "" };

const LegalEditor: React.FC = () => {
  const styles = useStyles();
  // Minimal state: input text, transcript, attachments, the doc target, and a
  // single `pending` string ("" = idle, otherwise the live status / busy flag).
  const [instruction, setInstruction] = useState<string>("");
  const [messages, setMessages] = useState<Message[]>([]);
  const [attachments, setAttachments] = useState<Chip[]>([]);
  const [target, setTarget] = useState<Target>(EMPTY_TARGET);
  const [pending, setPending] = useState<string>("");

  const busy = pending !== "";
  const canSend = !busy && instruction.trim().length > 0;
  const readyIds = attachments.filter((c) => c.status === "ready" && c.id).map((c) => c.id!);

  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const transcriptEndRef = useRef<HTMLDivElement | null>(null);
  const sessionId = useRef<string>(crypto.randomUUID());

  const pushMessage = (m: Omit<Message, "id">) =>
    setMessages((prev) => [...prev, { ...m, id: crypto.randomUUID() }]);
  const patchChip = (localId: string, patch: Partial<Chip>) =>
    setAttachments((prev) => prev.map((c) => (c.localId === localId ? { ...c, ...patch } : c)));

  // ---- selection context: auto-read, no manual button --------------------
  const refreshTarget = () => {
    getDocState()
      .then((s) =>
        setTarget({
          hasSelection: s.hasSelection,
          preview: s.selectionText,
          location: s.location,
          documentName: s.documentName,
        })
      )
      .catch(() => undefined);
  };

  useEffect(() => {
    refreshTarget();
    return subscribeSelection(refreshTarget);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    transcriptEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, busy]);

  // ---- attachments -------------------------------------------------------
  const onFilesSelected = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files ?? []);
    e.target.value = "";
    for (const file of files) void ingestFile(file);
  };

  const ingestFile = async (file: File) => {
    const localId = crypto.randomUUID();
    const abort = new AbortController();
    setAttachments((prev) => [...prev, { localId, name: file.name, status: "queued", abort }]);

    let id: string;
    try {
      const created = await uploadAttachment(file, sessionId.current, abort.signal);
      id = created.id;
      patchChip(localId, { id, status: created.status });
    } catch (err) {
      if ((err as Error).name === "AbortError") return patchChip(localId, { status: "cancelled" });
      return patchChip(localId, { status: "error", error: String((err as Error).message || err) });
    }

    for (;;) {
      if (abort.signal.aborted) return;
      await sleep(700, abort.signal);
      if (abort.signal.aborted) return;
      try {
        const state = await getAttachment(id, abort.signal);
        patchChip(localId, { status: state.status, error: state.error });
        if (TERMINAL.includes(state.status)) return;
      } catch {
        return;
      }
    }
  };

  const removeChip = (chip: Chip) => {
    chip.abort.abort();
    if (chip.id) void deleteAttachment(chip.id);
    setAttachments((prev) => prev.filter((c) => c.localId !== chip.localId));
  };

  // ---- send --------------------------------------------------------------
  const onSend = async () => {
    if (!canSend) return;
    const text = instruction.trim();
    const attachedCount = readyIds.length;
    pushMessage({
      role: "user",
      text,
      caption: attachedCount ? `${attachedCount} reference${attachedCount > 1 ? "s" : ""} attached` : undefined,
    });
    setInstruction("");
    setPending("Working…");
    try {
      const reference = attachedCount > 0 ? { attachmentIds: readyIds } : undefined;
      const r: EditResult = await runInstruction(text, setPending, reference);
      refreshTarget();
      if (!r.changed) {
        pushMessage({
          role: "assistant",
          kind: "info",
          text: "No change — the model returned the text unchanged for this instruction.",
        });
      } else {
        pushMessage({
          role: "assistant",
          kind: "result",
          text: r.edited,
          caption:
            r.mode === "draft"
              ? "Drafted and inserted as a tracked change — review it in the document."
              : "Edit written as a tracked change — review it in the document.",
        });
      }
    } catch (e) {
      pushMessage({ role: "assistant", kind: "error", text: String((e as Error).message || e) });
    } finally {
      setPending("");
    }
  };

  const onTextareaKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      void onSend();
    }
  };

  // Click anywhere in the box (except a button) focuses the textarea.
  const onBoxMouseDown = (e: React.MouseEvent) => {
    const el = e.target as HTMLElement;
    if (el.closest("button") || el === textareaRef.current) return;
    e.preventDefault();
    textareaRef.current?.focus();
  };

  const ctxLabel = target.hasSelection
    ? `Editing ${target.location}`
    : `Drafting into ${target.documentName || "the document"}`;

  return (
    <div className={styles.root}>
      <div className={styles.transcript}>
        {messages.length === 0 && (
          <div className={styles.empty}>
            <Text weight="semibold">Ask Silks AI to redline or draft</Text>
            <Text size={200}>
              Select a clause and describe the change, or attach a reference document and say how to
              use it. Every change lands as a tracked change you can accept or reject.
            </Text>
          </div>
        )}

        {messages.map((m) => (
          <div key={m.id} className={mergeClasses(styles.msgRow, m.role === "user" ? styles.userRow : styles.assistantRow)}>
            <div>
              <div
                className={mergeClasses(
                  styles.bubble,
                  m.role === "user" ? styles.userBubble : m.kind === "error" ? styles.errorBubble : styles.assistantBubble
                )}
              >
                {m.text}
              </div>
              {m.caption && <div className={styles.caption}>{m.caption}</div>}
            </div>
          </div>
        ))}

        {busy && (
          <div className={mergeClasses(styles.msgRow, styles.assistantRow)}>
            <div className={styles.typingRow}>
              <Spinner size="extra-tiny" />
              <span>{pending}</span>
            </div>
          </div>
        )}

        <div ref={transcriptEndRef} />
      </div>

      <div className={styles.dock}>
        <div className={styles.contextPill}>
          {target.hasSelection ? <DocumentTextRegular /> : <DocumentRegular />}
          <Tooltip content={target.hasSelection ? `“${truncate(target.preview, 140)}”` : ctxLabel} relationship="label">
            <span className={styles.contextText}>{ctxLabel}</span>
          </Tooltip>
        </div>

        <div className={styles.inputBox} onMouseDown={onBoxMouseDown}>
          {attachments.length > 0 && (
            <div className={styles.chipRow}>
              {attachments.map((c) => (
                <span
                  key={c.localId}
                  className={mergeClasses(styles.chip, c.status === "ready" && styles.chipReady, c.status === "error" && styles.chipError)}
                  title={c.status === "error" && c.error ? c.error : `${c.name} — ${STATE_LABEL[c.status]}`}
                >
                  {!TERMINAL.includes(c.status) ? (
                    <Spinner size="extra-tiny" />
                  ) : c.status === "ready" ? (
                    <CheckmarkCircleRegular />
                  ) : c.status === "error" ? (
                    <ErrorCircleRegular />
                  ) : (
                    <DocumentRegular />
                  )}
                  <span className={styles.chipName}>{c.name}</span>
                  <Button appearance="subtle" size="small" className={styles.chipBtn} icon={<DismissRegular />} aria-label="Remove attachment" onClick={() => removeChip(c)} />
                </span>
              ))}
            </div>
          )}

          <textarea
            ref={textareaRef}
            className={styles.textarea}
            rows={2}
            value={instruction}
            placeholder={target.hasSelection ? "Describe the edit…  (Enter to send)" : "Describe what to draft…  (Enter to send)"}
            onChange={(e) => setInstruction(e.target.value)}
            onKeyDown={onTextareaKeyDown}
          />

          <div className={styles.toolbar}>
            <Tooltip content="Attach a reference document (.docx, .pdf, .txt)" relationship="label">
              <Button appearance="subtle" icon={<AttachRegular />} onClick={() => fileInputRef.current?.click()} disabled={busy} aria-label="Attach a reference document" />
            </Tooltip>
            <input ref={fileInputRef} type="file" multiple accept=".pdf,.docx,.txt" style={{ display: "none" }} onChange={onFilesSelected} />
            <Tooltip content={target.hasSelection ? "Edit selection" : "Draft into document"} relationship="label">
              <Button
                appearance="primary"
                className={styles.sendBtn}
                icon={busy ? <Spinner size="extra-tiny" /> : <SendRegular />}
                onClick={onSend}
                disabled={!canSend}
                aria-label={target.hasSelection ? "Edit selection" : "Draft into document"}
              />
            </Tooltip>
          </div>
        </div>
      </div>
    </div>
  );
};

function truncate(s: string, n: number): string {
  return s.length > n ? s.slice(0, n - 1) + "…" : s;
}

// Promise sleep that resolves early if the signal aborts (stops chip polling
// promptly on remove).
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
