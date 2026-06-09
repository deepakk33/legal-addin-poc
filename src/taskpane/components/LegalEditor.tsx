import * as React from "react";
import { useState } from "react";
import {
  Button,
  Field,
  Textarea,
  Text,
  Spinner,
  Card,
  tokens,
  makeStyles,
} from "@fluentui/react-components";
import { runInstruction, getDocState, readSelection, EditResult } from "../word/editor";

/* global HTMLTextAreaElement */

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
  card: { padding: "10px" },
  label: { fontWeight: tokens.fontWeightSemibold, marginBottom: "4px" },
  mono: { whiteSpace: "pre-wrap", fontSize: tokens.fontSizeBase200 },
});

const LegalEditor: React.FC = () => {
  const styles = useStyles();
  const [instruction, setInstruction] = useState<string>("Make this clause mutual.");
  const [selection, setSelection] = useState<string>("");
  const [hasSelection, setHasSelection] = useState<boolean>(false);
  const [result, setResult] = useState<EditResult | null>(null);
  const [status, setStatus] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);

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
      const r = await runInstruction(instruction, setStatus);
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
        Legal AI redline
      </Text>
      <Text className={styles.status}>
        Select a clause and type an instruction to redline it, or select nothing and type an
        instruction to draft new text. Either way the change lands as a tracked change you can
        accept or reject.
      </Text>

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
              {result.mode === "draft" ? "Drafted (inserted as tracked change)" : "AI edit (written to document)"}
            </div>
            <Text className={styles.mono}>{result.edited}</Text>
          </Card>
        </div>
      )}
    </div>
  );
};

export default LegalEditor;
