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
import { editSelection, readSelection, EditResult } from "../word/editor";

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
  const [result, setResult] = useState<EditResult | null>(null);
  const [status, setStatus] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [busy, setBusy] = useState<boolean>(false);

  const onReadSelection = async () => {
    setError("");
    try {
      const text = await readSelection();
      setSelection(text.trim() || "(nothing selected)");
    } catch (e) {
      setError(String((e as Error).message || e));
    }
  };

  const onEdit = async () => {
    setError("");
    setResult(null);
    setBusy(true);
    setStatus("Reading selection…");
    try {
      const r = await editSelection(instruction, setStatus);
      setResult(r);
      setSelection(r.original);
      setStatus("Done — review the tracked change in the document.");
    } catch (e) {
      setError(String((e as Error).message || e));
      setStatus("");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.root}>
      <Text size={400} weight="semibold">
        Legal AI redline
      </Text>
      <Text className={styles.status}>
        Select a clause in the document, type an instruction, then Edit. The change lands as a
        tracked change you can accept or reject.
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
          Edit selection
        </Button>
        {busy && <Spinner size="tiny" />}
      </div>

      {status && <Text className={styles.status}>{status}</Text>}
      {error && <Text className={styles.error}>{error}</Text>}

      {result && (
        <div className={styles.preview}>
          <Card className={styles.card}>
            <div className={styles.label}>AI edit (written to document)</div>
            <Text className={styles.mono}>{result.edited}</Text>
          </Card>
        </div>
      )}
    </div>
  );
};

export default LegalEditor;
