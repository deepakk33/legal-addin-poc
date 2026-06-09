import Database from "better-sqlite3";
import * as fs from "fs";
import * as path from "path";

// Append-only audit log. POC writes one row per /api/edit with status 'pending'.
const DB_PATH = process.env.AUDIT_DB_PATH ?? path.join(__dirname, "audit.db");
const SCHEMA_PATH = path.join(__dirname, "schema.sql");

const db = new Database(DB_PATH);
db.pragma("journal_mode = WAL");
db.exec(fs.readFileSync(SCHEMA_PATH, "utf8"));

export interface AuditRow {
  docName?: string | null;
  instruction: string;
  modelName: string;
  modelVersion: string;
  originalText: string;
  editedText: string;
}

const insertStmt = db.prepare(
  `INSERT INTO edit_log
     (created_at, doc_name, instruction, model_name, model_version, original_text, edited_text, status)
   VALUES
     (@created_at, @doc_name, @instruction, @model_name, @model_version, @original_text, @edited_text, 'pending')`
);

// Returns the new row id.
export function recordEdit(row: AuditRow): number {
  const info = insertStmt.run({
    created_at: new Date().toISOString(),
    doc_name: row.docName ?? null,
    instruction: row.instruction,
    model_name: row.modelName,
    model_version: row.modelVersion,
    original_text: row.originalText,
    edited_text: row.editedText,
  });
  return Number(info.lastInsertRowid);
}

export function dbPath(): string {
  return DB_PATH;
}
