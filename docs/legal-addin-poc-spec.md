# Legal AI Word Add-in — POC Build Spec

> Self-contained build brief for Claude Code. The goal is a working proof of concept that proves the core loop: select text in Word → ask a local model to edit it → land the edit as a native tracked change the user can accept or reject → log it. The stack is production-grade in shape, but two layers are deliberately removed for the POC (see "Scope cuts").

---

## 1. POC goal (the one loop that must work)

A user has a `.docx` open in Microsoft Word. Our task pane add-in is loaded. The user:

1. Selects a span of text (e.g. a clause).
2. Types an instruction ("tighten this", "make this mutual", "redline against plain-English standard").
3. Clicks Edit.
4. The add-in sends the selected text + instruction to our local backend.
5. The backend calls **Ollama** (running locally) with a legal-tuned system prompt and returns the edited text.
6. The add-in turns on Track Changes and writes the edit back into the same range, so it appears as a native redline.
7. The user accepts or rejects it through Word's normal review flow.
8. Every edit is written to an append-only audit log.

If that loop runs reliably across a few clause types, the POC is a success.

---

## 2. Scope cuts (read this before building)

We are keeping the production architecture shape but removing two layers for the POC. **Do not build the removed layers, and do not build anything that depends on them.**

| Layer | Prod | POC | Why |
|---|---|---|---|
| AI model | Anthropic API, key held server-side | **Ollama, local** | No key, no cost, runs offline. Kept behind the same adapter so the swap back is one file. |
| Auth | Entra ID SSO + On-Behalf-Of | **None** | POC runs on a local machine for a single developer. |
| OneDrive / Graph sync | DriveItem versions + webhooks | **Removed** | Graph requires Entra auth, which we cut. POC works on a local `.docx` opened directly in Word. |
| Standalone document app | Re-reads file after sync | **Removed** | Same reason — the round-trip needs the sync layer. |
| Multi-tenant isolation | Required | **Removed** | Single user, single machine. |

**Knock-on effect to keep in mind:** because auth and Graph are gone, there is no "app re-reads from OneDrive" step in the POC. The document Word has open *is* the source of truth. Everything happens inside the one open Word session. This is intentional.

**What we deliberately keep** (these carry the actual value and de-risk the hard parts):

- Model-agnostic provider adapter, so Ollama → Anthropic later is a single-file change.
- Backend proxy pattern (client never calls the model directly), so the prod boundary already exists.
- Tracked-changes editing mechanism (the real technical risk worth proving).
- Append-only audit log (the legal differentiation; cheap to include now).
- Legal-tuned system prompt with anti-hallucination guardrails.

---

## 3. Architecture (POC)

```
┌─────────────────────────────┐
│  Microsoft Word (desktop)   │
│                             │
│  ┌───────────────────────┐  │
│  │  Task Pane Add-in     │  │   Office.js (Word.run)
│  │  React + Fluent UI    │  │   getSelection → insertText
│  │  TypeScript           │  │   changeTrackingMode = trackAll
│  └──────────┬────────────┘  │
└─────────────┼───────────────┘
              │  POST /api/edit  { text, instruction }
              ▼
┌─────────────────────────────┐
│  Local backend (Node)       │
│  Express/Fastify            │
│                             │
│  ModelProvider (interface)  │
│   └── OllamaProvider  ◄──────┼──► Ollama @ localhost:11434
│   └── AnthropicProvider (stub)
│                             │
│  Legal system prompt        │
│  Append-only audit log  ────┼──► SQLite (POC) / Postgres (prod)
└─────────────────────────────┘
```

---

## 4. Tech stack (POC)

| Layer | Choice | Notes |
|---|---|---|
| Editor / IDE | VS Code | Visual Studio add-in dev is deprecated; use VS Code. |
| Runtime | Node.js LTS | Generators, build, and backend all need it. |
| Scaffold | `yo office` (Yeoman) | Task Pane project, TypeScript. |
| Language | TypeScript | Client and server. |
| UI | React + Fluent UI React | Matches native Office look. |
| Manifest | **Add-in-only XML manifest** | Unified JSON manifest is still preview for Word as of 2026; XML is the production-safe choice. |
| Document layer | Office.js (Word JavaScript API) | Only sanctioned in-document editing surface. |
| Backend | Node + Express (or Fastify) | Proxies the model, holds the adapter, writes the audit log. |
| Model | **Ollama**, local | Native API at `http://localhost:11434`. |
| Model adapter | Custom `ModelProvider` interface | `OllamaProvider` live, `AnthropicProvider` stubbed. |
| Audit DB | SQLite (POC) → Postgres (prod) | Append-only. SQLite keeps the POC zero-config; schema is portable to Postgres. |
| Transport | HTTPS (dev cert) | Office requires HTTPS even in dev; the generator sets up the cert. |

**Model recommendation:** use a general instruct model, not a code model. `qwen2.5-coder` (already installed locally) is tuned for code and will underperform on legal prose. Pull a general instruct model instead, e.g. `qwen2.5:14b-instruct` or `llama3.1:8b-instruct`, and set it via an env var so it's swappable. Larger instruct models give noticeably better redlines if the machine can run them.

---

## 5. Prerequisites (developer machine)

- Node.js LTS installed.
- VS Code with the Office Add-in debugging extension.
- Ollama installed and running, with a general instruct model pulled:
  ```bash
  ollama pull qwen2.5:14b-instruct   # or llama3.1:8b-instruct
  ollama serve                       # if not already running as a service
  ```
- `yo` and the Office generator:
  ```bash
  npm install -g yo generator-office
  ```
- Word desktop (Windows or Mac) for sideloading. Test the web target later; desktop first.
- Dev HTTPS certificate trusted (the generator prompts for this on first run).

---

## 6. Suggested project structure

```
legal-addin-poc/
├── manifest.xml                 # add-in-only XML manifest
├── package.json
├── tsconfig.json
├── webpack.config.js
├── src/
│   └── taskpane/
│       ├── taskpane.html
│       ├── index.tsx            # mounts React
│       ├── components/
│       │   └── App.tsx          # instruction box, Edit button, status
│       └── word/
│           └── editor.ts        # all Word.run logic lives here
└── server/
    ├── index.ts                 # Express/Fastify app, CORS, HTTPS
    ├── routes/
    │   └── edit.ts              # POST /api/edit
    ├── providers/
    │   ├── ModelProvider.ts     # interface
    │   ├── OllamaProvider.ts     # live
    │   └── AnthropicProvider.ts  # stub, throws "not configured"
    ├── prompts/
    │   └── legal.ts             # system prompt
    └── db/
        ├── schema.sql
        └── audit.ts             # append-only writes
```

---

## 7. Build order (milestones)

1. **Scaffold + sideload.** `yo office` → Task Pane, React, TypeScript, XML manifest. Sideload the hello-world pane into Word. Confirm HTTPS dev cert is trusted.
2. **Read selection.** Wire the pane to read the current Word selection and display it. Proves Office.js plumbing.
3. **Backend proxy → Ollama.** Stand up the Node server with `POST /api/edit`. Implement `OllamaProvider`. Return edited text. Test the endpoint with curl before touching the client.
4. **Apply edit as tracked change.** Turn on `changeTrackingMode = trackAll`, write the edit back to the range, confirm it appears as a redline. Verify Word's Accept/Reject works on it.
5. **Audit log.** Write each edit to SQLite (instruction, model, original, output, status, timestamp).
6. **Legal prompt + guardrails.** Drop in the legal system prompt; verify the model does not invent clauses or citations and flags rather than fabricates.

Stop here for the POC. Auth, Graph sync, the standalone app, multi-tenant, and the vector/clause-library retrieval layer are all post-POC.

---

## 8. Core editing mechanism (client)

All document interaction runs inside `Word.run`. The endpoint is three Office.js calls plus the tracked-changes toggle.

```ts
// src/taskpane/word/editor.ts
export async function editSelection(instruction: string): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();

    const original = range.text;
    if (!original.trim()) {
      // surface "select some text first" in the UI
      return;
    }

    // backend proxies to Ollama; client never talks to the model directly
    const res = await fetch("https://localhost:3001/api/edit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text: original, instruction }),
    });
    const { text: edited } = await res.json();

    // land the change as a native redline
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    range.insertText(edited, Word.InsertLocation.replace);
    await context.sync();
  });
}
```

Notes for the implementer:
- Gate any newer Office.js APIs with `Office.context.requirements.isSetSupported` so it doesn't break on older Word builds.
- For the POC, scope edits to the selection only. Do not send whole documents to the model. Full-document passes and chunking are post-POC.
- `insertText` replace flattens formatting on the range. That is acceptable for the POC. The fidelity ladder (`insertOoxml`, content controls, word-level diff) is a prod concern; note it but don't build it now.

---

## 9. Backend spec

### Endpoint

`POST /api/edit`

Request:
```json
{ "text": "the selected clause text", "instruction": "make this mutual" }
```

Response:
```json
{ "text": "the edited clause text" }
```

Server flow: receive → build messages (legal system prompt + instruction + text) → call `ModelProvider.edit()` → write audit row → return edited text. Enable CORS for the add-in origin and serve over HTTPS (reuse the dev cert).

### Model adapter

Keep the provider boundary clean so the prod swap is trivial.

```ts
// server/providers/ModelProvider.ts
export interface EditRequest {
  text: string;
  instruction: string;
}
export interface ModelProvider {
  name(): string;          // for the audit log
  version(): string;       // model tag, for the audit log
  edit(req: EditRequest): Promise<string>;
}
```

```ts
// server/providers/OllamaProvider.ts
import { ModelProvider, EditRequest } from "./ModelProvider";
import { LEGAL_SYSTEM_PROMPT } from "../prompts/legal";

const MODEL = process.env.OLLAMA_MODEL ?? "qwen2.5:14b-instruct";
const HOST = process.env.OLLAMA_HOST ?? "http://localhost:11434";

export class OllamaProvider implements ModelProvider {
  name() { return "ollama"; }
  version() { return MODEL; }

  async edit({ text, instruction }: EditRequest): Promise<string> {
    const res = await fetch(`${HOST}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: MODEL,
        stream: false,
        messages: [
          { role: "system", content: LEGAL_SYSTEM_PROMPT },
          {
            role: "user",
            content:
              `Instruction: ${instruction}\n\n` +
              `Text to edit:\n${text}\n\n` +
              `Return ONLY the edited text, no commentary, no markdown.`,
          },
        ],
      }),
    });
    const data = await res.json();
    return data.message.content.trim();
  }
}
```

`AnthropicProvider` is a stub that throws "not configured for POC" — present so the interface and swap path are real, not built out.

Provider selection via env: `MODEL_PROVIDER=ollama` (default) picks `OllamaProvider`.

### Ollama API notes

- Native chat endpoint: `POST http://localhost:11434/api/chat`, body takes `model`, `messages`, `stream`. Set `stream: false` for the POC so you get one JSON response; content is at `data.message.content`.
- Ollama also exposes an OpenAI-compatible endpoint at `/v1/chat/completions` if a future provider prefers that shape.
- First call after `ollama serve` can be slow while the model loads into memory. Add a generous timeout on the backend fetch.

---

## 10. Audit log

Append-only. SQLite for the POC (zero config), schema written so it ports straight to Postgres.

```sql
-- server/db/schema.sql
CREATE TABLE IF NOT EXISTS edit_log (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  created_at    TEXT    NOT NULL,   -- ISO timestamp
  doc_name      TEXT,
  instruction   TEXT    NOT NULL,
  model_name    TEXT    NOT NULL,
  model_version TEXT    NOT NULL,
  original_text TEXT    NOT NULL,
  edited_text   TEXT    NOT NULL,
  status        TEXT    NOT NULL DEFAULT 'pending'  -- pending | accepted | rejected
);
```

For the POC, write the row on `/api/edit` with status `pending`. Capturing accept/reject status is a stretch goal (Word fires events for tracked-change resolution; wire it only if time allows). The point is to prove the log exists and records the full who/what/which-model trail that Word's own version history cannot.

---

## 11. Legal system prompt (starting point)

Keep it short, redline-focused, and guardrailed. Refine during milestone 6.

```
You are a legal drafting assistant. You edit contract and legal text precisely
according to the user's instruction.

Rules:
- Return only the edited text. No preamble, no explanation, no markdown.
- Preserve defined terms, capitalization of defined terms, and numbering exactly
  unless the instruction asks you to change them.
- Never invent clauses, parties, dates, figures, or legal citations. If the
  instruction asks for something the text does not support, make the minimal
  honest edit and do not fabricate.
- Prefer clear, plain drafting. Do not add boilerplate the instruction did not ask for.
```

---

## 12. POC acceptance criteria

The POC is done when all of these hold:

1. The task pane sideloads into Word desktop over HTTPS.
2. Selecting text and clicking Edit reads the correct selection.
3. The backend reaches Ollama and returns edited text for the selection.
4. The edit lands in the document as a native tracked change.
5. Word's Accept and Reject both work on that change.
6. Each edit produces a row in the audit log with instruction, model name, model version, original, and edited text.
7. Switching the model is an env-var change, and switching to Anthropic later is a single-file change (the stub proves the seam).

---

## 13. Deferred to prod (do not build now)

Entra SSO and On-Behalf-Of; Microsoft Graph file sync, versions, and webhooks; the standalone document app and its re-read loop; multi-tenant isolation; zero-retention / no-training model configuration and the M365-boundary ToS; the fidelity ladder beyond `insertText` (`insertOoxml`, content controls, word-level diff); full-document passes and chunking/retrieval; the vector DB, clause library, jurisdiction awareness, and playbook-compare mode; AppSource or private M365 admin-center distribution.
