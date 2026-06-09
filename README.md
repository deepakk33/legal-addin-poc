# Legal AI Word Add-in — POC

A Microsoft Word task-pane add-in that redlines selected legal text with an AI and lands the
edit as a **native tracked change** the lawyer accepts or rejects. Our own product (not Copilot),
bring-your-own-model behind an adapter. This is the POC that proves the one core loop:

> select text in Word → ask a local model to edit it → land it as a tracked change → audit-log it

Full brief in [`docs/legal-addin-poc-spec.md`](docs/legal-addin-poc-spec.md); product/strategy
context in [`docs/legal-word-addin-context-handoff.md`](docs/legal-word-addin-context-handoff.md).

## Layout

```
legal-addin-poc/
├── manifest.xml              # add-in-only XML manifest
├── src/taskpane/
│   ├── components/App.tsx        # pane shell
│   ├── components/LegalEditor.tsx # instruction box, Read selection, Edit button, preview
│   └── word/editor.ts            # all Word.run logic: getSelection → trackAll → insertText replace
├── server/
│   ├── index.ts                  # Express app over HTTPS (reuses Office dev cert)
│   ├── routes/edit.ts            # POST /api/edit
│   ├── providers/                # ModelProvider; OllamaProvider (live), AnthropicProvider (stub)
│   ├── prompts/legal.ts          # legal system prompt + guardrails
│   └── db/                       # SQLite append-only audit log (schema.sql, audit.ts)
└── docs/                         # the two canonical spec docs
```

## Requirements (any machine)

| What | Why | Notes |
|---|---|---|
| **Node.js 22 LTS** | `generator-office` + native `better-sqlite3` reject non-LTS node | Use nvm/fnm/Homebrew/installer. Node 25 is **rejected** — see Troubleshooting. |
| **Microsoft Word** (desktop) | the add-in sideloads into Word | Windows or Mac; web target untested. |
| **Ollama** | runs the model locally, offline, no API key | https://ollama.com |
| **gh / git** (optional) | only to clone | — |

> Everything runs **locally** — no cloud, no auth, no API key. The model is local (Ollama); the
> open `.docx` in Word is the source of truth.

## Quick start (fresh machine)

```bash
# 1. Clone
git clone https://github.com/deepakk33/legal-addin-poc.git
cd legal-addin-poc

# 2. Use Node 22 LTS (example with nvm; or fnm / Homebrew node@22)
nvm install 22 && nvm use 22
node -v        # must be v22.x

# 3. Pull the model (general INSTRUCT model — not a code model) and start Ollama
ollama pull llama3.1:8b
ollama serve   # leave running (skip if already a service)

# 4. Trust the local dev HTTPS certificate (one-time, interactive — installs a local CA)
npx office-addin-dev-certs install

# 5. Install deps (root = add-in, server = backend)
npm install
( cd server && npm install )
```

## Run (two terminals, both on Node 22)

**Terminal 1 — backend** (model proxy + audit log):
```bash
cd server
npm start          # https://localhost:3001
```
Config via env or `server/.env` (copy `server/.env.example`): `OLLAMA_MODEL` (default `llama3.1:8b`),
`OLLAMA_HOST`, `MODEL_PROVIDER` (`ollama` default; `anthropic` hits the stub), `BACKEND_PORT` (3001),
`ADDIN_ORIGIN` (`https://localhost:3000`), `OLLAMA_TIMEOUT_MS`, `AUDIT_DB_PATH`.

**Terminal 2 — add-in** (task pane + dev server + sideload), from the repo root:
```bash
npm start          # webpack dev server on https://localhost:3000, sideloads into Word
```
`npm start` auto-sideloads on Windows/Mac. If Word doesn't open the pane, sideload `manifest.xml`
manually (Word → **Insert → Add-ins → Upload My Add-in**), or `npm stop` then `npm start` again.

## Use it end to end

1. Open a `.docx` in Word. The **Legal AI Redline** pane appears (Home tab → Show Taskpane, if not).
2. **Select a clause** in the document.
3. (Optional) click **Read selection** to confirm the pane sees your text.
4. Type an **instruction** — e.g. *"make this mutual"*, *"tighten this"*, *"redline against plain-English standard"*.
5. Click **Edit selection**. The backend calls the local model; the edited text is written back into
   the same range **as a tracked change** (redline).
6. Resolve it with Word's normal **Review → Accept / Reject**.
7. Every edit is recorded in the append-only audit log (`server/db/audit.db`).

> First model call after `ollama serve` is slow (cold load) — backend timeout is generous
> (`OLLAMA_TIMEOUT_MS`, default 120s). Not a hang.

## Troubleshooting

- **`generator-office`/build complains about Node version** → you're on non-LTS node (e.g. 25).
  Switch to Node 22 (`nvm use 22`).
- **Backend crashes loading `better-sqlite3`** (ABI / `NODE_MODULE_VERSION` mismatch) → deps were
  installed under a different node. Re-run `npm install` in `server/` under Node 22, or
  `npm rebuild better-sqlite3`.
- **Add-in can't reach the backend / cert errors** → run `npx office-addin-dev-certs install` and
  restart both. The backend reuses the same Office dev cert as the add-in.
- **Model output is weak / code-like** → make sure `OLLAMA_MODEL` is a general **instruct** model
  (`llama3.1:8b`, `qwen2.5:14b-instruct`), not a `*-coder` model. Note `llama3.1:8b-instruct` is
  NOT a real Ollama tag — `llama3.1:8b` is already instruct-tuned.

## Verify without Word

Backend health + edit, with the dev cert trusted:
```bash
curl -k https://localhost:3001/health
curl -k https://localhost:3001/api/edit \
  -H 'Content-Type: application/json' \
  -d '{"text":"The Receiving Party shall not disclose Confidential Information.","instruction":"make this mutual"}'
```
Inspect the audit log:
```bash
sqlite3 server/db/audit.db 'select id, created_at, model_name, model_version, status, instruction from edit_log;'
```

> First model call after `ollama serve` is slow (cold load) — the backend timeout is generous
> (`OLLAMA_TIMEOUT_MS`, default 120s). Not a hang.

## Scope (POC)

**In:** the one loop above — selection-only edits, tracked-change insert, append-only audit log,
legal prompt with anti-hallucination guardrails, model-agnostic provider seam.

**Out (prod, deliberately not built):** Entra SSO/On-Behalf-Of, Microsoft Graph sync/versions/
webhooks, the standalone document app, multi-tenant isolation, the fidelity ladder beyond
`insertText` (OOXML / content controls / word-level diff), full-document passes & chunking,
vector DB / clause library / playbook-compare, AppSource / admin-center distribution.

Swapping the local model for Anthropic in prod is a single-file change behind `ModelProvider`
(`server/providers/AnthropicProvider.ts` is a deliberate stub proving the seam).
