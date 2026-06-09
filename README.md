# Legal AI Word Add-in — POC

A Microsoft Word task-pane add-in that redlines selected legal text with an AI and lands the
edit as a **native tracked change** the lawyer accepts or rejects. Our own product (not Copilot),
bring-your-own-model behind an adapter. This is the POC that proves the one core loop:

> select text in Word → ask an AI to edit it → land it as a tracked change → audit-log it

The model backend is swappable — **OpenAI (default), Claude, or local Ollama** — behind one
`ModelProvider` seam. Full brief in [`docs/legal-addin-poc-spec.md`](docs/legal-addin-poc-spec.md);
product/strategy context in [`docs/legal-word-addin-context-handoff.md`](docs/legal-word-addin-context-handoff.md).

## Layout

```
legal-addin-poc/
├── manifest.xml              # add-in-only XML manifest
├── src/taskpane/
│   ├── components/App.tsx        # pane shell
│   ├── components/LegalEditor.tsx # instruction box, Read selection, Edit button, preview
│   └── word/editor.ts            # all Word.run logic: getSelection → trackAll → insertText replace
├── server/
│   ├── index.ts                  # Express app over HTTPS (reuses Office dev cert), loads .env
│   ├── routes/edit.ts            # POST /api/edit
│   ├── providers/                # ModelProvider seam: OpenAIProvider, AnthropicProvider, OllamaProvider
│   ├── prompts/legal.ts          # legal system prompt + guardrails
│   └── db/                       # SQLite append-only audit log (schema.sql, audit.ts)
└── docs/                         # the two canonical spec docs
```

## Requirements (any machine)

| What | Why | Notes |
|---|---|---|
| **Node.js 22 LTS** | build tooling + native `better-sqlite3` reject non-LTS node | nvm/fnm/Homebrew/installer. Node 25 is **rejected** — see Troubleshooting. |
| **Microsoft Word** (desktop) | the add-in sideloads into Word | Windows or Mac; web target untested. |
| **A model API key** | the AI that does the editing | **OpenAI** key by default. Or a Claude key. Or **Ollama** (local, no key) — see Switching models. |
| **git** | to clone | repo is public. |

> Keys are read **server-side only** — they never reach the add-in client bundle. The open `.docx`
> in Word is the source of truth (this POC has no cloud file sync).

## Quick start (fresh machine)

```bash
# 1. Clone
git clone https://github.com/deepakk33/legal-addin-poc.git
cd legal-addin-poc

# 2. Use Node 22 LTS (example with nvm; or fnm / Homebrew node@22)
nvm install 22 && nvm use 22
node -v        # must be v22.x

# 3. Trust the local dev HTTPS certificate (one-time, interactive — installs a local CA)
npx office-addin-dev-certs install

# 4. Install deps (root = add-in, server = backend)
npm install
( cd server && npm install )

# 5. Configure the backend key (default backend is OpenAI)
cd server
cp .env.example .env
#    edit .env → paste OPENAI_API_KEY=sk-...
cd ..
```

> Using **Claude** or **Ollama** instead of OpenAI? See **Switching models** below before step 5.

## Run (two terminals, both on Node 22)

**Terminal 1 — backend** (model proxy + audit log):
```bash
cd server
npm start          # https://localhost:3001  (auto-loads server/.env via dotenv)
```
`server/.env` is gitignored — your key is never committed.

**Terminal 2 — add-in** (task pane + dev server + sideload), from the repo root:
```bash
npm start          # webpack dev server on https://localhost:3000, sideloads into Word
```
`npm start` auto-sideloads on Windows/Mac. If Word doesn't open the pane, sideload `manifest.xml`
manually (Word → **Insert → Add-ins → Upload My Add-in**), or `npm stop` then `npm start` again.

## Switching models — OpenAI / Claude / Ollama

The backend is model-agnostic behind a `ModelProvider` seam. Set `MODEL_PROVIDER` in `server/.env`
(or as an env var) and restart the backend. Nothing else changes.

| `MODEL_PROVIDER` | Provider | Key (server-side) | Model env (default) |
|---|---|---|---|
| `openai` (default) | OpenAI API | `OPENAI_API_KEY` | `OPENAI_MODEL` (`gpt-4o`) |
| `anthropic` | Claude API | `ANTHROPIC_API_KEY` | `ANTHROPIC_MODEL` (`claude-sonnet-4-6`) |
| `ollama` | local Ollama | none | `OLLAMA_MODEL` (`llama3.1:8b`) |

Set it in `server/.env`:
```bash
# OpenAI (default)
MODEL_PROVIDER=openai
OPENAI_API_KEY=sk-...

# Claude
MODEL_PROVIDER=anthropic
ANTHROPIC_API_KEY=sk-ant-...
ANTHROPIC_MODEL=claude-sonnet-4-6

# Ollama (local, no key) — first: ollama pull llama3.1:8b && ollama serve
MODEL_PROVIDER=ollama
OLLAMA_MODEL=llama3.1:8b
```

The provider name + model version are recorded in every audit-log row, so the trail shows
exactly which model produced each edit.

> Notes: use a general **instruct** model, not a code model (Ollama: `llama3.1:8b`, not `*-coder`).
> Claude Opus 4.x rejects `temperature`/`budget_tokens`, so the Anthropic provider omits them and
> relies on the "return only the edited text" system-prompt guardrail; the OpenAI provider sends
> `temperature: 0.2`. All non-streaming (a clause is short); raise `ANTHROPIC_MAX_TOKENS` for long passages.

## Use it end to end

1. Open a `.docx` in Word. The **Legal AI Redline** pane appears (Home tab → Show Taskpane, if not).
2. **Select a clause** in the document.
3. (Optional) click **Read selection** to confirm the pane sees your text.
4. Type an **instruction** — e.g. *"make this mutual"*, *"tighten this"*, *"redline against plain-English standard"*.
5. Click **Edit selection**. The backend calls the configured model; the edited text is written
   back into the same range **as a tracked change** (redline).
6. Resolve it with Word's normal **Review → Accept / Reject**.
7. Every edit is recorded in the append-only audit log (`server/db/audit.db`).

## Verify without Word

After the dev cert is trusted and a key is set:
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

## Troubleshooting

- **Build complains about Node version** → you're on non-LTS node (e.g. 25). Switch to Node 22 (`nvm use 22`).
- **Backend crashes loading `better-sqlite3`** (ABI / `NODE_MODULE_VERSION` mismatch) → deps were
  installed under a different node. Re-run `npm install` in `server/` under Node 22, or `npm rebuild better-sqlite3`.
- **Add-in can't reach the backend / cert errors** → run `npx office-addin-dev-certs install` and
  restart both. The backend reuses the same Office dev cert as the add-in.
- **`... API key is not set`** → the selected `MODEL_PROVIDER` has no key in `server/.env`. Add the matching
  key (`OPENAI_API_KEY` / `ANTHROPIC_API_KEY`), or switch to `MODEL_PROVIDER=ollama` (no key).
- **Ollama: weak / code-like output** → use a general **instruct** model (`llama3.1:8b`), not a `*-coder`
  model. `llama3.1:8b-instruct` is NOT a real Ollama tag — `llama3.1:8b` is already instruct-tuned.
  First call after `ollama serve` is slow (cold load), not a hang (`OLLAMA_TIMEOUT_MS`, default 120s).

## Scope (POC)

**In:** the one loop above — selection-only edits, tracked-change insert, append-only audit log,
legal prompt with anti-hallucination guardrails, model-agnostic provider seam (OpenAI / Claude / Ollama).

**Out (prod, deliberately not built):** Entra SSO/On-Behalf-Of, Microsoft Graph sync/versions/
webhooks, the standalone document app, multi-tenant isolation, the fidelity ladder beyond
`insertText` (OOXML / content controls / word-level diff), full-document passes & chunking,
vector DB / clause library / playbook-compare, AppSource / admin-center distribution.
