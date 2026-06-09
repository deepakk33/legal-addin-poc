# Project Context — Legal AI Word Add-in

> Handoff brief. Paste or upload this into a new chat to continue with full context. It is self-contained: it captures the concept, the decisions made, the technical mechanism, the stack, the open questions, and the risks.

---

## TL;DR for a fresh chat

We are designing an enterprise product: a standalone application that displays documents, with a Microsoft Word add-in (task pane) that lets a user edit the open document by asking an AI. It is **not** Microsoft Copilot — it is our own product, using our own API key to Claude and/or other models (key held server-side), and it is **specialized for legal documents**. Edits sync back to the application via OneDrive. "Marketplace" in earlier discussion meant Office add-ins (AppSource distribution), not a separate store.

---

## Product concept & decisions locked

- Standalone app shows the document to the user.
- A button opens that document in Microsoft Word. The file lives on OneDrive/SharePoint.
- In Word, the user is already authenticated into our task pane add-in.
- The user selects text (or requests a full-document pass) and asks the AI to edit.
- Edits are applied in the document, then AutoSave writes back to OneDrive, and the app re-reads the file.
- Own product, not Copilot. Bring-your-own model: Claude and/or others, behind a model-agnostic adapter.
- API keys live on the server, never in the add-in client bundle.
- Specialized for legal: redlining, clause libraries, jurisdiction awareness are the differentiation.

---

## Strategic context (the bottleneck analysis)

Three bottlenecks were identified, in priority order:

1. **Incumbent / distribution (biggest strategic risk).** Microsoft Copilot already does AI drafting, full-doc creation, and agentic multi-step editing inside Word, bundled into M365, with native Track Changes for AI edits. The existential question is "why not just use Copilot." The answer / wedge: deep legal specialization plus bring-your-own model choice, a workflow Copilot does not tailor for.
2. **Edit fidelity (the technical risk that breaks UX).** Applying AI edits into a real legal doc while preserving styles, clause numbering, defined terms, cross-references, and producing clean tracked changes is the hard part. Calling the AI is trivial; landing the edit cleanly is not.
3. **Sync / source-of-truth (the flagged operational risk).** When both the app and an open Word session can edit the same file, you must define authority or you get silent overwrites — very bad in a legal context.

---

## How the AI editing actually works (the precise mechanism)

Everything runs inside `Word.run`, which batches commands and flushes them with `context.sync()`.

Core endpoint = three Office.js calls:
1. `context.document.getSelection()` returns a Range for the highlighted text.
2. `range.load("text")` then `await context.sync()` pulls the text out.
3. `range.insertText(editedText, Word.InsertLocation.replace)` writes the AI version back.

For legal, set `context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll` **before** the write, so the edit lands as a native redline the lawyer accepts or rejects.

```js
await Word.run(async (context) => {
  const range = context.document.getSelection();   // what the user highlighted
  range.load("text");
  await context.sync();

  const original = range.text;

  // backend holds the key and calls Claude
  const edited = await fetch("/api/edit", {
    method: "POST",
    body: JSON.stringify({ text: original, instruction: userPrompt })
  }).then(r => r.json()).then(d => d.text);

  context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  range.insertText(edited, Word.InsertLocation.replace);   // lands as a redline
  await context.sync();
});
```

For "create a new doc," operate on `context.document.body` with `body.insertText(...)` / `body.insertParagraph(...)`.

### Fidelity ladder (choose per edit type)

| Method | Use when | Trade-off |
|---|---|---|
| `insertText` replace | Plain prose rewrite | Flattens formatting on the range |
| `insertOoxml` | Preserve/inject styles, structure | More complex to generate |
| Content controls (`insertContentControl`) | Bounded, lockable clause slots | Setup overhead |
| Word-level diff (e.g. office-word-diff) | Surgical word changes as clean redlines | Extra library, best legal UX |

Long contracts: scope to selection or chunk/retrieve. Never send an 80-page doc per call.

---

## Recommended stack

| Layer | Choice | Notes |
|---|---|---|
| Scaffold | `yo office` (Yeoman) or M365 Agents Toolkit | Microsoft's recommended generators |
| Language | TypeScript | |
| UI | React + Fluent UI React | Matches Office look |
| Manifest | Add-in-only XML manifest | Unified JSON manifest still preview, not production |
| Document layer | Office.js (Word JavaScript API) | Only sanctioned in-app editing surface |
| Backend | Node (Express/Fastify) | Holds secrets, proxies AI |
| AI | Anthropic API behind model-agnostic adapter | Swap/add models freely |
| Auth | Entra ID SSO (nested app auth) + On-Behalf-Of | SSO into Graph |
| File + sync | OneDrive/SharePoint + Microsoft Graph | Versions, webhooks, delta |
| Audit DB | Postgres, append-only | Compliance trail |
| Retrieval | Vector DB (pgvector/Pinecone) | Clause libraries, precedents |
| Hosting | Azure | HTTPS mandatory |

Environment gotchas: Node LTS; HTTPS required even in dev; build in VS Code (Visual Studio add-in dev deprecated as of VS 2026); test Word on Windows, Mac, and web separately and gate newer APIs with `Office.context.requirements.isSetSupported`.

---

## Version history & sync-back (two layers)

**Layer A — document state (free, do not rebuild).** OneDrive/SharePoint versions every save; AutoSave + Microsoft Graph (DriveItem versions endpoint) gives document history with no custom code.

**Layer B — legal audit trail (must build).** Append-only log per edit: user identity (Entra), the instruction, model + version, original text, AI output, accepted/rejected status, timestamp, matter/client ID. Word's version history cannot provide this.

**Sync mechanics.** Use Graph change notifications (webhooks) or delta queries, not polling. Define source of truth: recommended that the Word session is authoritative while the doc is open, and the app locks editing during that window. Handle save races.

---

## Security & distribution

- API keys server-side only.
- Multi-tenant isolation; zero cross-firm leakage.
- Zero-retention AI configuration, no training on customer data, encryption in transit and at rest.
- Be explicit in ToS about what content leaves the M365 boundary (legal privilege).
- Production = add-in-only XML manifest.
- Distribute via public AppSource (Partner Center validation) or private M365 admin center deployment for a single firm (faster, no public review).

---

## Legal specialization (the moat)

The editing window is commodity. These make it a legal product worth paying for:
- Legal-tuned system prompt for drafting and redlining.
- Clause library + retrieval over precedents (vector DB).
- Jurisdiction awareness.
- "Redline against our standard terms / playbook" comparison mode.
- Defined-term and citation consistency checks.
- Hallucination guardrails: never invent clauses or citations; flag, do not fabricate.

---

## Suggested build order

1. Hello-world task pane (`yo office`, React + TS, XML manifest), sideload in Word.
2. Read selection, display in pane.
3. Backend proxy → Claude → return edited text.
4. Apply edit as tracked change, with Keep/Discard.
5. OneDrive file + Graph version read; sync-back via webhook.
6. Append-only audit log DB.
7. SSO + multi-tenant isolation.
8. Legal specialization layer (retrieval, clause library, playbook compare).
9. AppSource or private deployment.

---

## Open questions still to decide

- Selection-only edits, full-document passes, or both.
- Single task pane vs richer two-pane review UI.
- Per-firm bring-your-own key vs metered backend billing.
- Public AppSource vs private single-firm deployment.
- Target platforms: desktop only, or also Word on the web.

---

## Top risks to keep front of mind

1. Source-of-truth conflicts causing silent overwrites (legal-critical).
2. Competing against free, bundled Copilot — the wedge must stay sharp.
3. Edit fidelity on complex legal formatting.
