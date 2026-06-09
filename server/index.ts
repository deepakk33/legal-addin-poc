import express from "express";
import cors from "cors";
import * as https from "https";
import devCerts from "office-addin-dev-certs";
import editRouter from "./routes/edit";
import { dbPath } from "./db/audit";

const PORT = Number(process.env.BACKEND_PORT ?? 3001);
// The add-in runs on the webpack dev server origin.
const ADDIN_ORIGIN = process.env.ADDIN_ORIGIN ?? "https://localhost:3000";

async function main() {
  const app = express();
  app.use(cors({ origin: ADDIN_ORIGIN }));
  app.use(express.json({ limit: "1mb" }));

  app.get("/health", (_req, res) => res.json({ ok: true }));
  app.use("/api", editRouter);

  // Reuse the Office dev cert so the add-in (HTTPS) can call us without cert errors.
  const httpsOptions = await devCerts.getHttpsServerOptions();

  https
    .createServer({ key: httpsOptions.key, cert: httpsOptions.cert }, app)
    .listen(PORT, () => {
      // eslint-disable-next-line no-console
      console.log(`Legal add-in backend on https://localhost:${PORT}`);
      // eslint-disable-next-line no-console
      console.log(`  provider:   ${process.env.MODEL_PROVIDER ?? "ollama"}`);
      // eslint-disable-next-line no-console
      console.log(`  ollama:     ${process.env.OLLAMA_MODEL ?? "llama3.1:8b"} @ ${process.env.OLLAMA_HOST ?? "http://localhost:11434"}`);
      // eslint-disable-next-line no-console
      console.log(`  audit db:   ${dbPath()}`);
      // eslint-disable-next-line no-console
      console.log(`  cors origin:${ADDIN_ORIGIN}`);
    });
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error("Failed to start backend:", err);
  process.exit(1);
});
