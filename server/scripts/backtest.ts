import "dotenv/config";
import * as fs from "fs";
import * as path from "path";
import { getProvider } from "../providers";
import { extractText } from "../attachments/extract";
import { buildArtifact, projectArtifact } from "../attachments/artifact";
import { ReferenceContext } from "../providers/ModelProvider";

// In-process backtest: runs reference docs + prompts through the SAME ingestion
// and edit/draft pipeline the add-in uses (no HTTP, no Word), and writes results
// to backtest-results.md for manual review. Tune the M9 prompts / artifact schema
// based on the output, then re-run.
//
//   PATH="/opt/homebrew/opt/node@22/bin:$PATH" tsx server/scripts/backtest.ts

interface Case {
  name: string;
  referenceFile?: string; // path relative to fixtures/, optional
  referenceMode: ReferenceContext["mode"];
  instruction: string;
  selectionText?: string; // present => edit mode; absent => draft mode
  notes?: string;
}

const FIXTURES_DIR = path.join(__dirname, "fixtures");
const CASES_PATH = path.join(FIXTURES_DIR, "cases.json");
const OUT_PATH = path.join(__dirname, "backtest-results.md");

async function main() {
  if (!fs.existsSync(CASES_PATH)) {
    // eslint-disable-next-line no-console
    console.error(`No cases file at ${CASES_PATH}. Create it (see sample-nda.txt).`);
    process.exit(1);
  }
  const cases: Case[] = JSON.parse(fs.readFileSync(CASES_PATH, "utf8"));
  const provider = getProvider();
  const lines: string[] = [
    `# Backtest results`,
    ``,
    `Provider: ${provider.name()} / ${provider.version()}`,
    ``,
  ];

  for (const c of cases) {
    // eslint-disable-next-line no-console
    console.log(`Running: ${c.name}`);
    lines.push(`## ${c.name}`, ``, `- mode: ${c.referenceMode}`, `- instruction: ${c.instruction}`);
    if (c.notes) lines.push(`- notes: ${c.notes}`);

    let reference: ReferenceContext | undefined;
    if (c.referenceFile) {
      const buf = fs.readFileSync(path.join(FIXTURES_DIR, c.referenceFile));
      const { text } = await extractText(buf, c.referenceFile);
      const artifact = await buildArtifact(provider, text);
      const projection = projectArtifact(artifact, c.referenceMode);
      reference = { mode: c.referenceMode, projection };
      lines.push(``, `### Projection`, "```", projection, "```");
    }

    const mode = c.selectionText ? "edit" : "draft";
    const output = await provider.edit({
      text: c.selectionText ?? "",
      instruction: c.instruction,
      mode,
      reference,
    });
    lines.push(``, `### Output (${mode})`, "```", output, "```", ``);
  }

  fs.writeFileSync(OUT_PATH, lines.join("\n"), "utf8");
  // eslint-disable-next-line no-console
  console.log(`\nWrote ${OUT_PATH}`);
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error("Backtest failed:", err);
  process.exit(1);
});
