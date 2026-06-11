import * as path from "path";
import { ocrPdf } from "./ocr";

// Extract plain text from an uploaded reference doc. Supported: .docx, .pdf, .txt.
// PDFs fall back to OCR only when the text layer is empty (scanned document).
// mammoth and pdf-parse are pure-JS; pdf-parse is required lazily because its
// index module misbehaves if pulled in as a process entry point.

export interface Extracted {
  text: string;
  // True when the PDF text layer was empty and we OCR'd the rendered pages.
  ocrUsed: boolean;
}

const MIN_TEXT_LEN = 20; // below this, treat a PDF as scanned (no real text layer)

export async function extractText(buf: Buffer, filename: string): Promise<Extracted> {
  const ext = path.extname(filename).toLowerCase();

  if (ext === ".txt") {
    return { text: buf.toString("utf8").trim(), ocrUsed: false };
  }

  if (ext === ".docx") {
    const mammoth = (await import("mammoth")).default ?? (await import("mammoth"));
    const { value } = await (mammoth as any).extractRawText({ buffer: buf });
    return { text: (value || "").trim(), ocrUsed: false };
  }

  if (ext === ".pdf") {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const pdfParse = require("pdf-parse") as (b: Buffer) => Promise<{ text: string }>;
    let text = "";
    try {
      const data = await pdfParse(buf);
      text = (data.text || "").trim();
    } catch {
      text = "";
    }
    if (text.length >= MIN_TEXT_LEN) {
      return { text, ocrUsed: false };
    }
    // Empty/near-empty text layer -> scanned PDF -> OCR fallback.
    const ocrText = (await ocrPdf(buf)).trim();
    return { text: ocrText, ocrUsed: true };
  }

  throw new Error(`Unsupported file type "${ext}". Upload a .docx, .pdf, or .txt.`);
}
