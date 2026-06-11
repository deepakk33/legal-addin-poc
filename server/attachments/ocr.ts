// OCR fallback for scanned PDFs (empty text layer). Heavy deps (pdfjs-dist,
// @napi-rs/canvas, tesseract.js) are loaded LAZILY via dynamic import and typed
// as `any`, so:
//   1. the rest of the backend builds/runs even if these deps fail to install;
//   2. a broken/missing OCR stack degrades to a clear error on that one card
//      rather than crashing the server.
// All are prebuilt/WASM — no native compile (keeps better-sqlite3 the only
// native module).

const OCR_UNAVAILABLE =
  "Scanned PDF — OCR is unavailable. Install the OCR deps (tesseract.js, pdfjs-dist, @napi-rs/canvas) or upload a text-based PDF/DOCX.";

// Cap pages so a huge scan can't hang ingestion. Tune as needed.
const MAX_OCR_PAGES = Number(process.env.OCR_MAX_PAGES ?? 15);

export async function ocrPdf(buf: Buffer): Promise<string> {
  let pdfjs: any;
  let canvasMod: any;
  let tesseract: any;
  try {
    // pdfjs-dist legacy build works under Node/CommonJS via dynamic import.
    pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    canvasMod = await import("@napi-rs/canvas");
    tesseract = await import("tesseract.js");
  } catch {
    throw new Error(OCR_UNAVAILABLE);
  }

  const data = new Uint8Array(buf);
  const doc = await pdfjs.getDocument({ data, disableWorker: true }).promise;
  const pageCount = Math.min(doc.numPages, MAX_OCR_PAGES);

  const worker = await tesseract.createWorker("eng");
  try {
    const pages: string[] = [];
    for (let i = 1; i <= pageCount; i++) {
      const page = await doc.getPage(i);
      const viewport = page.getViewport({ scale: 2 });
      const canvas = canvasMod.createCanvas(viewport.width, viewport.height);
      const ctx = canvas.getContext("2d");
      await page.render({ canvasContext: ctx as any, viewport }).promise;
      const png = canvas.toBuffer("image/png");
      const { data: ocr } = await worker.recognize(png);
      pages.push(ocr.text);
    }
    return pages.join("\n\n");
  } finally {
    await worker.terminate();
  }
}
