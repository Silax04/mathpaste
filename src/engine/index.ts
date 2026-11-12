/**
 * MathPaste conversion engine — shared by the task pane (method A) and the ribbon /
 * context-menu "Convert to Formula" command (method B).
 *
 *   LaTeX --(Temml)--> MathML --(OmmlConverter)--> OMML --(buildPackage)--> Flat OPC --> insertOoxml
 */
import { LatexError, latexToMathML, normalizeLatex } from "./latex";
import { OmmlConverter, OmmlOptions } from "./omml";
import { Segment, extractSegments } from "./extract";
import { ParagraphSpec, buildPackage, PackageOptions } from "./ooxml";

export { extractSegments, buildPackage, normalizeLatex, LatexError };
export type { Segment, ParagraphSpec };

export type Conversion =
  | { ok: true; latex: string; mathml: string; omml: string; display: boolean }
  | { ok: false; latex: string; error: string; kind: LatexError["kind"] };

export interface ConvertOptions extends OmmlOptions {
  display?: boolean;
}

/** Converts ONE LaTeX expression to an <m:oMath> element. Never throws. */
export function convertLatex(latex: string, opts: ConvertOptions = {}): Conversion {
  const display = opts.display ?? true;
  try {
    const mathml = latexToMathML(latex, display);
    const omml = new OmmlConverter(opts).convert(mathml);
    if (!/<m:(r|f|rad|sSub|sSup|sSubSup|d|func|nary|acc|bar|m|eqArr|limLow|limUpp|groupChr|borderBox)>/.test(omml)) {
      throw new LatexError("La conversion n'a produit aucun contenu.", "conversion");
    }
    return { ok: true, latex, mathml, omml, display };
  } catch (err) {
    if (err instanceof LatexError) return { ok: false, latex, error: err.message, kind: err.kind };
    const msg = err instanceof Error ? err.message : String(err);
    return { ok: false, latex, error: `Échec de la conversion OMML : ${msg}`, kind: "conversion" };
  }
}

export interface AnalyzedSegment extends Segment {
  conversion?: Conversion;
}

/** Splits the input into segments and converts every math segment independently. */
export function analyze(input: string, opts: OmmlOptions = {}): AnalyzedSegment[] {
  return extractSegments(input).map((s) =>
    s.kind === "math" ? { ...s, conversion: convertLatex(s.source, { ...opts, display: s.display }) } : s,
  );
}

export type LayoutMode = "lines" | "original";

/**
 * Builds the paragraphs to insert from analysed segments.
 *  - "lines":    each equation on its own line (display); text blocks kept as their own paragraphs.
 *  - "original": keeps the input's line structure, text and blank lines; equations inline in their line.
 * Failed conversions are skipped (they are reported separately).
 */
export function layoutParagraphs(input: string, segs: AnalyzedSegment[], mode: LayoutMode): ParagraphSpec[] {
  const paras: ParagraphSpec[] = [];
  if (mode === "lines") {
    for (const s of segs) {
      if (s.kind === "math") {
        if (s.conversion?.ok) paras.push({ parts: [{ kind: "math", omml: s.conversion.omml }], display: true });
      } else {
        for (const line of s.source.split("\n")) {
          const t = line.trim();
          if (t) paras.push({ parts: [{ kind: "text", text: t }] });
        }
      }
    }
    return paras;
  }

  let cur: ParagraphSpec = { parts: [] };
  const push = () => {
    // Trim whitespace-only text at both ends
    const parts = cur.parts.slice();
    while (parts.length && parts[0].kind === "text" && !(parts[0] as { text: string }).text.trim()) parts.shift();
    while (parts.length && parts[parts.length - 1].kind === "text" && !(parts[parts.length - 1] as { text: string }).text.trim()) parts.pop();
    if (parts[0]?.kind === "text") parts[0] = { kind: "text", text: (parts[0] as { text: string }).text.replace(/^\s+/, "") };
    const li = parts.length - 1;
    if (parts[li]?.kind === "text") parts[li] = { kind: "text", text: (parts[li] as { text: string }).text.replace(/\s+$/, "") };
    const maths = parts.filter((p) => p.kind === "math").length;
    paras.push({ parts, display: maths === 1 && parts.length === 1 });
    cur = { parts: [] };
  };
  for (const s of segs) {
    if (s.kind === "math") {
      if (s.conversion?.ok) cur.parts.push({ kind: "math", omml: s.conversion.omml });
    } else {
      const lines = s.source.split("\n");
      lines.forEach((line, idx) => {
        if (idx > 0) push();
        if (line) cur.parts.push({ kind: "text", text: line.replace(/\r/g, "") });
      });
    }
  }
  push();
  // Remove empty paragraphs at the start/end
  while (paras.length && paras[0].parts.length === 0) paras.shift();
  while (paras.length && paras[paras.length - 1].parts.length === 0) paras.pop();
  return paras;
}

/** One-call helper: input text -> Flat OPC package + per-expression report. */
export function buildInsertion(input: string, mode: LayoutMode, opts: OmmlOptions & PackageOptions = {}) {
  const segs = analyze(input, opts);
  const paragraphs = layoutParagraphs(input, segs, mode);
  const maths = segs.filter((s) => s.kind === "math");
  return {
    segments: segs,
    paragraphs,
    ooxml: paragraphs.length ? buildPackage(paragraphs, opts) : "",
    converted: maths.filter((s) => s.conversion?.ok).length,
    failed: maths.filter((s) => s.conversion && !s.conversion.ok),
  };
}
