/**
 * Word integration (Office.js). Both entry points use the shared engine:
 *  - insertText():       task pane "Insérer dans Word" (method A)
 *  - convertSelection(): ribbon / context menu / shortcut "Convert to Formula" (method B)
 */
import { buildInsertion, buildPackage, convertLatex, extractSegments, LayoutMode } from "../engine";

export interface ExpressionReport {
  latex: string;
  ok: boolean;
  error?: string;
}

export interface OperationReport {
  ok: boolean;
  message: string;
  converted: number;
  failures: ExpressionReport[];
}

export function isWordAvailable(): boolean {
  return typeof Office !== "undefined" && typeof Word !== "undefined" && !!Office.context?.document;
}

function wordError(err: unknown): string {
  if (err && typeof err === "object" && "code" in err) {
    const e = err as { code?: string; message?: string };
    if (e.code === "AccessDenied" || e.code === "ItemNotFound") return `Document Word inaccessible (${e.code}).`;
    return `Échec d'insertion dans Word : ${e.message ?? e.code}`;
  }
  return `Échec d'insertion dans Word : ${err instanceof Error ? err.message : String(err)}`;
}

// --------------------------------------------------------------------- Method A: task pane

/** Inserts the content of the task pane at the cursor position (never deletes selected text). */
export async function insertText(input: string, mode: LayoutMode): Promise<OperationReport> {
  if (!input.trim()) return { ok: false, message: "La zone de saisie est vide.", converted: 0, failures: [] };
  if (!isWordAvailable()) return { ok: false, message: "Document Word inaccessible : ouvrez MathPaste depuis Word.", converted: 0, failures: [] };

  try {
    return await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("isEmpty");
      sel.font.load("size");
      await ctx.sync();

      const size = typeof sel.font.size === "number" && sel.font.size > 0 ? Math.round(sel.font.size * 2) : undefined;
      const res = buildInsertion(input, mode, { sizeHalfPoints: size });
      const failures = res.failed.map((s) => ({ latex: s.source, ok: false, error: s.conversion && !s.conversion.ok ? s.conversion.error : "" }));

      if (res.converted === 0) {
        const none = res.segments.every((s) => s.kind !== "math");
        return {
          ok: false,
          message: none ? "Aucune expression LaTeX détectée." : "Aucune expression n'a pu être convertie.",
          converted: 0,
          failures,
        };
      }
      // Insert at the cursor; if text is selected, insert after it instead of replacing it.
      const target = sel.isEmpty ? sel : sel.getRange("End");
      target.insertOoxml(res.ooxml, "Replace");
      await ctx.sync();
      return {
        ok: failures.length === 0,
        message: `${res.converted} équation${res.converted > 1 ? "s" : ""} insérée${res.converted > 1 ? "s" : ""}` +
          (failures.length ? ` — ${failures.length} non convertie${failures.length > 1 ? "s" : ""}.` : "."),
        converted: res.converted,
        failures,
      };
    });
  } catch (err) {
    return { ok: false, message: wordError(err), converted: 0, failures: [] };
  }
}

// --------------------------------------------------------------------- Method B: selection

interface Part {
  range: Word.Range;
  text: string; // selected text of this paragraph
  paraText: string; // full paragraph text
  offset: number; // offset of this part in the combined string
}

const SEARCH_MAX = 255;

/** Escapes text for Word's (non-wildcard) search syntax. */
function escapeSearch(s: string): string {
  return s.replace(/\^/g, "^^").replace(/\t/g, "^t").replace(/[\v\n]/g, "^l").replace(/ /g, "^s");
}

/** Number of non-overlapping occurrences of `needle` that start before `pos`. */
function occurrencesBefore(hay: string, needle: string, pos: number): number {
  let n = 0;
  let i = hay.indexOf(needle);
  while (i !== -1 && i < pos) {
    n++;
    i = hay.indexOf(needle, i + needle.length);
  }
  return n;
}

/**
 * Converts the LaTeX found in the current selection into native Word equations,
 * replacing ONLY the detected expressions. Text outside the selection is never touched,
 * and an expression that cannot be converted is left unchanged.
 */
export async function convertSelection(): Promise<OperationReport> {
  if (!isWordAvailable()) return { ok: false, message: "Document Word inaccessible.", converted: 0, failures: [] };
  try {
    return await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load("isEmpty,text");
      const paras = sel.paragraphs;
      paras.load("items/text");
      await ctx.sync();

      if (sel.isEmpty || !sel.text.trim()) {
        return { ok: false, message: "Sélectionnez d'abord une expression LaTeX dans le document.", converted: 0, failures: [] };
      }

      // Selected part of each paragraph
      const raw = paras.items.map((p) => {
        const r = p.getRange("Content").intersectWithOrNullObject(sel);
        r.load("text,isNullObject");
        return { p, r };
      });
      await ctx.sync();

      const parts: Part[] = [];
      let combined = "";
      for (const { p, r } of raw) {
        if (r.isNullObject) continue;
        if (parts.length) combined += "\n";
        parts.push({ range: r, text: r.text, paraText: p.text, offset: combined.length });
        combined += r.text;
      }
      // Word uses \r (paragraph) and \v (line break); offsets are preserved (1 char each)
      const normalized = combined.replace(/[\r\v]/g, "\n");
      const segments = extractSegments(normalized).filter((s) => s.kind === "math");
      if (segments.length === 0) {
        return { ok: false, message: "Aucune expression LaTeX détectée dans la sélection.", converted: 0, failures: [] };
      }

      const partAt = (pos: number) => {
        let k = parts.length - 1;
        while (k > 0 && parts[k].offset > pos) k--;
        return k;
      };

      // 1) Locate every expression in the document (search inside the selected part only)
      type Pending = { seg: (typeof segments)[number]; locate: () => Word.Range | null; display: boolean };
      const pendings: Pending[] = [];
      const searches: Array<{ part: Part; text: string; results: Word.RangeCollection }> = [];
      const search = (part: Part, text: string) => {
        const s = { part, text, results: part.range.search(escapeSearch(text), { matchCase: true, matchWildcards: false }) };
        s.results.load("items");
        searches.push(s);
        return s;
      };
      const pick = (s: ReturnType<typeof search>, localStart: number) => {
        const k = occurrencesBefore(s.part.text, s.text, localStart);
        return s.results.items[k] ?? null;
      };

      for (const seg of segments) {
        const pi = partAt(seg.start);
        const pj = partAt(seg.end - 1);
        const P = parts[pi];
        const Q = parts[pj];
        const ls = seg.start - P.offset;
        const le = seg.end - Q.offset;
        const display =
          pi === pj && seg.source.trim() === P.paraText.replace(/[\r\v]/g, "\n").trim();

        const piece = (part: Part, s: number, e: number): (() => Word.Range | null) => {
          if (s === 0 && e === part.text.length) return () => part.range;
          const snippet = part.text.slice(s, e);
          if (snippet.length <= SEARCH_MAX) {
            const sr = search(part, snippet);
            return () => pick(sr, s);
          }
          const head = part.text.slice(s, s + 200);
          const tail = part.text.slice(e - 200, e);
          const hs = search(part, head);
          const ts = search(part, tail);
          return () => {
            const h = pick(hs, s);
            const t = pick(ts, e - 200);
            return h && t ? h.expandTo(t) : null;
          };
        };

        if (pi === pj) pendings.push({ seg, display, locate: piece(P, ls, le) });
        else {
          const a = piece(P, ls, P.text.length);
          const b = piece(Q, 0, le);
          pendings.push({ seg, display: false, locate: () => { const x = a(); const y = b(); return x && y ? x.expandTo(y) : null; } });
        }
      }
      await ctx.sync();

      const located = pendings.map((p) => ({ ...p, range: p.locate() }));
      for (const l of located) l.range?.font.load("size");
      await ctx.sync();

      // 2) Convert, then replace from the last expression to the first
      const failures: ExpressionReport[] = [];
      let converted = 0;
      for (const l of located.slice().reverse()) {
        const latex = l.seg.source;
        if (!l.range) {
          failures.push({ latex, ok: false, error: "Impossible de localiser l'expression dans le document." });
          continue;
        }
        const size = typeof l.range.font.size === "number" && l.range.font.size > 0 ? Math.round(l.range.font.size * 2) : undefined;
        const conv = convertLatex(latex, { display: l.display, sizeHalfPoints: size });
        if (!conv.ok) {
          failures.push({ latex, ok: false, error: conv.error });
          continue;
        }
        const pkg = buildPackage([{ parts: [{ kind: "math", omml: conv.omml }], display: l.display }]);
        try {
          l.range.insertOoxml(pkg, "Replace");
          await ctx.sync();
          converted++;
        } catch (err) {
          failures.push({ latex, ok: false, error: wordError(err) });
        }
      }
      failures.reverse();

      let message: string;
      if (converted && !failures.length) message = `${converted} expression${converted > 1 ? "s" : ""} convertie${converted > 1 ? "s" : ""}.`;
      else if (converted) message = `${converted} convertie${converted > 1 ? "s" : ""}, ${failures.length} laissée${failures.length > 1 ? "s" : ""} intacte${failures.length > 1 ? "s" : ""}.`;
      else message = "Aucune expression n'a pu être convertie ; le texte original est intact.";
      return { ok: failures.length === 0 && converted > 0, message, converted, failures };
    });
  } catch (err) {
    return { ok: false, message: wordError(err), converted: 0, failures: [] };
  }
}
