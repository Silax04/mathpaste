/**
 * Splits pasted / selected text into math and text segments.
 *
 * Handles:
 *  - explicit delimiters: $$…$$, \[…\], \(…\), $…$, \begin{equation|align|…}…\end{…}
 *  - Markdown code fences (```latex … ```)
 *  - raw LaTeX lines (ChatGPT/Claude style), one or several expressions, blank-line separated
 *  - expressions spanning several lines (unbalanced braces, continuation lines)
 *  - normal text mixed with an undelimited formula (heuristic)
 *
 * Offsets always refer to the ORIGINAL string so that callers can map segments back
 * to ranges in the Word document.
 */

export interface Segment {
  kind: "math" | "text";
  start: number;
  end: number;
  /** Original text of the segment (with delimiters for math). */
  source: string;
  /** Display equation (own paragraph) vs inline equation. */
  display: boolean;
}

const FUNCTION_WORDS = new Set([
  "sin", "cos", "tan", "cot", "sec", "csc", "sinh", "cosh", "tanh", "coth", "arcsin", "arccos", "arctan",
  "ln", "log", "lg", "exp", "lim", "max", "min", "det", "dim", "ker", "deg", "gcd", "sup", "inf", "arg", "mod", "Pr", "tr",
]);

const STOPWORDS_2 = new Set([
  "le", "la", "de", "et", "un", "en", "du", "au", "où", "ou", "on", "se", "ce", "sa", "si", "ne", "il", "je", "tu", "nous",
  "is", "of", "to", "in", "at", "an", "or", "if", "as", "be", "we", "it", "by", "so", "on", "no", "do",
]);

const DELIM_RE =
  /\$\$([\s\S]+?)\$\$|\\\[([\s\S]+?)\\\]|\\\(([\s\S]+?)\\\)|\\begin\{(equation\*?|align\*?|gather\*?|multline\*?|flalign\*?|displaymath|eqnarray\*?)\}[\s\S]*?\\end\{\4\}|(?<![\\$\w])\$(?![\s$])([^$\n]+?)(?<![\s\\])\$(?![\d$])/g;

export function extractSegments(input: string): Segment[] {
  // Work on a copy where code-fence marker lines are blanked (same length => same offsets)
  const work = input.replace(/^[ \t]*```[^\n]*$/gm, (m) => " ".repeat(m.length));

  const segs: Segment[] = [];
  let last = 0;
  DELIM_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = DELIM_RE.exec(work))) {
    if (m.index > last) segs.push(...analyzeUndelimited(work, input, last, m.index));
    const isInline = m[3] !== undefined || m[5] !== undefined;
    segs.push({
      kind: "math",
      start: m.index,
      end: m.index + m[0].length,
      source: input.slice(m.index, m.index + m[0].length),
      display: !isInline,
    });
    last = m.index + m[0].length;
  }
  if (last < work.length) segs.push(...analyzeUndelimited(work, input, last, work.length));

  // Inline-delimited math alone on its line is shown as display
  for (const s of segs) {
    if (s.kind === "math" && !s.display) {
      const ls = input.lastIndexOf("\n", s.start - 1) + 1;
      let le = input.indexOf("\n", s.end);
      if (le === -1) le = input.length;
      if (input.slice(ls, s.start).trim() === "" && input.slice(s.end, le).trim() === "") s.display = true;
    }
  }
  return mergeText(segs, input);
}

function mergeText(segs: Segment[], input: string): Segment[] {
  const out: Segment[] = [];
  for (const s of segs) {
    const prev = out[out.length - 1];
    if (prev && prev.kind === "text" && s.kind === "text" && prev.end === s.start) {
      prev.end = s.end;
      prev.source = input.slice(prev.start, prev.end);
    } else out.push({ ...s });
  }
  return out.filter((s) => s.end > s.start);
}

// --------------------------------------------------------------------------- undelimited text

interface Line {
  start: number;
  end: number; // exclusive, without "\n"
}

function analyzeUndelimited(work: string, input: string, a: number, b: number): Segment[] {
  const lines: Line[] = [];
  let p = a;
  while (p <= b) {
    let nl = work.indexOf("\n", p);
    if (nl === -1 || nl > b) nl = b;
    lines.push({ start: p, end: nl });
    if (nl >= b) break;
    p = nl + 1;
  }

  const out: Segment[] = [];
  const text = (s: number, e: number) => {
    if (e > s) out.push({ kind: "text", start: s, end: e, source: input.slice(s, e), display: false });
  };

  let cursor = a;
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const lt = work.slice(line.start, line.end);
    if (!lt.trim()) {
      i++;
      continue;
    }
    // Gather a chunk of consecutive non-blank lines forming one logical expression
    let j = i;
    let chunkEnd = line.end;
    while (j + 1 < lines.length) {
      const next = work.slice(lines[j + 1].start, lines[j + 1].end);
      if (!next.trim()) break;
      const cur = work.slice(line.start, chunkEnd);
      if (!hasMathIndicator(cur) && !hasMathIndicator(next)) break;
      if (isProse(cur) && !isUnbalanced(cur)) break;
      // Consecutive Markdown bullets ("- a", "- b") are separate items, not continuations
      if (BULLET_RE.test(lt) && BULLET_RE.test(next) && !isUnbalanced(cur)) break;
      if (isUnbalanced(cur) || endsWithContinuation(cur) || startsWithContinuation(next)) {
        j++;
        chunkEnd = lines[j].end;
      } else break;
    }

    const chunkStart = line.start;
    const segs = classifyChunk(work, input, chunkStart, chunkEnd);
    for (const s of segs) {
      if (s.kind === "math") {
        text(cursor, s.start);
        out.push(s);
        cursor = s.end;
      }
    }
    i = j + 1;
  }
  text(cursor, b);
  return out;
}

const BULLET_RE = /^\s*(?:[-*•]|\d+[.)])\s+\S/;

function hasMathIndicator(s: string): boolean {
  return /\\[a-zA-Z]+|\\[,;:!]|[_^]/.test(s);
}

function isUnbalanced(s: string): boolean {
  let depth = 0;
  for (let k = 0; k < s.length; k++) {
    const c = s[k];
    if (c === "\\") {
      k++;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}") depth--;
  }
  if (depth !== 0) return true;
  const count = (re: RegExp) => (s.match(re) ?? []).length;
  if (count(/\\left(?![a-zA-Z])/g) !== count(/\\right(?![a-zA-Z])/g)) return true;
  if (count(/\\begin\{/g) !== count(/\\end\{/g)) return true;
  return false;
}

function endsWithContinuation(s: string): boolean {
  return /(=|\+|-|\\cdot|\\times|\{|&|\\\\|\\left[(\[{|.]|\\left\\\{|,)\s*$/.test(s);
}

function startsWithContinuation(s: string): boolean {
  return /^\s*(=|\+|-(?!\s*[A-Za-zÀ-ÿ]{3})|\\cdot|\\times|\{|\}|&|\\right|\\\\|\)|\]|\\end\{|\\approx|\\leq|\\geq|<|>)/.test(s);
}

interface Token {
  start: number;
  end: number;
  text: string;
  cls: "prose" | "neutral" | "math";
}

/** Splits at whitespace occurring at brace depth 0. */
function tokenize(work: string, a: number, b: number): Token[] {
  const toks: Token[] = [];
  let depth = 0;
  let s = -1;
  for (let k = a; k <= b; k++) {
    const c = k < b ? work[k] : " ";
    if (c === "\\" && k + 1 < b) {
      if (s === -1) s = k;
      k++;
      continue;
    }
    if (c === "{") depth++;
    else if (c === "}") depth = Math.max(0, depth - 1);
    const ws = /\s/.test(c) && (depth === 0 || k === b);
    if (ws) {
      if (s !== -1) {
        toks.push({ start: s, end: k, text: work.slice(s, k), cls: "neutral" });
        s = -1;
      }
    } else if (s === -1) s = k;
  }
  for (const t of toks) t.cls = classifyToken(t.text);
  return toks;
}

function classifyToken(t: string): Token["cls"] {
  const bare = t.replace(/[.,;:!?»«"“”()]+$/u, "").replace(/^[«"“(]+/u, "");
  if (/^[A-Za-zÀ-ÖØ-öø-ÿ'’-]+$/u.test(bare)) {
    if (FUNCTION_WORDS.has(bare)) return "math";
    if (bare.length >= 3) return "prose";
    if (bare.length === 2 && STOPWORDS_2.has(bare.toLowerCase())) return "prose";
    return "neutral";
  }
  if (/^\d+([.,]\d+)?[.,;:]?$/.test(t)) return "neutral";
  if (/^[.,;:!?…]+$/.test(t)) return "neutral";
  return "math";
}

function isProse(s: string): boolean {
  const words = s.split(/\s+/).filter(Boolean);
  if (hasMathIndicator(s)) return false;
  return words.some((w) => classifyToken(w) === "prose");
}

function classifyChunk(work: string, input: string, a: number, b: number): Segment[] {
  // Skip leading markdown bullet / numbering ("- ", "* ", "1. ", "a) ")
  const raw = work.slice(a, b);
  const bullet = /^\s*(?:[-*•]|\d+[.)]|[a-z][)])\s+/.exec(raw);
  const startOff = bullet && raw.length > bullet[0].length ? a + bullet[0].length : a;

  const toks = tokenize(work, startOff, b);
  if (toks.length === 0) return [];
  const anyProse = toks.some((t) => t.cls === "prose");
  const anyMath = toks.some((t) => t.cls === "math");
  const content = work.slice(toks[0].start, toks[toks.length - 1].end);

  if (!anyMath) return [];
  if (!hasMathIndicator(content) && !/[=<>≤≥≈±+×÷*/−]/.test(content)) return [];

  const mk = (s: number, e: number, display: boolean): Segment => ({
    kind: "math", start: s, end: e, source: input.slice(s, e), display,
  });

  if (!anyProse) {
    return [mk(toks[0].start, trimTrailingPunct(work, toks[0].start, toks[toks.length - 1].end, false), true)];
  }

  // Mixed text: find maximal runs of non-prose tokens containing math tokens
  const segs: Segment[] = [];
  let k = 0;
  while (k < toks.length) {
    if (toks[k].cls === "prose") {
      k++;
      continue;
    }
    let e = k;
    while (e + 1 < toks.length && toks[e + 1].cls !== "prose") e++;
    const run = toks.slice(k, e + 1);
    const firstMath = run.findIndex((t) => t.cls === "math");
    if (firstMath !== -1) {
      let lastMath = run.length - 1;
      while (run[lastMath].cls !== "math") lastMath--;
      // keep at most one neutral token on each side (e.g. "x = 2", "C = …")
      let s = firstMath;
      if (s > 0 && /^[A-Za-z0-9]+$/.test(run[s - 1].text)) s--;
      let f = lastMath;
      if (f < run.length - 1 && /^[A-Za-z0-9]+[.,;:]?$/.test(run[f + 1].text)) f++;
      const start = run[s].start;
      const end = trimTrailingPunct(work, start, run[f].end, true);
      if (end > start) segs.push(mk(start, end, false));
    }
    k = e + 1;
  }
  return segs;
}

/** Removes sentence punctuation that follows a formula (", ." ";" ":"), outside braces. */
function trimTrailingPunct(work: string, start: number, end: number, aggressive: boolean): number {
  let e = end;
  while (e > start && /[.,;:]/.test(work[e - 1]) && work[e - 2] !== "\\") {
    if (!aggressive && work[e - 1] !== "." && work[e - 1] !== "," && work[e - 1] !== ";") break;
    e--;
  }
  return e;
}
