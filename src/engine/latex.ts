import temml from "temml";
import { CUSTOM_MACROS } from "./macros";

export class LatexError extends Error {
  constructor(
    message: string,
    public readonly kind: "empty" | "invalid" | "unsupported" | "conversion",
    public readonly detail?: string,
  ) {
    super(message);
    this.name = "LatexError";
  }
}

const ENV_WRAPPERS = /^\\begin\{(equation|equation\*|displaymath|math|gather\*?|align\*?|multline\*?|flalign\*?|eqnarray\*?)\}([\s\S]*)\\end\{\1\}$/;

/**
 * Normalises raw LaTeX copied from ChatGPT / Claude / Overleaf:
 * strips math delimiters, numbering commands, outer equation environments, trailing punctuation…
 */
export function normalizeLatex(input: string): string {
  let s = input.replace(/\r\n?/g, "\n").replace(/\u000b/g, "\n").trim();

  // Strip delimiters ($$…$$, \[…\], \(…\), $…$)
  for (let k = 0; k < 3; k++) {
    const before = s;
    if (s.startsWith("$$") && s.endsWith("$$") && s.length >= 4) s = s.slice(2, -2).trim();
    else if (s.startsWith("\\[") && s.endsWith("\\]")) s = s.slice(2, -2).trim();
    else if (s.startsWith("\\(") && s.endsWith("\\)")) s = s.slice(2, -2).trim();
    else if (s.startsWith("$") && s.endsWith("$") && s.length >= 2) s = s.slice(1, -1).trim();
    if (s === before) break;
  }

  // Outer numbered/display environments -> content (align-like become aligned)
  const env = ENV_WRAPPERS.exec(s);
  if (env) {
    const name = env[1].replace("*", "");
    const body = env[2].trim();
    if (/^(align|flalign|eqnarray)$/.test(name)) s = `\\begin{aligned}${body}\\end{aligned}`;
    else if (name === "gather") s = `\\begin{gathered}${body}\\end{gathered}`;
    else if (name === "multline") s = `\\begin{gathered}${body}\\end{gathered}`;
    else s = body;
  }

  // Remove labels / tags which have no meaning inside a Word equation
  s = s.replace(/\\label\{[^{}]*\}/g, "").replace(/\\tag\*?\{[^{}]*\}/g, "");

  // Remove LaTeX comments (unescaped %)
  s = s.replace(/(^|[^\\])%[^\n]*/g, "$1");

  // Trailing "\\" at the very end (common when copying a line out of an align block)
  s = s.replace(/\\\\\s*$/, "");

  // Lines are only whitespace for LaTeX
  s = s.replace(/\s*\n\s*/g, " ").trim();

  // A trailing sentence punctuation outside braces (", ." after the formula) is kept by LaTeX,
  // so we keep it too: no change here on purpose.
  return s;
}

/**
 * Converts LaTeX to MathML using Temml. Throws LatexError with a user-friendly French message.
 */
export function latexToMathML(latex: string, display: boolean): string {
  const src = normalizeLatex(latex);
  if (!src) throw new LatexError("Expression vide.", "empty");
  try {
    return temml.renderToString(src, {
      displayMode: display,
      xml: true,
      throwOnError: true,
      trust: false,
      strict: false,
      macros: { ...CUSTOM_MACROS },
    } as Parameters<typeof temml.renderToString>[1]);
  } catch (err) {
    throw translateTemmlError(err);
  }
}

function translateTemmlError(err: unknown): LatexError {
  const raw = err instanceof Error ? err.message : String(err);
  const msg = raw.replace(/^Temml parse error:\s*/i, "").split("\n")[0].trim();
  let m = /Unsupported function name: (\\\S+)/.exec(msg) ?? /Undefined control sequence: (\\\S+)/.exec(msg);
  if (m) return new LatexError(`Commande non supportée : ${m[1]}`, "unsupported", msg);
  m = /No such environment: (\S+)/.exec(msg);
  if (m) return new LatexError(`Environnement non supporté : ${m[1]}`, "unsupported", msg);
  if (/expected '\}'|Expected '\}'|Unexpected end of input|Extra \}|Unexpected character: '\}'/i.test(msg))
    return new LatexError("Accolades { } non équilibrées.", "invalid", msg);
  if (/Missing \\right|Missing \\left|Expected '\\right'|\\right/i.test(msg))
    return new LatexError("\\left et \\right ne sont pas appariés.", "invalid", msg);
  if (/Double superscript/i.test(msg)) return new LatexError("Double exposant (x^a^b) : ajoutez des accolades.", "invalid", msg);
  if (/Double subscript/i.test(msg)) return new LatexError("Double indice (x_a_b) : ajoutez des accolades.", "invalid", msg);
  if (/Expected group|Expected .* argument|Undefined argument/i.test(msg))
    return new LatexError("Argument manquant après une commande.", "invalid", msg);
  if (/Cannot read properties|undefined|null/i.test(msg))
    return new LatexError("Expression LaTeX invalide (structure incorrecte, ex. ^ ou _ sans argument).", "invalid", msg);
  return new LatexError(`Expression LaTeX invalide : ${msg}`, "invalid", msg);
}
