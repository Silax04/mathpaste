/**
 * Character tables used by the MathML -> OMML converter.
 * Extend these tables to support additional symbols/accents.
 */

/** Large operators rendered by Word as n-ary objects (m:nary). */
export const NARY_CHARS = new Set([
  "∑", "∏", "∐", "∫", "∬", "∭", "∮", "∯", "∰", "∱", "∲", "∳", "⨌",
  "⋃", "⋂", "⋁", "⋀", "⨁", "⨂", "⨀", "⨄", "⨆",
]);

/** Integral-like operators: limits are placed as sub/superscripts. */
export const INTEGRAL_CHARS = new Set(["∫", "∬", "∭", "∮", "∯", "∰", "∱", "∲", "∳", "⨌"]);

/** MathML accent character -> OMML combining accent character (m:acc/m:chr). */
export const ACCENT_MAP: Record<string, string> = {
  "ˆ": "̂", "^": "̂", "̂": "̂",
  "ˇ": "̌", "̌": "̌",
  "˜": "̃", "~": "̃", "̃": "̃",
  "‾": "̅", "¯": "̅", "̄": "̅", "̅": "̅", "_": "̅",
  "→": "⃗", "⃗": "⃗",
  "←": "⃖",
  "↔": "⃡",
  "˙": "̇", "̇": "̇", ".": "̇",
  "¨": "̈", "̈": "̈",
  "⃛": "⃛",
  "´": "́", "́": "́",
  "`": "̀", "̀": "̀",
  "˘": "̆", "̆": "̆",
  "˚": "̊", "̊": "̊",
};

/** Stretchy horizontal characters drawn by Word as group characters (m:groupChr). */
export const GROUP_CHARS = new Set(["⏞", "⏟", "⏜", "⏝", "⎴", "⎵", "→", "←", "↔", "⟶", "⟵", "⇒", "⇐", "⇔"]);

/** Over/under bar characters (m:bar). */
export const BAR_CHARS = new Set(["‾", "¯", "_", "̲", "―", "—", "─"]);

/**
 * Operators that terminate the operand of an n-ary operator (∑, ∫ …).
 * Everything else (letters, products, scripts, fractions) stays inside the operand.
 */
export const NARY_TERMINATORS = new Set([
  "+", "−", "-", "±", "∓", "=", "≠", "≈", "≡", "<", ">", "≤", "≥", "≪", "≫", "∼", "≃", "≅", "∝",
  "→", "←", "⇒", "⇔", "⟹", "⟺", "↦", ",", ";", "∈", "∉", "⊂", "⊆", "⊃", "⊇", "|", "∣", "∥",
]);

/** Function application (invisible operator emitted by Temml after \sin, \ln…). */
export const FUNC_APPLY = "⁡";
/** Invisible operators (times, separator, plus) that must not appear in Word. */
export const INVISIBLE_OPS = new Set(["⁡", "⁢", "⁣", "⁤"]);

/**
 * Converts a MathML <mspace width="..."> to the closest Unicode space understood by Word math.
 * Returns "" for zero/negative widths (Word has no negative space).
 */
export function spaceForWidth(width: string | undefined): string {
  if (!width) return "";
  const m = /^(-?[\d.]+)\s*(em|ex|pt|px|mu)?$/.exec(width.trim());
  if (!m) return "";
  let v = parseFloat(m[1]);
  const unit = m[2] ?? "em";
  if (unit === "ex") v *= 0.43;
  else if (unit === "pt") v /= 10;
  else if (unit === "px") v /= 13.33;
  else if (unit === "mu") v /= 18;
  if (!(v > 0.01)) return "";
  if (v >= 1.9) return " ".repeat(Math.round(v));
  if (v >= 0.9) return " "; // \quad
  if (v >= 0.45) return " "; // en space
  if (v >= 0.26) return " "; // \; thick (5/18 em)
  if (v >= 0.2) return " "; // \: medium (4/18 em)
  return " "; // \, thin (3/18 em)
}

export type MathStyle = "p" | "b" | "i" | "bi";
export type MathScript =
  | "roman" | "script" | "fraktur" | "double-struck" | "sans-serif" | "monospace";

export interface AlphaInfo {
  ch: string;
  sty?: MathStyle;
  scr?: MathScript;
}

/**
 * Maps a character of the Unicode "Mathematical Alphanumeric Symbols" block (U+1D400–U+1D7FF)
 * back to a plain character + OMML style/script, which is how Word stores e.g. \mathbf{F}.
 */
export function decodeMathAlpha(cp: number): AlphaInfo | null {
  // Latin letter blocks: 52 code points each (A–Z, a–z)
  const latin: Array<[number, MathStyle | undefined, MathScript | undefined]> = [
    [0x1d400, "b", "roman"],
    [0x1d434, "i", "roman"],
    [0x1d468, "bi", "roman"],
    [0x1d49c, "p", "script"],
    [0x1d4d0, "b", "script"],
    [0x1d504, "p", "fraktur"],
    [0x1d538, "p", "double-struck"],
    [0x1d56c, "b", "fraktur"],
    [0x1d5a0, "p", "sans-serif"],
    [0x1d5d4, "b", "sans-serif"],
    [0x1d608, "i", "sans-serif"],
    [0x1d63c, "bi", "sans-serif"],
    [0x1d670, "p", "monospace"],
  ];
  for (const [base, sty, scr] of latin) {
    if (cp >= base && cp < base + 52) {
      const k = cp - base;
      const ch = String.fromCharCode(k < 26 ? 65 + k : 97 + (k - 26));
      return { ch, sty, scr };
    }
  }
  if (cp === 0x1d6a4) return { ch: "ı", sty: "i" };
  if (cp === 0x1d6a5) return { ch: "ȷ", sty: "i" };
  // Greek blocks: 58 code points each (Α..Ω, ∇, α..ω, ∂, ϵ ϑ ϰ ϕ ϱ ϖ)
  const greekChars =
    "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡϴΣΤΥΦΧΨΩ∇αβγδεζηθικλμνξοπρςστυφχψω∂ϵϑϰϕϱϖ";
  const greek: Array<[number, MathStyle]> = [
    [0x1d6a8, "b"],
    [0x1d6e2, "i"],
    [0x1d71c, "bi"],
    [0x1d756, "b"],
    [0x1d790, "bi"],
  ];
  for (const [base, sty] of greek) {
    if (cp >= base && cp < base + 58) {
      return { ch: greekChars[cp - base], sty, scr: base >= 0x1d756 ? "sans-serif" : undefined };
    }
  }
  // Digits: bold, double-struck, sans, sans-bold, monospace
  const digits: Array<[number, MathStyle, MathScript]> = [
    [0x1d7ce, "b", "roman"],
    [0x1d7d8, "p", "double-struck"],
    [0x1d7e2, "p", "sans-serif"],
    [0x1d7ec, "b", "sans-serif"],
    [0x1d7f6, "p", "monospace"],
  ];
  for (const [base, sty, scr] of digits) {
    if (cp >= base && cp < base + 10) return { ch: String.fromCharCode(48 + cp - base), sty, scr };
  }
  return null;
}
