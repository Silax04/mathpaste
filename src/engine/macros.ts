/**
 * Extension point #1 — LaTeX-level macros.
 *
 * Any command added here is expanded by Temml before conversion, so it automatically
 * benefits from the full LaTeX -> MathML -> OMML pipeline. Use "#1", "#2"… for arguments.
 *
 * Example: "\\vect": "\\mathbf{#1}"
 */
export const CUSTOM_MACROS: Record<string, string> = {
  // Common shortcuts found in ChatGPT / Claude / Overleaf output
  "\\R": "\\mathbb{R}",
  "\\N": "\\mathbb{N}",
  "\\Z": "\\mathbb{Z}",
  "\\Q": "\\mathbb{Q}",
  "\\C": "\\mathbb{C}",
  "\\dd": "\\mathrm{d}",
  "\\e": "\\mathrm{e}",
  "\\abs": "\\left|#1\\right|",
  "\\norm": "\\left\\|#1\\right\\|",
  "\\dfrac": "\\frac{#1}{#2}",
  "\\tfrac": "\\frac{#1}{#2}",
  "\\cfrac": "\\frac{#1}{#2}",
  "\\displaystyle": "",
  "\\textstyle": "",
  "\\scriptstyle": "",
  "\\nonumber": "",
  "\\notag": "",
  "\\limits": "",
  "\\nolimits": "",
  "\\eqref": "(\\text{#1})",
  "\\ref": "\\text{#1}",
  "\\si": "\\mathrm{#1}",
  "\\SI": "#1\\,\\mathrm{#2}",
  "\\unit": "\\mathrm{#1}",
  "\\qty": "#1\\,\\mathrm{#2}",
};
