/**
 * Builds the Flat OPC package expected by Word's `insertOoxml` (Office.js) / `InsertXML` (COM).
 */
import { escapeXml } from "./xml";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/** One piece of a paragraph: plain text or an OMML equation. */
export type InlinePart = { kind: "text"; text: string } | { kind: "math"; omml: string };

/** A paragraph to insert. `display` = the paragraph holds a single display equation (m:oMathPara). */
export interface ParagraphSpec {
  parts: InlinePart[];
  display?: boolean;
}

export interface PackageOptions {
  sizeHalfPoints?: number;
}

function textRun(text: string, opts: PackageOptions): string {
  const rPr = opts.sizeHalfPoints ? `<w:rPr><w:sz w:val="${opts.sizeHalfPoints}"/><w:szCs w:val="${opts.sizeHalfPoints}"/></w:rPr>` : "";
  // Tabs are separate elements in WordprocessingML
  return text
    .split("\t")
    .map((chunk, i) => (i > 0 ? `<w:r>${rPr}<w:tab/></w:r>` : "") + (chunk ? `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(chunk)}</w:t></w:r>` : ""))
    .join("");
}

export function paragraphXml(p: ParagraphSpec, opts: PackageOptions = {}): string {
  const maths = p.parts.filter((x) => x.kind === "math");
  if (p.display && maths.length === 1 && p.parts.length === 1) {
    return `<w:p><m:oMathPara>${(maths[0] as { omml: string }).omml}</m:oMathPara></w:p>`;
  }
  const body = p.parts.map((x) => (x.kind === "math" ? x.omml : textRun(x.text, opts))).join("");
  return `<w:p>${body}</w:p>`;
}

/**
 * Returns a Flat OPC (pkg:package) string. When `inline` is true the content is a single paragraph
 * whose runs Word merges into the target paragraph (used to replace a selection in place).
 */
export function buildPackage(paragraphs: ParagraphSpec[], opts: PackageOptions = {}): string {
  const body = paragraphs.map((p) => paragraphXml(p, opts)).join("");
  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData><w:document xmlns:w="${W_NS}" xmlns:m="${M_NS}"><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  );
}
