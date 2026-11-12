/**
 * Minimal, dependency-free XML parser used to read the MathML produced by Temml.
 * Works identically in the browser (Word WebView2) and in Node (unit tests).
 */

export interface XNode {
  name: string; // element name without namespace prefix
  attrs: Record<string, string>;
  children: XNode[];
  text: string; // concatenated text content of direct text children
}

const ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: " ",
};

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, ent: string) => {
    if (ent[0] === "#") {
      const code = ent[1] === "x" || ent[1] === "X" ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[ent] ?? m;
  });
}

export function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export function parseXml(src: string): XNode {
  let i = 0;
  const root: XNode = { name: "#root", attrs: {}, children: [], text: "" };
  const stack: XNode[] = [root];

  while (i < src.length) {
    const lt = src.indexOf("<", i);
    const textEnd = lt === -1 ? src.length : lt;
    if (textEnd > i) {
      const t = decodeEntities(src.slice(i, textEnd));
      stack[stack.length - 1].text += t;
    }
    if (lt === -1) break;
    i = lt;
    if (src.startsWith("<!--", i)) {
      const end = src.indexOf("-->", i);
      i = end === -1 ? src.length : end + 3;
      continue;
    }
    if (src.startsWith("<?", i) || src.startsWith("<!", i)) {
      const end = src.indexOf(">", i);
      i = end === -1 ? src.length : end + 1;
      continue;
    }
    const gt = findTagEnd(src, i);
    const raw = src.slice(i + 1, gt);
    i = gt + 1;
    if (raw.startsWith("/")) {
      if (stack.length > 1) stack.pop();
      continue;
    }
    const selfClosing = raw.endsWith("/");
    const body = selfClosing ? raw.slice(0, -1) : raw;
    const m = /^\s*([^\s/>]+)/.exec(body);
    if (!m) continue;
    const qname = m[1];
    const name = qname.includes(":") ? qname.slice(qname.indexOf(":") + 1) : qname;
    const attrs: Record<string, string> = {};
    const attrRe = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let am: RegExpExecArray | null;
    const rest = body.slice(m[0].length);
    while ((am = attrRe.exec(rest))) {
      attrs[am[1]] = decodeEntities(am[3] ?? am[4] ?? "");
    }
    const node: XNode = { name, attrs, children: [], text: "" };
    stack[stack.length - 1].children.push(node);
    if (!selfClosing) stack.push(node);
  }
  return root;
}

function findTagEnd(src: string, start: number): number {
  let quote: string | null = null;
  for (let j = start + 1; j < src.length; j++) {
    const c = src[j];
    if (quote) {
      if (c === quote) quote = null;
    } else if (c === '"' || c === "'") {
      quote = c;
    } else if (c === ">") {
      return j;
    }
  }
  return src.length - 1;
}
