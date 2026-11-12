/**
 * MathML (as produced by Temml) -> OMML (Office Math Markup Language).
 *
 * Produces the same structures Word's own equation editor builds:
 *   fractions m:f, radicals m:rad, scripts m:sSub/m:sSup/m:sSubSup, delimiters m:d,
 *   functions m:func, n-ary operators m:nary, accents m:acc, bars m:bar, group characters,
 *   limits m:limLow/m:limUpp, matrices m:m and equation arrays m:eqArr.
 *
 * Extension point #2: add a handler to ELEMENT_HANDLERS for a new MathML element.
 */
import { XNode, escapeXml, parseXml } from "./xml";
import {
  ACCENT_MAP, BAR_CHARS, FUNC_APPLY, GROUP_CHARS, INTEGRAL_CHARS, INVISIBLE_OPS, NARY_CHARS,
  NARY_TERMINATORS, MathScript, MathStyle, decodeMathAlpha, spaceForWidth,
} from "./charmaps";

export interface OmmlOptions {
  /** Font size in half-points (w:sz). Undefined = inherit from paragraph. */
  sizeHalfPoints?: number;
  /** Math font (default Cambria Math, Word's default math font). */
  mathFont?: string;
}

interface Ctx {
  sty?: MathStyle;
  scr?: MathScript;
  color?: string;
  nor?: boolean;
}

type Item = { kind: "run"; text: string; ctx: Ctx; aln?: boolean } | { kind: "xml"; xml: string };

export class OmmlConverter {
  private readonly font: string;
  private readonly sz?: number;

  constructor(opts: OmmlOptions = {}) {
    this.font = opts.mathFont ?? "Cambria Math";
    this.sz = opts.sizeHalfPoints;
  }

  /** Converts a full <math> MathML string into an <m:oMath> element (string). */
  convert(mathml: string): string {
    const doc = parseXml(mathml);
    const math = doc.children.find((c) => c.name === "math");
    if (!math) throw new Error("MathML invalide : élément <math> introuvable.");
    const body = this.seq(math.children, {});
    return `<m:oMath>${body}</m:oMath>`;
  }

  // ---------------------------------------------------------------- run properties

  private wRPr(ctx: Ctx, forCtrl = false): string {
    let x = "";
    if (!ctx.nor) x += `<w:rFonts w:ascii="${this.font}" w:hAnsi="${this.font}"/>`;
    if (forCtrl) x += "<w:i/>";
    if (ctx.color) x += `<w:color w:val="${ctx.color}"/>`;
    if (this.sz) x += `<w:sz w:val="${this.sz}"/><w:szCs w:val="${this.sz}"/>`;
    return `<w:rPr>${x}</w:rPr>`;
  }

  private ctrl(ctx: Ctx = {}): string {
    return `<m:ctrlPr>${this.wRPr({ color: ctx.color }, true)}</m:ctrlPr>`;
  }

  private mRPr(ctx: Ctx, aln?: boolean): string {
    let x = "";
    if (ctx.nor) x += "<m:nor/>";
    else {
      if (ctx.scr && ctx.scr !== "roman") x += `<m:scr m:val="${ctx.scr}"/>`;
      if (ctx.sty) x += `<m:sty m:val="${ctx.sty}"/>`;
    }
    if (aln) x += "<m:aln/>";
    return x ? `<m:rPr>${x}</m:rPr>` : "";
  }

  private run(text: string, ctx: Ctx, aln?: boolean): string {
    return `<m:r>${this.mRPr(ctx, aln)}${this.wRPr(ctx)}<m:t xml:space="preserve">${escapeXml(text)}</m:t></m:r>`;
  }

  /** Serialises items, merging adjacent runs that share the same formatting. */
  private emit(items: Item[]): string {
    let out = "";
    let pending: { text: string; ctx: Ctx; aln?: boolean } | null = null;
    const key = (c: Ctx) => `${c.sty ?? ""}|${c.scr ?? ""}|${c.color ?? ""}|${c.nor ? 1 : 0}`;
    const flush = () => {
      if (pending) out += this.run(pending.text, pending.ctx, pending.aln);
      pending = null;
    };
    for (const it of items) {
      if (it.kind === "xml") {
        flush();
        out += it.xml;
      } else if (pending && !it.aln && key(pending.ctx) === key(it.ctx)) {
        pending.text += it.text;
      } else {
        flush();
        pending = { text: it.text, ctx: it.ctx, aln: it.aln };
      }
    }
    flush();
    return out;
  }

  // ---------------------------------------------------------------- sequences

  /** Converts a list of sibling MathML nodes to OMML content. */
  seq(nodes: XNode[], ctx: Ctx): string {
    return this.emit(this.seqItems(nodes, ctx));
  }

  private seqItems(nodes: XNode[], ctx: Ctx): Item[] {
    const list = this.flatten(nodes, ctx);
    const items: Item[] = [];
    for (let i = 0; i < list.length; i++) {
      const { n, ctx: c } = list[i];

      if (isMo(n) && INVISIBLE_OPS.has(n.text.trim())) continue;

      // --- function application: name ⁡ argument  ->  m:func
      const applyIdx = nextNonSpace(list, i + 1);
      if (applyIdx < list.length && isMo(list[applyIdx].n) && list[applyIdx].n.text.trim() === FUNC_APPLY) {
        const fName = this.emit(this.nodeItems(n, c, true));
        let j = nextNonSpace(list, applyIdx + 1);
        const argNodes: Array<{ n: XNode; ctx: Ctx }> = [];
        if (j < list.length && isFence(list[j].n)) {
          argNodes.push(list[j]);
          j++;
        } else {
          while (j < list.length && isArgAtom(list[j].n) && !isFunctionHead(list, j)) {
            argNodes.push(list[j]);
            j++;
          }
        }
        if (argNodes.length === 0) {
          // No argument (e.g. "\max" alone, "\arg\min"): keep the upright name only.
          items.push(...this.nodeItems(n, c, true));
          if (j < list.length && !isMo(list[j].n)) items.push({ kind: "run", text: " ", ctx: {} });
          i = applyIdx;
          continue;
        }
        const e = this.emit(argNodes.flatMap((a) => this.nodeItems(a.n, a.ctx)));
        items.push({
          kind: "xml",
          xml: `<m:func><m:funcPr>${this.ctrl(c)}</m:funcPr><m:fName>${fName}</m:fName><m:e>${e}</m:e></m:func>`,
        });
        i = j - 1;
        continue;
      }

      // --- n-ary operators: ∑ ∫ ∏ … with optional limits -> m:nary
      const nary = naryParts(n);
      if (nary) {
        let j = i + 1;
        const body: Array<{ n: XNode; ctx: Ctx }> = [];
        while (j < list.length && !(isMo(list[j].n) && NARY_TERMINATORS.has(list[j].n.text.trim()))) {
          body.push(list[j]);
          j++;
        }
        items.push({ kind: "xml", xml: this.nary(nary, body, c) });
        i = j - 1;
        continue;
      }

      items.push(...this.nodeItems(n, c));
    }
    return items;
  }

  /**
   * Flattens grouping-only elements (mrow, mstyle, mpadded…) into their parent sequence,
   * which is how Word represents juxtaposed terms. Fences, boxes, etc. are kept as nodes.
   */
  private flatten(nodes: XNode[], ctx: Ctx): Array<{ n: XNode; ctx: Ctx }> {
    const out: Array<{ n: XNode; ctx: Ctx }> = [];
    for (const n of nodes) {
      if (n.name === "semantics") {
        const first = n.children.find((c) => c.name !== "annotation" && c.name !== "annotation-xml");
        if (first) out.push(...this.flatten([first], ctx));
        continue;
      }
      if (n.name === "annotation" || n.name === "annotation-xml") continue;
      if ((n.name === "mrow" || n.name === "mstyle" || n.name === "mpadded" || n.name === "mtd") && !isFence(n) && !isBoxed(n)) {
        let c = ctx;
        const color = n.attrs.mathcolor ?? colorFromStyle(n.attrs.style);
        if (color) c = { ...c, color: normalizeColor(color) ?? c.color };
        let kids = n.children;
        // Function heads generated by Temml: [space] name ⁡ [space] -> drop automatic spaces
        if (kids.some((k) => isMo(k) && k.text.trim() === FUNC_APPLY)) kids = kids.filter((k) => k.name !== "mspace");
        out.push(...this.flatten(kids, c));
        continue;
      }
      out.push({ n, ctx });
    }
    return out;
  }

  // ---------------------------------------------------------------- single nodes

  private nodeItems(n: XNode, ctx: Ctx, asFuncName = false): Item[] {
    const handler = ELEMENT_HANDLERS[n.name];
    if (handler) return handler.call(this, n, ctx, asFuncName);
    // Unknown element: convert its children so no content is lost.
    return this.seqItems(n.children, ctx);
  }

  /** Converts a node as the content of an OMML argument (m:e, m:num, m:sub…). */
  arg(n: XNode | undefined, ctx: Ctx): string {
    if (!n) return "";
    return this.seq([n], ctx);
  }

  token(n: XNode, ctx: Ctx, kind: "mi" | "mn" | "mo" | "mtext"): Item[] {
    let text = n.text;
    if (kind === "mtext") {
      const tc = colorFromStyle(n.attrs.style);
      return text ? [{ kind: "run", text, ctx: { ...ctx, nor: true, ...(tc && normalizeColor(tc) ? { color: normalizeColor(tc) } : {}) } }] : [];
    }
    text = text.trim();
    if (!text) return [];
    if (kind === "mo") text = [...text].filter((ch) => !INVISIBLE_OPS.has(ch)).join("");
    if (!text) return [];

    const rawColor = n.attrs.mathcolor ?? colorFromStyle(n.attrs.style);
    const color = rawColor ? normalizeColor(rawColor) : undefined;
    const base: Ctx = color ? { ...ctx, color } : { ...ctx };
    const variant = n.attrs.mathvariant;

    // Decode Unicode mathematical alphanumerics (𝐅 -> F + bold), as Word stores them.
    const items: Item[] = [];
    const chars = [...text];
    let plain = "";
    const flushPlain = () => {
      if (plain) items.push({ kind: "run", text: plain, ctx: applyVariant(base, variant, kind, chars.length) });
      plain = "";
    };
    for (const ch of chars) {
      const info = decodeMathAlpha(ch.codePointAt(0)!);
      if (info) {
        flushPlain();
        const c: Ctx = { ...base, sty: info.sty, scr: info.scr };
        items.push({ kind: "run", text: info.ch, ctx: c });
      } else plain += ch;
    }
    flushPlain();
    return items;
  }

  frac(n: XNode, ctx: Ctx): Item[] {
    const [num, den] = n.children;
    const lt = (n.attrs.linethickness ?? "").trim();
    let type = "";
    if (/^0(\.0+)?(px|pt|em|ex|mu)?$/.test(lt)) type = `<m:type m:val="noBar"/>`;
    else if (n.attrs.bevelled === "true") type = `<m:type m:val="skw"/>`;
    return [{
      kind: "xml",
      xml: `<m:f><m:fPr>${type}${this.ctrl(ctx)}</m:fPr><m:num>${this.arg(num, ctx)}</m:num><m:den>${this.arg(den, ctx)}</m:den></m:f>`,
    }];
  }

  rad(base: XNode[], deg: XNode | undefined, ctx: Ctx): Item[] {
    const baseNodes = base.filter((b) => !(b.name === "mspace" && !spaceForWidth(b.attrs.width)));
    const degXml = deg ? this.arg(deg, ctx) : "";
    const pr = deg ? `<m:radPr>${this.ctrl(ctx)}</m:radPr>` : `<m:radPr><m:degHide m:val="1"/>${this.ctrl(ctx)}</m:radPr>`;
    return [{ kind: "xml", xml: `<m:rad>${pr}<m:deg>${degXml}</m:deg><m:e>${this.seq(baseNodes, ctx)}</m:e></m:rad>` }];
  }

  scripts(n: XNode, ctx: Ctx, asFuncName: boolean): Item[] {
    const [b, s1, s2] = n.children;
    const e = asFuncName ? this.emit(this.nodeItems(b, ctx, true)) : this.arg(b, ctx);
    const c = this.ctrl(ctx);
    if (n.name === "msub")
      return [{ kind: "xml", xml: `<m:sSub><m:sSubPr>${c}</m:sSubPr><m:e>${e}</m:e><m:sub>${this.arg(s1, ctx)}</m:sub></m:sSub>` }];
    if (n.name === "msup")
      return [{ kind: "xml", xml: `<m:sSup><m:sSupPr>${c}</m:sSupPr><m:e>${e}</m:e><m:sup>${this.arg(s1, ctx)}</m:sup></m:sSup>` }];
    return [{
      kind: "xml",
      xml: `<m:sSubSup><m:sSubSupPr>${c}</m:sSubSupPr><m:e>${e}</m:e><m:sub>${this.arg(s1, ctx)}</m:sub><m:sup>${this.arg(s2, ctx)}</m:sup></m:sSubSup>`,
    }];
  }

  underOver(n: XNode, ctx: Ctx, asFuncName: boolean): Item[] {
    const [b, s1, s2] = n.children;
    const e = asFuncName ? this.emit(this.nodeItems(b, ctx, true)) : this.arg(b, ctx);
    const c = this.ctrl(ctx);
    const opChar = (x: XNode | undefined) => (x && isMo(x) ? x.text.trim() : x && x.name === "mrow" && x.children.length === 1 && isMo(x.children[0]) ? x.children[0].text.trim() : "");

    if (n.name === "mover") {
      const ch = opChar(s1);
      const stretchy = s1?.attrs.stretchy === "true";
      if (ch === "⏞" || ch === "⏜" || ch === "⎴" || (stretchy && GROUP_CHARS.has(ch))) {
        return [{ kind: "xml", xml: `<m:groupChr><m:groupChrPr><m:chr m:val="${escapeXml(ch)}"/><m:pos m:val="top"/><m:vertJc m:val="bot"/>${c}</m:groupChrPr><m:e>${e}</m:e></m:groupChr>` }];
      }
      if (stretchy && BAR_CHARS.has(ch)) {
        return [{ kind: "xml", xml: `<m:bar><m:barPr><m:pos m:val="top"/>${c}</m:barPr><m:e>${e}</m:e></m:bar>` }];
      }
      if (ch && ACCENT_MAP[ch] && (n.attrs.accent === "true" || !stretchy)) {
        const acc = ACCENT_MAP[ch];
        const chr = acc === "̂" ? "" : `<m:chr m:val="${acc}"/>`;
        return [{ kind: "xml", xml: `<m:acc><m:accPr>${chr}${c}</m:accPr><m:e>${e}</m:e></m:acc>` }];
      }
      return [{ kind: "xml", xml: `<m:limUpp><m:limUppPr>${c}</m:limUppPr><m:e>${e}</m:e><m:lim>${this.arg(s1, ctx)}</m:lim></m:limUpp>` }];
    }
    if (n.name === "munder") {
      const ch = opChar(s1);
      const stretchy = s1?.attrs.stretchy === "true";
      if (ch === "⏟" || ch === "⏝" || ch === "⎵" || (stretchy && GROUP_CHARS.has(ch))) {
        const chr = ch === "⏟" ? "" : `<m:chr m:val="${escapeXml(ch)}"/>`;
        return [{ kind: "xml", xml: `<m:groupChr><m:groupChrPr>${chr}${c}</m:groupChrPr><m:e>${e}</m:e></m:groupChr>` }];
      }
      if (BAR_CHARS.has(ch) && (stretchy || n.attrs.accentunder === "true")) {
        return [{ kind: "xml", xml: `<m:bar><m:barPr>${c}</m:barPr><m:e>${e}</m:e></m:bar>` }];
      }
      return [{ kind: "xml", xml: `<m:limLow><m:limLowPr>${c}</m:limLowPr><m:e>${e}</m:e><m:lim>${this.arg(s1, ctx)}</m:lim></m:limLow>` }];
    }
    // munderover on a non n-ary base
    const low = `<m:limLow><m:limLowPr>${c}</m:limLowPr><m:e>${e}</m:e><m:lim>${this.arg(s1, ctx)}</m:lim></m:limLow>`;
    return [{ kind: "xml", xml: `<m:limUpp><m:limUppPr>${c}</m:limUppPr><m:e>${low}</m:e><m:lim>${this.arg(s2, ctx)}</m:lim></m:limUpp>` }];
  }

  nary(p: NaryParts, body: Array<{ n: XNode; ctx: Ctx }>, ctx: Ctx): string {
    const isInt = INTEGRAL_CHARS.has(p.chr);
    let pr = "";
    if (p.chr !== "∫") pr += `<m:chr m:val="${escapeXml(p.chr)}"/>`;
    pr += `<m:limLoc m:val="${isInt || !p.underOver ? "subSup" : "undOvr"}"/>`;
    if (!p.sub) pr += `<m:subHide m:val="1"/>`;
    if (!p.sup) pr += `<m:supHide m:val="1"/>`;
    const e = this.emit(body.flatMap((b) => this.nodeItems(b.n, b.ctx)));
    return `<m:nary><m:naryPr>${pr}${this.ctrl(ctx)}</m:naryPr><m:sub>${this.arg(p.sub, ctx)}</m:sub><m:sup>${this.arg(p.sup, ctx)}</m:sup><m:e>${e}</m:e></m:nary>`;
  }

  fence(n: XNode, ctx: Ctx): Item[] {
    const kids = n.children;
    const open = kids[0];
    const close = kids[kids.length - 1];
    const beg = open.text.trim();
    const end = close.text.trim();
    // Split on \middle separators
    const groups: XNode[][] = [[]];
    let sep = "";
    for (const k of kids.slice(1, -1)) {
      if (isMo(k) && k.attrs.stretchy === "true" && k.attrs.form === "infix" && k.attrs.fence !== "true") {
        sep = k.text.trim();
        groups.push([]);
      } else groups[groups.length - 1].push(k);
    }
    let pr = "";
    if (beg !== "(") pr += `<m:begChr m:val="${escapeXml(beg)}"/>`;
    if (sep) pr += `<m:sepChr m:val="${escapeXml(sep)}"/>`;
    if (end !== ")") pr += `<m:endChr m:val="${escapeXml(end)}"/>`;

    // Cases-like: "{" + table + nothing  -> left-aligned matrix (Word "cases" layout)
    const inner = groups.map((g) => {
      if (g.length === 1 && g[0].name === "mtable") return this.table(g[0], ctx, beg === "{" && end === "" ? "cases" : "matrix");
      return this.seq(g, ctx);
    });
    return [{ kind: "xml", xml: `<m:d><m:dPr>${pr}${this.ctrl(ctx)}</m:dPr>${inner.map((x) => `<m:e>${x}</m:e>`).join("")}</m:d>` }];
  }

  table(n: XNode, ctx: Ctx, mode?: "matrix" | "cases"): string {
    const rows = n.children.filter((r) => r.name === "mtr" || r.name === "mlabeledtr");
    const cells = rows.map((r) => (r.name === "mlabeledtr" ? r.children.slice(1) : r.children).filter((c) => c.name === "mtd"));
    const nCols = Math.max(1, ...cells.map((r) => r.length));
    const firstAligns = cells[0]?.map((c) => (c.attrs.class ?? "").includes("tml-right") ? "r" : (c.attrs.class ?? "").includes("tml-left") ? "l" : "c") ?? [];
    const aligned = mode !== "matrix" && mode !== "cases" && (nCols === 1 || firstAligns[0] === "r" || (n.attrs.class ?? "").includes("tml-jot") || n.attrs.displaystyle === "true");

    if (aligned) {
      // Equation array (aligned / gathered / split): Word stores alignment points as "&".
      const es = cells.map((rowCells) => {
        const items: Item[] = [];
        rowCells.forEach((cell, idx) => {
          if (idx > 0) items.push({ kind: "run", text: "&", ctx: {} });
          items.push(...this.seqItems(cell.children, ctx));
        });
        return `<m:e>${this.emit(items)}</m:e>`;
      });
      return `<m:eqArr><m:eqArrPr>${this.ctrl(ctx)}</m:eqArrPr>${es.join("")}</m:eqArr>`;
    }
    const jc = mode === "cases" ? "left" : "center";
    const mcs = `<m:mcs><m:mc><m:mcPr><m:count m:val="${nCols}"/><m:mcJc m:val="${jc}"/></m:mcPr></m:mc></m:mcs>`;
    const mrs = cells
      .map((rowCells) => {
        const es: string[] = [];
        for (let k = 0; k < nCols; k++) es.push(`<m:e>${rowCells[k] ? this.seq(rowCells[k].children, ctx) : ""}</m:e>`);
        return `<m:mr>${es.join("")}</m:mr>`;
      })
      .join("");
    return `<m:m><m:mPr>${mcs}${mode === "cases" ? `<m:cGp m:val="12"/>` : ""}${this.ctrl(ctx)}</m:mPr>${mrs}</m:m>`;
  }

  enclose(n: XNode, ctx: Ctx): Item[] {
    const notation = (n.attrs.notation ?? "box").split(/\s+/);
    const c = this.ctrl(ctx);
    const e = this.seq(n.children, ctx);
    if (notation.includes("top") && notation.length === 1)
      return [{ kind: "xml", xml: `<m:bar><m:barPr><m:pos m:val="top"/>${c}</m:barPr><m:e>${e}</m:e></m:bar>` }];
    if (notation.includes("bottom") && notation.length === 1)
      return [{ kind: "xml", xml: `<m:bar><m:barPr>${c}</m:barPr><m:e>${e}</m:e></m:bar>` }];
    const strike = notation.some((x) => /strike/.test(x));
    let pr = "";
    if (strike && !notation.includes("box")) pr += `<m:hideTop m:val="1"/><m:hideBot m:val="1"/><m:hideLeft m:val="1"/><m:hideRight m:val="1"/>`;
    if (notation.includes("horizontalstrike")) pr += `<m:strikeH m:val="1"/>`;
    if (notation.includes("verticalstrike")) pr += `<m:strikeV m:val="1"/>`;
    if (notation.includes("updiagonalstrike")) pr += `<m:strikeBLTR m:val="1"/>`;
    if (notation.includes("downdiagonalstrike")) pr += `<m:strikeTLBR m:val="1"/>`;
    return [{ kind: "xml", xml: `<m:borderBox><m:borderBoxPr>${pr}${c}</m:borderBoxPr><m:e>${e}</m:e></m:borderBox>` }];
  }
}

// ------------------------------------------------------------------ element handlers

type Handler = (this: OmmlConverter, n: XNode, ctx: Ctx, asFuncName: boolean) => Item[];

export const ELEMENT_HANDLERS: Record<string, Handler> = {
  mi(n, ctx) { return this.token(n, ctx, "mi"); },
  mn(n, ctx) { return this.token(n, ctx, "mn"); },
  mo(n, ctx) { return this.token(n, ctx, "mo"); },
  mtext(n, ctx) { return this.token(n, ctx, "mtext"); },
  ms(n, ctx) { return this.token(n, ctx, "mtext"); },
  mspace(n, ctx) {
    const s = spaceForWidth(n.attrs.width);
    return s ? [{ kind: "run", text: s, ctx: { color: ctx.color } }] : [];
  },
  mfrac(n, ctx) { return this.frac(n, ctx); },
  msqrt(n, ctx) { return this.rad(n.children, undefined, ctx); },
  mroot(n, ctx) { return this.rad([n.children[0]], n.children[1], ctx); },
  msub(n, ctx, f) { return this.scripts(n, ctx, f); },
  msup(n, ctx, f) { return this.scripts(n, ctx, f); },
  msubsup(n, ctx, f) { return this.scripts(n, ctx, f); },
  mover(n, ctx, f) { return this.underOver(n, ctx, f); },
  munder(n, ctx, f) { return this.underOver(n, ctx, f); },
  munderover(n, ctx, f) { return this.underOver(n, ctx, f); },
  menclose(n, ctx) { return this.enclose(n, ctx); },
  mtable(n, ctx) { return [{ kind: "xml", xml: this.table(n, ctx) }]; },
  mphantom(n, ctx) {
    return [{ kind: "xml", xml: `<m:phant><m:phantPr><m:show m:val="0"/>${""}</m:phantPr><m:e>${this.seq(n.children, ctx)}</m:e></m:phant>` }];
  },
  mrow(n, ctx) {
    if (isFence(n)) return this.fence(n, ctx);
    if (isBoxed(n)) return this.enclose({ ...n, attrs: { notation: "box" } }, ctx);
    return [{ kind: "xml", xml: this.seq(n.children, ctx) }];
  },
  mstyle(n, ctx) { return [{ kind: "xml", xml: this.seq(n.children, ctx) }]; },
  mpadded(n, ctx) { return [{ kind: "xml", xml: this.seq(n.children, ctx) }]; },
  merror(n) { throw new Error(`Erreur de rendu : ${n.text || "expression invalide"}`); },
};

// ------------------------------------------------------------------ helpers

interface NaryParts {
  chr: string;
  sub?: XNode;
  sup?: XNode;
  underOver: boolean;
}

function isMo(n: XNode): boolean {
  return n.name === "mo";
}

function isFence(n: XNode): boolean {
  if (n.name !== "mrow" || n.children.length < 2) return false;
  const a = n.children[0];
  const b = n.children[n.children.length - 1];
  return isMo(a) && isMo(b) && a.attrs.fence === "true" && b.attrs.fence === "true";
}

function isBoxed(n: XNode): boolean {
  return n.name === "mrow" && /border\s*:/.test(n.attrs.style ?? "");
}

function nextNonSpace(list: Array<{ n: XNode }>, from: number): number {
  let j = from;
  while (j < list.length && list[j].n.name === "mspace") j++;
  return j;
}

function isFunctionHead(list: Array<{ n: XNode }>, j: number): boolean {
  const k = nextNonSpace(list, j + 1);
  return k < list.length && isMo(list[k].n) && list[k].n.text.trim() === FUNC_APPLY;
}

/** Elements that can belong to an implicit function argument (sin x, ln 2x…). */
function isArgAtom(n: XNode): boolean {
  if (n.name === "mo") {
    const t = n.text.trim();
    return t === "′" || t === "″" || t === "!";
  }
  if (n.name === "mspace" || n.name === "mtext") return false;
  if (naryParts(n)) return false;
  return true;
}

function naryParts(n: XNode): NaryParts | null {
  const baseOf = (x: XNode | undefined): string | null => {
    if (!x) return null;
    if (isMo(x) && NARY_CHARS.has(x.text.trim())) return x.text.trim();
    if (x.name === "mrow" && x.children.length === 1) return baseOf(x.children[0]);
    return null;
  };
  if (isMo(n)) {
    const c = baseOf(n);
    return c ? { chr: c, underOver: false } : null;
  }
  const [b, s1, s2] = n.children;
  const chr = baseOf(b);
  if (!chr) return null;
  switch (n.name) {
    case "msub": return { chr, sub: s1, underOver: false };
    case "msup": return { chr, sup: s1, underOver: false };
    case "msubsup": return { chr, sub: s1, sup: s2, underOver: false };
    case "munder": return { chr, sub: s1, underOver: true };
    case "mover": return { chr, sup: s1, underOver: true };
    case "munderover": return { chr, sub: s1, sup: s2, underOver: true };
    default: return null;
  }
}

function applyVariant(ctx: Ctx, variant: string | undefined, kind: string, length: number): Ctx {
  const c = { ...ctx };
  switch (variant) {
    case "normal": c.sty = "p"; break;
    case "bold": c.sty = "b"; break;
    case "italic": c.sty = "i"; break;
    case "bold-italic": c.sty = "bi"; break;
    case "double-struck": c.scr = "double-struck"; c.sty = "p"; break;
    case "script": c.scr = "script"; c.sty = "p"; break;
    case "bold-script": c.scr = "script"; c.sty = "b"; break;
    case "fraktur": c.scr = "fraktur"; c.sty = "p"; break;
    case "bold-fraktur": c.scr = "fraktur"; c.sty = "b"; break;
    case "sans-serif": c.scr = "sans-serif"; c.sty = "p"; break;
    case "bold-sans-serif": c.scr = "sans-serif"; c.sty = "b"; break;
    case "sans-serif-italic": c.scr = "sans-serif"; c.sty = "i"; break;
    case "monospace": c.scr = "monospace"; c.sty = "p"; break;
    default:
      // MathML rule: a multi-character <mi> (sin, ln, max, \operatorname…) is upright.
      if (kind === "mi" && length > 1 && !c.sty) c.sty = "p";
  }
  return c;
}

function colorFromStyle(style: string | undefined): string | undefined {
  if (!style) return undefined;
  const m = /(?:^|;)\s*color\s*:\s*([^;]+)/.exec(style);
  return m ? m[1].trim() : undefined;
}

const NAMED_COLORS: Record<string, string> = {
  red: "FF0000", blue: "0000FF", green: "008000", black: "000000", white: "FFFFFF", orange: "FFA500",
  purple: "800080", gray: "808080", grey: "808080", magenta: "FF00FF", cyan: "00FFFF", yellow: "FFFF00",
  brown: "A52A2A", violet: "EE82EE", teal: "008080", olive: "808000", lime: "00FF00", pink: "FFC0CB",
};

function normalizeColor(c: string): string | undefined {
  const s = c.trim().toLowerCase();
  if (NAMED_COLORS[s]) return NAMED_COLORS[s];
  let m = /^#([0-9a-f]{6})$/.exec(s);
  if (m) return m[1].toUpperCase();
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(s);
  if (m) return (m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  return undefined;
}
