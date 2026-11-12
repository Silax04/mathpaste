import { describe, expect, it } from "vitest";
import { analyze, buildInsertion, convertLatex, extractSegments } from "../src/engine";
import { parseXml } from "../src/engine/xml";

const REF = [
  String.raw`C_{m1} = \frac{\varepsilon_0 \pi r^2}{d + r - \sqrt{r^2 + d^2}}`,
  String.raw`C_{m2} = \frac{\varepsilon_0 \pi r^2}{d}`,
  String.raw`C_{m3} = \varepsilon_0 \, r \left( \frac{\pi r}{d} + 2 \ln\!\left(\frac{8 \pi r}{d}\right) - 2 \right)`,
];

function omml(latex: string, display = true): string {
  const r = convertLatex(latex, { display });
  if (!r.ok) throw new Error(r.error);
  return r.omml;
}

/** Strips formatting so structure assertions stay readable. */
function skeleton(x: string): string {
  return x
    .replace(/<w:rPr>.*?<\/w:rPr>/g, "")
    .replace(/<m:ctrlPr><\/m:ctrlPr>/g, "")
    .replace(/ xml:space="preserve"/g, "");
}

function wellFormed(x: string) {
  // our parser is lenient; check tag balance explicitly
  const stack: string[] = [];
  for (const m of x.matchAll(/<(\/?)([\w:]+)[^>]*?(\/?)>/g)) {
    if (m[3] === "/") continue;
    if (m[1] === "/") expect(stack.pop()).toBe(m[2]);
    else stack.push(m[2]);
  }
  expect(stack).toEqual([]);
  parseXml(x);
}

describe("reference expressions", () => {
  it("C_m1: subscript, vertical fraction, square root", () => {
    const s = skeleton(omml(REF[0]));
    wellFormed(s);
    expect(s).toContain("<m:sSub><m:sSubPr></m:sSubPr><m:e><m:r><m:t>C</m:t></m:r></m:e><m:sub><m:r><m:t>m1</m:t></m:r></m:sub></m:sSub>");
    expect(s).toMatch(/<m:f><m:fPr><\/m:fPr><m:num>.*ε.*π.*<m:sSup>.*r.*2.*<\/m:num><m:den>.*<m:rad><m:radPr><m:degHide m:val="1"\/><\/m:radPr><m:deg><\/m:deg><m:e><m:sSup>/);
    expect(s).not.toContain("\\");
  });
  it("C_m2", () => {
    const s = skeleton(omml(REF[1]));
    expect(s).toContain("<m:den><m:r><m:t>d</m:t></m:r></m:den>");
  });
  it("C_m3: \\left( \\right) -> m:d, \\ln -> m:func, \\, -> thin space, \\! dropped", () => {
    const s = skeleton(omml(REF[2]));
    wellFormed(s);
    expect(s).toContain("<m:func><m:funcPr></m:funcPr><m:fName><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>ln</m:t></m:r></m:fName><m:e><m:d>");
    expect((s.match(/<m:d>/g) ?? []).length).toBe(2);
    expect(s).toContain("\u2009r");
    expect(s).toContain("−2</m:t>");
  });
  it("all runs use Cambria Math", () => {
    const x = omml(REF[0]);
    const runs = x.match(/<m:r>/g)!.length;
    expect((x.match(/<m:r>(<m:rPr>.*?<\/m:rPr>)?<w:rPr><w:rFonts w:ascii="Cambria Math"/g) ?? []).length).toBe(runs);
  });
  it("applies font size", () => {
    const r = convertLatex(REF[1], { sizeHalfPoints: 28 });
    expect(r.ok && r.omml).toContain('<w:sz w:val="28"/>');
  });
});

describe("LaTeX coverage", () => {
  const cases: Array<[string, RegExp | string]> = [
    [String.raw`\frac{a}{b}`, "<m:f>"],
    [String.raw`\sqrt{x}`, '<m:degHide m:val="1"/>'],
    [String.raw`\sqrt[3]{x}`, /<m:deg><m:r><m:t>3<\/m:t><\/m:r><\/m:deg>/],
    [String.raw`x^2`, "<m:sSup>"],
    [String.raw`x_i^2`, "<m:sSubSup>"],
    [String.raw`\alpha + \beta = \gamma`, "α+β=γ"],
    [String.raw`\ln(x)`, /<m:func>.*<m:t>ln<\/m:t>.*<m:e><m:d>/],
    [String.raw`\sin x`, /<m:func>.*sin.*<m:e><m:r><m:t>x<\/m:t>/],
    [String.raw`\sin^2\theta`, /<m:func><m:funcPr><\/m:funcPr><m:fName><m:sSup>.*sin.*<\/m:fName><m:e><m:r><m:t>θ/],
    [String.raw`\lim_{x \to 0} f(x)`, /<m:func>.*<m:fName><m:limLow>.*lim.*→/],
    [String.raw`a \approx b \neq c \pm d \cdot e`, "≈"],
    [String.raw`a\quad b`, "a\u2003b"],
    [String.raw`\sum_{i=1}^{n} i^2`, /<m:nary><m:naryPr><m:chr m:val="∑"\/><m:limLoc m:val="undOvr"\/>/],
    [String.raw`\int_0^\infty e^{-x}\,dx`, /<m:nary><m:naryPr><m:limLoc m:val="subSup"\/><\/m:naryPr><m:sub><m:r><m:t>0<\/m:t>/],
    [String.raw`\hat{x}`, /<m:acc><m:accPr><\/m:accPr>/],
    [String.raw`\vec{v}`, '<m:chr m:val="⃗"/>'],
    [String.raw`\bar{y}`, '<m:chr m:val="̅"/>'],
    [String.raw`\overline{AB}`, '<m:bar><m:barPr><m:pos m:val="top"/>'],
    [String.raw`\underbrace{a+b}_{n}`, /<m:limLow>.*<m:groupChr>/],
    [String.raw`\begin{pmatrix} a & b \\ c & d \end{pmatrix}`, /<m:d><m:dPr><\/m:dPr><m:e><m:m>.*<m:count m:val="2"\/>/],
    [String.raw`f(x) = \begin{cases} 1 & x>0 \\ 0 & \text{sinon} \end{cases}`, /<m:begChr m:val="\{"\/><m:endChr m:val=""\/>.*<m:mcJc m:val="left"\/>.*<m:nor\/>.*sinon/],
    [String.raw`\begin{aligned} a &= b+c \\ d &= e \end{aligned}`, /<m:eqArr>.*a&amp;=b\+c.*d&amp;=e/],
    [String.raw`\mathbf{F} = m\mathbf{a}`, /<m:sty m:val="b"\/>.*F/],
    [String.raw`\mathrm{d}x`, /<m:sty m:val="p"\/>.*<m:t>d<\/m:t>/],
    [String.raw`\mathbb{R}`, "ℝ"],
    [String.raw`\operatorname{tr}(A)`, /<m:func>.*tr/],
    [String.raw`\left. \frac{df}{dx} \right|_{x=0}`, /<m:sSub>.*<m:d><m:dPr><m:begChr m:val=""\/><m:endChr m:val="\|"\/>/],
    [String.raw`\left\{ a \middle| b \right\}`, /<m:sepChr m:val="\|"\/>/],
    [String.raw`\binom{n}{k}`, /<m:d>.*<m:type m:val="noBar"\/>/],
    [String.raw`\boxed{E=mc^2}`, "<m:borderBox>"],
    [String.raw`x' + f''`, "′′"],
    [String.raw`\text{vitesse } v = \frac{d}{t}`, /<m:nor\/>.*vitesse/],
    [String.raw`\R \to \R`, "ℝ"],
    [String.raw`\nabla \cdot \vec{E} = \frac{\rho}{\varepsilon_0}`, "∇"],
    [String.raw`\iint_D f\,dA`, '<m:chr m:val="∬"/>'],
    [String.raw`\Delta x`, /<m:sty m:val="p"\/>.*Δ/],
    [String.raw`\overset{!}{=}`, "<m:limUpp>"],
    [String.raw`\cancel{x}`, "<m:strikeBLTR"],
  ];
  for (const [latex, expected] of cases) {
    it(latex, () => {
      const s = skeleton(omml(latex));
      wellFormed(s);
      if (typeof expected === "string") expect(s).toContain(expected);
      else expect(s).toMatch(expected);
    });
  }

  it("color", () => { expect(omml(String.raw`\color{red}{x}`)).toContain('<w:color w:val="FF0000"/>'); });

  it("delimiters and environments are stripped", () => {
    expect(skeleton(omml("$$x^2$$"))).toContain("<m:sSup>");
    expect(skeleton(omml("\\[x^2\\]"))).toContain("<m:sSup>");
    expect(skeleton(omml("\\begin{equation} E = mc^2 \\label{eq:1} \\end{equation}"))).toContain("<m:sSup>");
    expect(skeleton(omml("\\begin{align} a &= b \\\\ c &= d \\end{align}"))).toContain("<m:eqArr>");
  });
});

describe("errors", () => {
  it("empty", () => {
    const r = convertLatex("   ");
    expect(r.ok).toBe(false);
    expect(!r.ok && r.kind).toBe("empty");
  });
  it("unbalanced braces", () => {
    const r = convertLatex(String.raw`\frac{a}{`);
    expect(r.ok).toBe(false);
    expect(!r.ok && r.error).toMatch(/Accolades|invalide|Argument/);
  });
  it("unsupported command", () => {
    const r = convertLatex(String.raw`\foo{x}`);
    expect(!r.ok && r.kind).toBe("unsupported");
    expect(!r.ok && r.error).toContain("\\foo");
  });
  it("malformed script", () => {
    expect(convertLatex("a^^b").ok).toBe(false);
  });
  it("one bad expression does not block the others", () => {
    const res = buildInsertion([REF[0], String.raw`\frac{a}{`, REF[2]].join("\n\n"), "lines");
    expect(res.converted).toBe(2);
    expect(res.failed).toHaveLength(1);
    expect((res.ooxml.match(/<m:oMathPara>/g) ?? []).length).toBe(2);
  });
});

describe("expression detection", () => {
  const maths = (input: string) => extractSegments(input).filter((s) => s.kind === "math").map((s) => s.source);

  it("three reference lines separated by blank lines", () => {
    expect(maths(REF.join("\n\n"))).toEqual(REF);
  });
  it("three reference lines without blank lines", () => {
    expect(maths(REF.join("\n"))).toEqual(REF);
  });
  it("Windows line endings / Word paragraph marks", () => {
    expect(maths(REF.join("\r\n"))).toEqual(REF.map((r) => r));
  });
  it("leading / trailing spaces are excluded", () => {
    const segs = extractSegments(`   ${REF[1]}   `);
    const m = segs.filter((s) => s.kind === "math");
    expect(m).toHaveLength(1);
    expect(m[0].source).toBe(REF[1]);
    expect(m[0].start).toBe(3);
  });
  it("expression spanning several lines (unbalanced braces)", () => {
    const input = "C_{m1} = \\frac{\\varepsilon_0 \\pi r^2}{d + r\n - \\sqrt{r^2 + d^2}}";
    expect(maths(input)).toEqual([input]);
  });
  it("expression continued on next line by an operator", () => {
    const input = "y = a x^2\n+ b x + c";
    expect(maths(input)).toEqual([input]);
  });
  it("aligned block over several lines", () => {
    const input = "\\begin{aligned}\na &= b \\\\\nc &= d\n\\end{aligned}";
    expect(maths(input)).toEqual([input]);
  });
  it("ChatGPT style \\[ … \\] blocks with prose", () => {
    const input = "La capacité vaut :\n\\[\nC = \\frac{q}{U}\n\\]\noù $q$ est la charge.";
    const segs = extractSegments(input);
    expect(segs.filter((s) => s.kind === "math").map((s) => s.source)).toEqual(["\\[\nC = \\frac{q}{U}\n\\]", "$q$"]);
    expect(segs.find((s) => s.source === "$q$")!.display).toBe(false);
  });
  it("$$ blocks and code fences", () => {
    expect(maths("```latex\nx^2 + y^2 = r^2\n```")).toEqual(["x^2 + y^2 = r^2"]);
    expect(maths("$$a^2$$ et $$b^2$$")).toEqual(["$$a^2$$", "$$b^2$$"]);
  });
  it("dollar amounts are not math", () => {
    expect(maths("Cela coûte $5 et $10.")).toEqual([]);
  });
  it("plain prose is not math", () => {
    expect(maths("Bonjour, ceci est un paragraphe normal.")).toEqual([]);
  });
  it("text mixed with an undelimited formula", () => {
    const input = String.raw`La capacité est C_{m2} = \frac{\varepsilon_0 \pi r^2}{d} où r est le rayon.`;
    expect(maths(input)).toEqual([String.raw`C_{m2} = \frac{\varepsilon_0 \pi r^2}{d}`]);
  });
  it("trailing sentence punctuation is excluded", () => {
    expect(maths(String.raw`E = mc^2.`)).toEqual(["E = mc^2"]);
    expect(maths(String.raw`on obtient x^2 + 1, donc`)).toEqual(["x^2 + 1"]);
  });
  it("markdown bullets are excluded", () => {
    expect(maths("- x^2 = 4\n- y_1 = 3")).toEqual(["x^2 = 4", "y_1 = 3"]);
  });
  it("simple equation without backslash", () => {
    expect(maths("a = b + c")).toEqual(["a = b + c"]);
  });
  it("integral with dx stays one expression", () => {
    expect(maths(String.raw`\int_0^1 f(x) dx`)).toEqual([String.raw`\int_0^1 f(x) dx`]);
  });
});

describe("layout", () => {
  it("lines mode: one display equation per paragraph", () => {
    const res = buildInsertion(REF.join("\n"), "lines");
    expect(res.paragraphs).toHaveLength(3);
    expect(res.paragraphs.every((p) => p.display)).toBe(true);
    expect(res.ooxml).toContain("<pkg:package");
    expect((res.ooxml.match(/<w:p>/g) ?? []).length).toBe(3);
  });
  it("original mode keeps text, blank lines and inline equations", () => {
    const input = "Titre\n\nOn a $x^2$ ici.\n\n" + REF[1];
    const res = buildInsertion(input, "original");
    expect(res.paragraphs.map((p) => p.parts.map((x) => x.kind).join("+"))).toEqual(["text", "", "text+math+text", "", "math"]);
    expect(res.paragraphs[4].display).toBe(true);
    expect(res.paragraphs[2].display).toBe(false);
  });
  it("analyze reports per-expression status", () => {
    const segs = analyze(REF.join("\n"));
    expect(segs.filter((s) => s.conversion?.ok)).toHaveLength(3);
  });
});
