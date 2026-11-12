/**
 * Generates the Flat OPC packages used by scripts/word-validate.ps1 to validate the
 * engine output inside the real Microsoft Word (COM InsertXML = same importer as insertOoxml).
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { buildInsertion, buildPackage, convertLatex } from "../src/engine";

const out = join(process.cwd(), "tests", "word-out");
mkdirSync(out, { recursive: true });

export const REF = [
  String.raw`C_{m1} = \frac{\varepsilon_0 \pi r^2}{d + r - \sqrt{r^2 + d^2}}`,
  String.raw`C_{m2} = \frac{\varepsilon_0 \pi r^2}{d}`,
  String.raw`C_{m3} = \varepsilon_0 \, r \left( \frac{\pi r}{d} + 2 \ln\!\left(\frac{8 \pi r}{d}\right) - 2 \right)`,
];

// 1. Task pane, "one per line" mode, the three reference equations in ONE operation
writeFileSync(join(out, "ref.xml"), buildInsertion(REF.join("\n\n"), "lines").ooxml, "utf8");

// 2. Coverage sheet: label + equation for many constructs
const coverage = [
  String.raw`\sum_{i=1}^{n} i^2 = \frac{n(n+1)(2n+1)}{6}`,
  String.raw`\int_0^\infty e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}`,
  String.raw`\lim_{x \to 0} \frac{\sin x}{x} = 1`,
  String.raw`\hat{x}, \bar{y}, \vec{v}, \dot{a}, \tilde{n}, \overline{AB}`,
  String.raw`\begin{pmatrix} a & b \\ c & d \end{pmatrix}^{-1} = \frac{1}{ad-bc}\begin{pmatrix} d & -b \\ -c & a \end{pmatrix}`,
  String.raw`f(x) = \begin{cases} x^2 & \text{si } x \geq 0 \\ -x & \text{sinon} \end{cases}`,
  String.raw`\begin{aligned} a &= b + c \\ d + e &= f \end{aligned}`,
  String.raw`\mathbf{F} = m\mathbf{a},\quad \mathrm{d}x,\quad \mathbb{R},\quad \Delta\Omega`,
  String.raw`\left. \frac{\partial f}{\partial x} \right|_{x=0} \approx \sqrt[3]{x} \pm \binom{n}{k} \neq \infty`,
  String.raw`\nabla \cdot \vec{E} = \frac{\rho}{\varepsilon_0}, \quad \sin^2\theta + \cos^2\theta = 1`,
  String.raw`\underbrace{a + b}_{n} \cdot \overbrace{c}^{k}, \quad \boxed{E = mc^2}`,
  String.raw`\log_2 x \leq \max_{i} x_i \times \alpha\beta\gamma`,
];
const paras = coverage.flatMap((l) => {
  const r = convertLatex(l);
  if (!r.ok) throw new Error(`${l}: ${r.error}`);
  return [{ parts: [{ kind: "text" as const, text: l }] }, { parts: [{ kind: "math" as const, omml: r.omml }], display: true }];
});
writeFileSync(join(out, "coverage.xml"), buildPackage(paras), "utf8");

// 3. Selection replacement payloads (what "Convert to Formula" inserts for a whole-paragraph selection)
REF.forEach((l, i) => {
  const r = convertLatex(l);
  if (!r.ok) throw new Error(r.error);
  writeFileSync(join(out, `sel${i + 1}.xml`), buildPackage([{ parts: [{ kind: "math", omml: r.omml }], display: true }]), "utf8");
});
// Inline (mixed text) payload
const inl = convertLatex(REF[1], { display: false });
if (!inl.ok) throw new Error(inl.error);
writeFileSync(join(out, "inline.xml"), buildPackage([{ parts: [{ kind: "math", omml: inl.omml }] }]), "utf8");
writeFileSync(join(out, "ref-latex.txt"), REF.join("\n"), "utf8");
console.log("fixtures written to", out);
