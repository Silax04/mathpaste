import { useEffect, useMemo, useRef, useState } from "react";
import { analyze, LayoutMode } from "../engine";
import { convertSelection, insertText, isWordAvailable, OperationReport } from "../office/word";
import { commandReports } from "./store";

const EXAMPLE = String.raw`C_{m1} = \frac{\varepsilon_0 \pi r^2}{d + r - \sqrt{r^2 + d^2}}

C_{m2} = \frac{\varepsilon_0 \pi r^2}{d}

C_{m3} = \varepsilon_0 \, r \left( \frac{\pi r}{d} + 2 \ln\!\left(\frac{8 \pi r}{d}\right) - 2 \right)`;

const STORAGE_KEY = "mathpaste.draft.v1";

function loadDraft(): { text: string; mode: LayoutMode } {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const d = JSON.parse(raw);
      return { text: typeof d.text === "string" ? d.text : "", mode: d.mode === "original" ? "original" : "lines" };
    }
  } catch {
    /* storage unavailable */
  }
  return { text: "", mode: "lines" };
}

function saveDraft(text: string, mode: LayoutMode) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ text, mode }));
  } catch {
    /* ignore */
  }
}

function useDebounced<T>(value: T, ms: number): T {
  const [v, setV] = useState(value);
  useEffect(() => {
    const t = setTimeout(() => setV(value), ms);
    return () => clearTimeout(t);
  }, [value, ms]);
  return v;
}

type Status = { kind: "idle" } | { kind: "busy"; label: string } | { kind: "done"; report: OperationReport; source: "pane" | "command" };

export function App() {
  const initial = useMemo(loadDraft, []);
  const [text, setText] = useState(initial.text);
  const [mode, setMode] = useState<LayoutMode>(initial.mode);
  const [showPreview, setShowPreview] = useState(false);
  const [status, setStatus] = useState<Status>(() => {
    const last = commandReports.latest();
    return last ? { kind: "done", report: last, source: "command" } : { kind: "idle" };
  });
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const debounced = useDebounced(text, 200);
  const wordOk = isWordAvailable();

  useEffect(() => saveDraft(text, mode), [text, mode]);
  useEffect(() => commandReports.subscribe((r) => setStatus({ kind: "done", report: r, source: "command" })), []);

  const segments = useMemo(() => (showPreview ? analyze(debounced) : []), [debounced, showPreview]);
  const maths = segments.filter((s) => s.kind === "math");
  const failed = maths.filter((s) => s.conversion && !s.conversion.ok).length;

  const busy = status.kind === "busy";

  async function onInsert() {
    setStatus({ kind: "busy", label: "Conversion et insertion…" });
    await new Promise((r) => setTimeout(r, 0)); // let the spinner paint
    const report = await insertText(text, mode);
    setStatus({ kind: "done", report, source: "pane" });
    if (report.failures.length) setShowPreview(true);
  }

  async function onConvertSelection() {
    setStatus({ kind: "busy", label: "Conversion de la sélection…" });
    const report = await convertSelection();
    setStatus({ kind: "done", report, source: "command" });
  }

  function onClear() {
    setText("");
    setStatus({ kind: "idle" });
    textareaRef.current?.focus();
  }

  return (
    <div className="app">
      <header className="header">
        <div className="logo" aria-hidden>√x</div>
        <div>
          <h1>MathPaste</h1>
          <p className="subtitle">LaTeX → équations Word natives</p>
        </div>
      </header>

      {!wordOk && <div className="banner warn">Document Word inaccessible : ouvrez MathPaste depuis Word pour insérer.</div>}

      <label className="field-label" htmlFor="latex">Code LaTeX</label>
      <textarea
        id="latex"
        ref={textareaRef}
        className="latex-input"
        spellCheck={false}
        placeholder={"Collez ici votre LaTeX (ChatGPT, Claude, Overleaf…)\n\nC_{m1} = \\frac{\\varepsilon_0 \\pi r^2}{d}"}
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && (e.ctrlKey || e.metaKey) && !busy) {
            e.preventDefault();
            void onInsert();
          }
        }}
      />

      <fieldset className="mode">
        <legend>Disposition</legend>
        <label>
          <input type="radio" name="mode" checked={mode === "lines"} onChange={() => setMode("lines")} />
          Une équation par ligne
        </label>
        <label>
          <input type="radio" name="mode" checked={mode === "original"} onChange={() => setMode("original")} />
          Conserver la disposition originale
        </label>
      </fieldset>

      <div className="actions">
        <button className="btn primary" disabled={busy || !text.trim() || !wordOk} onClick={onInsert} title="Ctrl+Entrée">
          {busy ? <span className="spinner" aria-hidden /> : null}
          Insérer dans Word
        </button>
        <div className="row">
          <button className={`btn ${showPreview ? "active" : ""}`} onClick={() => setShowPreview((v) => !v)} disabled={!text.trim()}>
            {showPreview ? "Masquer l'aperçu" : "Aperçu"}
          </button>
          <button className="btn" onClick={onClear} disabled={busy || !text}>Effacer</button>
        </div>
      </div>

      <StatusView status={status} />

      {showPreview && text.trim() && (
        <section className="preview" aria-live="polite">
          <div className="preview-head">
            <span>Aperçu</span>
            <span className="count">
              {maths.length} expression{maths.length > 1 ? "s" : ""}
              {failed ? <span className="bad"> · {failed} en erreur</span> : null}
            </span>
          </div>
          {maths.length === 0 && <p className="muted">Aucune expression LaTeX détectée.</p>}
          {maths.map((s, i) => (
            <div key={`${s.start}-${i}`} className={`eq ${s.conversion?.ok ? "ok" : "ko"}`}>
              <span className="eq-num">{i + 1}</span>
              {s.conversion?.ok ? (
                <div className="eq-render" dangerouslySetInnerHTML={{ __html: s.conversion.mathml }} />
              ) : (
                <div className="eq-error">
                  <strong>{s.conversion && !s.conversion.ok ? s.conversion.error : "Erreur"}</strong>
                  <code>{s.source.length > 140 ? s.source.slice(0, 140) + "…" : s.source}</code>
                </div>
              )}
            </div>
          ))}
        </section>
      )}

      <section className="card">
        <h2>Conversion directe dans le document</h2>
        <p className="muted">
          Sélectionnez du LaTeX dans votre document puis cliquez <b>Convert to Formula</b> (onglet <b>MathPaste</b> du ruban,
          clic droit, ou <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Maj</kbd>+<kbd>M</kbd>).
        </p>
        <button className="btn" onClick={onConvertSelection} disabled={busy || !wordOk}>Convertir la sélection</button>
      </section>

      <footer className="footer">
        {!text && (
          <button className="link" onClick={() => setText(EXAMPLE)}>Charger l'exemple</button>
        )}
        <span className="muted">Conversion 100 % locale · OMML natif</span>
      </footer>
    </div>
  );
}

function StatusView({ status }: { status: Status }) {
  if (status.kind === "idle") return null;
  if (status.kind === "busy")
    return (
      <div className="banner info" role="status">
        <span className="spinner dark" aria-hidden /> {status.label}
      </div>
    );
  const r = status.report;
  return (
    <div className={`banner ${r.ok ? "success" : r.converted ? "warn" : "error"}`} role="status">
      <div>
        {status.source === "command" ? <b>Convert to Formula : </b> : null}
        {r.message}
      </div>
      {r.failures.length > 0 && (
        <ul className="failures">
          {r.failures.map((f, i) => (
            <li key={i}>
              <code>{f.latex.length > 90 ? f.latex.slice(0, 90) + "…" : f.latex}</code>
              <span>{f.error}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
