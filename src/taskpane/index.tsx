/**
 * Entry point of the shared runtime: hosts the task pane UI AND the ribbon / context-menu /
 * keyboard-shortcut command "Convert to Formula" (same JavaScript context, same engine).
 */
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { convertSelection, OperationReport } from "../office/word";
import { commandReports } from "./store";
import "./styles.css";

async function runConvertCommand(event?: Office.AddinCommands.Event) {
  let report: OperationReport;
  try {
    report = await convertSelection();
  } catch (err) {
    report = { ok: false, message: String(err), converted: 0, failures: [] };
  }
  commandReports.publish(report);
  // Without feedback the user would not know why nothing happened: open the pane on problems.
  if (!report.ok) {
    try {
      await Office.addin.showAsTaskpane();
    } catch {
      /* SharedRuntime not available: the report is still shown next time the pane opens */
    }
  }
  event?.completed();
}

function mount() {
  const el = document.getElementById("root");
  if (el) createRoot(el).render(<App />);
}

if (typeof Office !== "undefined") {
  Office.onReady(() => {
    try {
      Office.actions.associate("convertSelection", runConvertCommand);
      Office.actions.associate("ConvertToFormulaShortcut", () => runConvertCommand());
    } catch {
      /* Office.actions unavailable (e.g. opened outside Word) */
    }
    mount();
  });
} else {
  mount();
}
