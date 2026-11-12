import type { OperationReport } from "../office/word";

/** Tiny pub/sub so that ribbon command results show up in the task pane (shared runtime). */
class ReportChannel {
  private last: (OperationReport & { at: number }) | null = null;
  private listeners = new Set<(r: OperationReport & { at: number }) => void>();

  publish(r: OperationReport) {
    this.last = { ...r, at: Date.now() };
    this.listeners.forEach((l) => l(this.last!));
  }

  latest() {
    return this.last;
  }

  subscribe(l: (r: OperationReport & { at: number }) => void): () => void {
    this.listeners.add(l);
    return () => this.listeners.delete(l);
  }
}

export const commandReports = new ReportChannel();
