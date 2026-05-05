import React from "react";
import { createRoot } from "react-dom/client";
import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import {
  AlertTriangle,
  CheckCircle2,
  FileSpreadsheet,
  FolderOpen,
  Loader2,
  RotateCcw,
  Save,
  Table2,
  XCircle,
} from "lucide-react";

import "./styles.css";
import { Button } from "./components/ui/button";
import { Progress } from "./components/ui/progress";
import { Separator } from "./components/ui/separator";
import { cn } from "./lib/utils";

const initialState = {
  poPath: "",
  workOrderPath: "",
  outputPath: "",
  status: "ready",
  summary: null,
  warnings: [],
  error: "",
};

function App() {
  const [state, setState] = React.useState(initialState);
  const canGenerate = state.poPath && state.workOrderPath && state.outputPath && state.status !== "generating";

  async function pickPurchaseOrder() {
    try {
      await pickFile({
        title: "Select purchase order CSV",
        filters: [{ name: "CSV Purchase Order", extensions: ["csv"] }],
        onSelect: (path) => setState((current) => ({ ...current, poPath: path, error: "" })),
      });
    } catch (error) {
      setState((current) => ({ ...current, error: String(error.message || error) }));
    }
  }

  async function pickWorkOrder() {
    try {
      await pickFile({
        title: "Select work order XLSX",
        filters: [{ name: "Excel Work Order", extensions: ["xlsx"] }],
        onSelect: (path) => setState((current) => ({ ...current, workOrderPath: path, error: "" })),
      });
    } catch (error) {
      setState((current) => ({ ...current, error: String(error.message || error) }));
    }
  }

  async function pickOutput() {
    try {
      const selected = await save({
        title: "Save completed workbook",
        defaultPath: suggestedOutputName(state.workOrderPath),
        filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
      });
      if (selected) {
        setState((current) => ({ ...current, outputPath: selected, error: "" }));
      }
    } catch (error) {
      setState((current) => ({ ...current, error: desktopOnlyMessage(error) }));
    }
  }

  async function generateWorkbook() {
    if (!canGenerate) return;
    setState((current) => ({ ...current, status: "generating", summary: null, warnings: [], error: "" }));

    try {
      const rawResult = await invoke("generate_workbook", {
        poPath: state.poPath,
        workOrderPath: state.workOrderPath,
        outputPath: state.outputPath,
      });
      const result = JSON.parse(rawResult);
      setState((current) => ({
        ...current,
        status: "complete",
        summary: result.summary,
        warnings: result.warnings || [],
      }));
    } catch (error) {
      setState((current) => ({
        ...current,
        status: "failed",
        error: String(error?.message || error),
      }));
    }
  }

  function reset() {
    setState(initialState);
  }

  return (
    <main className="min-h-full bg-background">
      <header className="border-b bg-card">
        <div className="mx-auto flex max-w-7xl flex-col gap-3 px-5 py-4 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-center gap-3">
            <div className="grid h-11 w-11 place-items-center rounded-md bg-primary text-primary-foreground">
              <Table2 className="h-5 w-5" />
            </div>
            <div>
              <h1 className="text-2xl font-bold leading-tight tracking-normal">Pella Order Automation</h1>
              <p className="text-sm text-muted-foreground">Offline desktop workbook generation</p>
            </div>
          </div>
          <StatusBadge status={state.status} />
        </div>
      </header>

      <div className="mx-auto grid max-w-7xl gap-4 px-5 py-5 lg:grid-cols-[minmax(0,1.2fr)_minmax(360px,0.8fr)]">
        <section className="space-y-4">
          <Panel>
            <PanelTitle
              title="Source Files"
              caption="Pick the CSV purchase order, XLSX work order, and where to save the completed workbook."
            />
            <div className="grid gap-3 md:grid-cols-2">
              <FileRow
                label="Purchase order"
                type="CSV"
                path={state.poPath}
                icon={FileSpreadsheet}
                onPick={pickPurchaseOrder}
              />
              <FileRow
                label="Work order"
                type="XLSX"
                path={state.workOrderPath}
                icon={Table2}
                onPick={pickWorkOrder}
              />
            </div>
            <FileRow label="Completed workbook" type="XLSX output" path={state.outputPath} icon={Save} onPick={pickOutput} />

            <Separator className="my-4" />

            <div className="flex flex-col gap-2 sm:flex-row">
              <Button className="h-11 sm:min-w-56" disabled={!canGenerate} onClick={generateWorkbook}>
                {state.status === "generating" ? <Loader2 className="h-4 w-4 animate-spin" /> : <CheckCircle2 className="h-4 w-4" />}
                Generate Workbook
              </Button>
              <Button className="h-11" variant="outline" disabled={state.status === "generating"} onClick={reset}>
                <RotateCcw className="h-4 w-4" />
                Reset
              </Button>
            </div>

            {state.status === "generating" && (
              <div className="mt-4">
                <Progress value={65} />
              </div>
            )}
          </Panel>

          <Panel>
            <PanelTitle title="Output Summary" caption={state.summary?.filename || "No workbook generated yet."} />
            {state.summary ? <SummaryGrid summary={state.summary} /> : <EmptyState text="Generate a workbook to see the extracted order details." />}
          </Panel>
        </section>

        <aside className="space-y-4">
          <Panel>
            <PanelTitle title="Run Status" caption={statusCaption(state.status, state.warnings.length)} />
            <StatusMessage status={state.status} error={state.error} />
          </Panel>

          <Panel>
            <PanelTitle title="Warnings" caption="Items the user should review in the completed workbook." count={state.warnings.length} />
            {state.warnings.length ? (
              <div className="space-y-3">
                {state.warnings.map((warning, index) => (
                  <WarningCard key={`${warning.code}-${index}`} warning={warning} />
                ))}
              </div>
            ) : (
              <EmptyState text="No warnings." />
            )}
          </Panel>
        </aside>
      </div>
    </main>
  );
}

function Panel({ children }) {
  return <div className="rounded-lg border bg-card p-5 shadow-soft">{children}</div>;
}

function PanelTitle({ title, caption, count }) {
  return (
    <div className="mb-4 flex items-start justify-between gap-4">
      <div className="min-w-0">
        <h2 className="text-lg font-bold tracking-normal">{title}</h2>
        <p className="break-words text-sm text-muted-foreground">{caption}</p>
      </div>
      {typeof count === "number" && <span className={cn("rounded-md px-2 py-1 text-xs font-bold", count ? "bg-amber-100 text-amber-900" : "bg-muted text-muted-foreground")}>{count}</span>}
    </div>
  );
}

function FileRow({ label, type, path, icon: Icon, onPick }) {
  return (
    <div className="grid gap-3 rounded-md border bg-slate-50 p-4 sm:grid-cols-[44px_minmax(0,1fr)_auto] sm:items-center">
      <div className={cn("grid h-11 w-11 place-items-center rounded-md", path ? "bg-primary text-primary-foreground" : "bg-muted text-muted-foreground")}>
        <Icon className="h-5 w-5" />
      </div>
      <div className="min-w-0">
        <div className="flex flex-wrap items-center gap-2">
          <p className="text-sm font-bold">{label}</p>
          <span className="rounded-md bg-muted px-2 py-0.5 text-xs font-semibold text-muted-foreground">{type}</span>
        </div>
        <p className="mt-1 break-words text-sm text-muted-foreground">{path || "No file selected"}</p>
      </div>
      <Button variant="outline" onClick={onPick}>
        <FolderOpen className="h-4 w-4" />
        Choose
      </Button>
    </div>
  );
}

function StatusBadge({ status }) {
  const config = {
    ready: "border-border bg-card text-muted-foreground",
    generating: "border-blue-200 bg-blue-50 text-blue-800",
    complete: "border-emerald-200 bg-emerald-50 text-emerald-800",
    failed: "border-red-200 bg-red-50 text-red-800",
  }[status];
  const label = {
    ready: "Ready",
    generating: "Generating",
    complete: "Complete",
    failed: "Needs attention",
  }[status];
  return <span className={cn("rounded-md border px-3 py-1 text-sm font-bold", config)}>{label}</span>;
}

function StatusMessage({ status, error }) {
  if (error) {
    return (
      <Message tone="error" icon={XCircle}>
        {error}
      </Message>
    );
  }
  if (status === "complete") {
    return (
      <Message tone="success" icon={CheckCircle2}>
        Workbook generated successfully.
      </Message>
    );
  }
  if (status === "generating") {
    return (
      <Message tone="info" icon={Loader2} spin>
        Generating workbook...
      </Message>
    );
  }
  return (
    <Message tone="info" icon={FileSpreadsheet}>
      Ready for source files.
    </Message>
  );
}

function Message({ tone, icon: Icon, spin = false, children }) {
  const tones = {
    info: "border-blue-200 bg-blue-50 text-blue-900",
    success: "border-emerald-200 bg-emerald-50 text-emerald-900",
    error: "border-red-200 bg-red-50 text-red-900",
  };
  return (
    <div className={cn("flex gap-3 rounded-md border p-3 text-sm", tones[tone])}>
      <Icon className={cn("mt-0.5 h-4 w-4 shrink-0", spin && "animate-spin")} />
      <p>{children}</p>
    </div>
  );
}

function WarningCard({ warning }) {
  return (
    <div className="flex gap-3 rounded-md border border-amber-200 bg-amber-50 p-3 text-sm text-amber-950">
      <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
      <div className="min-w-0">
        <p className="font-bold">{warning.code}</p>
        <p className="break-words">{warning.message}</p>
      </div>
    </div>
  );
}

function SummaryGrid({ summary }) {
  const items = [
    ["Order number", summary.order_number],
    ["Pella PO", summary.pella_po_number],
    ["Customer", summary.customer_name],
    ["Request dates", summary.request_dates?.join(", ") || "Blank"],
    ["Jamb lines", summary.jamb_line_count],
    ["Generated BOM rows", summary.generated_bom_rows],
    ["Output file", summary.filename],
  ];
  return (
    <div className="grid gap-3 sm:grid-cols-2">
      {items.map(([label, value]) => (
        <div key={label} className="min-w-0 rounded-md border bg-slate-50 p-3">
          <p className="text-xs font-semibold text-muted-foreground">{label}</p>
          <p className="mt-1 break-words text-sm font-bold">{value || "Unknown"}</p>
        </div>
      ))}
    </div>
  );
}

function EmptyState({ text }) {
  return <div className="rounded-md border border-dashed bg-slate-50 p-4 text-sm text-muted-foreground">{text}</div>;
}

async function pickFile({ title, filters, onSelect }) {
  try {
    const selected = await open({
      title,
      multiple: false,
      directory: false,
      filters,
    });
    if (typeof selected === "string") {
      onSelect(selected);
    }
  } catch (error) {
    throw new Error(desktopOnlyMessage(error));
  }
}

function desktopOnlyMessage(error) {
  const text = String(error?.message || error || "");
  if (text.includes("__TAURI__") || text.includes("not found") || text.includes("is not a function")) {
    return "This UI must be opened through the Tauri desktop app, not directly in a browser.";
  }
  return text || "The desktop file dialog could not be opened.";
}

function suggestedOutputName(workOrderPath) {
  const fallback = "Pella Completed Workbook.xlsx";
  if (!workOrderPath) return fallback;
  const normalized = workOrderPath.replace(/\\/g, "/");
  const fileName = normalized.split("/").pop() || fallback;
  return `${fileName.replace(/\.[^.]+$/, "")} - Completed.xlsx`;
}

function statusCaption(status, warningCount) {
  if (status === "generating") return "Processing files locally on this computer.";
  if (status === "complete") return warningCount ? `${warningCount} warning${warningCount === 1 ? "" : "s"} found.` : "No warnings found.";
  if (status === "failed") return "Generation stopped before creating a workbook.";
  return "Waiting for selections.";
}

createRoot(document.getElementById("root")).render(<App />);
