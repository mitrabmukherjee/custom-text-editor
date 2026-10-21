"use client";

import { useCallback, useRef, useState } from "react";
import {
  TranscriptEditor,
  type Author,
  type TrackedChange,
  type TranscriptEditorHandle,
} from "@/components/editor";
import { SAMPLES, type SampleId } from "@/lib/samples";

type LogEntry = { id: number; at: string; message: string };

const AUTHORS: Author[] = [
  { id: "typist-1", name: "Typist" },
  { id: "proof-1", name: "Proofreader" },
  { id: "qa-1", name: "QA reviewer" },
];

function countWords(text: string) {
  const t = text.trim();
  return t ? t.split(/\s+/).length : 0;
}

function sectionTitle(text: string) {
  return (
    <h2 className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-zinc-500">
      {text}
    </h2>
  );
}

function controlBtn(active = false) {
  return `rounded-md border px-2.5 py-1.5 text-xs font-medium ${
    active
      ? "border-zinc-500 bg-zinc-100 text-zinc-900"
      : "border-zinc-700 text-zinc-300 hover:bg-zinc-800"
  }`;
}

function downloadBuffer(buffer: ArrayBuffer, filename: string) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function EditorDemo() {
  const editorRef = useRef<TranscriptEditorHandle | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const logId = useRef(0);
  const [editable, setEditable] = useState(true);
  const [tracking, setTracking] = useState(false);
  const [author, setAuthor] = useState<Author>(AUTHORS[0]);
  const [stats, setStats] = useState({ words: 0, chars: 0 });
  const [pageCount, setPageCount] = useState(1);
  const [changes, setChanges] = useState<TrackedChange[]>([]);
  const [output, setOutput] = useState<{ label: string; text: string } | null>(
    null,
  );
  const [log, setLog] = useState<LogEntry[]>([]);

  const pushLog = useCallback((message: string) => {
    logId.current += 1;
    const entry = {
      id: logId.current,
      at: new Date().toLocaleTimeString(),
      message,
    };
    setLog((prev) => [entry, ...prev].slice(0, 12));
  }, []);

  const refreshStats = useCallback(() => {
    const plain = editorRef.current?.getPlainText() ?? "";
    setStats({ words: countWords(plain), chars: plain.length });
  }, []);

  const loadSample = (id: SampleId) => {
    editorRef.current?.loadPlainText(SAMPLES[id].text);
    refreshStats();
    pushLog(`Loaded "${SAMPLES[id].label}" via loadPlainText()`);
  };

  const openDocx = async (file: File | undefined) => {
    if (!file) return;
    try {
      await editorRef.current?.importDocx(await file.arrayBuffer());
      refreshStats();
      pushLog(`importDocx("${file.name}")`);
    } catch (e) {
      pushLog(`importDocx failed: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const saveDocx = async () => {
    const handle = editorRef.current;
    if (!handle) return;
    try {
      const buffer = await handle.exportDocx();
      downloadBuffer(buffer, "transcript.docx");
      pushLog(`exportDocx() → ${buffer.byteLength.toLocaleString()} bytes`);
    } catch (e) {
      pushLog(`exportDocx failed: ${e instanceof Error ? e.message : String(e)}`);
    }
  };

  return (
    <div className="flex min-h-0 flex-1">
      <main className="min-h-0 min-w-0 flex-1 bg-zinc-900 p-3">
        <TranscriptEditor
          editable={editable}
          initialContent={SAMPLES.deposition.text}
          author={author}
          trackChanges={tracking}
          editorRef={editorRef}
          studioActions={{
            onPlayPause: () => pushLog("studioActions.onPlayPause()"),
            onPrint: () => {
              pushLog("studioActions.onPrint()");
              document.documentElement.classList.add("printing");
              const done = () => {
                document.documentElement.classList.remove("printing");
                window.removeEventListener("afterprint", done);
              };
              window.addEventListener("afterprint", done);
              window.print();
            },
          }}
          onReady={() => {
            refreshStats();
            pushLog("onReady");
          }}
          onUpdate={refreshStats}
          onChangesUpdate={setChanges}
          onTrackChangesChange={(enabled) => {
            setTracking(enabled);
            pushLog(`Track changes ${enabled ? "on" : "off"}`);
          }}
          onLayout={({ pageCount: n }) => setPageCount(n)}
          className="h-full"
        />
      </main>

      <aside className="flex w-80 shrink-0 flex-col gap-5 overflow-y-auto border-l border-zinc-800 bg-zinc-950 p-4">
        <section>
          {sectionTitle("Mode")}
          <label className="flex items-center gap-2 text-sm text-zinc-300">
            <input
              type="checkbox"
              checked={!editable}
              onChange={(e) => {
                setEditable(!e.target.checked);
                pushLog(e.target.checked ? "Read-only on" : "Read-only off");
              }}
            />
            Read-only
          </label>
        </section>

        <section>
          {sectionTitle("Track changes")}
          <label className="flex items-center gap-2 text-sm text-zinc-300">
            <input
              type="checkbox"
              checked={tracking}
              onChange={(e) => setTracking(e.target.checked)}
            />
            Track changes
          </label>
          <div className="mt-2 flex flex-wrap gap-2">
            {AUTHORS.map((a) => (
              <button
                key={a.id}
                type="button"
                className={controlBtn(author.id === a.id)}
                onClick={() => {
                  setAuthor(a);
                  pushLog(`Author → ${a.name}`);
                }}
              >
                {a.name}
              </button>
            ))}
          </div>
          <p className="mt-2 text-xs text-zinc-500">
            {changes.length} change{changes.length === 1 ? "" : "s"} — shown as balloons beside
            the page (All Markup). Use Review → Accept all / Reject all, or click a balloon to
            jump to that change.
          </p>
          <div className="mt-2 flex gap-2">
            <button
              type="button"
              className="rounded border border-emerald-800 px-2 py-1 text-xs text-emerald-300 hover:bg-emerald-950 disabled:opacity-40"
              disabled={!editable || !changes.length}
              onClick={() => editorRef.current?.acceptAll()}
            >
              Accept all
            </button>
            <button
              type="button"
              className="rounded border border-rose-900 px-2 py-1 text-xs text-rose-300 hover:bg-rose-950 disabled:opacity-40"
              disabled={!editable || !changes.length}
              onClick={() => editorRef.current?.rejectAll()}
            >
              Reject all
            </button>
          </div>
        </section>

        <section>
          {sectionTitle("Content")}
          <div className="flex flex-wrap gap-2">
            {(Object.keys(SAMPLES) as SampleId[]).map((id) => (
              <button
                key={id}
                type="button"
                className={controlBtn()}
                onClick={() => loadSample(id)}
              >
                {SAMPLES[id].label}
              </button>
            ))}
          </div>
          <p className="mt-2 text-xs text-zinc-500">
            {stats.words.toLocaleString()} words ·{" "}
            {stats.chars.toLocaleString()} characters
            {` · ${pageCount} page${pageCount === 1 ? "" : "s"}`}
          </p>
        </section>

        <section>
          {sectionTitle("DOCX")}
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              className={controlBtn()}
              onClick={() => fileRef.current?.click()}
            >
              Open .docx
            </button>
            <button type="button" className={controlBtn()} onClick={() => void saveDocx()}>
              Save .docx
            </button>
            <input
              ref={fileRef}
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              className="hidden"
              onChange={(e) => void openDocx(e.target.files?.[0])}
            />
          </div>
          <p className="mt-2 text-xs text-zinc-600">
            Tracked changes are saved as Word revisions (w:ins / w:del / w:rPrChange).
          </p>
        </section>

        <section>
          {sectionTitle("Handle methods")}
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              className={controlBtn()}
              onClick={() =>
                setOutput({
                  label: "getPlainText() (accepted)",
                  text: editorRef.current?.getPlainText() ?? "",
                })
              }
            >
              getPlainText()
            </button>
            <button
              type="button"
              className={controlBtn()}
              onClick={() =>
                setOutput({
                  label: "getPlainText({ mode: \"original\" })",
                  text: editorRef.current?.getPlainText({ mode: "original" }) ?? "",
                })
              }
            >
              original
            </button>
            <button
              type="button"
              className={controlBtn()}
              onClick={() =>
                setOutput({
                  label: "getJson()",
                  text: JSON.stringify(editorRef.current?.getJson(), null, 2),
                })
              }
            >
              getJson()
            </button>
            <button
              type="button"
              className={controlBtn()}
              onClick={() => editorRef.current?.focus()}
            >
              focus()
            </button>
          </div>
          {output && (
            <div className="mt-3">
              <p className="mb-1 text-xs text-zinc-500">{output.label}</p>
              <pre className="max-h-64 overflow-auto whitespace-pre-wrap rounded-md border border-zinc-800 bg-zinc-900 p-2 text-[11px] text-zinc-300">
                {output.text || "(empty)"}
              </pre>
            </div>
          )}
        </section>

        <section>
          {sectionTitle("Events")}
          {log.length === 0 ? (
            <p className="text-xs text-zinc-600">No events yet.</p>
          ) : (
            <ul className="space-y-1 text-xs text-zinc-400">
              {log.map((entry) => (
                <li key={entry.id}>
                  <span className="text-zinc-600">{entry.at}</span>{" "}
                  {entry.message}
                </li>
              ))}
            </ul>
          )}
        </section>
      </aside>
    </div>
  );
}
