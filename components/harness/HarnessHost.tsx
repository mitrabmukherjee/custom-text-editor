"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import {
  isEditorMessage,
  type HostToEditorMessage,
  type PlainTextMode,
} from "@/lib/bridge/protocol";
import type { DocJSON } from "@/lib/pm/json";
import { SAMPLE_DEPOSITION } from "@/lib/samples";

const REQUEST_TIMEOUT_MS = 15000;

const AUTHORS = [
  { id: "typist-1", name: "Typist" },
  { id: "proof-1", name: "Proofreader" },
  { id: "qa-1", name: "QA reviewer" },
];

type LogEntry = { id: number; at: string; message: string };
type Pending = {
  resolve: (value: unknown) => void;
  reject: (err: Error) => void;
  timer: ReturnType<typeof setTimeout>;
};

function controlBtn(disabled = false) {
  return `rounded-md border border-zinc-700 px-2.5 py-1.5 text-xs font-medium text-zinc-300 ${
    disabled ? "cursor-not-allowed opacity-40" : "hover:bg-zinc-800"
  }`;
}

function statusPill(label: string, on: boolean) {
  return (
    <span
      className={`rounded-full px-2 py-0.5 text-[11px] font-medium ${
        on ? "bg-emerald-950 text-emerald-300" : "bg-zinc-800 text-zinc-500"
      }`}
    >
      {label}
    </span>
  );
}

function sectionTitle(text: string) {
  return (
    <h2 className="mb-2 text-[11px] font-semibold uppercase tracking-wide text-zinc-500">
      {text}
    </h2>
  );
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

/**
 * Plays the role av-tracking-project will play: owns the document, frames
 * /embed, and talks to it only through postMessage.
 */
export function HarnessHost() {
  const iframeRef = useRef<HTMLIFrameElement>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const pending = useRef(new Map<string, Pending>());
  const logId = useRef(0);
  const draftRef = useRef(SAMPLE_DEPOSITION);
  const editableRef = useRef(true);
  const [draft, setDraft] = useState(SAMPLE_DEPOSITION);
  const [editable, setEditable] = useState(true);
  const [tracking, setTracking] = useState(false);
  const [authorId, setAuthorId] = useState(AUTHORS[0].id);
  const [plainMode, setPlainMode] = useState<PlainTextMode>("accepted");
  const [ready, setReady] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const [changeEvents, setChangeEvents] = useState(0);
  const [stats, setStats] = useState({ pageCount: 1, changeCount: 0 });
  const [received, setReceived] = useState<{ label: string; text: string } | null>(null);
  const [log, setLog] = useState<LogEntry[]>([]);

  useEffect(() => {
    draftRef.current = draft;
    editableRef.current = editable;
  }, [draft, editable]);

  const pushLog = useCallback((message: string) => {
    logId.current += 1;
    const entry = {
      id: logId.current,
      at: new Date().toLocaleTimeString(),
      message,
    };
    setLog((prev) => [entry, ...prev].slice(0, 40));
  }, []);

  const send = useCallback(
    (msg: HostToEditorMessage, transfer?: Transferable[]) => {
      const win = iframeRef.current?.contentWindow;
      if (!win) return;
      win.postMessage(msg, window.location.origin, transfer);
      pushLog(`→ ${msg.type}`);
    },
    [pushLog],
  );

  const request = useCallback(
    <T,>(build: (requestId: string) => HostToEditorMessage, label: string) =>
      new Promise<T>((resolve, reject) => {
        const requestId = crypto.randomUUID();
        const timer = setTimeout(() => {
          pending.current.delete(requestId);
          reject(new Error(`Timed out waiting for ${label}`));
        }, REQUEST_TIMEOUT_MS);
        pending.current.set(requestId, {
          resolve: resolve as (value: unknown) => void,
          reject,
          timer,
        });
        send(build(requestId));
      }),
    [send],
  );

  useEffect(() => {
    const requests = pending.current;
    const settle = (requestId: string, value: unknown) => {
      const req = requests.get(requestId);
      if (!req) return;
      clearTimeout(req.timer);
      requests.delete(requestId);
      req.resolve(value);
    };
    const onMessage = (event: MessageEvent) => {
      if (event.source !== iframeRef.current?.contentWindow) return;
      if (event.origin !== window.location.origin) return;
      const msg = event.data;
      if (!isEditorMessage(msg)) return;

      switch (msg.type) {
        case "te:ready":
          pushLog("← te:ready");
          setReady(true);
          send({
            type: "te:load",
            plainText: draftRef.current,
            editable: editableRef.current,
          });
          break;
        case "te:loaded":
          pushLog("← te:loaded");
          setLoaded(true);
          break;
        case "te:change":
          setChangeEvents((n) => n + 1);
          setStats({ pageCount: msg.pageCount, changeCount: msg.changeCount });
          break;
        case "te:layout":
          pushLog(`← te:layout ${msg.pageCount} page(s)`);
          setStats((s) => ({ ...s, pageCount: msg.pageCount }));
          break;
        case "te:plain":
          pushLog(`← te:plain (${msg.text.length} chars)`);
          settle(msg.requestId, msg.text);
          break;
        case "te:json":
          pushLog("← te:json");
          settle(msg.requestId, msg.json);
          break;
        case "te:docx":
          pushLog(`← te:docx (${msg.buffer.byteLength.toLocaleString()} bytes)`);
          settle(msg.requestId, msg.buffer);
          break;
        case "te:action":
          pushLog(`← te:action ${msg.name}`);
          break;
        case "te:error": {
          pushLog(`← te:error ${msg.message}`);
          const req = msg.requestId ? requests.get(msg.requestId) : undefined;
          if (req && msg.requestId) {
            clearTimeout(req.timer);
            requests.delete(msg.requestId);
            req.reject(new Error(msg.message));
          }
          break;
        }
      }
    };
    window.addEventListener("message", onMessage);
    return () => {
      window.removeEventListener("message", onMessage);
      for (const req of requests.values()) clearTimeout(req.timer);
      requests.clear();
    };
  }, [pushLog, send]);

  const run = async (fn: () => Promise<void>) => {
    try {
      await fn();
    } catch (e) {
      pushLog(e instanceof Error ? e.message : "Request failed");
    }
  };

  const getPlain = () =>
    run(async () => {
      const text = await request<string>(
        (requestId) => ({ type: "te:get-plain", requestId, mode: plainMode }),
        "te:plain",
      );
      setReceived({ label: `te:plain (${plainMode})`, text });
    });

  const getJson = () =>
    run(async () => {
      const json = await request<DocJSON>(
        (requestId) => ({ type: "te:get-json", requestId }),
        "te:json",
      );
      setReceived({ label: "te:json", text: JSON.stringify(json, null, 2) });
    });

  const getDocx = () =>
    run(async () => {
      const buffer = await request<ArrayBuffer>(
        (requestId) => ({ type: "te:get-docx", requestId }),
        "te:docx",
      );
      downloadBuffer(buffer, "transcript.docx");
    });

  const uploadDocx = async (file: File | undefined) => {
    if (!file) return;
    const buffer = await file.arrayBuffer();
    setLoaded(false);
    send({ type: "te:load-docx", buffer, editable }, [buffer]);
    if (fileRef.current) fileRef.current.value = "";
  };

  const changeAuthor = (id: string) => {
    setAuthorId(id);
    const author = AUTHORS.find((a) => a.id === id)!;
    send({ type: "te:set-author", id: author.id, name: author.name });
  };

  return (
    <div className="flex min-h-0 flex-1">
      <main className="min-h-0 min-w-0 flex-1 bg-zinc-900 p-3">
        <iframe
          ref={iframeRef}
          src="/embed"
          title="Transcript editor"
          className="h-full w-full rounded-md border border-zinc-800 bg-zinc-950"
        />
      </main>

      <aside className="flex w-96 shrink-0 flex-col gap-5 overflow-y-auto border-l border-zinc-800 bg-zinc-950 p-4">
        <section className="flex flex-wrap items-center gap-2">
          {statusPill("ready", ready)}
          {statusPill("loaded", loaded)}
          <span className="text-xs text-zinc-500">
            {stats.pageCount} page{stats.pageCount === 1 ? "" : "s"} · {stats.changeCount}{" "}
            tracked · {changeEvents} te:change
          </span>
        </section>

        <section>
          {sectionTitle("Host document")}
          <textarea
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            rows={7}
            className="w-full rounded-md border border-zinc-800 bg-zinc-900 p-2 font-mono text-xs text-zinc-200"
          />
          <div className="mt-2 flex flex-wrap gap-2">
            <button
              type="button"
              className={controlBtn(!ready)}
              disabled={!ready}
              onClick={() => send({ type: "te:load", plainText: draft, editable })}
            >
              te:load
            </button>
            <button
              type="button"
              className={controlBtn(!loaded)}
              disabled={!loaded}
              onClick={() =>
                send({
                  type: "te:load-template",
                  styles: [
                    {
                      id: "DepoQ",
                      name: "Deposition Q",
                      type: "paragraph",
                      format: { fontSize: 12, spaceAfter: 6 },
                      gallery: true,
                    },
                  ],
                })
              }
            >
              te:load-template
            </button>
            <button
              type="button"
              className={controlBtn(!ready)}
              disabled={!ready}
              onClick={() => send({ type: "te:set-plain", plainText: draft })}
            >
              te:set-plain
            </button>
            <button
              type="button"
              className={controlBtn(!loaded)}
              disabled={!loaded}
              onClick={() => send({ type: "te:focus" })}
            >
              te:focus
            </button>
          </div>
          <label className="mt-3 flex items-center gap-2 text-sm text-zinc-300">
            <input
              type="checkbox"
              checked={!editable}
              disabled={!ready}
              onChange={(e) => {
                const next = !e.target.checked;
                setEditable(next);
                send({ type: "te:set-editable", editable: next });
              }}
            />
            Read-only (te:set-editable)
          </label>
        </section>

        <section>
          {sectionTitle("DOCX")}
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              className={controlBtn(!ready)}
              disabled={!ready}
              onClick={() => fileRef.current?.click()}
            >
              Upload → te:load-docx
            </button>
            <button
              type="button"
              className={controlBtn(!loaded)}
              disabled={!loaded}
              onClick={() => void getDocx()}
            >
              te:get-docx → download
            </button>
            <input
              ref={fileRef}
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              className="hidden"
              onChange={(e) => void uploadDocx(e.target.files?.[0])}
            />
          </div>
        </section>

        <section>
          {sectionTitle("Review")}
          <label className="flex items-center gap-2 text-sm text-zinc-300">
            <input
              type="checkbox"
              checked={tracking}
              disabled={!ready}
              onChange={(e) => {
                setTracking(e.target.checked);
                send({ type: "te:set-track-changes", enabled: e.target.checked });
              }}
            />
            Track changes (te:set-track-changes)
          </label>
          <label className="mt-2 flex items-center gap-2 text-sm text-zinc-300">
            Author
            <select
              value={authorId}
              disabled={!ready}
              onChange={(e) => changeAuthor(e.target.value)}
              className="rounded-md border border-zinc-700 bg-zinc-900 px-2 py-1 text-xs text-zinc-200"
            >
              {AUTHORS.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.name}
                </option>
              ))}
            </select>
          </label>
        </section>

        <section>
          {sectionTitle("Read back")}
          <div className="flex flex-wrap items-center gap-2">
            <select
              value={plainMode}
              onChange={(e) => setPlainMode(e.target.value as PlainTextMode)}
              className="rounded-md border border-zinc-700 bg-zinc-900 px-2 py-1 text-xs text-zinc-200"
            >
              <option value="accepted">accepted</option>
              <option value="original">original</option>
            </select>
            <button
              type="button"
              className={controlBtn(!loaded)}
              disabled={!loaded}
              onClick={() => void getPlain()}
            >
              te:get-plain
            </button>
            <button
              type="button"
              className={controlBtn(!loaded)}
              disabled={!loaded}
              onClick={() => void getJson()}
            >
              te:get-json
            </button>
          </div>
          {received && (
            <div className="mt-3">
              <p className="mb-1 text-xs text-zinc-500">{received.label}</p>
              <pre className="max-h-56 overflow-auto whitespace-pre-wrap rounded-md border border-zinc-800 bg-zinc-900 p-2 text-[11px] text-zinc-300">
                {received.text || "(empty)"}
              </pre>
            </div>
          )}
        </section>

        <section>
          {sectionTitle("Messages")}
          {log.length === 0 ? (
            <p className="text-xs text-zinc-600">Waiting for te:ready…</p>
          ) : (
            <ul className="space-y-1 font-mono text-[11px] text-zinc-400">
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
