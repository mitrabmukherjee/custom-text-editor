"use client";

import Link from "next/link";
import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from "react";
import {
  TranscriptEditor,
  type Author,
  type TranscriptEditorHandle,
} from "@/components/editor";
import { getAllowedParentOrigins } from "@/lib/bridge/origins";
import {
  isHostMessage,
  type EditorToHostMessage,
} from "@/lib/bridge/protocol";

const CHANGE_DEBOUNCE_MS = 400;

const noopSubscribe = () => () => {};

function allowedOriginSet() {
  return new Set([...getAllowedParentOrigins(), window.location.origin]);
}

/** Best guess at the embedding page's origin, only if it is on the allowlist. */
function detectParentOrigin(allowed: Set<string>): string | null {
  const ancestors = window.location.ancestorOrigins;
  let candidate: string | null = ancestors?.length ? ancestors[0] : null;
  if (!candidate && document.referrer) {
    try {
      candidate = new URL(document.referrer).origin;
    } catch {
      candidate = null;
    }
  }
  return candidate && allowed.has(candidate) ? candidate : null;
}

function errorMessage(e: unknown) {
  return e instanceof Error ? e.message : String(e);
}

export function EmbedBridge() {
  const framed = useSyncExternalStore(
    noopSubscribe,
    () => window.parent !== window,
    () => null,
  );
  const editorRef = useRef<TranscriptEditorHandle | null>(null);
  const parentOriginRef = useRef<string | null>(null);
  const changeTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const pendingDocx = useRef<ArrayBuffer | null>(null);
  const lastPageCount = useRef(0);
  const [doc, setDoc] = useState<{ plainText: string; key: number } | null>(
    null,
  );
  const [editable, setEditable] = useState(true);
  const [author, setAuthor] = useState<Author | undefined>(undefined);
  const [trackChanges, setTrackChanges] = useState(false);

  const post = useCallback((msg: EditorToHostMessage, transfer?: Transferable[]) => {
    const origin = parentOriginRef.current;
    if (!origin || window.parent === window) return;
    window.parent.postMessage(msg, origin, transfer);
  }, []);

  const remount = (plainText: string) =>
    setDoc((prev) => ({ plainText, key: (prev?.key ?? 0) + 1 }));

  const loadDocx = useCallback(
    async (buffer: ArrayBuffer) => {
      try {
        await editorRef.current!.importDocx(buffer);
        if (changeTimer.current) clearTimeout(changeTimer.current);
        post({ type: "te:loaded" });
      } catch (e) {
        post({ type: "te:error", message: `Could not open DOCX: ${errorMessage(e)}` });
      }
    },
    [post],
  );

  useEffect(() => {
    if (!framed) return undefined;
    const allowed = allowedOriginSet();
    parentOriginRef.current = detectParentOrigin(allowed);

    const onMessage = (event: MessageEvent) => {
      if (event.source !== window.parent) return;
      if (!allowed.has(event.origin)) return;
      if (!isHostMessage(event.data)) return;
      parentOriginRef.current = event.origin;
      const msg = event.data;
      const handle = editorRef.current;
      const notReady = (requestId: string) =>
        post({ type: "te:error", requestId, message: "Editor not ready" });

      switch (msg.type) {
        case "te:load":
          pendingDocx.current = null;
          remount(msg.plainText);
          if (msg.editable !== undefined) setEditable(msg.editable);
          break;
        case "te:load-docx":
          if (msg.editable !== undefined) setEditable(msg.editable);
          if (handle) {
            void loadDocx(msg.buffer);
          } else {
            pendingDocx.current = msg.buffer;
            remount("");
          }
          break;
        case "te:get-plain":
          if (!handle) notReady(msg.requestId);
          else {
            post({
              type: "te:plain",
              requestId: msg.requestId,
              text: handle.getPlainText({ mode: msg.mode }),
            });
          }
          break;
        case "te:get-json":
          if (!handle) notReady(msg.requestId);
          else post({ type: "te:json", requestId: msg.requestId, json: handle.getJson() });
          break;
        case "te:get-docx":
          if (!handle) notReady(msg.requestId);
          else {
            handle.exportDocx().then(
              (buffer) =>
                post({ type: "te:docx", requestId: msg.requestId, buffer }, [buffer]),
              (e) =>
                post({ type: "te:error", requestId: msg.requestId, message: errorMessage(e) }),
            );
          }
          break;
        case "te:set-plain":
          if (handle) handle.setPlainText(msg.plainText);
          else remount(msg.plainText);
          break;
        case "te:set-editable":
          setEditable(msg.editable);
          break;
        case "te:set-author":
          setAuthor({ id: msg.id, name: msg.name || msg.id });
          break;
        case "te:set-track-changes":
          setTrackChanges(msg.enabled);
          handle?.setTrackChanges(msg.enabled);
          break;
        case "te:focus":
          handle?.focus();
          break;
        case "te:load-template":
          if (!handle) {
            if (msg.requestId) notReady(msg.requestId);
            else post({ type: "te:error", message: "Editor not ready" });
          } else {
            try {
              handle.loadTemplateStyles(msg.styles);
              post(msg.requestId ? { type: "te:loaded", requestId: msg.requestId } : { type: "te:loaded" });
            } catch (e) {
              post({
                type: "te:error",
                message: errorMessage(e),
                requestId: msg.requestId,
              });
            }
          }
          break;
      }
    };

    window.addEventListener("message", onMessage);
    post({ type: "te:ready" });
    return () => {
      window.removeEventListener("message", onMessage);
      if (changeTimer.current) clearTimeout(changeTimer.current);
    };
  }, [framed, post, loadDocx]);

  const onUpdate = useCallback(() => {
    if (changeTimer.current) clearTimeout(changeTimer.current);
    changeTimer.current = setTimeout(() => {
      const handle = editorRef.current;
      post({
        type: "te:change",
        pageCount: handle?.getPageCount() ?? 1,
        changeCount: handle?.listChanges().length ?? 0,
      });
    }, CHANGE_DEBOUNCE_MS);
  }, [post]);

  const onReady = useCallback(() => {
    const buffer = pendingDocx.current;
    pendingDocx.current = null;
    if (buffer) void loadDocx(buffer);
    else post({ type: "te:loaded" });
  }, [loadDocx, post]);

  if (framed === null) return null;

  if (!framed) {
    return (
      <div className="flex h-screen items-center justify-center p-6">
        <div className="max-w-md text-center text-sm text-zinc-400">
          <p className="mb-2 font-medium text-zinc-200">
            This route is meant to be embedded in an iframe.
          </p>
          <p>
            A host page sends the document with <code>te:load</code> or{" "}
            <code>te:load-docx</code> and reads it back with{" "}
            <code>te:get-plain</code> or <code>te:get-docx</code>. Try it on the{" "}
            <Link href="/harness" className="text-zinc-100 underline">
              iframe harness
            </Link>
            .
          </p>
        </div>
      </div>
    );
  }

  if (!doc) {
    return (
      <div className="flex h-screen items-center justify-center text-sm text-zinc-400">
        Waiting for document…
      </div>
    );
  }

  return (
    <div className="h-screen">
      <TranscriptEditor
        key={doc.key}
        initialContent={doc.plainText}
        editable={editable}
        author={author}
        trackChanges={trackChanges}
        editorRef={editorRef}
        studioActions={{
          onPlayPause: () => post({ type: "te:action", name: "playPause" }),
          onPrint: () => post({ type: "te:action", name: "print" }),
        }}
        onReady={onReady}
        onUpdate={onUpdate}
        onTrackChangesChange={setTrackChanges}
        onLayout={({ pageCount }) => {
          if (pageCount === lastPageCount.current) return;
          lastPageCount.current = pageCount;
          post({ type: "te:layout", pageCount });
        }}
        className="h-full"
      />
    </div>
  );
}
