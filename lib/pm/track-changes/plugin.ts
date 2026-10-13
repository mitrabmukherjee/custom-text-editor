import { Plugin, PluginKey, type Command } from "prosemirror-state";

export type Author = { id: string; name: string };

export type TrackChangesState = {
  enabled: boolean;
  author: Author;
  /** Word's Lock Tracking: tracking can't be turned off until unlocked. */
  locked?: boolean;
};

export const trackChangesKey = new PluginKey<TrackChangesState>("trackChanges");

/** Transaction meta: apply this transaction as-is, never convert it. */
export const SKIP_TRACKING = "trackChanges.skip";

export function trackChangesPlugin(initial: () => TrackChangesState) {
  return new Plugin<TrackChangesState>({
    key: trackChangesKey,
    state: {
      init: () => initial(),
      apply(tr, prev) {
        const meta = tr.getMeta(trackChangesKey) as
          | Partial<TrackChangesState>
          | undefined;
        return meta ? { ...prev, ...meta } : prev;
      },
    },
    props: {
      attributes(state) {
        const enabled = trackChangesKey.getState(state)?.enabled;
        return { "data-track-changes": enabled ? "on" : "off" };
      },
    },
  });
}

export function setTrackChanges(enabled: boolean): Command {
  return (state, dispatch) => {
    if (!enabled && trackChangesKey.getState(state)?.locked) return false;
    dispatch?.(state.tr.setMeta(trackChangesKey, { enabled }));
    return true;
  };
}

/** Word's Ctrl+Shift+E. */
export const toggleTrackChanges: Command = (state, dispatch) => {
  const current = trackChangesKey.getState(state);
  return current ? setTrackChanges(!current.enabled)(state, dispatch) : false;
};

/** Lock tracking on (locking also turns it on), or unlock it. */
export function setTrackingLocked(locked: boolean): Command {
  return (state, dispatch) => {
    dispatch?.(state.tr.setMeta(trackChangesKey, locked ? { locked, enabled: true } : { locked }));
    return true;
  };
}

export function setAuthor(author: Author): Command {
  return (state, dispatch) => {
    dispatch?.(state.tr.setMeta(trackChangesKey, { author }));
    return true;
  };
}
