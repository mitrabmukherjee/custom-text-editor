import { TextSelection, type Command } from "prosemirror-state";
import { schema } from "./schema";

export function insertQaLine(role: "q" | "a"): Command {
  return (state, dispatch) => {
    const node = schema.nodes.qa_line.create({ role });
    if (!dispatch) return true;
    const tr = state.tr.replaceSelectionWith(node);
    const pos = tr.selection.from + 1;
    dispatch(tr.setSelection(TextSelection.near(tr.doc.resolve(pos))).scrollIntoView());
    return true;
  };
}

export function insertSpeakerLine(label = "THE COURT:"): Command {
  return (state, dispatch) => {
    const node = schema.nodes.speaker_line.create({ label });
    if (!dispatch) return true;
    const tr = state.tr.replaceSelectionWith(node);
    const pos = tr.selection.from + 1;
    dispatch(tr.setSelection(TextSelection.near(tr.doc.resolve(pos))).scrollIntoView());
    return true;
  };
}

export function insertTimestamp(time = "00:00:00"): Command {
  return (state, dispatch) => {
    const node = schema.nodes.timestamp.create({ time });
    if (!dispatch) return true;
    dispatch(state.tr.replaceSelectionWith(node).scrollIntoView());
    return true;
  };
}
