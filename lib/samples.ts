export const SAMPLE_DEPOSITION = `THE COURT: Good morning.

Q. Please state your name for the record.

A. My name is Jane Doe.

Q. Where do you reside?

A. I live in Springfield.`;

export const SAMPLE_LONG = Array.from({ length: 80 }, (_, i) => {
  const n = i + 1;
  return `Q. Question number ${n} on the record?\n\nA. Answer number ${n} with some additional detail for scrolling tests.`;
}).join("\n\n");

export const SAMPLES = {
  deposition: { label: "Sample deposition", text: SAMPLE_DEPOSITION },
  long: { label: "Long document", text: SAMPLE_LONG },
  empty: { label: "Empty", text: "" },
} as const;

export type SampleId = keyof typeof SAMPLES;
