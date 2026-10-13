export const DEFAULT_PARENT_ORIGINS = [
  "http://localhost:3000",
  "http://localhost:3001",
];

/** Parse a comma-separated origin list; falls back to the local dev origins. */
export function parseAllowedOrigins(raw: string | undefined): string[] {
  const list = String(raw ?? "")
    .split(",")
    .map((s) => s.trim().replace(/\/+$/, ""))
    .filter(Boolean);
  return list.length ? list : DEFAULT_PARENT_ORIGINS;
}

export function getAllowedParentOrigins(): string[] {
  return parseAllowedOrigins(process.env.NEXT_PUBLIC_ALLOWED_PARENT_ORIGINS);
}
