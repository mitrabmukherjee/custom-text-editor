/** Word's 16 highlight colours (`w:highlight` values). */
export const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: "#ffff00",
  green: "#00ff00",
  cyan: "#00ffff",
  magenta: "#ff00ff",
  blue: "#0000ff",
  red: "#ff0000",
  darkBlue: "#000080",
  darkCyan: "#008080",
  darkGreen: "#008000",
  darkMagenta: "#800080",
  darkRed: "#800000",
  darkYellow: "#808000",
  darkGray: "#808080",
  lightGray: "#c0c0c0",
  black: "#000000",
  white: "#ffffff",
};

/** Word's standard font colours, as shown in the Font Color menu. */
export const FONT_COLORS: { name: string; hex: string }[] = [
  { name: "Dark Red", hex: "#c00000" },
  { name: "Red", hex: "#ff0000" },
  { name: "Orange", hex: "#ffc000" },
  { name: "Yellow", hex: "#ffff00" },
  { name: "Light Green", hex: "#92d050" },
  { name: "Green", hex: "#00b050" },
  { name: "Light Blue", hex: "#00b0f0" },
  { name: "Blue", hex: "#0070c0" },
  { name: "Dark Blue", hex: "#002060" },
  { name: "Purple", hex: "#7030a0" },
];

const NAMES = new Map<string, string>([
  ...FONT_COLORS.map((c) => [c.hex, c.name] as [string, string]),
  ...Object.entries(HIGHLIGHT_COLORS).map(
    ([name, hex]) =>
      [hex, name.replace(/([A-Z])/g, " $1").replace(/^./, (c) => c.toUpperCase())] as [
        string,
        string,
      ],
  ),
]);

/** Normalise "FF0000" / "#f00" / "#FF0000" to "#ff0000"; null for "auto" or junk. */
export function normalizeHex(value: string | null | undefined): string | null {
  if (!value) return null;
  let v = value.trim().replace(/^#/, "").toLowerCase();
  if (/^[0-9a-f]{3}$/.test(v)) v = v.replace(/./g, (c) => c + c);
  return /^[0-9a-f]{6}$/.test(v) ? `#${v}` : null;
}

/** Word's wording for a colour in revision balloons, e.g. "Red" or "Custom Color(RGB(1,2,3))". */
export function colorName(hex: string): string {
  const known = NAMES.get(hex.toLowerCase());
  if (known) return known;
  const [r, g, b] = rgb(hex);
  return `Custom Color(RGB(${r},${g},${b}))`;
}

/** The `w:highlight` name for a colour, or null when it needs `w:shd` instead. */
export function highlightName(hex: string): string | null {
  const v = hex.toLowerCase();
  return Object.entries(HIGHLIGHT_COLORS).find(([, h]) => h === v)?.[0] ?? null;
}

function rgb(hex: string): [number, number, number] {
  const v = parseInt(hex.slice(1), 16);
  return [(v >> 16) & 255, (v >> 8) & 255, v & 255];
}

function luminance(hex: string) {
  const [r, g, b] = rgb(hex).map((c) => {
    const s = c / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/**
 * Colour to draw text with on the dark page: dark colours have their lightness
 * inverted (as Word's dark mode does) so black text stays readable.
 */
export function displayColor(hex: string): string {
  const [r, g, b] = rgb(hex).map((c) => c / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (l >= 0.5) return hex;
  const d = max - min;
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  let h = 0;
  if (d) {
    if (max === r) h = ((g - b) / d) % 6;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h *= 60;
  }
  const lightness = Math.round((1 - l) * 90 + 5);
  return `hsl(${Math.round((h + 360) % 360)} ${Math.round(s * 100)}% ${lightness}%)`;
}

/** Black or white text, whichever reads better on `background`. */
export function textOn(background: string): string {
  return luminance(background) > 0.35 ? "#000000" : "#ffffff";
}
