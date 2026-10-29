const PALETTE: ReadonlyArray<readonly [string, number, number, number]> = [
  ["white", 255, 255, 255],
  ["black", 0, 0, 0],
  ["gray", 128, 128, 128],
  ["dark gray", 64, 64, 64],
  ["light gray", 192, 192, 192],
  ["silver", 211, 211, 211],

  ["red", 255, 0, 0],
  ["dark red", 139, 0, 0],
  ["green", 0, 128, 0],
  ["dark green", 0, 100, 0],
  ["blue", 0, 0, 255],
  ["dark blue", 0, 0, 139],
  ["navy", 0, 0, 128],

  ["yellow", 255, 255, 0],
  ["orange", 255, 165, 0],
  ["purple", 128, 0, 128],
  ["pink", 255, 192, 203],
  ["cyan", 0, 255, 255],
  ["teal", 0, 128, 128],
  ["magenta", 255, 0, 255],

  ["light blue", 173, 216, 230],
  ["light green", 144, 238, 144],
  ["light yellow", 255, 255, 224],
  ["light pink", 255, 182, 193],
  ["light cyan", 224, 255, 255],

  ["pale blue", 217, 225, 242],
  ["pale green", 226, 239, 218],
  ["pale yellow", 255, 242, 204],
  ["pale orange", 252, 228, 214],
  ["pale gray", 242, 242, 242],

  ["beige", 245, 245, 220],
  ["ivory", 255, 255, 240],
  ["gold", 255, 215, 0],
  ["coral", 255, 127, 80],
  ["salmon", 250, 128, 114],
  ["olive", 128, 128, 0],
  ["maroon", 128, 0, 0],
  ["brown", 165, 42, 42],
  ["wheat", 245, 222, 179],
  ["lavender", 230, 230, 250],
  ["khaki", 240, 230, 140],
  ["indigo", 75, 0, 130],
];

function hexToRgb(hex: string): [number, number, number] | null {
  const clean = hex.replace(/^#/, "");
  if (clean.length !== 6) return null;
  const n = parseInt(clean, 16);
  if (isNaN(n)) return null;
  return [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff];
}

function rgbDistance(
  a: [number, number, number],
  b: readonly [number, number, number],
): number {
  const dr = a[0] - b[0];
  const dg = a[1] - b[1];
  const db = a[2] - b[2];
  return Math.sqrt(dr * dr + dg * dg + db * db);
}

const MATCH_THRESHOLD = 50;

export function hexToColorName(hex: string): string | null {
  const rgb = hexToRgb(hex);
  if (!rgb) return null;

  let bestName = "";
  let bestDist = Infinity;

  for (const [name, r, g, b] of PALETTE) {
    const dist = rgbDistance(rgb, [r, g, b]);
    if (dist < bestDist) {
      bestDist = dist;
      bestName = name;
    }
  }

  return bestDist <= MATCH_THRESHOLD ? bestName : null;
}

export function formatColorLabel(hex: string): string {
  const name = hexToColorName(hex.toUpperCase());
  return name ?? hex;
}

export function humanizeColorsInText(text: string): string {
  return text.replace(/#([0-9A-Fa-f]{6})\b/g, (match) => {
    const name = hexToColorName(match.toUpperCase());
    return name ?? match;
  });
}
