/*
 * Swatch name -> colour value mapping, modelled on InDesign's colorInfo():
 *   { name, space: "RGB"|"CMYK"|"LAB"|"GRAY", value: [...], hex }
 * where CMYK values are 0-100, RGB 0-255, and hex is the same naive conversion the InDesign script used.
 *
 * Resolution order for a swatch name:
 *   1. overrides: a { "Swatch Name": "#rrggbb" | {c,m,y,k} | {r,g,b} | [r,g,b] } map (e.g. brand colours)
 *   2. the document's own swatches, read from DesignCraft's swatch.save (.ase) export
 *   3. InDesign's default swatch library (the [Paper]/[Black] built-ins and the stock CMYK swatches)
 *   4. the name itself: "C=15 M=100 Y=90 K=10", "R=10 G=20 B=30", "#ff8800"
 * A name that resolves nowhere is returned as { name } only.
 */

const clamp255 = (n) => Math.max(0, Math.min(255, Math.round(n)));
const hex2 = (n) => clamp255(n).toString(16).padStart(2, "0");
const round = (n, p = 2) => Math.round(n * 10 ** p) / 10 ** p;

export function cmykToHex(c, m, y, k) {
  const kk = 1 - k / 100;
  return "#" + hex2(255 * (1 - c / 100) * kk) + hex2(255 * (1 - m / 100) * kk) + hex2(255 * (1 - y / 100) * kk);
}

export const rgbToHex = (r, g, b) => "#" + hex2(r) + hex2(g) + hex2(b);

function labToHex(L, a, b) {
  const fy = (L + 16) / 116;
  const fx = fy + a / 500;
  const fz = fy - b / 200;
  const f = (t) => (t ** 3 > 0.008856 ? t ** 3 : (t - 16 / 116) / 7.787);
  const X = 0.95047 * f(fx);
  const Y = f(fy);
  const Z = 1.08883 * f(fz);
  const lin = [3.2406 * X - 1.5372 * Y - 0.4986 * Z, -0.9689 * X + 1.8758 * Y + 0.0415 * Z, 0.0557 * X - 0.204 * Y + 1.057 * Z];
  const gamma = (v) => 255 * (v <= 0.0031308 ? 12.92 * v : 1.055 * v ** (1 / 2.4) - 0.055);
  return rgbToHex(gamma(lin[0]), gamma(lin[1]), gamma(lin[2]));
}

export function makeColor(name, space, value) {
  const v = value.map((n) => round(n, 2));
  const out = { name, space, value: v };
  if (space === "RGB") out.hex = rgbToHex(v[0], v[1], v[2]);
  else if (space === "CMYK") out.hex = cmykToHex(v[0], v[1], v[2], v[3]);
  else if (space === "LAB") out.hex = labToHex(v[0], v[1], v[2]);
  else if (space === "GRAY") out.hex = cmykToHex(0, 0, 0, v[0]);
  return out;
}

// InDesign's default swatch library, by name ("C=.. M=.. Y=.. K=..").
// Built-in swatches, as InDesign defines them.
const BUILTINS = {
  "[Paper]": ["CMYK", [0, 0, 0, 0]],
  "[Black]": ["CMYK", [0, 0, 0, 100]],
  "[Registration]": ["CMYK", [100, 100, 100, 100]]
};

const DEFAULT_CMYK = [
  [100, 0, 0, 0], [0, 100, 0, 0], [0, 0, 100, 0], [15, 100, 90, 10], [0, 90, 85, 0], [0, 80, 95, 0],
  [0, 50, 100, 0], [0, 35, 85, 0], [5, 0, 90, 0], [20, 0, 100, 0], [50, 0, 100, 0], [75, 0, 100, 0],
  [85, 10, 100, 10], [90, 30, 95, 30], [75, 0, 75, 0], [80, 10, 45, 0], [70, 15, 0, 0], [85, 50, 0, 0],
  [100, 95, 5, 0], [100, 100, 25, 25], [75, 100, 0, 0], [50, 100, 0, 0], [35, 100, 35, 10], [10, 100, 50, 0],
  [0, 95, 20, 0], [25, 25, 40, 0], [40, 45, 50, 5], [50, 50, 60, 25], [55, 60, 65, 40], [25, 40, 65, 0],
  [30, 50, 75, 10], [35, 60, 80, 25], [40, 65, 90, 35], [40, 70, 100, 50], [50, 70, 80, 70]
];

const cmykName = ([c, m, y, k]) => `C=${c} M=${m} Y=${y} K=${k}`;

// Colours encoded in the name itself, as InDesign names unnamed/process swatches.
export function parseSwatchName(name) {
  let m = /^\s*C\s*=\s*([\d.]+)\s+M\s*=\s*([\d.]+)\s+Y\s*=\s*([\d.]+)\s+K\s*=\s*([\d.]+)\s*$/i.exec(name);
  if (m) return makeColor(name, "CMYK", m.slice(1, 5).map(Number));
  m = /^\s*R\s*=\s*([\d.]+)\s+G\s*=\s*([\d.]+)\s+B\s*=\s*([\d.]+)\s*$/i.exec(name);
  if (m) return makeColor(name, "RGB", m.slice(1, 4).map(Number));
  m = /^\s*#?([0-9a-f]{6})\s*$/i.exec(name);
  if (m && /^#/.test(name.trim())) {
    const n = parseInt(m[1], 16);
    return makeColor(name, "RGB", [(n >> 16) & 255, (n >> 8) & 255, n & 255]);
  }
  return null;
}

// Reads an Adobe Swatch Exchange file: [{ name, space, value }] (CMYK as 0-100, RGB as 0-255, Gray as 0-100 K).
export function parseAse(buf) {
  if (buf.length < 12 || buf.toString("latin1", 0, 4) !== "ASEF") return [];
  const count = buf.readUInt32BE(8);
  const out = [];
  let off = 12;
  for (let i = 0; i < count && off + 6 <= buf.length; i++) {
    const type = buf.readUInt16BE(off);
    const len = buf.readUInt32BE(off + 2);
    const start = off + 6;
    off = start + len;
    if (type !== 0x0001) continue;
    const nameLen = buf.readUInt16BE(start);
    let name = "";
    for (let c = 0; c < nameLen - 1; c++) name += String.fromCharCode(buf.readUInt16BE(start + 2 + c * 2));
    let p = start + 2 + nameLen * 2;
    const model = buf.toString("latin1", p, p + 4).trim().toUpperCase();
    p += 4;
    const floats = (n) => Array.from({ length: n }, (_, j) => buf.readFloatBE(p + j * 4));
    if (model === "CMYK") out.push({ name, space: "CMYK", value: floats(4).map((v) => v * 100) });
    else if (model === "RGB") out.push({ name, space: "RGB", value: floats(3).map((v) => v * 255) });
    else if (model === "GRAY") out.push({ name, space: "GRAY", value: [(1 - floats(1)[0]) * 100] });
    else if (model === "LAB") {
      const [l, a, b] = floats(3);
      out.push({ name, space: "LAB", value: [l * 100, a, b] });
    }
  }
  return out;
}

function fromOverride(name, v) {
  if (typeof v === "string") {
    const n = parseInt(v.replace(/^#/, ""), 16);
    return makeColor(name, "RGB", [(n >> 16) & 255, (n >> 8) & 255, n & 255]);
  }
  if (Array.isArray(v)) return makeColor(name, "RGB", v);
  if ("c" in v) return makeColor(name, "CMYK", [v.c, v.m, v.y, v.k]);
  return makeColor(name, "RGB", [v.r, v.g, v.b]);
}

/**
 * dc: the control-channel client from export-regions.mjs (needs exec()).
 * overrides: optional { swatchName: colour } map that wins over everything else.
 */
export async function loadSwatchMap(dc, overrides = {}) {
  const map = new Map();
  for (const [name, [space, value]] of Object.entries(BUILTINS)) map.set(name, makeColor(name, space, value));
  for (const cmyk of DEFAULT_CMYK) map.set(cmykName(cmyk), makeColor(cmykName(cmyk), "CMYK", cmyk));
  try {
    const saved = await dc.exec("swatch.save", {});
    for (const s of parseAse(Buffer.from(saved.base64, "base64"))) map.set(s.name, makeColor(s.name, s.space, s.value));
  } catch {
    // Older builds or documents without colour swatches: the defaults and name parsing still apply.
  }
  for (const [name, v] of Object.entries(overrides)) map.set(name, fromOverride(name, v));

  return function colorFor(name) {
    if (!name || name === "[None]") return null;
    return map.get(name) || parseSwatchName(name) || { name };
  };
}
