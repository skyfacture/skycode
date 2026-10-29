#!/usr/bin/env node
/*
 * DesignCraft port of export-regions.jsx. DesignCraft has no UXP/ExtendScript host; it is scripted
 * over its JSON control channel, so this is a Node client for a running app:
 *
 *   designcraft --control 7979                       (start DesignCraft, open the scaffold document)
 * node export-regions.mjs [--port 7979] [--out file.regions.json] [--layer DYNAMIC_LAYER] [--swatches overrides.json]
 *
 * Setup in the scaffold document (one page per template: cover, content, back):
 *  - Put one frame per dynamic region on the layer named by --layer (default DYNAMIC_LAYER).
 *  - Name each frame (Object name, or object.label): cover: productDescription, productPartNumber,
 *    productHero; content: cell.1 ... cell.4 (reading order).
 *  - Name/label one frame per page "template:cover", "template:content" or "template:back"
 *    (any layer). It is exported as "templateFrame" with its geometry and fill/stroke swatch names.
 *    A single-page document named Cover/Content/Back falls back to its title.
 *  - For text regions, apply the paragraph/character style and put sample text in the frame.
 *
 * Differences from the InDesign script: DesignCraft cannot list labels, so frames are matched by
 * object name, plus object labels for the known region names (KNOWN_LABELS). Parent-page items are
 * not exposed by document.inspect: override them onto the page (layout.overrideParentItems) first.
 * Swatch colours are resolved to values by swatch-map.mjs (document swatches via swatch.save, then
 * InDesign's defaults, then the name). --swatches file.json maps names to "#rrggbb", {c,m,y,k} or [r,g,b].
 *
 * Output: mm (top-left origin) and PDF points (bottom-left origin), relative to the page's top-left.
 */
import net from "node:net";
import fs from "node:fs";
import { loadSwatchMap } from "./swatch-map.mjs";

const args = process.argv.slice(2);
const opt = (name, def) => {
  const i = args.indexOf("--" + name);
  return i >= 0 ? args[i + 1] : def;
};
const PORT = Number(opt("port", process.env.DESIGNCRAFT_CONTROL_PORT || 7979));
const GUIDE_LAYER = opt("layer", "DYNAMIC_LAYER");
let colorFor = (name) => (name && name !== "[None]" ? { name } : null);
const PT_TO_MM = 25.4 / 72;
const MM_TO_PT = 72 / 25.4;
const templateRe = /^\s*template\s*[:.=]\s*(\S+)\s*$/i;
const KNOWN_LABELS = [
  "productDescription", "productPartNumber", "productHero",
  ...Array.from({ length: 24 }, (_, i) => "cell." + (i + 1))
];

const round = (n, places) => {
  const f = 10 ** places;
  return Math.round(n * f) / f;
};

function connect() {
  return new Promise((resolve, reject) => {
    const sock = net.connect(PORT, "127.0.0.1");
    let buf = "";
    let nextId = 1;
    const pending = new Map();
    sock.setEncoding("utf8");
    sock.on("data", (d) => {
      buf += d;
      let nl;
      while ((nl = buf.indexOf("\n")) >= 0) {
        const line = buf.slice(0, nl).trim();
        buf = buf.slice(nl + 1);
        if (!line) continue;
        const msg = JSON.parse(line);
        const p = pending.get(msg.id);
        if (!p) continue;
        pending.delete(msg.id);
        msg.ok ? p.resolve(msg.result) : p.reject(new Error(msg.error || "request failed"));
      }
    });
    sock.on("error", (e) =>
      reject(new Error(`cannot connect to DesignCraft on 127.0.0.1:${PORT} (start it with: designcraft --control ${PORT}): ${e.message}`))
    );
    sock.on("connect", () =>
      resolve({
        call(method, params) {
          const id = nextId++;
          return new Promise((res, rej) => {
            pending.set(id, { resolve: res, reject: rej });
            sock.write(JSON.stringify({ id, method, params }) + "\n");
          });
        },
        exec(command, params = {}) {
          return this.call("engine.execute", { command, params });
        },
        close: () => sock.end()
      })
    );
  });
}

function leadingPt(chars, para) {
  const l = chars.leading;
  if (l && l.kind === "points") return round(l.value, 2);
  return round(chars.size * (para.autoLeading ?? 1.2), 2);
}

async function textStyle(dc, item) {
  await dc.exec("selection.set", { ids: [item.id], content: true });
  const a = await dc.exec("type.selectionAttrs", {});
  const c = a.chars;
  const p = a.para;
  return {
    paragraphStyle: a.paragraphStyle,
    fontFamily: c.fontFamily,
    fontStyle: c.fontStyle,
    // DesignCraft doesn't expose PostScript names; derive the usual "Family-Style" form.
    postscriptName: (c.fontFamily + "-" + c.fontStyle).replace(/\s+/g, ""),
    sizePt: round(c.size, 2),
    leadingPt: leadingPt(c, p),
    trackingThousandthsEm: c.tracking,
    align: p.align === "justify" || p.align.startsWith("justify") ? "justify" : p.align,
    allCaps: c.capitalization === "allCaps",
    color: colorFor(c.fill),
    colorTintPercent: round((c.fillTint ?? 1) * 100, 1)
  };
}

async function regionFor(dc, item, page, heightMm) {
  const [x0, y0, x1, y1] = item.bounds;
  const x = (x0 - page.bounds[0]) * PT_TO_MM;
  const y = (y0 - page.bounds[1]) * PT_TO_MM;
  const w = (x1 - x0) * PT_TO_MM;
  const h = (y1 - y0) * PT_TO_MM;
  const region = {
    x: round(x, 3),
    y: round(y, 3),
    width: round(w, 3),
    height: round(h, 3),
    pdf: {
      x: round(x * MM_TO_PT, 2),
      y: round((heightMm - y - h) * MM_TO_PT, 2),
      width: round(w * MM_TO_PT, 2),
      height: round(h * MM_TO_PT, 2)
    }
  };
  if (item.kind === "<text frame>" && item.story != null) {
    try {
      const story = await dc.exec("story.get", { story: item.story });
      if (story.text) {
        region.text = await textStyle(dc, item);
        region.sampleText = story.text.substring(0, 80);
      }
      const opts = await dc.exec("object.textFrameOptions", { ids: [item.id] }).catch(() => null);
      if (opts && opts.inset !== undefined) region.insetsMm = [].concat(opts.inset).map((v) => round(v * PT_TO_MM, 3));
      if (opts && opts.verticalJustification) region.verticalAlign = opts.verticalJustification;
    } catch {
      // Style details are best effort, as in the InDesign script.
    }
  }
  return region;
}

async function templateFrameInfo(dc, item, page, heightMm) {
  const info = await regionFor(dc, item, page, heightMm);
  info.type = item.kind.replace(/[<>]/g, "");
  info.fill = colorFor(item.fill);
  const weight = item.stroke ? item.stroke.weight : 0;
  info.stroke = weight > 0 && item.stroke ? colorFor(item.stroke.swatch) : null;
  info.strokeWeightPt = weight;
  info.printing = !item.hidden;
  return info;
}

async function main() {
  const dc = await connect();
  try {
    const doc = await dc.call("document.inspect", {});
    const items = [];
    const pages = [];
    for (const sp of doc.spreads) {
      for (const pg of sp.pages) pages.push(pg);
      for (const it of sp.items) items.push({ ...it, spread: sp.index });
    }
    if (!pages.length) throw new Error("Open the brochure scaffold first.");

    const swatchFile = opt("swatches", null);
    colorFor = await loadSwatchMap(dc, swatchFile ? JSON.parse(fs.readFileSync(swatchFile, "utf8")) : {});

    const layerNames = Object.fromEntries(doc.layers.map((l) => [l.id, l.name]));
    if (!doc.layers.some((l) => l.name === GUIDE_LAYER)) throw new Error(`Layer "${GUIDE_LAYER}" not found.`);

    // Labels cannot be listed, so look up the known ones and treat them as names.
    const labelOf = {};
    const labels = [...KNOWN_LABELS, ...["cover", "content", "back"].map((t) => "template:" + t)];
    for (const label of labels) {
      const r = await dc.exec("object.findByLabel", { label });
      for (const id of r.ids || []) labelOf[id] = labelOf[id] || label;
    }
    for (const it of items) it.tag = it.name || labelOf[it.id] || "";

    const title = String(doc.title || "document");
    const docTemplate = title.replace(/\.(designcraft|idml|indd)$/i, "").toLowerCase();
    const result = {
      source: title,
      exportedAt: new Date().toString(),
      units: { mm: "top-left origin, relative to page", pdf: "points, bottom-left origin, relative to page" },
      pages: []
    };
    const warnings = [];
    const entries = new Map();
    const pageOf = (it) => {
      const cx = (it.bounds[0] + it.bounds[2]) / 2;
      const cy = (it.bounds[1] + it.bounds[3]) / 2;
      return pages.find(
        (p, i) => doc.spreads.some((s) => s.pages.includes(p) && s.index === it.spread) &&
          cx >= p.bounds[0] && cx <= p.bounds[2] && cy >= p.bounds[1] && cy <= p.bounds[3]
      );
    };

    pages.forEach((p, i) => {
      const widthMm = round((p.bounds[2] - p.bounds[0]) * PT_TO_MM, 3);
      const heightMm = round((p.bounds[3] - p.bounds[1]) * PT_TO_MM, 3);
      const entry = {
        pageIndex: i,
        pageName: p.name,
        template: "page-" + (i + 1),
        widthMm,
        heightMm,
        widthPt: round(widthMm * MM_TO_PT, 2),
        heightPt: round(heightMm * MM_TO_PT, 2),
        regions: {}
      };
      result.pages.push(entry);
      entries.set(p, entry);
    });

    for (const it of items) {
      const m = templateRe.exec(it.tag);
      const page = pageOf(it);
      if (!m || !page) continue;
      const e = entries.get(page);
      e.template = m[1].toLowerCase();
      e.templateFrom = "label";
      e.templateFrame = await templateFrameInfo(dc, it, page, e.heightMm);
    }
    for (const e of result.pages) {
      if (!e.templateFrom && pages.length === 1 && /^(cover|content|back)$/.test(docTemplate)) {
        e.template = docTemplate;
        e.templateFrom = "file name";
      }
      if (!e.templateFrom) warnings.push(`Page ${e.pageName} has no template label (use e.g. "template:cover")`);
    }

    for (const it of items) {
      if (!it.tag || templateRe.test(it.tag)) continue;
      if (layerNames[it.layer] !== GUIDE_LAYER) continue;
      const page = pageOf(it);
      if (!page) {
        warnings.push(`"${it.tag}" is not on a document page (pasteboard?) - skipped`);
        continue;
      }
      const e = entries.get(page);
      const region = await regionFor(dc, it, page, e.heightMm);
      const dot = it.tag.indexOf(".");
      if (dot < 0) {
        if (e.regions[it.tag]) warnings.push(`Duplicate label ${it.tag} on ${page.name}`);
        e.regions[it.tag] = region;
      } else {
        const parent = it.tag.slice(0, dot);
        const child = it.tag.slice(dot + 1);
        e.regions[parent] = e.regions[parent] || {};
        e.regions[parent].children = e.regions[parent].children || {};
        e.regions[parent].children[child] = region;
      }
    }

    const REQUIRED = { cover: ["productDescription", "productPartNumber", "productHero"], content: ["cell"] };
    for (const pg of result.pages) {
      if (Object.keys(pg.regions).length === 0 && pg.template !== "back") {
        warnings.push("No labelled regions on page " + pg.pageName);
      }
      // Flatten cell.N children into an ordered array for the grid layout.
      if (pg.regions.cell && pg.regions.cell.children) {
        const cells = [];
        for (let c = 1; pg.regions.cell.children[String(c)]; c++) cells.push(pg.regions.cell.children[String(c)]);
        pg.cells = cells;
        delete pg.regions.cell;
        if (pg.template === "content" && cells.length !== 4) {
          warnings.push(`Content page ${pg.pageName} has ${cells.length} cells (expected 4: cell.1 ... cell.4)`);
        }
      }
      for (const req of REQUIRED[pg.template] || []) {
        const present = req === "cell" ? pg.cells && pg.cells.length : pg.regions[req];
        if (!present) warnings.push(`Template ${pg.template} (page ${pg.pageName}) is missing ${req}`);
      }
    }
    if (warnings.length) result.warnings = warnings;

    const target = opt("out", doc.path ? doc.path.replace(/\.[^.\\/]+$/, "") + ".regions.json" : "brochure.regions.json");
    fs.writeFileSync(target, JSON.stringify(result, null, 2), "utf8");
    console.log("Wrote " + target);
    if (warnings.length) console.log("Warnings:\n- " + warnings.join("\n- "));

    // Put the user's selection back (the style read selects each text frame).
    await dc.exec("selection.set", { ids: doc.selection.items.map((s) => s.id ?? s) });
  } finally {
    dc.close();
  }
}

main().catch((e) => {
  console.error("Export failed: " + e.message);
  process.exit(1);
});
