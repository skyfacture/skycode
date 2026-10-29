//@target indesign
/*
 * Exports brochure dynamic-region geometry and text styles to <document>.regions.json
 *
 * Setup in InDesign (one scaffold page per template: cover, contentpage1, contentpageN, back):
 *  - Put one frame (rectangle or text frame) per dynamic region on the layer named in GUIDE_LAYER.
 *    Static artwork (logo, QR code, fixed text) stays on normal layers and is NOT exported.
 *  - Name each frame in Window > Utilities > Script Label.
 *      cover:   productDescription, productPartNumber, productHero
 *      contentpage1: cell.1 (first XCON, on the page immediately after Cover)
 *      contentpageN: four cells, e.g. cell.2 ... cell.5 (reading order: left-to-right, top-to-bottom)
 *        Continuation cells may start at any positive number; they are exported in numeric order.
 *  - Label any frame (on any layer) on each page, or on its master/parent page, "template:cover",
 *    "template:contentpage1", "template:contentpageN" or "template:back".
 *    Template names are exported in lowercase. A label on the page wins over its master's. The separator
 *    must be a colon (a dot or = also works). If a single-page document is named
 *    Cover.indd, Contentpage1.indd, ContentpageN.indd or Back.indd, the file name is used
 *    when there is no label. Legacy Content.indd / template:content still supports four cells.
 *  - The template-labelled frame is also exported as "templateFrame" (geometry, fill colour/tint, stroke,
 *    opacity), so a background rectangle's colour can be read from the JSON.
 *  - Region frames can also sit on a master page (on the guide layer); they apply to every page using it.
 *  - For text regions, apply the paragraph/character style you want, and put sample text in the frame.
 *
 * Output: coordinates in mm (top-left origin, as shown in InDesign) and in PDF points (bottom-left
 * origin, ready for pdf-lib drawText/drawImage), both relative to the page's top-left corner.
 */

var GUIDE_LAYER = "DYNAMIC_LAYER";
var MM_TO_PT = 72 / 25.4;

function round(n, places) {
  var f = Math.pow(10, places);
  return Math.round(n * f) / f;
}

// ExtendScript has no JSON object, so serialise by hand.
function toJson(v, indent) {
  indent = indent || "";
  var next = indent + "  ";
  var i, parts, k;
  if (v === null || v === undefined) return "null";
  if (typeof v === "number") return isFinite(v) ? String(v) : "null";
  if (typeof v === "boolean") return v ? "true" : "false";
  if (typeof v === "string") {
    return '"' + v.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\n").replace(/\n/g, "\\n").replace(/\t/g, "\\t") + '"';
  }
  if (v instanceof Array) {
    if (v.length === 0) return "[]";
    var allScalar = true;
    for (i = 0; i < v.length; i++) if (typeof v[i] === "object" && v[i] !== null) allScalar = false;
    parts = [];
    for (i = 0; i < v.length; i++) parts.push(allScalar ? toJson(v[i]) : next + toJson(v[i], next));
    return allScalar ? "[" + parts.join(", ") + "]" : "[\n" + parts.join(",\n") + "\n" + indent + "]";
  }
  parts = [];
  for (k in v) if (v.hasOwnProperty(k)) parts.push(next + toJson(String(k)) + ": " + toJson(v[k], next));
  return parts.length ? "{\n" + parts.join(",\n") + "\n" + indent + "}" : "{}";
}

function hex2(n) {
  var s = Math.max(0, Math.min(255, Math.round(n))).toString(16);
  return s.length === 1 ? "0" + s : s;
}

function colorInfo(c) {
  if (!c || !c.isValid) return null;
  try {
    var v = c.colorValue;
    var out = { name: c.name, space: String(c.space), value: v };
    if (c.space === ColorSpace.RGB) {
      out.hex = "#" + hex2(v[0]) + hex2(v[1]) + hex2(v[2]);
    } else if (c.space === ColorSpace.CMYK) {
      // Naive conversion; use the CMYK values if the exact brand colour matters.
      var k = 1 - v[3] / 100;
      out.hex = "#" + hex2(255 * (1 - v[0] / 100) * k) + hex2(255 * (1 - v[1] / 100) * k) + hex2(255 * (1 - v[2] / 100) * k);
    }
    return out;
  } catch (e) {
    return { name: c.name };
  }
}

function justificationName(j) {
  if (j === Justification.CENTER_ALIGN) return "center";
  if (j === Justification.RIGHT_ALIGN) return "right";
  if (j === Justification.LEFT_ALIGN) return "left";
  if (j === Justification.FULLY_JUSTIFIED) return "justify";
  return String(j);
}

function textStyle(para) {
  var style = {};
  try { style.paragraphStyle = para.appliedParagraphStyle.name; } catch (e1) {}
  try {
    style.fontFamily = para.appliedFont.fontFamily;
    style.fontStyle = para.fontStyle;
    style.postscriptName = para.appliedFont.postscriptName;
    style.fontFile = String(para.appliedFont.location || "");
  } catch (e2) {}
  try {
    style.sizePt = round(para.pointSize, 2);
    style.leadingPt = para.leading === Leading.AUTO ? round(para.pointSize * (para.autoLeading / 100), 2) : round(para.leading, 2);
    style.trackingThousandthsEm = para.tracking;
    style.align = justificationName(para.justification);
    style.allCaps = para.capitalization === Capitalization.ALL_CAPS;
  } catch (e3) {}
  try { style.color = colorInfo(para.fillColor); } catch (e4) {}
  return style;
}

function regionFor(item, pageBounds, pageHeightMm) {
  var gb = item.geometricBounds; // [top, left, bottom, right] in current ruler units (mm)
  var x = gb[1] - pageBounds[1];
  var y = gb[0] - pageBounds[0];
  var w = gb[3] - gb[1];
  var h = gb[2] - gb[0];
  var region = {
    x: round(x, 3),
    y: round(y, 3),
    width: round(w, 3),
    height: round(h, 3),
    pdf: {
      x: round(x * MM_TO_PT, 2),
      y: round((pageHeightMm - y - h) * MM_TO_PT, 2),
      width: round(w * MM_TO_PT, 2),
      height: round(h * MM_TO_PT, 2)
    }
  };
  if (item.constructor.name === "TextFrame") {
    try {
      if (item.paragraphs.length > 0 && item.contents !== "") {
        region.text = textStyle(item.paragraphs[0]);
        region.sampleText = item.contents.substring(0, 80);
      }
      region.insetsMm = item.textFramePreferences.insetSpacing;
      region.verticalAlign = String(item.textFramePreferences.verticalJustification);
    } catch (e) {}
  }
  return region;
}

// The template marker frame often doubles as the page background (e.g. a full-page coloured rectangle).
function templateFrameInfo(item, page, pageHeightMm) {
  var info = regionFor(item, page.bounds, pageHeightMm);
  info.type = item.constructor.name;
  try { info.fill = colorInfo(item.fillColor); } catch (e1) {}
  try { if (item.fillTint >= 0) info.fillTintPercent = item.fillTint; } catch (e2) {}
  try {
    info.stroke = item.strokeWeight > 0 ? colorInfo(item.strokeColor) : null;
    info.strokeWeightPt = item.strokeWeight;
  } catch (e3) {}
  try { info.opacityPercent = item.transparencySettings.blendingSettings.opacity; } catch (e4) {}
  try { info.cornerRadiusMm = item.topLeftCornerRadius; } catch (e5) {}
  info.printing = item.nonprinting !== true;
  return info;
}

function main() {
  if (app.documents.length === 0) {
    alert("Open the brochure scaffold first.");
    return;
  }
  var doc = app.activeDocument;
  var layer = doc.layers.itemByName(GUIDE_LAYER);
  if (!layer.isValid) {
    alert('Layer "' + GUIDE_LAYER + '" not found.');
    return;
  }

  var vp = doc.viewPreferences;
  var oldH = vp.horizontalMeasurementUnits;
  var oldV = vp.verticalMeasurementUnits;
  vp.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
  vp.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;

  var result = {
    source: doc.name,
    exportedAt: new Date().toString(),
    units: { mm: "top-left origin, relative to page", pdf: "points, bottom-left origin, relative to page" },
    pages: []
  };
  var warnings = [];
  var pageMap = {};
  var items = layer.allPageItems;
  var i, j, key;

  for (i = 0; i < doc.pages.length; i++) {
    var p = doc.pages[i];
    var pb = p.bounds;
    var entry = {
      pageIndex: i,
      pageName: p.name,
      template: "page-" + (i + 1),
      widthMm: round(pb[3] - pb[1], 3),
      heightMm: round(pb[2] - pb[0], 3),
      regions: {}
    };
    entry.widthPt = round(entry.widthMm * MM_TO_PT, 2);
    entry.heightPt = round(entry.heightMm * MM_TO_PT, 2);
    result.pages.push(entry);
    pageMap[p.id] = entry;
  }

  var templateRe = /^\s*template\s*[:.=]\s*(\S+)\s*$/i;
  var docTemplate = doc.name.replace(/\.indd$/i, "").toLowerCase();
  var docPages = [];
  var masterOf = {}; // document page id -> the master page it uses
  var masterItems = {}; // master page id -> [{ item, label }]
  var allItems = doc.allPageItems;
  var m, mp, k;

  for (i = 0; i < doc.pages.length; i++) {
    var dp = doc.pages[i];
    docPages.push(dp);
    try {
      var am = dp.appliedMaster;
      if (am && am.isValid && am.pages.length > 0) {
        mp = am.pages.length === 1 ? am.pages[0] : (dp.side === PageSideOptions.LEFT_HAND ? am.pages[0] : am.pages[am.pages.length - 1]);
        masterOf[dp.id] = mp;
      }
    } catch (e) {}
  }

  for (m = 0; m < doc.masterSpreads.length; m++) {
    var mitems = doc.masterSpreads[m].allPageItems;
    for (k = 0; k < mitems.length; k++) {
      var ml = mitems[k].label;
      if (!ml) continue;
      var mpage = null;
      try { mpage = mitems[k].parentPage; } catch (e) {}
      if (!mpage) continue;
      if (!masterItems[mpage.id]) masterItems[mpage.id] = [];
      masterItems[mpage.id].push({ item: mitems[k], label: ml });
    }
  }

  // Template: label on the page, else label on its master page, else the file name.
  for (j = 0; j < allItems.length; j++) {
    var tl = allItems[j].label;
    if (!tl) continue;
    var tm = templateRe.exec(tl);
    if (!tm) continue;
    var tpage = null;
    try { tpage = allItems[j].parentPage; } catch (e) {}
    if (tpage && pageMap[tpage.id]) {
      pageMap[tpage.id].template = tm[1].toLowerCase();
      pageMap[tpage.id].templateFrom = "label";
      pageMap[tpage.id].templateFrame = templateFrameInfo(allItems[j], tpage, pageMap[tpage.id].heightMm);
    }
  }
  for (i = 0; i < docPages.length; i++) {
    var te = pageMap[docPages[i].id];
    var mlist = masterOf[docPages[i].id] ? masterItems[masterOf[docPages[i].id].id] : null;
    if (!te.templateFrom && mlist) {
      for (k = 0; k < mlist.length; k++) {
        var mt = templateRe.exec(mlist[k].label);
        if (mt) {
          te.template = mt[1].toLowerCase();
          te.templateFrom = "master page label";
          te.templateFrame = templateFrameInfo(mlist[k].item, masterOf[docPages[i].id], te.heightMm);
          break;
        }
      }
    }
    // A single-page template document can use its file name instead of a label.
    if (!te.templateFrom && doc.pages.length === 1 && /^(cover|content|contentpage1|contentpagen|back)$/.test(docTemplate)) {
      te.template = docTemplate;
      te.templateFrom = "file name";
    }
    if (!te.templateFrom) {
      warnings.push('Page ' + te.pageName + ' has no template label (use e.g. "template:cover")');
    }
  }

  function addRegion(item, label, page, pageEntry, fromMaster) {
    var dot = label.indexOf(".");
    var existing;
    if (dot < 0) {
      existing = pageEntry.regions[label];
      if (fromMaster && existing && !existing.children) return;
      var region = regionFor(item, page.bounds, pageEntry.heightMm);
      if (existing && existing.children) {
        region.children = existing.children;
      } else if (existing) {
        warnings.push("Duplicate label " + label + " on " + page.name);
      }
      pageEntry.regions[label] = region;
    } else {
      var parent = label.substring(0, dot);
      var child = label.substring(dot + 1);
      if (!pageEntry.regions[parent]) pageEntry.regions[parent] = {};
      if (!pageEntry.regions[parent].children) pageEntry.regions[parent].children = {};
      if (fromMaster && pageEntry.regions[parent].children[child]) return;
      pageEntry.regions[parent].children[child] = regionFor(item, page.bounds, pageEntry.heightMm);
    }
  }

  for (j = 0; j < items.length; j++) {
    var item = items[j];
    var label = item.label;
    if (!label) continue;
    if (templateRe.test(label)) continue;
    var page = null;
    try { page = item.parentPage; } catch (e) {}
    if (!page || !pageMap[page.id]) {
      warnings.push('"' + label + '" is not on a document page (pasteboard?) - skipped');
      continue;
    }
    addRegion(item, label, page, pageMap[page.id], false);
  }

  // Region frames placed on a master page apply to every page that uses it.
  for (i = 0; i < docPages.length; i++) {
    var mlist2 = masterOf[docPages[i].id] ? masterItems[masterOf[docPages[i].id].id] : null;
    if (!mlist2) continue;
    for (k = 0; k < mlist2.length; k++) {
      var mi = mlist2[k];
      if (templateRe.test(mi.label)) continue;
      var onGuideLayer = false;
      try { onGuideLayer = mi.item.itemLayer.name === GUIDE_LAYER; } catch (e) {}
      if (!onGuideLayer) continue;
      addRegion(mi.item, mi.label, masterOf[docPages[i].id], pageMap[docPages[i].id], true);
    }
  }

  vp.horizontalMeasurementUnits = oldH;
  vp.verticalMeasurementUnits = oldV;

  var REQUIRED = {
    cover: ["productDescription", "productPartNumber", "productHero"],
    content: ["cell"],
    contentpage1: ["cell"],
    contentpagen: ["cell"]
  };
  var CELL_COUNTS = { content: 4, contentpage1: 1, contentpagen: 4 };
  for (i = 0; i < result.pages.length; i++) {
    var pg = result.pages[i];
    var n = 0;
    for (key in pg.regions) n++;
    if (n === 0 && pg.template !== "back") warnings.push("No labelled regions on page " + pg.pageName);

    // Labels need not start at 1 (continuation pages normally use cell.2 ... cell.5).
    if (pg.regions.cell && pg.regions.cell.children) {
      var children = pg.regions.cell.children;
      var cellNumbers = [];
      var cells = [];
      var c;
      for (key in children) {
        if (!children.hasOwnProperty(key)) continue;
        if (/^[1-9]\d*$/.test(key)) {
          cellNumbers.push(Number(key));
        } else {
          warnings.push('Invalid cell label "cell.' + key + '" on page ' + pg.pageName + " (expected a positive integer)");
        }
      }
      cellNumbers.sort(function(a, b) { return a - b; });
      for (c = 0; c < cellNumbers.length; c++) cells.push(children[String(cellNumbers[c])]);
      pg.cells = cells;
      delete pg.regions.cell;
      if (pg.template === "contentpage1" && cells.length === 1 && cellNumbers[0] !== 1) {
        warnings.push("Template contentpage1 (page " + pg.pageName + ") must use cell.1");
      }
    }
    var expectedCells = CELL_COUNTS[pg.template];
    if (expectedCells && (!pg.cells || pg.cells.length !== expectedCells)) {
      warnings.push("Template " + pg.template + " (page " + pg.pageName + ") has " + (pg.cells ? pg.cells.length : 0) + " cells (expected " + expectedCells + ")");
    }

    var req = REQUIRED[pg.template];
    var r;
    if (req) {
      for (r = 0; r < req.length; r++) {
        var present = req[r] === "cell" ? pg.cells && pg.cells.length : pg.regions[req[r]];
        if (!present) warnings.push("Template " + pg.template + " (page " + pg.pageName + ") is missing " + req[r]);
      }
    }
  }
  if (warnings.length) result.warnings = warnings;

  var target = null;
  try {
    target = new File(doc.fullName.fsName.replace(/\.indd$/i, "") + ".regions.json");
  } catch (e) {
    target = new File("~/brochure.regions.json").saveDlg("Save regions JSON", "JSON:*.json");
  }
  if (!target) return;
  target.encoding = "UTF-8";
  target.open("w");
  target.write(toJson(result));
  target.close();

  alert("Wrote " + target.fsName + (warnings.length ? "\n\nWarnings:\n- " + warnings.join("\n- ") : ""));
}

function run() {
  try {
    main();
  } catch (e) {
    alert("Export failed: " + e.message + "\nLine " + e.line + " (error " + e.number + ")");
  }
}

app.doScript(run, ScriptLanguage.JAVASCRIPT, [], UndoModes.ENTIRE_SCRIPT, "Export brochure regions");