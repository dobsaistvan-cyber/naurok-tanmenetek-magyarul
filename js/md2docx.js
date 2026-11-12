/**
 * md2docx.js — Markdown → .docx konverter a Naurok tanársegéd oldalhoz.
 *
 * Böngészőben ÉS Node-ban is működik (UMD-szerű betöltés):
 *  - Böngésző: a `docx` globálist használja (js/docx-bundle.js), és
 *    `window.md2docx.convertMarkdownToDocx` néven érhető el.
 *  - Node: a globális `docx`-et vagy a `require("docx")` csomagot használja,
 *    és `module.exports` formában exportál.
 *
 * Fő függvény: convertMarkdownToDocx(markdownText, title)
 *  → Promise<Blob> böngészőben, Promise<Buffer> Node-ban.
 *
 * A parser a tervfájlok szerkezetét támogatja: #–###### címsorok,
 * **félkövér** és *dőlt* inline formázás (táblázatcellákban is),
 * felsorolás, számozott lista, táblázatok, vízszintes vonal,
 * idézetblokk (>), <br> sortörés és [szöveg](url) hivatkozások.
 * Minden szöveget egy-az-egyben megőriz (magyar, ukrán, angol vegyesen).
 */
(function (root, factory) {
  "use strict";
  var lib = null;
  if (typeof root !== "undefined" && root && root.docx) lib = root.docx;
  if (!lib && typeof require === "function") {
    try { lib = require("docx"); } catch (e) { lib = null; }
  }
  var api = factory(lib);
  if (typeof module === "object" && module.exports) {
    module.exports = api;
  } else if (typeof root !== "undefined" && root) {
    root.md2docx = api;
  }
})(typeof self !== "undefined" ? self : (typeof globalThis !== "undefined" ? globalThis : this), function (docx) {
  "use strict";

  // ---------------------------------------------------------------------------
  // Inline (soron belüli) elemzés
  // ---------------------------------------------------------------------------

  // Védett karakterek a `\*` jellegű escape-ekhez: a tokenizálás előtt
  // privát Unicode-karakterekre cseréljük őket, hogy ne zavarják a
  // formázás-felismerést, majd a végén visszaállítjuk.
  var ESCAPES = { "*": "\uE000", "\\": "\uE001", "`": "\uE002", "_": "\uE003", "[": "\uE004", "]": "\uE005" };
  var UNESCAPES = { "\uE000": "*", "\uE001": "\\", "\uE002": "`", "\uE003": "_", "\uE004": "[", "\uE005": "]" };

  function protectEscapes(s) {
    return s.replace(/\\([\\`*_\[\]])/g, function (_, ch) { return ESCAPES[ch] || ch; });
  }
  function restoreEscapes(s) {
    return s.replace(/[\uE000-\uE005]/g, function (ch) { return UNESCAPES[ch] || ch; });
  }

  // --- Kétfázisú inline elemzés ---
  // 1. fázis: a ***...*** és **...** szakaszokat helyőrzőkre cseréljük,
  //    így a beágyazott `*...**...**...*` (dőltben félkövér) is helyesen
  //    bomlik fel: a félkövér pár előbb kikerül, a dőlt ezután záródik.
  // 2. fázis: a maradékot szkenneljük *dőlt*, `kód` és [hivatkozás]
  //    után, a helyőrzőket visszaállítva (az alapstílussal összeolvasztva).
  var PH_OPEN = "\uE010", PH_CLOSE = "\uE011";

  function parseInlineSegment(seg, base) {
    var holders = [];
    function stash(text, bold, italics, literal) {
      holders.push({ text: text, bold: bold, italics: italics, literal: !!literal });
      return PH_OPEN + (holders.length - 1) + PH_CLOSE;
    }
    // 4+ csillagból álló futam (pl. cenzúrázott szó: serdül********k)
    // nem formázás, hanem tartalom: szó szerint megőrzendő.
    var s = seg.replace(/\*{4,}/g, function (run) { return stash(run, false, false, true); });
    s = s
      .replace(/\*\*\*(?=\S)([^*]+?)\*\*\*/g, function (_, inner) { return stash(inner, true, true); })
      // A félkövér tartalma *dőlt* párokat is tartalmazhat
      // (pl. **4. téma. *Az otthonom***), ezért a tartalom-minta
      // csillag nélküli részek és *...* párok sorozata.
      .replace(/\*\*(?=\S)((?:[^*]|\*[^*]*?\*)+?)\*\*/g, function (_, inner) { return stash(inner, true, false); });
    return scanInline(s, base, holders);
  }

  function scanInline(s, base, holders) {
    var specs = [];
    var buf = "";
    var i = 0, n = s.length;
    function flush() {
      if (buf) { specs.push({ text: restoreEscapes(buf), bold: !!base.bold, italics: !!base.italics }); buf = ""; }
    }
    while (i < n) {
      var ch = s.charAt(i);
      // Helyőrző visszaállítása (a félkövér szövegben lehet hivatkozás/kód is)
      if (ch === PH_OPEN) {
        var pm = s.slice(i).match(/^\uE010(\d+)\uE011/);
        if (pm) {
          flush();
          var h = holders[+pm[1]];
          var merged = { bold: h.bold || !!base.bold, italics: h.italics || !!base.italics };
          if (h.literal) {
            // Szó szerinti tartalom (pl. cenzúra-csillagok): nem elemezzük tovább.
            specs.push({ text: h.text, bold: merged.bold, italics: merged.italics });
          } else {
            specs.push.apply(specs, scanInline(h.text, merged, holders));
          }
          i += pm[0].length;
          continue;
        }
      }
      // Kód: `...`
      if (ch === "`") {
        var ce = s.indexOf("`", i + 1);
        if (ce > i + 1) {
          flush();
          specs.push({ text: s.slice(i + 1, ce), code: true, bold: !!base.bold, italics: !!base.italics });
          i = ce + 1;
          continue;
        }
        buf += ch; i++; continue;
      }
      // Hivatkozás: [szöveg](url)
      if (ch === "[") {
        var lm = s.slice(i).match(/^\[([^\]]+)\]\(([^)\s]+)\)/);
        if (lm) {
          flush();
          specs.push({ text: restoreEscapes(lm[1]), link: lm[2], bold: !!base.bold, italics: !!base.italics });
          i += lm[0].length;
          continue;
        }
        buf += ch; i++; continue;
      }
      // Dőlt: *...* — a nyitó csillag után nem állhat szóköz vagy csillag,
      // a záró csillag nem lehet ** pár második tagja.
      if (ch === "*") {
        var nxt = s.charAt(i + 1);
        if (nxt !== "" && nxt !== "*" && !/\s/.test(nxt)) {
          var j = i + 1, closed = -1;
          while (j < n) {
            if (s.charAt(j) === "*" && s.charAt(j - 1) !== "*") { closed = j; break; }
            j++;
          }
          if (closed > i + 1) {
            flush();
            specs.push.apply(specs, scanInline(s.slice(i + 1, closed), { bold: base.bold, italics: true }, holders));
            i = closed + 1;
            continue;
          }
        }
        buf += ch; i++; continue;
      }
      buf += ch; i++;
    }
    flush();
    return specs;
  }

  // Teljes inline szöveg felbontása; a <br> sortöréseket jelöli
  // (az adott szegmens első futamán `lineBreak: true`).
  function parseInline(text, base) {
    base = base || {};
    var protectedText = protectEscapes(String(text == null ? "" : text));
    var segments = protectedText.split(/<br\s*\/?>/i);
    var out = [];
    segments.forEach(function (seg, i) {
      var specs = parseInlineSegment(seg, base);
      if (i > 0) {
        if (specs.length) specs[0].lineBreak = true;
        else out.push({ text: "", lineBreak: true, bold: !!base.bold, italics: !!base.italics });
      }
      out.push.apply(out, specs);
    });
    // Üres szöveg (pl. üres táblázatcella): az alapstílust öröklő üres futam,
    // hogy a félkövér fejléc üres cellája is félkövér maradjon.
    if (!out.length) out.push({ text: "", bold: !!base.bold, italics: !!base.italics });
    return out;
  }

  // ---------------------------------------------------------------------------
  // Blokk-szintű elemzés
  // ---------------------------------------------------------------------------

  // Táblázatsor felbontása cellákra (a `\|` védett pipe-ot is kezeli).
  function splitRow(line) {
    var s = line.trim();
    if (s.charAt(0) === "|") s = s.slice(1);
    if (s.charAt(s.length - 1) === "|") s = s.slice(0, -1);
    var tmp = s.replace(/\\\|/g, "\uE006");
    return tmp.split("|").map(function (c) { return c.replace(/\uE006/g, "|").trim(); });
  }

  function isSeparator(line) {
    var t = line.trim();
    return t.charAt(0) === "|" && t.indexOf("-") >= 0 && /^\|[\s:|-]+\|$/.test(t);
  }

  // A Markdown-szöveget blokkok listájává alakítja.
  function parseBlocks(markdownText) {
    var lines = String(markdownText == null ? "" : markdownText).replace(/\r\n?/g, "\n").split("\n");
    var blocks = [];
    var i = 0;
    while (i < lines.length) {
      var line = lines[i];
      var trimmed = line.trim();
      if (!trimmed) { i++; continue; }
      var m;
      // Címsor: # ... ######
      if ((m = trimmed.match(/^(#{1,6})\s+(.*)$/))) {
        blocks.push({ type: "heading", level: m[1].length, text: m[2].trim() });
        i++;
        continue;
      }
      // Vízszintes vonal
      if (/^(-{3,}|_{3,}|\*{3,})$/.test(trimmed)) {
        blocks.push({ type: "hr" });
        i++;
        continue;
      }
      // Táblázat: fejléc + elválasztó sor + törzssorok
      if (trimmed.charAt(0) === "|" && i + 1 < lines.length && isSeparator(lines[i + 1])) {
        var header = splitRow(trimmed);
        var rows = [];
        i += 2;
        while (i < lines.length && lines[i].trim().charAt(0) === "|") {
          rows.push(splitRow(lines[i].trim()));
          i++;
        }
        blocks.push({ type: "table", header: header, rows: rows });
        continue;
      }
      // Idézetblokk: egymást követő `>` sorok egy bekezdéssé összevonva
      if (trimmed.charAt(0) === ">") {
        var parts = [];
        while (i < lines.length && lines[i].trim().charAt(0) === ">") {
          parts.push(lines[i].trim().replace(/^>\s?/, ""));
          i++;
        }
        blocks.push({ type: "quote", text: parts.join(" ") });
        continue;
      }
      // Felsorolás (beágyazott szint a behúzásból)
      if ((m = line.match(/^(\s*)[-*+]\s+(.*)$/))) {
        blocks.push({ type: "bullet", level: Math.min(8, Math.floor(m[1].length / 2)), text: m[2] });
        i++;
        continue;
      }
      // Számozott lista — a forrás sorszáma szövegként megmarad
      if ((m = line.match(/^(\s*)\d+[.)]\s+(.*)$/))) {
        blocks.push({ type: "numbered", level: Math.min(8, Math.floor(m[1].length / 2)), text: trimmed });
        i++;
        continue;
      }
      // Sima bekezdés (soronként egy bekezdés, a forrás sortöréseit megtartva)
      blocks.push({ type: "para", text: trimmed });
      i++;
    }
    return blocks;
  }

  // ---------------------------------------------------------------------------
  // docx építés
  // ---------------------------------------------------------------------------

  var BODY_SIZE = 22;   // 11pt (félpontokban)
  var TABLE_SIZE = 20;  // 10pt
  var HEADER_FILL = "D9D9D9"; // világosszürke fejléc-háttér
  var BORDER_COLOR = "B0B0B0";

  function specsToRuns(specs, size) {
    var runs = [];
    specs.forEach(function (s) {
      if (s.link) {
        runs.push(new docx.ExternalHyperlink({
          children: [new docx.TextRun({ text: s.text, color: "0563C1", underline: {}, size: size, bold: !!s.bold, italics: !!s.italics })],
          link: s.link
        }));
        return;
      }
      var opts = { text: s.text, bold: !!s.bold, italics: !!s.italics, size: size };
      if (s.code) opts.font = "Consolas";
      if (s.lineBreak) opts.break = 1;
      runs.push(new docx.TextRun(opts));
    });
    if (!runs.length) runs.push(new docx.TextRun({ text: "", size: size }));
    return runs;
  }

  function tableBorder() {
    return { style: docx.BorderStyle.SINGLE, size: 4, color: BORDER_COLOR, space: 0 };
  }

  function blockToDocx(block) {
    var headingLevels = [
      docx.HeadingLevel.HEADING_1, docx.HeadingLevel.HEADING_2,
      docx.HeadingLevel.HEADING_3, docx.HeadingLevel.HEADING_4,
      docx.HeadingLevel.HEADING_5, docx.HeadingLevel.HEADING_6
    ];
    switch (block.type) {
      case "heading":
        return new docx.Paragraph({
          heading: headingLevels[Math.min(6, block.level) - 1],
          children: specsToRuns(parseInline(block.text, {}), BODY_SIZE)
        });
      case "hr":
        return new docx.Paragraph({
          border: { bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: "999999", space: 1 } },
          spacing: { before: 120, after: 120 }
        });
      case "quote":
        return new docx.Paragraph({
          indent: { left: 720 },
          children: specsToRuns(parseInline(block.text, { italics: true }), BODY_SIZE)
        });
      case "bullet":
        return new docx.Paragraph({
          bullet: { level: block.level },
          children: specsToRuns(parseInline(block.text, {}), BODY_SIZE)
        });
      case "numbered": {
        var indent = block.level * 360;
        return new docx.Paragraph({
          indent: { left: 720 + indent, hanging: 360 },
          children: specsToRuns(parseInline(block.text, {}), BODY_SIZE)
        });
      }
      case "table": {
        var cols = block.header.length;
        block.rows.forEach(function (r) { if (r.length > cols) cols = r.length; });
        function pad(row) {
          var out = row.slice();
          while (out.length < cols) out.push("");
          return out;
        }
        function makeRow(cells, isHeader) {
          return new docx.TableRow({
            tableHeader: isHeader,
            children: pad(cells).map(function (cellText) {
              var cellOpts = {
                children: [new docx.Paragraph({ children: specsToRuns(parseInline(cellText, {}), TABLE_SIZE) })]
              };
              if (isHeader) {
                cellOpts.shading = { type: docx.ShadingType.CLEAR, fill: HEADER_FILL };
              }
              return new docx.TableCell(cellOpts);
            })
          });
        }
        // A fejlécsor félkövér: a cella-szöveget félkövér alapstílussal elemezzük.
        var headerRow = new docx.TableRow({
          tableHeader: true,
          children: pad(block.header).map(function (cellText) {
            return new docx.TableCell({
              shading: { type: docx.ShadingType.CLEAR, fill: HEADER_FILL },
              children: [new docx.Paragraph({ children: specsToRuns(parseInline(cellText, { bold: true }), TABLE_SIZE) })]
            });
          })
        });
        var bodyRows = block.rows.map(function (r) { return makeRow(r, false); });
        var b = tableBorder();
        return new docx.Table({
          width: { size: 100, type: docx.WidthType.PERCENTAGE },
          borders: { top: b, left: b, bottom: b, right: b, insideHorizontal: b, insideVertical: b },
          rows: [headerRow].concat(bodyRows)
        });
      }
      default: // "para"
        return new docx.Paragraph({
          children: specsToRuns(parseInline(block.text, {}), BODY_SIZE)
        });
    }
  }

  function deriveTitle(blocks) {
    for (var i = 0; i < blocks.length; i++) {
      if (blocks[i].type === "heading" && blocks[i].level === 1) return blocks[i].text;
    }
    for (var j = 0; j < blocks.length; j++) {
      if (blocks[j].type === "heading") return blocks[j].text;
    }
    return "Tanmenet";
  }

  function buildDocument(markdownText, title) {
    var blocks = parseBlocks(markdownText);
    var children = blocks.map(blockToDocx);
    return new docx.Document({
      creator: "Naurok Tanmenetek Magyarul",
      title: title || deriveTitle(blocks),
      styles: {
        default: {
          document: {
            run: { font: "Calibri", size: BODY_SIZE },
            paragraph: { spacing: { after: 120 } }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            // A4, keskeny margók (0,5" ≈ 1,27 cm)
            size: { width: 11906, height: 16838 },
            margin: { top: 720, right: 720, bottom: 720, left: 720 }
          }
        },
        children: children
      }]
    });
  }

  function isBrowser() {
    return typeof window !== "undefined" && typeof window.document !== "undefined" && typeof window.Blob !== "undefined";
  }

  function convertMarkdownToDocx(markdownText, title) {
    if (!docx) {
      return Promise.reject(new Error("A docx könyvtár nem töltődött be (hiányzik a js/docx-bundle.js)."));
    }
    var doc;
    try {
      doc = buildDocument(markdownText, title);
    } catch (err) {
      return Promise.reject(err);
    }
    if (isBrowser()) {
      return docx.Packer.toBlob(doc);
    }
    return docx.Packer.toBuffer(doc);
  }

  return {
    convertMarkdownToDocx: convertMarkdownToDocx,
    parseMarkdown: parseBlocks,
    parseInline: parseInline
  };
});
