(function () {
  // diffPad. Loaded by index.html after vendor/jszip.min.js and the readers
  // in js/formats/; styled by css/style.css. Everything runs in this browser
  // tab and nothing is sent anywhere.

  var state = { left: null, right: null }; // {fileName, blocks}

  var app = document.getElementById("app");
  var resetBtn = document.getElementById("reset-btn");
  var inputLeft = document.getElementById("file-input-left");
  var inputRight = document.getElementById("file-input-right");

  var lastDiffHtml = "";

  // Document content reaches the page inside attribute values as well as
  // between tags -- title=, data-href=, id= and href= all carry strings
  // taken straight out of the .docx -- so the quotes have to go too.
  // Without them an author name or a hyperlink target closes the
  // attribute it sits in and everything after it becomes markup.
  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
  function escapeHtmlVisibleWs(s) {
    return escapeHtml(s).replace(/ /g, '<span class="ws-dot">\u00b7</span>');
  }

  // Colours are the one place a document's string is interpolated into
  // CSS rather than into text, and escaping is the wrong tool there:
  // a colour is not arbitrary text that needs quoting, it is one of a
  // known set of values. OOXML writes w:color/@w:val and w:shd/@w:fill
  // as six hex digits or the word "auto", so accept exactly that and
  // discard anything else rather than trying to describe what is unsafe.
  function hexColor(v) {
    return typeof v === "string" && /^[0-9A-Fa-f]{6}$/.test(v)
      ? v.toUpperCase()
      : null;
  }

  // A toggle property in OOXML is on when the element is present, unless
  // its w:val says otherwise. Word writes <w:b w:val="0"/> whenever a run
  // turns off something its paragraph style switched on, so reading mere
  // presence as "on" renders text bold that the document says is not.
  function onOff(el) {
    if (!el) return false;
    var v = el.getAttribute("w:val");
    if (v === null) return true;
    return !(v === "0" || v === "false" || v === "off" || v === "none");
  }

  // ============================== THE DOCUMENT MODEL ==============================
  //
  // Whatever kind of file a document came from, it is reduced to the same
  // thing: a list of blocks, one per paragraph and one per table row. A
  // block carries runs of text ("segments"), each with the formatting it
  // ends up with once styles have been applied, and the paragraph's own
  // formatting. The readers below and in js/formats/ produce that list;
  // everything after it -- alignment, word diff, drawing -- works on blocks
  // and does not know what the file was.
  //
  // Sizes are in half-points, as Word counts them; spacing and indents are
  // in points.

  function clampNumber(v, lo, hi) {
    v = Number(v);
    if (!isFinite(v)) return null;
    return Math.max(lo, Math.min(hi, v));
  }

  // A font name is the other string besides a colour that reaches CSS, and
  // gets the same treatment: keep the characters a font name is made of and
  // drop the rest, rather than try to quote whatever arrives.
  function fontCss(name) {
    if (typeof name !== "string") return "";
    var clean = name
      .replace(/[^A-Za-z0-9 _.\-]/g, "")
      .trim()
      .slice(0, 60);
    if (!clean) return "";
    var generic = /courier|consol|mono|typewriter/i.test(clean)
      ? "monospace"
      : /sans|arial|helvet|calibri|verdana|tahoma|segoe|aptos/i.test(clean)
        ? "sans-serif"
        : /times|georgia|garamond|cambria|palatino|book|roman|serif|minion|century|constantia/i.test(
              clean,
            )
          ? "serif"
          : "sans-serif";
    return "font-family:'" + clean + "'," + generic + ";";
  }

  function sizeCss(halfPoints) {
    var n = clampNumber(halfPoints, 2, 400);
    return n === null ? "" : "font-size:" + n / 2 + "pt;";
  }

  // The formatting a run adds to the paragraph it sits in. `base` is what
  // the paragraph already gives every run, so an ordinary run adds nothing
  // and a heading does not repeat its size on each word.
  function runCss(s, base) {
    base = base || {};
    var css = "";
    if (!!s.bold !== !!base.bold)
      css += "font-weight:" + (s.bold ? "700" : "400") + ";";
    if (!!s.italic !== !!base.italic)
      css += "font-style:" + (s.italic ? "italic" : "normal") + ";";
    // Underline and strike-through are never left to the paragraph: a
    // decoration drawn by a parent cannot be switched off by a child.
    var deco = [];
    if (s.underline) deco.push("underline");
    if (s.strike) deco.push("line-through");
    if (deco.length) css += "text-decoration:" + deco.join(" ") + ";";
    var color = hexColor(s.color),
      baseColor = hexColor(base.color);
    if (color !== baseColor) css += "color:#" + (color || "000000") + ";";
    if (s.vert === "super" || s.vert === "sub") {
      css += "vertical-align:" + s.vert + ";font-size:0.75em;";
    } else if (s.size && s.size !== base.size) {
      css += sizeCss(s.size);
    }
    if (s.font && s.font !== base.font) css += fontCss(s.font);
    var hl = hexColor(s.highlight);
    if (hl) css += "background:#" + hl + ";";
    if (s.caps) css += "text-transform:uppercase;";
    return css;
  }

  // What the paragraph gives all of its runs.
  function baseCss(base) {
    if (!base) return "";
    var css = "";
    if (base.size) css += sizeCss(base.size);
    if (base.font) css += fontCss(base.font);
    if (base.bold) css += "font-weight:700;";
    if (base.italic) css += "font-style:italic;";
    var color = hexColor(base.color);
    if (color) css += "color:#" + color + ";";
    return css;
  }

  var ALIGN = { left: 1, center: 1, right: 1, justify: 1 };

  function pt(v, lo, hi) {
    var n = clampNumber(v, lo, hi);
    return n === null ? null : Math.round(n * 100) / 100 + "pt";
  }

  // The paragraph's own formatting: where it sits and how much room it has.
  function paraCss(block) {
    var p = block.para || {};
    var css = "";
    if (ALIGN[p.align] && p.align !== "left")
      css += "text-align:" + p.align + ";";
    var left = clampNumber(p.indL || 0, -72, 360) || 0;
    var first = clampNumber(p.indFirst || 0, -360, 360) || 0;
    if (block.isListItem) {
      // A list item's number hangs in its indent; the number is drawn
      // beside the text, so only the level is kept.
      left = 0;
      first = 0;
    }
    // A hanging indent pulls the first line back out; it can only pull it
    // as far as the paragraph is indented.
    if (left + first < 0) first = -left;
    if (left) css += "margin-left:" + pt(left, -72, 360) + ";";
    if (p.indR) {
      var r = pt(p.indR, 0, 360);
      if (r && r !== "0pt") css += "margin-right:" + r + ";";
    }
    if (first) css += "text-indent:" + pt(first, -360, 360) + ";";
    if (p.before !== undefined && p.before !== null) {
      var b = pt(p.before, 0, 72);
      if (b && b !== "0pt") css += "margin-top:" + b + ";";
    }
    if (p.after !== undefined && p.after !== null) {
      var a = pt(p.after, 0, 72);
      if (a) css += "margin-bottom:" + a + ";";
    }
    return css + baseCss(block.base);
  }

  // ---- numbers and bullets

  function toRoman(n) {
    if (n <= 0 || n >= 4000) return String(n);
    var out = "";
    [
      [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"],
      [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
    ].forEach(function (p) {
      while (n >= p[0]) {
        out += p[1];
        n -= p[0];
      }
    });
    return out;
  }

  // a, b ... z, aa, bb ... the way Word letters a list.
  function toLetters(n) {
    if (n <= 0) return String(n);
    var letter = String.fromCharCode(97 + ((n - 1) % 26));
    var out = "";
    for (var i = 0; i <= Math.floor((n - 1) / 26) && i < 20; i++) out += letter;
    return out;
  }

  // `fmt` is one of decimal, lowerLetter, upperLetter, lowerRoman,
  // upperRoman, decimalZero, none; anything else counts in digits.
  function formatNumber(n, fmt) {
    if (fmt === "lowerLetter") return toLetters(n);
    if (fmt === "upperLetter") return toLetters(n).toUpperCase();
    if (fmt === "lowerRoman") return toRoman(n);
    if (fmt === "upperRoman") return toRoman(n).toUpperCase();
    if (fmt === "decimalZero") return n < 10 ? "0" + n : String(n);
    if (fmt === "none") return "";
    return String(n);
  }

  function bulletFor(level) {
    return level >= 2 ? "▪" : level === 1 ? "◦" : "•";
  }

  // A bullet is stored as a character in a symbol font, where it sits in
  // the private-use area and means nothing in any other font.
  var SYMBOL_BULLETS = {
    "": "•",
    "·": "•",
    "": "▪",
    "§": "▪",
    "": "➢",
    "": "✓",
    "": "◆",
    "": "❖",
    "": "–",
    o: "◦",
  };

  function bulletGlyph(text, level) {
    text = String(text || "").trim();
    if (!text) return bulletFor(level);
    if (Object.prototype.hasOwnProperty.call(SYMBOL_BULLETS, text))
      return SYMBOL_BULLETS[text];
    if (/^[-]+$/.test(text) || text.length > 3) return bulletFor(level);
    return text;
  }

  // Counts list items as a document is read. `levels` describes one list:
  // for each level its format, the text of its number with %1, %2 ...
  // standing for the levels' counters, and where it starts.
  function listCounter() {
    var counters = Object.create(null);
    return function (listId, level, levels) {
      level = Math.max(0, Math.min(8, level | 0));
      var def = (levels && levels[level]) || null;
      if (!def || def.fmt === "bullet") {
        return {
          kind: "bullet",
          label: bulletGlyph(def ? def.text : "", level),
        };
      }
      var c = counters[listId] || (counters[listId] = []);
      for (var i = 0; i < level; i++) {
        if (c[i] === undefined) {
          var d = levels[i];
          c[i] = d && d.start !== undefined ? d.start : 1;
        }
      }
      c[level] =
        c[level] === undefined
          ? def.start !== undefined
            ? def.start
            : 1
          : c[level] + 1;
      c.length = level + 1;
      var text = def.text === undefined || def.text === null ? "%" + (level + 1) + "." : def.text;
      var label = text.replace(/%([1-9])/g, function (whole, d) {
        var idx = parseInt(d, 10) - 1;
        var lv = levels[idx];
        return c[idx] === undefined
          ? ""
          : formatNumber(c[idx], lv ? lv.fmt : "decimal");
      });
      return { kind: "number", label: label };
    };
  }

  // ---- blocks

  function makeParagraphBlock(pd, sectionLabel) {
    return {
      sectionLabel: sectionLabel,
      isHeading: !!pd.isHeading,
      headingLevel: pd.headingLevel || 0,
      isListItem: !!pd.isListItem,
      ilvl: pd.ilvl || 0,
      listLabel: pd.listLabel || "",
      listKind: pd.listKind || "",
      para: pd.para || {},
      base: pd.base || {},
      segments: pd.segments,
      isTableRow: false,
      isDivider: false,
    };
  }
  function makeTableRowBlock(cells, sectionLabel) {
    return {
      sectionLabel: sectionLabel,
      isTableRow: true,
      cells: cells,
      isDivider: false,
    };
  }
  function makeDividerBlock(label) {
    return {
      sectionLabel: "divider",
      isDivider: true,
      segments: [
        {
          text: "— " + label + " —",
          isDel: false,
          isImage: false,
        },
      ],
    };
  }

  // Images and footnote references are diffed as tokens, so they need a
  // spelling inside the plain text that the word tokeniser can pick out
  // again. Both are wrapped in their own marker character at each end:
  // the fingerprint of an image is base64, which has no delimiter of its
  // own, so an opening marker alone left "fig<mark>1892:AAB...</mark>one"
  // as a single word whenever an image sat against text with no space.
  // That word matched nothing in the token map and was printed verbatim.
  var IMG_MARK = "￼";
  var FN_MARK = "�";

  function imageToken(seg) {
    return IMG_MARK + (seg.imgFingerprint || seg.relId || "?") + IMG_MARK;
  }
  function footnoteToken(seg) {
    return FN_MARK + (seg.noteKind || "") + (seg.footnoteId || "?") + FN_MARK;
  }
  // Marker characters occurring in the document's own text would make a
  // token ambiguous, so they are dropped from the text used for
  // alignment. They stay in the rendered markup, so a paragraph that
  // differs only by one still shows up -- as a formatting change.
  function stripTokenMarks(s) {
    return String(s).replace(/[￼�]/g, "");
  }

  // Record every image and footnote segment under the exact token that
  // plainTextForDiff will emit for it, so the renderer can turn a token
  // back into the thing it stands for. Both sides go through the same two
  // helpers; building the key separately is how they drifted before.
  function indexTokens(block, segments) {
    segments.forEach(function (s) {
      if (s.isDel) return;
      if (s.isImage) block.imageTokens[imageToken(s)] = s;
      else if (s.isFootnoteRef || s.isFootnoteMarker)
        block.footnoteTokens[footnoteToken(s)] = s;
    });
  }

  function segmentPlainText(s) {
    if (s.isImage) return imageToken(s);
    if (s.isFootnoteRef || s.isFootnoteMarker) return footnoteToken(s);
    return stripTokenMarks(s.text);
  }

  function plainTextForDiff(segments) {
    return segments
      .filter(function (s) {
        return !s.isDel;
      })
      .map(segmentPlainText)
      .join("");
  }

  // Where each run falls in the plain text, so that a word picked out by
  // the word diff can be drawn with the formatting of the run it came from.
  function spansFor(segments) {
    var spans = [];
    var at = 0;
    segments.forEach(function (s) {
      if (s.isDel) return;
      var len = segmentPlainText(s).length;
      if (len && !s.isImage && !s.isFootnoteRef && !s.isFootnoteMarker)
        spans.push({ start: at, end: at + len, seg: s });
      at += len;
    });
    return spans;
  }

  function footnoteRefHtml(s) {
    return (
      '<sup class="fn-ref" id="' +
      escapeHtml(s.refAnchorId) +
      '"><a class="fn-ref-link" href="#' +
      escapeHtml(s.anchorId) +
      '" title="Jump to ' +
      (s.noteKind === "en" ? "endnote " : "footnote ") +
      escapeHtml(s.footnoteId) +
      '">' +
      escapeHtml(s.footnoteId) +
      "</a></sup>"
    );
  }

  function footnoteMarkerHtml(s) {
    return (
      '<sup class="fn-marker" id="' +
      escapeHtml(s.anchorId) +
      '"><a class="fn-marker-link" href="#' +
      escapeHtml(s.backAnchorId) +
      '" title="Back to text">' +
      escapeHtml(s.footnoteId) +
      "</a></sup>"
    );
  }

  // An image's data: URL is built here from bytes, never copied from a
  // document, but it is checked all the same before it becomes an
  // attribute.
  function safeImageSrc(url) {
    return typeof url === "string" &&
      /^data:image\/(png|jpeg|gif|bmp);base64,[A-Za-z0-9+/=]*$/.test(url)
      ? url
      : "";
  }

  function imageTagHtml(s) {
    var src = safeImageSrc(s.dataUrl);
    if (!src) return "[image]";
    var w = clampNumber(s.widthPx, 1, 2000);
    return (
      '<img src="' +
      src +
      '"' +
      (w ? ' style="width:' + Math.round(w) + 'px;"' : "") +
      ">"
    );
  }

  // One run's text, already escaped, wrapped in whatever the run carries:
  // its formatting, a tracked change, a link.
  function wrapRunHtml(s, inner, base) {
    var style = runCss(s, base);
    var cls = [];
    var title = "";
    if (s.isDel) {
      cls.push("trk-del");
      title = "Tracked deletion" + (s.author ? " by " + s.author : "");
    } else if (s.trackType === "ins") {
      cls.push("trk-ins");
      title = "Tracked insertion" + (s.author ? " by " + s.author : "");
    }
    if (s.href) {
      inner =
        '<a class="doc-link" data-href="' +
        escapeHtml(s.href) +
        '" title="' +
        escapeHtml(s.href) +
        '">' +
        inner +
        "</a>";
    }
    return (
      "<span" +
      (cls.length ? ' class="' + cls.join(" ") + '"' : "") +
      (style ? ' style="' + style + '"' : "") +
      (title ? ' title="' + escapeHtml(title) + '"' : "") +
      ">" +
      inner +
      "</span>"
    );
  }

  function renderSegmentsHtml(segments, base) {
    return segments
      .map(function (s) {
        if (s.isFootnoteRef) return footnoteRefHtml(s);
        if (s.isFootnoteMarker) return footnoteMarkerHtml(s);
        if (s.isImage) {
          var cls =
            "inline-img" +
            (s.trackType === "ins" ? " trk-ins" : "") +
            (s.trackType === "del" ? " trk-del" : "");
          var title = s.trackType
            ? "Tracked " + s.trackType + (s.author ? " by " + s.author : "")
            : "Embedded image";
          return (
            '<span class="' +
            cls +
            '" title="' +
            escapeHtml(title) +
            '">' +
            imageTagHtml(s) +
            "</span>"
          );
        }
        var text = escapeHtml(s.text)
          .replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;")
          .replace(/\n/g, "<br>");
        return wrapRunHtml(s, text, base);
      })
      .join("");
  }

  function cellStyleAttr(c) {
    var css = "";
    var w = clampNumber(c.width, 0, 100000);
    if (w) css += "flex-grow:" + Math.round(w) + ";";
    var fill = hexColor(c.fill);
    if (fill) css += "background:#" + fill + ";";
    if (ALIGN[c.align] && c.align !== "left") css += "text-align:" + c.align + ";";
    css += baseCss(c.base);
    return css ? ' style="' + css + '"' : "";
  }

  function rowCellHtml(c, inner) {
    return '<div class="row-cell"' + cellStyleAttr(c) + ">" + inner + "</div>";
  }

  // Give every block the two strings the rest of the page works from: the
  // text it is aligned on and the markup it is drawn with.
  function finalizeBlocks(blocks, docId) {
    blocks.forEach(function (b) {
      b.imageTokens = {};
      b.footnoteTokens = {};
      if (b.isTableRow) {
        b.text = b.cells
          .map(function (c) {
            return plainTextForDiff(c.segs);
          })
          .join(" | ");
        b.html = b.cells
          .map(function (c) {
            c.spans = spansFor(c.segs);
            return rowCellHtml(c, renderSegmentsHtml(c.segs, c.base));
          })
          .join("");
        b.cells.forEach(function (c) {
          indexTokens(b, c.segs);
        });
        b.paraKey = "";
      } else {
        b.text = plainTextForDiff(b.segments);
        b.html = renderSegmentsHtml(b.segments, b.base);
        b.spans = spansFor(b.segments);
        indexTokens(b, b.segments);
        b.paraCss = b.isDivider ? "" : paraCss(b);
        b.paraKey = b.isDivider
          ? ""
          : b.paraCss +
            (b.isHeading ? "h" + b.headingLevel : "") +
            (b.isListItem ? "li" + b.ilvl + b.listKind : "");
      }

      // b.html carries docId, which is freshly randomised on every read,
      // so the same file read twice produces two different strings for
      // any paragraph holding a footnote anchor. Comparing b.html to
      // decide "formatting changed" therefore reported every one of
      // those paragraphs as changed against an identical copy of itself.
      // The key is the same markup with the per-read id taken back out,
      // and it -- not b.html -- is what the alignment compares, along
      // with the paragraph's own formatting.
      var htmlKey = docId ? b.html.split(docId).join("") : b.html;
      b.fmtKey = b.paraKey + "\u0001" + htmlKey;

      // A paragraph that is entirely a tracked deletion (or otherwise has no
      // accepted text but still has visible markup, e.g. an in-source comment
      // anchor) must never be treated as an interchangeable "blank paragraph"
      // during alignment -- give it a key of its own so it surfaces as its
      // own change instead of silently absorbing into an unrelated blank.
      // The key is the markup itself and not a counter: the counter ran across
      // both documents at once, so the same deleted paragraph was handed a
      // different key on each side and always came out as one removal plus one
      // addition -- even when a file was compared with itself.
      if (b.text === "" && htmlKey !== "") {
        b.text = "\u0000" + htmlKey;
      }
    });
    return blocks;
  }

  function newDocId() {
    return "d" + Math.random().toString(36).slice(2, 9) + "_";
  }

  // ---- what the readers in js/formats/ hand back
  //
  // Those readers know nothing about this page. They return plain
  // descriptions -- { type: "p" | "row" | "divider" | "note" | "comment" }
  // -- and this turns them into blocks, adding the anchors that link a
  // footnote to its reference.
  function noteRefSegment(note, docId) {
    var kind = note.kind === "en" ? "en" : "";
    var id = String(note.id);
    return {
      text: id,
      isFootnoteRef: true,
      footnoteId: id,
      noteKind: kind,
      anchorId: docId + (kind || "fn") + "-" + id,
      refAnchorId: docId + (kind || "fn") + "ref-" + id,
      isDel: false,
      isImage: false,
    };
  }

  function noteMarkerSegment(note, docId) {
    var kind = note.kind === "en" ? "en" : "";
    var id = String(note.id);
    return {
      text: id,
      isFootnoteMarker: true,
      footnoteId: id,
      noteKind: kind,
      anchorId: docId + (kind || "fn") + "-" + id,
      backAnchorId: docId + (kind || "fn") + "ref-" + id,
      isDel: false,
      isImage: false,
    };
  }

  function segmentsFromRaw(raw, docId) {
    return (raw || []).map(function (s) {
      if (s.note) return noteRefSegment(s.note, docId);
      if (s.image) {
        return {
          text: IMG_MARK,
          isImage: true,
          isDel: false,
          relId: null,
          dataUrl: s.image.dataUrl || null,
          imgFingerprint: s.image.fingerprint || null,
          widthPx: s.image.widthPx || null,
          trackType: s.trackType || null,
          author: s.author || null,
        };
      }
      return {
        text: String(s.text === undefined || s.text === null ? "" : s.text),
        bold: !!s.bold,
        italic: !!s.italic,
        underline: !!s.underline,
        strike: !!s.strike,
        color: hexColor(s.color),
        size: s.size || null,
        font: s.font || null,
        highlight: hexColor(s.highlight),
        vert: s.vert || null,
        caps: !!s.caps,
        trackType: s.trackType || null,
        author: s.author || null,
        href: s.href || null,
        isDel: !!s.isDel,
        isImage: false,
      };
    });
  }

  // A reader that resolves no styles gives every run its formatting in
  // full. What all of a paragraph's runs agree on is the paragraph's, and
  // is said once.
  function commonBase(segs) {
    var base = null;
    segs.forEach(function (s) {
      if (s.isImage || s.isFootnoteRef || s.isFootnoteMarker || !s.text) return;
      if (/^\s*$/.test(s.text) && base) return;
      var own = {
        bold: !!s.bold,
        italic: !!s.italic,
        color: s.color || null,
        size: s.size || null,
        font: s.font || null,
      };
      if (!base) {
        base = own;
        return;
      }
      Object.keys(own).forEach(function (k) {
        if (base[k] !== own[k]) base[k] = k === "bold" || k === "italic" ? false : null;
      });
    });
    return base || {};
  }

  function blocksFromRaw(raw, docId) {
    var blocks = [];
    raw.forEach(function (r) {
      var section = r.section || "body";
      if (r.type === "divider") {
        blocks.push(makeDividerBlock(r.label));
      } else if (r.type === "row") {
        var row = makeTableRowBlock(
          (r.cells || []).map(function (c) {
            var cellSegs = segmentsFromRaw(c.segments, docId);
            return {
              segs: cellSegs,
              fill: hexColor(c.fill) === "FFFFFF" ? null : hexColor(c.fill),
              width: c.width || null,
              align: c.align || null,
              base: c.base || commonBase(cellSegs),
            };
          }),
          section,
        );
        row.rowFirst = !!r.first;
        row.rowLast = !!r.last;
        blocks.push(row);
      } else if (r.type === "note" || r.type === "comment") {
        var lead =
          r.type === "note"
            ? noteMarkerSegment(r, docId)
            : {
                text: "Comment (" + (r.author || "Unknown") + "): ",
                bold: true,
                isDel: false,
                isImage: false,
              };
        var noteSegs = segmentsFromRaw(r.segments, docId);
        blocks.push(
          makeParagraphBlock(
            {
              segments: [lead].concat(noteSegs),
              base: r.base || commonBase(noteSegs),
            },
            r.type === "comment" ? "comment" : r.kind === "en" ? "endnote" : "footnote",
          ),
        );
      } else {
        var list = r.list || null;
        var paraSegs = segmentsFromRaw(r.segments, docId);
        blocks.push(
          makeParagraphBlock(
            {
              isHeading: r.heading > 0,
              headingLevel: Math.max(0, Math.min(9, r.heading | 0)),
              isListItem: !!list,
              ilvl: list ? Math.max(0, Math.min(8, list.level | 0)) : 0,
              listLabel: list ? String(list.label || "") : "",
              listKind: list ? (list.kind === "number" ? "number" : "bullet") : "",
              para: r.para || {},
              base: r.base || commonBase(paraSegs),
              segments: paraSegs,
            },
            section,
          ),
        );
      }
    });
    return blocks;
  }

  // ============================== OOXML EXTRACTION ==============================

  function childrenLocal(el, name) {
    return Array.from(el.childNodes).filter(function (n) {
      return n.nodeType === 1 && n.localName === name;
    });
  }
  function allLocal(el, name) {
    return Array.from(el.getElementsByTagName("*")).filter(function (n) {
      return n.localName === name;
    });
  }
  function firstLocal(el, name) {
    var arr = childrenLocal(el, name);
    return arr.length ? arr[0] : null;
  }

  var HIGHLIGHT = {
    yellow: "FFFF00", green: "00FF00", cyan: "00FFFF", magenta: "FF00FF",
    blue: "0000FF", red: "FF0000", darkBlue: "000080", darkCyan: "008080",
    darkGreen: "008000", darkMagenta: "800080", darkRed: "800000",
    darkYellow: "808000", darkGray: "808080", lightGray: "C0C0C0",
    black: "000000", white: "FFFFFF",
  };

  // What a w:rPr says, and only what it says: a property it does not
  // mention is left out, so that the result can be laid over whatever the
  // style underneath gave.
  function readRunProps(rPr, theme) {
    var p = {};
    if (!rPr) return p;
    var b = firstLocal(rPr, "b"),
      bCs = firstLocal(rPr, "bCs");
    if (b || bCs) p.bold = onOff(b) || onOff(bCs);
    var i = firstLocal(rPr, "i"),
      iCs = firstLocal(rPr, "iCs");
    if (i || iCs) p.italic = onOff(i) || onOff(iCs);
    var u = firstLocal(rPr, "u");
    if (u) p.underline = onOff(u);
    var strike = firstLocal(rPr, "strike");
    if (strike) p.strike = onOff(strike);
    var caps = firstLocal(rPr, "caps");
    if (caps) p.caps = onOff(caps);
    var colorEl = firstLocal(rPr, "color");
    if (colorEl) {
      var cv = colorEl.getAttribute("w:val");
      if (hexColor(cv)) p.color = hexColor(cv);
      else if (cv === "auto") p.color = null;
    }
    var sz = firstLocal(rPr, "sz");
    if (sz) {
      var half = parseInt(sz.getAttribute("w:val") || "", 10);
      if (half > 0) p.size = half;
    }
    var fonts = firstLocal(rPr, "rFonts");
    if (fonts) {
      var face = fonts.getAttribute("w:ascii") || fonts.getAttribute("w:hAnsi");
      var themed = fonts.getAttribute("w:asciiTheme") || fonts.getAttribute("w:hAnsiTheme");
      if (!face && themed && theme)
        face = /^major/.test(themed) ? theme.major : theme.minor;
      if (face) p.font = face;
    }
    var hl = firstLocal(rPr, "highlight");
    if (hl) {
      var hv = hl.getAttribute("w:val");
      p.highlight = Object.prototype.hasOwnProperty.call(HIGHLIGHT, hv)
        ? HIGHLIGHT[hv]
        : null;
    }
    var va = firstLocal(rPr, "vertAlign");
    if (va) {
      var vv = va.getAttribute("w:val");
      p.vert = vv === "superscript" ? "super" : vv === "subscript" ? "sub" : null;
    }
    var rStyle = firstLocal(rPr, "rStyle");
    if (rStyle) p.styleId = rStyle.getAttribute("w:val");
    return p;
  }

  function twips(el, name) {
    var v = el.getAttribute(name);
    if (v === null || v === "") return undefined;
    var n = parseInt(v, 10);
    return isFinite(n) ? n / 20 : undefined;
  }

  var JC = {
    left: "left", start: "left", center: "center", right: "right",
    end: "right", both: "justify", distribute: "justify",
  };

  function readParaProps(pPr) {
    var p = {};
    if (!pPr) return p;
    var jc = firstLocal(pPr, "jc");
    if (jc) {
      var jv = jc.getAttribute("w:val");
      if (Object.prototype.hasOwnProperty.call(JC, jv)) p.align = JC[jv];
    }
    var ind = firstLocal(pPr, "ind");
    if (ind) {
      var left = twips(ind, "w:left");
      if (left === undefined) left = twips(ind, "w:start");
      if (left !== undefined) p.indL = left;
      var right = twips(ind, "w:right");
      if (right === undefined) right = twips(ind, "w:end");
      if (right !== undefined) p.indR = right;
      var hanging = twips(ind, "w:hanging"),
        firstLine = twips(ind, "w:firstLine");
      if (hanging !== undefined) p.indFirst = -hanging;
      else if (firstLine !== undefined) p.indFirst = firstLine;
    }
    var spacing = firstLocal(pPr, "spacing");
    if (spacing) {
      var before = twips(spacing, "w:before"),
        after = twips(spacing, "w:after");
      if (before !== undefined) p.before = before;
      if (after !== undefined) p.after = after;
    }
    var numPr = firstLocal(pPr, "numPr");
    if (numPr) {
      var numId = firstLocal(numPr, "numId"),
        ilvl = firstLocal(numPr, "ilvl");
      if (numId) p.numId = numId.getAttribute("w:val");
      if (ilvl) p.ilvl = parseInt(ilvl.getAttribute("w:val") || "0", 10) || 0;
    }
    var outline = firstLocal(pPr, "outlineLvl");
    if (outline) {
      var lvl = parseInt(outline.getAttribute("w:val") || "", 10);
      if (isFinite(lvl)) p.outline = lvl;
    }
    var pStyle = firstLocal(pPr, "pStyle");
    if (pStyle) p.styleId = pStyle.getAttribute("w:val");
    return p;
  }

  // word/styles.xml: the document's defaults and every style, by id.
  function readStyles(stylesXml, theme) {
    var styles = {
      byId: Object.create(null),
      run: {},
      para: {},
      defaultPara: null,
      resolved: Object.create(null),
    };
    if (!stylesXml || !stylesXml.documentElement) return styles;
    var root = stylesXml.documentElement;
    var dd = firstLocal(root, "docDefaults");
    if (dd) {
      var rd = firstLocal(dd, "rPrDefault"),
        pd = firstLocal(dd, "pPrDefault");
      if (rd) styles.run = readRunProps(firstLocal(rd, "rPr"), theme);
      if (pd) styles.para = readParaProps(firstLocal(pd, "pPr"));
    }
    childrenLocal(root, "style").forEach(function (st) {
      var id = st.getAttribute("w:styleId");
      if (id === null) return;
      var nameEl = firstLocal(st, "name"),
        basedOn = firstLocal(st, "basedOn");
      var type = st.getAttribute("w:type");
      styles.byId[id] = {
        type: type,
        name: (nameEl && nameEl.getAttribute("w:val")) || "",
        basedOn: basedOn ? basedOn.getAttribute("w:val") : null,
        run: readRunProps(firstLocal(st, "rPr"), theme),
        para: readParaProps(firstLocal(st, "pPr")),
      };
      var dflt = st.getAttribute("w:default");
      if (type === "paragraph" && (dflt === "1" || dflt === "true"))
        styles.defaultPara = id;
    });
    return styles;
  }

  // A style with everything it is based on laid underneath it.
  function resolveStyle(styles, id) {
    if (id === null || id === undefined || !styles.byId[id])
      return { run: {}, para: {}, heading: 0 };
    if (styles.resolved[id]) return styles.resolved[id];
    var chain = [];
    var seen = Object.create(null);
    for (var at = id; at !== null && styles.byId[at] && !seen[at]; ) {
      seen[at] = true;
      chain.unshift(styles.byId[at]);
      at = styles.byId[at].basedOn;
    }
    var out = { run: {}, para: {}, heading: 0 };
    chain.forEach(function (st) {
      Object.assign(out.run, st.run);
      Object.assign(out.para, st.para);
      // "heading 1" is the name Word gives the built-in style in every
      // language, whatever the style's id has been translated to.
      var m = /^heading ([1-9])$/i.exec(st.name);
      if (m) out.heading = parseInt(m[1], 10);
    });
    delete out.run.styleId;
    styles.resolved[id] = out;
    return out;
  }

  // word/numbering.xml: for each list a paragraph can name, its levels.
  function readNumbering(numXml) {
    var lists = Object.create(null);
    if (!numXml || !numXml.documentElement) return lists;
    var root = numXml.documentElement;
    function readLevels(parent, into) {
      childrenLocal(parent, "lvl").forEach(function (lvl) {
        var i = parseInt(lvl.getAttribute("w:ilvl") || "0", 10) || 0;
        var fmt = firstLocal(lvl, "numFmt"),
          text = firstLocal(lvl, "lvlText"),
          start = firstLocal(lvl, "start");
        var def = into[i] || (into[i] = {});
        if (fmt) def.fmt = fmt.getAttribute("w:val") || "decimal";
        if (text) def.text = text.getAttribute("w:val") || "";
        if (start) {
          var s = parseInt(start.getAttribute("w:val") || "", 10);
          if (isFinite(s)) def.start = s;
        }
      });
    }
    var abstracts = Object.create(null);
    childrenLocal(root, "abstractNum").forEach(function (a) {
      var levels = [];
      readLevels(a, levels);
      abstracts[a.getAttribute("w:abstractNumId")] = levels;
    });
    childrenLocal(root, "num").forEach(function (num) {
      var ref = firstLocal(num, "abstractNumId");
      var base = (ref && abstracts[ref.getAttribute("w:val")]) || [];
      var levels = base.map(function (d) {
        return d ? Object.assign({}, d) : d;
      });
      childrenLocal(num, "lvlOverride").forEach(function (ov) {
        var i = parseInt(ov.getAttribute("w:ilvl") || "0", 10) || 0;
        var so = firstLocal(ov, "startOverride");
        if (so) {
          var s = parseInt(so.getAttribute("w:val") || "", 10);
          if (isFinite(s)) (levels[i] || (levels[i] = {})).start = s;
        }
        readLevels(ov, levels);
      });
      lists[num.getAttribute("w:numId")] = levels;
    });
    return lists;
  }

  function emptyEnv(docId) {
    return {
      docId: docId || "",
      styles: readStyles(null),
      numbering: Object.create(null),
      theme: null,
      nextLabel: listCounter(),
    };
  }

  function extractParagraph(pEl, pendingImages, rels, basePart, env) {
    if (!env || typeof env !== "object") env = emptyEnv(env);
    var docId = env.docId;
    var styles = env.styles;

    var pPr = firstLocal(pEl, "pPr");
    var direct = readParaProps(pPr);
    var styleId =
      direct.styleId !== undefined && direct.styleId !== null
        ? direct.styleId
        : styles.defaultPara;
    var style = resolveStyle(styles, styleId);
    var para = Object.assign({}, styles.para, style.para, direct);
    var base = Object.assign({}, styles.run, style.run);

    var isHeading = false,
      headingLevel = 0;
    if (style.heading) {
      headingLevel = style.heading;
    } else if (styleId && /^Heading([1-9])/.test(styleId)) {
      headingLevel = parseInt(RegExp.$1, 10);
    } else if (para.outline !== undefined && para.outline >= 0 && para.outline < 9) {
      headingLevel = para.outline + 1;
    }
    isHeading = headingLevel > 0;

    // numId 0 is how a paragraph says it is not in the list its style
    // would otherwise put it in.
    var isListItem =
      para.numId !== undefined && para.numId !== null && para.numId !== "0";
    var ilvl = isListItem ? para.ilvl || 0 : 0;
    var label = isListItem
      ? env.nextLabel(para.numId, ilvl, env.numbering[para.numId])
      : null;

    var segments = [];

    function runProps(rPr) {
      var own = readRunProps(rPr, env.theme);
      var props = Object.assign({}, base);
      if (own.styleId) Object.assign(props, resolveStyle(styles, own.styleId).run);
      Object.assign(props, own);
      return props;
    }

    function walk(node, ctx) {
      Array.from(node.childNodes).forEach(function (child) {
        if (child.nodeType !== 1) return;
        var ln = child.localName;
        if (ln === "ins") {
          walk(
            child,
            Object.assign({}, ctx, {
              trackType: "ins",
              author: child.getAttribute("w:author") || "",
            }),
          );
        } else if (ln === "del") {
          walk(
            child,
            Object.assign({}, ctx, {
              trackType: "del",
              author: child.getAttribute("w:author") || "",
            }),
          );
        } else if (ln === "hyperlink") {
          var rid = child.getAttribute("r:id");
          var anchor = child.getAttribute("w:anchor");
          var href = null;
          if (rid && rels[rid]) href = rels[rid].target;
          else if (anchor) href = "#" + anchor;
          walk(child, Object.assign({}, ctx, { href: href }));
        } else if (ln === "r") {
          var props = runProps(firstLocal(child, "rPr"));
          Array.from(child.childNodes).forEach(function (tn) {
            if (tn.nodeType !== 1) return;
            if (tn.localName === "t" || tn.localName === "delText") {
              segments.push({
                text: tn.textContent,
                bold: !!props.bold,
                italic: !!props.italic,
                underline: !!props.underline,
                strike: !!props.strike,
                color: props.color || null,
                size: props.size || null,
                font: props.font || null,
                highlight: props.highlight || null,
                vert: props.vert || null,
                caps: !!props.caps,
                trackType: ctx.trackType || null,
                author: ctx.author || null,
                href: ctx.href || null,
                isDel: tn.localName === "delText",
                isImage: false,
              });
            } else if (tn.localName === "tab") {
              segments.push({
                text: "\t",
                bold: !!base.bold,
                italic: !!base.italic,
                color: base.color || null,
                size: base.size || null,
                font: base.font || null,
                isDel: false,
                isImage: false,
                trackType: ctx.trackType || null,
              });
            } else if (tn.localName === "br" || tn.localName === "cr") {
              segments.push({
                text: "\n",
                bold: !!base.bold,
                italic: !!base.italic,
                color: base.color || null,
                size: base.size || null,
                font: base.font || null,
                isDel: false,
                isImage: false,
                trackType: ctx.trackType || null,
              });
            } else if (tn.localName === "noBreakHyphen") {
              segments.push({
                text: "\u2011",
                bold: !!props.bold,
                italic: !!props.italic,
                color: props.color || null,
                size: props.size || null,
                font: props.font || null,
                isDel: false,
                isImage: false,
                trackType: ctx.trackType || null,
              });
            } else if (tn.localName === "drawing") {
              var blip = allLocal(tn, "blip")[0];
              var embed = blip ? blip.getAttribute("r:embed") : null;
              var extent = allLocal(tn, "extent")[0];
              var cx = extent ? parseInt(extent.getAttribute("cx") || "", 10) : NaN;
              var seg = {
                text: "￼",
                isImage: true,
                isDel: false,
                relId: embed,
                // English Metric Units: 9,525 to the pixel.
                widthPx: cx > 0 ? cx / 9525 : null,
                trackType: ctx.trackType || null,
                dataUrl: null,
                _rels: rels,
                _basePart: basePart,
              };
              segments.push(seg);
              if (embed) pendingImages.push(seg);
              // A drawing can carry a text box as well as a picture, and
              // its paragraphs sit below this point.
              walk(tn, ctx);
            } else if (
              tn.localName === "footnoteReference" ||
              tn.localName === "endnoteReference"
            ) {
              var noteSeg = noteRefSegment(
                {
                  kind: tn.localName === "endnoteReference" ? "en" : "fn",
                  id: tn.getAttribute("w:id"),
                },
                docId || "",
              );
              noteSeg.trackType = ctx.trackType || null;
              segments.push(noteSeg);
            } else {
              // Anything else inside a run can still hold text. A text
              // box -- VML w:pict, or the DrawingML equivalent -- holds
              // whole paragraphs, and skipping it meant a callout or a
              // sidebar was compared as though it were not in the
              // document. Its paragraphs are flattened into the
              // paragraph that anchors the box rather than becoming
              // blocks of their own, which is where they sit in the
              // reading order anyway.
              walk(tn, ctx);
            }
          });
        } else {
          walk(child, ctx);
        }
      });
    }
    walk(pEl, {});
    return {
      isHeading: isHeading,
      headingLevel: headingLevel,
      isListItem: isListItem,
      ilvl: ilvl,
      listLabel: label ? label.label : "",
      listKind: label ? label.kind : "",
      para: {
        align: para.align,
        indL: para.indL,
        indR: para.indR,
        indFirst: para.indFirst,
        before: para.before,
        after: para.after,
      },
      base: {
        bold: !!base.bold,
        italic: !!base.italic,
        color: base.color || null,
        size: base.size || null,
        font: base.font || null,
      },
      segments: segments,
    };
  }

  // Every w:p at or below `el`, in document order, without descending
  // into a paragraph itself. Used for the contents of a table cell, which
  // may hold a nested table whose paragraphs would otherwise be lost.
  function descendantParagraphs(el) {
    var out = [];
    (function walk(node) {
      Array.from(node.childNodes).forEach(function (child) {
        if (child.nodeType !== 1) return;
        if (child.localName === "p") out.push(child);
        else walk(child);
      });
    })(el);
    return out;
  }

  function extractBlocksFromContainer(
    containerEl,
    pendingImages,
    rels,
    sectionLabel,
    basePart,
    env,
  ) {
    if (!env || typeof env !== "object") env = emptyEnv(env);
    var blocks = [];

    function tableRows(tbl) {
      // The table's own column widths, for a cell that does not state one.
      var grid = [];
      var tblGrid = firstLocal(tbl, "tblGrid");
      if (tblGrid) {
        childrenLocal(tblGrid, "gridCol").forEach(function (gc) {
          grid.push(parseInt(gc.getAttribute("w:w") || "", 10) || 0);
        });
      }
      var rows = childrenLocal(tbl, "tr");
      rows.forEach(function (tr, ri) {
        var col = 0;
        var cells = childrenLocal(tr, "tc").map(function (tc) {
          var segs = [];
          var cellBase = null,
            cellAlign = null;
          // descendantParagraphs rather than direct children: a cell
          // holding a nested table used to contribute nothing but its
          // own loose paragraphs, so the nested rows vanished.
          descendantParagraphs(tc).forEach(function (p, idx) {
            if (idx > 0) segs.push({ text: "\n", isDel: false, isImage: false });
            var pd = extractParagraph(p, pendingImages, rels, basePart, env);
            if (idx === 0) {
              cellBase = pd.base;
              cellAlign = pd.para.align || null;
            }
            segs = segs.concat(pd.segments);
          });
          var tcPr = firstLocal(tc, "tcPr");
          var shd = tcPr && firstLocal(tcPr, "shd");
          var fill = shd ? hexColor(shd.getAttribute("w:fill")) : null;
          if (fill === "FFFFFF") fill = null;
          var spanEl = tcPr && firstLocal(tcPr, "gridSpan");
          var span = spanEl
            ? Math.max(1, parseInt(spanEl.getAttribute("w:val") || "1", 10) || 1)
            : 1;
          var tcW = tcPr && firstLocal(tcPr, "tcW");
          var width = 0;
          if (tcW && (tcW.getAttribute("w:type") || "dxa") === "dxa")
            width = parseInt(tcW.getAttribute("w:w") || "", 10) || 0;
          if (!width) {
            for (var g = col; g < col + span && g < grid.length; g++)
              width += grid[g];
          }
          col += span;
          return {
            segs: segs,
            fill: fill,
            width: width > 0 ? width : null,
            align: cellAlign,
            base: cellBase || {},
          };
        });
        var row = makeTableRowBlock(cells, sectionLabel);
        row.rowFirst = ri === 0;
        row.rowLast = ri === rows.length - 1;
        blocks.push(row);
      });
    }

    function visit(el) {
      Array.from(el.childNodes).forEach(function (child) {
        if (child.nodeType !== 1) return;
        var ln = child.localName;
        if (ln === "p") {
          blocks.push(
            makeParagraphBlock(
              extractParagraph(child, pendingImages, rels, basePart, env),
              sectionLabel,
            ),
          );
        } else if (ln === "tbl") {
          tableRows(child);
        } else {
          // Not a block itself, but it may contain them. Content controls
          // (w:sdt), text boxes (w:txbxContent), smart tags and custom XML
          // all sit between the body and its paragraphs. Looking only at
          // direct children meant a table of contents, a form field or a
          // whole templated section was never read at all -- and a diff
          // that never reads a section reports it as unchanged, which is
          // the one answer it must never give.
          visit(child);
        }
      });
    }

    visit(containerEl);
    return blocks;
  }

  function footnoteOrCommentBlock(
    containerEl,
    prefixSeg,
    pendingImages,
    rels,
    sectionLabel,
    basePart,
    env,
  ) {
    var segs = [prefixSeg];
    var base = null;
    childrenLocal(containerEl, "p").forEach(function (p, idx) {
      if (idx > 0) segs.push({ text: " ", isDel: false, isImage: false });
      var pd = extractParagraph(p, pendingImages, rels, basePart, env);
      if (idx === 0) base = pd.base;
      segs = segs.concat(pd.segments);
    });
    return makeParagraphBlock({ segments: segs, base: base || {} }, sectionLabel);
  }

  async function extractDocxStructured(file) {
    return extractDocxFromZip(await JSZip.loadAsync(await file.arrayBuffer()));
  }

  async function extractDocxFromZip(zip) {
    var docId = newDocId();

    async function xml(path) {
      var f = zip.file(path);
      if (!f) return null;
      var s = await f.async("string");
      return new DOMParser().parseFromString(s, "application/xml");
    }
    async function relsFor(partPath) {
      var idx = partPath.lastIndexOf("/");
      var dir = idx >= 0 ? partPath.slice(0, idx) : "";
      var name = idx >= 0 ? partPath.slice(idx + 1) : partPath;
      var relsPath = (dir ? dir + "/" : "") + "_rels/" + name + ".rels";
      var relsXmlDoc = await xml(relsPath);
      var map = {};
      if (relsXmlDoc) {
        allLocal(relsXmlDoc, "Relationship").forEach(function (r) {
          map[r.getAttribute("Id")] = {
            target: r.getAttribute("Target"),
            type: r.getAttribute("Type"),
            mode: r.getAttribute("TargetMode"),
          };
        });
      }
      return map;
    }
    function resolveTarget(basePartPath, target, mode) {
      if (
        mode === "External" ||
        /^https?:\/\//i.test(target || "") ||
        /^mailto:/i.test(target || "")
      )
        return target;
      var idx = basePartPath.lastIndexOf("/");
      var dir = idx >= 0 ? basePartPath.slice(0, idx) : "";
      var combined = (dir ? dir + "/" : "") + target;
      var parts = combined.split("/"),
        stack = [];
      parts.forEach(function (p) {
        if (p === "" || p === ".") return;
        if (p === "..") stack.pop();
        else stack.push(p);
      });
      return stack.join("/");
    }

    var docXml = await xml("word/document.xml");
    if (!docXml)
      throw new Error("Missing word/document.xml — not a valid .docx");
    var docRels = await relsFor("word/document.xml");

    // The fonts a theme stands for, the styles, and the lists: together
    // they are what turns "this paragraph is a Heading 1" into a size, a
    // weight and a colour.
    var theme = null;
    var themeXml = await xml("word/theme/theme1.xml");
    if (themeXml) {
      var major = allLocal(themeXml, "majorFont")[0],
        minor = allLocal(themeXml, "minorFont")[0];
      var majorLatin = major && firstLocal(major, "latin"),
        minorLatin = minor && firstLocal(minor, "latin");
      theme = {
        major: majorLatin ? majorLatin.getAttribute("typeface") : null,
        minor: minorLatin ? minorLatin.getAttribute("typeface") : null,
      };
    }
    var env = {
      docId: docId,
      theme: theme,
      styles: readStyles(await xml("word/styles.xml"), theme),
      numbering: readNumbering(await xml("word/numbering.xml")),
      nextLabel: listCounter(),
    };

    var pendingImages = [];
    var blocks = [];

    var headerParts = [],
      footerParts = [];
    Object.keys(docRels).forEach(function (id) {
      var r = docRels[id];
      if (!r.type) return;
      if (/\/header$/.test(r.type))
        headerParts.push(resolveTarget("word/document.xml", r.target, r.mode));
      if (/\/footer$/.test(r.type))
        footerParts.push(resolveTarget("word/document.xml", r.target, r.mode));
    });

    for (var hi = 0; hi < headerParts.length; hi++) {
      var hp = headerParts[hi];
      var hXml = await xml(hp);
      if (!hXml) continue;
      var hRels = await relsFor(hp);
      blocks.push(makeDividerBlock("HEADER"));
      blocks = blocks.concat(
        extractBlocksFromContainer(
          hXml.documentElement,
          pendingImages,
          hRels,
          "header",
          hp,
          env,
        ),
      );
    }

    var body = allLocal(docXml, "body")[0];
    if (body)
      blocks = blocks.concat(
        extractBlocksFromContainer(
          body,
          pendingImages,
          docRels,
          "body",
          "word/document.xml",
          env,
        ),
      );

    for (var fi = 0; fi < footerParts.length; fi++) {
      var fp = footerParts[fi];
      var fXml = await xml(fp);
      if (!fXml) continue;
      var fRels = await relsFor(fp);
      blocks.push(makeDividerBlock("FOOTER"));
      blocks = blocks.concat(
        extractBlocksFromContainer(
          fXml.documentElement,
          pendingImages,
          fRels,
          "footer",
          fp,
          env,
        ),
      );
    }

    // Footnotes and endnotes are the same thing in two parts of the
    // package, down to the separator entries that are not notes at all.
    var NOTE_PARTS = [
      { path: "word/footnotes.xml", el: "footnote", kind: "fn", label: "FOOTNOTES", section: "footnote" },
      { path: "word/endnotes.xml", el: "endnote", kind: "en", label: "ENDNOTES", section: "endnote" },
    ];
    for (var ni = 0; ni < NOTE_PARTS.length; ni++) {
      var np = NOTE_PARTS[ni];
      var notesXml = await xml(np.path);
      if (!notesXml) continue;
      var nRels = await relsFor(np.path);
      var notes = childrenLocal(notesXml.documentElement, np.el).filter(
        function (f) {
          var t = f.getAttribute("w:type");
          return (
            t !== "separator" &&
            t !== "continuationSeparator" &&
            t !== "continuationNotice"
          );
        },
      );
      if (!notes.length) continue;
      blocks.push(makeDividerBlock(np.label));
      for (var nj = 0; nj < notes.length; nj++) {
        blocks.push(
          footnoteOrCommentBlock(
            notes[nj],
            noteMarkerSegment(
              { kind: np.kind, id: notes[nj].getAttribute("w:id") },
              docId,
            ),
            pendingImages,
            nRels,
            np.section,
            np.path,
            env,
          ),
        );
      }
    }

    var commentsXml = await xml("word/comments.xml");
    if (commentsXml) {
      var cRels = await relsFor("word/comments.xml");
      var cs = childrenLocal(commentsXml.documentElement, "comment");
      if (cs.length) {
        blocks.push(makeDividerBlock("COMMENTS"));
        cs.forEach(function (c) {
          var author = c.getAttribute("w:author") || "Unknown";
          blocks.push(
            footnoteOrCommentBlock(
              c,
              {
                text: "Comment (" + author + "): ",
                bold: true,
                isDel: false,
                isImage: false,
              },
              pendingImages,
              cRels,
              "comment",
              "word/comments.xml",
              env,
            ),
          );
        });
      }
    }

    await Promise.all(
      pendingImages.map(async function (seg) {
        if (!seg.relId || !seg._rels || !seg._rels[seg.relId]) return;
        var r = seg._rels[seg.relId];
        var path = resolveTarget(seg._basePart, r.target, r.mode);
        var f = zip.file(path);
        if (!f) return;
        var base64 = await f.async("base64");
        var ext = (path.split(".").pop() || "png").toLowerCase();
        var mime =
          ext === "jpg" || ext === "jpeg"
            ? "image/jpeg"
            : ext === "gif"
              ? "image/gif"
              : ext === "bmp"
                ? "image/bmp"
                : "image/png";
        seg.dataUrl = "data:" + mime + ";base64," + base64;
        seg.imgFingerprint = base64.length + ":" + base64.slice(-16);
      }),
    );

    return finalizeBlocks(blocks, docId);
  }

  // ============================== OPENING A FILE ==============================
  //
  // .docx is read above. The other kinds -- OpenDocument, RTF, the binary
  // .doc, plain text, HTML -- each have a reader in js/formats/, which
  // registers itself on window.diffPadFormats when its script loads.
  //
  // What a file is, is decided by its first bytes wherever they say: a
  // ".doc" is very often RTF, or a .docx somebody renamed, and each of
  // those is opened as what it is. The name only has to be one this page
  // knows, so that a photograph or a spreadsheet is turned away with a
  // message rather than read as text.

  var KNOWN_NAME =
    /\.(docx|docm|dotx|dotm|doc|dot|rtf|odt|ott|fodt|txt|text|md|markdown|html?|xhtml)$/i;
  var ACCEPTED = "Word (.docx, .doc), RTF, OpenDocument (.odt), text or HTML";

  function userError(message) {
    var e = new Error(message);
    e.userMessage = message;
    return e;
  }

  function sniffFormat(bytes) {
    if (
      bytes.length > 3 &&
      bytes[0] === 0x50 &&
      bytes[1] === 0x4b &&
      (bytes[2] === 3 || bytes[2] === 5 || bytes[2] === 7)
    )
      return "zip";
    var OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    if (
      bytes.length > 8 &&
      OLE.every(function (b, i) {
        return bytes[i] === b;
      })
    )
      return "ole";
    // Past a byte-order mark and any blank space, what does the text open with?
    var head = "";
    for (var i = 0; i < bytes.length && head.length < 256; i++) {
      var c = bytes[i];
      if (c === 0 || c === 0xef || c === 0xbb || c === 0xbf || c === 0xff || c === 0xfe)
        continue;
      head += String.fromCharCode(c);
    }
    head = head.replace(/^\s+/, "");
    if (/^\{\\rtf/.test(head)) return "rtf";
    if (/<office:document[\s>]/.test(head)) return "fodt";
    if (/^(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*)*<(!doctype\s+html|html[\s>])/i.test(head))
      return "html";
    return null;
  }

  // Text of unknown encoding: a byte-order mark says so outright; failing
  // that it is UTF-8 if it is valid UTF-8, and the Windows code page that
  // older files were written in if it is not.
  function decodeText(bytes) {
    function dec(label, from, fatal) {
      return new TextDecoder(label, { fatal: !!fatal }).decode(
        from ? bytes.subarray(from) : bytes,
      );
    }
    if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe)
      return dec("utf-16le", 2);
    if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff)
      return dec("utf-16be", 2);
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)
      return dec("utf-8", 3);
    try {
      return dec("utf-8", 0, true);
    } catch (e) {
      return dec("windows-1252");
    }
  }

  function base64FromBytes(bytes) {
    var out = "";
    for (var i = 0; i < bytes.length; i += 0x6000) {
      out += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x6000));
    }
    return btoa(out);
  }

  // Enough of a picture's bytes to tell two pictures apart: the same form
  // the .docx reader uses, so one picture has one fingerprint.
  function imageFromBytes(bytes, mime) {
    var base64 = base64FromBytes(bytes);
    return {
      dataUrl: "data:" + mime + ";base64," + base64,
      fingerprint: base64.length + ":" + base64.slice(-16),
    };
  }

  // What a picture is, from how its bytes begin; null if it is not a kind
  // a browser can draw.
  function imageMime(bytes) {
    if (bytes.length > 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
      return "image/png";
    if (bytes.length > 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff)
      return "image/jpeg";
    if (bytes.length > 6 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x38)
      return "image/gif";
    if (bytes.length > 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
    return null;
  }

  // What every reader is given to work with.
  var READER_TOOLS = {
    formatNumber: formatNumber,
    bulletGlyph: bulletGlyph,
    listCounter: listCounter,
    hexColor: hexColor,
    decodeText: decodeText,
    imageFromBytes: imageFromBytes,
    imageMime: imageMime,
  };

  async function readDocument(file) {
    var buffer = await file.arrayBuffer();
    var bytes = new Uint8Array(buffer);
    var formats = window.diffPadFormats || {};
    var name = String(file.name || "");
    var kind = sniffFormat(bytes);

    function read(format, input) {
      if (typeof formats[format] !== "function")
        throw new Error("no reader is loaded for " + format);
      var docId = newDocId();
      return Promise.resolve(formats[format](input, READER_TOOLS)).then(
        function (raw) {
          return finalizeBlocks(blocksFromRaw(raw, docId), docId);
        },
      );
    }

    // Tell the two kinds of package apart by what is inside.
    if (kind === "zip" || /\.(docx|docm|dotx|dotm|odt|ott)$/i.test(name)) {
      var zip = await JSZip.loadAsync(buffer);
      if (!zip.file("word/document.xml") && zip.file("content.xml"))
        return read("odt", zip);
      return extractDocxFromZip(zip);
    }
    if (kind === "ole") return read("doc", bytes);
    if (kind === "rtf") return read("rtf", bytes);
    if (kind === "fodt" || /\.fodt$/i.test(name)) return read("fodt", decodeText(bytes));
    if (kind === "html" || /\.(html?|xhtml)$/i.test(name))
      return read("html", decodeText(bytes));
    if (/\.(doc|dot|rtf)$/i.test(name)) {
      // Named as a Word or RTF document and not one. A few really are plain
      // text; anything with binary in it is not something to show as text.
      for (var i = 0; i < bytes.length && i < 4096; i++) {
        if (bytes[i] === 0) throw new Error("not a document this page can read");
      }
    }
    return read("text", decodeText(bytes));
  }

  // ============================== DIFF ENGINE ==============================

  // Split on runs of whitespace, and on each of the two marked tokens.
  // Both markers are paired, so a token ends where its own closing marker
  // is rather than where the next space happens to fall -- which is what
  // lets an image sit directly against a word without swallowing it.
  var TOKEN_SPLIT =
    /(\s+|\uFFFC[^\uFFFC\uFFFD]*\uFFFC|\uFFFD[^\uFFFC\uFFFD]*\uFFFD)/;

  function tokenizeWords(text) {
    return text.split(TOKEN_SPLIT).filter(function (t) {
      return t.length > 0;
    });
  }

  // How the diff is computed depends on how big the problem is.
  //
  // The table below fills an (n+1) x (m+1) grid of 32-bit cells, so its
  // cost is the product of the two lengths: 2,000 blocks a side is 16 MB
  // and about half a second, 10,000 is 400 MB, and 20,000 is 1.6 GB and
  // takes the tab with it. Blocks are paragraphs and table rows, so that
  // is a book or a long table rather than a report -- but the same
  // function also diffs the words of a single paragraph against each
  // other, and a paragraph of 20,000 words reaches it far sooner.
  //
  // Rather than cap the input and refuse to answer, anything past the
  // limit goes to Myers' algorithm, whose cost is driven by the number of
  // differences rather than by the size of the input. Two revisions of the
  // same long document differ in very little, so that path is fast exactly
  // when the table is unaffordable. Everything below the limit still takes
  // the table, unchanged, so the output for ordinary documents is
  // identical to what it has always been.
  var LCS_TABLE_CELL_LIMIT = 4e6; // ~16 MB of Int32Array

  // generic LCS-based diff; eq defaults to === ; ops store {type, a, b}
  function lcsDiff(a, b, eq) {
    eq =
      eq ||
      function (x, y) {
        return x === y;
      };
    if ((a.length + 1) * (b.length + 1) <= LCS_TABLE_CELL_LIMIT) {
      return lcsDiffTable(a, b, eq);
    }
    return lcsDiffLarge(a, b, eq);
  }

  function lcsDiffTable(a, b, eq) {
    var n = a.length,
      m = b.length;
    var dp = new Array(n + 1);
    for (var i = 0; i <= n; i++) dp[i] = new Int32Array(m + 1);
    for (var i = n - 1; i >= 0; i--) {
      for (var j = m - 1; j >= 0; j--) {
        dp[i][j] = eq(a[i], b[j])
          ? dp[i + 1][j + 1] + 1
          : Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
    var ops = [];
    var i = 0,
      j = 0;
    while (i < n && j < m) {
      if (eq(a[i], b[j])) {
        ops.push({ type: "equal", a: a[i], b: b[j] });
        i++;
        j++;
      } else if (dp[i + 1][j] >= dp[i][j + 1]) {
        ops.push({ type: "delete", a: a[i] });
        i++;
      } else {
        ops.push({ type: "insert", b: b[j] });
        j++;
      }
    }
    while (i < n) {
      ops.push({ type: "delete", a: a[i] });
      i++;
    }
    while (j < m) {
      ops.push({ type: "insert", b: b[j] });
      j++;
    }
    return ops;
  }

  // Caps for the large path. MYERS_MAX_D bounds the memory kept for the
  // backtrack (about (D+1)^2 ints, so 16 MB at 2,000); MYERS_MAX_WORK
  // bounds the wall clock by counting comparisons, because the running
  // time depends on the input as well as on the number of differences.
  // Past either, the two sequences have so little in common that pairing
  // them up element by element says as much as an exact answer would, and
  // says it immediately.
  var MYERS_MAX_D = 2000;
  var MYERS_MAX_WORK = 2e7;

  function lcsDiffLarge(a, b, eq) {
    // Two revisions of one document usually share a long head and tail.
    // Trimming those first is what makes the middle small enough to be
    // worth an exact algorithm at all.
    var head = 0;
    var aEnd = a.length,
      bEnd = b.length;
    while (head < aEnd && head < bEnd && eq(a[head], b[head])) head++;
    while (aEnd > head && bEnd > head && eq(a[aEnd - 1], b[bEnd - 1])) {
      aEnd--;
      bEnd--;
    }

    var ops = [];
    for (var i = 0; i < head; i++) {
      ops.push({ type: "equal", a: a[i], b: b[i] });
    }

    var midA = a.slice(head, aEnd),
      midB = b.slice(head, bEnd);
    var mid = myersDiff(midA, midB, eq);
    if (mid === null) {
      // Out of budget: emit the middle as a wholesale replacement. Still
      // a correct edit script -- it reproduces both sides -- just not the
      // shortest one.
      for (var d = 0; d < midA.length; d++) {
        ops.push({ type: "delete", a: midA[d] });
      }
      for (var s = 0; s < midB.length; s++) {
        ops.push({ type: "insert", b: midB[s] });
      }
    } else {
      for (var o = 0; o < mid.length; o++) ops.push(mid[o]);
    }

    for (var t = aEnd; t < a.length; t++) {
      ops.push({ type: "equal", a: a[t], b: b[bEnd + (t - aEnd)] });
    }
    return ops;
  }

  // Myers' 1986 greedy shortest-edit-script search. It walks outwards by
  // edit distance D, keeping for each D the furthest point reached on
  // every diagonal, and stops as soon as one of them reaches the end --
  // so a pair that differs in ten places costs ten passes, whatever the
  // length of the input. Returns null if it runs past its budget.
  function myersDiff(a, b, eq) {
    var n = a.length,
      m = b.length;
    if (n === 0 || m === 0) {
      var only = [];
      for (var q = 0; q < n; q++) only.push({ type: "delete", a: a[q] });
      for (var w = 0; w < m; w++) only.push({ type: "insert", b: b[w] });
      return only;
    }

    var limit = Math.min(n + m, MYERS_MAX_D);
    var offset = limit + 1;
    var v = new Int32Array(2 * limit + 3);
    var trace = [];
    var work = 0;

    for (var d = 0; d <= limit; d++) {
      // Snapshot the diagonals reachable at this depth before extending
      // them; the backtrack replays these to recover the actual script.
      trace.push(v.slice(offset - d, offset + d + 1));
      for (var k = -d; k <= d; k += 2) {
        var x;
        if (k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1])) {
          x = v[offset + k + 1]; // came from above: an insertion
        } else {
          x = v[offset + k - 1] + 1; // came from the left: a deletion
        }
        var y = x - k;
        while (x < n && y < m && eq(a[x], b[y])) {
          x++;
          y++;
          work++;
        }
        work++;
        v[offset + k] = x;
        if (x >= n && y >= m) return myersBacktrack(a, b, trace, d);
        if (work > MYERS_MAX_WORK) return null;
      }
    }
    return null;
  }

  function myersBacktrack(a, b, trace, d) {
    var ops = [];
    var x = a.length,
      y = b.length;

    for (var step = d; step >= 0; step--) {
      var snap = trace[step]; // covers diagonals -step .. step at index k+step
      var k = x - y;
      var prevK;
      if (
        k === -step ||
        (k !== step && snap[k - 1 + step] < snap[k + 1 + step])
      ) {
        prevK = k + 1;
      } else {
        prevK = k - 1;
      }
      var prevX = step === 0 ? 0 : snap[prevK + step];
      var prevY = prevX - prevK;

      while (x > prevX && y > prevY) {
        x--;
        y--;
        ops.push({ type: "equal", a: a[x], b: b[y] });
      }
      if (step > 0) {
        if (x === prevX) {
          y--;
          ops.push({ type: "insert", b: b[y] });
        } else {
          x--;
          ops.push({ type: "delete", a: a[x] });
        }
      }
      x = prevX;
      y = prevY;
    }

    ops.reverse();
    return ops;
  }

  function buildRows(ops) {
    var rows = [];
    var delBuf = [],
      insBuf = [];

    function flush() {
      var len = Math.max(delBuf.length, insBuf.length);
      for (var k = 0; k < len; k++) {
        var d = delBuf[k],
          ins = insBuf[k];
        if (d !== undefined && ins !== undefined) {
          var wordOps = lcsDiff(tokenizeWords(d.text), tokenizeWords(ins.text));
          var row = {
            kind: "modified",
            left: d,
            right: ins,
            wordOps: wordOps,
          };
          // Two rows of a table with the same number of cells are compared
          // cell by cell, so the result can be drawn as a row of a table.
          if (
            d.isTableRow &&
            ins.isTableRow &&
            d.cells &&
            ins.cells &&
            d.cells.length === ins.cells.length
          ) {
            row.cellOps = d.cells.map(function (c, ci) {
              return lcsDiff(
                tokenizeWords(plainTextForDiff(c.segs)),
                tokenizeWords(plainTextForDiff(ins.cells[ci].segs)),
              );
            });
          }
          rows.push(row);
        } else if (d !== undefined) {
          rows.push({ kind: "removed", left: d, right: null });
        } else {
          rows.push({ kind: "added", left: null, right: ins });
        }
      }
      delBuf = [];
      insBuf = [];
    }

    ops.forEach(function (op) {
      if (op.type === "equal") {
        flush();
        var formatChanged = op.a.fmtKey !== op.b.fmtKey;
        rows.push({
          kind: formatChanged ? "format-changed" : "equal",
          left: op.a,
          right: op.b,
        });
      } else if (op.type === "delete") {
        delBuf.push(op.a);
      } else {
        insBuf.push(op.b);
      }
    });
    flush();
    return rows;
  }

  function renderImageToken(token, block) {
    if (!block || !block.imageTokens || token.charAt(0) !== "￼")
      return null;
    var seg = block.imageTokens[token];
    if (!seg) return null;
    return '<span class="inline-img">' + imageTagHtml(seg) + "</span>";
  }

  function renderFootnoteToken(token, block) {
    if (!block || !block.footnoteTokens || token.charAt(0) !== "�")
      return null;
    var seg = block.footnoteTokens[token];
    if (!seg) return null;
    if (seg.isFootnoteRef) return footnoteRefHtml(seg);
    if (seg.isFootnoteMarker) return footnoteMarkerHtml(seg);
    return null;
  }

  // The word diff of one block, drawn for one side. `leftBlock` and
  // `rightBlock` are what the two token streams came from: they turn a
  // marked token back into its picture or footnote marker, and -- where
  // they carry the positions of their runs -- give each word back the
  // bold, the colour or the link of the run it was taken from.
  function renderWordDiff(ops, side, leftBlock, rightBlock) {
    var own = side === "left" ? leftBlock : rightBlock;
    var spans = (own && own.spans) || null;
    var base = (own && own.base) || null;
    var pos = 0; // how far into this side's plain text the walk has got
    var si = 0;

    function plain(text, visibleWs) {
      return visibleWs ? escapeHtmlVisibleWs(text) : escapeHtml(text);
    }

    // `text` is the token at [pos, pos + text.length) on this side.
    function formatted(text, visibleWs) {
      if (!spans) return plain(text, visibleWs);
      var start = pos,
        end = pos + text.length;
      while (si < spans.length && spans[si].end <= start) si++;
      var out = "";
      var at = start;
      var k = si;
      while (at < end) {
        var sp = spans[k];
        if (!sp || sp.start > at) {
          // Between runs: an image, a footnote, the bar between two cells.
          var gapEnd = sp ? Math.min(sp.start, end) : end;
          out += plain(text.slice(at - start, gapEnd - start), visibleWs);
          at = gapEnd;
          continue;
        }
        var stop = Math.min(sp.end, end);
        out += wrapRunHtml(
          sp.seg,
          plain(text.slice(at - start, stop - start), visibleWs),
          base,
        );
        at = stop;
        if (stop === sp.end) k++;
      }
      return out;
    }

    var html = "";
    ops.forEach(function (o) {
      if (o.type === "equal") {
        var eqImg =
          renderImageToken(o.a, leftBlock) || renderImageToken(o.a, rightBlock);
        var eqFn =
          eqImg === null
            ? renderFootnoteToken(o.a, leftBlock) ||
              renderFootnoteToken(o.a, rightBlock)
            : null;
        html += eqImg || eqFn || formatted(o.a, false);
        pos += o.a.length;
      } else if (o.type === "delete" && side === "left") {
        var delImg = renderImageToken(o.a, leftBlock);
        var delFn =
          delImg === null ? renderFootnoteToken(o.a, leftBlock) : null;
        if (delImg) {
          html += '<span class="del">' + delImg + "</span>";
        } else if (delFn) {
          html += '<span class="del">' + delFn + "</span>";
        } else {
          var isWs = o.a.trim() === "" && o.a.length > 0;
          html += '<span class="del">' + formatted(o.a, isWs) + "</span>";
        }
        pos += o.a.length;
      } else if (o.type === "insert" && side === "right") {
        var insImg = renderImageToken(o.b, rightBlock);
        var insFn =
          insImg === null ? renderFootnoteToken(o.b, rightBlock) : null;
        if (insImg) {
          html += '<span class="ins">' + insImg + "</span>";
        } else if (insFn) {
          html += '<span class="ins">' + insFn + "</span>";
        } else {
          var isWs2 = o.b.trim() === "" && o.b.length > 0;
          html += '<span class="ins">' + formatted(o.b, isWs2) + "</span>";
        }
        pos += o.b.length;
      }
    });
    return html;
  }

  function sectionBadge(block) {
    if (
      !block ||
      !block.sectionLabel ||
      block.sectionLabel === "body" ||
      block.sectionLabel === "divider"
    )
      return "";
    return (
      '<span class="sec-badge">' + escapeHtml(block.sectionLabel) + "</span>"
    );
  }

  var FMT_TAG = '<span class="fmt-tag">formatting changed</span>';

  function styleAttr(css) {
    return css ? ' style="' + css + '"' : "";
  }

  // Put a block's content into the element that makes it look like what it
  // is: a paragraph, a heading, a list item with its number, a row of a
  // table. `lead` and `tail` are the page's own labels -- which part of the
  // document this is, that its formatting changed -- and go inside a
  // paragraph but outside a table row.
  function wrapBlockContent(block, innerHtml, lead, tail) {
    lead = lead || "";
    tail = tail || "";
    if (!block) return innerHtml;
    if (block.isDivider)
      return '<div class="block-divider"><span>' + innerHtml + "</span></div>";
    if (block.isTableRow) {
      // No "formatting changed" label on a row: there is nowhere to put
      // one that does not break the table apart, and the row's colour
      // already says it.
      return (
        lead +
        '<div class="row-cells' +
        (block.rowFirst ? " row-first" : "") +
        (block.rowLast ? " row-last" : "") +
        '">' +
        innerHtml +
        "</div>"
      );
    }
    var css = styleAttr(block.paraCss || "");
    if (block.isListItem) {
      var label =
        block.listLabel ||
        (block.ilvl >= 2 ? "▪" : block.ilvl === 1 ? "◦" : "•");
      return (
        '<div class="block-li' +
        (block.isHeading
          ? " block-heading h" +
            block.headingLevel +
            (block.base && block.base.size ? " styled" : "")
          : "") +
        '"' +
        css +
        '><span class="li-bullet' +
        (block.listKind === "number" ? " li-number" : "") +
        '" style="margin-left:' +
        block.ilvl * 18 +
        'pt;">' +
        escapeHtml(label) +
        '</span><span class="li-text">' +
        lead +
        innerHtml +
        tail +
        "</span></div>"
      );
    }
    if (block.isHeading)
      return (
        '<div class="block-heading h' +
        block.headingLevel +
        (block.base && block.base.size ? " styled" : "") +
        '"' +
        css +
        ">" +
        lead +
        innerHtml +
        tail +
        "</div>"
      );
    return '<div class="block-p"' + css + ">" + lead + innerHtml + tail + "</div>";
  }

  // A changed table row, drawn as a row: each cell with its own word diff.
  function renderRowDiff(cellOps, side, leftBlock, rightBlock) {
    var own = side === "left" ? leftBlock : rightBlock;
    function cellView(block, ci) {
      return {
        spans: block.cells[ci].spans,
        base: block.cells[ci].base,
        imageTokens: block.imageTokens,
        footnoteTokens: block.footnoteTokens,
      };
    }
    return own.cells
      .map(function (c, ci) {
        return rowCellHtml(
          c,
          renderWordDiff(
            cellOps[ci],
            side,
            cellView(leftBlock, ci),
            cellView(rightBlock, ci),
          ),
        );
      })
      .join("");
  }

  function renderCellForSide(
    block,
    kind,
    side,
    wordOps,
    leftBlock,
    rightBlock,
    cellOps,
    edge,
  ) {
    // `edge` is "table-start" or "table-end" on the row of the diff where a
    // table begins or has just ended.
    edge = edge ? " " + edge : "";
    if (!block) return '<div class="cell gap' + edge + '"></div>';
    var badge = sectionBadge(block);
    var innerHtml,
      cellClass,
      tail = "";
    if (kind === "modified") {
      var aTable = !!(leftBlock && leftBlock.isTableRow),
        bTable = !!(rightBlock && rightBlock.isTableRow);
      if (cellOps) {
        innerHtml = renderRowDiff(cellOps, side, leftBlock, rightBlock);
        cellClass = "cell modified";
      } else if (aTable || bTable) {
        // A row against a paragraph, or against a row with a different
        // number of cells: there is no word-by-word answer that could be
        // drawn as either, so each side is shown whole.
        innerHtml = block.html;
        cellClass = side === "left" ? "cell removed" : "cell added";
      } else {
        innerHtml = renderWordDiff(wordOps, side, leftBlock, rightBlock);
        cellClass = "cell modified";
      }
    } else if (kind === "removed") {
      innerHtml = block.html;
      cellClass = "cell removed";
    } else if (kind === "added") {
      innerHtml = block.html;
      cellClass = "cell added";
    } else if (kind === "format-changed") {
      innerHtml = block.html;
      tail = FMT_TAG;
      cellClass = "cell fmt-changed";
    } else {
      innerHtml = block.html;
      cellClass = "cell";
    }
    cellClass += edge;
    if (block.isTableRow) cellClass += " is-row";
    if (block.sectionLabel && block.sectionLabel !== "body")
      cellClass += " sec-" + block.sectionLabel.replace(/[^a-z]/g, "");
    return (
      '<div class="' +
      cellClass +
      '">' +
      wrapBlockContent(block, innerHtml, badge, tail) +
      "</div>"
    );
  }

  function renderBlockPreview(block, i, all) {
    var prev = all && i > 0 ? all[i - 1] : null;
    var wasRow = !!(prev && prev.isTableRow);
    var edge =
      block.isTableRow && !wasRow ? "table-start" : !block.isTableRow && wasRow ? "table-end" : "";
    return renderCellForSide(block, "equal", "left", null, block, block, null, edge);
  }

  // ============================== UI ==============================
  //
  // The page is a Windows 95 desktop with one diffPad window per document.
  // The desktop, taskbar, Start menu and About box are static markup in
  // index.html; each window -- title bar, menu bar, page and status bar -- is
  // drawn here on every render.

  var ROLE = { left: "Original", right: "Updated" };

  // Which window has the navy title bar. Win95 shows exactly one active
  // window; clicking a window or its taskbar button moves it.
  var activeSide = "left";

  var ICON_DIFFPAD =
    '<svg class="icon16" viewBox="0 0 16 16" shape-rendering="crispEdges" aria-hidden="true">' +
    '<path d="M2.5 1.5h8l3 3v10h-11z" fill="#fff" stroke="#000"/>' +
    '<path d="M10.5 1.5v3h3" fill="#c0c0c0" stroke="#000"/>' +
    '<path d="M4 6.5h8M4 8.5h8M4 10.5h8M4 12.5h6" stroke="#000080"/></svg>';

  var ICON_SWAP =
    '<svg viewBox="0 0 12 10" shape-rendering="crispEdges" aria-hidden="true">' +
    '<path d="M0 2h9V0h1v1h1v1h1v1h-1v1h-1v1H9V3H0z"/>' +
    '<path d="M12 7H3V5H2v1H1v1H0v1h1v1h1v1h1V8h9z"/></svg>';

  var ICON_DOC_BIG =
    '<svg class="icon32" viewBox="0 0 32 32" shape-rendering="crispEdges" aria-hidden="true">' +
    '<path d="M6.5 2.5h14l6 6v21h-20z" fill="#fff" stroke="#000"/>' +
    '<path d="M20.5 2.5v6h6" fill="#c0c0c0" stroke="#000"/>' +
    '<path d="M9 12.5h14M9 15.5h14M9 18.5h14M9 21.5h14M9 24.5h9" stroke="#000080"/></svg>';

  function windowTitle(side) {
    var doc = state[side];
    return (doc ? escapeHtml(doc.fileName) : "Untitled") + " - diffPad";
  }

  function swapButtonHtml(enabled) {
    return (
      '<button class="title-btn swap-btn" title="Swap original and updated" aria-label="Swap original and updated"' +
      (enabled ? "" : " disabled") +
      ">" +
      ICON_SWAP +
      "</button>"
    );
  }

  function menuItemHtml(action, label, side, enabled) {
    return (
      '<button class="menu-item" role="menuitem" data-action="' +
      action +
      '" data-for="' +
      side +
      '"' +
      (enabled ? "" : " disabled") +
      ">" +
      label +
      "</button>"
    );
  }

  function menuHtml(label, items) {
    return (
      '<div class="menu">' +
      '<button class="menu-title" aria-haspopup="true">' +
      label +
      "</button>" +
      '<div class="menu-list" role="menu">' +
      items
        .map(function (it) {
          return it === "-" ? '<div class="menu-sep"></div>' : it;
        })
        .join("") +
      "</div></div>"
    );
  }

  // Notepad's bare menu bar: no toolbar, no format bar, no ruler.
  function menuBarHtml(side) {
    var ready = !!(state.left && state.right);
    return (
      '<div class="menubar" role="menubar">' +
      menuHtml("<u>F</u>ile", [
        menuItemHtml("open", "<u>O</u>pen…", side, true),
        menuItemHtml("close", "<u>C</u>lose", side, !!state[side]),
        "-",
        menuItemHtml("export", "<u>E</u>xport diff…", side, ready),
      ]) +
      menuHtml("<u>H</u>elp", [
        menuItemHtml("about", "<u>A</u>bout diffPad", side, true),
      ]) +
      "</div>"
    );
  }

  function statusBarHtml(side) {
    var other = side === "left" ? "right" : "left";
    if (state.left && state.right) {
      return side === "left"
        ? '<span class="sb-main"><i class="swatch removed"></i><span id="removed-count">0</span> removed</span>' +
            '<span class="sb-cell">' +
            ROLE[side] +
            "</span>"
        : '<span class="sb-main"><i class="swatch added"></i><span id="added-count">0</span> added or changed' +
            '<i class="swatch fmt"></i><span id="fmt-count">0</span> formatting only</span>' +
            '<span class="sb-cell">' +
            ROLE[side] +
            "</span>";
    }
    var text = state[side]
      ? "Waiting for the " + ROLE[other].toLowerCase() + " document"
      : "For Help, press F1";
    return (
      '<span class="sb-main">' +
      text +
      "</span>" +
      '<span class="sb-cell">' +
      ROLE[side] +
      "</span>"
    );
  }

  // A loaded window has no dropzone, so a file opened into it from the File
  // menu reports its progress and its errors here instead.
  function windowMessagesHtml() {
    return (
      '<div class="dz-loading win-loading" style="display:none;"><span class="progress"></span>Reading document…</div>' +
      '<div class="dz-error" style="display:none;"></div>'
    );
  }

  function windowHtml(side, pageHtml) {
    return (
      '<section class="window' +
      (side === activeSide ? "" : " inactive") +
      '" data-win="' +
      side +
      '">' +
      '<div class="titlebar">' +
      ICON_DIFFPAD +
      '<span class="title-text">' +
      windowTitle(side) +
      "</span>" +
      swapButtonHtml(!!(state.left || state.right)) +
      "</div>" +
      menuBarHtml(side) +
      '<div class="page">' +
      pageHtml +
      "</div>" +
      '<div class="statusbar">' +
      statusBarHtml(side) +
      "</div>" +
      (state[side] ? windowMessagesHtml() : "") +
      "</section>"
    );
  }

  function renderTasks() {
    var tasks = document.getElementById("tasks");
    if (!tasks) return;
    tasks.innerHTML = ["left", "right"]
      .map(function (side) {
        var name = state[side] ? escapeHtml(state[side].fileName) : "Untitled";
        return (
          '<button class="task' +
          (side === activeSide ? " on" : "") +
          '" data-task="' +
          side +
          '" title="' +
          ROLE[side] +
          '">' +
          ICON_DIFFPAD +
          '<span class="task-text">' +
          name +
          " - diffPad</span></button>"
        );
      })
      .join("");
  }

  function setActive(side) {
    if (side !== "left" && side !== "right") return;
    activeSide = side;
    if (!app.querySelectorAll) return;
    app.querySelectorAll(".window").forEach(function (w) {
      w.classList.toggle("inactive", w.getAttribute("data-win") !== side);
    });
    var tasks = document.getElementById("tasks");
    if (tasks && tasks.querySelectorAll) {
      tasks.querySelectorAll(".task").forEach(function (t) {
        t.classList.toggle("on", t.getAttribute("data-task") === side);
      });
    }
  }

  function setMode() {
    var ready = !!(state.left && state.right);
    document.body.classList.toggle("mode-setup", !ready);
    document.body.classList.toggle("mode-diff", ready);
  }

  // Start over lives in the taskbar's Start menu, outside the windows.
  function updateGlobalCommands() {
    if (resetBtn) resetBtn.disabled = !(state.left || state.right);
  }

  function renderSetup() {
    setMode();

    function pageFor(side) {
      if (!state[side]) return dropzoneHtml(side);
      return (
        '<div class="page-scroll"><div class="doc-preview doc-sheet">' +
        state[side].blocks.map(renderBlockPreview).join("") +
        "</div></div>"
      );
    }

    app.innerHTML =
      '<div class="desk">' +
      '<div class="frames">' +
      windowHtml("left", pageFor("left")) +
      windowHtml("right", pageFor("right")) +
      "</div></div>";
    wireDropzones();
  }

  function dropzoneHtml(side) {
    var label = side === "left" ? "original" : "updated";
    return (
      '<div class="dropzone" data-side="' +
      side +
      '">' +
      ICON_DOC_BIG +
      '<p class="dz-title">Drag a document here or click to browse</p>' +
      '<p class="dz-sub">Open the ' +
      label +
      " document</p>" +
      '<p class="dz-kinds">.docx &middot; .doc &middot; .rtf &middot; .odt &middot; .txt &middot; .html</p>' +
      '<span class="btn95 dz-open"><span class="focus-ring">Open…</span></span>' +
      '<div class="dz-loading" style="display:none;"><span class="progress"></span>Reading document…</div>' +
      '<div class="dz-error" style="display:none;"></div>' +
      "</div>"
    );
  }

  function wireDropzones() {
    document.querySelectorAll(".dropzone").forEach(function (dz) {
      var side = dz.getAttribute("data-side");
      var input = side === "left" ? inputLeft : inputRight;

      dz.addEventListener("click", function (e) {
        // The error box sits inside the dropzone. Clicking it dismisses it
        // rather than opening the picker again underneath.
        if (
          e &&
          e.target &&
          e.target.closest &&
          e.target.closest(".dz-error")
        ) {
          hideError(dz);
          return;
        }
        input.click();
      });
      dz.addEventListener("dragover", function (e) {
        e.preventDefault();
        dz.classList.add("drag");
      });
      dz.addEventListener("dragleave", function () {
        dz.classList.remove("drag");
      });
      dz.addEventListener("drop", function (e) {
        e.preventDefault();
        dz.classList.remove("drag");
        if (e.dataTransfer.files && e.dataTransfer.files[0])
          handleFile(side, e.dataTransfer.files[0], dz);
      });
    });

    // Both routes into handleFile have to name the place they came from,
    // because that is the only place an error is shown. Choosing a file
    // through the picker used to pass nothing, so showDropzoneError returned
    // at its first line and a wrong file type or an unreadable document
    // produced no message, no spinner and no change on screen at all -- the
    // click simply did nothing.
    inputLeft.onchange = function () {
      if (inputLeft.files[0])
        handleFile("left", inputLeft.files[0], messageHostFor("left"));
      inputLeft.value = "";
    };
    inputRight.onchange = function () {
      if (inputRight.files[0])
        handleFile("right", inputRight.files[0], messageHostFor("right"));
      inputRight.value = "";
    };
  }

  function dropzoneFor(side) {
    return document.querySelector('.dropzone[data-side="' + side + '"]');
  }

  // An empty side shows progress and errors on its dropzone; a loaded side,
  // replaced through File > Open, shows them on its window.
  function messageHostFor(side) {
    return (
      dropzoneFor(side) ||
      document.querySelector('.window[data-win="' + side + '"]')
    );
  }

  function hideError(host) {
    host.classList.remove("error");
    var err = host.querySelector(".dz-error");
    if (err) err.style.display = "none";
  }

  function showDropzoneError(dz, message) {
    if (!dz) return;
    dz.classList.remove("drag");
    var loading = dz.querySelector(".dz-loading");
    if (loading) loading.style.display = "none";
    dz.classList.add("error");
    var err = dz.querySelector(".dz-error");
    err.style.display = "block";
    err.textContent = message;
  }

  function handleFile(side, file, dzEl) {
    // The name is checked before anything is read: see readDocument for
    // which names are known and what is then made of the contents.
    if (!KNOWN_NAME.test(file.name)) {
      showDropzoneError(
        dzEl,
        "That kind of file cannot be opened here. Open a " + ACCEPTED + " document.",
      );
      return;
    }
    if (dzEl) {
      var loading = dzEl.querySelector(".dz-loading");
      if (loading) loading.style.display = "block";
      dzEl.classList.remove("error");
      var err = dzEl.querySelector(".dz-error");
      if (err) err.style.display = "none";
    }
    readDocument(file)
      .then(function (blocks) {
        state[side] = { fileName: file.name, blocks: blocks };
        renderApp();
      })
      .catch(function (err) {
        console.error(err);
        showDropzoneError(
          dzEl,
          (err && err.userMessage) ||
            "Could not read that file. Is it a valid document?",
        );
      });
  }

  function renderDiff() {
    var ops = lcsDiff(state.left.blocks, state.right.blocks, function (x, y) {
      return x.text === y.text;
    });
    var rows = buildRows(ops);

    var removedCount = 0,
      addedCount = 0,
      fmtCount = 0;
    rows.forEach(function (r) {
      if (r.kind === "removed" || r.kind === "modified") removedCount++;
      if (r.kind === "added" || r.kind === "modified") addedCount++;
      if (r.kind === "format-changed") fmtCount++;
    });

    var gridHtml = rows
      .map(function (r, ri) {
        // A table is set a little apart from the text around it. The room
        // is given to both sides of a row at once, so that the two stay
        // level whichever of them has the table.
        var isRow = !!((r.left && r.left.isTableRow) || (r.right && r.right.isTableRow));
        var prev = rows[ri - 1];
        var wasRow = !!(
          prev &&
          ((prev.left && prev.left.isTableRow) || (prev.right && prev.right.isTableRow))
        );
        var edge = isRow && !wasRow ? "table-start" : !isRow && wasRow ? "table-end" : "";
        return (
          renderCellForSide(
            r.left,
            r.kind,
            "left",
            r.wordOps,
            r.left,
            r.right,
            r.cellOps,
            edge,
          ) +
          renderCellForSide(
            r.right,
            r.kind,
            "right",
            r.wordOps,
            r.left,
            r.right,
            r.cellOps,
            edge,
          )
        );
      })
      .join("");

    lastDiffHtml = gridHtml;

    // The two windows are drawn empty, and the diff lies over both of their
    // pages in a single scroller: one grid in one scroller is what keeps a
    // row's two sides level. The grid's column gap spans the inner edge of
    // the Original window, the desktop between the windows and the outer
    // edge of the Updated one, and is transparent, so the window frames show
    // through it.
    app.innerHTML =
      '<div class="desk diff-wrap">' +
      '<div class="frames">' +
      windowHtml("left", "") +
      windowHtml("right", "") +
      "</div>" +
      '<div class="diff-scroll"><div class="diff-grid doc-sheet">' +
      gridHtml +
      "</div></div>" +
      "</div>";

    mountOriginalScrollbar();

    document.getElementById("removed-count").textContent = removedCount;
    document.getElementById("added-count").textContent = addedCount;
    document.getElementById("fmt-count").textContent = fmtCount;
  }

  // The diff's scroller has one scrollbar, inside the Updated window's page.
  // This puts a second one inside the Original window's page: an empty track
  // with the diff's scroll range, kept in step with the diff both ways.
  var originalScrollbarObserver = null;

  function unmountOriginalScrollbar() {
    if (originalScrollbarObserver) originalScrollbarObserver.disconnect();
    originalScrollbarObserver = null;
  }

  function mountOriginalScrollbar() {
    unmountOriginalScrollbar();
    var wrap = app.querySelector(".diff-wrap");
    var scroller = app.querySelector(".diff-scroll");
    var grid = app.querySelector(".diff-grid");
    if (!wrap || !scroller || !grid || !wrap.appendChild) return;

    var bar = document.createElement("div");
    bar.className = "diff-scrollbar-left";
    bar.setAttribute("aria-hidden", "true");
    var track = document.createElement("div");
    // The scrollbar takes the bar's whole width, and a track left with no
    // width of its own does not count as overflow, so nothing would scroll.
    track.style.width = "1px";
    bar.appendChild(track);
    wrap.appendChild(bar);

    function layout() {
      // The grid leaves room for a scrollbar at the inner edge of the
      // Original page. An overlay scrollbar measures zero; it still needs
      // room to appear.
      var width = Math.max(scroller.offsetWidth - scroller.clientWidth, 12);
      wrap.style.setProperty("--sbw", width + "px");
      var firstCell = grid.firstElementChild;
      var wrapRect = wrap.getBoundingClientRect();
      var scrollerRect = scroller.getBoundingClientRect();
      bar.style.left =
        (firstCell
          ? firstCell.getBoundingClientRect().right
          : scrollerRect.left) -
        wrapRect.left +
        "px";
      bar.style.top = scrollerRect.top - wrapRect.top + "px";
      bar.style.width = width + "px";
      bar.style.height = scroller.clientHeight + "px";
      track.style.height = scroller.scrollHeight + "px";
      bar.scrollTop = scroller.scrollTop;
    }

    // Each side copies the other's position. The copy raises a scroll
    // event of its own, which is recognised and dropped: echoing it back
    // would pull a smooth scroll, still under way, back to where it was a
    // frame ago.
    var echo = [null, null];
    function follow(from, to, self, other) {
      from.addEventListener("scroll", function () {
        var expected = echo[self];
        echo[self] = null;
        if (expected !== null && Math.abs(from.scrollTop - expected) < 1)
          return;
        if (Math.abs(to.scrollTop - from.scrollTop) < 1) return;
        echo[other] = from.scrollTop;
        to.scrollTop = from.scrollTop;
      });
    }
    follow(scroller, bar, 0, 1);
    follow(bar, scroller, 1, 0);

    layout();
    if (typeof ResizeObserver === "function") {
      originalScrollbarObserver = new ResizeObserver(layout);
      originalScrollbarObserver.observe(scroller);
      originalScrollbarObserver.observe(grid);
    }
  }

  function renderApp() {
    unmountOriginalScrollbar();
    closeMenus();
    setMode();
    if (state.left && state.right) {
      renderDiff();
    } else {
      renderSetup();
    }
    updateGlobalCommands();
    renderTasks();
  }

  // ------------------------------------------------ menus and commands

  // The Start menu and the game windows belong to js/desktop.js; this file
  // looks after the two diffPad windows and the About box.
  var aboutBox = document.getElementById("about");

  function closeMenus() {
    if (app.querySelectorAll) {
      app.querySelectorAll(".menu.open").forEach(function (m) {
        m.classList.remove("open");
      });
    }
  }

  function anyMenuOpen() {
    return !!(app.querySelector && app.querySelector(".menu.open"));
  }

  function showAbout() {
    closeMenus();
    if (!aboutBox) return;
    aboutBox.hidden = false;
    var ok = aboutBox.querySelector("button");
    if (ok) ok.focus();
  }

  function hideAbout() {
    if (aboutBox) aboutBox.hidden = true;
  }

  function startOver() {
    state = { left: null, right: null };
    renderApp();
  }

  // The exported page stands alone, so the part of css/style.css that
  // draws a document is repeated here in the colours of a printed page.
  var EXPORT_CSS =
    "body{font-family:-apple-system,Helvetica,Arial,sans-serif;margin:0;padding:24px;background:#f3f2ee;color:#1a1a18;}" +
    "h2{font-size:15px;}" +
    ".diff-grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);column-gap:16px;}" +
    ".cell{padding:0 28px;background:#fff;font-family:'Times New Roman',Times,serif;font-size:11.5pt;line-height:1.35;white-space:pre-wrap;overflow-wrap:anywhere;min-width:0;border-left:1px solid #d8d6cd;border-right:1px solid #d8d6cd;}" +
    ".diff-grid>.cell:nth-child(-n+2){padding-top:24px;border-top:1px solid #d8d6cd;}.cell:nth-last-child(-n+2){padding-bottom:24px;border-bottom:1px solid #d8d6cd;}" +
    ".cell.gap{background:repeating-linear-gradient(135deg,#f3f2ee 0 4px,#fff 4px 8px);}" +
    ".cell.removed{background:#faece7;}.cell.removed .block-p,.cell.removed .block-li,.cell.removed .block-heading,.cell.removed .row-cells{color:#7a2c11;text-decoration:line-through;}" +
    ".cell.added{background:#e1f5ee;}.cell.added .block-p,.cell.added .block-li,.cell.added .block-heading,.cell.added .row-cells{color:#0b5544;}" +
    ".cell.fmt-changed{background:#fdf7e6;box-shadow:inset 3px 0 0 #e8c766;}" +
    ".del{background:#f6d5cb;color:#7a2c11;text-decoration:line-through;}.ins{background:#c4ecdc;color:#0b5544;}.ws-dot{opacity:.6;}" +
    ".block-p,.block-li,.block-heading{margin:0 0 8pt;min-height:1.35em;}" +
    ".block-heading{font-weight:700;margin-top:10pt;}.block-heading.h1{font-size:16pt;}.block-heading.h2{font-size:14pt;}.block-heading.h3{font-size:12pt;}.block-heading.styled{font-weight:400;font-size:inherit;}" +
    ".block-li{display:flex;gap:8px;align-items:baseline;margin-bottom:4pt;}.li-bullet{flex:none;min-width:14pt;}.li-text{min-width:0;flex:1;}" +
    ".block-divider{display:flex;align-items:center;gap:8px;margin:14px 0 8px;color:#8a897f;font:10px Helvetica,Arial,sans-serif;letter-spacing:.06em;}.block-divider:before,.block-divider:after{content:'';flex:1;border-top:1px dashed #b9b7ab;}" +
    ".row-cells{display:flex;border:1px solid #555;margin-top:-1px;}.cell.is-row{display:flex;flex-direction:column;}.cell.is-row>.row-cells{flex:1;}.cell.table-start{padding-top:4pt;}.cell.table-end{padding-top:8pt;}" +
    ".row-cell{flex:1 1 0;min-width:0;padding:3pt 5pt;border-left:1px solid #555;}.row-cell:first-child{border-left:none;}" +
    ".sec-badge{font:9px Helvetica,Arial,sans-serif;text-transform:uppercase;color:#8a897f;margin-right:6px;vertical-align:1px;}" +
    ".fmt-tag{display:inline-block;margin-left:8px;font:600 9px Helvetica,Arial,sans-serif;text-transform:uppercase;color:#8a6d1a;background:#fbf0d2;border-radius:3px;padding:1px 5px;white-space:nowrap;}" +
    ".inline-img img{max-width:100%;height:auto;vertical-align:middle;}" +
    ".trk-ins{text-decoration:underline wavy #1d9e75;}.trk-del{text-decoration:line-through;color:#8a4634;opacity:.85;}" +
    ".doc-link{color:#3c3489;text-decoration:underline;}.fn-ref,.fn-marker{font-size:.75em;line-height:0;}.fn-ref a,.fn-marker a{color:#3c3489;text-decoration:none;}";

  function exportDiff() {
    if (!lastDiffHtml) return;
    var doc =
      '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>diffPad export</title><style>' +
      EXPORT_CSS +
      '</style></head><body><h2 style="font-weight:500;">' +
      escapeHtml(state.left.fileName) +
      " &rarr; " +
      escapeHtml(state.right.fileName) +
      "</h2>" +
      '<div class="diff-grid doc-sheet">' +
      lastDiffHtml +
      "</div></body></html>";
    var blob = new Blob([doc], { type: "text/html" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "diffpad-export.html";
    // The anchor has to be in the document for the click to count in
    // Firefox, and the object URL has to outlive the click: revoking it
    // on the next line races the download the click has only just
    // started, and the browser may find the blob already gone.
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 1000);
  }

  function runAction(action, side) {
    closeMenus();
    if (action === "open") {
      (side === "right" ? inputRight : inputLeft).click();
    } else if (action === "close") {
      state[side] = null;
      renderApp();
    } else if (action === "export") {
      exportDiff();
    } else if (action === "reset") {
      startOver();
    } else if (action === "about") {
      showAbout();
    }
  }

  // Everything inside the windows is redrawn on each render, so their clicks
  // are handled once, here, rather than wired to each new element.
  app.addEventListener("click", function (e) {
    var t = e && e.target && e.target.closest ? e.target : null;
    if (!t) return;

    var swap = t.closest(".swap-btn");
    if (swap) {
      if (swap.disabled) return;
      var tmp = state.left;
      state.left = state.right;
      state.right = tmp;
      renderApp();
      return;
    }

    var title = t.closest(".menu-title");
    if (title) {
      var menu = title.parentNode;
      var wasOpen = menu.classList.contains("open");
      closeMenus();
      if (!wasOpen) menu.classList.add("open");
      return;
    }

    var item = t.closest(".menu-item");
    if (item) {
      if (!item.disabled)
        runAction(
          item.getAttribute("data-action"),
          item.getAttribute("data-for"),
        );
      return;
    }

    var err = t.closest(".window > .dz-error");
    if (err) hideError(err.parentNode);
  });

  // Win95 menus follow the pointer from one title to the next once one of
  // them is open.
  app.addEventListener("mouseover", function (e) {
    var t = e && e.target && e.target.closest ? e.target : null;
    var title = t && t.closest(".menu-title");
    if (!title || !anyMenuOpen()) return;
    var menu = title.parentNode;
    if (menu.classList.contains("open")) return;
    closeMenus();
    menu.classList.add("open");
  });

  // Clicking anywhere in a window makes it the active one. In the diff the
  // pages lie under the shared scroller, so the side is read from where the
  // click landed rather than from the element it landed on.
  app.addEventListener("mousedown", function (e) {
    if (!app.getBoundingClientRect || typeof e.clientX !== "number") return;
    var r = app.getBoundingClientRect();
    setActive(e.clientX < r.left + r.width / 2 ? "left" : "right");
  });

  if (document.addEventListener) {
    document.addEventListener("mousedown", function (e) {
      var t = e.target;
      if (t && t.closest && t.closest("#app .menu")) return;
      closeMenus();
    });
    document.addEventListener("keydown", function (e) {
      // While a game window is the active one, the keyboard is the game's.
      if (document.body.classList.contains("game-focus")) return;
      if (e.key === "Escape") {
        closeMenus();
        hideAbout();
      } else if (e.key === "F1") {
        e.preventDefault();
        showAbout();
      }
    });
  }

  var tasksEl = document.getElementById("tasks");
  if (tasksEl) {
    tasksEl.addEventListener("click", function (e) {
      var task = e.target && e.target.closest && e.target.closest(".task");
      if (task) setActive(task.getAttribute("data-task"));
    });
  }

  if (aboutBox) {
    aboutBox.addEventListener("click", function (e) {
      if (e.target && e.target.closest && e.target.closest("[data-close]"))
        hideAbout();
    });
  }

  if (resetBtn) {
    resetBtn.addEventListener("click", function () {
      closeMenus();
      startOver();
    });
  }

  renderApp();
})();
