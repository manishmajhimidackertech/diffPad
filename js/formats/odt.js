/*
 * OpenDocument Text: .odt and .ott (a zip holding content.xml and
 * styles.xml) and .fodt (the same document as one flat XML file).
 *
 * A reader takes the file and returns a list of plain descriptions --
 * paragraphs, table rows, notes, comments -- that js/app.js turns into the
 * blocks it compares. See "what the readers hand back" there.
 */
(function () {
  "use strict";

  var formats = (window.diffPadFormats = window.diffPadFormats || {});

  function kids(el, name) {
    var out = [];
    for (var i = 0; i < el.childNodes.length; i++) {
      var n = el.childNodes[i];
      if (n.nodeType === 1 && (!name || n.localName === name)) out.push(n);
    }
    return out;
  }
  function kid(el, name) {
    return kids(el, name)[0] || null;
  }
  function descendants(el, name) {
    var all = el.getElementsByTagName("*");
    var out = [];
    for (var i = 0; i < all.length; i++) {
      if (all[i].localName === name) out.push(all[i]);
    }
    return out;
  }

  // "0.5in", "1.27cm", "12pt" -> points. A percentage is relative to
  // something this reader does not track, and comes back undefined.
  var UNIT = { pt: 1, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, px: 0.75, pc: 12 };
  function points(v) {
    var m = /^(-?[0-9.]+)(pt|in|cm|mm|px|pc)$/.exec(String(v || "").trim());
    if (!m) return undefined;
    var n = parseFloat(m[1]);
    return isFinite(n) ? n * UNIT[m[2]] : undefined;
  }

  function colour(v) {
    var m = /^#([0-9a-fA-F]{6})$/.exec(String(v || "").trim());
    return m ? m[1].toUpperCase() : null;
  }

  // What a style:text-properties element says, and only what it says.
  function textProps(el, fonts) {
    var p = {};
    if (!el) return p;
    var w = el.getAttribute("fo:font-weight");
    if (w !== null) p.bold = w === "bold" || parseInt(w, 10) >= 600;
    var s = el.getAttribute("fo:font-style");
    if (s !== null) p.italic = s === "italic" || s === "oblique";
    var u = el.getAttribute("style:text-underline-style");
    if (u !== null) p.underline = u !== "none";
    var lt = el.getAttribute("style:text-line-through-style");
    if (lt === null) lt = el.getAttribute("style:text-line-through-type");
    if (lt !== null) p.strike = lt !== "none";
    var c = el.getAttribute("fo:color");
    if (c !== null) p.color = colour(c);
    var size = points(el.getAttribute("fo:font-size"));
    if (size !== undefined && size > 0) p.size = Math.round(size * 2);
    var font = el.getAttribute("style:font-name");
    if (font !== null) p.font = (fonts && fonts[font]) || font;
    else if (el.getAttribute("fo:font-family") !== null)
      p.font = el.getAttribute("fo:font-family").replace(/['"]/g, "");
    var bg = el.getAttribute("fo:background-color");
    if (bg !== null) p.highlight = colour(bg);
    var pos = el.getAttribute("style:text-position");
    if (pos !== null) {
      p.vert = /^(super|[1-9])/.test(pos) ? "super" : /^(sub|-)/.test(pos) ? "sub" : null;
    }
    var tt = el.getAttribute("fo:text-transform");
    if (tt !== null) p.caps = tt === "uppercase";
    return p;
  }

  var ALIGN = { start: "left", left: "left", end: "right", right: "right", center: "center", justify: "justify" };

  function paraProps(el) {
    var p = {};
    if (!el) return p;
    var a = el.getAttribute("fo:text-align");
    if (a !== null && ALIGN[a]) p.align = ALIGN[a];
    [
      ["fo:margin-left", "indL"],
      ["fo:margin-right", "indR"],
      ["fo:text-indent", "indFirst"],
      ["fo:margin-top", "before"],
      ["fo:margin-bottom", "after"],
    ].forEach(function (pair) {
      var v = points(el.getAttribute(pair[0]));
      if (v !== undefined) p[pair[1]] = v;
    });
    return p;
  }

  var NUM_FORMAT = { 1: "decimal", a: "lowerLetter", A: "upperLetter", i: "lowerRoman", I: "upperRoman" };

  // Every style in the document, from every part that declares any.
  function readStyles(roots) {
    var styles = {
      byKey: Object.create(null),
      defaults: Object.create(null),
      lists: Object.create(null),
      resolved: Object.create(null),
    };
    var fonts = Object.create(null);
    roots.forEach(function (root) {
      descendants(root, "font-face").forEach(function (f) {
        var name = f.getAttribute("style:name");
        var family = f.getAttribute("svg:font-family");
        if (name && family) fonts[name] = family.replace(/['"]/g, "");
      });
    });
    function entry(st) {
      var outline = parseInt(st.getAttribute("style:default-outline-level") || "", 10);
      var cell = kid(st, "table-cell-properties");
      var col = kid(st, "table-column-properties");
      return {
        parent: st.getAttribute("style:parent-style-name"),
        display: st.getAttribute("style:display-name") || st.getAttribute("style:name") || "",
        outline: outline > 0 ? outline : 0,
        text: textProps(kid(st, "text-properties"), fonts),
        para: paraProps(kid(st, "paragraph-properties")),
        fill: cell ? colour(cell.getAttribute("fo:background-color")) : null,
        width: col ? points(col.getAttribute("style:column-width")) : undefined,
      };
    }
    roots.forEach(function (root) {
      ["styles", "automatic-styles"].forEach(function (holder) {
        kids(root, holder).forEach(function (h) {
          kids(h, "style").forEach(function (st) {
            var key = st.getAttribute("style:family") + ":" + st.getAttribute("style:name");
            // A document's own automatic styles come first and win.
            if (!styles.byKey[key]) styles.byKey[key] = entry(st);
          });
          kids(h, "default-style").forEach(function (st) {
            var family = st.getAttribute("style:family");
            if (!styles.defaults[family]) styles.defaults[family] = entry(st);
          });
          kids(h, "list-style").forEach(function (ls) {
            var name = ls.getAttribute("style:name");
            if (styles.lists[name]) return;
            var levels = [];
            kids(ls).forEach(function (lv) {
              var i = (parseInt(lv.getAttribute("text:level") || "1", 10) || 1) - 1;
              if (lv.localName === "list-level-style-bullet") {
                levels[i] = { fmt: "bullet", text: lv.getAttribute("text:bullet-char") || "" };
              } else if (lv.localName === "list-level-style-number") {
                var nf = lv.getAttribute("style:num-format");
                var shown = parseInt(lv.getAttribute("text:display-levels") || "1", 10) || 1;
                var body = "";
                for (var d = Math.max(0, i - shown + 1); d <= i; d++)
                  body += (body ? "." : "") + "%" + (d + 1);
                var start = parseInt(lv.getAttribute("text:start-value") || "", 10);
                levels[i] = {
                  fmt: nf ? NUM_FORMAT[nf] || "decimal" : "none",
                  text: nf
                    ? (lv.getAttribute("style:num-prefix") || "") + body + (lv.getAttribute("style:num-suffix") || "")
                    : "",
                  start: isFinite(start) ? start : undefined,
                };
              }
            });
            styles.lists[name] = levels;
          });
        });
      });
    });
    return styles;
  }

  // A style with its parents and the family's defaults laid underneath it.
  function resolve(styles, family, name) {
    var key = family + ":" + name;
    if (styles.resolved[key]) return styles.resolved[key];
    var chain = [];
    var seen = Object.create(null);
    for (var at = name; at && styles.byKey[family + ":" + at] && !seen[at]; ) {
      seen[at] = true;
      chain.unshift(styles.byKey[family + ":" + at]);
      at = chain[0].parent;
    }
    var dflt = styles.defaults[family];
    var out = { text: {}, para: {}, heading: 0, fill: null, width: undefined };
    if (dflt) {
      Object.assign(out.text, dflt.text);
      Object.assign(out.para, dflt.para);
    }
    chain.forEach(function (st) {
      Object.assign(out.text, st.text);
      Object.assign(out.para, st.para);
      if (st.fill) out.fill = st.fill;
      if (st.width !== undefined) out.width = st.width;
      var m = /^Heading[ _](?:20_)?([1-9])$/.exec(st.display);
      if (st.outline) out.heading = st.outline;
      else if (m) out.heading = parseInt(m[1], 10);
    });
    styles.resolved[key] = out;
    return out;
  }

  function read(roots, images, tools) {
    var styles = readStyles(roots);
    var nextLabel = tools.listCounter();
    var footnotes = [],
      endnotes = [],
      comments = [];
    var listSerial = 0;
    var lastListId = Object.create(null);
    // An insertion being tracked runs from a change-start to its
    // change-end, which may be paragraphs apart.
    var inserting = null;

    // Tracked changes are kept apart from the text: the text holds a mark,
    // and the region the mark names says what was done and by whom.
    var regions = Object.create(null);
    roots.forEach(function (root) {
      descendants(root, "changed-region").forEach(function (region) {
        var id = region.getAttribute("text:id") || region.getAttribute("xml:id");
        var what = kids(region)[0];
        if (!id || !what) return;
        var creator = descendants(what, "creator")[0];
        regions[id] = {
          type: what.localName,
          author: creator ? creator.textContent : "",
          el: what,
        };
      });
    });

    function segment(text, props, ctx) {
      return {
        text: text,
        bold: props.bold,
        italic: props.italic,
        underline: props.underline,
        strike: props.strike,
        color: props.color,
        size: props.size,
        font: props.font,
        highlight: props.highlight,
        vert: props.vert,
        caps: props.caps,
        href: ctx.href || null,
        isDel: !!ctx.deleted,
        trackType: ctx.deleted ? "del" : inserting ? "ins" : null,
        author: ctx.deleted ? ctx.author : inserting ? inserting.author : null,
      };
    }

    // Everything a paragraph holds, flattened into runs.
    function inline(node, props, ctx, out) {
      for (var i = 0; i < node.childNodes.length; i++) {
        var n = node.childNodes[i];
        if (n.nodeType === 3 || n.nodeType === 4) {
          // White space in the file is only white space: runs of it are
          // one space, and real spacing is written as text:s and text:tab.
          var text = n.textContent.replace(/[\t\r\n ]+/g, " ");
          if (text) out.push(segment(text, props, ctx));
          continue;
        }
        if (n.nodeType !== 1) continue;
        var ln = n.localName;
        if (ln === "span") {
          var st = resolve(styles, "text", n.getAttribute("text:style-name"));
          inline(n, Object.assign({}, props, st.text), ctx, out);
        } else if (ln === "s") {
          var count = Math.min(1000, parseInt(n.getAttribute("text:c") || "1", 10) || 1);
          out.push(segment(new Array(count + 1).join(" "), props, ctx));
        } else if (ln === "tab") {
          out.push(segment("\t", props, ctx));
        } else if (ln === "line-break") {
          out.push(segment("\n", props, ctx));
        } else if (ln === "a") {
          inline(n, props, Object.assign({}, ctx, { href: n.getAttribute("xlink:href") }), out);
        } else if (ln === "note") {
          var isEnd = n.getAttribute("text:note-class") === "endnote";
          var list = isEnd ? endnotes : footnotes;
          var id = String(list.length + 1);
          var body = kid(n, "note-body");
          var segs = [];
          var base = null;
          if (body) {
            paragraphsBelow(body).forEach(function (p, idx) {
              var ps = resolve(styles, "paragraph", p.getAttribute("text:style-name"));
              if (idx === 0) base = baseOf(ps.text);
              else segs.push({ text: " " });
              inline(p, ps.text, {}, segs);
            });
          }
          list.push({ type: "note", kind: isEnd ? "en" : "fn", id: id, segments: segs, base: base || {} });
          out.push({ note: { kind: isEnd ? "en" : "fn", id: id } });
        } else if (ln === "annotation") {
          var who = kid(n, "creator");
          var csegs = [];
          kids(n, "p").forEach(function (p, idx) {
            if (idx > 0) csegs.push({ text: " " });
            inline(p, resolve(styles, "paragraph", p.getAttribute("text:style-name")).text, {}, csegs);
          });
          comments.push({ type: "comment", author: who ? who.textContent : "Unknown", segments: csegs });
        } else if (ln === "frame") {
          frame(n, props, ctx, out);
        } else if (ln === "change") {
          var gone = regions[n.getAttribute("text:change-id")];
          if (gone && gone.type === "deletion") {
            paragraphsBelow(gone.el).forEach(function (p, idx) {
              var dctx = { deleted: true, author: gone.author };
              if (idx > 0) out.push(segment(" ", props, dctx));
              inline(p, props, dctx, out);
            });
          }
        } else if (ln === "change-start") {
          var added = regions[n.getAttribute("text:change-id")];
          if (added && added.type === "insertion") inserting = { author: added.author };
        } else if (ln === "change-end") {
          inserting = null;
        } else if (
          ln === "annotation-end" ||
          ln === "tracked-changes" ||
          ln === "soft-page-break" ||
          ln === "bookmark" ||
          ln === "bookmark-start" ||
          ln === "bookmark-end" ||
          ln === "reference-mark" ||
          ln === "reference-mark-start" ||
          ln === "reference-mark-end"
        ) {
          // Marks, not text.
        } else {
          // Fields, ruby, and whatever else: the text inside is the text.
          inline(n, props, ctx, out);
        }
      }
    }

    // A frame holds a picture or a text box.
    function frame(el, props, ctx, out) {
      var img = kid(el, "image");
      if (img) {
        var width = points(el.getAttribute("svg:width"));
        var found = images(img);
        out.push({
          image: {
            dataUrl: found ? found.dataUrl : null,
            fingerprint: found ? found.fingerprint : img.getAttribute("xlink:href"),
            widthPx: width ? (width * 96) / 72 : null,
          },
          trackType: inserting ? "ins" : null,
          author: inserting ? inserting.author : null,
        });
      }
      kids(el, "text-box").forEach(function (box) {
        paragraphsBelow(box).forEach(function (p, idx) {
          if (idx > 0) out.push({ text: " " });
          inline(p, resolve(styles, "paragraph", p.getAttribute("text:style-name")).text, ctx, out);
        });
      });
    }

    // Every paragraph at or below `el`, without going into a paragraph.
    function paragraphsBelow(el) {
      var out = [];
      (function walk(node) {
        kids(node).forEach(function (c) {
          if (c.localName === "p" || c.localName === "h") out.push(c);
          else if (c.localName !== "tracked-changes") walk(c);
        });
      })(el);
      return out;
    }

    function baseOf(text) {
      return {
        bold: !!text.bold,
        italic: !!text.italic,
        color: text.color || null,
        size: text.size || null,
        font: text.font || null,
      };
    }

    function paragraph(el, section, listCtx) {
      var st = resolve(styles, "paragraph", el.getAttribute("text:style-name"));
      var segs = [];
      inline(el, st.text, {}, segs);
      var heading = 0;
      if (el.localName === "h")
        heading = parseInt(el.getAttribute("text:outline-level") || "1", 10) || 1;
      else if (st.heading) heading = st.heading;
      var list = null;
      if (listCtx && listCtx.pending) {
        listCtx.pending = false;
        var label = nextLabel(listCtx.id, listCtx.level, styles.lists[listCtx.style]);
        if (label.label || label.kind === "bullet")
          list = { level: listCtx.level, label: label.label, kind: label.kind };
      }
      return {
        type: "p",
        section: section,
        heading: heading,
        list: list,
        para: st.para,
        base: baseOf(st.text),
        segments: segs,
      };
    }

    function table(el, section, out) {
      var widths = [];
      (function columns(node) {
        kids(node).forEach(function (c) {
          if (c.localName === "table-column") {
            var w = resolve(styles, "table-column", c.getAttribute("table:style-name")).width;
            var n = Math.min(64, parseInt(c.getAttribute("table:number-columns-repeated") || "1", 10) || 1);
            for (var i = 0; i < n; i++) widths.push(w || 0);
          } else if (c.localName === "table-columns" || c.localName === "table-header-columns") {
            columns(c);
          }
        });
      })(el);
      var rows = [];
      (function gather(node) {
        kids(node).forEach(function (c) {
          if (c.localName === "table-row") rows.push(c);
          else if (c.localName === "table-header-rows" || c.localName === "table-rows" || c.localName === "table-row-group") gather(c);
        });
      })(el);
      rows.forEach(function (tr, ri) {
        var col = 0;
        var cells = [];
        kids(tr).forEach(function (tc) {
          if (tc.localName === "covered-table-cell") {
            col++;
            return;
          }
          if (tc.localName !== "table-cell") return;
          var span = Math.max(1, parseInt(tc.getAttribute("table:number-columns-spanned") || "1", 10) || 1);
          var width = 0;
          for (var g = col; g < col + span && g < widths.length; g++) width += widths[g];
          col += span;
          var segs = [];
          var base = null,
            align = null;
          paragraphsBelow(tc).forEach(function (p, idx) {
            var ps = resolve(styles, "paragraph", p.getAttribute("text:style-name"));
            if (idx === 0) {
              base = baseOf(ps.text);
              align = ps.para.align || null;
            } else {
              segs.push({ text: "\n" });
            }
            inline(p, ps.text, {}, segs);
          });
          cells.push({
            segments: segs,
            fill: resolve(styles, "table-cell", tc.getAttribute("table:style-name")).fill,
            // Twentieths of a point, as the other readers give widths.
            width: width ? width * 20 : null,
            align: align,
            base: base || {},
          });
        });
        out.push({ type: "row", section: section, cells: cells, first: ri === 0, last: ri === rows.length - 1 });
      });
    }

    function blocks(parent, section, listCtx, out) {
      kids(parent).forEach(function (el) {
        var ln = el.localName;
        if (ln === "p" || ln === "h") {
          out.push(paragraph(el, section, listCtx));
        } else if (ln === "list") {
          var styleName = el.getAttribute("text:style-name") || (listCtx && listCtx.style) || "";
          var level = listCtx ? listCtx.level + 1 : 0;
          var id;
          if (listCtx) {
            id = listCtx.id;
          } else if (el.getAttribute("text:continue-numbering") === "true" && lastListId[styleName]) {
            id = lastListId[styleName];
          } else {
            id = "L" + ++listSerial;
            lastListId[styleName] = id;
          }
          kids(el).forEach(function (item) {
            if (item.localName !== "list-item" && item.localName !== "list-header") return;
            blocks(
              item,
              section,
              { style: styleName, level: level, id: id, pending: item.localName === "list-item" },
              out,
            );
          });
        } else if (ln === "table") {
          table(el, section, out);
        } else if (ln === "frame") {
          var segs = [];
          frame(el, {}, {}, segs);
          if (segs.length) out.push({ type: "p", section: section, segments: segs });
        } else if (
          ln === "tracked-changes" ||
          ln === "sequence-decls" ||
          ln === "variable-decls" ||
          ln === "user-field-decls" ||
          ln === "forms"
        ) {
          // Declarations, not content.
        } else {
          // Sections, a table of contents and its body, an index.
          blocks(el, section, listCtx, out);
        }
      });
    }

    var out = [];

    // Headers and footers live with the page styles.
    var headers = [],
      footers = [];
    roots.forEach(function (root) {
      kids(root, "master-styles").forEach(function (ms) {
        kids(ms, "master-page").forEach(function (page) {
          kids(page).forEach(function (part) {
            var into = /^header/.test(part.localName) ? headers : /^footer/.test(part.localName) ? footers : null;
            if (!into) return;
            var got = [];
            blocks(part, into === headers ? "header" : "footer", null, got);
            if (got.length) into.push(got);
          });
        });
      });
    });
    headers.forEach(function (h) {
      out.push({ type: "divider", label: "HEADER" });
      out = out.concat(h);
    });

    roots.forEach(function (root) {
      kids(root, "body").forEach(function (body) {
        kids(body, "text").forEach(function (text) {
          blocks(text, "body", null, out);
        });
      });
    });

    footers.forEach(function (f) {
      out.push({ type: "divider", label: "FOOTER" });
      out = out.concat(f);
    });
    if (footnotes.length) out.push({ type: "divider", label: "FOOTNOTES" });
    out = out.concat(footnotes);
    if (endnotes.length) out.push({ type: "divider", label: "ENDNOTES" });
    out = out.concat(endnotes);
    if (comments.length) out.push({ type: "divider", label: "COMMENTS" });
    return out.concat(comments);
  }

  function parse(xml, what) {
    var doc = new DOMParser().parseFromString(xml, "application/xml");
    var root = doc && doc.documentElement;
    if (!root || root.localName === "parsererror" || descendants(root, "parsererror").length)
      throw new Error(what + " is not well-formed XML");
    return root;
  }

  formats.odt = async function (zip, tools) {
    var contentFile = zip.file("content.xml");
    if (!contentFile) throw new Error("Missing content.xml — not a valid .odt");
    var roots = [parse(await contentFile.async("string"), "content.xml")];
    var stylesFile = zip.file("styles.xml");
    if (stylesFile) roots.push(parse(await stylesFile.async("string"), "styles.xml"));

    // Pictures are files in the package. Read them all first, so the walk
    // through the document does not have to wait for any.
    var loaded = Object.create(null);
    var hrefs = Object.create(null);
    roots.forEach(function (root) {
      descendants(root, "image").forEach(function (img) {
        var href = img.getAttribute("xlink:href");
        if (href) hrefs[href.replace(/^\.\//, "")] = true;
      });
    });
    await Promise.all(
      Object.keys(hrefs).map(async function (href) {
        var f = zip.file(href);
        if (!f) return;
        var bytes = await f.async("uint8array");
        var mime = tools.imageMime(bytes);
        if (mime) loaded[href] = tools.imageFromBytes(bytes, mime);
      }),
    );
    return read(
      roots,
      function (img) {
        var href = img.getAttribute("xlink:href");
        return href ? loaded[href.replace(/^\.\//, "")] || null : null;
      },
      tools,
    );
  };

  formats.fodt = function (xml, tools) {
    var root = parse(xml, "the document");
    if (root.localName !== "document") throw new Error("not a flat OpenDocument file");
    return read(
      [root],
      function (img) {
        // A flat file carries each picture inside it, in base64.
        var data = kid(img, "binary-data");
        if (!data) return null;
        var raw;
        try {
          raw = atob(data.textContent.replace(/\s+/g, ""));
        } catch (e) {
          return null;
        }
        var bytes = new Uint8Array(raw.length);
        for (var i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
        var mime = tools.imageMime(bytes);
        return mime ? tools.imageFromBytes(bytes, mime) : null;
      },
      tools,
    );
  };
})();
