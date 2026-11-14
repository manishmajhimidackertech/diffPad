/*
 * Plain text and HTML.
 *
 * Plain text -- .txt, and Markdown read as the text it is -- is one
 * paragraph to a line. HTML is read for its structure: paragraphs, headings,
 * lists, tables, and the handful of inline elements that mean bold, italic,
 * a link and the like. The page is parsed, never displayed and never run:
 * its scripts, styles and remote images are not loaded, only read past.
 *
 * Both return the same descriptions the other readers do. See "what the
 * readers hand back" in js/app.js.
 */
(function () {
  "use strict";

  var formats = (window.diffPadFormats = window.diffPadFormats || {});

  formats.text = function (text) {
    var lines = String(text).replace(/\r\n?/g, "\n").split("\n");
    // A file ends with a line end, not with an empty last paragraph.
    if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
    return lines.map(function (line) {
      return {
        type: "p",
        // Form feeds and the like are not text.
        segments: line ? [{ text: line.replace(/[\u0000-\u0008\u000b-\u001f]/g, "") }] : [],
        para: { after: 0 },
      };
    });
  };

  var SKIP = {
    script: 1, style: 1, head: 1, title: 1, template: 1, noscript: 1,
    iframe: 1, object: 1, embed: 1, svg: 1, math: 1, canvas: 1, audio: 1,
    video: 1, select: 1, textarea: 1,
  };
  var BLOCK = {
    p: 1, div: 1, section: 1, article: 1, main: 1, aside: 1, header: 1,
    footer: 1, nav: 1, blockquote: 1, pre: 1, address: 1, figure: 1,
    figcaption: 1, form: 1, fieldset: 1, dl: 1, dt: 1, dd: 1, center: 1,
    details: 1, summary: 1, hr: 1, body: 1, html: 1,
  };
  var ALIGN = { left: "left", center: "center", right: "right", justify: "justify" };

  formats.html = function (html, tools) {
    var doc = new DOMParser().parseFromString(String(html), "text/html");
    var root = (doc && (doc.body || doc.documentElement)) || null;
    if (!root) throw new Error("the page has no body");

    var out = [];
    var segs = [];
    var para = {};
    var lists = []; // the lists the walk is inside: { ordered, n }
    var item = null; // the list item whose first paragraph is still to come
    var heading = 0;
    var pre = 0;

    function tag(el) {
      return String(el.localName || el.nodeName || "").toLowerCase();
    }

    function blank() {
      return !segs.some(function (s) {
        return s.image || /\S/.test(s.text || "");
      });
    }

    // End the paragraph being built, if there is anything in it.
    function flush() {
      if (blank()) {
        segs = [];
        return;
      }
      // What ends up at the edges of a paragraph is the page's layout.
      if (!pre) {
        var first = segs[0],
          last = segs[segs.length - 1];
        if (first.text !== undefined) first.text = first.text.replace(/^\s+/, "");
        if (last.text !== undefined) last.text = last.text.replace(/\s+$/, "");
      }
      var list = null;
      if (item) {
        var level = lists.length - 1;
        list = item.ordered
          ? { level: level, kind: "number", label: item.n + "." }
          : { level: level, kind: "bullet", label: tools.bulletGlyph("", level) };
        item = null;
      }
      out.push({ type: "p", heading: heading, list: list, para: para, segments: segs });
      segs = [];
    }

    function alignOf(el) {
      var a = String(el.getAttribute("align") || "").toLowerCase();
      if (ALIGN[a]) return ALIGN[a];
      var m = /text-align\s*:\s*(left|center|right|justify)/i.exec(el.getAttribute("style") || "");
      return m ? m[1].toLowerCase() : null;
    }

    function addText(text, props) {
      if (!pre) {
        text = text.replace(/[\t\r\n\f ]+/g, " ");
        // A space after a space, or at the start of a paragraph, is
        // nothing.
        var last = segs[segs.length - 1];
        if (text.charAt(0) === " " && (!last || (last.text !== undefined && /[ \n]$/.test(last.text))))
          text = text.slice(1);
      }
      if (!text) return;
      segs.push(Object.assign({ text: text }, props));
    }

    function imageOf(el) {
      var src = el.getAttribute("src") || "";
      var m = /^data:(image\/(?:png|jpeg|gif|bmp));base64,([A-Za-z0-9+/=\s]+)$/.exec(src);
      var width = parseInt(el.getAttribute("width") || "", 10);
      if (m) {
        var b64 = m[2].replace(/\s+/g, "");
        return {
          dataUrl: "data:" + m[1] + ";base64," + b64,
          fingerprint: b64.length + ":" + b64.slice(-16),
          widthPx: width > 0 ? width : null,
        };
      }
      // A picture kept somewhere else is known only by where.
      return { dataUrl: null, fingerprint: "src:" + src.slice(-120) };
    }

    function cellSegments(td) {
      // A cell is read as a small document of its own, and its paragraphs
      // joined.
      var keep = { out: out, segs: segs, para: para, lists: lists, item: item, heading: heading };
      out = [];
      segs = [];
      para = {};
      lists = [];
      item = null;
      heading = 0;
      walk(td, { bold: tag(td) === "th" });
      flush();
      var got = [];
      out.forEach(function (b, i) {
        var own = b.type === "row" ? [].concat.apply([], b.cells.map(function (c) { return c.segments.concat([{ text: " " }]); })) : b.segments;
        if (i > 0) got.push({ text: "\n" });
        got = got.concat(own);
      });
      out = keep.out;
      segs = keep.segs;
      para = keep.para;
      lists = keep.lists;
      item = keep.item;
      heading = keep.heading;
      return got;
    }

    function table(el) {
      flush();
      var rows = [];
      (function gather(node) {
        for (var i = 0; i < node.childNodes.length; i++) {
          var c = node.childNodes[i];
          if (c.nodeType !== 1) continue;
          var t = tag(c);
          if (t === "tr") rows.push(c);
          else if (t === "thead" || t === "tbody" || t === "tfoot") gather(c);
        }
      })(el);
      rows.forEach(function (tr, ri) {
        var cells = [];
        for (var i = 0; i < tr.childNodes.length; i++) {
          var td = tr.childNodes[i];
          if (td.nodeType !== 1 || (tag(td) !== "td" && tag(td) !== "th")) continue;
          var span = Math.max(1, Math.min(64, parseInt(td.getAttribute("colspan") || "1", 10) || 1));
          cells.push({ segments: cellSegments(td), width: span, align: alignOf(td) });
        }
        if (cells.length) out.push({ type: "row", cells: cells, first: ri === 0, last: ri === rows.length - 1 });
      });
    }

    function walk(node, props) {
      for (var i = 0; i < node.childNodes.length; i++) {
        var n = node.childNodes[i];
        if (n.nodeType === 3) {
          addText(n.textContent, props);
          continue;
        }
        if (n.nodeType !== 1) continue;
        var t = tag(n);
        if (SKIP[t]) continue;
        if (t === "br") {
          segs.push(Object.assign({ text: "\n" }, props));
        } else if (t === "img") {
          segs.push({ image: imageOf(n) });
        } else if (t === "table") {
          table(n);
        } else if (t === "ul" || t === "ol") {
          flush();
          var start = parseInt(n.getAttribute("start") || "1", 10);
          lists.push({ ordered: t === "ol", n: (isFinite(start) ? start : 1) - 1 });
          walk(n, props);
          flush();
          lists.pop();
        } else if (t === "li") {
          flush();
          var list = lists[lists.length - 1];
          if (list) {
            list.n++;
            item = { ordered: list.ordered, n: list.n };
          }
          walk(n, props);
          flush();
          item = null;
        } else if (/^h[1-6]$/.test(t)) {
          flush();
          heading = parseInt(t.charAt(1), 10);
          para = { align: alignOf(n) };
          walk(n, props);
          flush();
          heading = 0;
          para = {};
        } else if (BLOCK[t]) {
          flush();
          var keepPara = para;
          var a = alignOf(n) || (t === "center" ? "center" : null);
          para = a ? { align: a } : t === "blockquote" ? { indL: 30 } : keepPara;
          if (t === "pre") pre++;
          walk(n, t === "pre" ? Object.assign({}, props, { font: "Courier New" }) : props);
          flush();
          if (t === "pre") pre--;
          para = keepPara;
        } else {
          var next = props;
          if (t === "b" || t === "strong") next = Object.assign({}, props, { bold: true });
          else if (t === "i" || t === "em" || t === "cite" || t === "var") next = Object.assign({}, props, { italic: true });
          else if (t === "u") next = Object.assign({}, props, { underline: true });
          else if (t === "s" || t === "strike") next = Object.assign({}, props, { strike: true });
          else if (t === "del") next = Object.assign({}, props, { isDel: true, trackType: "del" });
          else if (t === "ins") next = Object.assign({}, props, { trackType: "ins" });
          else if (t === "sup") next = Object.assign({}, props, { vert: "super" });
          else if (t === "sub") next = Object.assign({}, props, { vert: "sub" });
          else if (t === "code" || t === "kbd" || t === "samp" || t === "tt") next = Object.assign({}, props, { font: "Courier New" });
          else if (t === "mark") next = Object.assign({}, props, { highlight: "FFFF00" });
          else if (t === "a" && n.getAttribute("href")) next = Object.assign({}, props, { href: n.getAttribute("href") });
          walk(n, next);
        }
      }
    }

    walk(root, {});
    flush();
    return out;
  };
})();
