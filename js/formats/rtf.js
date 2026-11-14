/*
 * Rich Text Format.
 *
 * RTF is a stream of groups in braces, control words that change the state
 * of the group they are in, and text. This walks the stream once, keeping a
 * stack of states, and builds the same list of paragraphs, table rows, notes
 * and comments the other readers do. See "what the readers hand back" in
 * js/app.js.
 */
(function () {
  "use strict";

  var formats = (window.diffPadFormats = window.diffPadFormats || {});

  // Groups whose contents are not the document's text.
  var SKIP = {
    info: 1, filetbl: 1, listtable: 1, listoverridetable: 1, rsidtbl: 1,
    generator: 1, themedata: 1, colorschememapping: 1, latentstyles: 1,
    datastore: 1, xmlnstbl: 1, pgdsctbl: 1, nonshppict: 1, shprslt: 1, sp: 1,
    bkmkstart: 1, bkmkend: 1, xe: 1, tc: 1, ftnsep: 1, ftnsepc: 1, ftncn: 1,
    aftnsep: 1, aftnsepc: 1, aftncn: 1, objdata: 1, objclass: 1, objname: 1,
    mmathPr: 1, fontemb: 1, panose: 1, falt: 1, template: 1, private: 1,
    userprops: 1, docvar: 1, pnseclvl: 1, atnid: 1, atndate: 1, atnref: 1,
    atnparent: 1, formfield: 1, wgrffmtfilter: 1, upr: 0,
  };

  // Control words that stand for a character.
  var CHARS = {
    tab: "\t", line: "\n", emdash: "—", endash: "–", bullet: "•",
    lquote: "‘", rquote: "’", ldblquote: "“", rdblquote: "”",
    emspace: " ", enspace: " ", qmspace: " ", zwj: "‍",
    zwnj: "‌",
  };

  // A font's character set, as its code page.
  var CHARSET_CP = {
    0: 1252, 77: 10000, 128: 932, 129: 949, 130: 1361, 134: 936, 136: 950,
    161: 1253, 162: 1254, 163: 1258, 177: 1255, 178: 1256, 186: 1257,
    204: 1251, 222: 874, 238: 1250,
  };
  var CP_LABEL = {
    932: "shift_jis", 936: "gbk", 949: "euc-kr", 950: "big5",
    10000: "macintosh", 874: "windows-874", 65001: "utf-8",
  };

  var decoders = Object.create(null);
  function decodeBytes(bytes, cp) {
    var ascii = true;
    for (var i = 0; i < bytes.length; i++) {
      if (bytes[i] > 0x7f) {
        ascii = false;
        break;
      }
    }
    if (ascii || cp === "symbol") return String.fromCharCode.apply(null, bytes);
    var label = CP_LABEL[cp] || (cp >= 1250 && cp <= 1258 ? "windows-" + cp : "windows-1252");
    if (!decoders[label]) {
      try {
        decoders[label] = new TextDecoder(label);
      } catch (e) {
        decoders[label] = new TextDecoder("windows-1252");
      }
    }
    return decoders[label].decode(new Uint8Array(bytes));
  }

  function hexValue(c) {
    if (c >= 48 && c <= 57) return c - 48;
    if (c >= 97 && c <= 102) return c - 87;
    if (c >= 65 && c <= 70) return c - 55;
    return -1;
  }

  function newStory(kind) {
    return {
      kind: kind,
      out: [],
      segs: [],
      cells: [],
      cellSegs: [],
      cellOpen: false,
      row: { cellx: [], fills: [], left: 0 },
      label: null,
      author: null,
      note: null,
    };
  }

  formats.rtf = function (bytes, tools) {
    var fonts = Object.create(null); // number -> { name, cp }
    var colors = [];
    var styles = Object.create(null); // "s1" -> { name, heading }
    var authors = [];
    var ansicp = 1252;
    var defaultFont = null;

    var body = newStory("body");
    var headers = [],
      footers = [],
      footnotes = [],
      endnotes = [],
      comments = [];
    var lastAuthor = "";

    function defaultChr() {
      return { f: defaultFont };
    }

    var state = {
      dest: "body",
      story: body,
      chr: defaultChr(),
      par: {},
      uc: 1,
      field: null,
      star: false,
      first: true, // nothing has been seen in this group yet
    };
    var stack = [];

    // Text waiting to be decoded: bytes, in the code page of the font they
    // were written in.
    var pending = [];
    var pendingCp = null;
    var skipChars = 0; // what is left to skip after a \u

    // ---- building paragraphs

    function segment(text) {
      var c = state.chr;
      var font = c.f !== null && c.f !== undefined ? fonts[c.f] : null;
      var color = c.cf ? colors[c.cf] : null;
      var hl = c.highlight ? colors[c.highlight] : null;
      return {
        text: text,
        bold: !!c.b,
        italic: !!c.i,
        underline: !!c.ul,
        strike: !!c.strike,
        caps: !!c.caps,
        color: color || null,
        highlight: hl || null,
        size: c.fs || 24,
        font: font ? font.name : null,
        vert: c.vert || null,
        href: state.field ? state.field.href : null,
        isDel: !!c.deleted,
        trackType: c.deleted ? "del" : c.revised ? "ins" : null,
        author: c.deleted
          ? authors[c.revauthdel] || authors[c.revauth] || ""
          : c.revised
            ? authors[c.revauth] || ""
            : null,
      };
    }

    function emit(text) {
      if (!text) return;
      var d = state.dest;
      if (d === "body") {
        if (state.chr.v) return; // hidden text
        var segs = state.story.segs;
        var seg = segment(text);
        var last = segs[segs.length - 1];
        // Runs that differ in nothing are one run.
        if (last && last.text !== undefined && sameRun(last, seg)) last.text += text;
        else segs.push(seg);
      } else if (d === "name") {
        state.holder.text += text;
      } else if (d === "fldinst") {
        if (state.field) state.field.inst += text;
      } else if (d === "listtext") {
        state.story.label = (state.story.label || "") + text;
      } else if (d === "atnauthor") {
        lastAuthor += text;
      }
    }

    function sameRun(a, b) {
      for (var k in b) {
        if (k !== "text" && a[k] !== b[k]) return false;
      }
      return true;
    }

    function flush() {
      if (!pending.length) return;
      var text = decodeBytes(pending, pendingCp);
      pending = [];
      emit(text);
    }

    function pushByte(b) {
      if (skipChars > 0) {
        skipChars--;
        return;
      }
      var font = state.chr.f !== null && state.chr.f !== undefined ? fonts[state.chr.f] : null;
      var cp = font && font.cp ? font.cp : ansicp;
      if (pending.length && cp !== pendingCp) flush();
      pendingCp = cp;
      pending.push(b);
    }

    function paraProps() {
      var p = state.par;
      var out = {};
      if (p.align) out.align = p.align;
      if (p.li) out.indL = p.li / 20;
      if (p.ri) out.indR = p.ri / 20;
      if (p.fi) out.indFirst = p.fi / 20;
      out.before = (p.sb || 0) / 20;
      out.after = (p.sa || 0) / 20;
      return out;
    }

    function listOf(story) {
      if (story.label === null) return null;
      var raw = story.label.replace(/[\t ]+$/, "").trim();
      story.label = null;
      if (!raw) return null;
      var numbered = /[0-9A-Za-z]/.test(raw) && raw !== "o";
      return {
        level: state.par.ilvl || 0,
        label: numbered ? raw : tools.bulletGlyph(raw, state.par.ilvl || 0),
        kind: numbered ? "number" : "bullet",
      };
    }

    function endParagraph() {
      flush();
      var story = state.story;
      if (state.par.intbl) {
        // A paragraph inside a cell: the cell goes on.
        story.cellSegs = story.cellSegs.concat(story.segs, [{ text: "\n" }]);
        story.cellOpen = true;
        story.segs = [];
        story.label = null;
        return;
      }
      var style = state.par.s !== undefined ? styles["s" + state.par.s] : null;
      var heading = style && style.heading ? style.heading : 0;
      if (!heading && state.par.outline !== undefined && state.par.outline < 9)
        heading = state.par.outline + 1;
      story.out.push({
        type: "p",
        heading: heading,
        list: listOf(story),
        para: paraProps(),
        segments: story.segs,
      });
      story.segs = [];
    }

    function endCell() {
      flush();
      var story = state.story;
      var segs = story.cellSegs.concat(story.segs);
      story.cells.push({ segments: segs, align: state.par.align || null });
      story.cellSegs = [];
      story.segs = [];
      story.cellOpen = false;
      story.label = null;
    }

    function endRow() {
      var story = state.story;
      if (story.cellOpen || story.segs.length) endCell();
      var row = story.row;
      var prev = row.left || 0;
      story.cells.forEach(function (c, i) {
        // A cell that ends in the paragraph break of its last paragraph.
        var s = c.segments;
        if (s.length && s[s.length - 1].text === "\n") s.pop();
        if (row.cellx[i] !== undefined) {
          c.width = Math.max(0, row.cellx[i] - prev) || null;
          prev = row.cellx[i];
        }
        c.fill = row.fills[i] ? colors[row.fills[i]] || null : null;
      });
      if (story.cells.length) story.out.push({ type: "row", cells: story.cells });
      story.cells = [];
    }

    function finishStory(story) {
      flush();
      if (story.cells.length || story.cellOpen) {
        var keep = state.story;
        state.story = story;
        endRow();
        state.story = keep;
      }
      if (story.segs.length) {
        story.out.push({ type: "p", heading: 0, list: null, para: {}, segments: story.segs });
        story.segs = [];
      }
      // Rows that follow one another are one table.
      story.out.forEach(function (b, i) {
        if (b.type !== "row") return;
        b.first = !(story.out[i - 1] && story.out[i - 1].type === "row");
        b.last = !(story.out[i + 1] && story.out[i + 1].type === "row");
      });
    }

    // ---- groups

    function openGroup() {
      flush();
      stack.push(state);
      state = {
        dest: state.dest,
        story: state.story,
        chr: Object.assign({}, state.chr),
        par: state.par, // paragraph formatting belongs to the story, not the group
        uc: state.uc,
        field: state.field,
        holder: state.holder,
        pict: state.pict,
        star: false,
        first: true,
      };
      if (state.dest === "fonttbl" || state.dest === "stylesheet" || state.dest === "revtbl") {
        // One entry of the table: its name is the text up to the semicolon.
        state.holder = { text: "", chr: {}, par: {} };
        if (state.dest === "stylesheet") {
          state.chr = state.holder.chr;
          state.par = state.holder.par;
        }
        state.entry = state.dest;
        state.dest = "name";
      } else if (state.dest === "name") {
        // A group inside an entry -- an alternative font name, say.
        state.dest = "skip";
      }
    }

    function closeGroup() {
      flush();
      var done = state;
      state = stack.pop();
      if (!state) throw new Error("unbalanced braces");
      if (done.entry === "fonttbl") {
        var h = done.holder;
        if (h.f !== undefined) {
          fonts[h.f] = {
            name: h.text.replace(/;\s*$/, "").trim(),
            cp: h.charset === 2 ? "symbol" : CHARSET_CP[h.charset] || h.cpg || null,
          };
        }
      } else if (done.entry === "stylesheet") {
        var name = done.holder.text.replace(/;\s*$/, "").trim();
        var key = done.holder.styleKey || "s0";
        var m = /^heading ([1-9])$/i.exec(name);
        styles[key] = { name: name, heading: m ? parseInt(m[1], 10) : 0 };
      } else if (done.entry === "revtbl") {
        authors.push(done.holder.text.replace(/;\s*$/, "").trim());
      } else if (done.ownsStory) {
        var story = done.story;
        var keepPar = state.par;
        finishStory(story);
        state.par = keepPar;
        if (story.kind === "footnote") {
          var list = story.note.kind === "en" ? endnotes : footnotes;
          story.note.id = String(list.length + 1);
          list.push({
            type: "note",
            kind: story.note.kind,
            id: story.note.id,
            segments: flatten(story.out),
          });
        } else if (story.kind === "comment") {
          comments.push({ type: "comment", author: story.author || "Unknown", segments: flatten(story.out) });
        }
      } else if (done.isFldinst && done.field) {
        var link = /HYPERLINK\s+(?:\\l\s+)?"([^"]*)"/i.exec(done.field.inst);
        if (link) done.field.href = (/\\l\s+"/.test(done.field.inst) ? "#" : "") + link[1];
      } else if (done.ownsPict) {
        endPicture(done.pict);
      }
    }

    // A note or a comment is one block: its paragraphs run together.
    function flatten(blocks) {
      var segs = [];
      blocks.forEach(function (b, i) {
        var own = b.type === "row" ? [].concat.apply([], b.cells.map(function (c) { return c.segments.concat([{ text: " " }]); })) : b.segments;
        if (i > 0 && own.length) segs.push({ text: " " });
        segs = segs.concat(own);
      });
      return segs;
    }

    function endPicture(pict) {
      var hex = pict.hex;
      var data = new Uint8Array((hex.length >> 1) + pict.bin.length);
      var n = 0;
      for (var i = 0; i + 1 < hex.length; i += 2) data[n++] = (hex[i] << 4) | hex[i + 1];
      for (var j = 0; j < pict.bin.length; j++) data[n++] = pict.bin[j];
      data = data.subarray(0, n);
      if (!n || state.dest !== "body") return;
      var mime = tools.imageMime(data);
      var image = mime
        ? tools.imageFromBytes(data, mime)
        : { dataUrl: null, fingerprint: n + ":" + Array.prototype.slice.call(data, Math.max(0, n - 12)).join(".") };
      image.widthPx = pict.wgoal ? pict.wgoal / 15 : null;
      state.story.segs.push({
        image: image,
        trackType: state.chr.revised ? "ins" : null,
      });
    }

    function startStory(kind) {
      flush();
      var story = newStory(kind);
      state.story = story;
      state.ownsStory = true;
      state.dest = "body";
      state.par = {};
      state.chr = defaultChr();
      return story;
    }

    // ---- control words

    function destination(word) {
      switch (word) {
        case "fonttbl":
        case "stylesheet":
        case "revtbl":
        case "colortbl":
          state.dest = word;
          if (word === "colortbl") state.color = { r: 0, g: 0, b: 0, set: false };
          return true;
        case "header":
        case "headerl":
        case "headerr":
        case "headerf":
          headers.push(startStory("header"));
          return true;
        case "footer":
        case "footerl":
        case "footerr":
        case "footerf":
          footers.push(startStory("footer"));
          return true;
        case "footnote":
          if (state.dest !== "body") {
            state.dest = "skip";
            return true;
          }
          var note = { kind: "fn", id: "" };
          state.story.segs.push({ note: note });
          startStory("footnote").note = note;
          return true;
        case "annotation":
          var c = startStory("comment");
          c.author = lastAuthor.trim();
          return true;
        case "atnauthor":
          lastAuthor = "";
          state.dest = "atnauthor";
          return true;
        case "field":
          state.field = { inst: "", href: null };
          return true;
        case "fldinst":
          state.dest = "fldinst";
          state.isFldinst = true;
          return true;
        case "fldrslt":
        case "shppict":
        case "shpinst":
        case "shptxt":
        case "shp":
        case "ud":
        case "result":
          return true; // read straight through
        case "listtext":
        case "pntext":
          state.dest = state.dest === "body" ? "listtext" : "skip";
          state.story.label = "";
          return true;
        case "pict":
          if (state.dest !== "body") {
            state.dest = "skip";
            return true;
          }
          state.pict = { hex: [], bin: [], wgoal: 0 };
          state.ownsPict = true;
          state.dest = "pict";
          return true;
        case "upr":
          // {\upr{ANSI version}{\*\ud{Unicode version}}}: keep the second.
          state.dest = "upr";
          return true;
      }
      if (Object.prototype.hasOwnProperty.call(SKIP, word) && SKIP[word]) {
        state.dest = "skip";
        return true;
      }
      return false;
    }

    function control(word, param, hasParam) {
      var wasFirst = state.first,
        star = state.star;
      state.first = false;
      state.star = false;

      if (state.dest === "skip") return;
      if (state.dest === "upr") {
        // Inside \upr only the \ud group is read.
        if (word === "ud") state.dest = stack.length > 1 ? stack[stack.length - 1].dest : "body";
        else return;
      }

      if (wasFirst || star) {
        if (destination(word)) return;
        if (star) {
          // "\*" marks a destination a reader may ignore if it does not
          // know it, and this one does not.
          state.dest = "skip";
          return;
        }
      }

      if (Object.prototype.hasOwnProperty.call(CHARS, word)) {
        flush();
        emit(CHARS[word]);
        return;
      }

      var n = hasParam ? param : 1;
      var chr = state.chr,
        par = state.par;
      var holder = state.dest === "name" ? state.holder : null;

      switch (word) {
        // -- the document
        case "ansicpg":
          ansicp = param;
          return;
        case "mac":
          ansicp = 10000;
          return;
        case "deff":
          defaultFont = param;
          if (stack.length === 1) chr.f = param;
          return;
        case "uc":
          state.uc = Math.max(0, param);
          return;
        case "u":
          flush();
          emit(String.fromCharCode(param < 0 ? param + 65536 : param));
          skipChars = state.uc;
          return;
        // -- tables of fonts, colours, styles
        case "fcharset":
          if (holder) holder.charset = param;
          return;
        case "cpg":
          if (holder) holder.cpg = param;
          return;
        case "red":
        case "green":
        case "blue":
          if (state.dest === "colortbl" && state.color) {
            state.color[word.charAt(0)] = param & 255;
            state.color.set = true;
          }
          return;
        case "cs":
        case "ds":
        case "ts":
          if (holder && state.entry === "stylesheet") holder.styleKey = word + param;
          return;
        // -- paragraphs
        case "par":
        case "sect":
          if (state.dest === "body") endParagraph();
          return;
        case "pard":
          state.par = par = holder ? (holder.par = {}) : {};
          return;
        case "cell":
          if (state.dest === "body") endCell();
          return;
        case "row":
          if (state.dest === "body") endRow();
          return;
        case "nestcell":
          flush();
          emit("\n");
          return;
        case "trowd":
          state.story.row = { cellx: [], fills: [], left: 0, fill: 0 };
          return;
        case "trleft":
          state.story.row.left = param;
          return;
        case "clcbpat":
          state.story.row.fill = param;
          return;
        case "cellx":
          state.story.row.cellx.push(param);
          state.story.row.fills.push(state.story.row.fill || 0);
          state.story.row.fill = 0;
          return;
        case "intbl":
          par.intbl = true;
          return;
        case "ql":
          par.align = "left";
          return;
        case "qc":
          par.align = "center";
          return;
        case "qr":
          par.align = "right";
          return;
        case "qj":
        case "qd":
          par.align = "justify";
          return;
        case "li":
        case "ri":
        case "fi":
        case "sb":
        case "sa":
        case "ilvl":
          par[word] = param;
          return;
        case "outlinelevel":
          par.outline = param;
          return;
        case "s":
          if (holder && state.entry === "stylesheet") holder.styleKey = "s" + param;
          else par.s = param;
          return;
        // -- characters
        case "plain":
          var keep = { deleted: chr.deleted, revised: chr.revised, revauth: chr.revauth, revauthdel: chr.revauthdel };
          state.chr = Object.assign(defaultChr(), keep);
          if (holder && state.entry === "stylesheet") holder.chr = state.chr;
          return;
        case "f":
          if (holder && state.entry === "fonttbl") holder.f = param;
          flush();
          chr.f = param;
          return;
        case "fs":
          chr.fs = param;
          return;
        case "b":
        case "i":
        case "strike":
        case "caps":
        case "v":
        case "deleted":
        case "revised":
          flush();
          chr[word] = n !== 0;
          return;
        case "cf":
        case "highlight":
        case "revauth":
        case "revauthdel":
          flush();
          chr[word] = param;
          return;
        case "ulnone":
          flush();
          chr.ul = false;
          return;
        case "super":
          flush();
          chr.vert = "super";
          return;
        case "sub":
          flush();
          chr.vert = "sub";
          return;
        case "up":
        case "dn":
          // Raised or lowered by so many half-points: how some writers
          // spell superscript and subscript.
          flush();
          chr.vert = n > 0 ? (word === "up" ? "super" : "sub") : null;
          return;
        case "nosupersub":
          flush();
          chr.vert = null;
          return;
        // -- pictures and notes
        case "picwgoal":
          if (state.pict) state.pict.wgoal = param;
          return;
        case "ftnalt":
          if (state.story.note) state.story.note.kind = "en";
          return;
      }
      // Every kind of underline is an underline.
      if (/^ul/.test(word) && word !== "ulc") {
        flush();
        chr.ul = n !== 0;
      }
    }

    // ---- the stream

    var len = bytes.length;
    var i = 0;
    while (i < len) {
      var c = bytes[i];
      if (c === 0x7b) {
        openGroup();
        i++;
      } else if (c === 0x7d) {
        if (!stack.length) break; // text after the closing brace is not part of the document
        closeGroup();
        i++;
      } else if (c === 0x5c) {
        var d = bytes[i + 1];
        if (d === undefined) break;
        if ((d >= 97 && d <= 122) || (d >= 65 && d <= 90)) {
          var start = i + 1;
          var j = start;
          while (j < len && ((bytes[j] >= 97 && bytes[j] <= 122) || (bytes[j] >= 65 && bytes[j] <= 90))) j++;
          var word = String.fromCharCode.apply(null, bytes.subarray(start, j));
          var hasParam = false,
            param = 0,
            neg = false;
          if (bytes[j] === 0x2d && bytes[j + 1] >= 48 && bytes[j + 1] <= 57) {
            neg = true;
            j++;
          }
          while (j < len && bytes[j] >= 48 && bytes[j] <= 57) {
            hasParam = true;
            param = param * 10 + (bytes[j] - 48);
            j++;
          }
          if (neg) param = -param;
          if (bytes[j] === 0x20) j++; // the space that ends a control word is part of it
          i = j;
          if (word === "bin") {
            var count = Math.max(0, Math.min(param, len - i));
            if (state.pict && state.dest === "pict") {
              for (var k = 0; k < count; k++) state.pict.bin.push(bytes[i + k]);
            }
            i += count;
          } else {
            if (word !== "u") flush();
            control(word, param, hasParam);
          }
        } else if (d === 0x27) {
          var hi = hexValue(bytes[i + 2]),
            lo = hexValue(bytes[i + 3]);
          i += 4;
          state.first = false;
          if (hi >= 0 && lo >= 0 && state.dest !== "skip" && state.dest !== "pict") pushByte((hi << 4) | lo);
        } else if (d === 0x2a) {
          state.star = true;
          i += 2;
        } else {
          i += 2;
          state.first = false;
          if (state.dest === "skip") continue;
          if (d === 0x5c || d === 0x7b || d === 0x7d) pushByte(d);
          else if (d === 0x7e) {
            flush();
            emit(" ");
          } else if (d === 0x5f) {
            flush();
            emit("‑");
          } else if (d === 0x0a || d === 0x0d) {
            flush();
            if (state.dest === "body") endParagraph();
          }
          // "\-", "\|", "\:" mark things that are not text.
        }
      } else if (c === 0x0d || c === 0x0a) {
        i++; // line ends in the file mean nothing
      } else {
        i++;
        var dest = state.dest;
        if (dest === "skip" || dest === "upr") continue;
        state.first = false;
        if (dest === "pict") {
          var hv = hexValue(c);
          if (hv >= 0 && state.pict) state.pict.hex.push(hv);
        } else if (dest === "colortbl") {
          if (c === 0x3b) {
            var col = state.color;
            colors.push(
              col && col.set
                ? ((1 << 24) | (col.r << 16) | (col.g << 8) | col.b).toString(16).slice(1).toUpperCase()
                : null,
            );
            state.color = { r: 0, g: 0, b: 0, set: false };
          }
        } else if (dest === "fonttbl" || dest === "stylesheet" || dest === "revtbl") {
          // Between the entries of a table.
        } else {
          pushByte(c);
        }
      }
    }
    flush();
    finishStory(body);

    var out = [];
    function add(label, section, stories) {
      stories.forEach(function (s) {
        if (!s.out.length) return;
        out.push({ type: "divider", label: label });
        s.out.forEach(function (b) {
          b.section = section;
          out.push(b);
        });
      });
    }
    add("HEADER", "header", headers);
    out = out.concat(body.out);
    add("FOOTER", "footer", footers);
    if (footnotes.length) out.push({ type: "divider", label: "FOOTNOTES" });
    out = out.concat(footnotes);
    if (endnotes.length) out.push({ type: "divider", label: "ENDNOTES" });
    out = out.concat(endnotes);
    if (comments.length) out.push({ type: "divider", label: "COMMENTS" });
    return out.concat(comments);
  };
})();
