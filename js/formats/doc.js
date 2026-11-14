/*
 * The binary Word document: .doc and .dot, as written by Word 97 to 2003
 * and still by anything asked to "save as Word 97-2003".
 *
 * The file is a small file system (a "compound file") holding a few
 * streams. WordDocument holds the text and, in 512-byte pages, the
 * formatting of each run and paragraph; the table stream holds the piece
 * table that says where in WordDocument the text is, the style sheet, the
 * lists, and the positions of the footnotes, headers and comments.
 *
 * This reads the text of every part of the document, its paragraph and
 * character formatting through the style sheet, tables, lists, fields,
 * tracked changes and pictures, and returns the same descriptions the other
 * readers do. See "what the readers hand back" in js/app.js.
 */
(function () {
  "use strict";

  var formats = (window.diffPadFormats = window.diffPadFormats || {});

  function fail(message) {
    var e = new Error(message);
    e.userMessage = message;
    throw e;
  }

  function view(bytes) {
    return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  // ---------------------------------------------------------------- compound file

  var END_OF_CHAIN = 0xfffffffe;
  var FREE = 0xffffffff;

  function readCompoundFile(bytes) {
    var dv = view(bytes);
    var sectorSize = 1 << dv.getUint16(0x1e, true);
    var miniSize = 1 << dv.getUint16(0x20, true);
    if (sectorSize < 512 || sectorSize > 4096 || miniSize > sectorSize)
      throw new Error("not a compound file");
    var dirStart = dv.getUint32(0x30, true);
    var cutoff = dv.getUint32(0x38, true);
    var miniFatStart = dv.getUint32(0x3c, true);
    var difatStart = dv.getUint32(0x44, true);
    var perSector = sectorSize / 4;
    var maxSectors = Math.floor(bytes.length / sectorSize) + 1;

    function offsetOf(sector) {
      return (sector + 1) * sectorSize;
    }

    // The sectors that hold the allocation table are listed in the header
    // and, past the first 109, in a chain of their own.
    var fatSectors = [];
    for (var i = 0; i < 109; i++) {
      var s = dv.getUint32(0x4c + i * 4, true);
      if (s < END_OF_CHAIN) fatSectors.push(s);
    }
    for (var d = difatStart, guard = 0; d < END_OF_CHAIN && guard < maxSectors; guard++) {
      var base = offsetOf(d);
      for (var k = 0; k < perSector - 1; k++) {
        var fs = dv.getUint32(base + k * 4, true);
        if (fs < END_OF_CHAIN) fatSectors.push(fs);
      }
      d = dv.getUint32(base + (perSector - 1) * 4, true);
    }
    var fat = new Uint32Array(fatSectors.length * perSector);
    fatSectors.forEach(function (sector, n) {
      var at = offsetOf(sector);
      for (var e = 0; e < perSector; e++) {
        fat[n * perSector + e] = at + e * 4 + 4 <= bytes.length ? dv.getUint32(at + e * 4, true) : FREE;
      }
    });

    function chain(start, table) {
      var out = [];
      for (var at = start; at < END_OF_CHAIN && at < table.length && out.length <= table.length; at = table[at])
        out.push(at);
      return out;
    }

    function readChain(start, size) {
      var sectors = chain(start, fat);
      var total = size === undefined ? sectors.length * sectorSize : Math.min(size, sectors.length * sectorSize);
      var out = new Uint8Array(total);
      for (var n = 0, done = 0; n < sectors.length && done < total; n++) {
        var from = offsetOf(sectors[n]);
        var take = Math.min(sectorSize, total - done, Math.max(0, bytes.length - from));
        if (take <= 0) break;
        out.set(bytes.subarray(from, from + take), done);
        done += sectorSize;
      }
      return out;
    }

    var dir = readChain(dirStart);
    var ddv = view(dir);
    var entries = [];
    for (var p = 0; p + 128 <= dir.length; p += 128) {
      var nameLen = Math.min(64, ddv.getUint16(p + 0x40, true));
      var name = "";
      for (var c = 0; c + 2 < nameLen; c += 2) name += String.fromCharCode(ddv.getUint16(p + c, true));
      entries.push({
        name: name,
        type: dir[p + 0x42],
        left: ddv.getUint32(p + 0x44, true),
        right: ddv.getUint32(p + 0x48, true),
        child: ddv.getUint32(p + 0x4c, true),
        start: ddv.getUint32(p + 0x74, true),
        size: ddv.getUint32(p + 0x78, true),
      });
    }
    if (!entries.length) throw new Error("the compound file has no directory");
    var root = entries[0];
    var miniStream = readChain(root.start, root.size);
    var miniFatBytes = readChain(miniFatStart);
    var miniFat = new Uint32Array(miniFatBytes.length >> 2);
    var mdv = view(miniFatBytes);
    for (var m = 0; m < miniFat.length; m++) miniFat[m] = mdv.getUint32(m * 4, true);

    // Only the streams directly under the root: an embedded document has
    // a WordDocument stream of its own, further down.
    var top = Object.create(null);
    var todo = [root.child];
    var seen = Object.create(null);
    while (todo.length) {
      var id = todo.pop();
      if (id >= entries.length || seen[id]) continue;
      seen[id] = true;
      var en = entries[id];
      if (en.type === 2) top[en.name.toLowerCase()] = en;
      todo.push(en.left, en.right);
    }

    return {
      stream: function (wanted) {
        var entry = top[wanted.toLowerCase()];
        if (!entry) return null;
        if (entry.size >= cutoff) return readChain(entry.start, entry.size);
        var sectors = chain(entry.start, miniFat);
        var out = new Uint8Array(entry.size);
        for (var n = 0, done = 0; n < sectors.length && done < entry.size; n++) {
          var from = sectors[n] * miniSize;
          var take = Math.min(miniSize, entry.size - done);
          out.set(miniStream.subarray(from, from + take), done);
          done += miniSize;
        }
        return out;
      },
    };
  }

  // ---------------------------------------------------------------- properties

  // A run of property changes ("sprms"): each is a two-byte code whose top
  // three bits say how long its operand is.
  function eachSprm(g, fn) {
    var i = 0;
    while (i + 2 <= g.length) {
      var sprm = g[i] | (g[i + 1] << 8);
      i += 2;
      var size;
      switch (sprm >> 13) {
        case 0:
        case 1:
          size = 1;
          break;
        case 2:
        case 4:
        case 5:
          size = 2;
          break;
        case 3:
          size = 4;
          break;
        case 7:
          size = 3;
          break;
        default:
          if (sprm === 0xd608) {
            // The table definition counts itself in two bytes.
            size = (g[i] | (g[i + 1] << 8)) + 1;
          } else if (sprm === 0xc615 && g[i] === 255) {
            var del = g[i + 1] || 0;
            var add = g[i + 2 + del * 4] || 0;
            size = 1 + 1 + del * 4 + 1 + add * 3;
          } else {
            size = 1 + (g[i] || 0);
          }
      }
      if (i + size > g.length) return;
      fn(sprm, i, size);
      i += size;
    }
  }

  var ICO = [
    null, "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00",
    "FFFFFF", "000080", "008080", "008000", "800080", "800000", "808000",
    "808080", "C0C0C0",
  ];

  function hex2(n) {
    return (n < 16 ? "0" : "") + n.toString(16).toUpperCase();
  }

  // 0 and 1 are off and on; 128 is "as the style has it" and 129 the
  // opposite of that.
  function toggle(v, styleValue) {
    if (v === 0) return false;
    if (v === 1) return true;
    if (v === 0x80) return !!styleValue;
    if (v === 0x81) return !styleValue;
    return !!v;
  }

  function applyChp(chp, g, style, fonts) {
    eachSprm(g, function (sprm, at) {
      var b = g[at];
      var w = g[at] | (g[at + 1] << 8);
      switch (sprm) {
        case 0x0835: chp.bold = toggle(b, style.bold); break;
        case 0x0836: chp.italic = toggle(b, style.italic); break;
        case 0x0837: chp.strike = toggle(b, style.strike); break;
        case 0x083b: chp.caps = toggle(b, style.caps); break;
        case 0x083c: chp.hidden = toggle(b, style.hidden); break;
        case 0x0800: chp.deleted = b !== 0; break;
        case 0x0801: chp.inserted = b !== 0; break;
        case 0x0855: chp.spec = b !== 0; break;
        case 0x0806: chp.data = b !== 0; break;
        case 0x080a: chp.ole = b !== 0; break;
        case 0x0856: chp.obj = b !== 0; break;
        case 0x2a3e: chp.underline = b !== 0; break;
        case 0x2a42: chp.color = ICO[b] || null; break;
        case 0x2a0c: chp.highlight = ICO[b] || null; break;
        case 0x2a48: chp.vert = b === 1 ? "super" : b === 2 ? "sub" : null; break;
        case 0x4a43: chp.size = w; break;
        case 0x4a4f: chp.font = fonts[w] || null; break;
        case 0x4804: chp.insAuthor = w; break;
        case 0x4863: chp.delAuthor = w; break;
        case 0x4a30: chp.istd = w; break;
        case 0x6870:
          chp.color = g[at + 3] === 0xff ? null : hex2(g[at]) + hex2(g[at + 1]) + hex2(g[at + 2]);
          break;
        case 0x6a03:
          chp.pic = (g[at] | (g[at + 1] << 8) | (g[at + 2] << 16) | (g[at + 3] << 24)) >>> 0;
          break;
        case 0x6a09:
          chp.symbol = { font: fonts[w] || "", ch: g[at + 2] | (g[at + 3] << 8) };
          chp.spec = true;
          break;
      }
    });
    return chp;
  }

  function int16(g, at) {
    var v = g[at] | (g[at + 1] << 8);
    return v > 32767 ? v - 65536 : v;
  }

  function applyPap(pap, g) {
    eachSprm(g, function (sprm, at, size) {
      var b = g[at];
      switch (sprm) {
        case 0x2403:
        case 0x2461:
          pap.align = ["left", "center", "right", "justify", "justify", "justify"][b] || "left";
          break;
        case 0x2416: pap.inTable = b !== 0; break;
        case 0x2417: pap.rowEnd = b !== 0; break;
        case 0x6649: pap.depth = g[at] | (g[at + 1] << 8); break;
        case 0x244b: pap.innerCell = b !== 0; break;
        case 0x244c: pap.innerRowEnd = b !== 0; break;
        case 0x260a: pap.ilvl = b; break;
        case 0x460b: pap.ilfo = g[at] | (g[at + 1] << 8); break;
        case 0x2640: pap.outline = b; break;
        case 0x840f:
        case 0x845e: pap.indL = int16(g, at) / 20; break;
        case 0x840e:
        case 0x845d: pap.indR = int16(g, at) / 20; break;
        case 0x8411:
        case 0x8460: pap.indFirst = int16(g, at) / 20; break;
        case 0xa413: pap.before = (g[at] | (g[at + 1] << 8)) / 20; break;
        case 0xa414: pap.after = (g[at] | (g[at + 1] << 8)) / 20; break;
        case 0xd608:
          // The row's cells: how many, and where each one's edges are.
          var count = g[at + 2];
          var edges = [];
          for (var c = 0; c <= count && at + 3 + c * 2 + 2 <= at + size; c++)
            edges.push(int16(g, at + 3 + c * 2));
          pap.edges = edges;
          break;
        case 0xd612:
          // Each cell's shading: ten bytes, the background colour in the
          // second four.
          var fills = [];
          for (var s = at + 1; s + 10 <= at + size; s += 10)
            fills.push(g[s + 7] === 0xff ? null : hex2(g[s + 4]) + hex2(g[s + 5]) + hex2(g[s + 6]));
          pap.fills = fills;
          break;
      }
    });
    return pap;
  }

  // ---------------------------------------------------------------- the document

  var CP1252_HIGH =
    "€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008dŽ\u008f" +
    "\u0090‘’“”•–—˜™š›œ\u009džŸ";

  var NFC = { 0: "decimal", 1: "upperRoman", 2: "lowerRoman", 3: "upperLetter", 4: "lowerLetter", 22: "decimalZero", 23: "bullet", 255: "none" };

  formats.doc = function (bytes, tools) {
    var cf = readCompoundFile(bytes);
    var wd = cf.stream("WordDocument");
    if (!wd || wd.length < 512) {
      fail("That is not a Word document: it may be another kind of Office file with a .doc name.");
    }
    var dv = view(wd);
    var nFib = dv.getUint16(2, true);
    var flags = dv.getUint16(0x0a, true);
    if (flags & 0x0100) fail("That document is password-protected, and cannot be read here.");
    if (nFib < 0x00c1) return readOldWord(wd, dv);

    var table = cf.stream(flags & 0x0200 ? "1Table" : "0Table");
    if (!table) throw new Error("the document has no table stream");
    var tv = view(table);
    var dataStream = cf.stream("Data");

    var pairs = dv.getUint16(152, true);
    function part(index) {
      if (index >= pairs) return { fc: 0, lcb: 0 };
      var fc = dv.getUint32(154 + index * 8, true),
        lcb = dv.getUint32(154 + index * 8 + 4, true);
      if (fc + lcb > table.length) return { fc: 0, lcb: 0 };
      return { fc: fc, lcb: lcb };
    }
    var ccp = {
      text: dv.getInt32(76, true),
      ftn: dv.getInt32(80, true),
      hdd: dv.getInt32(84, true),
      mcr: dv.getInt32(88, true),
      atn: dv.getInt32(92, true),
      edn: dv.getInt32(96, true),
      txbx: dv.getInt32(100, true),
    };

    // ---- the text, by way of the piece table

    var pieces = [];
    (function () {
      var clx = part(33);
      var at = clx.fc,
        end = clx.fc + clx.lcb;
      while (at < end && table[at] === 1) at += 3 + tv.getUint16(at + 1, true);
      if (at >= end || table[at] !== 2) throw new Error("the document has no piece table");
      var lcb = tv.getUint32(at + 1, true);
      at += 5;
      var n = Math.floor((lcb - 4) / 12);
      for (var i = 0; i < n; i++) {
        var raw = tv.getUint32(at + (n + 1) * 4 + i * 8 + 2, true);
        var compressed = (raw & 0x40000000) !== 0;
        var fc = raw & 0x3fffffff;
        pieces.push({
          start: tv.getInt32(at + i * 4, true),
          end: tv.getInt32(at + (i + 1) * 4, true),
          fc: compressed ? fc >> 1 : fc,
          width: compressed ? 1 : 2,
        });
      }
    })();

    var chunks = [];
    var total = 0;
    pieces.forEach(function (p) {
      // Characters between pieces, if a damaged file has any, keep their
      // positions.
      while (total < p.start) {
        chunks.push(" ");
        total++;
      }
      var s = "";
      var count = Math.max(0, p.end - p.start);
      for (var i = 0; i < count; i++) {
        var at = p.fc + i * p.width;
        if (at + p.width > wd.length) {
          s += " ";
          continue;
        }
        var c = p.width === 1 ? wd[at] : wd[at] | (wd[at + 1] << 8);
        if (p.width === 1 && c >= 0x80 && c <= 0x9f) s += CP1252_HIGH.charAt(c - 0x80);
        else s += String.fromCharCode(c);
        if (s.length >= 8192) {
          chunks.push(s);
          s = "";
        }
      }
      chunks.push(s);
      total = p.end;
    });
    var text = chunks.join("");

    function pieceAt(cp) {
      var lo = 0,
        hi = pieces.length - 1;
      while (lo <= hi) {
        var mid = (lo + hi) >> 1;
        if (cp < pieces[mid].start) hi = mid - 1;
        else if (cp >= pieces[mid].end) lo = mid + 1;
        else return pieces[mid];
      }
      return null;
    }

    // ---- formatting pages

    function formattingRuns(index, isPara) {
      var p = part(index);
      var n = Math.floor((p.lcb - 4) / 8);
      var runs = [];
      for (var i = 0; i < n; i++) {
        var page = tv.getUint32(p.fc + (n + 1) * 4 + i * 4, true) * 512;
        if (page + 512 > wd.length) continue;
        var count = wd[page + 511];
        var after = page + (count + 1) * 4;
        for (var r = 0; r < count; r++) {
          var run = {
            start: dv.getUint32(page + r * 4, true),
            end: dv.getUint32(page + (r + 1) * 4, true),
            g: null,
            istd: 0,
          };
          var off = wd[after + r * (isPara ? 13 : 1)] * 2;
          if (off) {
            var at = page + off;
            if (isPara) {
              var cb = wd[at];
              var len;
              if (cb === 0) {
                len = wd[at + 1] * 2;
                at += 2;
              } else {
                len = cb * 2 - 1;
                at += 1;
              }
              if (len >= 2 && at + len <= page + 512) {
                run.istd = dv.getUint16(at, true);
                run.g = wd.subarray(at + 2, at + len);
              }
            } else if (at + 1 + wd[at] <= page + 512) {
              run.g = wd.subarray(at + 1, at + 1 + wd[at]);
            }
          }
          runs.push(run);
        }
      }
      runs.sort(function (a, b) {
        return a.start - b.start;
      });
      return runs;
    }
    var chpRuns = formattingRuns(12, false);
    var papRuns = formattingRuns(13, true);

    function runAt(runs, fc) {
      var lo = 0,
        hi = runs.length - 1;
      while (lo <= hi) {
        var mid = (lo + hi) >> 1;
        if (fc < runs[mid].start) hi = mid - 1;
        else if (fc >= runs[mid].end) lo = mid + 1;
        else return runs[mid];
      }
      return null;
    }

    // ---- fonts, styles, lists, authors

    function utf16(src, at, count) {
      var s = "";
      var v = view(src);
      for (var i = 0; i < count && at + i * 2 + 2 <= src.length; i++)
        s += String.fromCharCode(v.getUint16(at + i * 2, true));
      return s;
    }

    var fonts = [];
    (function () {
      var p = part(15);
      if (!p.lcb) return;
      var count = tv.getUint16(p.fc, true);
      var at = p.fc + 4;
      for (var i = 0; i < count && at < p.fc + p.lcb; i++) {
        var size = table[at];
        var name = "";
        for (var c = at + 1 + 39; c + 2 <= at + 1 + size; c += 2) {
          var ch = tv.getUint16(c, true);
          if (!ch) break;
          name += String.fromCharCode(ch);
        }
        fonts.push(name);
        at += size + 1;
      }
    })();

    var styles = [];
    (function () {
      var p = part(1);
      if (p.lcb < 4) return;
      var cbStshi = tv.getUint16(p.fc, true);
      var count = tv.getUint16(p.fc + 2, true);
      var cbBase = tv.getUint16(p.fc + 4, true);
      var at = p.fc + 2 + cbStshi;
      var end = p.fc + p.lcb;
      for (var i = 0; i < count && at + 2 <= end; i++) {
        var cb = tv.getUint16(at, true);
        at += 2;
        if (!cb || at + cb > end) {
          styles.push(null);
          at += cb;
          continue;
        }
        var w1 = tv.getUint16(at + 2, true);
        var st = {
          sti: tv.getUint16(at, true) & 0x0fff,
          kind: w1 & 0x0f,
          base: w1 >> 4,
          papx: null,
          chpx: null,
        };
        var cupx = tv.getUint16(at + 4, true) & 0x0f;
        var cch = tv.getUint16(at + cbBase, true);
        st.name = utf16(table, at + cbBase + 2, cch);
        var u = at + cbBase + 2 + cch * 2 + 2;
        for (var k = 0; k < cupx && u + 2 <= at + cb; k++) {
          if ((u - at) & 1) u++;
          var cbUpx = tv.getUint16(u, true);
          u += 2;
          if (u + cbUpx > at + cb) break;
          var bytesOf = table.subarray(u, u + cbUpx);
          if (st.kind === 1 && k === 0) st.papx = bytesOf.subarray(Math.min(2, bytesOf.length));
          else if ((st.kind === 1 && k === 1) || (st.kind === 2 && k === 0)) st.chpx = bytesOf;
          u += cbUpx;
        }
        styles.push(st);
        at += cb;
      }
    })();

    // A style with everything it is based on laid underneath it.
    var styleCache = Object.create(null);
    function styleOf(istd) {
      if (styleCache[istd]) return styleCache[istd];
      var chain = [];
      var seen = Object.create(null);
      for (var at = istd; at < styles.length && styles[at] && !seen[at]; at = styles[at].base) {
        seen[at] = true;
        chain.unshift(styles[at]);
      }
      var out = { pap: {}, chp: {}, heading: 0 };
      chain.forEach(function (st) {
        if (st.papx) applyPap(out.pap, st.papx);
        if (st.chpx) applyChp(out.chp, st.chpx, out.chp, fonts);
        // The built-in headings have fixed numbers, whatever they are called.
        if (st.sti >= 1 && st.sti <= 9) out.heading = st.sti;
      });
      styleCache[istd] = out;
      return out;
    }

    var listLevels = Object.create(null); // list id -> levels
    var listOfFormat = []; // ilfo - 1 -> list id
    try {
      (function () {
        var p = part(73);
        if (p.lcb < 2) return;
        var count = tv.getUint16(p.fc, true);
        var at = p.fc + 2 + count * 28;
        for (var i = 0; i < count; i++) {
          var head = p.fc + 2 + i * 28;
          var id = tv.getInt32(head, true);
          var simple = (table[head + 26] & 1) !== 0;
          var levels = [];
          for (var l = 0; l < (simple ? 1 : 9); l++) {
            var nfc = table[at + 4];
            var start = tv.getInt32(at, true);
            var cbChpx = table[at + 24],
              cbPapx = table[at + 25];
            at += 28 + cbPapx + cbChpx;
            var cch = tv.getUint16(at, true);
            var raw = utf16(table, at + 2, cch);
            at += 2 + cch * 2;
            // A level's number is its text with a small number standing
            // for each level's counter.
            var shown = raw.replace(/[\u0000-\u0008]/g, function (ch) {
              return "%" + (ch.charCodeAt(0) + 1);
            });
            levels.push({ fmt: NFC[nfc] || "decimal", text: shown, start: start });
          }
          listLevels[id] = levels;
        }
        var lfo = part(74);
        if (lfo.lcb < 4) return;
        var n = tv.getUint32(lfo.fc, true);
        for (var f = 0; f < n && lfo.fc + 4 + f * 16 + 4 <= lfo.fc + lfo.lcb; f++)
          listOfFormat.push(tv.getInt32(lfo.fc + 4 + f * 16, true));
      })();
    } catch (e) {
      // Lists without their definitions are still lists: they get bullets.
      listLevels = Object.create(null);
      listOfFormat = [];
    }
    var nextLabel = tools.listCounter();

    var authors = [];
    (function () {
      var p = part(51);
      if (p.lcb < 6 || tv.getUint16(p.fc, true) !== 0xffff) return;
      var count = tv.getUint16(p.fc + 2, true);
      var extra = tv.getUint16(p.fc + 4, true);
      var at = p.fc + 6;
      for (var i = 0; i < count && at + 2 <= p.fc + p.lcb; i++) {
        var cch = tv.getUint16(at, true);
        authors.push(utf16(table, at + 2, cch));
        at += 2 + cch * 2 + extra;
      }
    })();

    // ---- where the notes and comments are

    function positions(index, dataSize) {
      var p = part(index);
      if (p.lcb < 4) return { cps: [], data: 0, n: 0 };
      var n = dataSize ? Math.floor((p.lcb - 4) / (4 + dataSize)) : Math.floor(p.lcb / 4) - 1;
      var cps = [];
      for (var i = 0; i <= n; i++) cps.push(tv.getInt32(p.fc + i * 4, true));
      return { cps: cps, data: p.fc + (n + 1) * 4, n: n };
    }
    var starts = {
      ftn: ccp.text,
      hdd: ccp.text + ccp.ftn,
      atn: ccp.text + ccp.ftn + ccp.hdd + ccp.mcr,
      edn: ccp.text + ccp.ftn + ccp.hdd + ccp.mcr + ccp.atn,
      txbx: ccp.text + ccp.ftn + ccp.hdd + ccp.mcr + ccp.atn + ccp.edn,
    };
    var noteRefs = Object.create(null); // cp -> note
    var footnotes = [],
      endnotes = [];
    [
      { refs: 2, txt: 3, kind: "fn", start: starts.ftn, list: footnotes },
      { refs: 46, txt: 47, kind: "en", start: starts.edn, list: endnotes },
    ].forEach(function (def) {
      var refs = positions(def.refs, 2);
      var txt = positions(def.txt, 0);
      for (var i = 0; i < refs.n && i + 1 < txt.cps.length; i++) {
        var note = { type: "note", kind: def.kind, id: String(i + 1), from: def.start + txt.cps[i], to: def.start + txt.cps[i + 1] };
        noteRefs[refs.cps[i]] = note;
        def.list.push(note);
      }
    });

    // ---- pictures

    function pictureAt(offset) {
      if (!dataStream || offset + 8 > dataStream.length) return null;
      var pv = view(dataStream);
      var lcb = pv.getInt32(offset, true);
      var header = pv.getUint16(offset + 4, true);
      if (lcb <= header || offset + lcb > dataStream.length) return null;
      var raw = dataStream.subarray(offset + header, offset + lcb);
      // The picture is wrapped in drawing records; the PNG or JPEG itself
      // starts wherever its signature is.
      var found = null;
      for (var i = 0; i + 8 < raw.length; i++) {
        if (raw[i] === 0x89 && raw[i + 1] === 0x50 && raw[i + 2] === 0x4e && raw[i + 3] === 0x47 && raw[i + 4] === 0x0d) {
          found = raw.subarray(i);
          break;
        }
        if (raw[i] === 0xff && raw[i + 1] === 0xd8 && raw[i + 2] === 0xff && (raw[i + 3] & 0xf0) === 0xe0) {
          found = raw.subarray(i);
          break;
        }
      }
      var image;
      var mime = found ? tools.imageMime(found) : null;
      if (mime) {
        image = tools.imageFromBytes(found, mime);
      } else {
        var sum = 0;
        for (var k = 0; k < raw.length; k++) sum = (sum * 31 + raw[k]) >>> 0;
        image = { dataUrl: null, fingerprint: raw.length + ":" + sum };
      }
      if (header >= 0x24) {
        var goal = pv.getUint16(offset + 0x1c, true),
          scale = pv.getUint16(offset + 0x20, true) || 1000;
        if (goal) image.widthPx = ((goal / 15) * scale) / 1000;
      }
      return image;
    }

    // ---- one part of the document, as blocks

    function story(from, to, section) {
      var out = [];
      var segs = [];
      var cells = [],
        cellSegs = [];
      var fields = [];
      var chp = null,
        chpEnd = from;
      to = Math.min(to, text.length);

      function formattingAt(cp) {
        var piece = pieceAt(cp);
        if (!piece) {
          chpEnd = cp + 1;
          return {};
        }
        var fc = piece.fc + (cp - piece.start) * piece.width;
        var end = piece.end;
        var pr = runAt(papRuns, fc);
        if (pr) end = Math.min(end, cp + Math.ceil((pr.end - fc) / piece.width));
        var cr = runAt(chpRuns, fc);
        if (cr) end = Math.min(end, cp + Math.ceil((cr.end - fc) / piece.width));
        chpEnd = Math.max(cp + 1, end);
        var style = styleOf(pr ? pr.istd : 0).chp;
        var out = Object.assign({}, style);
        if (cr && cr.g) {
          // A character style sits between the paragraph's and the run's own.
          var probe = applyChp({}, cr.g, style, fonts);
          if (probe.istd !== undefined && styles[probe.istd] && styles[probe.istd].kind === 2)
            Object.assign(out, styleOf(probe.istd).chp);
          applyChp(out, cr.g, style, fonts);
        }
        return out;
      }

      function papAt(cp) {
        var piece = pieceAt(cp);
        var pr = piece ? runAt(papRuns, piece.fc + (cp - piece.start) * piece.width) : null;
        var style = styleOf(pr ? pr.istd : 0);
        var pap = Object.assign({}, style.pap);
        if (pr && pr.g) applyPap(pap, pr.g);
        pap.heading = style.heading;
        pap.base = style.chp;
        return pap;
      }

      function href() {
        for (var i = fields.length - 1; i >= 0; i--) {
          if (fields[i].href) return fields[i].href;
        }
        return null;
      }

      function push(str) {
        if (!str) return;
        // Inside a field, before its separator, is the field's code.
        var top = fields[fields.length - 1];
        if (top && !top.result) {
          top.code += str;
          return;
        }
        if (chp.hidden) return;
        var link = href();
        var last = segs[segs.length - 1];
        if (last && last.chp === chp && last.href === link && last.text !== undefined) {
          last.text += str;
          return;
        }
        segs.push({
          chp: chp,
          text: str,
          bold: !!chp.bold,
          italic: !!chp.italic,
          underline: !!chp.underline,
          strike: !!chp.strike,
          caps: !!chp.caps,
          color: chp.color || null,
          highlight: chp.highlight || null,
          size: chp.size || 20,
          font: chp.font || null,
          vert: chp.vert || null,
          href: link,
          isDel: !!chp.deleted,
          trackType: chp.deleted ? "del" : chp.inserted ? "ins" : null,
          author: chp.deleted
            ? authors[chp.delAuthor] || authors[chp.insAuthor] || ""
            : chp.inserted
              ? authors[chp.insAuthor] || ""
              : null,
        });
      }

      function clean(list) {
        list.forEach(function (s) {
          delete s.chp;
        });
        return list;
      }

      function endParagraph(cp) {
        var pap = papAt(cp);
        var mark = text.charCodeAt(cp);
        var inTable = pap.inTable || pap.depth > 0;
        if (mark === 7 || (inTable && mark === 13)) {
          if (mark === 7 && pap.rowEnd) {
            // The mark that ends a row carries the row's definition.
            var edges = pap.edges || [];
            cells.forEach(function (c, i) {
              if (edges[i + 1] !== undefined) c.width = Math.max(0, edges[i + 1] - edges[i]) || null;
              if (pap.fills && pap.fills[i]) c.fill = pap.fills[i];
            });
            if (cells.length) out.push({ type: "row", section: section, cells: cells });
            cells = [];
            cellSegs = [];
            segs = [];
          } else if (mark === 7) {
            cells.push({ segments: clean(cellSegs.concat(segs)), align: pap.align || null });
            cellSegs = [];
            segs = [];
          } else if (pap.innerRowEnd) {
            segs = [];
          } else {
            // A paragraph inside a cell, or a cell of a table inside one.
            cellSegs = cellSegs.concat(segs, [{ text: "\n" }]);
            segs = [];
          }
          return;
        }
        var list = null;
        if (pap.ilfo > 0 && pap.ilfo < 2047) {
          var listId = listOfFormat[pap.ilfo - 1];
          var label = nextLabel("f" + pap.ilfo, pap.ilvl || 0, listLevels[listId]);
          if (label.label || label.kind === "bullet")
            list = { level: pap.ilvl || 0, label: label.label, kind: label.kind };
        }
        var heading = pap.heading;
        if (!heading && pap.outline !== undefined && pap.outline < 9) heading = pap.outline + 1;
        var base = pap.base || {};
        out.push({
          type: "p",
          section: section,
          heading: heading,
          list: list,
          para: {
            align: pap.align,
            indL: pap.indL,
            indR: pap.indR,
            indFirst: pap.indFirst,
            before: pap.before || 0,
            after: pap.after || 0,
          },
          base: {
            bold: !!base.bold,
            italic: !!base.italic,
            color: base.color || null,
            size: base.size || 20,
            font: base.font || null,
          },
          segments: clean(segs),
        });
        segs = [];
      }

      var buffer = "";
      function flush() {
        if (buffer) {
          push(buffer);
          buffer = "";
        }
      }

      for (var cp = from; cp < to; cp++) {
        if (cp >= chpEnd || chp === null) {
          flush();
          chp = formattingAt(cp);
        }
        var c = text.charCodeAt(cp);
        if (c >= 32 && !(chp.spec && c === 40 && chp.symbol)) {
          buffer += c === 0x2002 || c === 0x2003 ? " " : text.charAt(cp);
          continue;
        }
        flush();
        if (c === 13 || c === 7 || c === 12) {
          endParagraph(cp);
        } else if (c === 9) {
          push("\t");
        } else if (c === 11) {
          push("\n");
        } else if (c === 30) {
          push("‑");
        } else if (c === 0x13) {
          fields.push({ code: "", result: false, href: null });
        } else if (c === 0x14) {
          var open = fields[fields.length - 1];
          if (open) {
            open.result = true;
            var m = /HYPERLINK\s+(\\l\s+)?"([^"]*)"/i.exec(open.code);
            if (m) open.href = (m[1] ? "#" : "") + m[2];
          }
        } else if (c === 0x15) {
          fields.pop();
        } else if (c === 40 && chp.symbol) {
          push(tools.bulletGlyph(String.fromCharCode(chp.symbol.ch), 0));
        } else if (c === 1 && chp.spec && chp.pic !== undefined && !chp.data && !chp.ole && !chp.obj) {
          var top = fields[fields.length - 1];
          if (!(top && !top.result)) {
            var image = pictureAt(chp.pic);
            if (image) segs.push({ image: image, trackType: chp.inserted ? "ins" : null });
          }
        } else if (c === 2 && chp.spec && noteRefs[cp] && section === "body") {
          segs.push({ note: { kind: noteRefs[cp].kind, id: noteRefs[cp].id } });
        }
        // Anything else below a space marks something that is not text: a
        // drawn object, a comment's anchor, a note's own number.
      }
      flush();
      if (cells.length || cellSegs.length) {
        if (cellSegs.length || segs.length) cells.push({ segments: clean(cellSegs.concat(segs)) });
        out.push({ type: "row", section: section, cells: cells });
        segs = [];
      }
      if (segs.length) out.push({ type: "p", section: section, segments: clean(segs) });

      out.forEach(function (b, i) {
        if (b.type !== "row") return;
        b.cells.forEach(function (cell) {
          var s = cell.segments;
          if (s.length && s[s.length - 1].text === "\n") s.pop();
        });
        b.first = !(out[i - 1] && out[i - 1].type === "row");
        b.last = !(out[i + 1] && out[i + 1].type === "row");
      });
      return out;
    }

    function hasText(blocks) {
      return blocks.some(function (b) {
        return b.type === "row" || b.segments.some(function (s) {
          return s.image || s.note || /\S/.test(s.text || "");
        });
      });
    }

    function flatten(blocks) {
      var segs = [];
      blocks.forEach(function (b, i) {
        var own = b.type === "row" ? [].concat.apply([], b.cells.map(function (c) { return c.segments.concat([{ text: " " }]); })) : b.segments;
        if (i > 0 && own.length) segs.push({ text: " " });
        segs = segs.concat(own);
      });
      // A note's text ends with the mark that ends the note.
      while (segs.length && segs[segs.length - 1].text !== undefined && !/\S/.test(segs[segs.length - 1].text)) segs.pop();
      return segs;
    }

    // ---- putting the parts in order

    var result = [];

    // Headers and footers come in sixes, one six to a section, after six
    // entries that are the separators drawn above footnotes.
    var hdd = positions(11, 0);
    var headerBlocks = [],
      footerBlocks = [];
    for (var h = 6; h + 1 < hdd.cps.length; h++) {
      if (hdd.cps[h + 1] <= hdd.cps[h]) continue;
      var kind = (h - 6) % 6;
      var isHeader = kind === 0 || kind === 1 || kind === 4;
      var got = story(starts.hdd + hdd.cps[h], starts.hdd + hdd.cps[h + 1], isHeader ? "header" : "footer");
      // Each ends with an empty paragraph that is not part of it.
      while (got.length && got[got.length - 1].type === "p" && !hasText([got[got.length - 1]])) got.pop();
      if (!hasText(got)) continue;
      // The header for odd pages and the one for even pages are usually
      // the same header, stored twice.
      var key = JSON.stringify(got);
      var into = isHeader ? headerBlocks : footerBlocks;
      if (
        !into.some(function (b) {
          return JSON.stringify(b) === key;
        })
      )
        into.push(got);
    }
    headerBlocks.forEach(function (b) {
      result.push({ type: "divider", label: "HEADER" });
      result = result.concat(b);
    });

    result = result.concat(story(0, ccp.text, "body"));

    if (ccp.txbx > 0) {
      var boxes = story(starts.txbx, starts.txbx + ccp.txbx, "text box").filter(function (b) {
        return hasText([b]);
      });
      result = result.concat(boxes);
    }

    footerBlocks.forEach(function (b) {
      result.push({ type: "divider", label: "FOOTER" });
      result = result.concat(b);
    });

    [
      { list: footnotes, label: "FOOTNOTES", section: "footnote" },
      { list: endnotes, label: "ENDNOTES", section: "endnote" },
    ].forEach(function (def) {
      if (!def.list.length) return;
      result.push({ type: "divider", label: def.label });
      def.list.forEach(function (note) {
        result.push({ type: "note", kind: note.kind, id: note.id, segments: flatten(story(note.from, note.to, def.section)) });
      });
    });

    var atnRefs = positions(4, 30);
    var atnTxt = positions(5, 0);
    var comments = [];
    for (var a = 0; a < atnRefs.n && a + 1 < atnTxt.cps.length; a++) {
      var initials = utf16(table, atnRefs.data + a * 30 + 2, Math.min(9, tv.getUint16(atnRefs.data + a * 30, true)));
      comments.push({
        type: "comment",
        author: initials || "Unknown",
        segments: flatten(story(starts.atn + atnTxt.cps[a], starts.atn + atnTxt.cps[a + 1], "comment")),
      });
    }
    if (comments.length) result.push({ type: "divider", label: "COMMENTS" });
    return result.concat(comments);
  };

  // Word 6 and Word 95 keep their text in one run of single bytes. Their
  // formatting is laid out differently and is not read: the text, in its
  // paragraphs, is.
  function readOldWord(wd, dv) {
    var from = dv.getUint32(0x18, true),
      to = Math.min(wd.length, dv.getUint32(0x1c, true));
    if (to <= from) fail("That is a Word document of a kind too old to read here.");
    var raw = new TextDecoder("windows-1252").decode(wd.subarray(from, to));
    var out = [];
    var depth = 0;
    var line = "";
    for (var i = 0; i < raw.length; i++) {
      var c = raw.charCodeAt(i);
      if (c === 0x13) depth++;
      else if (c === 0x14) depth = Math.max(0, depth - 1);
      else if (c === 0x15) continue;
      else if (depth > 0) continue;
      else if (c === 13 || c === 12 || c === 7) {
        out.push({ type: "p", segments: line ? [{ text: line }] : [] });
        line = "";
      } else if (c === 9) line += "\t";
      else if (c === 11) line += "\n";
      else if (c >= 32) line += raw.charAt(i);
    }
    if (line) out.push({ type: "p", segments: [{ text: line }] });
    return out;
  }
})();
