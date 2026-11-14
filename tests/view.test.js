"use strict";

// The document as a word processor shows it: what a style, a list
// definition and a table's grid contribute to a block, and how a change is
// drawn without losing any of that.

const test = require("node:test");
const assert = require("node:assert");

const { loadApp } = require("./harness.js");
const d = require("./docx.js");

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const stylesXml = (inner) => `<?xml version="1.0"?><w:styles ${W}>${inner}</w:styles>`;
const numberingXml = (inner) => `<?xml version="1.0"?><w:numbering ${W}>${inner}</w:numbering>`;

const withParts = (body, parts) => d.buildDocx({ body, extra: parts });

// ---------------------------------------------------------------- styles

const STYLES = stylesXml(
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:rFonts w:ascii="Calibri"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Titre1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="2E74B5"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Sub"><w:name w:val="Sub"/><w:basedOn w:val="Titre1"/><w:rPr><w:sz w:val="26"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="LoopA"><w:name w:val="LoopA"/><w:basedOn w:val="LoopB"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="LoopB"><w:name w:val="LoopB"/><w:basedOn w:val="LoopA"/></w:style>',
);

test("a paragraph takes its size, weight, colour and alignment from its style", async () => {
  const app = loadApp();
  const [h, p] = await app.blocks(
    withParts([d.para(d.run("Heading"), { style: "Titre1" }), d.para(d.run("Body"))], {
      "word/styles.xml": STYLES,
    }),
  );
  assert.ok(h.isHeading && h.headingLevel === 1, "a style named 'heading 1' is a heading whatever its id");
  for (const want of ["text-align:center;", "font-size:16pt;", "font-weight:700;", "color:#2E74B5;", "margin-bottom:8pt;"]) {
    assert.ok(h.paraCss.includes(want), want + " in " + h.paraCss);
  }
  assert.strictEqual(h.html, "<span>Heading</span>", "and its runs do not repeat any of it");
  assert.ok(p.paraCss.includes("font-size:11pt;") && p.paraCss.includes("'Calibri'"), p.paraCss);
  assert.ok(!p.isHeading);
});

test("a style inherits from the one it is based on", async () => {
  const app = loadApp();
  const [b] = await app.blocks(
    withParts([d.para(d.run("Sub"), { style: "Sub" })], { "word/styles.xml": STYLES }),
  );
  assert.ok(b.paraCss.includes("font-size:13pt;"), "its own size");
  assert.ok(b.paraCss.includes("font-weight:700;") && b.paraCss.includes("text-align:center;"), "the rest inherited");
  assert.strictEqual(b.headingLevel, 1);
});

test("a run that switches off what its style switched on is drawn switched off", async () => {
  const app = loadApp();
  const [b] = await app.blocks(
    withParts(
      [
        d.para(
          [d.run("bold "), '<w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>not</w:t></w:r>'],
          { style: "Titre1" },
        ),
      ],
      { "word/styles.xml": STYLES },
    ),
  );
  assert.strictEqual(b.html, '<span>bold </span><span style="font-weight:400;">not</span>');
});

test("a character style formats the run that names it", async () => {
  const app = loadApp();
  const [b] = await app.blocks(
    withParts(
      [d.para('<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>strong</w:t></w:r>')],
      { "word/styles.xml": STYLES },
    ),
  );
  assert.strictEqual(b.html, '<span style="font-weight:700;">strong</span>');
});

test("styles based on each other in a circle do not hang the reader", async () => {
  const app = loadApp();
  const [b] = await app.blocks(
    withParts([d.para(d.run("text"), { style: "LoopA" })], { "word/styles.xml": STYLES }),
  );
  assert.strictEqual(b.text, "text");
});

test("a style id that is also a property of every object is only a style id", async () => {
  const app = loadApp();
  const [b] = await app.blocks(d.buildDocx({ body: [d.para(d.run("text"), { style: "constructor" })] }));
  assert.strictEqual(b.text, "text");
  assert.strictEqual(b.html, "<span>text</span>");
});

test("the same text with a different alignment is a formatting change", async () => {
  const app = loadApp();
  const centred = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' + d.run("same words") + "</w:p>";
  await app.diff(d.buildDocx({ body: [d.para(d.run("same words"))] }), d.buildDocx({ body: [centred] }));
  assert.deepStrictEqual(app.counts(), { removed: 0, added: 0, fmt: 1 });
});

test("measurements that are not numbers never reach the page", async () => {
  const app = loadApp();
  const evil = '<w:p><w:pPr><w:ind w:left="720;background:url(x)" w:hanging="9999999"/><w:spacing w:after="x&quot;&gt;"/></w:pPr>' +
    '<w:r><w:rPr><w:sz w:val="24;x"/><w:rFonts w:ascii="A&quot;&gt;&lt;b&gt;\'"/></w:rPr><w:t>t</w:t></w:r></w:p>';
  const [b] = await app.blocks(d.buildDocx({ body: [evil] }));
  const all = b.paraCss + b.html;
  assert.ok(!/url\(|<b>|">/.test(all.replace('">t<', "")), all);
  assert.match(b.paraCss, /^margin-left:36pt;text-indent:-36pt;/);
});

// ---------------------------------------------------------------- lists

const NUMBERING = numberingXml(
  '<w:abstractNum w:abstractNumId="0">' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/></w:lvl></w:abstractNum>' +
    '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1"/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
    '<w:num w:numId="3"><w:abstractNumId w:val="2"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride></w:num>',
);

const item = (text, numId, ilvl) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl || 0}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${d.run(text)}</w:p>`;

test("list items are numbered as the list's definition says", async () => {
  const app = loadApp();
  const blocks = await app.blocks(
    withParts(
      [item("a", 1), item("b", 1), item("c", 1, 1), item("d", 1, 1), item("e", 1), item("f", 1, 1), item("dot", 2), item("iv", 3), item("v", 3)],
      { "word/numbering.xml": NUMBERING },
    ),
  );
  assert.deepStrictEqual(
    Array.from(blocks, (b) => b.listLabel),
    ["1.", "2.", "2.a)", "2.b)", "3.", "3.a)", "•", "IV", "V"],
  );
  assert.deepStrictEqual(Array.from(blocks, (b) => b.listKind).slice(5), ["number", "bullet", "number", "number"]);
});

test("a list with no definition still shows as a list", async () => {
  const app = loadApp();
  const html = await app.diff(d.buildDocx({ body: [item("point", 9)] }), d.buildDocx({ body: [item("point", 9)] }));
  assert.match(html, /<span class="li-bullet" style="margin-left:0pt;">•<\/span>/);
});

test("numId 0 takes a paragraph out of its style's list", async () => {
  const app = loadApp();
  const [b] = await app.blocks(d.buildDocx({ body: [item("not an item", 0)] }));
  assert.ok(!b.isListItem);
});

test("an item's number is not part of what is compared", async () => {
  // Adding an item renumbers every item after it; none of those changed.
  const app = loadApp();
  const parts = { "word/numbering.xml": NUMBERING };
  await app.diff(
    withParts([item("one", 1), item("three", 1)], parts),
    withParts([item("one", 1), item("two", 1), item("three", 1)], parts),
  );
  assert.deepStrictEqual(app.counts(), { removed: 0, added: 1, fmt: 0 });
  assert.ok(app.app.innerHTML.includes(">3.</span>"), "but the new number is the one shown");
});

// ---------------------------------------------------------------- tables

const cell = (text, w, extra) =>
  `<w:tc><w:tcPr>${w ? `<w:tcW w:w="${w}" w:type="dxa"/>` : ""}${extra || ""}</w:tcPr>${d.para(d.run(text))}</w:tc>`;

test("a table's cells keep their widths, from the cell or from the grid", async () => {
  const app = loadApp();
  const tbl =
    '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell("a", 1500)}${cell("b", 4500)}</w:tr>` +
    `<w:tr>${cell("c")}${cell("d", 0, '<w:gridSpan w:val="2"/>')}</w:tr></w:tbl>`;
  const rows = await app.blocks(d.buildDocx({ body: [tbl] }));
  assert.deepStrictEqual(Array.from(rows[0].cells, (c) => c.width), [1500, 4500]);
  assert.deepStrictEqual(Array.from(rows[1].cells, (c) => c.width), [1000, 5000]);
  assert.ok(rows[0].html.includes("flex-grow:1500;"));
  assert.ok(rows[0].rowFirst && !rows[0].rowLast && rows[1].rowLast);
});

test("a changed row is drawn as a row, with the change inside its cell", async () => {
  const app = loadApp();
  const mk = (v) => d.buildDocx({ body: [d.table([[d.para(d.run("Leads")), d.para(d.run(v))]])] });
  const html = await app.diff(mk("1,530"), mk("1,610"));
  assert.deepStrictEqual(app.counts(), { removed: 1, added: 1, fmt: 0 });
  assert.match(html, /<div class="row-cell"><span>Leads<\/span><\/div><div class="row-cell"><span class="del"><span>1,530<\/span><\/span><\/div>/);
  assert.match(html, /<div class="row-cell"><span class="ins"><span>1,610<\/span><\/span><\/div>/);
  assert.ok(!html.includes("Leads | "), "not as its cells strung together");
});

test("rows with different numbers of cells are shown whole, old against new", async () => {
  const app = loadApp();
  const two = d.buildDocx({ body: [d.table([[d.para(d.run("a")), d.para(d.run("b"))]])] });
  const three = d.buildDocx({ body: [d.table([[d.para(d.run("a")), d.para(d.run("b")), d.para(d.run("c"))]])] });
  const html = await app.diff(two, three);
  assert.ok(html.includes('class="cell removed table-start is-row"'), html.slice(0, 600));
  assert.ok(html.includes('class="cell added table-start is-row"'));
});

test("the row where a table starts, and the one after it ends, are marked on both sides", async () => {
  const app = loadApp();
  const doc = d.buildDocx({
    body: [d.para(d.run("before")), d.table([[d.para(d.run("x"))], [d.para(d.run("y"))]]), d.para(d.run("after"))],
  });
  const html = await app.diff(doc, doc);
  assert.strictEqual((html.match(/class="cell table-start is-row"/g) || []).length, 2);
  assert.strictEqual((html.match(/class="cell table-end"/g) || []).length, 2);
});

// ---------------------------------------------------------------- changes keep their formatting

test("a word in a changed paragraph keeps the formatting of its run", async () => {
  const app = loadApp();
  const mk = (last) =>
    d.buildDocx({
      body: [d.para([d.run("plain "), d.run("bold", { bold: true }), d.run(" and "), d.run(last, { italic: true })])],
    });
  const html = await app.diff(mk("old"), mk("new"));
  assert.match(html, /<span>plain<\/span><span> <\/span><span style="font-weight:700;">bold<\/span>/);
  assert.match(html, /<span class="del"><span style="font-style:italic;">old<\/span><\/span>/);
  assert.match(html, /<span class="ins"><span style="font-style:italic;">new<\/span><\/span>/);
});

test("a link in a changed paragraph is still a link", async () => {
  const app = loadApp();
  const mk = (tail) =>
    d.buildDocx({
      body: [d.para([d.hyperlink("rIdL", d.run("the site")), d.run(" " + tail)])],
      rels: [{ id: "rIdL", type: d.TYPE.hyperlink, target: "https://example.com/" }],
    });
  const html = await app.diff(mk("one"), mk("two"));
  assert.ok(html.includes('<a class="doc-link" data-href="https://example.com/"'), html.slice(0, 900));
});

test("a changed heading is still drawn as a heading", async () => {
  const app = loadApp();
  const mk = (t) => withParts([d.para(d.run(t), { style: "Titre1" })], { "word/styles.xml": STYLES });
  const html = await app.diff(mk("Results for March"), mk("Results for April"));
  assert.match(html, /<div class="block-heading h1 styled" style="[^"]*font-size:16pt;[^"]*">/);
});

// ---------------------------------------------------------------- the rest of the document

test("endnotes are read, and linked to their references", async () => {
  const app = loadApp();
  const blocks = await app.blocks(
    d.buildDocx({
      body: [d.para([d.run("Claim"), '<w:r><w:endnoteReference w:id="2"/></w:r>'])],
      extra: {
        "word/endnotes.xml":
          `<?xml version="1.0"?><w:endnotes ${W}><w:endnote w:type="separator" w:id="0"><w:p/></w:endnote>` +
          `<w:endnote w:id="2">${d.para(d.run("The source."))}</w:endnote></w:endnotes>`,
      },
    }),
  );
  const note = blocks.find((b) => b.sectionLabel === "endnote");
  assert.ok(note && note.text.endsWith("The source."));
  assert.ok(blocks.some((b) => b.isDivider && b.text.includes("ENDNOTES")));
  const anchor = /href="#([^"]+)"/.exec(blocks[0].html)[1];
  assert.ok(note.html.includes('id="' + anchor + '"'), "the reference points at the note");
});

test("a footnote and an endnote with the same number are told apart", async () => {
  const app = loadApp();
  const fn = d.buildDocx({ body: [d.para([d.run("x"), d.footnoteRefRun("1")])] });
  const en = d.buildDocx({ body: [d.para([d.run("x"), '<w:r><w:endnoteReference w:id="1"/></w:r>'])] });
  await app.diff(fn, en);
  const c = app.counts();
  assert.ok(c.removed === 1 && c.added === 1, JSON.stringify(c));
});

test("a non-breaking hyphen is part of the text", async () => {
  const app = loadApp();
  const [b] = await app.blocks(d.buildDocx({ body: [d.para("<w:r><w:t>non</w:t><w:noBreakHyphen/><w:t>stop</w:t></w:r>")] }));
  assert.strictEqual(b.text, "non‑stop");
});

test("an empty paragraph is still a paragraph of the page", async () => {
  const app = loadApp();
  const doc = d.buildDocx({ body: [d.para(d.run("a")), "<w:p/>", d.para(d.run("b"))] });
  const html = await app.diff(doc, doc);
  assert.strictEqual((html.match(/<div class="block-p"><\/div>/g) || []).length, 2, "one blank line a side");
});

test("the exported page carries the rules that draw a document", async () => {
  const app = loadApp();
  await app.diff(d.simpleDocx(["a"]), d.simpleDocx(["b"]));
  await app.menu("export");
  const html = app.exportedHtml();
  for (const rule of [".row-cells{", ".block-divider", ".block-li{", ".cell.table-end{", ".block-heading.styled{"]) {
    assert.ok(html.includes(rule), rule + " is in the export");
  }
});
