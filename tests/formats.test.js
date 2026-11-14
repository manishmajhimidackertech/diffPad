"use strict";

// The documents that are not .docx: OpenDocument, RTF, the binary .doc,
// plain text. tests/fixtures/ holds the sample pair saved in each of them by
// LibreOffice, so every reader is run on a real file of its kind; the rest
// are small documents written out here to probe one thing at a time.
//
// HTML is the one reader not run here. It needs an HTML parser, which a
// browser has and this suite's stand-in for DOMParser does not, so it is
// exercised in tests/browser/ instead.

const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");

const { loadApp, makeFile, ROOT } = require("./harness.js");
const d = require("./docx.js");

const fixture = (name) => fs.readFileSync(path.join(ROOT, "tests", "fixtures", name));
const KINDS = ["doc", "rtf", "odt", "fodt"];

// Blocks are made inside the page's own realm, so anything taken from them is
// copied into an array of this one before it is compared.
const list = (blocks, pick) => Array.from(blocks, pick);
const bodyTexts = (blocks) =>
  list(blocks.filter((b) => b.sectionLabel === "body"), (b) => b.text);

async function read(buf, name) {
  return loadApp().read(Buffer.isBuffer(buf) ? buf : Buffer.from(buf, "latin1"), name);
}

async function counts(left, right, names) {
  const app = loadApp();
  app.ctx.state.left = { fileName: names[0], blocks: await app.read(left, names[0]) };
  app.ctx.state.right = { fileName: names[1], blocks: await app.read(right, names[1]) };
  app.ctx.renderApp();
  return { counts: app.counts(), html: app.app.innerHTML };
}

// ---------------------------------------------------------------- the sample, in each format

for (const kind of KINDS) {
  test(`.${kind}: the body reads the same as the .docx it was saved from`, async () => {
    const want = bodyTexts(await read(fixture("report_v1.docx"), "report_v1.docx"));
    const got = bodyTexts(await read(fixture("report_v1." + kind), "report_v1." + kind));
    assert.deepStrictEqual(got, want);
  });

  test(`.${kind}: headings, list items, table rows and notes are what they are`, async () => {
    const blocks = await read(fixture("report_v1." + kind), "report_v1." + kind);
    const find = (text) => blocks.find((b) => b.text === text);

    const title = find("Q2 2026 Marketing Performance Report");
    assert.ok(title.isHeading && title.headingLevel === 1, "the title is a level 1 heading");
    assert.strictEqual(find("Key Metrics").headingLevel, 2);

    const item = find("Southeast Asia pilot");
    assert.ok(item.isListItem, "a list item is a list item");
    assert.strictEqual(item.ilvl, 1, "at its level");
    assert.strictEqual(find("Expanded paid social presence into two new regions.").ilvl, 0);

    const row = find("Leads Generated | 1,240 | 1,530");
    assert.ok(row.isTableRow);
    assert.strictEqual(row.cells.length, 3);
    const rows = blocks.filter((b) => b.isTableRow);
    assert.ok(rows[0].rowFirst && rows[rows.length - 1].rowLast, "the table's ends are marked");
    assert.ok(rows[0].html.includes("background:#1E3C72;"), "the header row keeps its shading");
    assert.ok(rows[0].html.includes("font-weight:700"), "and its bold");

    const link = blocks.find((b) => b.text.includes("internal analytics portal"));
    assert.ok(
      link.html.includes('data-href="https://analytics.example.com/dashboard"'),
      "the hyperlink keeps its target",
    );

    const ref = blocks.find((b) => b.sectionLabel === "body" && b.html.includes('class="fn-ref"'));
    assert.ok(ref, "the footnote reference is in the body");
    const note = blocks.find((b) => b.sectionLabel === "footnote");
    assert.ok(note.text.includes("Particularly Maria Chen"), "and the note is with the notes");

    assert.ok(
      blocks.some((b) => b.sectionLabel === "header" && b.text.startsWith("Acme Corp")),
      "the header is read",
    );
    assert.strictEqual(
      blocks.filter((b) => b.isDivider && b.text.includes("HEADER")).length,
      1,
      "once",
    );
    assert.ok(
      blocks.some((b) => b.html.includes('<img src="data:image/png;base64,')),
      "the picture is shown",
    );
  });

  test(`.${kind}: compared with itself there is nothing to report`, async () => {
    const buf = fixture("report_v2." + kind);
    const r = await counts(buf, buf, ["a." + kind, "b." + kind]);
    assert.deepStrictEqual(r.counts, { removed: 0, added: 0, fmt: 0 });
  });

  test(`.${kind}: the two versions differ where the .docx pair differs`, async () => {
    const r = await counts(fixture("report_v1." + kind), fixture("report_v2." + kind), [
      "report_v1." + kind,
      "report_v2." + kind,
    ]);
    assert.ok(r.counts.removed >= 10 && r.counts.added >= 20, JSON.stringify(r.counts));
    assert.ok(r.html.includes("A revised methodology was applied"), "the added paragraph is there");
    assert.ok(/<span class="del">.*summarizes/.test(r.html), "and the reworded one is marked word by word");
  });

  test(`.${kind}: a replaced picture is a change`, async () => {
    const a = await read(fixture("report_v1." + kind), "a." + kind);
    const b = await read(fixture("report_v2." + kind), "b." + kind);
    const pic = (blocks) => blocks.find((x) => x.text.charAt(0) === "￼").text;
    assert.notStrictEqual(pic(a), pic(b));
  });
}

test("a document can be compared with the same document in another format", async () => {
  const r = await counts(fixture("report_v1.docx"), fixture("report_v1.doc"), ["a.docx", "a.doc"]);
  // The footer's page number is a field, which .docx stores without a value.
  assert.ok(r.counts.removed <= 1 && r.counts.added <= 1, JSON.stringify(r.counts));
});

// ---------------------------------------------------------------- what a file is

test("a file is opened as what it is, whatever its name says", async () => {
  const rtfNamedDoc = await read(fixture("report_v1.rtf"), "report.doc");
  assert.ok(rtfNamedDoc.some((b) => b.text === "Key Metrics"), "RTF saved as .doc");
  const docxNamedDoc = await read(fixture("report_v1.docx"), "report.doc");
  assert.ok(docxNamedDoc.some((b) => b.text === "Key Metrics"), ".docx renamed .doc");
  const odtNamedDocx = await read(fixture("report_v1.odt"), "report.docx");
  assert.ok(odtNamedDocx.some((b) => b.text === "Key Metrics"), ".odt renamed .docx");
  const docNamedRtf = await read(fixture("report_v1.doc"), "report.rtf");
  assert.ok(docNamedRtf.some((b) => b.text === "Key Metrics"), ".doc renamed .rtf");
});

test("every kind is accepted through the file picker", async () => {
  for (const kind of KINDS) {
    const app = loadApp();
    await app.chooseFile("left", makeFile("report_v1." + kind, fixture("report_v1." + kind)));
    assert.ok(app.ctx.state.left, kind + " loaded");
  }
});

test("binary rubbish named .doc is refused rather than shown as text", async () => {
  const app = loadApp();
  await app.chooseFile("left", makeFile("x.doc", Buffer.from([1, 2, 0, 0, 3, 0, 255, 0])));
  assert.strictEqual(app.ctx.state.left, null);
  assert.strictEqual(
    app.dropzone("left").querySelector(".dz-error").textContent,
    "Could not read that file. Is it a valid document?",
  );
});

test("a password-protected .doc says that it is", async () => {
  const buf = Buffer.from(fixture("report_v1.doc"));
  // The Word stream begins with the two bytes EC A5; bit 0 of the byte at
  // offset 11 from there says the text is encrypted.
  const at = buf.indexOf(Buffer.from([0xec, 0xa5]));
  assert.ok(at > 0 && at % 512 === 0, "found the start of the Word stream");
  buf[at + 11] |= 0x01;
  const app = loadApp();
  await app.chooseFile("left", makeFile("secret.doc", buf));
  assert.strictEqual(app.ctx.state.left, null);
  assert.match(app.dropzone("left").querySelector(".dz-error").textContent, /password-protected/);
});

test("a damaged file of any kind is refused or read, and never hangs", async () => {
  for (const kind of KINDS) {
    const whole = fixture("report_v2." + kind);
    for (const keep of [0.07, 0.2, 0.41, 0.63, 0.88, 0.97]) {
      const cut = whole.subarray(0, Math.floor(whole.length * keep));
      const started = Date.now();
      await read(cut, "cut." + kind).catch(() => {});
      assert.ok(Date.now() - started < 3000, `${kind} cut to ${keep} took too long`);
    }
    const holed = Buffer.from(whole);
    holed.fill(0xff, Math.floor(holed.length / 3), Math.floor(holed.length / 3) + 700);
    await read(holed, "holed." + kind).catch(() => {});
  }
});

// ---------------------------------------------------------------- RTF

const rtf = (body) => "{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0 Times New Roman;}{\\f1\\fcharset204 Arial;}}" + body + "}";

test("rtf: paragraphs, and the three ways of writing a character", async () => {
  const blocks = await read(
    rtf("\\pard caf\\'e9 \\u8364? 5\\par \\{braces\\} and a back\\\\slash\\par {\\f1 \\'cf\\'f0\\'e8}\\par"),
    "a.rtf",
  );
  assert.deepStrictEqual(
    list(blocks, (b) => b.text),
    ["café € 5", "{braces} and a back\\slash", "При"],
  );
});

test("rtf: character formatting is scoped to its group", async () => {
  const blocks = await read(rtf("\\pard plain {\\b bold {\\i both}} plain\\par"), "a.rtf");
  const html = blocks[0].html;
  assert.match(html, /<span>plain <\/span><span style="font-weight:700;">bold <\/span><span style="font-weight:700;font-style:italic;">both<\/span><span> plain<\/span>/);
});

test("rtf: unknown destinations, hidden text and the fallback after \\u are skipped", async () => {
  const blocks = await read(
    rtf("{\\*\\generator Nobody 1.0;}{\\info{\\title Secret title}}\\pard seen{\\v unseen} \\uc2\\u960 pi done\\par"),
    "a.rtf",
  );
  assert.deepStrictEqual(list(blocks, (b) => b.text), ["seen π done"]);
});

test("rtf: a table, a field and a footnote", async () => {
  const blocks = await read(
    rtf(
      "\\trowd\\cellx2000\\cellx6000\\pard\\intbl one\\cell two\\par more\\cell\\row" +
        "\\pard See {\\field{\\*\\fldinst HYPERLINK \"https://example.com/x\"}{\\fldrslt the site}}" +
        "{\\super\\chftn}{\\footnote\\pard {\\super\\chftn} The note.}\\par",
    ),
    "a.rtf",
  );
  assert.ok(blocks[0].isTableRow);
  assert.strictEqual(blocks[0].text, "one | two\nmore");
  assert.deepStrictEqual(list(blocks[0].cells, (c) => c.width), [2000, 4000]);
  assert.ok(blocks[1].html.includes('data-href="https://example.com/x"'));
  assert.ok(blocks[1].html.includes('class="fn-ref"'));
  assert.ok(!blocks[1].text.includes("HYPERLINK"), "the field's code is not text");
  const note = blocks.find((b) => b.sectionLabel === "footnote");
  assert.ok(note.text.endsWith("The note."));
});

test("rtf: a numbered item takes its number from the file", async () => {
  const blocks = await read(
    rtf("\\pard{\\listtext\\pard\\plain 3.\\tab}\\ls1\\ilvl0 third\\par\\pard{\\listtext\\pard\\plain\\f0 \\'b7\\tab}\\ls2 dot\\par\\pard plain\\par"),
    "a.rtf",
  );
  assert.ok(blocks[0].isListItem && blocks[0].listLabel === "3." && blocks[0].listKind === "number");
  assert.ok(blocks[1].isListItem && blocks[1].listKind === "bullet");
  assert.strictEqual(blocks[1].listLabel, "•");
  assert.ok(!blocks[2].isListItem);
  assert.deepStrictEqual(list(blocks, (b) => b.text), ["third", "dot", "plain"]);
});

test("rtf: tracked changes are shown as such", async () => {
  const blocks = await read(rtf("\\pard kept {\\deleted gone }{\\revised new}\\par"), "a.rtf");
  assert.strictEqual(blocks[0].text, "kept new", "deleted text is not part of the comparison");
  assert.ok(blocks[0].html.includes('class="trk-del"'));
  assert.ok(blocks[0].html.includes('class="trk-ins"'));
});

test("rtf: a font name cannot carry anything into the page", async () => {
  const blocks = await read(
    "{\\rtf1\\ansi{\\fonttbl{\\f0 x';background:url(//evil.example/a)\"><img src=x onerror=alert(1)>;}}\\f0 text\\par}",
    "a.rtf",
  );
  const all = list(blocks, (b) => b.html + " " + (b.paraCss || "")).join(" ");
  // What is left of the name is letters, digits and spaces, inside quotes it
  // cannot close.
  assert.ok(!/url\(|<img|[";]\s*>/.test(all), all);
  assert.match(all, /font-family:'[A-Za-z0-9 _.-]*',sans-serif;/);
});

test("rtf: braces that never close still give what was read", async () => {
  const blocks = await read("{\\rtf1\\ansi{\\b one\\par two", "a.rtf");
  assert.deepStrictEqual(list(blocks, (b) => b.text), ["one", "two"]);
});

// ---------------------------------------------------------------- OpenDocument

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:xlink="http://www.w3.org/1999/xlink" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/"';

const fodt = (styles, body) =>
  `<?xml version="1.0" encoding="UTF-8"?><office:document ${NS} office:mimetype="application/vnd.oasis.opendocument.text">` +
  `<office:automatic-styles>${styles}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document>`;

test("odt: spaces, tabs and line breaks are the elements that stand for them", async () => {
  const blocks = await read(
    Buffer.from(fodt("", '<text:p>a<text:s text:c="3"/>b<text:tab/>c<text:line-break/>d\n   e</text:p>')),
    "a.fodt",
  );
  assert.strictEqual(blocks[0].text, "a   b\tc\nd e");
});

test("odt: a style's formatting reaches the text through its parent", async () => {
  const styles =
    '<style:style style:name="Base" style:family="text"><style:text-properties fo:font-weight="bold" fo:color="#1a2b3c"/></style:style>' +
    '<style:style style:name="T1" style:family="text" style:parent-style-name="Base"><style:text-properties fo:font-style="italic"/></style:style>' +
    '<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style>';
  const blocks = await read(
    Buffer.from(fodt(styles, '<text:p text:style-name="P1">plain <text:span text:style-name="T1">styled</text:span></text:p>')),
    "a.fodt",
  );
  assert.ok(blocks[0].paraCss.includes("text-align:center;"), blocks[0].paraCss);
  assert.match(blocks[0].html, /<span style="font-weight:700;font-style:italic;color:#1A2B3C;">styled<\/span>/);
});

test("odt: lists are numbered as their style says", async () => {
  const styles =
    '<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/>' +
    '<text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")"/></text:list-style>' +
    '<text:list-style style:name="L2"><text:list-level-style-bullet text:level="1" text:bullet-char="•"/></text:list-style>';
  const body =
    '<text:list text:style-name="L1"><text:list-item><text:p>one</text:p></text:list-item>' +
    '<text:list-item><text:p>two</text:p><text:list><text:list-item><text:p>inner</text:p></text:list-item><text:list-item><text:p>inner 2</text:p></text:list-item></text:list></text:list-item>' +
    '<text:list-item><text:p>three</text:p></text:list-item></text:list>' +
    '<text:list text:style-name="L2"><text:list-item><text:p>dot</text:p></text:list-item></text:list>';
  const blocks = await read(Buffer.from(fodt(styles, body)), "a.fodt");
  assert.deepStrictEqual(
    list(blocks, (b) => b.listLabel),
    ["1.", "2.", "a)", "b)", "3.", "•"],
  );
  assert.deepStrictEqual(list(blocks, (b) => b.ilvl), [0, 0, 1, 1, 0, 0]);
});

test("odt: tracked changes and comments are read", async () => {
  const body =
    '<text:tracked-changes><text:changed-region text:id="c1"><text:deletion><office:change-info><dc:creator>Ann</dc:creator></office:change-info><text:p>removed words </text:p></text:deletion></text:changed-region>' +
    '<text:changed-region text:id="c2"><text:insertion><office:change-info><dc:creator>Bob</dc:creator></office:change-info></text:insertion></text:changed-region></text:tracked-changes>' +
    '<text:p>kept <text:change text:change-id="c1"/><text:change-start text:change-id="c2"/>added<text:change-end text:change-id="c2"/> end' +
    "<office:annotation><dc:creator>Cy</dc:creator><text:p>Check this.</text:p></office:annotation></text:p>";
  const blocks = await read(Buffer.from(fodt("", body)), "a.fodt");
  assert.strictEqual(blocks[0].text, "kept added end");
  assert.ok(blocks[0].html.includes('title="Tracked deletion by Ann"'));
  assert.ok(blocks[0].html.includes('title="Tracked insertion by Bob"'));
  const comment = blocks.find((b) => b.sectionLabel === "comment");
  assert.strictEqual(comment.text, "Comment (Cy): Check this.");
});

test("odt: nothing in the file reaches the page as markup", async () => {
  const evil = '&quot;&gt;&lt;img src=x onerror=alert(1)&gt;';
  const styles = `<style:style style:name="T1" style:family="text"><style:text-properties fo:color="red;${evil}" style:font-name="x';}${evil}"/></style:style>`;
  const body = `<text:p><text:a xlink:href="${evil}"><text:span text:style-name="T1">link ${evil}</text:span></text:a></text:p>`;
  const blocks = await read(Buffer.from(fodt(styles, body)), "a.fodt");
  const html = blocks[0].html;
  assert.ok(!/<img/.test(html), html);
  assert.ok(!html.includes("color:"), "a colour that is not a colour is dropped: " + html);
  assert.match(html, /font-family:'[A-Za-z0-9 _.-]*',sans-serif;/);
  assert.ok(html.includes('data-href="&quot;&gt;&lt;img src=x onerror=alert(1)&gt;"'), html);
});

test("odt: a package without content.xml is refused", async () => {
  const zip = d.buildDocx({ omitDocument: true, extra: { mimetype: "application/vnd.oasis.opendocument.text" } });
  await assert.rejects(() => read(zip, "broken.odt"));
});

// ---------------------------------------------------------------- plain text

test("text: one paragraph to a line, whatever ends the lines", async () => {
  const blocks = await read(Buffer.from("first\r\nsecond\n\nfourth\rfifth\n"), "notes.txt");
  assert.deepStrictEqual(list(blocks, (b) => b.text), ["first", "second", "", "fourth", "fifth"]);
});

test("text: the encoding is worked out from the bytes", async () => {
  const utf8 = await read(Buffer.from("naïve café"), "a.txt");
  assert.strictEqual(utf8[0].text, "naïve café");
  const bom = await read(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from("with a mark")]), "a.txt");
  assert.strictEqual(bom[0].text, "with a mark");
  const utf16 = await read(Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from("wide é", "utf16le")]), "a.txt");
  assert.strictEqual(utf16[0].text, "wide é");
  const windows = await read(Buffer.from([0x63, 0x61, 0x66, 0xe9, 0x20, 0x93, 0x71, 0x94]), "a.txt");
  assert.strictEqual(windows[0].text, "café “q”");
});

test("text: Markdown is compared as the text it is", async () => {
  const r = await counts(Buffer.from("# Title\n\nSome *text*.\n"), Buffer.from("# Title\n\nSome **text**.\n"), ["a.md", "b.md"]);
  assert.deepStrictEqual(r.counts, { removed: 1, added: 1, fmt: 0 });
  assert.ok(r.html.includes('<span class="ins">'));
});
