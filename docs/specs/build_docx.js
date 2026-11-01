// build.py が書き出した中間 JSON から Word（A4縦）を作る。
// 使い方: node build_docx.js <input.json> <output.docx>
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, ImageRun, LineRuleType,
  Footer, Header, PageNumber, TabStopType,
} = require("docx");

const [, , input, output] = process.argv;
const { title, blocks } = JSON.parse(fs.readFileSync(input, "utf8"));

const FONT = { ascii: "Arial", hAnsi: "Arial", eastAsia: "Yu Gothic", cs: "Arial" };
const NAVY = "1F4E79", LIGHT = "DDE7F3", GRID = "A6A6A6";
const PAGE_W = 11906, MARGIN = 1134, CONTENT_W = PAGE_W - MARGIN * 2; // A4・左右20mm

const wide = (s) => [...s].reduce((n, ch) => n + (/[　-鿿＀-￯]/.test(ch) ? 2 : 1), 0);
const text = (rs) => rs.map((r) => r.t).join("");
const toRuns = (rs, extra = {}) => rs.map((r) => new TextRun({
  text: r.t, bold: r.b || extra.bold, size: extra.size,
  font: r.code ? { ascii: "Consolas", hAnsi: "Consolas", eastAsia: "Yu Gothic" } : undefined,
  color: extra.color,
}));

const border = { style: BorderStyle.SINGLE, size: 4, color: GRID };
const cellBorders = { top: border, bottom: border, left: border, right: border };

function table(b) {
  const all = [b.header, ...b.rows];
  const n = b.header.length;
  const weights = [...Array(n)].map((_, j) => Math.min(40, Math.max(...all.map((r) => wide(text(r[j] || []))))) + 4);
  const sum = weights.reduce((a, w) => a + w, 0);
  const widths = weights.map((w) => Math.max(1000, Math.round((w / sum) * CONTENT_W)));
  const diff = CONTENT_W - widths.reduce((a, w) => a + w, 0);
  widths[widths.indexOf(Math.max(...widths))] += diff;
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: all.map((row, k) => new TableRow({
      tableHeader: k === 0,
      cantSplit: true,
      children: widths.map((w, j) => new TableCell({
        width: { size: w, type: WidthType.DXA },
        borders: cellBorders,
        shading: k === 0 ? { fill: LIGHT, type: ShadingType.CLEAR, color: "auto" } : undefined,
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [new Paragraph({ spacing: { after: 0 }, children: toRuns(row[j] || [], { bold: k === 0, size: 18 }) })],
      })),
    })),
  });
}

const children = [];
let listNo = 0; // 番号付きリストごとに 1 から振り直す
for (const b of blocks) {
  switch (b.type) {
    case "title":
      children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun({ text: b.text })] }));
      break;
    case "meta":
      children.push(new Paragraph({ spacing: { after: 240 }, children: toRuns(b.runs, { size: 18, color: "595959" }) }));
      break;
    case "h2":
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: b.text })] }));
      break;
    case "h3":
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: b.text })] }));
      break;
    case "p":
      children.push(new Paragraph({ children: toRuns(b.runs) }));
      break;
    case "list":
      listNo += 1;
      for (const it of b.items) {
        if (it.kind === "task") {
          children.push(new Paragraph({ indent: { left: 360, hanging: 360 }, spacing: { after: 60 },
            children: [new TextRun({ text: "☐\t" }), ...toRuns(it.runs)],
            tabStops: [{ type: TabStopType.LEFT, position: 360 }] }));
        } else {
          const ref = it.kind === "number" && it.level === 0 ? "numbers" : "bullets";
          children.push(new Paragraph({ numbering: { reference: ref, level: it.level, instance: listNo }, spacing: { after: 60 }, children: toRuns(it.runs) }));
        }
      }
      children.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
      break;
    case "table":
      children.push(table(b));
      children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
      break;
    case "image": {
      const data = fs.readFileSync(b.path);
      const w = data.readUInt32BE(16), h = data.readUInt32BE(20); // PNG ヘッダー
      const px = Math.round((CONTENT_W / 1440) * 96);
      children.push(new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 60, line: 240, lineRule: LineRuleType.AUTO },
        children: [new ImageRun({ type: "png", data, transformation: { width: px, height: Math.round((h / w) * px) },
          altText: { title: b.caption, description: b.caption, name: b.caption } })] }));
      children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
        children: [new TextRun({ text: "図：" + b.caption, size: 18, color: "595959" })] }));
      break;
    }
  }
}

const doc = new Document({
  title,
  styles: {
    default: { document: { run: { font: FONT, size: 21 }, paragraph: { spacing: { after: 120, line: 300, lineRule: LineRuleType.AUTO } } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", run: { size: 34, bold: true, color: NAVY, font: FONT },
        paragraph: { spacing: { after: 80 } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: NAVY, font: FONT },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0, keepNext: true,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, color: NAVY, font: FONT },
        paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "・", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 280 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "・", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 840, hanging: 280 } } } },
      ] },
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 360 } } } },
      ] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1134, bottom: 1134, left: MARGIN, right: MARGIN } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: title, size: 16, color: "808080" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES], size: 18, color: "808080" })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => fs.writeFileSync(output, buf));
