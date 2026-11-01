"""仕様書（Markdown）から Word（.docx）と Excel（.xlsx、A4縦印刷）を作る。

使い方: python docs/specs/build.py
  src/*.md を読み、out/ に <名前>.docx と <名前>.xlsx を出力する。
  Word は build_docx.js（npm の docx）で作るため、ここでは中間 JSON を書き出して呼び出す。
"""
import json
import math
import re
import subprocess
import unicodedata
from pathlib import Path

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties

HERE = Path(__file__).resolve().parent
SRC, OUT = HERE / "src", HERE / "out"

# ---------------------------------------------------------------- Markdown → ブロック
INLINE = re.compile(r"(\*\*.+?\*\*|`.+?`)")


def runs(text):
    out = []
    for part in INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            out.append({"t": part[2:-2], "b": True})
        elif part.startswith("`") and part.endswith("`"):
            out.append({"t": part[1:-1], "code": True})
        else:
            out.append({"t": part})
    return out


def plain(text):
    return "".join(r["t"] for r in runs(text))


def parse(md):
    blocks, lines, i = [], md.splitlines(), 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        if line.startswith("# "):
            blocks.append({"type": "title", "text": line[2:]})
        elif line.startswith("## "):
            blocks.append({"type": "h2", "text": line[3:]})
        elif line.startswith("### "):
            blocks.append({"type": "h3", "text": line[4:]})
        elif line.startswith("!["):
            m = re.match(r"!\[(.*?)\]\((.*?)\)", line)
            blocks.append({"type": "image", "caption": m.group(1), "path": str((SRC / m.group(2)).resolve())})
        elif line.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{3,}:?", c) for c in cells):
                    rows.append([runs(c) for c in cells])
                i += 1
            blocks.append({"type": "table", "header": rows[0], "rows": rows[1:]})
            continue
        elif re.match(r"(- |\d+\. )", line):
            items = []
            while i < len(lines) and re.match(r"(\s*)(- \[ \] |- |\d+\. )", lines[i]):
                m = re.match(r"(\s*)(- \[ \] |- |\d+\. )(.*)", lines[i])
                kind = "task" if "[ ]" in m.group(2) else ("number" if m.group(2)[0].isdigit() else "bullet")
                items.append({"level": 1 if len(m.group(1)) >= 2 else 0, "kind": kind, "runs": runs(m.group(3))})
                i += 1
            blocks.append({"type": "list", "items": items})
            continue
        else:
            kind = "meta" if line.startswith("作成日") else "p"
            blocks.append({"type": kind, "runs": runs(line)})
        i += 1
    return blocks


# ---------------------------------------------------------------- Excel（A4縦）
FONT = "Yu Gothic"
NAVY, LIGHT, GRID = "1F4E79", "DDE7F3", "A6A6A6"
NCOL, COLW = 12, 7.4            # B〜M の12列で本文幅を作る
FIRST = 2                        # B列
thin = Side(style="thin", color=GRID)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def width_units(s):
    """Excel の列幅単位での見かけの幅（全角=2、半角=1）。"""
    return sum(2 if unicodedata.east_asian_width(ch) in "WFA" else 1 for ch in s)


def lines_needed(text, span_cols, indent=0):
    cap = max(4, span_cols * COLW * 0.92 - indent * 2)
    return sum(max(1, math.ceil(width_units(seg) / cap)) for seg in text.split("\n"))


def put(ws, row, c0, c1, value, font=None, fill=None, align=None, border=None):
    ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
    cell = ws.cell(row=row, column=c0, value=value)
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = align or Alignment(wrap_text=True, vertical="top")
    if border:
        for c in range(c0, c1 + 1):
            ws.cell(row=row, column=c).border = border
    return cell


def table_spans(header, rows):
    n = len(header)
    weights = []
    for j in range(n):
        texts = [plain_runs(header[j])] + [plain_runs(r[j]) for r in rows if j < len(r)]
        weights.append(min(40, max(width_units(t) for t in texts)) + 4)
    spans = [max(2, round(w / sum(weights) * NCOL)) for w in weights]
    while sum(spans) > NCOL:
        spans[spans.index(max(spans))] -= 1
    while sum(spans) < NCOL:
        spans[spans.index(max(spans))] += 1
    return spans


def plain_runs(rs):
    return "".join(r["t"] for r in rs)


def build_xlsx(title, blocks, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "仕様書"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 1.5
    for c in range(FIRST, FIRST + NCOL):
        ws.column_dimensions[get_column_letter(c)].width = COLW
    ws.column_dimensions[get_column_letter(FIRST + NCOL)].width = 1.5
    last = FIRST + NCOL - 1
    body = Font(name=FONT, size=10)
    LH = 15  # 1行あたりの高さ（pt）
    CAP = 730  # A4縦・幅合わせ印刷で1ページに収まる行の高さの合計（pt）の目安
    r, y = 1, 0.0

    def height(row, h):
        nonlocal y
        ws.row_dimensions[row].height = h
        y += h

    def ensure(h):
        """次に置く h pt がページに収まらなければ、その前で改ページする。"""
        nonlocal y
        if y > 0 and y + h > CAP:
            ws.row_breaks.append(Break(id=r - 1))
            y = 0.0
            return True
        return False

    def gap(h=6):
        nonlocal r
        if y == 0:
            return
        height(r, h)
        r += 1

    for b in blocks:
        t = b["type"]
        if t == "title":
            put(ws, r, FIRST, last, b["text"], Font(name=FONT, size=16, bold=True, color=NAVY),
                align=Alignment(vertical="center", wrap_text=True))
            height(r, 30)
            r += 1
        elif t == "meta":
            put(ws, r, FIRST, last, plain_runs(b["runs"]), Font(name=FONT, size=9, color="595959"))
            height(r, LH)
            r += 1
            gap()
        elif t == "h2":
            ensure(22 + 4 + 90)  # 見出しは次の内容と同じページに置く
            gap(8)
            put(ws, r, FIRST, last, b["text"], Font(name=FONT, size=12, bold=True, color="FFFFFF"),
                PatternFill("solid", fgColor=NAVY), Alignment(vertical="center", indent=1))
            height(r, 22)
            r += 1
            gap(4)
        elif t == "h3":
            ensure(20 + 3 + 80)
            cell = put(ws, r, FIRST, last, b["text"], Font(name=FONT, size=10.5, bold=True, color=NAVY),
                       align=Alignment(vertical="bottom"))
            for c in range(FIRST, last + 1):
                ws.cell(row=r, column=c).border = Border(bottom=Side(style="thin", color=NAVY))
            height(r, 20)
            r += 1
            gap(3)
        elif t == "p":
            text = plain_runs(b["runs"])
            bold = len(b["runs"]) == 1 and b["runs"][0].get("b")
            ensure(lines_needed(text, NCOL) * LH + 3)
            put(ws, r, FIRST, last, text, Font(name=FONT, size=10, bold=bool(bold)))
            height(r, lines_needed(text, NCOL) * LH + 3)
            r += 1
            gap(4)
        elif t == "list":
            n = 0
            for it in b["items"]:
                text = plain_runs(it["runs"])
                if it["kind"] == "number" and it["level"] == 0:
                    n += 1
                    mark = f"{n}."
                else:
                    mark = {"task": "☐", "bullet": "・", "number": f"{n}."}[it["kind"]]
                ind = it["level"]
                ensure(lines_needed(text, NCOL - 1 - ind) * LH + 3)
                c0 = FIRST + ind
                ws.cell(row=r, column=c0, value=mark).font = body
                ws.cell(row=r, column=c0).alignment = Alignment(horizontal="right", vertical="top")
                put(ws, r, c0 + 1, last, text, body)
                height(r, lines_needed(text, NCOL - 1 - ind) * LH + 3)
                r += 1
            gap(4)
        elif t == "table":
            spans = table_spans(b["header"], b["rows"])
            queue = [(True, b["header"])] + [(False, row) for row in b["rows"]]
            while queue:
                is_head, row = queue.pop(0)
                k = 0 if is_head else 1
                h = max(lines_needed(plain_runs(row[j]) if j < len(row) else "", span) for j, span in enumerate(spans))
                need = h * 14 + 5 + (40 if is_head else 0)
                if ensure(need) and not is_head:
                    queue.insert(0, (False, row))  # 改ページ後は見出し行をもう一度置く
                    queue.insert(0, (True, b["header"]))
                    continue
                c = FIRST
                for j, span in enumerate(spans):
                    text = plain_runs(row[j]) if j < len(row) else ""
                    font = Font(name=FONT, size=9.5, bold=(k == 0))
                    fill = PatternFill("solid", fgColor=LIGHT) if k == 0 else None
                    put(ws, r, c, c + span - 1, text, font, fill,
                        Alignment(wrap_text=True, vertical="center" if k == 0 else "top"), BOX)
                    c += span
                height(r, h * 14 + 5)
                r += 1
            gap(6)
        elif t == "image":
            img = XLImage(b["path"])
            target_w = int(NCOL * (COLW * 7 + 5) * 0.97)
            scale = target_w / img.width
            img.width, img.height = target_w, int(img.height * scale)
            rows_needed = math.ceil(img.height / 20) + 1
            ensure(rows_needed * 15 + 20)
            ws.add_image(img, f"{get_column_letter(FIRST)}{r}")
            for k in range(rows_needed):
                height(r + k, 15)
            r += rows_needed
            put(ws, r, FIRST, last, "図：" + b["caption"], Font(name=FONT, size=9, color="595959"),
                align=Alignment(horizontal="center"))
            height(r, LH)
            r += 1
            gap()

    ws.print_area = f"A1:{get_column_letter(last + 1)}{r}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top, ws.page_margins.bottom = 0.6, 0.6
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = "&P / &N"
    ws.oddFooter.center.size = 9
    ws.oddHeader.right.text = title
    ws.oddHeader.right.size = 8
    wb.properties.title = title
    wb.save(path)


# ---------------------------------------------------------------- main
def main():
    OUT.mkdir(exist_ok=True)
    for md in sorted(SRC.glob("*.md")):
        blocks = parse(md.read_text(encoding="utf-8"))
        title = next(b["text"] for b in blocks if b["type"] == "title")
        build_xlsx(title, blocks, OUT / f"{md.stem}.xlsx")
        tmp = OUT / f".{md.stem}.json"
        tmp.write_text(json.dumps({"title": title, "blocks": blocks}, ensure_ascii=False), encoding="utf-8")
        subprocess.run(["node", str(HERE / "build_docx.js"), str(tmp), str(OUT / f"{md.stem}.docx")], check=True)
        tmp.unlink()
        print("built", md.stem)


if __name__ == "__main__":
    main()
