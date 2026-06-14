from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Iterable

from bs4 import BeautifulSoup, NavigableString, Tag
import markdown as md
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    HRFlowable,
    ListFlowable,
    ListItem,
    PageBreak,
    PageTemplate,
    Paragraph,
    Preformatted,
    Spacer,
    Table,
    TableStyle,
)


TARGET_DOCS = [
    "DOKUMENTATION_ANWENDER.md",
    "DOKUMENTATION_TECHNIK.md",
    "SCHNELLSTART.md",
]


def cmyk_to_rgb(c: float, m: float, y: float, k: float) -> tuple[float, float, float]:
    r = (1.0 - c) * (1.0 - k)
    g = (1.0 - m) * (1.0 - k)
    b = (1.0 - y) * (1.0 - k)
    return r, g, b


PRIMARY_RGB = cmyk_to_rgb(0.0, 0.60, 1.0, 0.0)  # 0/60/100/0
PRIMARY = colors.Color(*PRIMARY_RGB)
PRIMARY_DARK = colors.Color(0.45, 0.20, 0.0)
TEXT = colors.Color(0.16, 0.16, 0.16)
MUTED = colors.Color(0.42, 0.42, 0.42)
SURFACE = colors.Color(0.97, 0.97, 0.97)
BORDER = colors.Color(0.86, 0.86, 0.86)


class DocTheme:
    def __init__(self) -> None:
        base = getSampleStyleSheet()
        self.title = ParagraphStyle(
            "DocTitle",
            parent=base["Title"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=28,
            textColor=colors.white,
            alignment=TA_CENTER,
            spaceAfter=8,
        )
        self.subtitle = ParagraphStyle(
            "DocSubtitle",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=11,
            leading=14,
            textColor=colors.white,
            alignment=TA_CENTER,
            spaceAfter=0,
        )
        self.h1 = ParagraphStyle(
            "H1",
            parent=base["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=PRIMARY_DARK,
            spaceBefore=14,
            spaceAfter=8,
        )
        self.h2 = ParagraphStyle(
            "H2",
            parent=base["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=18,
            textColor=PRIMARY_DARK,
            spaceBefore=12,
            spaceAfter=6,
        )
        self.h3 = ParagraphStyle(
            "H3",
            parent=base["Heading3"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=16,
            textColor=TEXT,
            spaceBefore=10,
            spaceAfter=4,
        )
        self.body = ParagraphStyle(
            "Body",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=10.5,
            leading=15,
            textColor=TEXT,
            spaceAfter=5,
        )
        self.quote = ParagraphStyle(
            "Quote",
            parent=self.body,
            leftIndent=10,
            borderPadding=6,
            borderColor=PRIMARY,
            borderWidth=0.8,
            borderLeft=True,
            backColor=SURFACE,
            textColor=MUTED,
            spaceBefore=4,
            spaceAfter=7,
        )
        self.code = ParagraphStyle(
            "CodeInline",
            parent=self.body,
            fontName="Courier",
            backColor=colors.Color(0.95, 0.95, 0.95),
        )
        self.caption = ParagraphStyle(
            "Caption",
            parent=self.body,
            fontSize=9,
            leading=12,
            textColor=MUTED,
            alignment=TA_RIGHT,
        )


def _draw_page_chrome(canvas, doc: BaseDocTemplate) -> None:
    width, height = A4
    canvas.saveState()

    canvas.setFillColor(colors.white)
    canvas.rect(0, 0, width, height, stroke=0, fill=1)

    canvas.setFillColor(PRIMARY)
    canvas.rect(0, height - 14 * mm, width, 14 * mm, stroke=0, fill=1)

    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(15 * mm, 14 * mm, width - 15 * mm, 14 * mm)

    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(MUTED)
    canvas.drawRightString(width - 15 * mm, 9 * mm, f"Seite {canvas.getPageNumber()}")

    canvas.restoreState()


def _inline_html(tag: Tag) -> str:
    parts: list[str] = []
    for child in tag.children:
        if isinstance(child, NavigableString):
            parts.append(str(child).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
            continue

        if not isinstance(child, Tag):
            continue

        name = child.name.lower()
        if name in {"strong", "b"}:
            parts.append(f"<b>{_inline_html(child)}</b>")
        elif name in {"em", "i"}:
            parts.append(f"<i>{_inline_html(child)}</i>")
        elif name == "code":
            parts.append(f"<font name='Courier'>{_inline_html(child)}</font>")
        elif name == "a":
            href = child.get("href", "")
            label = _inline_html(child)
            parts.append(f"<a href='{href}' color='{PRIMARY.hexval()}'>{label}</a>")
        elif name == "br":
            parts.append("<br/>")
        else:
            parts.append(_inline_html(child))
    return "".join(parts).strip()


def _list_to_flowable(tag: Tag, theme: DocTheme, ordered: bool) -> ListFlowable:
    items: list[ListItem] = []
    for li in tag.find_all("li", recursive=False):
        text = []
        sublists: list[ListFlowable] = []

        for child in li.children:
            if isinstance(child, NavigableString):
                text.append(str(child))
            elif isinstance(child, Tag):
                if child.name in {"ul", "ol"}:
                    sublists.append(_list_to_flowable(child, theme, ordered=child.name == "ol"))
                else:
                    text.append(_inline_html(child))

        li_story = [Paragraph("".join(text).strip(), theme.body)]
        li_story.extend(sublists)
        items.append(ListItem(li_story))

    bullet_type = "1" if ordered else "bullet"
    return ListFlowable(items, bulletType=bullet_type, leftIndent=14, bulletFontName="Helvetica")


def _table_to_flowable(tag: Tag, theme: DocTheme) -> Table:
    rows: list[list[str]] = []
    for tr in tag.find_all("tr"):
        cells: list[str] = []
        for cell in tr.find_all(["th", "td"]):
            cells.append(_inline_html(cell) or " ")
        if cells:
            rows.append(cells)

    if not rows:
        rows = [["(leere Tabelle)"]]

    col_count = max(len(r) for r in rows)
    for r in rows:
        if len(r) < col_count:
            r.extend([" "] * (col_count - len(r)))

    tbl_data: list[list[Paragraph]] = []
    for row_idx, row in enumerate(rows):
        row_cells: list[Paragraph] = []
        for cell in row:
            style = theme.body if row_idx > 0 else theme.h3
            row_cells.append(Paragraph(cell, style))
        tbl_data.append(row_cells)

    table = Table(tbl_data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), PRIMARY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.4, BORDER),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return table


def markdown_to_story(title: str, markdown_text: str, source_name: str, theme: DocTheme) -> list:
    html = md.markdown(
        markdown_text,
        extensions=["extra", "tables", "fenced_code", "sane_lists", "toc"],
    )
    soup = BeautifulSoup(html, "html.parser")

    story: list = []

    # Cover
    story.append(Spacer(1, 45 * mm))
    story.append(Paragraph(title, theme.title))
    story.append(Paragraph("Bewerbungsverwaltung · Dokumentation", theme.subtitle))
    story.append(Spacer(1, 8 * mm))
    story.append(HRFlowable(width="45%", thickness=1.4, color=colors.white, hAlign="CENTER"))
    story.append(Spacer(1, 6 * mm))
    generated = dt.datetime.now().strftime("%d.%m.%Y %H:%M")
    story.append(Paragraph(f"Quelle: {source_name}<br/>Generiert am: {generated}", theme.subtitle))
    story.append(PageBreak())

    for node in soup.contents:
        if isinstance(node, NavigableString):
            stripped = str(node).strip()
            if stripped:
                story.append(Paragraph(stripped, theme.body))
            continue

        if not isinstance(node, Tag):
            continue

        name = node.name.lower()

        if name == "h1":
            story.append(Paragraph(_inline_html(node), theme.h1))
        elif name == "h2":
            story.append(Paragraph(_inline_html(node), theme.h2))
        elif name == "h3":
            story.append(Paragraph(_inline_html(node), theme.h3))
        elif name == "p":
            story.append(Paragraph(_inline_html(node), theme.body))
        elif name == "blockquote":
            story.append(Paragraph(_inline_html(node), theme.quote))
        elif name == "ul":
            story.append(_list_to_flowable(node, theme, ordered=False))
            story.append(Spacer(1, 2 * mm))
        elif name == "ol":
            story.append(_list_to_flowable(node, theme, ordered=True))
            story.append(Spacer(1, 2 * mm))
        elif name == "pre":
            code_text = node.get_text("", strip=False).rstrip()
            story.append(
                Preformatted(
                    code_text,
                    ParagraphStyle(
                        "Pre",
                        parent=theme.body,
                        fontName="Courier",
                        fontSize=9,
                        leading=12,
                        backColor=colors.Color(0.95, 0.95, 0.95),
                        borderColor=BORDER,
                        borderWidth=0.5,
                        borderPadding=6,
                    ),
                )
            )
            story.append(Spacer(1, 2 * mm))
        elif name == "hr":
            story.append(HRFlowable(width="100%", color=BORDER, thickness=0.8))
            story.append(Spacer(1, 2 * mm))
        elif name == "table":
            story.append(_table_to_flowable(node, theme))
            story.append(Spacer(1, 2 * mm))
        else:
            text = _inline_html(node)
            if text:
                story.append(Paragraph(text, theme.body))

    return story


def _resolve_source_files(docs_dir: Path, file_names: Iterable[str]) -> list[Path]:
    resolved: list[Path] = []
    by_lower = {p.name.lower(): p for p in docs_dir.glob("*.md")}

    for name in file_names:
        key = name.lower()
        if key in by_lower:
            resolved.append(by_lower[key])
        else:
            candidate = docs_dir / name
            if candidate.exists():
                resolved.append(candidate)
            else:
                raise FileNotFoundError(f"Dokument nicht gefunden: {name}")

    return resolved


def build_pdf(source_md: Path, output_pdf: Path) -> None:
    theme = DocTheme()
    markdown_text = source_md.read_text(encoding="utf-8")

    title = source_md.stem.replace("_", " ")
    story = markdown_to_story(title=title, markdown_text=markdown_text, source_name=source_md.name, theme=theme)

    output_pdf.parent.mkdir(parents=True, exist_ok=True)

    frame = Frame(16 * mm, 18 * mm, A4[0] - 32 * mm, A4[1] - 38 * mm, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    template = PageTemplate(id="default", frames=[frame], onPage=_draw_page_chrome)

    doc = BaseDocTemplate(
        str(output_pdf),
        pagesize=A4,
        title=title,
        author="Bewerbungsverwaltung",
    )
    doc.addPageTemplates([template])
    doc.build(story)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Konvertiert ausgewählte Markdown-Dokumentationen in modern gestaltete PDFs.")
    parser.add_argument("--project-root", type=Path, default=Path(__file__).resolve().parents[2], help="Projektwurzel (Default: automatisch erkannt).")
    parser.add_argument("--out-dir", type=Path, default=None, help="Ausgabeordner für PDFs (Default: <project>/docs/pdf).")
    parser.add_argument(
        "--files",
        nargs="+",
        default=TARGET_DOCS,
        help="Markdown-Dateien aus docs/, die konvertiert werden sollen.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = args.project_root.resolve()
    docs_dir = project_root / "docs"
    out_dir = args.out_dir.resolve() if args.out_dir else (docs_dir / "pdf")

    if not docs_dir.exists():
        raise FileNotFoundError(f"docs-Verzeichnis nicht gefunden: {docs_dir}")

    files = _resolve_source_files(docs_dir, args.files)
    for source_file in files:
        output_file = out_dir / f"{source_file.stem}.pdf"
        build_pdf(source_file, output_file)
        print(f"[PDF] erstellt: {output_file}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
