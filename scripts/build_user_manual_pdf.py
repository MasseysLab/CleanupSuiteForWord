"""Build and publish the CleanupSuite User Manual PDF."""

from __future__ import annotations

import argparse
import html
import re
import shutil
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    ListFlowable,
    ListItem,
    PageTemplate,
    Paragraph,
    Preformatted,
)

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "docs" / "CleanupSuite_User_Manual_v0.9.0-alpha.md"
DEFAULT_OUTPUT = ROOT / "output" / "pdf" / "CleanupSuite_User_Manual_v0.9.0-alpha.pdf"
CANONICAL_OUTPUT = ROOT / "documents" / "CleanupSuite_User_Manual.pdf"
VERSIONED_OUTPUT = ROOT / "documents" / "CleanupSuite_User_Manual_v0.9.0-alpha.pdf"

NAVY = colors.HexColor("#17324D")
BLUE = colors.HexColor("#245B89")
TEXT = colors.HexColor("#20252D")
MUTED = colors.HexColor("#5D6977")
PALE_BLUE = colors.HexColor("#EEF5FB")
RULE = colors.HexColor("#CDD8E3")


def fonts() -> tuple[str, str, str]:
    folder = Path("C:/Windows/Fonts")
    regular, bold, mono = folder / "segoeui.ttf", folder / "segoeuib.ttf", folder / "consola.ttf"
    if regular.exists() and bold.exists():
        pdfmetrics.registerFont(TTFont("CleanupSans", str(regular)))
        pdfmetrics.registerFont(TTFont("CleanupSans-Bold", str(bold)))
        if mono.exists():
            pdfmetrics.registerFont(TTFont("CleanupMono", str(mono)))
        return "CleanupSans", "CleanupSans-Bold", "CleanupMono" if mono.exists() else "Courier"
    return "Helvetica", "Helvetica-Bold", "Courier"


def inline(text: str) -> str:
    codes: list[str] = []

    def hold_code(match: re.Match[str]) -> str:
        codes.append(html.escape(match.group(1)))
        return "@@CODE%d@@" % (len(codes) - 1)

    tick = chr(96)
    text = re.sub(re.escape(tick) + "([^" + re.escape(tick) + "]+)" + re.escape(tick), hold_code, text)
    text = html.escape(text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)

    def link(match: re.Match[str]) -> str:
        url = html.unescape(match.group(2))
        return '<a href="%s" color="#245B89"><u>%s</u></a>' % (url, match.group(1))

    text = re.sub(r"\[([^]]+)\]\((https?://[^)]+|[^)]+\.pdf)\)", link, text)
    for index, value in enumerate(codes):
        text = text.replace("@@CODE%d@@" % index, '<font name="Courier">%s</font>' % value)
    return text


def styles(regular: str, bold: str, mono: str) -> dict[str, ParagraphStyle]:
    sample = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("Title", parent=sample["Title"], fontName=bold, fontSize=25,
                                leading=30, textColor=NAVY, alignment=0, spaceAfter=16),
        "h1": ParagraphStyle("H1", parent=sample["Heading1"], fontName=bold, fontSize=16,
                             leading=20, textColor=NAVY, spaceBefore=16, spaceAfter=7, keepWithNext=True),
        "h2": ParagraphStyle("H2", parent=sample["Heading2"], fontName=bold, fontSize=12.5,
                             leading=16, textColor=BLUE, spaceBefore=12, spaceAfter=5, keepWithNext=True),
        "body": ParagraphStyle("Body", parent=sample["BodyText"], fontName=regular, fontSize=9.4,
                               leading=13.4, textColor=TEXT, spaceAfter=7),
        "meta": ParagraphStyle("Meta", parent=sample["BodyText"], fontName=regular, fontSize=9.2,
                               leading=13, textColor=MUTED, leftIndent=12, borderColor=RULE,
                               borderWidth=0.6, borderPadding=8, backColor=PALE_BLUE, spaceAfter=9),
        "quote": ParagraphStyle("Quote", parent=sample["BodyText"], fontName=regular, fontSize=9.2,
                                leading=13, textColor=MUTED, leftIndent=15, borderPadding=7,
                                backColor=PALE_BLUE, spaceAfter=8),
        "list": ParagraphStyle("List", parent=sample["BodyText"], fontName=regular, fontSize=9.4,
                               leading=13.2, textColor=TEXT),
        "code": ParagraphStyle("Code", parent=sample["Code"], fontName=mono, fontSize=8.2,
                               leading=11, textColor=TEXT, leftIndent=8, rightIndent=8,
                               borderColor=RULE, borderWidth=0.5, borderPadding=7,
                               backColor=colors.HexColor("#F6F8FA"), spaceAfter=9),
    }


def markdown_story(source: str, style: dict[str, ParagraphStyle]) -> list:
    story: list = []
    paragraph: list[str] = []
    items: list[str] = []
    list_kind: str | None = None
    code: list[str] = []
    in_code = False
    title_seen = False
    metadata: list[str] = []

    def flush_paragraph() -> None:
        if paragraph:
            story.append(Paragraph(inline(" ".join(paragraph)), style["body"]))
            paragraph.clear()

    def flush_metadata() -> None:
        if metadata:
            story.append(Paragraph("<br/>".join(inline(value) for value in metadata), style["meta"]))
            metadata.clear()

    def flush_list() -> None:
        nonlocal list_kind
        if items:
            flow_items = [ListItem(Paragraph(inline(value), style["list"]), leftIndent=10) for value in items]
            list_options = {
                "bulletType": "1" if list_kind == "number" else "bullet",
                "leftIndent": 20,
                "bulletFontName": style["list"].fontName,
                "bulletFontSize": 8,
                "spaceAfter": 7,
            }
            if list_kind == "number":
                list_options["start"] = "1"
            else:
                list_options["bulletChar"] = "•"
            story.append(ListFlowable(flow_items, **list_options))
            items.clear()
            list_kind = None

    fence = chr(96) * 3
    for raw in source.splitlines():
        line = raw.rstrip()
        if line.startswith(fence):
            flush_paragraph()
            flush_metadata()
            flush_list()
            if in_code:
                story.append(Preformatted("\n".join(code), style["code"]))
                code.clear()
            in_code = not in_code
            continue
        if in_code:
            code.append(line)
            continue
        if not line.strip():
            flush_paragraph()
            flush_metadata()
            flush_list()
            continue
        heading = re.match(r"^(#{1,3})\s+(.+)$", line)
        if heading:
            flush_paragraph()
            flush_metadata()
            flush_list()
            level = len(heading.group(1))
            if level == 1 and not title_seen:
                story.append(Paragraph(inline(heading.group(2)), style["title"]))
                title_seen = True
            else:
                story.append(Paragraph(inline(heading.group(2)), style["h1"] if level == 2 else style["h2"]))
            continue
        if title_seen and line.startswith("**") and ":**" in line:
            flush_paragraph()
            flush_list()
            metadata.append(line)
            continue
        flush_metadata()
        bullet = re.match(r"^-\s+(.+)$", line)
        numbered = re.match(r"^\d+\.\s+(.+)$", line)
        if bullet or numbered:
            flush_paragraph()
            kind = "number" if numbered else "bullet"
            if list_kind and list_kind != kind:
                flush_list()
            list_kind = kind
            items.append((numbered or bullet).group(1))
            continue
        flush_list()
        if line.startswith("> "):
            flush_paragraph()
            story.append(Paragraph(inline(line[2:]), style["quote"]))
        else:
            paragraph.append(line.strip())

    flush_paragraph()
    flush_metadata()
    flush_list()
    if code:
        story.append(Preformatted("\n".join(code), style["code"]))
    return story


class ManualDocument(BaseDocTemplate):
    def __init__(self, filename: str, regular: str, **kwargs):
        super().__init__(filename, **kwargs)
        self.regular = regular
        frame = Frame(self.leftMargin, self.bottomMargin, self.width, self.height, id="manual")
        self.addPageTemplates(PageTemplate(id="manual", frames=[frame], onPage=self.page_chrome))

    def page_chrome(self, canvas, document) -> None:
        canvas.saveState()
        canvas.setStrokeColor(RULE)
        canvas.setLineWidth(0.5)
        canvas.line(document.leftMargin, 0.58 * inch, letter[0] - document.rightMargin, 0.58 * inch)
        canvas.setFont(self.regular, 7.5)
        canvas.setFillColor(MUTED)
        canvas.drawString(document.leftMargin, 0.38 * inch, "CleanupSuite User Manual · 0.9.0 Alpha · MasseysLab")
        canvas.drawRightString(letter[0] - document.rightMargin, 0.38 * inch, str(document.page))
        canvas.restoreState()


def build(input_path: Path, output_path: Path) -> None:
    regular, bold, mono = fonts()
    style = styles(regular, bold, mono)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = ManualDocument(str(output_path), regular, pagesize=letter, leftMargin=0.72 * inch,
                              rightMargin=0.72 * inch, topMargin=0.68 * inch, bottomMargin=0.76 * inch,
                              title="CleanupSuite User Manual", author="MasseysLab",
                              subject="CleanupSuite For Word 0.9.0 Alpha user manual")
    document.build(markdown_story(input_path.read_text(encoding="utf-8"), style))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    build(args.input, args.output)
    CANONICAL_OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.output, CANONICAL_OUTPUT)
    shutil.copy2(args.output, VERSIONED_OUTPUT)
    print(args.output)
    print(CANONICAL_OUTPUT)
    print(VERSIONED_OUTPUT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
