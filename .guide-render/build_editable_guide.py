from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "KULLANIM_KILAVUZU.md"
IMAGE = ROOT / "KULLANIM_KILAVUZU_GORSELLERI" / "epias_sayac_kaydi.png"
OUTPUT = ROOT / "KULLANIM_KILAVUZU_DUZENLENEBILIR.docx"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
for prefix, uri in NS.items():
    if prefix not in {"rel", "ct"}:
        ET.register_namespace(prefix, uri)
ET.register_namespace("rel", NS["rel"])
ET.register_namespace("ct", NS["ct"])


def q(name: str) -> str:
    prefix, local = name.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def element(parent: ET.Element, name: str, attrs: dict[str, str] | None = None) -> ET.Element:
    return ET.SubElement(
        parent,
        q(name),
        {k if k.startswith("{") else q(k) if ":" in k else k: v for k, v in (attrs or {}).items()},
    )


def set_spacing(ppr: ET.Element, before: int = 0, after: int = 100, line: int = 276) -> None:
    element(ppr, "w:spacing", {"w:before": str(before), "w:after": str(after), "w:line": str(line), "w:lineRule": "auto"})


class PackageBuilder:
    def __init__(self) -> None:
        self.hyperlinks: list[tuple[str, str]] = []
        self.relationships: list[tuple[str, str, str, bool]] = []
        self.next_rel = 5
        self.next_num = 1
        self.numbering = self.make_numbering()
        self.body = ET.Element(q("w:body"))
        self.image_rel = "rId4"
        self.relationships.extend([
            ("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml", False),
            ("rId2", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml", False),
            ("rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", "header1.xml", False),
            (self.image_rel, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "media/epias_sayac_kaydi.png", False),
            ("rId5", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "footer1.xml", False),
            ("rId6", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml", False),
        ])
        self.next_rel = 7

    def make_numbering(self) -> ET.Element:
        root = ET.Element(q("w:numbering"))
        for abstract_id, fmt, marker, font in ((0, "decimal", "%1.", "Arial"), (1, "bullet", "•", "Symbol")):
            abstract = element(root, "w:abstractNum", {"w:abstractNumId": str(abstract_id)})
            multi = element(abstract, "w:multiLevelType", {"w:val": "singleLevel"})
            lvl = element(abstract, "w:lvl", {"w:ilvl": "0"})
            element(lvl, "w:start", {"w:val": "1"})
            element(lvl, "w:numFmt", {"w:val": fmt})
            element(lvl, "w:lvlText", {"w:val": marker})
            element(lvl, "w:lvlJc", {"w:val": "left"})
            ppr = element(lvl, "w:pPr")
            tabs = element(ppr, "w:tabs")
            element(tabs, "w:tab", {"w:val": "num", "w:pos": "420"})
            element(ppr, "w:ind", {"w:left": "420", "w:hanging": "300"})
            rpr = element(lvl, "w:rPr")
            element(rpr, "w:rFonts", {"w:ascii": font, "w:hAnsi": font})
            if fmt == "decimal":
                element(rpr, "w:b")
                element(rpr, "w:color", {"w:val": "D9782A"})
        return root

    def new_num_id(self, kind: str) -> str:
        num_id = self.next_num
        self.next_num += 1
        num = element(self.numbering, "w:num", {"w:numId": str(num_id)})
        element(num, "w:abstractNumId", {"w:val": "0" if kind == "ol" else "1"})
        return str(num_id)

    def add_run(self, parent: ET.Element, text: str, *, bold: bool = False, code: bool = False, hyperlink: str | None = None) -> None:
        target = parent
        if hyperlink:
            rel_id = f"rId{self.next_rel}"
            self.next_rel += 1
            self.relationships.append((rel_id, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", hyperlink, True))
            target = element(parent, "w:hyperlink", {"r:id": rel_id, "w:history": "1"})
        run = element(target, "w:r")
        rpr = element(run, "w:rPr")
        if bold:
            element(rpr, "w:b")
        if code:
            element(rpr, "w:rFonts", {"w:ascii": "Consolas", "w:hAnsi": "Consolas"})
            element(rpr, "w:sz", {"w:val": "18"})
            element(rpr, "w:shd", {"w:val": "clear", "w:fill": "F1F4F7"})
        if hyperlink:
            element(rpr, "w:color", {"w:val": "145EA8"})
            element(rpr, "w:u", {"w:val": "single"})
        node = element(run, "w:t")
        if text[:1].isspace() or text[-1:].isspace():
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        node.text = text

    def add_inline(self, paragraph: ET.Element, text: str) -> None:
        token = re.compile(r"(\[[^\]]+\]\([^)]+\)|\*\*.+?\*\*|`[^`]+`)")
        for part in token.split(text):
            if not part:
                continue
            link = re.fullmatch(r"\[([^\]]+)\]\(([^)]+)\)", part)
            if link:
                self.add_run(paragraph, link.group(1), hyperlink=link.group(2))
            elif part.startswith("**") and part.endswith("**"):
                self.add_run(paragraph, part[2:-2], bold=True)
            elif part.startswith("`") and part.endswith("`"):
                self.add_run(paragraph, part[1:-1], code=True)
            else:
                self.add_run(paragraph, html.unescape(part))

    def paragraph(self, text: str = "", style: str = "Normal", *, num_id: str | None = None, quote: bool = False) -> ET.Element:
        p = element(self.body, "w:p")
        ppr = element(p, "w:pPr")
        element(ppr, "w:pStyle", {"w:val": style})
        if num_id:
            numpr = element(ppr, "w:numPr")
            element(numpr, "w:ilvl", {"w:val": "0"})
            element(numpr, "w:numId", {"w:val": num_id})
        if style in {"Title", "Heading1", "Heading2"}:
            element(ppr, "w:keepNext")
            element(ppr, "w:keepLines")
        if quote:
            element(ppr, "w:ind", {"w:left": "180", "w:right": "140"})
            element(ppr, "w:shd", {"w:val": "clear", "w:fill": "F1F8F8"})
            borders = element(ppr, "w:pBdr")
            element(borders, "w:left", {"w:val": "single", "w:sz": "22", "w:space": "7", "w:color": "218E91"})
        if text:
            self.add_inline(p, text)
        return p

    def add_flow(self) -> None:
        labels = [
            ("ZIP'e sağ tıklayın", "İndirilenler'deki SKF-Yerel-Paket-1.3.zip"),
            ("Tümünü ayıkla", "Açılan menüde bu seçeneğe basın"),
            ("Gözat > Masaüstü", "SKF-Yerel klasörünü seçin"),
            ("Klasörü Seç > Ayıkla", "Dosyalar Masaüstü'ne çıkarılır"),
        ]
        widths = [2160, 440, 2160, 440, 2160, 440, 2178]
        table = element(self.body, "w:tbl")
        props = element(table, "w:tblPr")
        element(props, "w:tblW", {"w:w": "9978", "w:type": "dxa"})
        element(props, "w:tblLayout", {"w:type": "fixed"})
        margins = element(props, "w:tblCellMar")
        for side, value in (("top", "90"), ("start", "85"), ("bottom", "90"), ("end", "85")):
            element(margins, f"w:{side}", {"w:w": value, "w:type": "dxa"})
        borders = element(props, "w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            element(borders, f"w:{side}", {"w:val": "single", "w:sz": "5", "w:space": "0", "w:color": "D8E2EC"})
        grid = element(table, "w:tblGrid")
        for width in widths:
            element(grid, "w:gridCol", {"w:w": str(width)})
        row = element(table, "w:tr")
        element(element(row, "w:trPr"), "w:cantSplit")
        label_index = 0
        for index, width in enumerate(widths):
            cell = element(row, "w:tc")
            tcpr = element(cell, "w:tcPr")
            element(tcpr, "w:tcW", {"w:w": str(width), "w:type": "dxa"})
            element(tcpr, "w:vAlign", {"w:val": "center"})
            is_arrow = index % 2 == 1
            if is_arrow:
                p = element(cell, "w:p")
                ppr = element(p, "w:pPr")
                element(ppr, "w:jc", {"w:val": "center"})
                run = element(p, "w:r")
                rpr = element(run, "w:rPr")
                element(rpr, "w:color", {"w:val": "218E91"})
                element(rpr, "w:sz", {"w:val": "28"})
                text = element(run, "w:t")
                text.text = "›"
                continue
            element(tcpr, "w:shd", {"w:val": "clear", "w:fill": "F7F9FB"})
            number_p = element(cell, "w:p")
            number_r = element(number_p, "w:r")
            number_rpr = element(number_r, "w:rPr")
            element(number_rpr, "w:b")
            element(number_rpr, "w:color", {"w:val": "D9782A"})
            element(number_rpr, "w:sz", {"w:val": "20"})
            element(number_r, "w:t").text = str(label_index + 1)
            title_p = element(cell, "w:p")
            title_ppr = element(title_p, "w:pPr")
            set_spacing(title_ppr, 0, 30, 240)
            title_r = element(title_p, "w:r")
            title_rpr = element(title_r, "w:rPr")
            element(title_rpr, "w:b")
            element(title_rpr, "w:color", {"w:val": "163F75"})
            element(title_rpr, "w:sz", {"w:val": "17"})
            element(title_r, "w:t").text = labels[label_index][0]
            detail_p = element(cell, "w:p")
            detail_ppr = element(detail_p, "w:pPr")
            set_spacing(detail_ppr, 0, 0, 220)
            detail_r = element(detail_p, "w:r")
            detail_rpr = element(detail_r, "w:rPr")
            element(detail_rpr, "w:color", {"w:val": "526373"})
            element(detail_rpr, "w:sz", {"w:val": "15"})
            element(detail_r, "w:t").text = labels[label_index][1]
            label_index += 1

    def add_image(self, alt: str) -> None:
        with Image.open(IMAGE) as image:
            width_px, height_px = image.size
        max_width = int(3.25 * 914400)
        max_height = int(2.44 * 914400)
        scale = min(max_width / width_px, max_height / height_px)
        cx, cy = int(width_px * scale), int(height_px * scale)
        p = self.paragraph(style="Image")
        run = element(p, "w:r")
        drawing = element(run, "w:drawing")
        inline = element(drawing, "wp:inline", {"distT": "0", "distB": "0", "distL": "0", "distR": "0"})
        element(inline, "wp:extent", {"cx": str(cx), "cy": str(cy)})
        element(inline, "wp:docPr", {"id": "1", "name": "EPİAŞ Sayaç Kaydı", "descr": alt})
        frame = element(inline, "wp:cNvGraphicFramePr")
        element(frame, "a:graphicFrameLocks", {"noChangeAspect": "1"})
        graphic = element(inline, "a:graphic")
        graphic_data = element(graphic, "a:graphicData", {"uri": NS["pic"]})
        picture = element(graphic_data, "pic:pic")
        nonvisual = element(picture, "pic:nvPicPr")
        element(nonvisual, "pic:cNvPr", {"id": "0", "name": "epias_sayac_kaydi.png", "descr": alt})
        element(nonvisual, "pic:cNvPicPr")
        fill = element(picture, "pic:blipFill")
        element(fill, "a:blip", {"r:embed": self.image_rel})
        stretch = element(fill, "a:stretch")
        element(stretch, "a:fillRect")
        shape = element(picture, "pic:spPr")
        transform = element(shape, "a:xfrm")
        element(transform, "a:off", {"x": "0", "y": "0"})
        element(transform, "a:ext", {"cx": str(cx), "cy": str(cy)})
        geometry = element(shape, "a:prstGeom", {"prst": "rect"})
        element(geometry, "a:avLst")
        caption = self.paragraph(style="Caption")
        self.add_inline(caption, alt)

    def parse_markdown(self, text: str) -> None:
        lines = text.splitlines()
        paragraph_lines: list[str] = []
        list_kind: str | None = None
        list_num_id: str | None = None

        def flush_paragraph() -> None:
            if paragraph_lines:
                self.paragraph(" ".join(x.strip() for x in paragraph_lines))
                paragraph_lines.clear()

        def close_list() -> None:
            nonlocal list_kind, list_num_id
            list_kind, list_num_id = None, None

        i = 0
        while i < len(lines):
            raw = lines[i]
            line = raw.strip()
            if not line:
                flush_paragraph()
                close_list()
                i += 1
                continue
            if line.startswith('<div class="zip-flow"'):
                flush_paragraph()
                close_list()
                while i < len(lines) and "</div>" not in lines[i]:
                    i += 1
                if i < len(lines):
                    i += 1
                self.add_flow()
                continue
            if line.startswith("<img "):
                flush_paragraph()
                close_list()
                alt_match = re.search(r'alt="([^"]*)"', line)
                self.add_image(alt_match.group(1) if alt_match else "EPİAŞ ekran görüntüsü")
                i += 1
                continue
            heading = re.match(r"^(#{1,3})\s+(.+)$", line)
            if heading:
                flush_paragraph()
                close_list()
                level = len(heading.group(1))
                self.paragraph(heading.group(2), {1: "Title", 2: "Heading1", 3: "Heading2"}[level])
                i += 1
                continue
            if line.startswith("> "):
                flush_paragraph()
                close_list()
                quote_lines = []
                while i < len(lines) and lines[i].strip().startswith("> "):
                    quote_lines.append(lines[i].strip()[2:])
                    i += 1
                self.paragraph(" ".join(quote_lines), "Quote", quote=True)
                continue
            ordered = re.match(r"^\d+\.\s+(.+)$", line)
            unordered = re.match(r"^-\s+(.+)$", line)
            if ordered or unordered:
                flush_paragraph()
                wanted = "ol" if ordered else "ul"
                if list_kind != wanted:
                    close_list()
                    list_kind = wanted
                    list_num_id = self.new_num_id(wanted)
                self.paragraph(ordered.group(1) if ordered else unordered.group(1), num_id=list_num_id)
                i += 1
                continue
            if line.startswith("<"):
                flush_paragraph()
                close_list()
                i += 1
                continue
            close_list()
            paragraph_lines.append(line)
            i += 1
        flush_paragraph()

    def make_document(self) -> ET.Element:
        document = ET.Element(q("w:document"))
        body = element(document, "w:body")
        self.body = body
        self.parse_markdown(SOURCE.read_text(encoding="utf-8"))
        sect = element(body, "w:sectPr")
        element(sect, "w:headerReference", {"w:type": "default", "r:id": "rId3"})
        element(sect, "w:footerReference", {"w:type": "default", "r:id": "rId5"})
        element(sect, "w:pgSz", {"w:w": "11906", "w:h": "16838"})
        element(sect, "w:pgMar", {"w:top": "1134", "w:right": "964", "w:bottom": "1134", "w:left": "964", "w:header": "454", "w:footer": "454", "w:gutter": "0"})
        element(sect, "w:cols", {"w:space": "708"})
        element(sect, "w:docGrid", {"w:linePitch": "360"})
        return document

    def make_styles(self) -> ET.Element:
        root = ET.Element(q("w:styles"))
        defaults = element(root, "w:docDefaults")
        rdefaults = element(defaults, "w:rPrDefault")
        rpr = element(rdefaults, "w:rPr")
        element(rpr, "w:rFonts", {"w:ascii": "Arial", "w:hAnsi": "Arial", "w:eastAsia": "Arial", "w:cs": "Arial"})
        element(rpr, "w:sz", {"w:val": "20"})
        element(rpr, "w:color", {"w:val": "263442"})
        pdefaults = element(defaults, "w:pPrDefault")
        pp = element(pdefaults, "w:pPr")
        set_spacing(pp, 0, 95, 276)
        styles = [
            ("Normal", "Normal", 20, "263442", False, 0, 95, 276),
            ("Title", "Title", 46, "163F75", True, 0, 100, 260),
            ("Heading1", "Heading 1", 32, "163F75", True, 260, 110, 260),
            ("Heading2", "Heading 2", 23, "C96F25", True, 190, 65, 260),
            ("Quote", "Quote", 20, "24485A", False, 70, 120, 276),
            ("Image", "Image", 18, "263442", False, 50, 20, 240),
            ("Caption", "Caption", 16, "637282", False, 0, 100, 240),
        ]
        for style_id, name, size, color, bold, before, after, line in styles:
            style_attrs = {"w:type": "paragraph", "w:styleId": style_id}
            if style_id == "Normal":
                style_attrs["w:default"] = "1"
            style = element(root, "w:style", style_attrs)
            element(style, "w:name", {"w:val": name})
            if style_id != "Normal":
                element(style, "w:basedOn", {"w:val": "Normal"})
            ppr = element(style, "w:pPr")
            set_spacing(ppr, before, after, line)
            if style_id in {"Title", "Heading1", "Heading2"}:
                element(ppr, "w:keepNext")
                element(ppr, "w:keepLines")
            if style_id == "Heading1":
                borders = element(ppr, "w:pBdr")
                element(borders, "w:bottom", {"w:val": "single", "w:sz": "8", "w:space": "4", "w:color": "D9E2EC"})
            rpr = element(style, "w:rPr")
            element(rpr, "w:rFonts", {"w:ascii": "Arial", "w:hAnsi": "Arial", "w:eastAsia": "Arial", "w:cs": "Arial"})
            element(rpr, "w:sz", {"w:val": str(size)})
            element(rpr, "w:color", {"w:val": color})
            if bold:
                element(rpr, "w:b")
        return root

    def make_header(self) -> ET.Element:
        root = ET.Element(q("w:hdr"))
        p = element(root, "w:p")
        ppr = element(p, "w:pPr")
        tabs = element(ppr, "w:tabs")
        element(tabs, "w:tab", {"w:val": "right", "w:pos": "9978"})
        borders = element(ppr, "w:pBdr")
        element(borders, "w:bottom", {"w:val": "single", "w:sz": "16", "w:space": "5", "w:color": "DF7F2D"})
        for text, color, bold in (("CK ENERJİ", "163F75", True), ("YEREL WINDOWS UYGULAMASI  ·  v1.3", "163F75", True)):
            if text.startswith("YEREL"):
                tabrun = element(p, "w:r")
                element(tabrun, "w:tab")
            run = element(p, "w:r")
            rpr = element(run, "w:rPr")
            element(rpr, "w:rFonts", {"w:ascii": "Arial", "w:hAnsi": "Arial"})
            element(rpr, "w:sz", {"w:val": "16"})
            element(rpr, "w:color", {"w:val": color})
            if bold:
                element(rpr, "w:b")
            element(run, "w:t").text = text
        return root

    def make_footer(self) -> ET.Element:
        root = ET.Element(q("w:ftr"))
        p = element(root, "w:p")
        ppr = element(p, "w:pPr")
        tabs = element(ppr, "w:tabs")
        element(tabs, "w:tab", {"w:val": "right", "w:pos": "9978"})
        borders = element(ppr, "w:pBdr")
        element(borders, "w:top", {"w:val": "single", "w:sz": "5", "w:space": "5", "w:color": "D6DEE7"})
        left = element(p, "w:r")
        leftpr = element(left, "w:rPr")
        element(leftpr, "w:color", {"w:val": "708090"})
        element(leftpr, "w:sz", {"w:val": "16"})
        element(left, "w:t").text = "CK Enerji  ·  SKF Yerel Uygulama Kullanım Kılavuzu"
        tabrun = element(p, "w:r")
        element(tabrun, "w:tab")
        label = element(p, "w:r")
        labelpr = element(label, "w:rPr")
        element(labelpr, "w:color", {"w:val": "708090"})
        element(labelpr, "w:sz", {"w:val": "16"})
        element(label, "w:t").text = "Sayfa "
        for instruction in (" PAGE ", " / ", " NUMPAGES "):
            if instruction.strip() in {"PAGE", "NUMPAGES"}:
                run = element(p, "w:r")
                element(run, "w:fldChar", {"w:fldCharType": "begin"})
                run = element(p, "w:r")
                instr = element(run, "w:instrText", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
                instr.text = instruction
                run = element(p, "w:r")
                element(run, "w:fldChar", {"w:fldCharType": "separate"})
                run = element(p, "w:r")
                element(run, "w:t").text = "1"
                run = element(p, "w:r")
                element(run, "w:fldChar", {"w:fldCharType": "end"})
            else:
                run = element(p, "w:r")
                element(run, "w:t").text = instruction
        return root

    def make_settings(self) -> ET.Element:
        root = ET.Element(q("w:settings"))
        element(root, "w:updateFields", {"w:val": "true"})
        return root

    def make_rels(self) -> ET.Element:
        root = ET.Element(f"{{{NS['rel']}}}Relationships")
        for rel_id, rel_type, target, external in self.relationships:
            attrs = {"Id": rel_id, "Type": rel_type, "Target": target}
            if external:
                attrs["TargetMode"] = "External"
            ET.SubElement(root, f"{{{NS['rel']}}}Relationship", attrs)
        return root

    def write(self) -> None:
        document = self.make_document()
        styles = self.make_styles()
        header = self.make_header()
        footer = self.make_footer()
        settings = self.make_settings()
        content_types = ET.Element(f"{{{NS['ct']}}}Types")
        ET.SubElement(content_types, f"{{{NS['ct']}}}Default", {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"})
        ET.SubElement(content_types, f"{{{NS['ct']}}}Default", {"Extension": "xml", "ContentType": "application/xml"})
        ET.SubElement(content_types, f"{{{NS['ct']}}}Default", {"Extension": "png", "ContentType": "image/png"})
        for part, kind in (
            ("document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
            ("styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"),
            ("numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"),
            ("header1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"),
            ("footer1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"),
            ("settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"),
            ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
            ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
        ):
            part_name = part if part.startswith("/") else f"/word/{part}"
            ET.SubElement(content_types, f"{{{NS['ct']}}}Override", {"PartName": part_name, "ContentType": kind})
        package_rels = ET.Element(f"{{{NS['rel']}}}Relationships")
        ET.SubElement(package_rels, f"{{{NS['rel']}}}Relationship", {"Id": "rId1", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "Target": "word/document.xml"})
        ET.SubElement(package_rels, f"{{{NS['rel']}}}Relationship", {"Id": "rId2", "Type": "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "Target": "docProps/core.xml"})
        ET.SubElement(package_rels, f"{{{NS['rel']}}}Relationship", {"Id": "rId3", "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "Target": "docProps/app.xml"})
        core = ET.fromstring(b'''<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>SKF Local Application User Guide</dc:title><dc:subject>Editable user guide</dc:subject><dc:creator>CK Enerji</dc:creator></cp:coreProperties>''')
        app = ET.fromstring(b'''<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Word</Application></Properties>''')
        with zipfile.ZipFile(OUTPUT, "w", zipfile.ZIP_DEFLATED) as package:
            package.writestr("[Content_Types].xml", ET.tostring(content_types, encoding="utf-8", xml_declaration=True))
            package.writestr("_rels/.rels", ET.tostring(package_rels, encoding="utf-8", xml_declaration=True))
            package.writestr("docProps/core.xml", ET.tostring(core, encoding="utf-8", xml_declaration=True))
            package.writestr("docProps/app.xml", ET.tostring(app, encoding="utf-8", xml_declaration=True))
            package.writestr("word/document.xml", ET.tostring(document, encoding="utf-8", xml_declaration=True))
            package.writestr("word/styles.xml", ET.tostring(styles, encoding="utf-8", xml_declaration=True))
            package.writestr("word/numbering.xml", ET.tostring(self.numbering, encoding="utf-8", xml_declaration=True))
            package.writestr("word/header1.xml", ET.tostring(header, encoding="utf-8", xml_declaration=True))
            package.writestr("word/footer1.xml", ET.tostring(footer, encoding="utf-8", xml_declaration=True))
            package.writestr("word/settings.xml", ET.tostring(settings, encoding="utf-8", xml_declaration=True))
            package.writestr("word/_rels/document.xml.rels", ET.tostring(self.make_rels(), encoding="utf-8", xml_declaration=True))
            package.write(IMAGE, "word/media/epias_sayac_kaydi.png")


if __name__ == "__main__":
    PackageBuilder().write()
    with zipfile.ZipFile(OUTPUT) as docx:
        bad_file = docx.testzip()
        if bad_file:
            raise RuntimeError(f"Invalid DOCX package member: {bad_file}")
        for name in ("word/document.xml", "word/styles.xml", "word/numbering.xml", "word/header1.xml", "word/footer1.xml", "word/settings.xml", "[Content_Types].xml", "_rels/.rels", "word/_rels/document.xml.rels"):
            ET.fromstring(docx.read(name))
    print(f"Created editable Word guide: {OUTPUT.name}")
