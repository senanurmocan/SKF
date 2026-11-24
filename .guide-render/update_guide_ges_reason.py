from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from lxml import etree as ET


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "KULLANIM_KILAVUZU_DUZENLENEBILIR_GUNCEL.docx"
OUTPUT = ROOT / "KULLANIM_KILAVUZU_DUZENLENEBILIR_GUNCEL_NEXT.docx"
DOCUMENT_XML = "word/document.xml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def paragraph_text(paragraph: ET._Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def replace_paragraph(paragraph: ET._Element, text: str) -> None:
    properties = paragraph.find("w:pPr", NS)
    for child in list(paragraph):
        if child is not properties:
            paragraph.remove(child)
    run = ET.SubElement(paragraph, f"{{{W}}}r")
    ET.SubElement(run, f"{{{W}}}t").text = text


def main() -> None:
    if OUTPUT.exists():
        raise SystemExit(f"Output already exists; refusing to overwrite: {OUTPUT}")
    with ZipFile(SOURCE, "r") as source_zip:
        root = ET.fromstring(source_zip.read(DOCUMENT_XML))
        body = root.find("w:body", NS)
        if body is None:
            raise RuntimeError("Word document body not found")
        ges_updated = False
        comparison_updated = False
        for paragraph in body.findall("w:p", NS):
            text = paragraph_text(paragraph)
            if text.startswith("GES Üretimi (isteğe bağlı):"):
                replace_paragraph(
                    paragraph,
                    "GES Üretimi (isteğe bağlı): Tedarikçi LUM bildirimi biçimindeki .xlsx dosyasını GES Üretimi düğmesiyle seçin. Bu dosya, dağıtım ile uzlaştırma arasındaki farkın GES üretimiyle ilişkili olup olmadığını değerlendirmenize yardımcı olmak için karşılaştırmaya eklenir. Üretim miktarları ETSO koduna göre eşleştirilir; GES dosyasında bulunmayan ETSO'larda 0 gösterilir.",
                )
                ges_updated = True
            elif text.startswith("Tam karşılaştırmayı çalıştırdıktan sonra sonuç bölümündeki"):
                replace_paragraph(
                    paragraph,
                    "Tam karşılaştırmayı çalıştırdıktan sonra sonuç bölümündeki ETSO Bazlı Tüketim Karşılaştırmasını İndir düğmesine basarak tek Excel raporunu alın. Raporda tüm ETSO kayıtları ve müşteri adlarıyla birlikte bizim kestiğimiz, dağıtımın ilettiği, uzlaştırma tüketimi ve GES üretimi bulunur. GES sütunu, dağıtım ile uzlaştırma farkının GES üretimiyle ilişkili olup olmadığını ETSO bazında incelemenizi sağlar.",
                )
                comparison_updated = True
        if not ges_updated or not comparison_updated:
            raise RuntimeError("Expected GES guide paragraphs were not found")

        document_xml = ET.tostring(
            root,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )
        with ZipFile(OUTPUT, "w") as output_zip:
            for item in source_zip.infolist():
                data = document_xml if item.filename == DOCUMENT_XML else source_zip.read(item.filename)
                output_zip.writestr(item, data)

    print(f"Updated DOCX: {OUTPUT}")
    print(f"Bytes: {OUTPUT.stat().st_size}")


if __name__ == "__main__":
    main()
