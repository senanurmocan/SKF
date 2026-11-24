from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "KULLANIM_KILAVUZU_DUZENLENEBILIR.docx"
OUTPUT = ROOT / "KULLANIM_KILAVUZU_DUZENLENEBILIR_GUNCEL.docx"
DOCUMENT_XML = "word/document.xml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
ET.register_namespace("w", W)


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def make_paragraph(text: str, style: str = "Normal") -> ET.Element:
    paragraph = ET.Element(f"{{{W}}}p")
    properties = ET.SubElement(paragraph, f"{{{W}}}pPr")
    paragraph_style = ET.SubElement(properties, f"{{{W}}}pStyle")
    paragraph_style.set(f"{{{W}}}val", style)
    if style.startswith("Heading"):
        ET.SubElement(properties, f"{{{W}}}keepNext")
        ET.SubElement(properties, f"{{{W}}}keepLines")
    run = ET.SubElement(paragraph, f"{{{W}}}r")
    node = ET.SubElement(run, f"{{{W}}}t")
    node.text = text
    return paragraph


def replace_paragraph(paragraph: ET.Element, text: str) -> None:
    properties = paragraph.find("w:pPr", NS)
    for child in list(paragraph):
        if child is not properties:
            paragraph.remove(child)
    run = ET.SubElement(paragraph, f"{{{W}}}r")
    node = ET.SubElement(run, f"{{{W}}}t")
    node.text = text


def insert_after(body: ET.Element, anchor: str, paragraphs: list[ET.Element]) -> None:
    for index, child in enumerate(list(body)):
        if child.tag == f"{{{W}}}p" and anchor in paragraph_text(child):
            for offset, paragraph in enumerate(paragraphs, start=1):
                body.insert(index + offset, paragraph)
            return
    raise RuntimeError(f"Guide insertion point not found: {anchor}")


def main() -> None:
    if OUTPUT.exists():
        raise SystemExit(f"Output already exists; refusing to overwrite: {OUTPUT}")

    with ZipFile(SOURCE, "r") as source_zip:
        root = ET.fromstring(source_zip.read(DOCUMENT_XML))
        body = root.find("w:body", NS)
        if body is None:
            raise RuntimeError("Word document body not found")

        changed_folder_note = False
        changed_appendix_note = False
        for paragraph in body.findall("w:p", NS):
            text = paragraph_text(paragraph)
            if "Kaydederken bu isimlerden biri" in text:
                replace_paragraph(
                    paragraph,
                    "Şirket klasörü adını aşağıdaki şirket adlarından veya tanımlı ad varyasyonlarından biriyle yazın. Büyük/küçük harf biçimi önemli değildir: AKEDAŞ ve akedaş aynı bölge olarak algılanır. Her şirket için tek klasör kullanın.",
                )
                changed_folder_note = True
            elif "Dağıtım klasörlerini adlandırırken aşağıdaki kalın şirket adlarını" in text:
                replace_paragraph(
                    paragraph,
                    "Dağıtım klasörlerini aşağıdaki kalın şirket adlarıyla veya aynı satırda parantez içinde verilen tanımlı ad varyasyonlarıyla adlandırın. Uygulama adları büyük/küçük harfe duyarsız eşleştirir; örneğin AKEDAŞ ve akedaş ikisi de tanınır. Her şirket için yalnızca bir klasör açın ve bütün şirket klasörlerini seçilecek tek ana klasörün içinde tutun.",
                )
                changed_appendix_note = True

        if not changed_folder_note:
            raise RuntimeError("Main folder naming paragraph was not found")
        if not changed_appendix_note:
            raise RuntimeError("Company name appendix paragraph was not found")

        insert_after(
            body,
            "Uzlaştırma Tüketim:",
            [
                make_paragraph(
                    "GES Üretimi (isteğe bağlı): Tedarikçi LUM bildirimi biçimindeki .xlsx dosyasını GES Üretimi düğmesiyle seçin. Uygulama üretim miktarını ETSO koduna göre eşleştirir ve tüketim karşılaştırmasına ekler."
                )
            ],
        )

        insert_after(
            body,
            "Yalnızca Bizim Kestiğimiz Faturalar kutusunu işaretleyin.",
            [
                make_paragraph("Yalnızca Dağıtım SKF'sini Hazırlama", "Heading2"),
                make_paragraph(
                    "Dağıtım ana klasörünü seçin, Yalnızca Dağıtım SKF'si kutusunu işaretleyin ve Yalnızca Dağıtım SKF'sini Hazırla düğmesine basın. Bu işlemde fatura dosyası gerekmez."
                ),
                make_paragraph(
                    "Analiz tamamlanınca Birleştirilmiş Dağıtım SKF Dosyasını İndir düğmesi görünür. Bu dosya, okunan dağıtım bölgelerindeki SKF kayıtlarının birleştirilmiş ve düzenlenmiş halidir. Ekrandaki bölge sayısı 21'den azsa uyarıdaki eksik bölgeleri kontrol edin."
                ),
            ],
        )

        insert_after(
            body,
            "Sonuç, kaynak dosyaların üzerine yazılmaz;",
            [
                make_paragraph("ETSO Bazlı Tüketim Karşılaştırması", "Heading2"),
                make_paragraph(
                    "Tam karşılaştırmayı çalıştırdıktan sonra sonuç bölümündeki ETSO Bazlı Tüketim Karşılaştırmasını İndir düğmesine basarak tek Excel raporunu alın. Raporda tüm ETSO kayıtları ve müşteri adlarıyla birlikte bizim kestiğimiz, dağıtımın ilettiği, uzlaştırma tüketimi ve GES üretimi bulunur."
                ),
                make_paragraph(
                    "Fark sütunu Excel formülüyle yalnızca Dağıtım eksi Uzlaştırma olarak hesaplanır. Farkı olmayan ETSO'lar da raporda kalır. Sayısal hücreler virgüllü biçimde ve düzenlenebilir olarak saklanır; GES verisi bulunmayan ETSO'larda 0 gösterilir. Renkler yalnızca fark sütununu belirtir."
                ),
                make_paragraph(
                    "SKF ve tüketim karşılaştırması çıktılarında Kesin ve Taslak kayıtları birlikte yer alır; durumlarına göre ayırma yapılmaz. Tam karşılaştırmada Diğer Bedel Tablosu çıktısı SKF önizlemesinin altında, Ek Kalem Bedel Farklarını İndir düğmesiyle alınır."
                ),
            ],
        )

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
