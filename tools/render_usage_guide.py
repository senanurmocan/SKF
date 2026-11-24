from __future__ import annotations

import html
import re
import shutil
import subprocess
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
GUIDE = ROOT / "KULLANIM_KILAVUZU.md"
IMAGE = ROOT / "KULLANIM_KILAVUZU_GORSELLERI" / "epias_sayac_kaydi.png"
OUTPUTS = (
    ROOT / "KULLANIM_KILAVUZU_ICMAL_GUNCEL_DUZENLENMIS.pdf",
    ROOT / "KULLANIM_KILAVUZU_ICMAL_GUNCEL_DUZENLENMIS_KUCULTULMUS.pdf",
)
SCRATCH = ROOT / ".guide-render"
EDGE_CANDIDATES = (
    Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
)


CSS = r"""
@page {
  size: A4;
  margin: 19mm 17mm 20mm;
  @top-left {
    content: "CK ENERJİ";
    color: #163f75;
    font: 700 8pt Arial, sans-serif;
    vertical-align: bottom;
    border-bottom: 2pt solid #df7f2d;
    padding-bottom: 3mm;
  }
  @top-right {
    content: "YEREL WINDOWS UYGULAMASI  ·  v1.3.2";
    color: #163f75;
    font: 700 8pt Arial, sans-serif;
    text-align: right;
    vertical-align: bottom;
    border-bottom: 2pt solid #df7f2d;
    padding-bottom: 3mm;
  }
  @bottom-left {
    content: "CK Enerji  ·  SKF Yerel Uygulama Kullanım Kılavuzu";
    color: #708090;
    font: 8pt Arial, sans-serif;
    vertical-align: top;
    border-top: 0.6pt solid #d6dee7;
    padding-top: 3mm;
  }
  @bottom-right {
    content: "Sayfa " counter(page) " / " counter(pages);
    color: #708090;
    font: 8pt Arial, sans-serif;
    text-align: right;
    vertical-align: top;
    border-top: 0.6pt solid #d6dee7;
    padding-top: 3mm;
  }
}

* { box-sizing: border-box; }
html { font-family: Arial, "Segoe UI", sans-serif; color: #263442; }
body { margin: 0; font-size: 9.4pt; line-height: 1.32; }
h1, h2, h3, p, ul, ol, blockquote, figure { margin-top: 0; }
h1 { color: #163f75; font-size: 24pt; line-height: 1.08; margin: 0 0 5pt; }
h1 + p { color: #596c7c; font-size: 11pt; margin: 0 0 13pt; }
h2 {
  color: #163f75; font-size: 16pt; line-height: 1.12;
  margin: 15pt 0 6pt; padding-bottom: 4pt;
  border-bottom: 1.5pt solid #d9e2ec;
  break-after: avoid-page; page-break-after: avoid;
}
h2::after { content: ""; display: block; width: 18mm; border-bottom: 2.3pt solid #df7f2d; margin-bottom: -5.2pt; }
h3 { color: #c96f25; font-size: 11.5pt; line-height: 1.18; margin: 10pt 0 4pt; break-after: avoid-page; page-break-after: avoid; }
p { margin-bottom: 6pt; orphans: 2; widows: 2; }
ul, ol { margin: 3pt 0 7pt; padding-left: 19pt; }
li { margin: 0 0 3.2pt; padding-left: 2pt; break-inside: avoid; page-break-inside: avoid; }
ol li::marker { color: #d9782a; font-weight: 700; }
ul li::marker { color: #218e91; }
strong { color: #243c54; }
code { font-family: Consolas, monospace; font-size: 0.92em; background: #f1f4f7; padding: 1pt 3pt; border-radius: 2pt; }
a { color: #145ea8; text-decoration: none; overflow-wrap: anywhere; }
blockquote {
  margin: 8pt 0 10pt; padding: 8pt 10pt;
  border-left: 3pt solid #218e91; border-radius: 2pt;
  background: #f1f8f8; color: #24485a;
  break-inside: avoid-page; page-break-inside: avoid;
}
.zip-flow {
  display: grid; grid-template-columns: 1fr 12pt 1fr 12pt 1.25fr 12pt 1fr;
  align-items: stretch; gap: 4pt; margin: 8pt 0 10pt;
  padding: 7pt; border: 0.7pt solid #d9e2ec; border-left: 3pt solid #df7f2d;
  border-radius: 5pt; background: #f7f9fb;
  break-inside: avoid-page; page-break-inside: avoid;
}
.zip-flow-step {
  min-height: 48pt; padding: 6pt; border: 0.6pt solid #d8e2ec;
  border-radius: 4pt; background: #fff;
}
.zip-flow-step strong { display: block; color: #163f75; font-size: 8.5pt; line-height: 1.2; margin: 2pt 0; }
.zip-flow-step small { display: block; color: #526373; font-size: 7.4pt; line-height: 1.18; }
.zip-flow-number { display: inline-block; width: 14pt; height: 14pt; line-height: 14pt; text-align: center; border-radius: 50%; background: #df7f2d; color: #fff; font-size: 7.4pt; font-weight: 700; }
.zip-flow-arrow { color: #218e91; text-align: center; align-self: center; font-size: 14pt; font-weight: 700; }
figure { margin: 7pt 0 10pt; text-align: left; break-inside: avoid-page; page-break-inside: avoid; }
figure img { display: block; width: 118mm; max-width: 100%; max-height: 62mm; object-fit: contain; object-position: left top; border: 0.6pt solid #d8e2ec; border-radius: 4pt; padding: 3pt; }
hr { border: 0; border-top: 1pt solid #d8e2ec; margin: 9pt 0; }
@media print { img { print-color-adjust: exact; -webkit-print-color-adjust: exact; } }
"""


def inline(text: str) -> str:
    token = re.compile(r"(\[[^\]]+\]\([^)]+\)|\*\*.+?\*\*|`[^`]+`)")
    result: list[str] = []
    for part in token.split(text):
        if not part:
            continue
        link = re.fullmatch(r"\[([^\]]+)\]\(([^)]+)\)", part)
        if link:
            label, target = link.groups()
            result.append(
                f'<a href="{html.escape(target, quote=True)}">{inline(label)}</a>'
            )
        elif part.startswith("**") and part.endswith("**"):
            result.append(f"<strong>{inline(part[2:-2])}</strong>")
        elif part.startswith("`") and part.endswith("`"):
            result.append(f"<code>{html.escape(part[1:-1])}</code>")
        else:
            result.append(html.escape(part))
    return "".join(result)


def render_markdown(source: str, image_path: Path) -> str:
    lines = source.splitlines()
    output: list[str] = []
    paragraph: list[str] = []
    list_kind: str | None = None

    def close_paragraph() -> None:
        if paragraph:
            output.append(f"<p>{inline(' '.join(line.strip() for line in paragraph))}</p>")
            paragraph.clear()

    def close_list() -> None:
        nonlocal list_kind
        if list_kind:
            output.append(f"</{list_kind}>")
            list_kind = None

    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.strip()
        if not line:
            close_paragraph()
            close_list()
            i += 1
            continue

        if line.startswith('<div class="zip-flow"'):
            close_paragraph()
            close_list()
            block = [line]
            i += 1
            while i < len(lines) and "</div>" not in lines[i]:
                block.append(lines[i].strip())
                i += 1
            if i < len(lines):
                block.append(lines[i].strip())
                i += 1
            output.append("\n".join(block))
            continue

        if line.startswith("<img "):
            close_paragraph()
            close_list()
            image_uri = image_path.resolve().as_uri()
            tag = re.sub(r'src="[^"]+"', f'src="{image_uri}"', line, count=1)
            output.append(f"<figure>{tag}</figure>")
            i += 1
            continue

        heading = re.match(r"^(#{1,3})\s+(.+)$", line)
        if heading:
            close_paragraph()
            close_list()
            level = len(heading.group(1))
            output.append(f"<h{level}>{inline(heading.group(2))}</h{level}>")
            i += 1
            continue

        if line.startswith("> "):
            close_paragraph()
            close_list()
            quoted: list[str] = []
            while i < len(lines) and lines[i].strip().startswith("> "):
                quoted.append(lines[i].strip()[2:])
                i += 1
            output.append(f"<blockquote>{inline(' '.join(quoted))}</blockquote>")
            continue

        ordered = re.match(r"^\d+\.\s+(.+)$", line)
        unordered = re.match(r"^-\s+(.+)$", line)
        if ordered or unordered:
            close_paragraph()
            wanted = "ol" if ordered else "ul"
            if list_kind != wanted:
                close_list()
                list_kind = wanted
                output.append(f"<{list_kind}>")
            value = ordered.group(1) if ordered else unordered.group(1)
            output.append(f"<li>{inline(value)}</li>")
            i += 1
            continue

        close_list()
        if line.startswith("<"):
            close_paragraph()
            output.append(line)
            i += 1
            continue
        paragraph.append(line)
        i += 1

    close_paragraph()
    close_list()
    return "\n".join(output)


def make_html(markdown: str, image_path: Path) -> str:
    content = render_markdown(markdown, image_path)
    return f"""<!doctype html>
<html lang="tr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>{CSS}</style>
</head>
<body>{content}</body>
</html>"""


def find_edge() -> Path:
    for candidate in EDGE_CANDIDATES:
        if candidate.is_file():
            return candidate
    raise FileNotFoundError("Microsoft Edge bulunamadı; kılavuz PDF'leri üretilmedi.")


def print_pdf(edge: Path, html_path: Path, pdf_path: Path, profile: Path) -> None:
    if pdf_path.exists():
        pdf_path.unlink()
    command = [
        str(edge),
        "--headless=new",
        "--disable-gpu",
        "--disable-extensions",
        "--no-first-run",
        "--no-pdf-header-footer",
        "--allow-file-access-from-files",
        f"--user-data-dir={profile}",
        f"--print-to-pdf={pdf_path}",
        html_path.resolve().as_uri(),
    ]
    subprocess.run(command, check=True, timeout=90, capture_output=True)
    if not pdf_path.is_file() or pdf_path.stat().st_size < 20_000:
        raise RuntimeError(f"PDF oluşturulamadı: {pdf_path}")


def main() -> None:
    markdown = GUIDE.read_text(encoding="utf-8")
    edge = find_edge()
    SCRATCH.mkdir(exist_ok=True)

    compact_image = SCRATCH / "epias_sayac_kaydi_compact.jpg"
    with Image.open(IMAGE) as source:
        image = source.convert("RGB")
        image.thumbnail((720, 720), Image.Resampling.LANCZOS)
        image.save(compact_image, "JPEG", quality=58, optimize=True)

    jobs = []
    for index, image_path in enumerate((IMAGE, compact_image)):
        kind = "full" if index == 0 else "compact"
        html_path = SCRATCH / f"usage-guide-{kind}.html"
        temporary_pdf = SCRATCH / f"usage-guide-{kind}.pdf"
        profile = SCRATCH / f"edge-profile-{kind}"
        html_path.write_text(make_html(markdown, image_path), encoding="utf-8")
        print_pdf(edge, html_path, temporary_pdf, profile)
        jobs.append((temporary_pdf, OUTPUTS[index]))

    for temporary_pdf, output in jobs:
        shutil.copy2(temporary_pdf, output)
        print(f"{output.name}: {output.stat().st_size:,} bytes")


if __name__ == "__main__":
    main()
