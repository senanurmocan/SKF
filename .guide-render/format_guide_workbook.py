from __future__ import annotations

import shutil
from copy import copy
from pathlib import Path

from openpyxl import load_workbook


root = Path(__file__).resolve().parents[1]
target = root / "SKF_Kullanim_Kilavuzu_Gorsel_Alanli.xlsx"
temporary = Path(__file__).with_name("SKF_Kullanim_Kilavuzu_Gorsel_Alanli.updated.xlsx")
workbook = load_workbook(target)

steps = [
    "1) Klavyede Windows + R tuşlarına basın; shell:desktop yazıp Enter'a basın.",
    "2) Masaüstünde boş alana sağ tıklayın > Yeni > Klasör; klasöre SKF-Yerel adını verin.",
    "3) İndirilenler'de SKF-Yerel-Paket-1.3.zip dosyasına sağ tıklayın > Tümünü ayıkla.",
    "4) Açılan pencerede Gözat'a basın.",
    "5) Masaüstü > SKF-Yerel klasörünü seçin > Klasörü Seç.",
    "6) Hedefin Masaüstü\\SKF-Yerel olduğunu kontrol edin > Ayıkla.",
    "7) Çıkarma bitince SKF-Yerel > SKF-Yerel-Paket-1.3 klasörünü açın.",
]

main_sheet = workbook["Adım Adım"]
main_sheet["B8"] = "\n".join(steps)
main_sheet["C8"] = (
    "ZIP dosyası İndirilenler klasöründe kalır. Uygulamayı Masaüstü\\SKF-Yerel "
    "içinden başlatın. Paket dosyalarını ayrı ayrı taşımayın veya silmeyin. "
    "Masaüstüne yazma izniniz yoksa yazma izniniz bulunan başka bir konum seçin."
)
main_sheet.row_dimensions[8].height = 132
for row_index, height in {
    6: 70,
    7: 80,
    9: 74,
    10: 65,
    11: 70,
    12: 74,
    13: 68,
    14: 60,
    15: 65,
    16: 86,
    17: 70,
    18: 65,
}.items():
    main_sheet.row_dimensions[row_index].height = height
for row in main_sheet.iter_rows(min_row=6, max_row=18, min_col=1, max_col=3):
    for cell in row:
        cell.alignment = copy(cell.alignment)
        cell.alignment = cell.alignment.copy(wrap_text=True, vertical="top")
        if cell.column in (2, 3):
            cell.font = copy(cell.font)
            cell.font = cell.font.copy(size=10)
main_sheet.sheet_view.showGridLines = False
main_sheet.print_area = "A1:D18"
main_sheet.print_title_rows = "5:5"
main_sheet.page_setup.orientation = "landscape"
main_sheet.page_setup.paperSize = main_sheet.PAPERSIZE_A4
main_sheet.page_setup.fitToWidth = 1
main_sheet.page_setup.fitToHeight = 2
main_sheet.sheet_properties.pageSetUpPr.fitToPage = True
main_sheet.page_margins.left = 0.25
main_sheet.page_margins.right = 0.25
main_sheet.page_margins.top = 0.35
main_sheet.page_margins.bottom = 0.35
main_sheet.print_options.horizontalCentered = True

for sheet_name, header_row, print_area, row_heights in (
    ("Sorun Giderme", "4:4", "A1:C10", {1: 38, 2: 28, 3: 8, 4: 30, 5: 68, 6: 62, 7: 60, 8: 60, 9: 64, 10: 60}),
    ("İndirme Bağlantıları", "1:1", "A1:B5", {1: 40, 2: 28, 3: 8, 4: 32, 5: 32}),
):
    sheet = workbook[sheet_name]
    sheet.sheet_view.showGridLines = False
    sheet.print_area = print_area
    sheet.print_title_rows = header_row
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.35
    sheet.page_margins.bottom = 0.35
    sheet.print_options.horizontalCentered = True
    for row_index, height in row_heights.items():
        sheet.row_dimensions[row_index].height = height

workbook.save(temporary)
check = load_workbook(temporary, read_only=False)
assert check["Adım Adım"]["B8"].value.startswith("1) Klavyede Windows + R")
assert "Tümünü ayıkla" in check["Adım Adım"]["B8"].value
assert "Masaüstü > SKF-Yerel" in check["Adım Adım"]["B8"].value
assert check["Adım Adım"].page_setup.fitToHeight == 2
assert len(check["Adım Adım"]._images) == len(main_sheet._images)
check.close()
workbook.close()
shutil.copy2(temporary, target)
print(f"Updated {target.name}")
