#!/usr/bin/env python3
import glob

import xlrd
import xlwt
from xlutils.copy import copy

SOURCE = glob.glob(
    "/Users/probox/Documents/qogoz2excel/incoming/beka-2026-10-22/3. *гиёхванлик.xls"
)[0]
OUTPUT = (
    "/Users/probox/Documents/qogoz2excel/incoming/beka-2026-10-22/"
    "3. Сўровнома 2026 гиёхванлик (1-respondent, to'ldirilgan).xls"
)

RESPONDENT_ROW = 5
RESPONDENT_NUMBER = 1

DATA = {
    1: 1,
    13: 1,
    19: 1,
    28: 1,
    45: 1,
    50: 1,
    52: 1,
    69: 1,
    77: 1,
    81: 1,
    84: 1,
    89: 1,
    100: 1,
    106: 1,
    113: "Alisher Navoiy",
    118: 1,
    133: 1,
    134: 1,
}

NOTES = [
    ("Respondent", "1 (Excel 6-qator)"),
    ("Manba", "Beka yuborgan bo'sh shablon + 5 ta rasm (22.10.2026)"),
    ("I.1", "Aniq belgilangan"),
    ("II.3", "Aniq belgilangan"),
    ("III.1", "Aniq belgilangan"),
    ("IV.1", "Aniq belgilangan"),
    ("V.8", "«Bilmayman» — 45-ustun"),
    ("VI.5", "Aniq belgilangan"),
    ("VII.1", "Aniq belgilangan"),
    ("VIII.9", "«Bilmayman» — 69-ustun"),
    ("IX.9", "«Bilmayman» — 77-ustun"),
    ("X.4", "Aniq belgilangan"),
    ("XI.1", "Aniq belgilangan"),
    ("XII.1", "Aniq belgilangan"),
    ("XIII.1", "Aniq belgilangan"),
    ("XIV.1", "Erkak"),
    ("XV", "«14 ёшда» — mos variant yo'q, 108–112 bo'sh qoldirildi"),
    ("XVI", "«Alisher Navoiy» — 113-ustun"),
    ("XVII.5", "«O'quvchi» — 118-ustun"),
    ("XVIII.15", "«O'quvchi» — 133-ustun"),
    ("XIX.1", "Aniq belgilangan"),
    ("XIX.5", "Kichik belgi bor, tasdiqlanmagan — 138-ustun bo'sh"),
    ("Sarlavha", "«Ravshan S. 9-sinf» taxminiy"),
]


def main():
    book = xlrd.open_workbook(SOURCE, formatting_info=True)
    workbook = copy(book)
    sheet = workbook.get_sheet(0)
    centered = xlwt.easyxf(
        "align: horiz center, vert center;"
        "border: left thin, right thin, top thin, bottom thin;"
    )
    left = xlwt.easyxf(
        "align: horiz left, vert center;"
        "border: left thin, right thin, top thin, bottom thin;"
    )

    sheet.write(RESPONDENT_ROW, 0, RESPONDENT_NUMBER, centered)
    for column, value in DATA.items():
        style = centered if isinstance(value, int) else left
        sheet.write(RESPONDENT_ROW, column, value, style)

    notes = workbook.add_sheet("Izohlar")
    notes.write(0, 0, "Maydon", xlwt.easyxf("font: bold on;"))
    notes.write(0, 1, "Izoh", xlwt.easyxf("font: bold on;"))
    notes.col(0).width = 6000
    notes.col(1).width = 18000
    for row, (field, text) in enumerate(NOTES, start=1):
        notes.write(row, 0, field)
        notes.write(row, 1, text)

    workbook.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
