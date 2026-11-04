#!/usr/bin/env python3
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path("/Users/probox/Documents/qogoz2excel/incoming/beka-2026-10-22")
OUT = BASE / "Beka_sorovnoma.xlsx"
OCR = Path("/tmp/beka_ocr.json")

QUESTIONS = [
    {
        "no": "I",
        "q": "«Соғлом турмуш тарзи» иборасини қандай тасаввур қиласиз? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": ["1"],
        "marked_text": "Жисмоний тарбия ва спорт билан шуғулланиш",
        "hand": "",
        "note": "",
        "page": "1-sahifa (I–III)",
        "status": "Aniq",
    },
    {
        "no": "II",
        "q": "Нима деб ўйлайсиз, «соғлом турмуш тарзи» жамиятга умуман керакми? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": ["3"],
        "marked_text": "Ҳаётимиз ва турмуш тарзимизда қийинчилик ва муаммолар жуда кўп, бундай шароитда соғлом турмуш тарзига амал қилиб яшашнинг деярли имкони йўқ",
        "hand": "",
        "note": "",
        "page": "1-sahifa (I–III)",
        "status": "Aniq",
    },
    {
        "no": "III",
        "q": "Сиз ўзингиз «соғлом турмуш тарзи»га амал қиласизми? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": ["1"],
        "marked_text": "Амал қиламан, ҳар куни спорт билан шуғулланаман ва соғлиғим учун зарарли таом ҳамда ичимликларни истеъмол қилмайман",
        "hand": "",
        "note": "",
        "page": "1-sahifa (I–III)",
        "status": "Aniq",
    },
    {
        "no": "IV",
        "q": "Сизнингча бугунги кунда аҳоли, хусусан ёшлар ўртасида гиёҳванд моддаларга рўжу қўйиш, психотроп таблеткалар қабул қилиш кўникмалари мавжудми? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": ["1"],
        "marked_text": "Йўқ, мавжуд эмас, чунки қонунчиликда бу иллатларга қарши жиддий курашилади",
        "hand": "",
        "note": "",
        "page": "2-sahifa (IV–VII)",
        "status": "Aniq",
    },
    {
        "no": "V",
        "q": "Нима деб ўйлайсиз, бугунги кунда айрим ёшлар ўртасида қайси турдаги гиёҳванд моддаларнинг истеъмоли мавжуд бўлиши мумкин? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": [],
        "marked_text": "",
        "hand": "Bilmayman",
        "note": "8-вариант «Бошқаси (ёзинг)» қаторига «Bilmayman» деб ёзилган.",
        "page": "2-sahifa (IV–VII)",
        "status": "Qo'lyozma (aniq)",
    },
    {
        "no": "VI",
        "q": "Айта оласизми бугунги кунда айрим ёшлар ўртасида қайси турдаги психотроп маҳсулотлар истеъмоли мавжуд? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": ["5"],
        "marked_text": "Психотроп таблеткалар тўғрисида эшитганман, аммо унинг қандай турлари мавжудлигини билмайман",
        "hand": "",
        "note": "",
        "page": "2-sahifa (IV–VII)",
        "status": "Aniq",
    },
    {
        "no": "VII",
        "q": "Айта оласизми, ўз ҳаётингиз мобайнида гиёҳванд моддалар қабул қилганмисиз? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": ["1"],
        "marked_text": "Йўқ, ҳеч қачон қабул қилмаганман, олган тарбиям йўл қўймайди",
        "hand": "",
        "note": "",
        "page": "3-sahifa (VIII–X)",
        "status": "Aniq",
    },
    {
        "no": "VIII",
        "q": "Гиёҳванд моддалар ва психотроп маҳсулотлар молиявий жиҳатдан қиммат туради, сизнингча ёшлар уларни қандай қўлга киритишлари мумкин? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": [],
        "marked_text": "",
        "hand": "Bilmayman",
        "note": "9-вариант «Бошқаси (ёзинг)» қаторига «Bilmayman» деб ёзилган.",
        "page": "3-sahifa (VIII–X)",
        "status": "Qo'lyozma (aniq)",
    },
    {
        "no": "IX",
        "q": "Нима деб ўйлайсиз, ёшлар қайси сабабларга кўра гиёҳванд моддалар ва психотроп маҳсулотлар қабул қиладилар? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": [],
        "marked_text": "",
        "hand": "Bilmayman",
        "note": "9-вариант «Бошқаси (ёзинг)» қаторига «Bilmayman» деб ёзилган.",
        "page": "3-sahifa (VIII–X)",
        "status": "Qo'lyozma (aniq)",
    },
    {
        "no": "X",
        "q": "Нима деб ўйлайсиз, гиёҳванд моддалар ва психотроп маҳсулотларга рўжу қўйишлик қайси ёшдан бошланиши мумкин?",
        "marked": ["4"],
        "marked_text": "У ёш талабалик, ёшгингиз руҳияти бутунлай ўзгариб кетган ҳисобига ҳам пайдо бўлиши мумкин",
        "hand": "",
        "note": "",
        "page": "4-sahifa (XI–XIII)",
        "status": "Aniq",
    },
    {
        "no": "XI",
        "q": "Нима деб ўйлайсиз гиёҳванд моддалар ва психотроп маҳсулотларга қарамлик қиз болалар ўртасида ҳам мавжудми? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": ["1"],
        "marked_text": "Аввалдан қиз болалар ўртасида бу иллатга мойиллик бўлмаган ва ҳозир ҳам мавжуд эмас",
        "hand": "",
        "note": "",
        "page": "4-sahifa (XI–XIII)",
        "status": "Aniq",
    },
    {
        "no": "XII",
        "q": "Нима деб ўйлайсиз, наҳотки гиёҳванд моддалар ёки психотроп маҳсулотлар истеъмоли инсонни ҳаёт чарчоқларидан халос этади? (учтагача жавоб вариантини белгилаш мумкин)",
        "marked": ["1"],
        "marked_text": "Йўқ, асло ундай эмас – бунинг оқибати фожиа билан тугаши аниқ",
        "hand": "",
        "note": "",
        "page": "4-sahifa (XI–XIII)",
        "status": "Aniq",
    },
    {
        "no": "XIII",
        "q": "Агар сиз гиёҳванд моддалар ва психотроп маҳсулотлар истеъмол қилмасангиз, уни қандай сабаблар билан изоҳлаган бўлардингиз?",
        "marked": ["1"],
        "marked_text": "Биринчи навбатда ота-онамга, кейин устозларимга миннатдорчилик билдирардим, ҳамма вақт ушбу иллатнинг яхшиликка олиб келмаслигини синғдирганлари учун",
        "hand": "",
        "note": "",
        "page": "4-sahifa (XI–XIII)",
        "status": "Aniq",
    },
    {
        "no": "XIV",
        "q": "Жинсингиз?",
        "marked": ["1"],
        "marked_text": "эркак",
        "hand": "",
        "note": "",
        "page": "5-sahifa (XIV–XIX)",
        "status": "Aniq",
    },
    {
        "no": "XV",
        "q": "Ёшингиз?",
        "marked": [],
        "marked_text": "",
        "hand": "14 ёшда",
        "note": "1-вариант (18–24) ёнига ёзилган; вариант айлана билан белгиланмаган.",
        "page": "5-sahifa (XIV–XIX)",
        "status": "Qo'lyozma (aniq)",
    },
    {
        "no": "XVI",
        "q": "Ҳудудингиз (Маҳалла номи)",
        "marked": [],
        "marked_text": "",
        "hand": "Alisher Navoiy",
        "note": "",
        "page": "5-sahifa (XIV–XIX)",
        "status": "Qo'lyozma (aniq)",
    },
    {
        "no": "XVII",
        "q": "Маълумотингиз?",
        "marked": [],
        "marked_text": "",
        "hand": "Ўқувчи",
        "note": "5-вариант «Бошқаси (ёзинг)» қаторига ёзилган.",
        "page": "5-sahifa (XIV–XIX)",
        "status": "Qo'lyozma (aniq)",
    },
    {
        "no": "XVIII",
        "q": "Фаолиятингиз?",
        "marked": [],
        "marked_text": "",
        "hand": "Ўқувчи",
        "note": "15-вариант «Бошқаси (ёзинг)» қаторига ёзилган.",
        "page": "5-sahifa (XIV–XIX)",
        "status": "Qo'lyozma (aniq)",
    },
    {
        "no": "XIX",
        "q": "Оилада кимлар билан истиқомат қиласиз?",
        "marked": ["1", "5?"],
        "marked_text": "1 – тўлиқ оилада (бобо, буви, ота, она, фарзанд); 5 – яқин қариндошлар билан (oxirida kichik belgi bor)",
        "mark_texts": {
            "1": "тўлиқ оилада (бобо, буви, ота, она, фарзанд)",
            "5?": "яқин қариндошлар билан",
        },
        "hand": "",
        "note": "1-вариант аниқ айлана қилинган. 5-вариант охиридаги кичик белги тасдиқланиши керак.",
        "page": "5-sahifa (XIV–XIX)",
        "status": "Tekshirish",
    },
]

PROFILE = [
    ("Жинс", "Эркак (XIV.1)", "Aniq"),
    ("Ёш", "14 ёшда", "Qo'lyozma (aniq)"),
    ("Ҳудуд / маҳалла", "Alisher Navoiy", "Qo'lyozma (aniq)"),
    ("Маълумот", "Ўқувчи (XVII.5)", "Qo'lyozma (aniq)"),
    ("Фаолият", "Ўқувчи (XVIII.15)", "Qo'lyozma (aniq)"),
    ("Оила", "Тўлиқ оилада (XIX.1); 5-вариантда ҳам белги бор", "Tekshirish"),
    ("Сарлавҳадаги ёзув", "Ravshan S. 9-sinf (taxminiy)", "Aniqlanmadi"),
]

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True)
MARK_FILL = PatternFill("solid", fgColor="C6EFCE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
THIN_WRAP = Alignment(vertical="top", wrap_text=True)


def style_header(ws, row=1):
    for cell in ws[row]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def widths(ws, values):
    for index, width in enumerate(values, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width


def main():
    workbook = Workbook()

    info = workbook.active
    info.title = "Ma'lumot"
    info.append(["Beka – qo'lda to'ldirilgan so'rovnoma"])
    info.append(["Manba", "Telegram: @bekmirzo_jumaboyev, 2026-10-22 17:21"])
    info.append(["Sahifalar", "5 ta rasm: I–III, IV–VII, VIII–X, XI–XIII, XIV–XIX"])
    info.append(["Ishlov", "Lokal OCR + qo'lda tekshiruv; internetga yuborilmadi"])
    info.append(
        [
            "Muhim",
            "«Tekshirish» deb belgilangan joylar va qo'lyozmalar rasmda qayta tekshirilishi kerak.",
        ]
    )
    info.append(["Natija fayli", str(OUT)])
    info.column_dimensions["A"].width = 22
    info.column_dimensions["B"].width = 95
    info["A1"].font = Font(bold=True, size=14)
    for row in info.iter_rows(min_row=1):
        for cell in row:
            cell.alignment = THIN_WRAP

    ws = workbook.create_sheet("Natija")
    ws.append(
        [
            "№",
            "Savol",
            "Belgilangan variant",
            "Belgilangan javob matni",
            "Qo'lyozma javob",
            "Izoh",
            "Sahifa",
            "Holat",
        ]
    )
    for item in QUESTIONS:
        ws.append(
            [
                item["no"],
                item["q"],
                ", ".join(item["marked"]),
                item["marked_text"],
                item["hand"],
                item["note"],
                item["page"],
                item["status"],
            ]
        )
        row = ws.max_row
        if item["marked"]:
            for column in (3, 4):
                ws.cell(row=row, column=column).fill = MARK_FILL
        if item["status"] == "Tekshirish":
            ws.cell(row=row, column=8).fill = WARN_FILL
        for cell in ws[row]:
            cell.alignment = THIN_WRAP
    style_header(ws)
    widths(ws, [8, 62, 16, 62, 26, 38, 24, 20])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    marks = workbook.create_sheet("Belgilar")
    marks.append(["Savol", "Variant", "Variant matni", "Belgi", "Sahifa", "Holat"])
    for item in QUESTIONS:
        for number in item["marked"]:
            text = item.get("mark_texts", {}).get(
                number,
                item["marked_text"] if number == item["marked"][0] else "",
            )
            marks.append(
                [
                    item["no"],
                    number,
                    text,
                    "Aylana" if number != "5?" else "Kichik belgi (tekshirish)",
                    item["page"],
                    item["status"],
                ]
            )
    style_header(marks)
    widths(marks, [10, 12, 75, 26, 24, 20])
    marks.freeze_panes = "A2"

    profile = workbook.create_sheet("Profil")
    profile.append(["Ko'rsatkich", "Qiymat", "Holat"])
    for row in PROFILE:
        profile.append(row)
    style_header(profile)
    widths(profile, [26, 60, 22])
    profile.freeze_panes = "A2"

    text_sheet = workbook.create_sheet("To'liq matn (OCR)")
    text_sheet.append(["Sahifa", "Tartib", "Matn", "Ishonch", "Manba fayl"])
    data = json.loads(OCR.read_text())
    page_map = [
        ("1", "1-sahifa (I–III)", "photo_2026-10-22_12-21-37.jpg"),
        ("2", "2-sahifa (IV–VII)", "photo_2026-10-22_12-21-37 (2).jpg"),
        ("3", "3-sahifa (VIII–X)", "photo_2026-10-22_12-21-37 (1).jpg"),
        ("4", "4-sahifa (XI–XIII)", "photo_2026-10-22_12-21-37 (4).jpg"),
        ("5", "5-sahifa (XIV–XIX)", "photo_2026-10-22_12-21-37 (3).jpg"),
    ]
    for page_key, label, source in page_map:
        number = 0
        for line in data[page_key]["lines"]:
            if line["top"] < 0.12:
                continue
            if line["text"].strip() in {"End", "Enter", "Del", "Ins"}:
                continue
            number += 1
            text_sheet.append(
                [label, number, line["text"], line["conf"], source]
            )
    style_header(text_sheet)
    widths(text_sheet, [24, 8, 100, 10, 42])
    text_sheet.freeze_panes = "A2"
    text_sheet.auto_filter.ref = text_sheet.dimensions

    workbook.save(OUT)
    print(OUT)


if __name__ == "__main__":
    main()
