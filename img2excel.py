#!/usr/bin/env python3
import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from statistics import median

from ocrmac import ocrmac
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image

try:
    import cv2
    import numpy as np
except ImportError:
    cv2 = None

try:
    import pymupdf as fitz
except ImportError:
    fitz = None

try:
    from pillow_heif import register_heif_opener

    register_heif_opener()
except ImportError:
    pass


@dataclass
class Line:
    text: str
    confidence: float
    x: float
    right: float
    top: float
    bottom: float

    @property
    def center(self):
        return (self.top + self.bottom) / 2


@dataclass
class Cell:
    text: str
    x: float


@dataclass
class Row:
    anchor: float
    lines: list = field(default_factory=list)


def normalize_text(text):
    for char in "\u2018\u2019\u02bb\u02bc`\u2032":
        text = text.replace(char, "'")
    text = re.sub(r"\s+", " ", text).strip()
    return text.replace("N\u00ba", "\u2116")


def ocr_lines(image, framework, languages, min_confidence):
    kwargs = dict(
        framework=framework,
        recognition_level="accurate",
        language_preference=list(languages),
        confidence_threshold=min_confidence,
        detail=True,
    )
    if framework == "livetext":
        kwargs["unit"] = "line"
    try:
        annotations = ocrmac.OCR(image, **kwargs).recognize()
    except Exception as exc:
        if framework != "livetext":
            raise
        print(f"Live Text ishlamadi ({exc}), Vision ishlatiladi", file=sys.stderr)
        kwargs.pop("unit")
        kwargs["framework"] = "vision"
        annotations = ocrmac.OCR(image, **kwargs).recognize()

    lines = []
    for text, confidence, bbox in annotations:
        text = normalize_text(text)
        if not text:
            continue
        x, y, w, h = bbox
        lines.append(
            Line(
                text=text,
                confidence=float(confidence),
                x=x,
                right=x + w,
                top=1 - (y + h),
                bottom=1 - y,
            )
        )
    return lines


def group_rows(lines):
    heights = [line.bottom - line.top for line in lines if line.bottom > line.top]
    tolerance = 0.55 * median(heights) if heights else 0.01
    rows = []
    for line in sorted(lines, key=lambda item: (item.center, item.x)):
        target = next(
            (row for row in reversed(rows) if abs(line.center - row.anchor) <= tolerance),
            None,
        )
        if target is None:
            rows.append(Row(anchor=line.center, lines=[line]))
        else:
            target.lines.append(line)
    rows.sort(key=lambda row: row.anchor)
    return rows


def row_cells(row, gap_threshold):
    lines = sorted(row.lines, key=lambda item: item.x)
    cells = [[lines[0]]]
    for line in lines[1:]:
        if line.x - cells[-1][-1].right > gap_threshold:
            cells.append([line])
        else:
            cells[-1].append(line)
    return [
        Cell(text=" ".join(item.text for item in cell), x=cell[0].x)
        for cell in cells
    ]


def to_table(lines):
    heights = [line.bottom - line.top for line in lines if line.bottom > line.top]
    gap_threshold = max(0.012, 1.4 * median(heights)) if heights else 0.02
    return [row_cells(row, gap_threshold) for row in group_rows(lines)]


def detect_header(rows):
    for index in range(len(rows) - 1):
        current = [cell.text for cell in rows[index]]
        following = [cell.text for cell in rows[index + 1]]
        if (
            len(current) >= 2
            and len(current) == len(following)
            and all(len(text) <= 30 for text in current)
            and any(not text.replace(" ", "").isdigit() for text in current)
        ):
            return index
    return None


def align_table(rows, header_index):
    if header_index is None or len(rows[header_index]) < 2:
        return [[cell.text for cell in row] for row in rows]

    header = rows[header_index]
    column_starts = [cell.x for cell in header]
    aligned = []
    for index, row in enumerate(rows):
        if index <= header_index:
            aligned.append([cell.text for cell in row])
            continue
        slots = [[] for _ in column_starts]
        for cell in row:
            column = min(
                range(len(column_starts)),
                key=lambda i: abs(cell.x - column_starts[i]),
            )
            slots[column].append(cell.text)
        aligned.append([" ".join(slot) for slot in slots])
    return aligned


def natural_key(value):
    return [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", value)
    ]


def sort_rows(rows, column_index, descending):
    filled = []
    empty = []
    for row in rows:
        value = row[column_index].strip() if column_index < len(row) else ""
        (filled if value else empty).append((value, row))
    filled.sort(key=lambda item: natural_key(item[0]), reverse=descending)
    return [row for _, row in filled] + [row for _, row in empty]


def rotation_matrix(width, height, angle, expand):
    center = (width / 2, height / 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    if not expand:
        return matrix, (width, height)
    cos, sin = abs(matrix[0, 0]), abs(matrix[0, 1])
    new_width = int(height * sin + width * cos)
    new_height = int(height * cos + width * sin)
    matrix[0, 2] += new_width / 2 - center[0]
    matrix[1, 2] += new_height / 2 - center[1]
    return matrix, (new_width, new_height)


def detect_skew(image, max_angle=8.0, step=0.5):
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2GRAY)
    if gray.shape[1] > 1200:
        scale = 1200 / gray.shape[1]
        gray = cv2.resize(
            gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA
        )
    _, binary = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU
    )
    height, width = binary.shape
    best_angle, best_score = 0.0, -1.0
    for angle in np.arange(-max_angle, max_angle + 0.01, step):
        matrix, size = rotation_matrix(width, height, angle, expand=False)
        rotated = cv2.warpAffine(
            binary, matrix, size, flags=cv2.INTER_LINEAR, borderValue=0
        )
        score = float(np.var(rotated.sum(axis=1, dtype=np.float64)))
        if score > best_score:
            best_angle, best_score = float(angle), score
    return best_angle


def rotate_image(image, angle):
    array = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    height, width = array.shape[:2]
    matrix, size = rotation_matrix(width, height, angle, expand=True)
    rotated = cv2.warpAffine(
        array,
        matrix,
        size,
        flags=cv2.INTER_CUBIC,
        borderValue=(255, 255, 255),
    )
    return Image.fromarray(cv2.cvtColor(rotated, cv2.COLOR_BGR2RGB))


def load_sources(paths, dpi):
    for raw_path in paths:
        path = Path(raw_path)
        if not path.exists():
            raise SystemExit(f"Fayl topilmadi: {path}")
        if path.suffix.lower() == ".pdf":
            if fitz is None:
                raise SystemExit("PDF uchun pymupdf kerak: pip install pymupdf")
            document = fitz.open(path)
            for number, page in enumerate(document, start=1):
                pixmap = page.get_pixmap(dpi=dpi)
                image = Image.frombytes(
                    "RGB", (pixmap.width, pixmap.height), pixmap.samples
                )
                yield f"{path.name} ({number}-bet)", image
        else:
            with Image.open(path) as image:
                yield path.name, image.convert("RGB")


def fold(text):
    text = re.sub(r"[^0-9a-z\u0430-\u044f\u0451\u0400-\u04ff]", "", text.casefold())
    return text.replace("l", "i").replace("1", "i")


def resolve_sort_index(value, header):
    if re.fullmatch(r"\d+", value):
        index = int(value) - 1
        if index < 0:
            raise SystemExit("Ustun raqami 1 dan boshlanadi")
        return index
    if header:
        query = fold(value)
        for index, cell in enumerate(header):
            if value.casefold() in cell.casefold() or (
                query and query in fold(cell)
            ):
                return index
    raise SystemExit(f"Ustun topilmadi: {value}")


def write_sheet(sheet, rows, is_header_row):
    for row in rows:
        sheet.append(row)
    max_columns = max((len(row) for row in rows), default=0)
    if is_header_row and rows:
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DCE6F1")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for column in range(1, max_columns + 1):
        width = max(
            (len(str(row[column - 1])) for row in rows if column - 1 < len(row)),
            default=8,
        )
        sheet.column_dimensions[get_column_letter(column)].width = min(
            max(10, width + 2), 60
        )
    sheet.freeze_panes = "A2" if is_header_row else "A1"
    if rows:
        sheet.auto_filter.ref = sheet.dimensions


def main():
    parser = argparse.ArgumentParser(
        description="Qog'oz rasm/PDF/skanerdagi matnni Excelga aylantiradi"
    )
    parser.add_argument("sources", nargs="+", help="Rasm, PDF yoki papka yo'li")
    parser.add_argument("-o", "--output", help="Natija .xlsx fayli")
    parser.add_argument(
        "--framework",
        choices=["vision", "livetext"],
        default="vision",
        help="vision barqaror, livetext qo'lyozmani yaxshiroq o'qishi mumkin",
    )
    parser.add_argument(
        "--lang",
        nargs="+",
        default=["en-US", "ru-RU"],
        help="Til ustuvorligi, masalan: --lang en-US ru-RU",
    )
    parser.add_argument(
        "--sort-by", help="Saralash ustuni: raqam (1 dan) yoki sarlavha nomi"
    )
    parser.add_argument("--desc", action="store_true", help="Teskari tartibda saralash")
    parser.add_argument(
        "--header-row",
        type=int,
        default=-1,
        help="-1 avtomatik, 0 sarlavhasiz, N - N-qator sarlavha (1 dan)",
    )
    parser.add_argument("--min-confidence", type=float, default=0.0)
    parser.add_argument("--dpi", type=int, default=300, help="PDF uchun render DPI")
    parser.add_argument(
        "--no-deskew", action="store_true", help="Rasmni avtomatik to'g'rilamaslik"
    )
    args = parser.parse_args()

    sources = []
    for raw_path in args.sources:
        path = Path(raw_path)
        if path.is_dir():
            sources.extend(
                str(item)
                for item in sorted(path.iterdir())
                if item.suffix.lower() in {".png", ".jpg", ".jpeg", ".heic", ".pdf"}
            )
        else:
            sources.append(str(path))
    if not sources:
        raise SystemExit("Rasm yoki PDF topilmadi")

    raw_lines = []
    prepared = []
    for label, image in load_sources(sources, args.dpi):
        if not args.no_deskew and cv2 is not None:
            angle = detect_skew(image)
            if abs(angle) >= 0.3:
                image = rotate_image(image, angle)
                print(
                    f"{label}: {angle:+.1f}\u00b0 to'g'rilandi", file=sys.stderr
                )
        lines = ocr_lines(image, args.framework, args.lang, args.min_confidence)
        if not lines:
            print(f"Matn topilmadi: {label}", file=sys.stderr)
            continue
        for order, line in enumerate(
            sorted(lines, key=lambda item: (item.top, item.x)), start=1
        ):
            raw_lines.append(
                [
                    label,
                    order,
                    line.text,
                    round(line.confidence, 3),
                    round(line.x, 4),
                    round(line.top, 4),
                ]
            )
        cells = to_table(lines)
        if args.header_row == -1:
            header_index = detect_header(cells)
        elif args.header_row == 0:
            header_index = None
        else:
            header_index = args.header_row - 1
            if header_index >= len(cells):
                raise SystemExit(f"{label}: {args.header_row}-qator yo'q")
        prepared.append((label, align_table(cells, header_index), header_index))

    if not prepared:
        raise SystemExit("Hech qanday matn topilmadi")

    merged = len(prepared) > 1
    first_table = prepared[0][1]
    first_header_index = prepared[0][2]
    header = (
        first_table[first_header_index] if first_header_index is not None else None
    )
    if merged and header is not None:
        header = ["Manba"] + header

    rows = []
    for label, table, header_index in prepared:
        data = table[header_index + 1 :] if header_index is not None else table
        rows.extend(([label] + row) if merged else row for row in data)

    combined_header = header

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Jadval"
    write_sheet(
        sheet,
        ([combined_header] + rows) if combined_header is not None else rows,
        combined_header is not None,
    )

    if args.sort_by:
        if combined_header is not None:
            column_index = resolve_sort_index(args.sort_by, combined_header)
            sheet_header = combined_header
        elif args.sort_by.isdigit():
            column_index = int(args.sort_by) - 1
            sheet_header = None
        else:
            raise SystemExit(
                "Sarlavha topilmadi; ustun raqamini bering yoki --header-row N"
            )
        sorted_sheet = workbook.create_sheet("Saralangan")
        write_sheet(
            sorted_sheet,
            ([sheet_header] if sheet_header is not None else [])
            + sort_rows(rows, column_index, args.desc),
            sheet_header is not None,
        )

    raw_header = ["Manba", "Qator", "Matn", "Ishonch", "X", "Y"]
    raw_sheet = workbook.create_sheet("Matn")
    write_sheet(raw_sheet, [raw_header] + raw_lines, True)

    output = Path(args.output) if args.output else Path(sources[0]).with_suffix(".xlsx")
    workbook.save(output)
    print(f"Tayyor: {output}")


if __name__ == "__main__":
    main()
