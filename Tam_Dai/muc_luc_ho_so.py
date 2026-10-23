# -*- coding: utf-8 -*-
"""
muc_luc_ho_so.py — Chuyển "MỤC LỤC HỒ SƠ" từ sổ Excel (.xlsx) sang Word (.docx)
in được ngay (tất cả trang nằm ngang, Times New Roman, số trang ở đầu trang).

MẪU EXCEL ĐẦU VÀO
-----------------
Một mẫu dùng chung cho hàng trăm phông lưu trữ. Tên sheet QUY ƯỚC (chữ viết tắt, đây là
cách nhận diện chính): TB (Tờ bìa) → VT (Bảng chữ cái viết tắt) → PL (Phụ lục tra cứu) →
ML (Bảng thống kê hồ sơ) → CTKT (Chứng từ kết thúc). Tệp cũ chưa đổi tên sheet vẫn được
nhận theo nội dung (60 dòng × 20 cột đầu), theo từ khoá trong tên sheet, rồi theo thứ tự
sheet; sheet "Biên mục" (nếu có) bị bỏ qua. Trong mỗi sheet không dùng địa chỉ ô cố định:

  1. Tờ bìa            — có ô "MỤC LỤC HỒ SƠ". Các dòng phía trên là tên cơ quan
                         (dòng 2 thường đỏ, đậm); các dòng phía dưới là tên phông và
                         các khoá/nhiệm kỳ; khối cuối gồm "Phông số: …", "Mục lục
                         số: …", "Số trang:" ở bên trái và "Thời hạn bảo quản" /
                         "Vĩnh viễn" ở bên phải; sau cùng là dòng địa danh – thời
                         gian ("Phú Ninh, 7/2026").
  2. Bảng chữ viết tắt — hàng tiêu đề có ô "CHỮ VIẾT TẮT" (SỐ TT | CHỮ VIẾT TẮT |
                         NGHĨA LÀ | GHI CHÚ); các dòng phía trên là tựa đề. Dòng chỉ
                         có số thứ tự mà không có chữ bị bỏ.
  3. Phụ lục tra cứu   — ô "PHỤ LỤC TRA CỨU"; hàng tiêu đề "Đề mục | Tiêu đề nhóm |
                         Trang số | Ghi chú"; hàng đánh số "(1) (2) …"; rồi dữ liệu
                         (đậm/thường theo Excel).
  4. Bảng thống kê hồ sơ (sheet chính, BẮT BUỘC) — hàng tiêu đề có cột Hồ sơ số ("Số
                         hồ sơ" hoặc "Hồ sơ số (đơn vị bảo quản)") và cột "Tên nhóm …"; 13
                         hoặc 14 cột (có/không cột "Chú giải"), bề rộng cột đặt theo tên
                         cột; hàng đánh số "(1) … (13)"; thân
                         bảng gồm dòng TÊN NHÓM (cột Số hồ sơ trống, cột Tên nhóm có
                         chữ) và dòng HỒ SƠ (cột Số hồ sơ có số). Số cặp/phông/mục
                         lục/hồ sơ là SỐ trong Excel sẽ được đệm 0 (2/3/2/4 chữ số).
  5. Chứng từ kết thúc — tên sheet chứa "CTKT" / "Chứng từ" là yếu tố xác định; nội
                         dung có ô "CHỨNG TỪ KẾT THÚC" (hoặc viết tắt "CTKT", sẽ được in
                         đầy đủ), các dòng nội dung (chép nguyên văn), dòng ngày tháng
                         và khối chữ ký chia trái/phải theo vị trí (PHÊ DUYỆT / chức
                         danh bên trái; ngày tháng, Người lập, tên người ký bên phải;
                         dòng trống trong Excel thành khoảng cách).
  Sheet "Biên mục" (tên sheet "Biên mục" hoặc nội dung MỤC LỤC VĂN BẢN, TÀI LIỆU) KHÔNG
  thuộc đầu ra và bị bỏ qua.
  Thiếu tờ bìa / chữ viết tắt / phụ lục / chứng từ thì bỏ qua mục đó; thiếu bảng
  thống kê hồ sơ thì báo lỗi ConvertError (thông báo tiếng Việt).

BỐ CỤC WORD ĐẦU RA (5 mục = 5 section, A4 ngang mặc định hoặc Letter ngang)
---------------------------------------------------------------------------
  Mục 1  Tờ bìa: khung viền trang, tên cơ quan (màu đỏ nếu Excel đỏ), gạch ngắn,
         tựa "MỤC LỤC HỒ SƠ", tên phông/khoá; khối "Phông số / Mục lục số / Số
         trang / Thời hạn bảo quản" là bảng nổi neo ở ĐÁY vùng chữ; dòng địa danh.
  Mục 2  Bảng chữ viết tắt: bảng 4 cột, hàng tiêu đề lặp lại khi sang trang.
  Mục 3  Phụ lục tra cứu: bảng 4 cột, 2 hàng tiêu đề lặp lại khi sang trang.
  Mục 4  Bảng thống kê hồ sơ: MỘT bảng duy nhất 13 cột, 2 hàng tiêu đề lặp lại ở
         mỗi trang, dòng tên nhóm in đậm căn trái (các ô khác của dòng đó giữ nguyên theo
         Excel), dòng hồ sơ không bị cắt qua trang.
         Cỡ chữ bảng tuỳ chọn (mặc định 8,5 pt; hàng tiêu đề nhỏ hơn 1,5 pt).
  Mục 5  Chứng từ kết thúc và khối chữ ký hai bên.
  Số trang: trường PAGE ở giữa đầu trang, Times New Roman 13, KHÔNG in trên tờ bìa;
  tuỳ chọn page_number_start = 2 (trang 2 mang số 2, mặc định) hoặc 1 (trang 2
  mang số 1, không tính bìa).

Thư viện cần: openpyxl, python-docx. Dùng như mô-đun (app.py gọi convert_workbook)
hoặc dòng lệnh:
    python3 muc_luc_ho_so.py A.xlsx [B.xlsx ...] [-o THU_MUC] [--start 2|1]
                             [--page A4|Letter] [--font 8.5]
Mặc định kết quả được ghi vào thư mục "ket_qua" cạnh tệp Excel.
"""

from __future__ import annotations

import argparse
import datetime as dt
import io
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from typing import Callable, Sequence

import openpyxl
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips

__all__ = [
    "DEFAULT_OPTIONS", "ConvertError", "WorkbookData", "read_workbook",
    "build_document", "convert_workbook", "convert_file", "output_name_for",
]

# ---------------------------------------------------------------------------
# Hằng số
# ---------------------------------------------------------------------------
DEFAULT_OPTIONS = {"page_size": "A4", "page_number_start": 2, "table_font_pt": 8.5}

FONT_NAME = "Times New Roman"
INCH = 1440                                   # twips
PAGE_SIZES = {"A4": (16838, 11906), "Letter": (15840, 12240)}   # (rộng, cao) nằm ngang
HEADER_DISTANCE = 708                         # khoảng cách đầu/chân trang

# Lề từng mục: (trên, phải, dưới, trái) — twips
MARGINS_COVER = (INCH, INCH, INCH, INCH)
MARGINS_ABBR = (INCH, INCH, INCH, INCH)
MARGINS_APPENDIX = (INCH, INCH // 2, INCH, INCH // 2)
MARGINS_MAIN = (INCH // 2, INCH // 2, INCH // 2, INCH // 2)
MARGINS_CLOSING = (INCH, INCH, INCH, INCH)

# Tỉ lệ cột các bảng (đo từ tệp mẫu), dùng khi số cột trùng với mẫu
ABBR_TABLE_WIDTH = int(8.3 * INCH)
ABBR_PROPORTIONS = (1764, 3368, 4671, 2064)
APPENDIX_PROPORTIONS = (1697, 7611, 1961, 1961)
APPENDIX_ROW_HEIGHT = 345        # twips (0,24"): 2 hàng tiêu đề + 37 dòng vừa 2 trang như mẫu
EXCEL_DEFAULT_COL_WIDTH = 8.43

# Vai trò cột của bảng chính, nhận theo chữ trong ô tiêu đề; số chữ số đệm 0
ROLE_PATTERNS = (
    ("TEN", ("tên nhóm",)),
    ("HOSO", ("số hồ sơ", "hồ sơ số")),      # "Số hồ sơ" | "Hồ sơ số (đơn vị bảo quản)"
    ("PHONG", ("phông",)),
    ("MUCLUC", ("mục lục",)),
    ("CAP", ("cặp", "hộp")),                  # "Số cặp (hộp)" | "Hộp (cặp) số"
)
ZERO_PAD = {"HOSO": 4, "PHONG": 3, "MUCLUC": 2, "CAP": 2}

# Bề rộng tương đối từng cột bảng thống kê theo CHỮ trong ô tiêu đề (đo từ tệp mẫu 13 cột; cột lạ
# như "Chú giải" lấy bề rộng cột hẹp). Xét theo thứ tự để "tên nhóm và tên hồ sơ" không bị nhầm
# với "hồ sơ số", "số tài liệu" không bị nhầm với "số trang".
MAIN_COL_WIDTH_RULES = (
    ("tên nhóm", 4969), ("thư mục", 593), ("tài liệu", 776), ("số trang", 607),
    ("thời gian", 998), ("thời hạn", 895), ("độ mật", 593), ("chú giải", 593),
    ("từ kh", 593), ("ghi chú", 593), ("mục lục", 608), ("phông", 607),
    ("cặp", 607), ("hộp", 607), ("số hồ sơ", 998), ("hồ sơ số", 998),
)
MAIN_COL_WIDTH_DEFAULT = 593

DETECT_ROWS, DETECT_COLS = 60, 20            # vùng quét để nhận diện sheet
CLOSING_TITLE = "CHỨNG TỪ KẾT THÚC"          # tiêu đề in khi sheet CTKT viết tắt / thiếu tiêu đề

# Thứ tự phần tử con theo lược đồ OOXML (để chèn XML đúng chỗ)
_SECTPR_SEQ = (
    "w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr", "w:type",
    "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType",
    "w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection",
    "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange",
)
_TBLPR_SEQ = (
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
    "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd",
    "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
    "w:tblDescription", "w:tblPrChange",
)
_TRPR_SEQ = (
    "w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore", "w:wAfter",
    "w:cantSplit", "w:trHeight", "w:tblHeader", "w:tblCellSpacing", "w:jc", "w:hidden",
    "w:ins", "w:del", "w:trPrChange",
)
_PPR_SEQ = (
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
    "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
)
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

CENTER, LEFT, JUSTIFY = (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT,
                         WD_ALIGN_PARAGRAPH.JUSTIFY)


class ConvertError(Exception):
    """Lỗi dữ liệu đầu vào hoặc tuỳ chọn; thông báo tiếng Việt để hiển thị cho người dùng."""


# ---------------------------------------------------------------------------
# Mô hình dữ liệu đọc từ Excel
# ---------------------------------------------------------------------------
@dataclass
class TextLine:
    """Một dòng chữ kèm định dạng lấy từ Excel (đậm, màu RGB 'RRGGBB' hoặc None).
    space_before: khoảng cách (pt) phía trên dòng, mô phỏng các dòng trống trong Excel."""
    text: str
    bold: bool = False
    color: str | None = None
    space_before: float = 0.0


@dataclass
class Cell:
    """Một ô bảng: chữ đã chuẩn hoá (đã đệm 0 nếu cần) và cờ đậm từ Excel."""
    text: str
    bold: bool = False


@dataclass
class BodyRow:
    """Dòng thân bảng chính: kind = 'group' (tên nhóm) hoặc 'data' (hồ sơ)."""
    kind: str
    cells: list[Cell]


@dataclass
class CoverData:
    header_lines: list[TextLine]        # tên cơ quan (các dòng trước "MỤC LỤC HỒ SƠ")
    title: TextLine                     # "MỤC LỤC HỒ SƠ"
    subtitle_lines: list[TextLine]      # tên phông, khoá/nhiệm kỳ
    bottom_left: list[TextLine]         # "Phông số: 001", "Mục lục số: 01", "Số trang:"
    bottom_right: list[TextLine]        # "Thời hạn bảo quản", "Vĩnh viễn"
    place_date: list[TextLine]          # "Phú Ninh, 7/2026"


@dataclass
class TableData:
    """Bảng chữ viết tắt / phụ lục tra cứu."""
    title_lines: list[TextLine]
    headers: list[str]
    numbering: list[str] | None         # hàng "(1) (2) …" nếu có
    rows: list[list[Cell]]
    col_widths: list[float]             # độ rộng cột trong Excel (ký tự)


@dataclass
class MainTableData:
    """Bảng thống kê hồ sơ."""
    title_lines: list[TextLine]
    headers: list[str]
    numbering: list[str] | None
    roles: dict[str, int]               # vai trò -> chỉ số cột (HOSO, TEN, PHONG, MUCLUC, CAP)
    rows: list[BodyRow]
    col_widths: list[float]

    @property
    def n_data_rows(self) -> int:
        return sum(1 for r in self.rows if r.kind == "data")

    @property
    def n_group_rows(self) -> int:
        return sum(1 for r in self.rows if r.kind == "group")


@dataclass
class ClosingData:
    title: TextLine
    body_lines: list[TextLine]
    left_lines: list[TextLine]          # khối chữ ký bên trái (PHÊ DUYỆT …)
    right_lines: list[TextLine]         # khối bên phải (dòng ngày tháng, Người lập …)


@dataclass
class WorkbookData:
    main_table: MainTableData
    cover: CoverData | None = None
    abbreviations: TableData | None = None
    appendix: TableData | None = None
    closing: ClosingData | None = None
    sheet_names: dict[str, str] = field(default_factory=dict)   # phần -> tên sheet đã dùng


# ---------------------------------------------------------------------------
# Đọc Excel: chuẩn hoá chữ, nhận diện sheet, tách từng phần
# ---------------------------------------------------------------------------
class _PartNotFound(Exception):
    """Sheet không có điểm neo cần thiết (dùng nội bộ, chuyển thành bỏ qua / ConvertError)."""


@dataclass
class _XCell:
    """Ô Excel có chữ: cột (1-based), giá trị gốc, chữ đã chuẩn hoá, đậm, màu, số dòng (1-based)."""
    col: int
    raw: object
    text: str
    bold: bool
    color: str | None
    row: int = 0


_INNER_WS = re.compile(r"[^\S\n]+")      # khoảng trắng trừ xuống dòng
_ANY_WS = re.compile(r"\s+")
_NUMBERING = re.compile(r"^\(\d+\)$")


def normalize_text(value) -> str:
    """None→''; số nguyên→'12'; số thực nguyên→'12'; ngày→'dd/mm/yyyy'; chuỗi: bỏ khoảng
    trắng thừa, gộp nhiều dấu cách thành một, GIỮ xuống dòng (\\n)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(round(value, 10))
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, dt.time):
        return value.strftime("%H:%M")
    text = unicodedata.normalize("NFC", str(value)).replace("\r\n", "\n").replace("\r", "\n")
    lines = [_INNER_WS.sub(" ", ln).strip() for ln in text.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def _key(text: str) -> str:
    """Khoá so khớp: NFC, chữ thường, mọi khoảng trắng (kể cả xuống dòng) thành một dấu cách."""
    return _ANY_WS.sub(" ", unicodedata.normalize("NFC", text)).strip().lower()


def _font_color(cell) -> str | None:
    """Màu chữ RGB 'RRGGBB' nếu Excel đặt màu rõ ràng khác đen; màu theme/mặc định → None."""
    color = cell.font.color
    if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
        return None
    rgb = color.rgb[-6:].upper()
    return None if rgb == "000000" else rgb


def _is_int(value) -> bool:
    return (isinstance(value, int) and not isinstance(value, bool)) or (
        isinstance(value, float) and value.is_integer())


MAX_EMPTY_RUN = 500        # dừng quét sheet sau chừng này dòng trống liên tiếp (sheet bị định dạng tới dòng 1.048.576)


def _text_rows(ws) -> list[list[_XCell]]:
    """Mọi dòng có chữ của sheet, mỗi dòng là danh sách ô có chữ (theo thứ tự cột). Dừng sớm khi gặp
    quá nhiều dòng trống liên tiếp: có tệp Excel định dạng cả triệu dòng trống phía dưới bảng."""
    rows = []
    empty_run = 0
    for row in ws.iter_rows():
        cells = [_XCell(c.column, c.value, text, bool(c.font.b), _font_color(c), c.row)
                 for c in row if (text := normalize_text(c.value))]
        if cells:
            rows.append(cells)
            empty_run = 0
        else:
            empty_run += 1
            if empty_run > MAX_EMPTY_RUN:
                break
    return rows


def _line_of(cells: Sequence[_XCell]) -> TextLine:
    """Ghép các ô của một dòng thành một dòng chữ; đậm/màu lấy theo ô đầu tiên."""
    return TextLine(" ".join(c.text for c in cells), cells[0].bold, cells[0].color)


def _row_has(cells: Sequence[_XCell], needle: str) -> bool:
    return any(needle in _key(c.text) for c in cells)


def _find_row(rows, pred: Callable[[Sequence[_XCell]], bool], start=0) -> int | None:
    return next((i for i in range(start, len(rows)) if pred(rows[i])), None)


def _excel_col_width(ws, col: int) -> float:
    """Độ rộng cột Excel (đơn vị ký tự); openpyxl gộp các cột cùng độ rộng thành khoảng min..max."""
    for dim in ws.column_dimensions.values():
        if dim.width and (dim.min or 0) <= col <= (dim.max or 0):
            return float(dim.width)
    return float(ws.sheet_format.defaultColWidth or EXCEL_DEFAULT_COL_WIDTH)


# --- nhận diện sheet --------------------------------------------------------
def _sheet_keys(ws) -> list[list[str]]:
    """Khoá so khớp của các ô chữ trong vùng quét đầu sheet, nhóm theo dòng.
    Giới hạn max_col theo bề rộng sheet vì iter_rows TẠO ô còn thiếu, nếu không ws.max_column
    sẽ bị nới rộng và làm sai ranh giới trái/phải của tờ bìa."""
    out = []
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, DETECT_ROWS),
                            max_col=min(ws.max_column, DETECT_COLS)):
        keys = [_key(c.value) for c in row if isinstance(c.value, str) and c.value.strip()]
        if keys:
            out.append(keys)
    return out


def _has(rows_keys, needle: str) -> bool:
    return any(needle in k for row in rows_keys for k in row)


def _has_exact(rows_keys, needle: str) -> bool:
    return any(k == needle for row in rows_keys for k in row)


def _strip_accents(text: str) -> str:
    """'Chứng từ kết thúc' → 'chung tu ket thuc' (chữ thường, bỏ dấu, đ→d) để so TÊN sheet."""
    text = unicodedata.normalize("NFD", text.lower().replace("đ", "d"))
    return _ANY_WS.sub(" ", "".join(ch for ch in text if not unicodedata.combining(ch))).strip()


def _name_has(ws, *needles: str) -> bool:
    """Tên sheet (bỏ dấu) chứa một trong các cụm `needles` (viết không dấu) như một từ trọn vẹn,
    vd. 'CTKT', 'ctkt 2026', 'Chứng từ kết thúc', 'Tờ bìa', 'Biên mục'."""
    name = _strip_accents(ws.title)
    return any(re.search(r"(?<![a-z0-9])%s(?![a-z0-9])" % re.escape(n), name) for n in needles)


def _is_main_header_keys(keys: Sequence[str]) -> bool:
    """Hàng tiêu đề bảng thống kê: có cột Hồ sơ số ("Số hồ sơ" / "Hồ sơ số…") VÀ cột "Tên nhóm…",
    và là một hàng nhiều cột (tránh nhầm với dòng "Hồ sơ số: 0001" của sheet Biên mục)."""
    return (len(keys) >= 5
            and any(any(p in k for p in dict(ROLE_PATTERNS)["HOSO"]) for k in keys)
            and any("tên nhóm" in k for k in keys))


def _is_main_sheet(rows_keys) -> bool:
    return _has(rows_keys, "bảng thống kê hồ sơ") or any(_is_main_header_keys(row) for row in rows_keys)


def _is_closing_name(ws, _keys=None) -> bool:
    """Tên sheet cho biết đây là Chứng từ kết thúc ('CTKT', 'Chứng từ…', '…kết thúc')."""
    return _name_has(ws, "ctkt", "chung tu", "ket thuc")


def _is_bien_muc(ws, keys) -> bool:
    """Sheet 'Biên mục' (MỤC LỤC VĂN BẢN, TÀI LIỆU) — không thuộc đầu ra; nhận theo tên hoặc nội dung."""
    return _name_has(ws, "bien muc") or _has(keys, "mục lục văn bản")


# Quy tắc nhận diện từng phần: (tên phần, các điều kiện (ws, keys) theo thứ tự ưu tiên — nội dung
# trước, tên sheet sau —, chỉ số dự phòng theo THỨ TỰ SHEET CHUẨN: Tờ bìa 0 → Bảng chữ viết tắt 1 →
# Phụ lục tra cứu 2 → Bảng thống kê 3 → [Biên mục, bị bỏ qua] → Chứng từ kết thúc = sheet cuối).
# Riêng Chứng từ kết thúc: TÊN sheet chứa "CTKT" là yếu tố quyết định nên xét tên trước.
def _name_is(ws, abbr: str) -> bool:
    """Tên sheet đúng bằng chữ viết tắt quy ước (không phân biệt hoa thường, bỏ khoảng trắng thừa)."""
    return _strip_accents(ws.title) == abbr


# Tên sheet QUY ƯỚC của mẫu: TB (Tờ bìa) → VT (Bảng chữ cái viết tắt) → PL (Phụ lục tra cứu) →
# ML (Bảng thống kê hồ sơ) → CTKT (Chứng từ kết thúc). Đây là cách nhận diện chính; các điều kiện
# theo nội dung / từ khoá trong tên chỉ là dự phòng cho tệp cũ chưa đổi tên sheet.
SHEET_ABBR = {"cover": "tb", "abbreviations": "vt", "appendix": "pl", "main_table": "ml", "closing": "ctkt"}

_PART_RULES = (
    ("main_table", (lambda ws, t: _name_is(ws, "ml"),
                    lambda ws, t: _is_main_sheet(t)), 3),
    ("closing", (lambda ws, t: _name_is(ws, "ctkt"),
                 _is_closing_name,
                 lambda ws, t: _has(t, "chứng từ kết thúc") or _has_exact(t, "ctkt")), -1),
    ("appendix", (lambda ws, t: _name_is(ws, "pl"),
                  lambda ws, t: _has(t, "phụ lục tra cứu"),
                  lambda ws, t: _name_has(ws, "phu luc", "tra cuu")), 2),
    ("abbreviations", (lambda ws, t: _name_is(ws, "vt"),
                       lambda ws, t: _has(t, "chữ viết tắt"),
                       lambda ws, t: _name_has(ws, "viet tat", "chu cai")), 1),
    ("cover", (lambda ws, t: _name_is(ws, "tb"),
               lambda ws, t: _has_exact(t, "mục lục hồ sơ"),
               lambda ws, t: _has(t, "mục lục hồ sơ"),
               lambda ws, t: _name_has(ws, "bia")), 0),
)


# --- tờ bìa -----------------------------------------------------------------
def _parse_cover(ws) -> CoverData:
    rows = _text_rows(ws)
    if not rows:
        raise _PartNotFound("sheet trống")
    # Ranh giới trái/phải = 60% bề rộng sheet (không chỉ bề rộng phần có chữ, kẻo bìa không có
    # cột phải sẽ xếp mọi ô sang phải); ô chữ đầu tiên của dòng luôn thuộc bên trái.
    split_col = 0.6 * max(ws.max_column, max(c.col for cells in rows for c in cells))

    def left(cells):
        return [c for c in cells if c is cells[0] or c.col <= split_col]

    def right(cells):
        return [c for c in cells if c is not cells[0] and c.col > split_col]

    def left_key(cells):
        return _key(" ".join(c.text for c in left(cells)))

    title_i = _find_row(rows, lambda cells: _key(" ".join(c.text for c in cells)) == "mục lục hồ sơ")
    if title_i is None:
        title_i = _find_row(rows, lambda cells: _row_has(cells, "mục lục hồ sơ"))
    if title_i is None:
        raise _PartNotFound("không có ô 'MỤC LỤC HỒ SƠ'")

    header_lines = [_line_of(cells) for cells in rows[:title_i]]
    title = _line_of(rows[title_i])

    start = _find_row(rows, lambda cells: "phông số" in left_key(cells), title_i + 1)
    if start is None:
        return CoverData(header_lines, title, [_line_of(c) for c in rows[title_i + 1:]], [], [], [])
    subtitle_lines = [_line_of(cells) for cells in rows[title_i + 1:start]]

    end = _find_row(rows, lambda cells: "số trang" in left_key(cells), start)
    if end is None:   # không có "Số trang": khối kéo dài tới dòng nhãn "…:" cuối cùng
        end = max(i for i in range(start, len(rows)) if i == start or ":" in left_key(rows[i]))
    bottom_left = [_line_of(left(cells)) for cells in rows[start:end + 1] if left(cells)]
    bottom_right = [_line_of(right(cells)) for cells in rows[start:end + 1] if right(cells)]
    place_date = [_line_of(cells) for cells in rows[end + 1:]]
    return CoverData(header_lines, title, subtitle_lines, bottom_left, bottom_right, place_date)


# --- bảng chữ viết tắt, phụ lục tra cứu ------------------------------------
def _parse_grid(ws, header_pred, drop_row=None) -> TableData:
    """Bảng đơn giản: tựa đề → hàng tiêu đề (xác định bằng header_pred) → hàng đánh số
    '(n)' (nếu có) → dữ liệu. drop_row(cells) → True để bỏ dòng (vd. chỉ có số TT)."""
    rows = _text_rows(ws)
    hi = _find_row(rows, lambda cells: len(cells) >= 2 and header_pred(cells))
    if hi is None:
        raise _PartNotFound("không tìm thấy hàng tiêu đề bảng")
    header = rows[hi]
    cols = [c.col for c in header]
    body_start = hi + 1
    numbering = None
    if body_start < len(rows) and all(_NUMBERING.match(c.text) for c in rows[body_start]):
        by_col = {c.col: c.text for c in rows[body_start]}
        numbering = [by_col.get(col, "") for col in cols]
        body_start += 1
    data = []
    for cells in rows[body_start:]:
        by_col = {c.col: c for c in cells}
        row = [Cell(by_col[col].text, by_col[col].bold) if col in by_col else Cell("") for col in cols]
        if any(c.text for c in row) and not (drop_row and drop_row(row)):
            data.append(row)
    return TableData(
        title_lines=[_line_of(cells) for cells in rows[:hi]],
        headers=[c.text for c in header],
        numbering=numbering,
        rows=data,
        col_widths=[_excel_col_width(ws, col) for col in cols],
    )


def _parse_abbreviations(ws) -> TableData:
    return _parse_grid(
        ws,
        header_pred=lambda cells: _row_has(cells, "chữ viết tắt"),
        drop_row=lambda row: not any(c.text for c in row[1:]),   # chỉ có số TT → bỏ
    )


def _parse_appendix(ws) -> TableData:
    return _parse_grid(
        ws,
        header_pred=lambda cells: _row_has(cells, "đề mục") or _row_has(cells, "tiêu đề nhóm"),
    )


# --- bảng thống kê hồ sơ ---------------------------------------------------
def _parse_main_table(ws) -> MainTableData:
    rows = _text_rows(ws)
    hi = _find_row(rows, lambda cells: _is_main_header_keys([_key(c.text) for c in cells]))
    if hi is None:
        raise _PartNotFound("không tìm thấy hàng tiêu đề có 'Số hồ sơ'/'Hồ sơ số' và 'Tên nhóm'")
    header = rows[hi]
    cols = [c.col for c in header]

    roles: dict[str, int] = {}
    for j, c in enumerate(header):
        k = _key(c.text)
        for role, patterns in ROLE_PATTERNS:
            if role not in roles and any(p in k for p in patterns):
                roles[role] = j
    pad_by_index = {idx: ZERO_PAD[role] for role, idx in roles.items() if role in ZERO_PAD}

    body_start = hi + 1
    numbering = None
    if body_start < len(rows) and all(_NUMBERING.match(c.text) for c in rows[body_start]):
        by_col = {c.col: c.text for c in rows[body_start]}
        numbering = [by_col.get(col, "") for col in cols]
        body_start += 1

    hoso, ten = roles["HOSO"], roles["TEN"]
    body: list[BodyRow] = []
    for cells in rows[body_start:]:
        by_col = {c.col: c for c in cells}
        row_cells = []
        for j, col in enumerate(cols):
            xc = by_col.get(col)
            if xc is None:
                row_cells.append(Cell(""))
                continue
            text = xc.text
            if j in pad_by_index and _is_int(xc.raw):
                text = text.zfill(pad_by_index[j])
            row_cells.append(Cell(text, xc.bold))
        if not any(c.text for c in row_cells):
            continue
        kind = "group" if not row_cells[hoso].text and row_cells[ten].text else "data"
        body.append(BodyRow(kind, row_cells))

    return MainTableData(
        title_lines=[_line_of(cells) for cells in rows[:hi]],
        headers=[c.text for c in header],
        numbering=numbering,
        roles=roles,
        rows=body,
        col_widths=[_excel_col_width(ws, col) for col in cols],
    )


# --- chứng từ kết thúc ------------------------------------------------------
_SIG_DATE = re.compile(r"^[^,:;]{1,40},\s*ngày\b")      # "Phú Ninh, Ngày ….. tháng 7 năm 2026"


def _is_date_line(c: _XCell) -> bool:
    """Dòng địa danh – ngày tháng của khối chữ ký: '<Địa danh>, Ngày …' ở bất kỳ cột nào, hoặc
    'ngày … tháng/năm …' nằm bên phải (từ cột G). Dòng nội dung kiểu 'Tài liệu có thời gian từ ngày
    01 tháng 01 năm 2006' ở cột trái không bị nhầm."""
    k = _key(c.text)
    return bool(_SIG_DATE.match(k)) or (c.col > 6 and "ngày" in k and ("tháng" in k or "năm" in k))


def _is_signature_cell(c: _XCell) -> bool:
    """Ô mở đầu khối chữ ký: 'PHÊ DUYỆT' / 'Người lập…' ở bất kỳ cột nào, hoặc dòng ngày tháng."""
    k = _key(c.text)
    return "phê duyệt" in k or "người lập" in k or _is_date_line(c)


def _split_signature(cells: Sequence[_XCell]) -> tuple[list[TextLine], list[TextLine]]:
    """Chia các ô của khối chữ ký thành hai cột trái/phải theo VỊ TRÍ TƯƠNG ĐỐI (điểm giữa của
    các cột có chữ, không cố định ở cột G vì mỗi tệp đặt khác nhau); dòng ngày tháng và 'Người lập'
    luôn bên phải, 'PHÊ DUYỆT' luôn bên trái. Dòng trống giữa hai ô (vd. chức danh → tên người ký)
    được giữ lại thành khoảng cách phía trên (13 pt mỗi dòng, tối đa 4 dòng)."""
    if not cells:
        return [], []
    mid = (min(c.col for c in cells) + max(c.col for c in cells)) / 2
    left: list[tuple[int, TextLine]] = []
    right: list[tuple[int, TextLine]] = []
    for c in cells:
        k = _key(c.text)
        if _is_date_line(c) or "người lập" in k:
            side = right
        elif "phê duyệt" in k:
            side = left
        else:
            side = left if c.col <= mid else right
        side.append((c.row, TextLine(c.text, c.bold, c.color)))

    def finish(block):
        block.sort(key=lambda rc: rc[0])
        lines = []
        for i, (row, ln) in enumerate(block):
            if i:
                gap = row - block[i - 1][0] - 1
                ln.space_before = 13.0 * min(max(gap, 0), 4)
            lines.append(ln)
        return lines

    return finish(left), finish(right)


def _is_closing_title(cells: Sequence[_XCell]) -> bool:
    """Dòng tiêu đề chứng từ: 'CHỨNG TỪ KẾT THÚC' hoặc viết tắt 'CTKT'."""
    k = _key(" ".join(c.text for c in cells))
    return "chứng từ kết thúc" in k or k == "ctkt"


def _parse_closing(ws, trusted: bool = False) -> ClosingData:
    """trusted=True: sheet đã được xác định là Chứng từ kết thúc theo TÊN sheet, nên không bắt buộc
    có dòng tiêu đề (khi thiếu sẽ in tiêu đề chuẩn). Tiêu đề viết tắt 'CTKT' được in đầy đủ."""
    rows = _text_rows(ws)
    # Điểm neo tiêu đề phải nằm trong 3 dòng chữ đầu; dòng lạ phía trên bị bỏ qua
    title_i = _find_row(rows[:3], _is_closing_title)
    if title_i is None and not (trusted and rows):
        raise _PartNotFound("không có dòng 'CHỨNG TỪ KẾT THÚC' (hoặc 'CTKT') ở đầu sheet")
    if title_i is None:
        title, body_from = TextLine(CLOSING_TITLE, bold=True), 0
    else:
        rows = rows[title_i:]
        title, body_from = _line_of(rows[0]), 1
        if _key(title.text) == "ctkt":
            title = TextLine(CLOSING_TITLE, True, title.color)
    sig = _find_row(rows, lambda cells: any(_is_signature_cell(c) for c in cells), body_from)
    if sig is None:
        sig = len(rows)
    left, right = _split_signature([c for cells in rows[sig:] for c in cells])
    return ClosingData(
        title=title,
        body_lines=[_line_of(cells) for cells in rows[body_from:sig]],
        left_lines=left,
        right_lines=right,
    )


_PARSERS = {
    "cover": _parse_cover,
    "abbreviations": _parse_abbreviations,
    "appendix": _parse_appendix,
    "main_table": _parse_main_table,
    "closing": _parse_closing,
}


def _extract(wb) -> WorkbookData:
    """Nhận diện sheet theo nội dung → tên sheet → thứ tự sheet chuẩn, rồi tách 5 phần."""
    sheets = wb.worksheets
    keys = [_sheet_keys(ws) for ws in sheets]
    # Sheet "Biên mục" không bao giờ được gán cho phần nào (kể cả khi dự phòng theo thứ tự)
    taken: set[int] = {i for i, ws in enumerate(sheets) if _is_bien_muc(ws, keys[i])}
    parts: dict[str, object] = {}
    names: dict[str, str] = {}
    for part, predicates, fallback in _PART_RULES:
        idx = next((i for pred in predicates for i in range(len(sheets))
                    if i not in taken and pred(sheets[i], keys[i])), None)
        if idx is None and fallback >= 0:
            idx = fallback if fallback < len(sheets) and fallback not in taken else None
        elif idx is None:      # Chứng từ kết thúc: sheet cuối cùng chưa dùng (sau khi bỏ Biên mục)
            idx = next((i for i in range(len(sheets) - 1, -1, -1) if i not in taken), None)
        if idx is None:
            continue
        try:
            if part == "closing":
                parts[part] = _parse_closing(sheets[idx], trusted=_is_closing_name(sheets[idx]))
            else:
                parts[part] = _PARSERS[part](sheets[idx])
        except _PartNotFound as exc:
            if part == "main_table":
                raise ConvertError(
                    "Không tìm thấy bảng thống kê hồ sơ (sheet có hàng tiêu đề "
                    "'Số hồ sơ'/'Hồ sơ số' và 'Tên nhóm và tên hồ sơ'): %s." % exc) from None
            continue
        taken.add(idx)
        names[part] = sheets[idx].title
    if "main_table" not in parts:
        raise ConvertError("Tệp Excel không có sheet BẢNG THỐNG KÊ HỒ SƠ "
                           "(hàng tiêu đề 'Số hồ sơ'/'Hồ sơ số' và 'Tên nhóm và tên hồ sơ').")
    return WorkbookData(
        main_table=parts["main_table"],
        cover=parts.get("cover"),
        abbreviations=parts.get("abbreviations"),
        appendix=parts.get("appendix"),
        closing=parts.get("closing"),
        sheet_names=names,
    )


def read_workbook(src) -> WorkbookData:
    """Đọc sổ Excel từ đường dẫn, bytes hoặc đối tượng tệp → WorkbookData."""
    if isinstance(src, (bytes, bytearray, memoryview)):
        src = io.BytesIO(bytes(src))
    elif isinstance(src, os.PathLike):
        src = os.fspath(src)
    try:
        wb = openpyxl.load_workbook(src, data_only=True)
    except Exception as exc:   # tệp hỏng, không phải .xlsx, không tồn tại…
        raise ConvertError("Không mở được tệp Excel (cần tệp .xlsx hợp lệ): %s" % exc) from None
    try:
        return _extract(wb)
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Tiện ích dựng Word (XML cấp thấp)
# ---------------------------------------------------------------------------
def _el(tag: str, **attrs):
    """Tạo phần tử w:tag với các thuộc tính w:… (giá trị chuyển sang chuỗi)."""
    el = OxmlElement(tag)
    for name, value in attrs.items():
        el.set(qn("w:" + name), str(value))
    return el


def _set_child(parent, el, seq: Sequence[str]):
    """Thay (hoặc chèn) phần tử con `el` vào `parent` đúng vị trí theo thứ tự lược đồ `seq`."""
    tag = "w:" + el.tag.rsplit("}", 1)[1]
    for old in parent.findall(qn(tag)):
        parent.remove(old)
    parent.insert_element_before(el, *seq[seq.index(tag) + 1:])


def _style_run(run, size_pt: float, bold=False, color: str | None = None):
    font = run.font
    font.name = FONT_NAME
    font.size = Pt(size_pt)
    font.bold = bool(bold)
    rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT_NAME)
    if color:
        font.color.rgb = RGBColor.from_string(color)


def _style_paragraph_mark(paragraph, size_pt: float):
    """Font/cỡ của dấu đoạn (quyết định chiều cao đoạn trống, ô trống)."""
    rpr = _el("w:rPr")
    rpr.append(_el("w:rFonts", ascii=FONT_NAME, hAnsi=FONT_NAME, cs=FONT_NAME, eastAsia=FONT_NAME))
    rpr.append(_el("w:sz", val=int(round(size_pt * 2))))
    _set_child(paragraph._p.get_or_add_pPr(), rpr, _PPR_SEQ)


def _add_text(paragraph, text: str, size_pt: float, bold=False, color=None):
    """Thêm một run; '\\n' trong chữ thành ngắt dòng trong đoạn."""
    run = paragraph.add_run()
    _style_run(run, size_pt, bold, color)
    first, *rest = text.split("\n")
    run.text = first
    for line in rest:
        run.add_break()
        run.add_text(line)
    return run


def _add_par(container, text="", *, size=13.0, bold=False, color=None, align=CENTER,
             before=0.0, after=0.0, line=1.0, left=None, right=None, keep_next=False,
             reuse_first=False):
    """Thêm đoạn vào Document hoặc ô bảng. reuse_first=True: dùng lại đoạn trống sẵn có
    của ô. before/after tính bằng pt; left/right là thụt lề tính bằng twips."""
    p = container.paragraphs[0] if reuse_first else container.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), line
    if left is not None:
        pf.left_indent = Twips(left)
    if right is not None:
        pf.right_indent = Twips(right)
    if keep_next:
        pf.keep_with_next = True
    _style_paragraph_mark(p, size)
    if text:
        _add_text(p, text, size, bold, color)
    return p


def _distribute(total: int, proportions: Sequence[float]) -> list[int]:
    """Chia `total` twips theo tỉ lệ; phần dư dồn vào cột rộng nhất để tổng khớp tuyệt đối."""
    s = float(sum(proportions))
    widths = [int(total * p / s) for p in proportions]
    widths[widths.index(max(widths))] += total - sum(widths)
    return widths


def _add_table(doc, n_rows: int, widths: Sequence[int], *, borders=None,
               cell_margins=None, floating_bottom=False):
    """Bảng bố cục cố định, căn giữa, độ rộng = sum(widths).
    borders=(ngoài, trong) là cỡ viền 1/8 pt, None = không viền.
    cell_margins=(trên, trái, dưới, phải) twips. floating_bottom: bảng nổi neo đáy lề."""
    table = doc.add_table(rows=n_rows, cols=len(widths))
    table.autofit = False                              # w:tblLayout fixed
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblpr = table._tbl.tblPr
    total = sum(widths)
    _set_child(tblpr, _el("w:tblW", w=total, type="dxa"), _TBLPR_SEQ)
    if floating_bottom:
        _set_child(tblpr, _el("w:tblpPr", leftFromText=180, rightFromText=180, vertAnchor="margin",
                              horzAnchor="margin", tblpXSpec="center", tblpYSpec="bottom"), _TBLPR_SEQ)
        _set_child(tblpr, _el("w:tblOverlap", val="never"), _TBLPR_SEQ)
    tblborders = _el("w:tblBorders")
    outer, inner = borders or (None, None)
    for side, size in (("top", outer), ("left", outer), ("bottom", outer), ("right", outer),
                       ("insideH", inner), ("insideV", inner)):
        if size:
            tblborders.append(_el("w:" + side, val="single", sz=size, space=0, color="000000"))
        else:
            tblborders.append(_el("w:" + side, val="none", sz=0, space=0, color="auto"))
    _set_child(tblpr, tblborders, _TBLPR_SEQ)
    if cell_margins:
        mar = _el("w:tblCellMar")
        for side, value in zip(("top", "left", "bottom", "right"), cell_margins):
            mar.append(_el("w:" + side, w=value, type="dxa"))
        _set_child(tblpr, mar, _TBLPR_SEQ)
    for column, w in zip(table.columns, widths):
        column.width = Twips(w)
    return table


def _row_props(row, *, header=False, cant_split=False, min_height=None):
    """Hàng lặp lại đầu trang (tblHeader), không cắt qua trang (cantSplit), cao tối thiểu."""
    trpr = row._tr.get_or_add_trPr()
    if cant_split:
        _set_child(trpr, _el("w:cantSplit"), _TRPR_SEQ)
    if min_height:
        row.height = Twips(min_height)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    if header:
        _set_child(trpr, _el("w:tblHeader"), _TRPR_SEQ)


def _fill_cell(cell, text: str, *, width: int, size: float, bold=False, align=CENTER,
               left=None, keep_next=False):
    cell.width = Twips(width)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _add_par(cell, text, size=size, bold=bold, align=align, left=left, keep_next=keep_next,
             reuse_first=True)


def _page_setup(section, page_size: str, margins) -> int:
    """Đặt khổ giấy nằm ngang + lề (trên, phải, dưới, trái); trả về bề rộng vùng chữ (twips)."""
    width, height = PAGE_SIZES[page_size]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Twips(width), Twips(height)
    top, right, bottom, left = margins
    section.top_margin, section.right_margin = Twips(top), Twips(right)
    section.bottom_margin, section.left_margin = Twips(bottom), Twips(left)
    section.header_distance = section.footer_distance = Twips(HEADER_DISTANCE)
    return width - left - right


# ---------------------------------------------------------------------------
# Dựng từng mục
# ---------------------------------------------------------------------------
def _build_cover(doc, section, cover: CoverData, page_size: str, _font_pt):
    text_w = _page_setup(section, page_size, MARGINS_COVER)
    for ln in cover.header_lines:
        _add_par(doc, ln.text, size=18, bold=ln.bold, color=ln.color)
    if cover.header_lines:
        _add_rule(doc, text_w)
    _add_par(doc, cover.title.text, size=22, bold=True, color=cover.title.color, before=70, after=14)
    for i, ln in enumerate(cover.subtitle_lines):
        _add_par(doc, ln.text, size=14, bold=ln.bold, color=ln.color,
                 before=24 if i == 1 else 0, left=int(1.3 * INCH), right=int(1.3 * INCH))
    _add_cover_bottom(doc, cover, text_w)


def _add_rule(doc, text_w: int):
    """Gạch ngang ngắn (~3,3") căn giữa dưới tên cơ quan: đoạn trống có viền dưới."""
    indent = max(0, (text_w - int(3.3 * INCH)) // 2)
    p = _add_par(doc, size=6, left=indent, right=indent)
    pbdr = _el("w:pBdr")
    pbdr.append(_el("w:bottom", val="single", sz=8, space=1, color="000000"))
    _set_child(p._p.get_or_add_pPr(), pbdr, _PPR_SEQ)


def _add_cover_bottom(doc, cover: CoverData, text_w: int):
    """Khối 'Phông số / Mục lục số / Số trang | Thời hạn bảo quản' + địa danh, là bảng nổi
    không viền neo ở đáy vùng chữ (không xê dịch khi tên phông dài ngắn khác nhau)."""
    n = max(len(cover.bottom_left), len(cover.bottom_right))
    if n == 0 and not cover.place_date:
        return
    widths = _distribute(text_w, (62, 38))
    table = _add_table(doc, n + (1 if cover.place_date else 0), widths, floating_bottom=True)
    rows = list(table.rows)
    for i in range(n):
        left_cell, right_cell = rows[i].cells
        if i < len(cover.bottom_left):
            _fill_label_value(left_cell, cover.bottom_left[i], widths[0])
        else:
            _fill_cell(left_cell, "", width=widths[0], size=15)
        ln = cover.bottom_right[i] if i < len(cover.bottom_right) else None
        _fill_cell(right_cell, ln.text if ln else "", width=widths[1], size=15,
                   bold=ln.bold if ln else False)
    if cover.place_date:
        first, second = rows[n].cells
        cell = first.merge(second)
        cell.width = Twips(text_w)
        for i, ln in enumerate(cover.place_date):
            _add_par(cell, ln.text, size=15, bold=ln.bold, color=ln.color,
                     before=6 if i == 0 else 0, reuse_first=(i == 0))


def _fill_label_value(cell, line: TextLine, width: int):
    """'Phông số: 001' → nhãn thường + giá trị sau dấu ':' in đậm; thụt lề trái 0,3"."""
    cell.width = Twips(width)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = _add_par(cell, size=15, align=LEFT, left=int(0.3 * INCH), reuse_first=True)
    label, sep, value = line.text.partition(":")
    if sep:
        _add_text(p, label.strip() + ":", 15, line.bold, line.color)
        if value.strip():
            _add_text(p, " " + value.strip(), 15, True, line.color)
    else:
        _add_text(p, line.text, 15, line.bold, line.color)


def _build_abbreviations(doc, section, abbr: TableData, page_size: str, _font_pt):
    text_w = _page_setup(section, page_size, MARGINS_ABBR)
    for ln in abbr.title_lines:
        _add_par(doc, ln.text, size=12, bold=True, after=12)
    proportions = ABBR_PROPORTIONS if len(abbr.headers) == len(ABBR_PROPORTIONS) else abbr.col_widths
    widths = _distribute(min(ABBR_TABLE_WIDTH, text_w), proportions)
    table = _add_table(doc, 1 + len(abbr.rows), widths, borders=(12, 6),
                       cell_margins=(0, int(0.08 * INCH), 0, int(0.08 * INCH)))
    rows = iter(table.rows)
    header = next(rows)
    for cell, text, w in zip(header.cells, abbr.headers, widths):
        _fill_cell(cell, text, width=w, size=12, bold=True)
    _row_props(header, header=True, min_height=int(0.36 * INCH))
    for row, data in zip(rows, abbr.rows):
        for j, (cell, c, w) in enumerate(zip(row.cells, data, widths)):
            _fill_cell(cell, c.text, width=w, size=12, bold=c.bold, align=CENTER if j == 0 else LEFT)
        _row_props(row, cant_split=True, min_height=int(0.26 * INCH))


def _build_appendix(doc, section, appendix: TableData, page_size: str, _font_pt):
    text_w = _page_setup(section, page_size, MARGINS_APPENDIX)
    for ln in appendix.title_lines:
        _add_par(doc, ln.text, size=13, bold=True, after=10)
    proportions = (APPENDIX_PROPORTIONS if len(appendix.headers) == len(APPENDIX_PROPORTIONS)
                   else appendix.col_widths)
    widths = _distribute(int(text_w * 0.95), proportions)
    n_header = 1 + (1 if appendix.numbering else 0)
    table = _add_table(doc, n_header + len(appendix.rows), widths, borders=(6, 6),
                       cell_margins=(0, int(0.06 * INCH), 0, int(0.06 * INCH)))
    rows = iter(table.rows)
    header = next(rows)
    for cell, text, w in zip(header.cells, appendix.headers, widths):
        _fill_cell(cell, text, width=w, size=13, bold=True)
    _row_props(header, header=True, min_height=APPENDIX_ROW_HEIGHT)
    if appendix.numbering:
        numbering = next(rows)
        for cell, text, w in zip(numbering.cells, appendix.numbering, widths):
            _fill_cell(cell, text, width=w, size=13)
        _row_props(numbering, header=True, min_height=APPENDIX_ROW_HEIGHT)
    for row, data in zip(rows, appendix.rows):
        for j, (cell, c, w) in enumerate(zip(row.cells, data, widths)):
            _fill_cell(cell, c.text, width=w, size=13, bold=c.bold, align=LEFT if j == 1 else CENTER)
        _row_props(row, cant_split=True, min_height=APPENDIX_ROW_HEIGHT)


def _main_proportions(headers: Sequence[str]) -> list[int]:
    """Bề rộng tương đối của từng cột theo chữ trong ô tiêu đề (xem MAIN_COL_WIDTH_RULES)."""
    out = []
    for h in headers:
        k = _key(h)
        out.append(next((w for needle, w in MAIN_COL_WIDTH_RULES if needle in k), MAIN_COL_WIDTH_DEFAULT))
    return out


def _build_main_table(doc, section, main: MainTableData, page_size: str, font_pt: float):
    text_w = _page_setup(section, page_size, MARGINS_MAIN)
    last = len(main.title_lines) - 1
    for i, ln in enumerate(main.title_lines):
        _add_par(doc, ln.text, size=11, bold=True, before=24 if i == 0 else 0,
                 after=10 if i == last else 0)
    widths = _distribute(text_w, _main_proportions(main.headers))
    n_header = 1 + (1 if main.numbering else 0)
    table = _add_table(doc, n_header + len(main.rows), widths, borders=(6, 6),
                       cell_margins=(int(0.02 * INCH), int(0.04 * INCH),
                                     int(0.02 * INCH), int(0.04 * INCH)))
    small = font_pt - 1.5
    rows = iter(table.rows)
    header = next(rows)
    for cell, text, w in zip(header.cells, main.headers, widths):
        _fill_cell(cell, text, width=w, size=small, bold=True)
    _row_props(header, header=True, min_height=int(0.5 * INCH))
    if main.numbering:
        numbering = next(rows)
        for cell, text, w in zip(numbering.cells, main.numbering, widths):
            _fill_cell(cell, text, width=w, size=small)
        _row_props(numbering, header=True)
    ten = main.roles["TEN"]
    for row, body in zip(rows, main.rows):
        is_group = body.kind == "group"
        for j, (cell, c, w) in enumerate(zip(row.cells, body.cells, widths)):
            if j == ten:
                # Tên nhóm: đậm, căn trái; tên hồ sơ: đậm theo Excel, căn đều
                bold, align = (True, LEFT) if is_group else (c.bold, JUSTIFY)
            else:
                bold, align = c.bold, CENTER          # mọi ô khác giữ nguyên như Excel
            # dòng tên nhóm đi cùng dòng kế tiếp (không nằm cuối trang một mình)
            _fill_cell(cell, c.text, width=w, size=font_pt, bold=bold, align=align, keep_next=is_group)
        _row_props(row, cant_split=True)


def _build_closing(doc, section, closing: ClosingData, page_size: str, _font_pt):
    text_w = _page_setup(section, page_size, MARGINS_CLOSING)
    _add_par(doc, closing.title.text, size=13, bold=True, after=14)
    for ln in closing.body_lines:
        _add_par(doc, ln.text, size=13, bold=ln.bold, align=LEFT, left=int(0.75 * INCH),
                 after=4, line=1.15)
    if not (closing.left_lines or closing.right_lines):
        return
    widths = _distribute(text_w, (1, 1))
    table = _add_table(doc, 1, widths)
    for cell, lines, w in zip(table.rows[0].cells, (closing.left_lines, closing.right_lines), widths):
        cell.width = Twips(w)
        _style_paragraph_mark(cell.paragraphs[0], 13)
        for i, ln in enumerate(lines):
            _add_par(cell, ln.text, size=13, bold=ln.bold, color=ln.color,
                     before=18 if i == 0 else ln.space_before, reuse_first=(i == 0))


# ---------------------------------------------------------------------------
# Tài liệu: style mặc định, số trang, khung viền
# ---------------------------------------------------------------------------
def _new_document():
    """Tài liệu trống: Times New Roman toàn bộ, Normal 13 pt, giãn dòng đơn, không cách đoạn."""
    doc = Document()
    styles = doc.styles.element
    rfonts = styles.xpath("w:docDefaults/w:rPrDefault/w:rPr/w:rFonts")[0]
    rfonts.attrib.clear()
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT_NAME)
    spacing = styles.xpath("w:docDefaults/w:pPrDefault/w:pPr/w:spacing")[0]
    spacing.attrib.clear()
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    normal = doc.styles["Normal"]
    normal.font.name = FONT_NAME
    normal.font.size = Pt(13)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
    normal.element.rPr.rFonts.set(qn("w:cs"), FONT_NAME)
    pf = normal.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(0), 1.0
    return doc


def _add_page_field(paragraph, size_pt: float):
    """Trường PAGE: fldChar begin / instrText / separate / chữ đệm / end."""
    def run():
        r = paragraph.add_run()
        _style_run(r, size_pt)
        return r

    run()._r.append(_el("w:fldChar", fldCharType="begin"))
    instr = _el("w:instrText")
    instr.set(_XML_SPACE, "preserve")
    instr.text = " PAGE "
    run()._r.append(instr)
    run()._r.append(_el("w:fldChar", fldCharType="separate"))
    run().text = "2"
    run()._r.append(_el("w:fldChar", fldCharType="end"))


def _finalise_sections(doc, page_number_start: int, has_cover: bool):
    """Mục 1: khung viền trang, đánh số trang, đầu trang riêng (trống) cho tờ bìa; đầu trang
    mặc định chứa số trang và được các mục sau kế thừa. Các mục sau: bỏ titlePg / pgNumType /
    pgBorders mà python-docx sao chép khi tạo section."""
    first = doc.sections[0]
    sectpr = first._sectPr
    if has_cover:
        borders = _el("w:pgBorders", offsetFrom="text")
        for side in ("top", "left", "bottom", "right"):
            borders.append(_el("w:" + side, val="single", sz=12, space=24, color="000000"))
        _set_child(sectpr, borders, _SECTPR_SEQ)
        first.different_first_page_header_footer = True
        first.first_page_header.is_linked_to_previous = False      # đầu trang trống cho bìa
    # Bìa là trang 1: trang 2 mang số 2 (start=1) hoặc số 1 (start=0). Không bìa: bắt đầu từ 1.
    start = 1 if (page_number_start == 2 or not has_cover) else 0
    _set_child(sectpr, _el("w:pgNumType", start=start), _SECTPR_SEQ)
    header = first.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.alignment = CENTER
    _style_paragraph_mark(p, 13)
    _add_page_field(p, 13)
    for section in doc.sections[1:]:
        section._sectPr.remove_all("w:pgBorders", "w:pgNumType", "w:titlePg")


def _normalise_options(options) -> dict:
    """Gộp với DEFAULT_OPTIONS và kiểm tra; sai → ConvertError (tiếng Việt)."""
    opts = dict(DEFAULT_OPTIONS)
    if options:
        unknown = sorted(set(options) - set(DEFAULT_OPTIONS))
        if unknown:
            raise ConvertError("Tuỳ chọn không hợp lệ: %s" % ", ".join(unknown))
        opts.update({k: v for k, v in options.items() if v is not None})

    page = str(opts["page_size"]).strip()
    page_size = {k.lower(): k for k in PAGE_SIZES}.get(page.lower())
    if page_size is None:
        raise ConvertError("Khổ giấy không hợp lệ: %r (chỉ nhận A4 hoặc Letter)." % page)

    try:
        start = float(opts["page_number_start"])
    except (TypeError, ValueError):
        start = float("nan")
    if start not in (1.0, 2.0):
        raise ConvertError("Cách đánh số trang không hợp lệ: %r (trang 2 mang số 1 hoặc 2)."
                           % (opts["page_number_start"],))

    try:
        font = float(opts["table_font_pt"])
    except (TypeError, ValueError):
        font = float("nan")
    if not 6 <= font <= 14:
        raise ConvertError("Cỡ chữ bảng không hợp lệ: %r (cho phép 6–14 pt)."
                           % (opts["table_font_pt"],))

    return {"page_size": page_size, "page_number_start": int(start), "table_font_pt": font}


_BUILDERS = (
    ("cover", _build_cover),
    ("abbreviations", _build_abbreviations),
    ("appendix", _build_appendix),
    ("main_table", _build_main_table),
    ("closing", _build_closing),
)


def build_document(data: WorkbookData, options=None):
    """Dựng tài liệu Word (docx.document.Document) từ WorkbookData theo tuỳ chọn."""
    opts = _normalise_options(options)
    doc = _new_document()
    section = doc.sections[0]
    first = True
    for part, build in _BUILDERS:
        content = getattr(data, part)
        if content is None:
            continue
        if not first:
            section = doc.add_section(WD_SECTION.NEW_PAGE)
        build(doc, section, content, opts["page_size"], opts["table_font_pt"])
        first = False
    _finalise_sections(doc, opts["page_number_start"], has_cover=data.cover is not None)
    return doc


def convert_workbook(src, options=None) -> bytes:
    """Excel (đường dẫn | bytes | đối tượng tệp) → nội dung tệp .docx (bytes)."""
    doc = build_document(read_workbook(src), options)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def output_name_for(xlsx_filename) -> str:
    """'ĐU Tam Đại Vĩnh viễn.xlsx' → 'ĐU Tam Đại Vĩnh viễn.docx' (chỉ tên tệp, bỏ thư mục)."""
    base = os.path.basename(os.fspath(xlsx_filename))
    stem, _ext = os.path.splitext(base)
    return (stem or base) + ".docx"


def convert_file(xlsx_path, out_path=None, options=None) -> str:
    """Chuyển một tệp Excel, ghi .docx và trả về đường dẫn đã ghi.
    out_path=None → thư mục 'ket_qua' cạnh tệp Excel; out_path là thư mục (hoặc không có đuôi
    .docx) → ghi vào thư mục đó với tên tệp Excel đổi đuôi .docx."""
    xlsx_path = os.fspath(xlsx_path)
    if out_path is None:
        out_path = os.path.join(os.path.dirname(os.path.abspath(xlsx_path)), "ket_qua")
    out_path = os.fspath(out_path)
    if os.path.isdir(out_path) or not out_path.lower().endswith(".docx"):
        out_path = os.path.join(out_path, output_name_for(xlsx_path))
    content = convert_workbook(xlsx_path, options)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "wb") as f:
        f.write(content)
    return out_path


# ---------------------------------------------------------------------------
# Dòng lệnh
# ---------------------------------------------------------------------------
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Chuyển MỤC LỤC HỒ SƠ từ Excel (.xlsx) sang Word (.docx) in được ngay.")
    ap.add_argument("xlsx", nargs="+", help="tệp Excel đầu vào (một hoặc nhiều)")
    ap.add_argument("-o", "--out-dir", help="thư mục xuất (mặc định: thư mục 'ket_qua' cạnh mỗi tệp)")
    ap.add_argument("--start", type=int, choices=(1, 2), default=DEFAULT_OPTIONS["page_number_start"],
                    help="số in trên trang 2: 2 = bìa là trang 1, không in số (mặc định); "
                         "1 = không tính bìa")
    ap.add_argument("--page", type=lambda s: s.strip().capitalize(), choices=tuple(PAGE_SIZES),
                    default=DEFAULT_OPTIONS["page_size"], help="khổ giấy nằm ngang: A4 (mặc định) hoặc Letter")
    ap.add_argument("--font", type=float, default=DEFAULT_OPTIONS["table_font_pt"],
                    help="cỡ chữ bảng thống kê hồ sơ, pt (mặc định 8.5; cho phép 6–14)")
    args = ap.parse_args(argv)
    options = {"page_size": args.page, "page_number_start": args.start, "table_font_pt": args.font}

    failed = 0
    for path in args.xlsx:
        try:
            written = convert_file(path, args.out_dir, options)
        except ConvertError as exc:
            failed += 1
            print("LỖI  %s: %s" % (path, exc), file=sys.stderr)
        else:
            print("XONG %s -> %s" % (path, written))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
