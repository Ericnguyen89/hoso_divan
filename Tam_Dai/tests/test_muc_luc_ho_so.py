# -*- coding: utf-8 -*-
"""
Kiểm thử cho muc_luc_ho_so.py — chỉ dùng unittest của thư viện chuẩn (không cần pytest).

Chạy toàn bộ:    cd "<thư mục dự án>" && python3 -m unittest -v
Chạy một lớp:    python3 -m unittest -v tests.test_muc_luc_ho_so.TestSections

Dữ liệu kiểm thử:
  * tệp mẫu "ĐU Tam Đại Vĩnh viễn.xlsx" trong thư mục Tam_Dai (các lớp dùng mẫu bị bỏ qua nếu thiếu);
  * các sổ Excel tổng hợp dựng ngay trong bộ nhớ bằng openpyxl (lớp TestSynthetic).
Kết quả .docx được kiểm tra ở mức XML (lxml trên word/document.xml và các header) chứ không
render bằng Word.
"""
from __future__ import annotations

import datetime as dt
import io
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
if PROJECT not in sys.path:
    sys.path.insert(0, PROJECT)

import openpyxl  # noqa: E402
from lxml import etree  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

import muc_luc_ho_so as m  # noqa: E402

SAMPLE_DIR = PROJECT   # tệp mẫu "ĐU Tam Đại Vĩnh viễn.xlsx/.docx" đặt ngay trong Tam_Dai/ (không đưa lên git)
SAMPLE_XLSX = os.path.join(SAMPLE_DIR, "ĐU Tam Đại Vĩnh viễn.xlsx")
SAMPLE_DOCX = os.path.join(SAMPLE_DIR, "ĐU Tam Đại Vĩnh viễn.docx")
HAS_SAMPLE = os.path.isfile(SAMPLE_XLSX)

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
NS = {"w": NSMAP["w"]}
INCH = 1440
A4_LANDSCAPE = ("16838", "11906")
LETTER_LANDSCAPE = ("15840", "12240")

# Thứ tự con của w:sectPr theo lược đồ OOXML (kiểm tra thứ tự phần tử)
SECTPR_ORDER = [
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
    "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
    "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
]

# Lề mong đợi từng mục (trên, phải, dưới, trái) — twips
EXPECTED_MARGINS = [
    (INCH, INCH, INCH, INCH),                      # 1 tờ bìa
    (INCH, INCH, INCH, INCH),                      # 2 chữ viết tắt
    (INCH, INCH // 2, INCH, INCH // 2),            # 3 phụ lục
    (INCH // 2, INCH // 2, INCH // 2, INCH // 2),  # 4 bảng thống kê
    (INCH, INCH, INCH, INCH),                      # 5 chứng từ kết thúc
]

MAIN_HEADERS = [
    "Số\ncặp\n(hộp)", "Phông\nsố", "Mục\nlục\nsố", "Số hồ sơ",
    "Tên nhóm và tên hồ sơ (đơn vị bảo quản)", "Thời gian bắt đầu và kết thúc",
    "Số\ntrang", "Số\ntài\nliệu", "Thời\nhạn\nbảo\nquản", "Độ mật", "Từ\nkhoá", "Ghi\nchú",
    "Tên thư mục hồ sơ",
]


# ---------------------------------------------------------------------------
# Tiện ích chung
# ---------------------------------------------------------------------------
def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return "{%s}%s" % (NSMAP[prefix], local)


def norm(text) -> str:
    """NFC, mọi khoảng trắng (kể cả xuống dòng) thành một dấu cách, bỏ khoảng trắng hai đầu."""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", text or "")).strip()


def norm_row(cells) -> list[str]:
    return [norm(c) for c in cells]


def xl_text(value) -> str:
    """Chuẩn hoá giá trị ô Excel giống như quy ước của mô-đun (để dựng dữ liệu mong đợi)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(value)
    if isinstance(value, (dt.datetime, dt.date)):
        return value.strftime("%d/%m/%Y")
    return norm(str(value))


def _is_int(value) -> bool:
    return (isinstance(value, int) and not isinstance(value, bool)) or (
        isinstance(value, float) and value.is_integer())


def p_text(p) -> str:
    """Chữ của một đoạn: w:t nối lại, w:br → '\\n', w:tab (trong run) → '\\t'."""
    buf = []
    for node in p.iter(qn("w:t"), qn("w:br"), qn("w:tab")):
        if node.tag == qn("w:t"):
            buf.append(node.text or "")
        elif node.getparent() is not None and node.getparent().tag == qn("w:r"):
            buf.append("\n" if node.tag == qn("w:br") else "\t")
    return "".join(buf)


def tc_text(tc) -> str:
    """Chữ của một ô bảng: các đoạn nối bằng '\\n'."""
    return "\n".join(p_text(p) for p in tc.findall("w:p", NS))


def row_texts(tr) -> list[str]:
    return [tc_text(tc) for tc in tr.findall("w:tc", NS)]


def table_rows(tbl):
    return tbl.findall("w:tr", NS)


def table_matrix(tbl) -> list[list[str]]:
    return [row_texts(tr) for tr in table_rows(tbl)]


def grid_cols(tbl) -> int:
    return len(tbl.findall("w:tblGrid/w:gridCol", NS))


def is_tbl_header(tr) -> bool:
    return tr.find("w:trPr/w:tblHeader", NS) is not None


def has_cant_split(tr) -> bool:
    return tr.find("w:trPr/w:cantSplit", NS) is not None


def runs_with_text(el):
    return [r for r in el.iter(qn("w:r"))
            if "".join(t.text or "" for t in r.findall("w:t", NS)).strip()]


def _rpr_flag(r, tag) -> bool:
    rpr = r.find("w:rPr", NS)
    el = rpr.find("w:" + tag, NS) if rpr is not None else None
    return el is not None and el.get(qn("w:val"), "1") not in ("0", "false", "off")


def run_bold(r) -> bool:
    return _rpr_flag(r, "b")


def run_color(r):
    el = r.find("w:rPr/w:color", NS)
    return None if el is None else el.get(qn("w:val"))


def run_size(r):
    el = r.find("w:rPr/w:sz", NS)
    return None if el is None else int(el.get(qn("w:val")))


def el_bold(el) -> bool:
    """Mọi run có chữ trong phần tử đều đậm (và có ít nhất một run)."""
    runs = runs_with_text(el)
    return bool(runs) and all(run_bold(r) for r in runs)


def el_any_bold(el) -> bool:
    return any(run_bold(r) for r in runs_with_text(el))


def para_jc(p):
    el = p.find("w:pPr/w:jc", NS)
    return None if el is None else el.get(qn("w:val"))


def first_cell_jc(tc):
    return para_jc(tc.find("w:p", NS))


class DocxView:
    """Nhìn vào bên trong một tệp .docx (bytes) bằng lxml."""

    def __init__(self, data: bytes):
        self.data = data
        self.zip = zipfile.ZipFile(io.BytesIO(data))
        self.root = etree.fromstring(self.zip.read("word/document.xml"))
        self.body = self.root.find("w:body", NS)
        rels = etree.fromstring(self.zip.read("word/_rels/document.xml.rels"))
        self.rels = {rel.get("Id"): rel.get("Target") for rel in rels}

    @property
    def sectprs(self):
        """Mọi w:sectPr theo thứ tự tài liệu (= thứ tự section)."""
        return self.root.findall(".//w:sectPr", NS)

    @property
    def tables(self):
        return self.body.findall("w:tbl", NS)

    def section_blocks(self, i: int):
        """Các phần tử con của w:body thuộc section thứ i (0-based)."""
        blocks, idx = [], 0
        for child in self.body:
            if child.tag == qn("w:sectPr"):
                break
            if idx == i:
                blocks.append(child)
            if child.tag == qn("w:p") and child.find("w:pPr/w:sectPr", NS) is not None:
                idx += 1
        return blocks

    def section_text(self, i: int) -> str:
        return norm(" ".join(p_text(p) for blk in self.section_blocks(i) for p in blk.iter(qn("w:p"))))

    @property
    def all_text(self) -> str:
        return norm(" ".join(p_text(p) for p in self.root.iter(qn("w:p"))))

    def part_root(self, rid: str):
        return etree.fromstring(self.zip.read("word/" + self.rels[rid]))

    def header_refs(self, sect) -> dict:
        return {h.get(qn("w:type")): h.get(qn("r:id")) for h in sect.findall("w:headerReference", NS)}

    def main_table(self):
        found = [t for t in self.tables if grid_cols(t) == 13]
        assert len(found) == 1, "mong đợi đúng một bảng 13 cột, thấy %d" % len(found)
        return found[0]

    def table_by_first_cell(self, text: str, ncols: int):
        for t in self.tables:
            rows = table_rows(t)
            if grid_cols(t) == ncols and rows and norm(row_texts(rows[0])[0]) == text:
                return t
        raise AssertionError("không thấy bảng %d cột bắt đầu bằng %r" % (ncols, text))


# ---------------------------------------------------------------------------
# Dữ liệu dùng chung (chuyển mẫu một lần cho cả bộ kiểm thử)
# ---------------------------------------------------------------------------
_CACHE: dict = {}


def sample_data() -> "m.WorkbookData":
    if "data" not in _CACHE:
        _CACHE["data"] = m.read_workbook(SAMPLE_XLSX)
    return _CACHE["data"]


def sample_view() -> DocxView:
    if "view" not in _CACHE:
        _CACHE["view"] = DocxView(m.convert_workbook(SAMPLE_XLSX))
    return _CACHE["view"]


def sample_wb():
    if "wb" not in _CACHE:
        _CACHE["wb"] = openpyxl.load_workbook(SAMPLE_XLSX, data_only=True)
    return _CACHE["wb"]


def build_view(data, options) -> DocxView:
    doc = m.build_document(data, options)
    buf = io.BytesIO()
    doc.save(buf)
    return DocxView(buf.getvalue())


def find_sheet(wb, needle: str):
    """Sheet đầu tiên có ô chứa `needle` (không phân biệt hoa thường) trong 60 dòng × 20 cột đầu."""
    needle = needle.lower()
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 60), max_col=20):
            for c in row:
                if isinstance(c.value, str) and needle in norm(c.value).lower():
                    return ws
    raise AssertionError("không thấy sheet chứa %r" % needle)


def excel_table(ws, header_pred, pad_roles=False, drop=None):
    """Dựng (headers, numbering, rows) mong đợi từ một sheet Excel theo quy ước của SPEC."""
    rows = [[c for c in row if xl_text(c.value)] for row in ws.iter_rows()]
    rows = [r for r in rows if r]
    hi = next(i for i, r in enumerate(rows) if header_pred([xl_text(c.value) for c in r]))
    cols = [c.column for c in rows[hi]]
    headers = [xl_text(c.value) for c in rows[hi]]
    pad = {}
    if pad_roles:
        for j, h in enumerate(headers):
            k = h.lower()
            for n, pats in ((4, ("số hồ sơ",)), (3, ("phông",)), (2, ("mục lục",)), (2, ("cặp", "hộp"))):
                if any(p in k for p in pats):
                    pad[j] = n
                    break
    body = hi + 1
    numbering = None
    if body < len(rows) and all(re.match(r"^\(\d+\)$", xl_text(c.value)) for c in rows[body]):
        by = {c.column: xl_text(c.value) for c in rows[body]}
        numbering = [by.get(col, "") for col in cols]
        body += 1
    out = []
    for r in rows[body:]:
        by = {c.column: c for c in r}
        texts = []
        for j, col in enumerate(cols):
            c = by.get(col)
            v = c.value if c is not None else None
            t = xl_text(v)
            if j in pad and _is_int(v):
                t = t.zfill(pad[j])
            texts.append(t)
        if any(texts) and not (drop and drop(texts)):
            out.append(texts)
    return headers, numbering, out


class RowsMixin:
    def assertRowsEqual(self, expected, got, label):
        """So sánh hai danh sách dòng; báo dòng khác đầu tiên (mong đợi / nhận được)."""
        for i, (e, g) in enumerate(zip(expected, got)):
            if e != g:
                self.fail("%s: dòng %d khác nhau\n  mong đợi: %r\n  nhận được: %r"
                          % (label, i, e, g))
        self.assertEqual(len(expected), len(got),
                         "%s: số dòng khác nhau (mong đợi %d, nhận được %d); dòng thừa/thiếu đầu tiên: %r"
                         % (label, len(expected), len(got),
                            (expected + got)[min(len(expected), len(got))]))


# ---------------------------------------------------------------------------
# 1. API cơ bản
# ---------------------------------------------------------------------------
class TestBasics(unittest.TestCase):
    def test_default_options(self):
        self.assertEqual(m.DEFAULT_OPTIONS,
                         {"page_size": "A4", "page_number_start": 2, "table_font_pt": 8.5})

    def test_output_name_for(self):
        self.assertEqual(m.output_name_for("ĐU Tam Đại Vĩnh viễn.xlsx"), "ĐU Tam Đại Vĩnh viễn.docx")
        self.assertEqual(m.output_name_for("/a/b/Phông 2 - Vĩnh viễn.XLSX"), "Phông 2 - Vĩnh viễn.docx")
        self.assertEqual(m.output_name_for("x.xlsm"), "x.docx")
        self.assertEqual(m.output_name_for(pathlib.Path("thu muc/tệp.xlsx")), "tệp.docx")

    @unittest.skipUnless(HAS_SAMPLE, "thiếu tệp mẫu")
    def test_convert_returns_docx_bytes(self):
        data = sample_view().data
        self.assertIsInstance(data, bytes)
        self.assertTrue(data.startswith(b"PK"))
        names = zipfile.ZipFile(io.BytesIO(data)).namelist()
        self.assertIn("word/document.xml", names)
        self.assertIn("[Content_Types].xml", names)

    @unittest.skipUnless(HAS_SAMPLE, "thiếu tệp mẫu")
    def test_read_workbook_accepts_bytes_and_filelike(self):
        with open(SAMPLE_XLSX, "rb") as f:
            raw = f.read()
        d1 = m.read_workbook(raw)
        with open(SAMPLE_XLSX, "rb") as f:
            d2 = m.read_workbook(f)
        self.assertEqual(len(d1.main_table.rows), len(sample_data().main_table.rows))
        self.assertEqual(len(d2.main_table.rows), len(sample_data().main_table.rows))
        self.assertEqual(d1.sheet_names["main_table"], "ĐU Tam Đại VV")

    @unittest.skipUnless(HAS_SAMPLE, "thiếu tệp mẫu")
    def test_convert_file_default_dir_is_ket_qua_next_to_input(self):
        with tempfile.TemporaryDirectory() as tmp:
            copy = os.path.join(tmp, "ĐU Tam Đại Vĩnh viễn.xlsx")
            shutil.copyfile(SAMPLE_XLSX, copy)
            written = m.convert_file(copy)
            self.assertEqual(written, os.path.join(tmp, "ket_qua", "ĐU Tam Đại Vĩnh viễn.docx"))
            self.assertTrue(os.path.isfile(written))
            with open(written, "rb") as f:
                self.assertTrue(f.read(2) == b"PK")
            self.assertFalse(os.path.exists(os.path.join(tmp, "ĐU Tam Đại Vĩnh viễn.docx")),
                             "không được ghi .docx cạnh tệp Excel khi không chỉ định thư mục")


# ---------------------------------------------------------------------------
# 2. Section, lề, đầu trang, tuỳ chọn
# ---------------------------------------------------------------------------
@unittest.skipUnless(HAS_SAMPLE, "thiếu tệp mẫu")
class TestSections(unittest.TestCase):
    def test_five_landscape_a4_sections(self):
        sects = sample_view().sectprs
        self.assertEqual(len(sects), 5)
        for i, s in enumerate(sects):
            pgsz = s.find("w:pgSz", NS)
            self.assertIsNotNone(pgsz, "section %d thiếu pgSz" % (i + 1))
            self.assertEqual((pgsz.get(qn("w:w")), pgsz.get(qn("w:h"))), A4_LANDSCAPE, "section %d" % (i + 1))
            self.assertEqual(pgsz.get(qn("w:orient")), "landscape", "section %d" % (i + 1))

    def test_section1_titlepg_pgnum_borders(self):
        s = sample_view().sectprs[0]
        self.assertIsNotNone(s.find("w:titlePg", NS))
        pgnum = s.find("w:pgNumType", NS)
        self.assertIsNotNone(pgnum)
        self.assertEqual(pgnum.get(qn("w:start")), "1")
        borders = s.find("w:pgBorders", NS)
        self.assertIsNotNone(borders)
        self.assertEqual(borders.get(qn("w:offsetFrom")), "text")
        sides = {etree.QName(b).localname: b for b in borders}
        self.assertEqual(set(sides), {"top", "left", "bottom", "right"})
        for b in sides.values():
            self.assertEqual(b.get(qn("w:val")), "single")
            self.assertEqual(b.get(qn("w:sz")), "12")

    def test_sections_2_to_5_inherit(self):
        for i, s in enumerate(sample_view().sectprs[1:], start=2):
            for tag in ("titlePg", "pgNumType", "pgBorders", "headerReference", "footerReference"):
                self.assertIsNone(s.find("w:" + tag, NS), "section %d không được có %s" % (i, tag))

    def test_sectpr_child_order(self):
        for i, s in enumerate(sample_view().sectprs, start=1):
            names = [etree.QName(c).localname for c in s]
            idx = [SECTPR_ORDER.index(n) for n in names]
            self.assertEqual(idx, sorted(idx), "section %d: thứ tự con sectPr sai: %s" % (i, names))

    def test_margins_per_section(self):
        for i, (s, exp) in enumerate(zip(sample_view().sectprs, EXPECTED_MARGINS), start=1):
            mar = s.find("w:pgMar", NS)
            got = tuple(int(mar.get(qn("w:" + k))) for k in ("top", "right", "bottom", "left"))
            self.assertEqual(got, exp, "section %d lề (trên, phải, dưới, trái)" % i)
            self.assertEqual(mar.get(qn("w:header")), "708", "section %d" % i)
            self.assertEqual(mar.get(qn("w:footer")), "708", "section %d" % i)

    def test_headers_page_field(self):
        view = sample_view()
        refs = view.header_refs(view.sectprs[0])
        self.assertEqual(set(refs), {"first", "default"})
        default = view.part_root(refs["default"])
        instr = [t.text for t in default.iter(qn("w:instrText"))]
        self.assertEqual(instr, [" PAGE "])
        fld = [f.get(qn("w:fldCharType")) for f in default.iter(qn("w:fldChar"))]
        self.assertEqual(fld, ["begin", "separate", "end"])
        p = default.find(".//w:p", NS)
        self.assertEqual(para_jc(p), "center")
        sizes = {run_size(r) for r in default.iter(qn("w:r"))}
        self.assertEqual(sizes, {26}, "số trang phải 13pt (sz=26)")
        for rf in default.iter(qn("w:rFonts")):
            self.assertEqual(rf.get(qn("w:ascii")), "Times New Roman")
        first = view.part_root(refs["first"])
        self.assertEqual(list(first.iter(qn("w:instrText"))), [])
        self.assertEqual(list(first.iter(qn("w:fldChar"))), [])
        self.assertEqual(norm(" ".join(p_text(p) for p in first.iter(qn("w:p")))), "")

    def test_option_page_number_start_1(self):
        view = build_view(sample_data(), {"page_number_start": 1})
        self.assertEqual(view.sectprs[0].find("w:pgNumType", NS).get(qn("w:start")), "0")
        for s in view.sectprs[1:]:
            self.assertIsNone(s.find("w:pgNumType", NS))

    def test_option_page_size_letter(self):
        view = build_view(sample_data(), {"page_size": "Letter"})
        self.assertEqual(len(view.sectprs), 5)
        for s in view.sectprs:
            pgsz = s.find("w:pgSz", NS)
            self.assertEqual((pgsz.get(qn("w:w")), pgsz.get(qn("w:h"))), LETTER_LANDSCAPE)
            self.assertEqual(pgsz.get(qn("w:orient")), "landscape")

    def test_option_table_font(self):
        view = build_view(sample_data(), {"table_font_pt": 10})
        tbl = view.main_table()
        rows = table_rows(tbl)
        self.assertEqual({run_size(r) for r in runs_with_text(rows[0])}, {17}, "tiêu đề = 10 − 1.5 = 8.5pt")
        self.assertEqual({run_size(r) for r in runs_with_text(rows[2])}, {20}, "thân bảng 10pt")

    def test_invalid_options_raise_convert_error(self):
        data = sample_data()
        for bad in ({"page_size": "A3"}, {"page_size": ""}, {"page_number_start": 3},
                    {"page_number_start": "x"}, {"table_font_pt": 5}, {"table_font_pt": 14.5},
                    {"table_font_pt": "abc"}, {"khong_co": 1}):
            with self.assertRaises(m.ConvertError, msg="tuỳ chọn %r phải bị từ chối" % bad):
                m.build_document(data, bad)
        with self.assertRaises(m.ConvertError):
            m.convert_workbook(SAMPLE_XLSX, {"page_size": "A3"})

    def test_valid_option_spellings(self):
        # chuỗi số và chữ thường vẫn được chấp nhận (app.py / CLI có thể truyền kiểu khác nhau)
        view = build_view(sample_data(), {"page_size": "letter", "page_number_start": "1", "table_font_pt": "9"})
        pgsz = view.sectprs[0].find("w:pgSz", NS)
        self.assertEqual((pgsz.get(qn("w:w")), pgsz.get(qn("w:h"))), LETTER_LANDSCAPE)
        self.assertEqual(view.sectprs[0].find("w:pgNumType", NS).get(qn("w:start")), "0")


# ---------------------------------------------------------------------------
# 3. Các bảng của tệp mẫu
# ---------------------------------------------------------------------------
@unittest.skipUnless(HAS_SAMPLE, "thiếu tệp mẫu")
class TestSampleTables(RowsMixin, unittest.TestCase):
    def test_table_count_and_order(self):
        tables = sample_view().tables
        self.assertEqual([grid_cols(t) for t in tables], [2, 4, 4, 13, 2],
                         "bìa (bảng nổi 2 cột), chữ viết tắt, phụ lục, bảng chính, khối chữ ký")

    def test_abbreviations_table(self):
        tbl = sample_view().table_by_first_cell("SỐ TT", 4)
        rows = table_rows(tbl)
        self.assertEqual(len(rows), 1 + 14)
        self.assertTrue(is_tbl_header(rows[0]))
        self.assertFalse(any(is_tbl_header(r) for r in rows[1:]))
        self.assertTrue(all(has_cant_split(r) for r in rows[1:]))
        mat = [norm_row(r) for r in table_matrix(tbl)]
        self.assertEqual(mat[0], ["SỐ TT", "CHỮ VIẾT TẮT", "NGHĨA LÀ", "GHI CHÚ"])
        self.assertEqual(mat[1], ["1", "BCH", "Ban Chấp hành", ""])
        self.assertEqual(mat[-1], ["14", "UBND", "Ủy ban Nhân dân", ""])
        self.assertNotIn("15", [r[0] for r in mat], "dòng chỉ có số TT 15 phải bị bỏ")
        self.assertTrue(el_bold(rows[0]))
        self.assertEqual(first_cell_jc(rows[1].findall("w:tc", NS)[0]), "center")
        self.assertEqual(first_cell_jc(rows[1].findall("w:tc", NS)[1]), "left")

    def test_abbreviations_content_equals_excel(self):
        ws = find_sheet(sample_wb(), "chữ viết tắt")
        _, _, expected = excel_table(ws, lambda h: any("chữ viết tắt" in x.lower() for x in h),
                                     drop=lambda t: not any(t[1:]))
        got = [norm_row(r) for r in table_matrix(sample_view().table_by_first_cell("SỐ TT", 4))[1:]]
        self.assertRowsEqual([norm_row(r) for r in expected], got, "bảng chữ viết tắt")

    def test_appendix_table(self):
        tbl = sample_view().table_by_first_cell("Đề mục", 4)
        rows = table_rows(tbl)
        self.assertEqual(len(rows), 2 + 37)
        self.assertTrue(is_tbl_header(rows[0]) and is_tbl_header(rows[1]))
        self.assertFalse(any(is_tbl_header(r) for r in rows[2:]))
        mat = [norm_row(r) for r in table_matrix(tbl)]
        self.assertEqual(mat[0], ["Đề mục", "Tiêu đề nhóm", "Trang số", "Ghi chú"])
        self.assertEqual(mat[1], ["(1)", "(2)", "(3)", "(4)"])
        self.assertEqual(mat[2], ["I.", "ĐẢNG BỘ XÃ TAM ĐẠI, KHÓA XIX, NHIỆM KỲ 2005-2010", "", ""])
        self.assertEqual(mat[-1], ["02.02.02.", "Hồ sơ do UBKT Đảng ủy xã Tam Đại kiểm tra, giám sát", "", ""])
        self.assertTrue(el_bold(rows[2]), "dòng 'I.' đậm theo Excel")
        self.assertFalse(el_any_bold(rows[3]), "dòng '01.' thường theo Excel")

    def test_appendix_content_equals_excel(self):
        ws = find_sheet(sample_wb(), "phụ lục tra cứu")
        _, numbering, expected = excel_table(ws, lambda h: any("đề mục" in x.lower() for x in h))
        mat = table_matrix(sample_view().table_by_first_cell("Đề mục", 4))
        self.assertEqual(norm_row(mat[1]), norm_row(numbering))
        self.assertRowsEqual([norm_row(r) for r in expected], [norm_row(r) for r in mat[2:]], "phụ lục tra cứu")

    def test_main_table_shape(self):
        tbl = sample_view().main_table()
        rows = table_rows(tbl)
        self.assertEqual(grid_cols(tbl), 13)
        self.assertEqual(len(rows), 2 + 37 + 229)
        self.assertTrue(is_tbl_header(rows[0]) and is_tbl_header(rows[1]))
        self.assertFalse(any(is_tbl_header(r) for r in rows[2:]))
        self.assertTrue(all(has_cant_split(r) for r in rows[2:]))
        for i, r in enumerate(rows):
            self.assertEqual(len(r.findall("w:tc", NS)), 13, "dòng %d" % i)
        layout = tbl.find("w:tblPr/w:tblLayout", NS)
        self.assertIsNotNone(layout)
        self.assertEqual(layout.get(qn("w:type")), "fixed")

    def test_main_header_rows(self):
        rows = table_rows(sample_view().main_table())
        self.assertEqual(row_texts(rows[0]), MAIN_HEADERS, "giữ xuống dòng trong ô tiêu đề")
        self.assertEqual(len(list(rows[0].findall("w:tc", NS)[0].iter(qn("w:br")))), 2)
        self.assertEqual(row_texts(rows[1]),
                         ["(1)", "(2)", "(3)", "(4)", "(5)", "(6)", "(7)", "(8)", "(9)", "(10)", "(12)", "(13)", ""])
        self.assertTrue(el_bold(rows[0]))
        self.assertEqual({run_size(r) for r in runs_with_text(rows[0])}, {14}, "tiêu đề 7pt = 8.5 − 1.5")

    def test_row_0001(self):
        rows = table_rows(sample_view().main_table())
        body = [(r, norm_row(row_texts(r))) for r in rows[2:]]
        hits = [(r, t) for r, t in body if t[3] == "0001"]
        self.assertEqual(len(hits), 1)
        tr, texts = hits[0]
        self.assertTrue(texts[4].startswith("Nghị quyết của Đảng ủy xã Tam Đại"), texts[4])
        self.assertEqual(texts[:4] + texts[5:],
                         ["01", "001", "01", "0001", "16/02/2006-19/02/2008", "61", "16", "Vĩnh viễn", "", "", "", ""])
        tcs = tr.findall("w:tc", NS)
        for j in range(4):
            self.assertTrue(el_bold(tcs[j]), "cột %d của dòng hồ sơ phải đậm" % (j + 1))
        self.assertFalse(el_any_bold(tcs[4]), "tên hồ sơ không đậm")
        self.assertEqual(first_cell_jc(tcs[4]), "both", "tên hồ sơ căn đều")
        self.assertEqual(first_cell_jc(tcs[3]), "center")
        self.assertEqual({run_size(r) for r in runs_with_text(tr)}, {17}, "thân bảng 8.5pt")

    # dòng tên nhóm duy nhất của Excel mẫu còn có giá trị ở cột khác ngoài Tên nhóm (B113='001', C113='01')
    STRAY_GROUP = "01.02.01.03. Nhiệm kỳ 2019-2022"

    def test_group_rows(self):
        rows = table_rows(sample_view().main_table())
        groups = [r for r in rows[2:] if not norm(row_texts(r)[3]) and norm(row_texts(r)[4])]
        self.assertEqual(len(groups), 37)
        first = norm_row(row_texts(groups[0]))
        self.assertEqual(first[4], "I. ĐẢNG BỘ XÃ TAM ĐẠI, KHÓA XIX, NHIỆM KỲ 2005-2010")
        others = [r for r in groups if norm(row_texts(r)[4]) != self.STRAY_GROUP]
        self.assertEqual(len(others), 36)
        for r in others:
            texts = norm_row(row_texts(r))
            self.assertEqual([t for j, t in enumerate(texts) if j != 4], [""] * 12, texts)
        for r in groups:
            texts = norm_row(row_texts(r))
            tc = r.findall("w:tc", NS)[4]
            self.assertTrue(el_bold(tc), "tên nhóm phải đậm: %r" % texts[4])
            self.assertEqual(first_cell_jc(tc), "left", "tên nhóm căn trái: %r" % texts[4])

    def test_last_row_odd_ids_and_int_cells(self):
        body = [norm_row(r) for r in table_matrix(sample_view().main_table())[2:]]
        self.assertEqual(body[-1][3], "0227")
        ids = [r[3] for r in body if r[3]]
        self.assertEqual(len(ids), 229)
        for odd in ("0014A", "0078a", "0169a"):
            self.assertIn(odd, ids)
        row2 = next(r for r in body if r[3] == "0002")
        self.assertEqual(row2[6:8], ["27", "17"], "G11/H11 là số nguyên trong Excel → '27', '17'")

    def test_group_row_with_stray_values_in_excel(self):
        """Excel mẫu: dòng tên nhóm '01.02.01.03. Nhiệm kỳ 2019-2022' (không có Số hồ sơ) còn có
        '001'/'01' ở cột B/C. Quyết định: dòng tên nhóm in TRUNG THỰC theo Excel → hai giá trị đó
        phải có mặt (đậm theo Excel, căn giữa); ô Tên nhóm vẫn đậm + căn trái."""
        rows = table_rows(sample_view().main_table())
        hits = [r for r in rows[2:] if norm(row_texts(r)[4]) == self.STRAY_GROUP]
        self.assertEqual(len(hits), 1)
        tr = hits[0]
        self.assertEqual(norm_row(row_texts(tr)), ["", "001", "01", "", self.STRAY_GROUP] + [""] * 8)
        tcs = tr.findall("w:tc", NS)
        self.assertTrue(el_bold(tcs[4]))
        self.assertEqual(first_cell_jc(tcs[4]), "left")
        ws = find_sheet(sample_wb(), "bảng thống kê hồ sơ")
        xl_row = next(r for r in ws.iter_rows() if xl_text(r[4].value) == self.STRAY_GROUP)
        for j in (1, 2):
            self.assertEqual(el_bold(tcs[j]), bool(xl_row[j].font.b), "đậm theo Excel ở cột %d" % (j + 1))
            self.assertEqual(first_cell_jc(tcs[j]), "center")

    def test_main_full_content_equals_excel(self):
        ws = find_sheet(sample_wb(), "bảng thống kê hồ sơ")
        headers, numbering, expected = excel_table(
            ws, lambda h: any("số hồ sơ" in x.lower() for x in h) and any("tên nhóm" in x.lower() for x in h),
            pad_roles=True)
        mat = table_matrix(sample_view().main_table())
        self.assertEqual(norm_row(mat[0]), norm_row(headers))
        self.assertEqual(norm_row(mat[1]), norm_row(numbering))
        self.assertRowsEqual([norm_row(r) for r in expected], [norm_row(r) for r in mat[2:]], "bảng thống kê hồ sơ")


# ---------------------------------------------------------------------------
# 4. Nội dung bìa, chứng từ kết thúc, font
# ---------------------------------------------------------------------------
@unittest.skipUnless(HAS_SAMPLE, "thiếu tệp mẫu")
class TestSampleContent(unittest.TestCase):
    def test_cover_texts(self):
        text = sample_view().section_text(0)
        for needle in ("ĐẢNG CỘNG SẢN VIỆT NAM", "ĐẢNG ỦY XÃ PHÚ NINH", "MỤC LỤC HỒ SƠ",
                       "PHÔNG LƯU TRỮ CƠ QUAN LÃNH ĐẠO ĐẢNG XÃ TAM ĐẠI",
                       "KHÓA XIX, NHIỆM KỲ 2005-2010; KHÓA XX, NHIỆM KỲ 2010-2015; "
                       "KHÓA XXI, NHIỆM KỲ 2015-2020; KHÓA XXII, NHIỆM KỲ 2020-2025",
                       "Phông số: 001", "Mục lục số: 01", "Số trang:", "Thời hạn bảo quản",
                       "Vĩnh viễn", "Phú Ninh, 7/2026"):
            self.assertIn(needle, text)

    def test_cover_org_line_red_bold_and_sizes(self):
        view = sample_view()
        runs = {norm(p_text(r)): r for r in view.root.iter(qn("w:r")) if norm(p_text(r))}
        org = runs["ĐẢNG ỦY XÃ PHÚ NINH"]
        self.assertTrue(run_bold(org))
        self.assertEqual(run_color(org), "FF0000")
        self.assertEqual(run_size(org), 36)
        party = runs["ĐẢNG CỘNG SẢN VIỆT NAM"]
        self.assertFalse(run_bold(party))
        self.assertIsNone(run_color(party))
        title = runs["MỤC LỤC HỒ SƠ"]
        self.assertTrue(run_bold(title))
        self.assertEqual(run_size(title), 44)

    def test_cover_bottom_block_is_floating_table(self):
        tbl = sample_view().tables[0]
        tblp = tbl.find("w:tblPr/w:tblpPr", NS)
        self.assertIsNotNone(tblp)
        self.assertEqual(tblp.get(qn("w:vertAnchor")), "margin")
        self.assertEqual(tblp.get(qn("w:tblpYSpec")), "bottom")
        self.assertIsNotNone(tbl.find("w:tblPr/w:tblOverlap", NS))
        mat = [norm_row(r) for r in table_matrix(tbl)]
        self.assertEqual(mat[0], ["Phông số: 001", "Thời hạn bảo quản"])
        self.assertEqual(mat[1], ["Mục lục số: 01", "Vĩnh viễn"])
        self.assertEqual(mat[2], ["Số trang:", ""])
        self.assertEqual(mat[3], ["Phú Ninh, 7/2026"])
        last = table_rows(tbl)[-1].find("w:tc", NS)
        self.assertEqual(last.find("w:tcPr/w:gridSpan", NS).get(qn("w:val")), "2")
        # giá trị sau dấu ':' in đậm, nhãn thường
        tc = table_rows(tbl)[0].find("w:tc", NS)
        bold_by_text = {norm(p_text(r)): run_bold(r) for r in runs_with_text(tc)}
        self.assertEqual(bold_by_text, {"Phông số:": False, "001": True})
        self.assertTrue(el_bold(table_rows(tbl)[1].findall("w:tc", NS)[1]), "'Vĩnh viễn' đậm theo Excel")

    def test_closing_texts(self):
        view = sample_view()
        text = view.section_text(4)
        for needle in ("CHỨNG TỪ KẾT THÚC",
                       "Mục lục này gồm: 229 hồ sơ (đơn vị bảo quản)",
                       "Viết bằng chữ: Bốn trăm hai mươi hai (đơn vị bảo quản)",
                       "Đánh số từ hồ sơ (đơn vị bảo quản) số 0001 đến hồ sơ (đơn vị bảo quản) số 0227",
                       "Trùng hồ sơ (đơn vị bảo quản): 03 (0014a, 0078a, 0169a)",
                       "Khuyết hồ sơ (đơn vị bảo quản): 01 (0067)",
                       "Trong đó có 229 hồ sơ (ĐVBQ) vĩnh viễn",
                       "Mục lục này gồm: trang. Được đánh số liên tục từ 01 đến",
                       "PHÊ DUYỆT", "CHÁNH VĂN PHÒNG", "Người lập mục lục hồ sơ",
                       "Phú Ninh, Ngày ….... tháng 7 năm 2026"):
            self.assertIn(needle, text)
        runs = {norm(p_text(r)): r for r in view.root.iter(qn("w:r")) if norm(p_text(r))}
        self.assertTrue(run_bold(runs["PHÊ DUYỆT"]))
        self.assertTrue(run_bold(runs["CHÁNH VĂN PHÒNG"]))
        self.assertFalse(run_bold(runs["Người lập mục lục hồ sơ"]))
        sig = view.tables[-1]
        self.assertEqual(grid_cols(sig), 2)
        left, right = table_rows(sig)[0].findall("w:tc", NS)
        self.assertEqual([norm(p_text(p)) for p in left.findall("w:p", NS)], ["PHÊ DUYỆT", "CHÁNH VĂN PHÒNG"])
        self.assertEqual([norm(p_text(p)) for p in right.findall("w:p", NS)],
                         ["Phú Ninh, Ngày ….... tháng 7 năm 2026", "Người lập mục lục hồ sơ"])

    def test_sections_hold_expected_parts(self):
        view = sample_view()
        self.assertIn("CHỮ VIẾT TẮT", view.section_text(1))
        self.assertIn("PHỤ LỤC TRA CỨU", view.section_text(2))
        self.assertIn("BẢNG THỐNG KÊ HỒ SƠ", view.section_text(3))
        self.assertIn("CHỨNG TỪ KẾT THÚC", view.section_text(4))

    def test_bien_muc_sheet_ignored(self):
        text = sample_view().all_text
        self.assertNotIn("MỤC LỤC VĂN BẢN, TÀI LIỆU", text)
        self.assertNotIn("03a-NQ/ĐU", text)
        self.assertNotIn("Mã hồ sơ:", text)

    def test_fonts_times_new_roman(self):
        view = sample_view()
        for rf in view.root.iter(qn("w:rFonts")):
            for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                self.assertEqual(rf.get(qn(attr)), "Times New Roman", etree.tostring(rf))
        styles = etree.fromstring(view.zip.read("word/styles.xml"))
        normal = next(s for s in styles.iter(qn("w:style")) if s.get(qn("w:styleId")) == "Normal")
        self.assertEqual(normal.find("w:rPr/w:sz", NS).get(qn("w:val")), "26")
        self.assertEqual(normal.find("w:rPr/w:rFonts", NS).get(qn("w:ascii")), "Times New Roman")


# ---------------------------------------------------------------------------
# 5. Sổ Excel tổng hợp (robustness)
# ---------------------------------------------------------------------------
DEFAULT_NAMES = {"cover": "Tờ bìa", "abbr": "Bảng chữ cái viết tắt", "appendix": "Phụ lục tra cứu",
                 "main": "Phông X VV", "closing": "Chứng từ kết thúc"}
ALL_PARTS = ("cover", "abbr", "appendix", "main", "closing")
DEFAULT_MAIN_ROWS = [
    {"E": "I. NHÓM THỨ NHẤT\nDÒNG THỨ HAI", "bold": True},
    {"A": 1, "B": 1, "C": 1, "D": 7, "E": "Hồ sơ A\nDòng hai", "F": dt.datetime(2024, 5, 3),
     "G": 27, "H": 3, "I": "Vĩnh viễn"},
    {"B": "001", "C": "01", "D": "0008", "E": "  Hồ sơ B  ", "F": "01/01/2020-02/02/2021",
     "G": "5", "H": "2", "I": "Vĩnh viễn"},
]


def _set(ws, ref, value, bold=False, color=None):
    c = ws[ref]
    c.value = value
    if bold or color:
        c.font = Font(bold=bold, color=color)
    return c


def _fill_cover(ws):
    _set(ws, "A1", "ĐẢNG CỘNG SẢN VIỆT NAM")
    _set(ws, "A2", "ĐẢNG ỦY XÃ THỬ NGHIỆM", bold=True, color="FFFF0000")
    _set(ws, "A5", "MỤC LỤC HỒ SƠ", bold=True)
    _set(ws, "A7", "PHÔNG LƯU TRỮ THỬ NGHIỆM", bold=True)
    _set(ws, "A8", "KHÓA I, NHIỆM KỲ 2020-2025", bold=True)
    _set(ws, "A10", "   Phông số: 002")
    _set(ws, "L10", "Thời hạn bảo quản")
    _set(ws, "A11", "   Mục lục số: 03")
    _set(ws, "L11", "Vĩnh viễn", bold=True)
    _set(ws, "A12", "   Số trang:  ")
    _set(ws, "A13", "Nơi Thử, 1/2026", bold=True)


def _fill_abbr(ws):
    _set(ws, "A1", "PHÔNG LƯU TRỮ THỬ NGHIỆM", bold=True)
    for col, h in zip("ABCD", ("SỐ TT", "CHỮ VIẾT TẮT", "NGHĨA LÀ", "GHI CHÚ")):
        _set(ws, col + "3", h, bold=True)
    _set(ws, "A4", 1)
    _set(ws, "B4", "BCH")
    _set(ws, "C4", "Ban Chấp hành")
    _set(ws, "A5", 2)                      # chỉ có số TT → phải bị bỏ


def _fill_appendix(ws):
    _set(ws, "A1", "PHỤ LỤC TRA CỨU", bold=True)
    for col, h in zip("ABCD", ("Đề mục", "Tiêu đề nhóm", "Trang số", "Ghi chú")):
        _set(ws, col + "2", h, bold=True)
    for col, n in zip("ABCD", ("(1)", "(2)", "(3)", "(4)")):
        _set(ws, col + "3", n)
    _set(ws, "A4", "I.", bold=True)
    _set(ws, "B4", "NHÓM THỨ NHẤT", bold=True)
    _set(ws, "A5", "01.")
    _set(ws, "B5", "Nhóm con")


def _fill_main(ws, rows, ncols=13):
    _set(ws, "E2", "BẢNG THỐNG KÊ HỒ SƠ", bold=True)
    _set(ws, "E3", "PHÔNG LƯU TRỮ THỬ NGHIỆM", bold=True)
    for j, h in enumerate(MAIN_HEADERS[:ncols]):
        ws.cell(row=5, column=j + 1, value=h).font = Font(bold=True)
    for j in range(ncols):
        ws.cell(row=6, column=j + 1, value="(%d)" % (j + 1))
    for i, spec in enumerate(rows):
        for letter, value in spec.items():
            if letter == "bold":
                continue
            ci = openpyxl.utils.column_index_from_string(letter)
            if ci > ncols:
                continue
            c = ws.cell(row=7 + i, column=ci, value=value)
            if spec.get("bold") or letter in "ABCD":
                c.font = Font(bold=True)


def _fill_closing(ws):
    _set(ws, "B2", "        CHỨNG TỪ KẾT THÚC", bold=True)
    _set(ws, "B4", "Mục lục này gồm: 02 hồ sơ (đơn vị bảo quản)")
    _set(ws, "B5", "Đánh số từ hồ sơ số 0007 đến hồ sơ số 0008")
    _set(ws, "I7", "Nơi Thử, Ngày … tháng 1 năm 2026")
    _set(ws, "A8", "PHÊ DUYỆT", bold=True)
    _set(ws, "J8", "Người lập mục lục hồ sơ")
    _set(ws, "A9", "CHÁNH VĂN PHÒNG", bold=True)


def make_workbook(parts=ALL_PARTS, names=None, main_rows=None, n_main_cols=13) -> bytes:
    """Sổ Excel tổng hợp theo mẫu (bytes). parts: thứ tự sheet; names: tên sheet tương ứng."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for i, part in enumerate(parts):
        ws = wb.create_sheet(names[i] if names else DEFAULT_NAMES[part])
        if part == "main":
            _fill_main(ws, main_rows if main_rows is not None else DEFAULT_MAIN_ROWS, n_main_cols)
        else:
            {"cover": _fill_cover, "abbr": _fill_abbr, "appendix": _fill_appendix,
             "closing": _fill_closing}[part](ws)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class TestSynthetic(unittest.TestCase):
    def test_default_synthetic_workbook_converts(self):
        view = DocxView(m.convert_workbook(make_workbook()))
        self.assertEqual(len(view.sectprs), 5)
        self.assertEqual([grid_cols(t) for t in view.tables], [2, 4, 4, 13, 2])
        abbr = view.table_by_first_cell("SỐ TT", 4)
        self.assertEqual([norm_row(r) for r in table_matrix(abbr)][1:], [["1", "BCH", "Ban Chấp hành", ""]])
        self.assertEqual(len(table_rows(view.main_table())), 2 + 3)

    def test_sheets_renamed_and_reordered_detected_by_content(self):
        parts = ("closing", "main", "appendix", "abbr", "cover")
        names = ["S1", "S2", "S3", "S4", "S5"]
        raw = make_workbook(parts, names)
        data = m.read_workbook(raw)
        self.assertEqual(data.sheet_names, {"closing": "S1", "main_table": "S2", "appendix": "S3",
                                            "abbreviations": "S4", "cover": "S5"})
        view = DocxView(m.convert_workbook(raw))
        self.assertEqual(len(view.sectprs), 5)
        self.assertEqual([grid_cols(t) for t in view.tables], [2, 4, 4, 13, 2])
        self.assertIn("ĐẢNG CỘNG SẢN VIỆT NAM", view.section_text(0))
        self.assertIn("MỤC LỤC HỒ SƠ", view.section_text(0))
        self.assertIn("CHỮ VIẾT TẮT", view.section_text(1))
        self.assertIn("PHỤ LỤC TRA CỨU", view.section_text(2))
        self.assertIn("BẢNG THỐNG KÊ HỒ SƠ", view.section_text(3))
        self.assertIn("CHỨNG TỪ KẾT THÚC", view.section_text(4))

    def test_only_main_sheet(self):
        raw = make_workbook(("main",))
        data = m.read_workbook(raw)
        self.assertIsNone(data.cover)
        self.assertIsNone(data.abbreviations)
        self.assertIsNone(data.appendix)
        self.assertIsNone(data.closing)
        view = DocxView(m.convert_workbook(raw))
        self.assertEqual(len(view.sectprs), 1)
        self.assertEqual([grid_cols(t) for t in view.tables], [13])
        text = view.all_text
        self.assertIn("BẢNG THỐNG KÊ HỒ SƠ", text)
        for absent in ("CHỮ VIẾT TẮT", "PHỤ LỤC TRA CỨU", "CHỨNG TỪ KẾT THÚC", "MỤC LỤC HỒ SƠ"):
            self.assertNotIn(absent, text)
        s = view.sectprs[0]
        self.assertIsNone(s.find("w:titlePg", NS), "không có bìa thì trang 1 vẫn in số trang")
        self.assertIsNone(s.find("w:pgBorders", NS))
        self.assertEqual(s.find("w:pgNumType", NS).get(qn("w:start")), "1")

    def test_missing_main_table_raises(self):
        with self.assertRaises(m.ConvertError) as ctx:
            m.read_workbook(make_workbook(("cover", "abbr", "appendix", "closing")))
        self.assertTrue(str(ctx.exception).strip())
        self.assertIn("hồ sơ", str(ctx.exception).lower())
        wb = openpyxl.Workbook()                 # sổ hoàn toàn trống
        buf = io.BytesIO()
        wb.save(buf)
        with self.assertRaises(m.ConvertError):
            m.convert_workbook(buf.getvalue())

    def test_int_values_zero_padded_and_datetime(self):
        view = DocxView(m.convert_workbook(make_workbook()))
        body = [norm_row(r) for r in table_matrix(view.main_table())[2:]]
        self.assertEqual(body[0], ["", "", "", "", "I. NHÓM THỨ NHẤT DÒNG THỨ HAI"] + [""] * 8)
        self.assertEqual(body[1], ["01", "001", "01", "0007", "Hồ sơ A Dòng hai", "03/05/2024",
                                   "27", "3", "Vĩnh viễn", "", "", "", ""])
        self.assertEqual(body[2], ["", "001", "01", "0008", "Hồ sơ B", "01/01/2020-02/02/2021",
                                   "5", "2", "Vĩnh viễn", "", "", "", ""])

    def test_string_ids_not_padded(self):
        rows = [{"B": "1", "C": "1", "D": "12A", "E": "Hồ sơ lẻ", "I": "Vĩnh viễn"}]
        view = DocxView(m.convert_workbook(make_workbook(("main",), main_rows=rows)))
        body = [norm_row(r) for r in table_matrix(view.main_table())[2:]]
        self.assertEqual(body[0][:5], ["", "1", "1", "12A", "Hồ sơ lẻ"], "chuỗi trong Excel giữ nguyên, không đệm 0")

    def test_newline_in_title_kept_as_line_break(self):
        view = DocxView(m.convert_workbook(make_workbook()))
        rows = table_rows(view.main_table())
        group_tc = rows[2].findall("w:tc", NS)[4]
        data_tc = rows[3].findall("w:tc", NS)[4]
        self.assertEqual(len(list(group_tc.iter(qn("w:br")))), 1)
        self.assertEqual(tc_text(group_tc), "I. NHÓM THỨ NHẤT\nDÒNG THỨ HAI")
        self.assertEqual(len(list(data_tc.iter(qn("w:br")))), 1)
        self.assertEqual(tc_text(data_tc), "Hồ sơ A\nDòng hai")

    def test_non_xlsx_input_raises_convert_error(self):
        with self.assertRaises(m.ConvertError):
            m.convert_workbook(b"day khong phai la tep excel\n")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            z.writestr("a.txt", "x")
        with self.assertRaises(m.ConvertError):
            m.read_workbook(buf.getvalue())
        with self.assertRaises(m.ConvertError):
            m.read_workbook(os.path.join(tempfile.gettempdir(), "khong_ton_tai_%d.xlsx" % os.getpid()))

    def test_main_table_with_fewer_columns(self):
        view = DocxView(m.convert_workbook(make_workbook(("main",), n_main_cols=9)))
        tbl = view.main_table_any = [t for t in view.tables if grid_cols(t) == 9]
        self.assertEqual(len(tbl), 1)
        body = [norm_row(r) for r in table_matrix(tbl[0])[2:]]
        self.assertEqual(body[1], ["01", "001", "01", "0007", "Hồ sơ A Dòng hai", "03/05/2024", "27", "3", "Vĩnh viễn"])

    def test_cover_synthetic_values(self):
        data = m.read_workbook(make_workbook())
        cover = data.cover
        self.assertEqual([ln.text for ln in cover.header_lines], ["ĐẢNG CỘNG SẢN VIỆT NAM", "ĐẢNG ỦY XÃ THỬ NGHIỆM"])
        self.assertEqual(cover.header_lines[1].color, "FF0000")
        self.assertEqual([ln.text for ln in cover.bottom_left], ["Phông số: 002", "Mục lục số: 03", "Số trang:"])
        self.assertEqual([ln.text for ln in cover.bottom_right], ["Thời hạn bảo quản", "Vĩnh viễn"])
        self.assertEqual([ln.text for ln in cover.place_date], ["Nơi Thử, 1/2026"])
        closing = data.closing
        self.assertEqual(closing.title.text, "CHỨNG TỪ KẾT THÚC")
        self.assertEqual([ln.text for ln in closing.left_lines], ["PHÊ DUYỆT", "CHÁNH VĂN PHÒNG"])
        self.assertEqual([ln.text for ln in closing.right_lines],
                         ["Nơi Thử, Ngày … tháng 1 năm 2026", "Người lập mục lục hồ sơ"])

    def test_bien_muc_sheet_never_used_as_closing_fallback(self):
        """Thiếu sheet Chứng từ kết thúc và sheet cuối là 'Biên mục' (MỤC LỤC VĂN BẢN, TÀI LIỆU):
        không được lấy sheet đó làm chứng từ kết thúc theo dự phòng thứ tự sheet."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _fill_main(wb.create_sheet("Phông X VV"), DEFAULT_MAIN_ROWS)
        bm = wb.create_sheet("Biên mục")
        for ref, v in (("A1", "MỤC LỤC VĂN BẢN, TÀI LIỆU"), ("A2", "Mã hồ sơ: 0001-01"), ("A3", "Hồ sơ số: 0001"),
                       ("A4", "Số\nTT"), ("B4", "Số, ký\nhiệu"), ("C4", "Ngày tháng"), ("D4", "Tên loại và trích yếu"),
                       ("A5", "01"), ("B5", "03a-NQ/ĐU"), ("C5", "16/02/2006"),
                       ("D5", "Nghị quyết về nhiệm vụ quốc phòng năm 2006"),
                       ("A6", "02"), ("B6", "04-NQ/ĐU"), ("C6", "20/02/2006"), ("D6", "Nghị quyết xây dựng Đảng")):
            bm[ref] = v
        buf = io.BytesIO()
        wb.save(buf)
        raw = buf.getvalue()
        data = m.read_workbook(raw)
        self.assertNotIn("closing", data.sheet_names)
        self.assertIsNone(data.closing)
        self.assertEqual(data.sheet_names, {"main_table": "Phông X VV"})
        view = DocxView(m.convert_workbook(raw))
        self.assertEqual(len(view.sectprs), 1)
        xml = view.zip.read("word/document.xml").decode("utf-8")
        self.assertNotIn("MỤC LỤC VĂN BẢN", xml)
        self.assertNotIn("03a-NQ/ĐU", xml)

    def test_closing_body_line_with_date_words_stays_in_body(self):
        """Dòng nội dung ở cột B có 'ngày … tháng … năm' KHÔNG phải là dòng ngày tháng của khối chữ ký
        (quy tắc ngày tháng chỉ áp dụng cho ô ở cột > 6); 'PHÊ DUYỆT'/'Người lập' vẫn bất kể cột."""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _fill_main(wb.create_sheet("Phông X VV"), DEFAULT_MAIN_ROWS)
        cl = wb.create_sheet("Chứng từ kết thúc")
        _set(cl, "B2", "CHỨNG TỪ KẾT THÚC", bold=True)
        _set(cl, "B4", "Mục lục này gồm: 02 hồ sơ (đơn vị bảo quản)")
        _set(cl, "B5", "Tài liệu có thời gian từ ngày 01 tháng 01 năm 2006 đến 2010")
        _set(cl, "B6", "Khuyết hồ sơ (đơn vị bảo quản): không")
        _set(cl, "I12", "Nơi Thử, Ngày ….... tháng 7 năm 2026")
        _set(cl, "A13", "PHÊ DUYỆT", bold=True)
        _set(cl, "J13", "Người lập mục lục hồ sơ")
        _set(cl, "A14", "CHÁNH VĂN PHÒNG", bold=True)
        buf = io.BytesIO()
        wb.save(buf)
        raw = buf.getvalue()
        closing = m.read_workbook(raw).closing
        self.assertEqual([ln.text for ln in closing.body_lines],
                         ["Mục lục này gồm: 02 hồ sơ (đơn vị bảo quản)",
                          "Tài liệu có thời gian từ ngày 01 tháng 01 năm 2006 đến 2010",
                          "Khuyết hồ sơ (đơn vị bảo quản): không"])
        self.assertEqual([ln.text for ln in closing.left_lines], ["PHÊ DUYỆT", "CHÁNH VĂN PHÒNG"])
        self.assertEqual([ln.text for ln in closing.right_lines],
                         ["Nơi Thử, Ngày ….... tháng 7 năm 2026", "Người lập mục lục hồ sơ"])
        view = DocxView(m.convert_workbook(raw))
        blocks = view.section_blocks(len(view.sectprs) - 1)
        body_pars = [norm(p_text(b)) for b in blocks if b.tag == qn("w:p")]
        self.assertIn("Tài liệu có thời gian từ ngày 01 tháng 01 năm 2006 đến 2010", body_pars)
        sig = view.tables[-1]
        self.assertEqual(grid_cols(sig), 2)
        left, right = table_rows(sig)[0].findall("w:tc", NS)
        self.assertNotIn("Tài liệu có thời gian", tc_text(left) + tc_text(right))
        self.assertEqual([norm(p_text(p)) for p in left.findall("w:p", NS)], ["PHÊ DUYỆT", "CHÁNH VĂN PHÒNG"])
        self.assertEqual([norm(p_text(p)) for p in right.findall("w:p", NS)],
                         ["Nơi Thử, Ngày ….... tháng 7 năm 2026", "Người lập mục lục hồ sơ"])


# ---------------------------------------------------------------------------
# 6. Dòng lệnh
# ---------------------------------------------------------------------------
def _run_cli(args, timeout=300):
    env = dict(os.environ, PYTHONDONTWRITEBYTECODE="1", PYTHONIOENCODING="utf-8")
    return subprocess.run([sys.executable, os.path.join(PROJECT, "muc_luc_ho_so.py")] + list(args),
                          cwd=PROJECT, capture_output=True, text=True, encoding="utf-8",
                          errors="replace", env=env, timeout=timeout)


@unittest.skipUnless(HAS_SAMPLE, "thiếu tệp mẫu")
class TestCLI(unittest.TestCase):
    def test_cli_converts_sample_into_out_dir(self):
        has_docx = os.path.isfile(SAMPLE_DOCX)
        before = os.stat(SAMPLE_DOCX).st_mtime_ns if has_docx else None
        with tempfile.TemporaryDirectory() as tmp:
            r = _run_cli([SAMPLE_XLSX, "-o", tmp])
            self.assertEqual(r.returncode, 0, "stderr: %s" % r.stderr)
            self.assertIn("XONG", r.stdout)
            out = os.path.join(tmp, "ĐU Tam Đại Vĩnh viễn.docx")
            self.assertTrue(os.path.isfile(out), "không thấy %s; stdout=%s" % (out, r.stdout))
            with open(out, "rb") as f:
                blob = f.read()
            self.assertEqual(blob[:2], b"PK")
            self.assertEqual(len(DocxView(blob).sectprs), 5)
        if has_docx:
            self.assertEqual(os.stat(SAMPLE_DOCX).st_mtime_ns, before, "tệp .docx mẫu của người dùng bị đụng tới")

    def test_cli_options_take_effect(self):
        with tempfile.TemporaryDirectory() as tmp:
            r = _run_cli([SAMPLE_XLSX, "-o", tmp, "--start", "1", "--page", "Letter", "--font", "10"])
            self.assertEqual(r.returncode, 0, "stderr: %s" % r.stderr)
            with open(os.path.join(tmp, "ĐU Tam Đại Vĩnh viễn.docx"), "rb") as f:
                view = DocxView(f.read())
        self.assertEqual(view.sectprs[0].find("w:pgNumType", NS).get(qn("w:start")), "0")
        for s in view.sectprs:
            pgsz = s.find("w:pgSz", NS)
            self.assertEqual((pgsz.get(qn("w:w")), pgsz.get(qn("w:h"))), LETTER_LANDSCAPE)
        rows = table_rows(view.main_table())
        self.assertEqual({run_size(r) for r in runs_with_text(rows[2])}, {20})
        self.assertEqual({run_size(r) for r in runs_with_text(rows[0])}, {17})

    def test_cli_bad_file_reports_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = os.path.join(tmp, "hong.xlsx")
            with open(bad, "w", encoding="utf-8") as f:
                f.write("khong phai excel")
            r = _run_cli([bad, "-o", tmp])
            self.assertEqual(r.returncode, 1)
            self.assertIn("LỖI", r.stderr)
            self.assertFalse(os.path.exists(os.path.join(tmp, "hong.docx")))

    def test_cli_rejects_bad_options(self):
        r = _run_cli([SAMPLE_XLSX, "--page", "A3"])
        self.assertNotEqual(r.returncode, 0)
        r = _run_cli([SAMPLE_XLSX, "--start", "3"])
        self.assertNotEqual(r.returncode, 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
