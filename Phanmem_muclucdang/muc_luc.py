# -*- coding: utf-8 -*-
"""
Tạo MỤC LỤC VĂN BẢN, TÀI LIỆU (.docx, khổ A4 NẰM NGANG) từ Excel biên mục.

Cách hoạt động
--------------
- Đọc TẤT CẢ các sheet trong tệp .xlsx. MỖI SHEET -> MỘT tệp .docx riêng
  (tiện in ấn từng phông/đơn vị).
- Trong một sheet, gom các dòng theo cột "Hồ sơ số" (cột A, tự điền xuống -
  ffill). Mỗi hồ sơ -> 1 bảng mục lục 14 cột, mỗi hồ sơ nằm trên trang mới.
- Kích thước trang, lề và tỉ lệ 14 cột lấy theo tệp mẫu sample.docx
  (A4 ngang 11.69" x 8.27").

Cấu trúc Excel (mỗi sheet)
--------------------------
    Dòng 1: tựa "MỤC LỤC TÀI LIỆU"
    Dòng 2: hàng tiêu đề (Hồ sơ số | Số TT | Số, ký hiệu | ...)
    Dòng 3: hàng số cột "(3) (4) ... (16)"   (bỏ qua)
    Dòng 4+: dữ liệu
    Cột:  A=Hồ sơ số  B=Số TT  C=Số, ký hiệu  D=Ngày tháng
          E=Tên loại và trích yếu  F=Tác giả  G=Người ký  H=Độ mật
          I=Loại bản  J=Trang số  K=Số trang  L=Từ khoá  M=Ghi chú
          N=Số lượng tệp (file)  O=Tên tệp tài liệu
"""

import io
import os
import re
import sys
import importlib.util
import subprocess


def ensure_deps():
    """Tự cài thư viện còn thiếu bằng đúng Python đang chạy, rồi nạp lại tiến
    trình — nhờ vậy chạy được bằng Python hệ thống mà KHÔNG cần môi trường ảo
    (venv). Trên Python bị 'khóa' theo PEP 668 (Homebrew...), sẽ lần lượt thử
    --user rồi --break-system-packages."""
    need = []
    for module, pip_name in (("docx", "python-docx"),
                             ("openpyxl", "openpyxl")):
        if importlib.util.find_spec(module) is None:
            need.append(pip_name)
    if not need:
        return
    print("• Thiếu thư viện: %s — đang tự cài đặt..." % ", ".join(need))
    base = [sys.executable, "-m", "pip", "install", "--quiet"]
    in_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    attempts = ([[]] if in_venv else [["--user"]]) + [["--break-system-packages"], []]
    last = None
    for extra in attempts:
        r = subprocess.run(base + extra + need,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        if r.returncode == 0:
            break
        last = r.stdout
    else:
        if last:
            sys.stdout.buffer.write(last)
        sys.exit("Không tự cài được. Hãy chạy thủ công:\n"
                 "  %s -m pip install %s" % (sys.executable, " ".join(need)))
    os.execv(sys.executable, [sys.executable] + sys.argv)


ensure_deps()

import openpyxl
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------------------------------------------------------------------------
# Thông số trang & bảng lấy theo sample.docx (A4 nằm ngang)
# ---------------------------------------------------------------------------
PAGE_W_IN = 11.69          # A4 ngang: chiều rộng
PAGE_H_IN = 8.27           # A4 ngang: chiều cao
MARGIN_LEFT_IN = 0.71
MARGIN_RIGHT_IN = 0.25
MARGIN_TOP_IN = 0.29
MARGIN_BOTTOM_IN = 0.23

USABLE_W_IN = PAGE_W_IN - MARGIN_LEFT_IN - MARGIN_RIGHT_IN   # ~10.73"

# 14 cột đầu ra. Tiêu đề hiển thị (\n = xuống dòng trong ô).
HEADERS = [
    "Số\nTT",              # 0  <- Excel cột B
    "Số, ký\nhiệu",        # 1  <- C
    "Ngày tháng",          # 2  <- D
    "Tên loại và trích yếu",  # 3  <- E
    "Tác giả",             # 4  <- F
    "Người ký",            # 5  <- G
    "Độ mật",              # 6  <- H
    "Loại bản",            # 7  <- I
    "Trang\nsố",           # 8  <- J
    "Số\ntrang",           # 9  <- K
    "Từ\nkhoá",            # 10 <- L
    "Ghi\nchú",            # 11 <- M
    "Số lượng tệp\n(file)",  # 12 <- N
    "Tên tệp\ntài liệu",   # 13 <- O
]

# Độ rộng 14 cột (inch) theo sample.docx; sẽ được co giãn để lấp đủ bề ngang.
_SAMPLE_WIDTHS = [0.283, 0.828, 0.793, 2.074, 0.895, 0.751, 0.626,
                  0.481, 0.496, 0.481, 0.561, 0.524, 0.524, 1.26]
_scale = USABLE_W_IN / sum(_SAMPLE_WIDTHS)
COL_WIDTHS_IN = [w * _scale for w in _SAMPLE_WIDTHS]

# Cột "Tên loại và trích yếu" căn trái (dễ đọc), các cột còn lại căn giữa.
COL_LEFT = 3

FONT_NAME = "Times New Roman"
SIZE_TITLE = Pt(14)      # "MỤC LỤC VĂN BẢN, TÀI LIỆU"
SIZE_SUB = Pt(13)        # "Mã hồ sơ:" / "Tên hồ sơ:"
SIZE_HEADER = Pt(11)     # hàng tiêu đề bảng
SIZE_BODY = Pt(11)       # dữ liệu

# Số cột Excel (1-based) tương ứng 14 cột đầu ra: B..O = 2..15
_EXCEL_COLS = list(range(2, 16))


# ---------------------------------------------------------------------------
# Tiện ích XML cho bảng
# ---------------------------------------------------------------------------
def _twips(inches):
    return str(int(round(inches * 1440)))


def _set_table_layout(table):
    """Đặt bảng rộng cố định, viền đơn, không đệm ô, layout 'fixed'."""
    tbl = table._tbl
    if tbl.tblPr is not None:
        tbl.remove(tbl.tblPr)
    tblPr = OxmlElement("w:tblPr")
    tbl.insert(0, tblPr)

    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), _twips(USABLE_W_IN))
    tblW.set(qn("w:type"), "dxa")
    tblPr.append(tblW)

    tblInd = OxmlElement("w:tblInd")
    tblInd.set(qn("w:w"), "0")
    tblInd.set(qn("w:type"), "dxa")
    tblPr.append(tblInd)

    tblCellMar = OxmlElement("w:tblCellMar")
    for side, val in (("top", "0"), ("left", "43"), ("bottom", "0"), ("right", "43")):
        el = OxmlElement("w:%s" % side)
        el.set(qn("w:w"), val)
        el.set(qn("w:type"), "dxa")
        tblCellMar.append(el)
    tblPr.append(tblCellMar)

    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")
    tblPr.append(tblLayout)

    tblBorders = OxmlElement("w:tblBorders")
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement("w:%s" % name)
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        tblBorders.append(b)
    tblPr.append(tblBorders)


def _set_cell_width(cell, width_in):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcW")):
        tcPr.remove(old)
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), _twips(width_in))
    tcW.set(qn("w:type"), "dxa")
    tcPr.append(tcW)


def _mark_header_row(row):
    """Đánh dấu hàng lặp lại ở đầu mỗi trang khi bảng bị tách trang."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))


def _fill_cell(cell, text, size, bold, align, width_in):
    """Ghi nội dung 1 ô (giữ font Times New Roman, xử lý xuống dòng \\n)."""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    para = cell.paragraphs[0]
    pf = para.paragraph_format
    pf.space_before = Pt(1)
    pf.space_after = Pt(1)
    pf.line_spacing = 1.0
    para.alignment = align

    lines = str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    run = para.add_run(lines[0])
    for line in lines[1:]:
        run.add_break()
        run.add_text(line)
    run.bold = bold
    run.font.name = FONT_NAME
    run.font.size = size
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)

    _set_cell_width(cell, width_in)


# ---------------------------------------------------------------------------
# Đọc dữ liệu Excel
# ---------------------------------------------------------------------------
def _norm(s):
    return re.sub(r"\s+", "", (s or "").lower())


def _txt(v):
    return "" if v is None else str(v).strip()


def read_sheet_groups(ws):
    """Đọc 1 sheet -> [(ho_so_so, [ [14 giá trị], ... ]), ...] đã sắp xếp.

    Tự dò hàng tiêu đề (ô có 'Số TT'), bỏ hàng số cột '(3)...(16)', rồi đọc
    dữ liệu; cột 'Hồ sơ số' (A) được điền xuống (ffill) để gom nhóm.

    Đọc TOÀN BỘ sheet một lần bằng iter_rows (nhanh, kể cả workbook read_only)
    rồi xử lý theo chỉ số — tránh truy cập ngẫu nhiên ws.cell() vốn rất chậm
    ở chế độ read_only.
    """
    # Vật chất hoá toàn bộ ô cần dùng: cột A (1) + 14 cột dữ liệu (2..15)
    data = [row[:15] for row in ws.iter_rows(min_row=1, max_col=15,
                                             values_only=True)]

    def col(row, c):   # c: 1-based
        return _txt(row[c - 1]) if c - 1 < len(row) else ""

    # 1) Tìm hàng tiêu đề: ô cột B (Số TT) chuẩn hoá == "sốtt"
    header_idx = None
    for i in range(min(len(data), 12)):
        if _norm(col(data[i], 2)) in ("sốtt", "stt"):
            header_idx = i
            break
    start = (header_idx + 1) if header_idx is not None else 0

    # 2) Bỏ hàng số cột "(3) (4) ..." nếu có ngay dưới tiêu đề
    if start < len(data) and re.fullmatch(r"\(\d+\)", col(data[start], 2)):
        start += 1

    # 3) Đọc dữ liệu, ffill "Hồ sơ số"
    groups = {}
    order = []
    cur_hs = ""
    for row in data[start:]:
        a = col(row, 1)
        if a:
            cur_hs = a
        vals = [col(row, c) for c in _EXCEL_COLS]
        if not any(vals):          # dòng trống hoàn toàn -> bỏ
            continue
        # Bỏ dòng lặp lại tiêu đề (nếu dán nhiều bảng vào cùng sheet)
        if _norm(vals[0]) in ("sốtt", "stt") or re.fullmatch(r"\(\d+\)", vals[0]):
            continue
        key = cur_hs or "(không rõ)"
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(vals)

    return sorted(((k, groups[k]) for k in order),
                  key=lambda kv: _sort_key(kv[0]))


def _sort_key(key):
    m = re.search(r"\d+", key or "")
    return (int(m.group()) if m else 1_000_000_000, key or "")


# ---------------------------------------------------------------------------
# Dựng tài liệu Word cho 1 sheet
# ---------------------------------------------------------------------------
def _new_document():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT_NAME
    st.font.size = SIZE_BODY
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)

    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width = Inches(PAGE_W_IN)
    sec.page_height = Inches(PAGE_H_IN)
    sec.left_margin = Inches(MARGIN_LEFT_IN)
    sec.right_margin = Inches(MARGIN_RIGHT_IN)
    sec.top_margin = Inches(MARGIN_TOP_IN)
    sec.bottom_margin = Inches(MARGIN_BOTTOM_IN)
    return doc


def _add_heading(doc, ma_ho_so, ten_ho_so):
    def line(text, size):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(text)
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = size
        run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
        return p

    line("MỤC LỤC VĂN BẢN, TÀI LIỆU", SIZE_TITLE)
    line("Mã hồ sơ: %s" % ma_ho_so, SIZE_SUB)
    line("Hồ sơ số: %s" % ten_ho_so, SIZE_SUB)


def _add_table(doc, rows):
    table = doc.add_table(rows=1, cols=14)
    table.style = "Table Grid"
    table.allow_autofit = False
    table.autofit = False
    _set_table_layout(table)

    # Hàng tiêu đề
    hdr = table.rows[0].cells
    for i, text in enumerate(HEADERS):
        _fill_cell(hdr[i], text, SIZE_HEADER, True,
                   WD_ALIGN_PARAGRAPH.CENTER, COL_WIDTHS_IN[i])
    _mark_header_row(table.rows[0])

    # Dữ liệu
    for vals in rows:
        cells = table.add_row().cells
        for i in range(14):
            # Lấy đúng nguyên nội dung từ Excel (kể cả cột O - Tên tệp tài liệu)
            text = vals[i]
            align = (WD_ALIGN_PARAGRAPH.LEFT if i == COL_LEFT
                     else WD_ALIGN_PARAGRAPH.CENTER)
            _fill_cell(cells[i], text, SIZE_BODY, False, align, COL_WIDTHS_IN[i])


def build_sheet_docx(groups, phong_so, muc_luc, progress=None):
    """Dựng .docx cho 1 sheet -> bytes. groups: kết quả read_sheet_groups()."""
    doc = _new_document()
    total = len(groups)
    for gi, (ho_so_so, rows) in enumerate(groups):
        if gi > 0:
            doc.add_page_break()
        # Mã hồ sơ (đầu bảng) = Phông số - Mục lục số; Hồ sơ số = Hồ sơ số.
        ma_ho_so = build_ma_ho_so(phong_so, muc_luc)
        _add_heading(doc, ma_ho_so, ho_so_so)
        _add_table(doc, rows)
        if progress is not None:
            progress(gi + 1, total)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def build_ma_ho_so(phong_so, muc_luc):
    """Ghép Mã hồ sơ = Phông số - Mục lục số (bỏ phần rỗng)."""
    parts = [p for p in (phong_so, muc_luc) if (p or "").strip()]
    return "-".join(str(p).strip() for p in parts)


def _safe_name(s):
    s = re.sub(r'[\\/:*?"<>|]+', "_", (s or "").strip())
    return s or "sheet"


def generate_from_bytes(data_bytes, phong_so="001", muc_luc="01", progress=None):
    """Nhận bytes .xlsx -> danh sách kết quả cho từng sheet.

    Trả về list các dict:
        {"sheet": tên sheet, "docx_name": tên tệp .docx,
         "docx_bytes": bytes, "ho_so": số hồ sơ, "rows": tổng số dòng}
    progress(done, total): báo tiến độ theo TỔNG SỐ HỒ SƠ của cả tệp.
    """
    wb = openpyxl.load_workbook(io.BytesIO(data_bytes), read_only=True,
                                data_only=True)
    # Đọc trước để biết tổng số hồ sơ (cho thanh tiến độ)
    per_sheet = []
    for name in wb.sheetnames:
        groups = read_sheet_groups(wb[name])
        if groups:
            per_sheet.append((name, groups))
    wb.close()

    if not per_sheet:
        raise ValueError("Không đọc được dữ liệu hồ sơ nào từ tệp Excel.")

    total_hs = sum(len(g) for _, g in per_sheet)
    done = 0
    results = []
    for name, groups in per_sheet:
        def cb(cur, tot, _base=done):
            if progress:
                progress(_base + cur, total_hs)

        docx_bytes = build_sheet_docx(groups, phong_so, muc_luc, progress=cb)
        done += len(groups)
        results.append({
            "sheet": name,
            "docx_name": "MLVB_%s.docx" % _safe_name(name),
            "docx_bytes": docx_bytes,
            "ho_so": len(groups),
            "rows": sum(len(r) for _, r in groups),
        })
    return results


# ---------------------------------------------------------------------------
# Chạy dòng lệnh (thử nhanh)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse
    import os

    ap = argparse.ArgumentParser(description="Tạo mục lục văn bản (A4 ngang).")
    ap.add_argument("xlsx", help="Tệp Excel đầu vào")
    ap.add_argument("--phong", default="001", help="Phông số")
    ap.add_argument("--muc-luc", default="01", help="Mục lục số")
    ap.add_argument("--outdir", default=".", help="Thư mục xuất .docx")
    args = ap.parse_args()

    with open(args.xlsx, "rb") as f:
        data = f.read()
    res = generate_from_bytes(data, phong_so=args.phong, muc_luc=args.muc_luc,
                              progress=lambda d, t: print("  %d/%d hồ sơ" % (d, t), end="\r"))
    print()
    os.makedirs(args.outdir, exist_ok=True)
    for r in res:
        out = os.path.join(args.outdir, r["docx_name"])
        with open(out, "wb") as f:
            f.write(r["docx_bytes"])
        print("✓ %-20s %3d hồ sơ, %4d dòng -> %s"
              % (r["sheet"], r["ho_so"], r["rows"], out))
