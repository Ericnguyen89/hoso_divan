# -*- coding: utf-8 -*-
"""
In NHÃN HỘP từ Excel — THEO KHUÔN "sample.docx" (form-fill, kiểu report form VB).

BẢN HỢP NHẤT (in_nhanh.py + in_nhan.py):
    - Nền tảng lấy từ in_nhanh.py: KHÔNG dựng layout bằng code. Tệp "sample.docx"
      là KHUÔN thiết kế sẵn (khổ trang, bảng, kích thước ô, viền); chương trình chỉ:
        1. Nhân bản bảng khuôn cho từng trang (mỗi bảng = 2 hàng x 2 cột = 4 nhãn),
        2. Đổ dữ liệu Excel vào từng ô theo đúng "bản đồ định dạng" của khuôn.
      Muốn đổi thiết kế (kích thước, viền, lề...) -> chỉ cần sửa sample.docx bằng
      Word, KHÔNG phải sửa code.
    - Kế thừa từ in_nhan.py: chỉ số CO BỀ NGANG nội dung (CONTENT_IND_TW) — thụt
      lề trái/phải cố định cho mọi đoạn trong ô, vùng chữ luôn hẹp hơn ô nên chữ
      dài tự xuống dòng sớm, KHÔNG tràn ra mép viền.

Bản đồ định dạng 1 ô nhãn (phân tích từ sample.docx, nền Times New Roman 14pt):
    [dòng trống đầu]
    Tên phông (cột E)             : 14pt ĐẬM, canh giữa
      - dòng bắt đầu "KHÓA"       : 13pt ĐẬM
      - dòng bắt đầu "("          : 12pt NGHIÊNG
    -----                          : 14pt thường
    [1 dòng trống]
    Nội dung (cột D)              : 13pt; dòng mã "01.", "01.02." ... ĐẬM
    [dòng trống đệm - tính động để đẩy khối dưới xuống đáy]
    ĐVBQ: Từ {B} đến {C}          : 17pt ĐẬM
    Cặp (Hộp) số: {A}             : 17pt ĐẬM
    [chừa ~2 dòng hở dưới đáy]

Cấu trúc Excel (mỗi sheet, dòng tiêu đề bỏ qua):
    A: Hộp số | B: Từ hồ sơ số | C: Đến hồ sơ số | D: Nội dung | E: Tiêu đề
"""

import io
import os
import re
import sys
import copy
import math
import importlib.util
import subprocess


def ensure_deps():
    need = []
    for module, pip_name in (("docx", "python-docx"), ("openpyxl", "openpyxl")):
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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

_HERE = os.path.dirname(os.path.abspath(__file__))
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# -------------------- Thông số đọc từ khuôn / bản đồ định dạng --------------
TEMPLATE_NAME = "sample.docx"   # tệp khuôn (đặt cạnh chương trình)
COLS = 2                        # 2 nhãn / hàng (theo khuôn)
ROWS_PER_PAGE = 2               # 2 hàng / trang (theo khuôn)
LABELS_PER_PAGE = COLS * ROWS_PER_PAGE

# Cỡ chữ (pt) theo bản đồ định dạng của khuôn
SZ_HEADER = 14                  # tên phông (nền mặc định của khuôn)
SZ_KHOA = 13                    # dòng "KHÓA ..."
SZ_DATE = 12                    # dòng "(Từ ngày ...)"
SZ_SEP = 14                     # dòng -----
SZ_MID = 13                     # khối nội dung (cột D)
SZ_FOOT = 17                    # ĐVBQ / Cặp (Hộp) số
FONT_NAME = "Times New Roman"

# CHỈ SỐ CO BỀ NGANG NỘI DUNG (kế thừa từ in_nhan.py cũ):
# thụt lề trái/phải cho MỌI đoạn trong ô (twips) -> vùng chữ hẹp hơn ô,
# chữ dài tự xuống dòng sớm, không tràn ra mép viền.
# 170 twips ~ 0.3cm mỗi bên (muốn hở nhiều hơn: 227 ~ 0.4cm, 284 ~ 0.5cm).
CONTENT_IND_TW = 170

BOTTOM_GAP_LINES = 2            # chừa ~2 dòng hở dưới ĐVBQ (không thêm đoạn thật)
LINE_FACTOR = 1.18              # chiều cao dòng / cỡ chữ (ước lượng)
CHAR_FACTOR = 0.50              # bề rộng ký tự TB / cỡ chữ (ước lượng)

_CODE_RE = re.compile(r"^\s*\d+(\.\d+)*\.")     # "01." / "01.02." ...


# -------------------- Nạp khuôn --------------------
def _template_path():
    p = os.environ.get("TEM_TEMPLATE") or os.path.join(_HERE, TEMPLATE_NAME)
    if not os.path.exists(p):
        raise ValueError("Không tìm thấy tệp khuôn '%s' trong thư mục chương trình."
                         % TEMPLATE_NAME)
    return p


# -------------------- Tiện ích XML --------------------
def _clone(el):
    return copy.deepcopy(el)


def _cell_w_tw(tbl):
    """Bề rộng cột (twips) đọc từ tblGrid của khuôn."""
    grid = tbl.find(qn("w:tblGrid"))
    g = grid.findall(qn("w:gridCol"))[0]
    return int(g.get(qn("w:w")))


def _cellmar_lr_tw(tbl):
    """Đệm trái/phải của ô đọc từ tblCellMar của khuôn (mặc định 15)."""
    pr = tbl.find(qn("w:tblPr"))
    cm = pr.find(qn("w:tblCellMar")) if pr is not None else None
    if cm is not None:
        left = cm.find(qn("w:left"))
        if left is not None:
            return int(left.get(qn("w:w")))
    return 15


def _row_h_tw(tr):
    """Chiều cao hàng (twips) đọc từ trHeight của khuôn."""
    trPr = tr.find(qn("w:trPr"))
    h = trPr.find(qn("w:trHeight")) if trPr is not None else None
    return int(h.get(qn("w:val"))) if h is not None else 7120


def _find_proto_cell(tbl):
    """Tìm Ô NHÃN MẪU trong khuôn: ô đầu tiên (theo thứ tự tài liệu, kể cả trong
    bảng lồng) có nhiều đoạn trực tiếp — chính là ô chứa KHUNG VẼ (hình chữ nhật
    neo ở đoạn 0) + nội dung nhãn mẫu. Trả (ô, hàng chứa ô)."""
    for tc in tbl.iter(qn("w:tc")):
        if len(tc.findall(qn("w:p"))) > 1:
            return tc, tc.getparent()
    raise ValueError("Không nhận ra ô nhãn mẫu trong tệp khuôn sample.docx.")


def _build_master(sample_tbl):
    """Dựng bảng khuôn SẠCH cho 1 trang (2 hàng x 2 cột, mọi ô = bản sao ô nhãn
    mẫu). Khuôn sample.docx thực tế chứa hàng nhãn thứ 2 dưới dạng BẢNG LỒNG
    trong ô (do thao tác dán trong Word) -> không nhân bản nguyên bảng được,
    phải dựng phẳng lại từ ô mẫu. Trả (bảng, chiều cao hàng tw)."""
    proto_tc, proto_tr = _find_proto_cell(sample_tbl)
    row_h = _row_h_tw(proto_tr)

    tbl = OxmlElement("w:tbl")
    tbl.append(_clone(sample_tbl.find(qn("w:tblPr"))))    # style/viền/đệm như khuôn
    tbl.append(_clone(sample_tbl.find(qn("w:tblGrid"))))  # 2 cột 4696tw như khuôn
    for _ in range(ROWS_PER_PAGE):
        tr = OxmlElement("w:tr")
        trPr = OxmlElement("w:trPr")
        h = OxmlElement("w:trHeight")
        h.set(qn("w:val"), str(row_h))
        h.set(qn("w:hRule"), "exact")
        trPr.append(h)
        trPr.append(OxmlElement("w:cantSplit"))
        tr.append(trPr)
        for _ in range(COLS):
            tr.append(_clone(proto_tc))
        tbl.append(tr)
    _pin_table_width(tbl)
    return tbl, row_h


def _pin_table_width(tbl):
    """GHIM bề rộng bảng = tổng cột (dxa). Khuôn để tblW='auto' khiến Word co
    giãn cột theo nội dung -> 2 cột lệch nhau, viền không khớp ô (lỗi đã gặp).
    Đây là chỉnh sửa DUY NHẤT so với khuôn."""
    pr = tbl.find(qn("w:tblPr"))
    grid = tbl.find(qn("w:tblGrid"))
    total = sum(int(g.get(qn("w:w"))) for g in grid.findall(qn("w:gridCol")))
    for old in pr.findall(qn("w:tblW")):
        pr.remove(old)
    w = OxmlElement("w:tblW")
    w.set(qn("w:w"), str(total))
    w.set(qn("w:type"), "dxa")
    style = pr.find(qn("w:tblStyle"))
    if style is not None:
        style.addnext(w)
    else:
        pr.insert(0, w)


def _write_para(p, text, size_pt, bold, italic, center=True):
    """Ghi đè nội dung + định dạng cho đoạn p (giữ nguyên phần tử đoạn của khuôn)."""
    # pPr: thụt lề 2 bên (co bề ngang chữ) + canh giữa
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    for tag in ("w:ind", "w:jc"):
        for old in pPr.findall(qn(tag)):
            pPr.remove(old)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(CONTENT_IND_TW))
    ind.set(qn("w:right"), str(CONTENT_IND_TW))
    jc = None
    if center:
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center")
    # theo schema pPr: ind/jc phải đứng TRƯỚC rPr (định dạng dấu đoạn) nếu có
    anchor = pPr.find(qn("w:rPr"))
    for el in (ind, jc):
        if el is None:
            continue
        if anchor is not None:
            anchor.addprevious(el)
        else:
            pPr.append(el)

    # gộp về 1 run duy nhất
    runs = p.findall(qn("w:r"))
    if not runs:
        r0 = OxmlElement("w:r")
        p.append(r0)
    else:
        r0 = runs[0]
        for extra in runs[1:]:
            p.remove(extra)

    rPr = r0.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r0.insert(0, rPr)
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT_NAME)

    def flag(tag, on):
        for e in rPr.findall(qn(tag)):
            rPr.remove(e)
        if on:
            rPr.append(OxmlElement(tag))
    flag("w:b", bold)
    flag("w:bCs", bold)
    flag("w:i", italic)
    flag("w:iCs", italic)
    for tag in ("w:sz", "w:szCs"):
        for e in rPr.findall(qn(tag)):
            rPr.remove(e)
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(int(round(size_pt * 2))))
        rPr.append(el)

    # nội dung chữ
    for ch in list(r0):
        if ch.tag in (qn("w:t"), qn("w:br"), qn("w:tab")):
            r0.remove(ch)
    lines = str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for i, line in enumerate(lines):
        if i > 0:
            r0.append(OxmlElement("w:br"))
        t = OxmlElement("w:t")
        t.set(XML_SPACE, "preserve")
        t.text = line
        r0.append(t)


def _write_cell(tc, items):
    """Đổ danh sách (text,size,bold,italic) vào ô khuôn: TÁI DÙNG đoạn có sẵn,
    thêm nếu thiếu, xóa bớt nếu thừa (giữ >=1 đoạn để ô hợp lệ)."""
    paras = tc.findall(qn("w:p"))
    for i, (text, size, bold, italic) in enumerate(items):
        if i < len(paras):
            p = paras[i]
        else:
            p = _clone(paras[-1])
            tc.append(p)
        _write_para(p, text, size, bold, italic)
    keep = max(1, len(items))
    for p in paras[keep:]:
        tc.remove(p)


# -------------------- Ước lượng chiều cao (đệm động) --------------------
def _est_height(text, size_pt, bold, text_width_pt):
    line_h = size_pt * LINE_FACTOR
    if not text:
        return line_h
    factor = CHAR_FACTOR * (1.12 if bold else 1.0)   # chữ đậm rộng hơn ~12%
    cpl = max(6, int(text_width_pt / (size_pt * factor)))
    n = 0
    for seg in str(text).split("\n"):
        n += max(1, math.ceil(len(seg) / cpl))
    return n * line_h


# -------------------- Nội dung nhãn theo bản đồ định dạng --------------------
def _split(text):
    """Tách dòng, bỏ dòng trống đầu/cuối nhưng GIỮ dòng trống ở giữa
    (khoảng cách người dùng cố ý đặt trong Excel)."""
    lines = str(text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines = [ln.strip() for ln in lines]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return lines


def _label_items(lb, inner_pt, text_width_pt):
    """Dàn nội dung 1 ô theo bản đồ định dạng của khuôn."""
    items = [("", SZ_MID, False, False)]                 # dòng trống đầu

    for line in _split(lb["tieu_de"]):                   # khối đầu (cột E)
        ls = line.lstrip()
        if ls.startswith("("):
            items.append((line, SZ_DATE, False, True))   # (Từ ngày...) nghiêng
        elif ls.upper().startswith("KHÓA") or ls.upper().startswith("KHOA"):
            items.append((line, SZ_KHOA, True, False))   # KHÓA... 13pt đậm
        else:
            items.append((line, SZ_HEADER, True, False)) # tên phông 14pt đậm
    items.append(("-----", SZ_SEP, False, False))
    items.append(("", SZ_MID, False, False))             # 1 dòng trống

    for line in _split(lb["noi_dung"]):                  # khối giữa (cột D)
        if line:
            items.append((line, SZ_MID, bool(_CODE_RE.match(line)), False))
        else:
            items.append(("", SZ_MID, False, False))     # giữ dòng trống cố ý

    tu, den = lb["tu"], lb["den"]
    dvbq = "ĐVBQ: Từ %s đến %s" % (tu, den) if den else "ĐVBQ: Từ %s" % tu
    foot = [(dvbq, SZ_FOOT, True, False),
            ("Cặp (Hộp) số: %s" % lb["hop"], SZ_FOOT, True, False)]

    # Đệm động: đẩy ĐVBQ xuống sao cho dưới nó còn ~BOTTOM_GAP_LINES dòng hở
    empty_h = SZ_MID * LINE_FACTOR
    used = sum(_est_height(t, s, b, text_width_pt) for (t, s, b, i) in items + foot)
    room = inner_pt - used - BOTTOM_GAP_LINES * empty_h
    filler = max(0, int(room / empty_h))
    items += [("", SZ_MID, False, False)] * filler
    items += foot
    return items


# -------------------- Đọc dữ liệu Excel --------------------
def _txt(v):
    return "" if v is None else str(v).strip()


def _norm(s):
    return re.sub(r"\s+", "", (s or "").lower())


def read_labels(wb):
    labels = []
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows(min_row=1, max_col=5, values_only=True):
            def c(i):
                return _txt(row[i]) if i < len(row) and row[i] is not None else ""
            hop, tu, den, noi_dung, tieu_de = c(0), c(1), c(2), c(3), c(4)
            if _norm(hop) in ("hộpsố", "hopso"):
                continue
            if not any([hop, tu, den, noi_dung, tieu_de]):
                continue
            labels.append({"hop": hop, "tu": tu, "den": den,
                           "noi_dung": noi_dung, "tieu_de": tieu_de})
    return labels


# -------------------- Đổ dữ liệu vào khuôn --------------------
def _page_break_p():
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    return p


def build_docx(labels, template_bytes, progress=None):
    doc = Document(io.BytesIO(template_bytes))
    body = doc.element.body

    tbls = body.findall(qn("w:tbl"))
    if not tbls:
        raise ValueError("Tệp khuôn không có bảng nhãn.")
    master, row_h = _build_master(tbls[0])   # bảng SẠCH 2x2 dựng từ ô nhãn mẫu

    # Thông số hình học đọc TỪ KHUÔN (bề rộng chữ trừ thêm phần thụt lề 2 bên
    # -> ước lượng xuống dòng khớp với vùng chữ đã co hẹp)
    cell_w = _cell_w_tw(master)
    mar_lr = _cellmar_lr_tw(master)
    text_width_pt = (cell_w - 2 * (mar_lr + CONTENT_IND_TW)) / 20.0
    inner_pt = row_h / 20.0 - 4              # chiều cao chữ trong ô (pt, trừ an toàn)

    # Dọn thân tài liệu: chỉ giữ sectPr cuối (khuôn có section-break rải rác
    # trong các đoạn -> bỏ hết để còn 1 section duy nhất)
    sectPr = body.find(qn("w:sectPr"))
    for ch in list(body):
        if ch is not sectPr:
            body.remove(ch)

    def add(el):
        if sectPr is not None:
            sectPr.addprevious(el)
        else:
            body.append(el)

    total = len(labels)
    first = True
    for start in range(0, total, LABELS_PER_PAGE):
        page = labels[start:start + LABELS_PER_PAGE]
        if not first:
            add(_page_break_p())
        first = False

        tbl = _clone(master)
        cells = [tc for tr in tbl.findall(qn("w:tr"))
                 for tc in tr.findall(qn("w:tc"))]
        for idx in range(LABELS_PER_PAGE):
            if idx < len(page):
                items = _label_items(page[idx], inner_pt, text_width_pt)
            else:
                # Ô chưa dùng: xóa chữ, GIỮ khung vẽ (nhãn trống như file mẫu)
                items = [("", SZ_MID, False, False)]
            _write_cell(cells[idx], items)
        add(tbl)
        if progress is not None:
            progress(min(start + LABELS_PER_PAGE, total), total)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def generate_from_bytes(data_bytes, progress=None, template_bytes=None):
    """API cho app.py: trả (docx_bytes, số_nhãn)."""
    if template_bytes is None:
        with open(_template_path(), "rb") as f:
            template_bytes = f.read()
    wb = openpyxl.load_workbook(io.BytesIO(data_bytes), read_only=True,
                                data_only=True)
    labels = read_labels(wb)
    wb.close()
    if not labels:
        raise ValueError("Không tìm thấy dữ liệu nhãn trong Excel.")
    return build_docx(labels, template_bytes, progress=progress), len(labels)


# -------------------- CLI --------------------
def main(argv):
    if len(argv) < 2:
        print("Cách dùng: python in_nhan.py <du_lieu.xlsx> [ket_qua.docx]")
        return 1
    src = argv[1]
    dst = argv[2] if len(argv) > 2 else os.path.splitext(src)[0] + " - Nhãn.docx"
    with open(src, "rb") as f:
        docx_bytes, n = generate_from_bytes(
            f.read(),
            progress=lambda d, t: print("  %d/%d" % (d, t), end="\r"))
    with open(dst, "wb") as f:
        f.write(docx_bytes)
    print("\n✓ Đã tạo %d nhãn -> %s" % (n, dst))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
