# -*- coding: utf-8 -*-
"""
In NHÃN HỘP từ Excel — THEO KHUÔN "Sample.docx" (form-fill, kiểu report form VB).

Nền tảng: KHÔNG dựng layout bằng code. Tệp "Sample.docx" là KHUÔN thiết kế sẵn
(khổ trang, bảng, kích thước ô, viền, KHUNG VẼ trong ô); chương trình chỉ:
    1. Nhân bản bảng khuôn cho từng trang (mỗi bảng = 2 hàng x 2 cột = 4 nhãn),
    2. Đổ dữ liệu Excel vào từng ô theo đúng "bản đồ định dạng" của khuôn.
Muốn đổi thiết kế (kích thước, viền, lề, khung vẽ...) -> chỉ cần sửa Sample.docx
bằng Word, KHÔNG phải sửa code.

CHỈ SỐ CO BỀ NGANG nội dung (CONTENT_IND_TW): thụt lề trái/phải cố định cho mọi
đoạn trong ô, vùng chữ luôn hẹp hơn ô nên chữ dài tự xuống dòng sớm, KHÔNG tràn
ra mép viền. Khuôn Sample.docx cũng dùng đúng giá trị này (w:ind 170).

===============================================================================
CẤU TRÚC EXCEL ("Dữ liệu hộp.xlsx") — mỗi dòng dữ liệu = một nhãn
===============================================================================
    Số hộp | Từ Hồ sơ | Đến Hồ sơ | Số lượng hồ sơ | Tên phông | Năm

Dòng 1 là tiêu đề cột. Chương trình đọc THEO TÊN CỘT nên thứ tự cột đặt thế nào
cũng được, thừa cột cũng không sao. Mọi sheet trong tệp đều được đọc.

BẢN ĐỒ ĐỊNH DẠNG 1 Ô NHÃN (theo bảng thứ nhất của Sample.docx — các bảng sau
trong tệp mẫu là thiết kế khác, chỉ để tham khảo, chương trình không dựng):

    [dòng trống đầu — GIỮ khung vẽ neo ở đây]  13pt
    Tên phông                                  14pt ĐẬM (canh giữa)
      - dòng bắt đầu "KHÓA"                    13pt ĐẬM
      - dòng bắt đầu "("                       12pt NGHIÊNG
    -----                                      14pt
    [2 dòng trống]
    HỘP SỐ                                     20pt ĐẬM
    {Số hộp}                                   100pt ĐẬM   (đệm 0 -> 2 chữ số)
    GỒM {Số lượng hồ sơ} ĐVBQ                  20pt ĐẬM
    Từ {Từ Hồ sơ} đến {Đến Hồ sơ}              20pt ĐẬM    (đệm 0 -> 4 chữ số)
    [dòng trống đệm — tính động để đẩy dòng NĂM xuống đáy]
    NĂM {Năm}                                  20pt ĐẬM
"""

import io
import os
import re
import sys
import copy
import math
import datetime
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
TEMPLATE_NAME = "Sample.docx"   # tệp khuôn (đặt cạnh chương trình)
COLS = 2                        # 2 nhãn / hàng (theo khuôn)
ROWS_PER_PAGE = 2               # 2 hàng / trang (theo khuôn)
LABELS_PER_PAGE = COLS * ROWS_PER_PAGE

# Cỡ chữ (pt) theo bản đồ định dạng của khuôn. Nền mặc định của khuôn là
# Times New Roman 14pt (w:sz 28 trong docDefaults).
SZ_ANCHOR = 13                  # dòng trống đầu (đoạn neo khung vẽ)
SZ_HEADER = 14                  # tên phông
SZ_KHOA = 13                    # dòng "KHÓA ..."
SZ_DATE = 12                    # dòng "(Từ ngày ...)"
SZ_SEP = 14                     # dòng -----
SZ_LABEL = 20                   # "HỘP SỐ"
SZ_NUM = 100                    # số hộp (chữ số khổ lớn)
SZ_INFO = 20                    # "GỒM n ĐVBQ" / "Từ ... đến ..."
SZ_YEAR = 20                    # "NĂM ..."
FONT_NAME = "Times New Roman"

GAP_AFTER_SEP = 2               # số dòng trống giữa "-----" và "HỘP SỐ"

# Đệm 0 cho số thứ tự (dùng cho cả hai kiểu nhãn, theo đúng khuôn Sample.docx).
# Chuỗi không phải số (vd "01-A") được giữ nguyên.
PAD_BOX = 2                     # số hộp   ("1"  -> "01")
PAD_FILE = 4                    # số hồ sơ ("14" -> "0014")

# CHỈ SỐ CO BỀ NGANG NỘI DUNG: thụt lề trái/phải cho MỌI đoạn trong ô (twips)
# -> vùng chữ hẹp hơn ô, chữ dài tự xuống dòng sớm, không tràn ra mép viền.
# 170 twips ~ 0.3cm mỗi bên (muốn hở nhiều hơn: 227 ~ 0.4cm, 284 ~ 0.5cm).
CONTENT_IND_TW = 170

# Dòng NĂM nằm sát đáy đúng như khuôn -> không chừa thêm dòng hở. Biên an toàn
# đã có ở inner_pt (-4pt) và ở LINE_FACTOR (1.18 > ~1.15 thực tế).
BOTTOM_GAP_LINES = 0
LINE_FACTOR = 1.18              # chiều cao dòng / cỡ chữ (ước lượng)
CHAR_FACTOR = 0.50              # bề rộng ký tự TB / cỡ chữ (ước lượng)

# Tự co cỡ chữ khi tên phông dài quá khổ nhãn: thử lần lượt các tỉ lệ dưới đây,
# lấy tỉ lệ đầu tiên vừa ô. MIN_PT là sàn để chữ vẫn đọc được.
SHRINK_STEPS = (1.0, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6)
MIN_PT = 7.0


# -------------------- Nạp khuôn --------------------
def _template_path():
    """Đường dẫn tệp khuôn. Tìm không phân biệt chữ hoa/thường để chương trình
    chạy được cả trên Linux (hệ tệp phân biệt) khi tên tệp là 'sample.docx'."""
    env = os.environ.get("TEM_TEMPLATE")
    if env:
        if not os.path.exists(env):
            raise ValueError("Không tìm thấy tệp khuôn '%s' (biến TEM_TEMPLATE)." % env)
        return env
    want = TEMPLATE_NAME.lower()
    try:
        for name in sorted(os.listdir(_HERE)):
            if name.lower() == want:
                return os.path.join(_HERE, name)
    except OSError:
        pass
    raise ValueError("Không tìm thấy tệp khuôn '%s' trong thư mục chương trình."
                     % TEMPLATE_NAME)


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
    raise ValueError("Không nhận ra ô nhãn mẫu trong tệp khuôn %s." % TEMPLATE_NAME)


def _build_master(sample_tbl):
    """Dựng bảng khuôn SẠCH cho 1 trang (2 hàng x 2 cột, mọi ô = bản sao ô nhãn
    mẫu). Không nhân bản nguyên bảng của khuôn vì hàng nhãn thứ 2 có thể nằm
    dưới dạng BẢNG LỒNG trong ô (do thao tác dán trong Word) -> dựng phẳng lại
    từ ô mẫu. Trả (bảng, chiều cao hàng tw)."""
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
    """GHIM bề rộng bảng = tổng cột (dxa). Nếu khuôn để tblW='auto', Word co
    giãn cột theo nội dung -> 2 cột lệch nhau, viền không khớp ô (lỗi đã gặp)."""
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
    """Ghi đè nội dung + định dạng cho đoạn p (giữ nguyên phần tử đoạn của khuôn:
    w:spacing, và mọi thứ trong run đầu KHÔNG phải chữ — tức KHUNG VẼ neo ở
    đoạn 0 vẫn còn nguyên)."""
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
    thêm nếu thiếu, xóa bớt nếu thừa (giữ >=1 đoạn để ô hợp lệ — đoạn 0 mang
    khung vẽ nên luôn được giữ)."""
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


def _total_height(items, text_width_pt):
    return sum(_est_height(t, s, b, text_width_pt) for (t, s, b, _i) in items)


def _fill_to_bottom(body, foot, used, inner_pt, gap_lines, filler_pt):
    """Chèn dòng trống giữa `body` và `foot` sao cho khối `foot` nằm sát đáy ô,
    còn chừa ~`gap_lines` dòng hở bên dưới. Trả danh sách đoạn hoàn chỉnh."""
    empty_h = filler_pt * LINE_FACTOR
    room = inner_pt - used - gap_lines * empty_h
    filler = max(0, int(room / empty_h))
    return body + [("", filler_pt, False, False)] * filler + foot


def _sc(size_pt, scale):
    """Co cỡ chữ, làm tròn về bước 0.5pt (w:sz tính theo nửa point)."""
    return max(MIN_PT, round(size_pt * scale * 2) / 2.0)


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


def _header_items(tieu_de, scale):
    """Khối đầu nhãn: tên phông (cho phép nhiều dòng) + dòng kẻ "-----"."""
    out = []
    for line in _split(tieu_de):
        ls = line.lstrip()
        if ls.startswith("("):
            # (Từ ngày...) nghiêng
            out.append((line, _sc(SZ_DATE, scale), False, True))
        elif ls.upper().startswith("KHÓA") or ls.upper().startswith("KHOA"):
            # KHÓA... 13pt đậm
            out.append((line, _sc(SZ_KHOA, scale), True, False))
        else:
            # tên phông 14pt đậm
            out.append((line, _sc(SZ_HEADER, scale), True, False))
    out.append(("-----", SZ_SEP, False, False))
    return out


def _build_items(lb, scale):
    """Dựng các đoạn của 1 ô nhãn ở tỉ lệ cỡ chữ `scale`.
    Trả (khối trên, khối đáy) — khối đáy sẽ được đẩy xuống sát đáy ô."""
    body = [("", SZ_ANCHOR, False, False)]          # dòng trống đầu (khung vẽ)
    body += _header_items(lb["tieu_de"], scale)
    body += [("", SZ_HEADER, False, False)] * GAP_AFTER_SEP

    body.append(("HỘP SỐ", SZ_LABEL, True, False))
    body.append((_pad(lb["hop"], PAD_BOX), SZ_NUM, True, False))

    sl = _so_luong(lb)
    if sl:
        body.append(("GỒM %s ĐVBQ" % sl, SZ_INFO, True, False))
    body.append((_range_text(lb, PAD_FILE), SZ_INFO, True, False))

    nam = lb["nam"]
    foot = [("NĂM %s" % nam, SZ_YEAR, True, False)] if nam else []
    return body, foot


def _label_items(lb, inner_pt, text_width_pt):
    """Dàn nội dung 1 ô nhãn theo bản đồ định dạng của khuôn.

    Hàng nhãn có trHeight kiểu "exact", nên phần chữ vượt khổ ô sẽ bị Word CẮT
    MẤT mà không báo gì — nhãn hộp mất chữ là nhãn sai. Vì vậy nếu ước lượng
    thấy tràn thì tự co cỡ chữ tên phông cho vừa; các dòng cố định (số hộp khổ
    lớn, GỒM ... ĐVBQ, Từ ... đến ..., NĂM) giữ nguyên cỡ vì đó là phần phải
    đọc được từ xa."""
    for scale in SHRINK_STEPS:
        body, foot = _build_items(lb, scale)
        used = _total_height(body + foot, text_width_pt)
        if used <= inner_pt or scale == SHRINK_STEPS[-1]:
            return _fill_to_bottom(body, foot, used, inner_pt,
                                   BOTTOM_GAP_LINES, SZ_HEADER)


def _range_text(lb, pad):
    """'Từ 0001 đến 0014' — bỏ phần 'đến' nếu Excel không ghi cột Đến."""
    tu, den = _pad(lb["tu"], pad), _pad(lb["den"], pad)
    if tu and den:
        return "Từ %s đến %s" % (tu, den)
    return "Từ %s" % (tu or den)


def _so_luong(lb):
    """Số ĐVBQ: lấy cột 'Số lượng hồ sơ'; nếu trống thì suy ra từ Từ/Đến."""
    if lb["so_luong"]:
        return lb["so_luong"]
    tu, den = lb["tu"], lb["den"]
    if tu.isdigit() and den.isdigit() and int(den) >= int(tu):
        return str(int(den) - int(tu) + 1)
    return ""


# -------------------- Đọc dữ liệu Excel --------------------
def _txt(v):
    """Ô Excel -> chuỗi. Số nguyên lưu dạng float (13.0) phải ra '13', không
    phải '13.0'; ngày tháng ra dd/mm/yyyy."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, datetime.datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, datetime.date):
        return v.strftime("%d/%m/%Y")
    return str(v).strip()


def _year_txt(v):
    """Cột 'Năm': nếu người dùng nhập 2011 mà Excel hiểu thành ngày thì chỉ lấy
    năm, tránh in ra '01/01/2011'."""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return str(v.year)
    return _txt(v)


def _pad(v, width):
    """Đệm 0 cho số thứ tự ('14' -> '0014'); chuỗi không phải số giữ nguyên."""
    s = _txt(v)
    return s.zfill(width) if s.isdigit() else s


def _norm(s):
    return re.sub(r"\s+", "", (s or "").lower())


# Bản đồ tên cột (đã chuẩn hóa: bỏ trắng, chữ thường) -> khóa nội bộ.
_COL_ALIASES = (
    ("hop",      ("sốhộp", "hộpsố", "sohop", "hopso",
                  "cặphộpsố", "cặp(hộp)số", "caphopso", "cap(hop)so")),
    ("tu",       ("từhồsơ", "từhồsơsố", "tuhoso", "tuhososo", "từ", "tu")),
    ("den",      ("đếnhồsơ", "đếnhồsơsố", "denhoso", "denhososo", "đến", "den")),
    ("so_luong", ("sốlượnghồsơ", "soluonghoso", "sốlượng", "soluong",
                  "sốđvbq", "sodvbq", "sốlượngđvbq", "soluongdvbq")),
    ("tieu_de",  ("tênphông", "tenphong", "phông", "phong",
                  "tiêuđề", "tieude")),
    ("nam",      ("năm", "nam")),
)

# Vị trí cột khi sheet KHÔNG có dòng tiêu đề nhận ra được (đọc theo thứ tự A-F).
_DEFAULT_COLS = {"hop": 0, "tu": 1, "den": 2, "so_luong": 3,
                 "tieu_de": 4, "nam": 5}


def _map_header(row):
    """Nhận dạng dòng tiêu đề -> {khóa: chỉ số cột}. Trả {} nếu không phải
    dòng tiêu đề (nhận ra được ít nhất 2 cột thì mới coi là tiêu đề)."""
    cols = {}
    for i, v in enumerate(row):
        name = _norm(_txt(v))
        if not name:
            continue
        for key, aliases in _COL_ALIASES:
            if name in aliases and key not in cols:
                cols[key] = i
                break
    return cols if len(cols) >= 2 else {}


def read_labels(wb):
    """Đọc mọi sheet -> danh sách nhãn. Mỗi sheet tự nhận dòng tiêu đề và kiểu
    nhãn, nên một tệp Excel có thể trộn cả hai định dạng."""
    labels = []
    for name in wb.sheetnames:
        ws = wb[name]
        cols = None
        for row in ws.iter_rows(values_only=True):
            if cols is None:
                if not any(_txt(v) for v in row):
                    continue                    # dòng trống trước dòng tiêu đề
                found = _map_header(row)
                if found:
                    cols = found
                    continue                    # bỏ qua chính dòng tiêu đề
                cols = _DEFAULT_COLS

            def c(key, conv=_txt):
                i = cols.get(key)
                if i is None or i >= len(row):
                    return ""
                return conv(row[i])

            lb = {"hop": c("hop"), "tu": c("tu"), "den": c("den"),
                  "so_luong": c("so_luong"), "tieu_de": c("tieu_de"),
                  "nam": c("nam", _year_txt)}
            if not any(lb.values()):
                continue                        # dòng trống
            labels.append(lb)
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
                items = [("", SZ_ANCHOR, False, False)]
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
