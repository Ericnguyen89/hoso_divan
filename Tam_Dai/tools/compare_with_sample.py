#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tools/compare_with_sample.py — So sánh NỘI DUNG CHỮ giữa tệp Word do muc_luc_ho_so.py tạo ra và
tệp Word mẫu ("ĐU Tam Đại Vĩnh viễn.docx", vốn là bản PDF→Word của bản in Excel nên bị mất một số
dấu cách trong từ và tách bảng thành nhiều bảng với hàng tiêu đề lặp lại).

Cách so sánh
  * Bảng: gom các ô theo LOẠI bảng — chữ viết tắt (4 cột, ô đầu "SỐ TT"), phụ lục tra cứu (4 cột,
    "Đề mục"), bảng thống kê hồ sơ (13 cột) — nối các bảng bị tách của mẫu lại, bỏ hàng tiêu đề /
    hàng đánh số "(n)" lặp lại, rồi so từng dòng theo thứ tự (difflib để nhận ra dòng thiếu/thừa).
  * Đoạn văn (bìa, chứng từ kết thúc, tựa bảng): mọi đoạn ngoài các bảng trên + ô của bảng không
    thuộc 3 loại (bảng bìa / chữ ký của ta). Mỗi đoạn của ta phải xuất hiện trong mẫu; mỗi đoạn của
    mẫu phải xuất hiện trong của ta hoặc ghép được từ các đoạn của ta (PDF hay gộp dòng).
  * Chuẩn hoá: NFC, BỎ MỌI KHOẢNG TRẮNG, chữ thường.

Cách dùng
  python3 tools/compare_with_sample.py                 # chuyển mẫu .xlsx trong bộ nhớ rồi so với .docx mẫu
  python3 tools/compare_with_sample.py --generated ket_qua/x.docx
  python3 tools/compare_with_sample.py --xlsx A.xlsx --reference A_mau.docx [-v]
Mã thoát: 0 = không khác biệt; 1 = có dòng bảng khác / đoạn chữ thiếu; 2 = lỗi đầu vào.
"""
from __future__ import annotations

import argparse
import difflib
import io
import os
import re
import sys
import unicodedata
import zipfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
NUMBERING = re.compile(r"^\(\d+\)$")
# (tên loại, số cột, khoá ô đầu của hàng tiêu đề)
TABLE_TYPES = (("abbreviations", 4, "sốtt"), ("appendix", 4, "đềmục"), ("main", 13, "sốcặp(hộp)"))
SAMPLE_DIR = PROJECT   # tệp mẫu đặt ngay trong Tam_Dai/
DEFAULT_XLSX = os.path.join(SAMPLE_DIR, "ĐU Tam Đại Vĩnh viễn.xlsx")
DEFAULT_REF = os.path.join(SAMPLE_DIR, "ĐU Tam Đại Vĩnh viễn.docx")


def qn(tag: str) -> str:
    """'w:p' hoặc 'p' → '{ns}p'."""
    return "{%s}%s" % (W, tag.split(":")[-1])


def key(text: str) -> str:
    """Khoá so sánh: NFC, bỏ mọi khoảng trắng, chữ thường."""
    return re.sub(r"\s+", "", unicodedata.normalize("NFC", text or "")).lower()


def para_text(p) -> str:
    buf = []
    for node in p.iter(qn("w:t"), qn("w:br"), qn("w:tab")):
        if node.tag == qn("w:t"):
            buf.append(node.text or "")
        elif node.getparent() is not None and node.getparent().tag == qn("w:r"):
            buf.append("\n" if node.tag == qn("w:br") else "\t")
    return "".join(buf)


def row_texts(tr) -> list[str]:
    """Chữ các ô của một hàng; ô gộp ngang (gridSpan=n) được đệm thêm n-1 ô trống."""
    out = []
    for tc in tr.findall("w:tc", NS):
        out.append("\n".join(para_text(p) for p in tc.findall("w:p", NS)))
        span = tc.find("w:tcPr/w:gridSpan", NS)
        if span is not None:
            out.extend([""] * (int(span.get(qn("w:val"))) - 1))
    return out


def load_root(src):
    data = src if isinstance(src, (bytes, bytearray)) else open(src, "rb").read()
    return etree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("word/document.xml"))


def extract(root) -> list:
    """Các khối của thân tài liệu theo thứ tự: ('p', text) hoặc ('tbl', số cột, [[ô]])."""
    items = []
    for child in root.find("w:body", NS):
        if child.tag == qn("w:p"):
            items.append(("p", para_text(child)))
        elif child.tag == qn("w:tbl"):
            ncols = len(child.findall("w:tblGrid/w:gridCol", NS))
            rows = [row_texts(tr) for tr in child.findall("w:tr", NS)]
            items.append(("tbl", ncols, rows))
    return items


def split_tables(items):
    """→ (dòng dữ liệu theo loại bảng, số bảng theo loại, các đoạn chữ khác theo thứ tự)."""
    by_type = {name: [] for name, _, _ in TABLE_TYPES}
    n_tables = {name: 0 for name, _, _ in TABLE_TYPES}
    texts = []
    for item in items:
        if item[0] == "p":
            if key(item[1]):
                texts.append(item[1])
            continue
        _, ncols, rows = item
        first = key(rows[0][0]) if rows and rows[0] else ""
        typ = next((name for name, n, k in TABLE_TYPES if ncols == n and first.startswith(k)), None)
        if typ is None:                                    # bảng bìa / chữ ký của ta: từng đoạn một
            texts.extend(line for r in rows for t in r for line in t.split("\n") if key(line))
            continue
        n_tables[typ] += 1
        header_key = next(k for name, _, k in TABLE_TYPES if name == typ)
        for r in rows:
            keys = [key(c) for c in r]
            nonempty = [k for k in keys if k]
            if not nonempty:
                continue
            if keys[0].startswith(header_key) or all(NUMBERING.match(k) for k in nonempty):
                continue                                   # hàng tiêu đề / đánh số lặp lại
            by_type[typ].append(r)
    return by_type, n_tables, texts


def fmt_row(row, width=70) -> str:
    cells = []
    for c in row:
        c = re.sub(r"\s+", " ", c).strip()
        cells.append(c if len(c) <= width else c[:width - 1] + "…")
    return " | ".join(cells)


def compare_rows(name, expected, got, verbose=False):
    ek = [tuple(key(c) for c in r) for r in expected]
    gk = [tuple(key(c) for c in r) for r in got]
    sm = difflib.SequenceMatcher(None, ek, gk, autojunk=False)
    identical, diffs = 0, []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            identical += i2 - i1
        elif tag == "replace":
            for k in range(max(i2 - i1, j2 - j1)):
                e = expected[i1 + k] if i1 + k < i2 else None
                g = got[j1 + k] if j1 + k < j2 else None
                diffs.append((i1 + k if e is not None else None, j1 + k if g is not None else None, e, g))
        elif tag == "delete":
            diffs.extend((i, None, expected[i], None) for i in range(i1, i2))
        elif tag == "insert":
            diffs.extend((None, j, None, got[j]) for j in range(j1, j2))
    for ri, gi, e, g in diffs:
        what = ("THIẾU trong tệp sinh ra" if g is None else
                "THỪA trong tệp sinh ra" if e is None else "KHÁC")
        cols = ""
        if e is not None and g is not None:
            bad = [str(c + 1) for c, (a, b) in enumerate(zip(ek[ri], gk[gi])) if a != b]
            if len(e) != len(g):
                bad.append("số ô %d≠%d" % (len(e), len(g)))
            cols = " (cột %s)" % ", ".join(bad)
        print("  %s%s — mẫu dòng %s / sinh ra dòng %s" % (what, cols, ri, gi))
        if e is not None:
            print("     mong đợi: %s" % fmt_row(e))
        if g is not None:
            print("     nhận được: %s" % fmt_row(g))
    return identical, len(diffs)


def composed_of(target: str, parts: set) -> bool:
    """target có ghép được từ các khoá trong `parts` (theo thứ tự bất kỳ, có thể lặp)?"""
    n = len(target)
    reach = [False] * (n + 1)
    reach[0] = True
    for i in range(n):
        if not reach[i]:
            continue
        for p in parts:
            if target.startswith(p, i):
                reach[i + len(p)] = True
    return reach[n]


def compare_texts(ref_texts, gen_texts, verbose=False):
    ref_blob = "".join(key(t) for t in ref_texts)
    gen_blob = "".join(key(t) for t in gen_texts)
    gen_keys = {key(t) for t in gen_texts if key(t)}
    missing_in_ref = [t for t in gen_texts if key(t) not in ref_blob]
    exact, composed, missing_in_gen = 0, [], []
    for t in ref_texts:
        k = key(t)
        if k in gen_blob:
            exact += 1
        elif composed_of(k, gen_keys):
            composed.append(t)
        else:
            missing_in_gen.append(t)
    print("== Đoạn chữ ngoài bảng (bìa, tựa đề, chứng từ kết thúc) ==")
    print("  tệp sinh ra: %d đoạn, %d có trong mẫu, %d KHÔNG có trong mẫu"
          % (len(gen_texts), len(gen_texts) - len(missing_in_ref), len(missing_in_ref)))
    for t in missing_in_ref:
        print("     KHÔNG CÓ TRONG MẪU: %r" % re.sub(r"\s+", " ", t).strip())
    print("  tệp mẫu: %d đoạn, %d có nguyên văn trong tệp sinh ra, %d ghép được từ các đoạn của tệp sinh ra "
          "(PDF gộp dòng), %d KHÔNG tìm thấy"
          % (len(ref_texts), exact, len(composed), len(missing_in_gen)))
    for t in composed:
        print("     ghép từ nhiều đoạn: %r" % re.sub(r"\s+", " ", t).strip())
    for t in missing_in_gen:
        print("     KHÔNG TÌM THẤY: %r" % re.sub(r"\s+", " ", t).strip())
    if verbose:
        print("  -- đoạn của tệp sinh ra --")
        for t in gen_texts:
            print("     ", re.sub(r"\s+", " ", t).strip())
        print("  -- đoạn của tệp mẫu --")
        for t in ref_texts:
            print("     ", re.sub(r"\s+", " ", t).strip())
    return len(missing_in_ref), len(missing_in_gen)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--generated", help="tệp .docx đã sinh ra (mặc định: chuyển --xlsx trong bộ nhớ)")
    ap.add_argument("--xlsx", default=DEFAULT_XLSX, help="tệp Excel để chuyển khi không có --generated")
    ap.add_argument("--reference", default=DEFAULT_REF, help="tệp .docx mẫu")
    ap.add_argument("-v", "--verbose", action="store_true", help="in mọi đoạn chữ của hai tệp")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.reference):
        print("Không thấy tệp mẫu: %s" % args.reference, file=sys.stderr)
        return 2
    if args.generated:
        gen_src = args.generated
        print("Tệp sinh ra : %s" % gen_src)
    else:
        sys.path.insert(0, PROJECT)
        import muc_luc_ho_so
        if not os.path.isfile(args.xlsx):
            print("Không thấy tệp Excel: %s" % args.xlsx, file=sys.stderr)
            return 2
        gen_src = muc_luc_ho_so.convert_workbook(args.xlsx)
        print("Tệp sinh ra : chuyển từ %s (trong bộ nhớ, %d byte)" % (args.xlsx, len(gen_src)))
    print("Tệp mẫu     : %s" % args.reference)

    ref_tables, ref_n, ref_texts = split_tables(extract(load_root(args.reference)))
    gen_tables, gen_n, gen_texts = split_tables(extract(load_root(gen_src)))

    labels = {"abbreviations": "Bảng chữ viết tắt (4 cột)", "appendix": "Phụ lục tra cứu (4 cột)",
              "main": "Bảng thống kê hồ sơ (13 cột)"}
    summary, failed = [], False
    for name, _, _ in TABLE_TYPES:
        exp, got = ref_tables[name], gen_tables[name]
        print("== %s == mẫu: %d bảng, %d dòng dữ liệu; sinh ra: %d bảng, %d dòng dữ liệu"
              % (labels[name], ref_n[name], len(exp), gen_n[name], len(got)))
        identical, ndiff = compare_rows(name, exp, got, args.verbose)
        print("  so sánh %d dòng: %d giống, %d khác" % (max(len(exp), len(got)), identical, ndiff))
        summary.append("%s %d/%d giống" % (name, identical, max(len(exp), len(got))))
        failed |= ndiff > 0
    miss_ref, miss_gen = compare_texts(ref_texts, gen_texts, args.verbose)
    failed |= miss_ref > 0 or miss_gen > 0
    summary.append("đoạn chữ: %d của ta thiếu trong mẫu, %d của mẫu không tìm thấy" % (miss_ref, miss_gen))
    print("SUMMARY: " + "; ".join(summary) + (" — CÓ KHÁC BIỆT" if failed else " — KHỚP"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
