# -*- coding: utf-8 -*-
"""
Chuyển các sheet "MỤC LỤC TÀI LIỆU" (chưa có ML) sang cấu trúc sheet "... - ML".
Cách dùng: python3 xuat-ml.py <file_vao.xlsx> [file_ra.xlsx]
  - Không truyền file_ra: ghi ra "<file_vao> - ML.xlsx" (không đụng file gốc).
  - Mọi sheet chưa có ML -> 1 sheet "<tên gốc> - ML" tương ứng trong file ra.
Số TT, ngày, tháng luôn đệm số 0 về 2 chữ số (1 -> 01, 5/3/2010 -> 05/03/2010).

Chỉ dùng thư viện chuẩn của Python (>= 3.6), KHÔNG cần openpyxl, KHÔNG cần venv.
"""
import re, sys, zipfile
import xml.etree.ElementTree as ET
from datetime import date, datetime, timedelta

# ============================ ĐỌC .XLSX ============================
NS  = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NSR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
# Mã định dạng ngày/giờ dựng sẵn của Excel
DATE_FMT_IDS = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))

def _col_index(ref):
    """'C15' -> 3"""
    n = 0
    for ch in ref:
        if ch.isdigit(): break
        n = n * 26 + (ord(ch.upper()) - 64)
    return n

def _is_date_fmt(code):
    code = re.sub(r'\[[^\]]*\]|"[^"]*"|\\.', "", code or "")
    return any(ch in code for ch in "ymdhs")

def _serial_to_dt(n, d1904):
    if d1904: return datetime(1904, 1, 1) + timedelta(days=n)
    if n < 60: n += 1                       # Excel coi 29/02/1900 là ngày có thật
    return datetime(1899, 12, 30) + timedelta(days=n)

def _si_text(el):
    for rph in el.findall(NS + "rPh"): el.remove(rph)
    return "".join(t.text or "" for t in el.iter(NS + "t"))

def _read_shared_strings(zf):
    if "xl/sharedStrings.xml" not in zf.namelist(): return []
    with zf.open("xl/sharedStrings.xml") as fp:
        return [_si_text(si) for si in ET.parse(fp).getroot() if si.tag == NS + "si"]

def _read_date_styles(zf):
    """Tập chỉ số style ứng với ô định dạng ngày tháng."""
    if "xl/styles.xml" not in zf.namelist(): return set()
    with zf.open("xl/styles.xml") as fp:
        root = ET.parse(fp).getroot()
    custom = {int(f.get("numFmtId")): f.get("formatCode")
              for f in root.iter(NS + "numFmt")}
    out, xfs = set(), root.find(NS + "cellXfs")
    if xfs is None: return out
    for i, xf in enumerate(xfs):
        fid = int(xf.get("numFmtId") or 0)
        if fid in DATE_FMT_IDS or (fid in custom and _is_date_fmt(custom[fid])):
            out.add(i)
    return out

def _cell_value(el, shared, date_styles, d1904):
    t = el.get("t")
    if t == "inlineStr":
        node = el.find(NS + "is")
        return _si_text(node) if node is not None else None
    v = el.find(NS + "v")
    if v is None or v.text is None: return None
    s = v.text
    if t == "s":  return shared[int(s)] if int(s) < len(shared) else ""
    if t == "b":  return s == "1"
    if t in ("str", "e"): return s
    try: num = float(s)
    except ValueError: return s
    si = el.get("s")
    if si is not None and int(si) in date_styles:
        return _serial_to_dt(num, d1904)
    return int(num) if num.is_integer() else num

class Sheet:
    """Sheet đã đọc vào bộ nhớ: rows[r][c] theo chỉ số 0."""
    def __init__(self, title, rows): self.title, self.rows = title, rows
    @property
    def max_row(self): return len(self.rows)
    def cell(self, r, c):
        """Giá trị ô theo chỉ số 1 (trả None nếu ngoài vùng)."""
        if 1 <= r <= len(self.rows):
            row = self.rows[r - 1]
            if 1 <= c <= len(row): return row[c - 1]
        return None

def _read_sheet(zf, part, title, shared, date_styles, d1904):
    rows = []
    with zf.open(part) as fp:
        cells = []
        for _, el in ET.iterparse(fp, events=("end",)):
            tag = el.tag[el.tag.rfind("}") + 1:]
            if tag == "c":
                ref = el.get("r")
                ci = _col_index(ref) if ref else (cells[-1][0] + 1 if cells else 1)
                cells.append((ci, _cell_value(el, shared, date_styles, d1904)))
            elif tag == "row":
                ri = int(el.get("r") or len(rows) + 1)
                while len(rows) < ri - 1: rows.append([])
                width = max((c for c, _ in cells), default=0)
                row = [None] * width
                for ci, v in cells:
                    if 1 <= ci <= width: row[ci - 1] = v
                rows.append(row)
                cells = []
                el.clear()
    while rows and not any(v is not None and str(v).strip() != "" for v in rows[-1]):
        rows.pop()
    return Sheet(title, rows)

def doc_xlsx(path):
    """Đọc file .xlsx -> danh sách Sheet, theo đúng thứ tự trong workbook."""
    with zipfile.ZipFile(path) as zf:
        shared, date_styles = _read_shared_strings(zf), _read_date_styles(zf)
        with zf.open("xl/workbook.xml") as fp:
            wb = ET.parse(fp).getroot()
        pr = wb.find(NS + "workbookPr")
        d1904 = pr is not None and (pr.get("date1904") in ("1", "true"))
        with zf.open("xl/_rels/workbook.xml.rels") as fp:
            rels = {r.get("Id"): r.get("Target") for r in ET.parse(fp).getroot()}
        out = []
        for sh in wb.iter(NS + "sheet"):
            target = rels.get(sh.get(NSR + "id"), "")
            part = target[1:] if target.startswith("/") else "xl/" + target.lstrip("./")
            if part not in zf.namelist(): continue
            out.append(_read_sheet(zf, part, sh.get("name"), shared, date_styles, d1904))
        return out

# ============================ GHI .XLSX ============================
S_MAC, S_TIEUDE, S_HEADER, S_GIUA, S_TRAI = 0, 1, 2, 3, 4   # chỉ số style

_ESC = {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "\n": "&#10;"}
_XAU = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

def _esc(s):
    return "".join(_ESC.get(ch, ch) for ch in _XAU.sub("", str(s)).replace("\r\n", "\n").replace("\r", "\n"))

def _col_letter(n):
    s = ""
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><sz val="13"/><name val="Times New Roman"/><family val="1"/></font>
<font><b/><sz val="13"/><name val="Times New Roman"/><family val="1"/></font>
</fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="5">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="49" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="49" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="49" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="49" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

class SheetOut:
    def __init__(self, title):
        self.title, self.widths = title, []
        self.cells, self.heights, self.merges = {}, {}, []
    def dat(self, r, c, value=None, style=S_MAC):
        self.cells.setdefault(r, {})[c] = (value, style)
    def gop(self, r, c1, c2):
        self.merges.append("%s%d:%s%d" % (_col_letter(c1), r, _col_letter(c2), r))
    @property
    def max_row(self): return max(self.cells) if self.cells else 0

    def _chunks(self):
        yield '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        yield '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        if self.widths:
            yield "<cols>%s</cols>" % "".join(
                '<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i, i, w)
                for i, w in enumerate(self.widths, 1))
        yield "<sheetData>"
        for r in sorted(self.cells):
            ht = self.heights.get(r)
            buf = ['<row r="%d"%s>' % (r, ' ht="%s" customHeight="1"' % ht if ht else "")]
            for c in sorted(self.cells[r]):
                v, st = self.cells[r][c]
                ref = "%s%d" % (_col_letter(c), r)
                if v is None or v == "":
                    buf.append('<c r="%s" s="%d"/>' % (ref, st))
                else:
                    buf.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                               % (ref, st, _esc(v)))
            buf.append("</row>")
            yield "".join(buf)
        yield "</sheetData>"
        if self.merges:
            yield '<mergeCells count="%d">%s</mergeCells>' % (
                len(self.merges), "".join('<mergeCell ref="%s"/>' % m for m in self.merges))
        yield "</worksheet>"

class WorkbookOut:
    def __init__(self): self.sheets = []
    def create_sheet(self, title):
        sh = SheetOut(title); self.sheets.append(sh); return sh

    def save(self, path):
        n = len(self.sheets)
        parts = ["/xl/worksheets/sheet%d.xml" % i for i in range(1, n + 1)]
        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            + "".join('<Override PartName="%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % p for p in parts)
            + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>")
        root_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>")
        wb_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i)
                      for i in range(1, n + 1))
            + '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (n + 1)
            + "</Relationships>")
        workbook = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
            + "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (_esc(sh.title), i, i)
                      for i, sh in enumerate(self.sheets, 1))
            + "</sheets></workbook>")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types)
            zf.writestr("_rels/.rels", root_rels)
            zf.writestr("xl/workbook.xml", workbook)
            zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
            zf.writestr("xl/styles.xml", STYLES_XML)
            for i, sh in enumerate(self.sheets, 1):
                info = zipfile.ZipInfo("xl/worksheets/sheet%d.xml" % i)
                info.compress_type = zipfile.ZIP_DEFLATED
                with zf.open(info, "w") as fp:          # ghi theo luồng, ít tốn bộ nhớ
                    for chunk in sh._chunks(): fp.write(chunk.encode("utf-8"))

# ========================= LOGIC NGHIỆP VỤ =========================
TITLE = "MỤC LỤC VĂN BẢN, TÀI LIỆU"
HEADERS = ["Số\nTT", "Số, ký\nhiệu", "Ngày tháng", "Tên loại và trích yếu", "Tác giả",
           "Người ký", "Độ mật", "Loại bản", "Trang số", "Số\ntrang", "Từ\nkhoá",
           "Ghi\nchú", "Số lượng tệp (file)", "Tên tệp tài liệu"]
NCOL = len(HEADERS)
COL_WIDTHS_PX = [35, 77, 77, 224, 91, 84, 56, 56, 56, 49, 49, 49, 56, 98]  # theo sheet mẫu
LEFT_ALIGN_COLS = {4}
IDX_STT, IDX_NGAY = 0, 2

def text(v):
    if v is None: return ""
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return str(v).strip()

def pad2(n): return str(int(n)).zfill(2)

def chuan_hoa_ngay(v):
    """dd/mm/yyyy (hoặc mm/yyyy, yyyy) với ngày/tháng 2 chữ số; không nhận dạng được -> giữ nguyên."""
    if v is None: return ""
    if isinstance(v, (datetime, date)):
        return f"{pad2(v.day)}/{pad2(v.month)}/{v.year}"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        s = str(int(v))
        if len(s) in (7, 8):                       # ddmmyyyy lưu dạng số
            s8 = s.zfill(8); d, m, y = int(s8[:2]), int(s8[2:4]), s8[4:]
            if 1 <= d <= 31 and 1 <= m <= 12: return f"{pad2(d)}/{pad2(m)}/{y}"
        return s
    s = text(v)
    m = re.fullmatch(r"(\d{1,2})\s*[/.\-]\s*(\d{1,2})\s*[/.\-]\s*(\d{4})", s)
    if m: return f"{pad2(m[1])}/{pad2(m[2])}/{m[3]}"
    m = re.fullmatch(r"(\d{1,2})\s*[/.\-]\s*(\d{4})", s)
    if m: return f"{pad2(m[1])}/{m[2]}"
    m = re.fullmatch(r"(\d{1,2})(\d{2})(\d{4})", s)
    if m and 1 <= int(m[1]) <= 31 and 1 <= int(m[2]) <= 12:
        return f"{pad2(m[1])}/{m[2]}/{m[3]}"
    return s

def chuan_hoa_stt(v):   s = text(v); return s.zfill(2) if s.isdigit() else s
def chuan_hoa_ho_so(v): s = text(v); return s.zfill(4) if s.isdigit() else s

def ma_ho_so(sheet_name):
    m = re.search(r"(\d{4}-\d{2})", sheet_name); return m.group(1) if m else ""

def ten_sheet_ml(sheet_name): return (sheet_name + " - ML")[:31]

def find_header_row(ws):
    for r in range(1, 20):
        a = text(ws.cell(r, 1)).lower()
        if a.startswith("hồ sơ số") or a.startswith("số hồ sơ"): return r
    raise ValueError(f"Sheet '{ws.title}': không tìm thấy dòng tiêu đề 'Hồ sơ số'")

def read_source(ws):
    start = find_header_row(ws) + 1
    if text(ws.cell(start, 2)).startswith("("): start += 1            # dòng (3),(4)...
    groups, order, current = {}, [], None
    for row in ws.rows[start - 1:]:
        row = list(row) + [None] * (15 - len(row))
        if all(text(x) == "" for x in row[:15]): continue
        if text(row[0]) != "":
            current = chuan_hoa_ho_so(row[0])
            if current not in groups: groups[current] = []; order.append(current)
        if current is None: continue
        data = [text(x) for x in row[1:15]]
        data[IDX_STT]  = chuan_hoa_stt(row[1])
        data[IDX_NGAY] = chuan_hoa_ngay(row[3])
        groups[current].append(data)
    return [(k, groups[k]) for k in order]

def write_block(ws, row, ma, so_hs, rows):
    def title_line(r, s):
        ws.dat(r, 1, s, S_TIEUDE)
        ws.gop(r, 1, NCOL)
        ws.heights[r] = 18
    title_line(row, TITLE); title_line(row + 1, f"Mã hồ sơ: {ma}"); title_line(row + 2, f"Hồ sơ số: {so_hs}")
    r = row + 3
    for c, h in enumerate(HEADERS, 1):
        ws.dat(r, c, h, S_HEADER)
    ws.heights[r] = 48
    r += 1
    for data in rows:
        for c, v in enumerate(data, 1):
            ws.dat(r, c, v if v != "" else None, S_TRAI if c in LEFT_ALIGN_COLS else S_GIUA)
        r += 1
    return r + 1                                    # 1 dòng trống ngăn khối

def build_ml_sheet(wb_out, src_ws):
    ws = wb_out.create_sheet(ten_sheet_ml(src_ws.title))
    ws.widths = [round(px / 7, 2) for px in COL_WIDTHS_PX]
    ma, row, n_hs = ma_ho_so(src_ws.title), 1, 0
    for so_hs, rows in read_source(src_ws):
        row = write_block(ws, row, ma, so_hs, rows); n_hs += 1
    return ws, n_hs

def main(path_in, path_out=None):
    if path_out is None:
        path_out = re.sub(r"\.xlsx$", "", path_in, flags=re.I) + " - ML.xlsx"
    wb_out = WorkbookOut()
    for ws in doc_xlsx(path_in):
        if ws.title.rstrip().endswith("- ML"): continue          # đã là sheet ML -> bỏ qua
        out, n_hs = build_ml_sheet(wb_out, ws)
        print(f"{ws.title!r} -> {out.title!r}: {n_hs} hồ sơ, {out.max_row} dòng")
    wb_out.save(path_out); print("Đã ghi:", path_out)

if __name__ == "__main__":
    if len(sys.argv) < 2: print(__doc__); sys.exit(1)
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
