# -*- coding: utf-8 -*-
"""
Web app (Flask) chuyển "MỤC LỤC HỒ SƠ" từ Excel (.xlsx) sang Word (.docx)
hàng loạt — mỗi tệp Excel cho ra một tệp Word khổ ngang gồm: Tờ bìa, Bảng chữ
viết tắt, Phụ lục tra cứu, Bảng thống kê hồ sơ và Chứng từ kết thúc.

- Chọn / kéo thả NHIỀU tệp .xlsx (hoặc tệp .zip chứa .xlsx — máy chủ tự giải nén).
- Xử lý tuần tự ở luồng nền, tiến độ realtime, tải từng tệp hoặc tất cả (.zip).
- Phần chuyển đổi nằm trong mô-đun `muc_luc_ho_so.py` (cùng thư mục).

Chạy:
    python3 app.py
(Lần đầu thiếu thư viện, app sẽ tự cài rồi chạy lại.)
Mở:  http://127.0.0.1:5039        (đổi cổng/địa chỉ bằng biến môi trường PORT / HOST)
Khi chạy trong server.py (cổng chung), app được gắn tại /tamdai/ — tab "IN DỰ ÁN TAM ĐẠI".
"""

import io
import os
import re
import sys
import uuid
import logging
import zipfile
import datetime
import threading
import subprocess
import unicodedata
import importlib.util
from urllib.parse import quote


def ensure_deps():
    """Tự cài các thư viện còn thiếu bằng đúng Python đang chạy, rồi nạp lại.
    Nhờ vậy app chạy được mà không cần tạo/kích hoạt môi trường ảo (venv)."""
    need = []
    for module, pip_name in (("flask", "flask"),
                             ("docx", "python-docx"),
                             ("openpyxl", "openpyxl")):
        if importlib.util.find_spec(module) is None:
            need.append(pip_name)
    if not need:
        return
    print("• Thiếu thư viện: %s — đang tự cài đặt..." % ", ".join(need))
    base = [sys.executable, "-m", "pip", "install", "--quiet"]
    # Trong venv: cài thẳng. Ngoài venv: ưu tiên user site, rồi phá khóa PEP668.
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
    # Nạp lại tiến trình để dùng thư viện vừa cài
    os.execv(sys.executable, [sys.executable] + sys.argv)


ensure_deps()

from flask import (Flask, Response, request, jsonify, send_file)  # noqa: E402
from werkzeug.exceptions import RequestEntityTooLarge, NotFound  # noqa: E402

# Mô-đun lõi: đọc Excel -> dựng docx (cùng thư mục với app.py). Khi gộp vào
# server.py (tab "IN DỰ ÁN TAM ĐẠI", mount tại /tamdai) app này vẫn giữ nguyên.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import muc_luc_ho_so  # noqa: E402

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("muc_luc_ho_so.web")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024      # 500 MB mỗi lượt tải lên
try:                                                       # JSON giữ nguyên tiếng Việt
    app.json.ensure_ascii = False                          # Flask >= 2.3
except AttributeError:                                     # Flask cũ hơn
    app.config["JSON_AS_ASCII"] = False

JOB_TTL_SECONDS = 2 * 3600            # job tự xóa sau 2 giờ
MAX_FILES_PER_JOB = 500               # tránh một lượt quá lớn làm đầy RAM
MAX_EXPANDED_BYTES = 1024 * 1024 * 1024  # tổng dung lượng .xlsx giải nén từ .zip
ALLOWED_EXT = (".xlsx", ".xlsm")
DOCX_MIME = ("application/vnd.openxmlformats-officedocument."
             "wordprocessingml.document")
ZIP_MIME = "application/zip"

# Bộ nhớ tạm cho công việc: job_id -> dict trạng thái + kết quả (chỉ 1 tiến trình)
JOBS = {}
JOBS_LOCK = threading.Lock()


# ----------------------------------------------------------------------------
# Tiện ích
# ----------------------------------------------------------------------------
def cleanup():
    """Xóa job đã xong quá 2 giờ (job còn chạy giữ lại tối đa 6 giờ)."""
    now = datetime.datetime.now()
    with JOBS_LOCK:
        for k in list(JOBS):
            job = JOBS[k]
            age = (now - job["created"]).total_seconds()
            if age > JOB_TTL_SECONDS and job["status"] != "running":
                JOBS.pop(k, None)
            elif age > 3 * JOB_TTL_SECONDS:
                JOBS.pop(k, None)


def _default_options():
    opts = dict(getattr(muc_luc_ho_so, "DEFAULT_OPTIONS", None) or {})
    opts.setdefault("page_size", "A4")
    opts.setdefault("page_number_start", 2)
    opts.setdefault("table_font_pt", 8.5)
    return opts


def _parse_options(form):
    """Đọc tùy chọn từ form; sai -> ValueError (thông báo tiếng Việt)."""
    opts = _default_options()

    v = (form.get("page_number_start") or "").strip()
    if v:
        if v not in ("1", "2"):
            raise ValueError("Cách đánh số trang không hợp lệ (chỉ nhận 1 hoặc 2).")
        opts["page_number_start"] = int(v)

    v = (form.get("page_size") or "").strip()
    if v:
        sizes = {"a4": "A4", "letter": "Letter"}
        if v.lower() not in sizes:
            raise ValueError("Khổ giấy không hợp lệ (chỉ nhận A4 hoặc Letter).")
        opts["page_size"] = sizes[v.lower()]

    v = (form.get("table_font_pt") or "").strip().replace(",", ".")
    if v:
        try:
            f = float(v)
        except ValueError:
            raise ValueError("Cỡ chữ bảng không hợp lệ.")
        if not (6 <= f <= 14):
            raise ValueError("Cỡ chữ bảng phải nằm trong khoảng 6–14 pt.")
        opts["table_font_pt"] = f
    return opts


def _zip_entry_name(info):
    """Tên mục trong zip: nếu không có cờ UTF-8 thì thử giải mã lại (zip Windows)."""
    name = info.filename
    if not (info.flag_bits & 0x800):
        try:
            name = name.encode("cp437").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return name


def _expand_zip(zip_name, data):
    """Lấy các tệp .xlsx bên trong một tệp .zip; bỏ qua __MACOSX và tệp khác."""
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise ValueError("Tệp '%s' không phải tệp .zip hợp lệ." % zip_name)
    out, total = [], 0
    with zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            path = _zip_entry_name(info).replace("\\", "/")
            parts = path.split("/")
            base = parts[-1].strip()
            if ("__MACOSX" in parts[:-1] or base.startswith("._")
                    or base.startswith(".") or not base):
                continue
            if not base.lower().endswith(ALLOWED_EXT):
                continue
            total += info.file_size
            if total > MAX_EXPANDED_BYTES:
                raise ValueError("Tệp '%s' giải nén quá lớn (hơn %d MB)."
                                 % (zip_name, MAX_EXPANDED_BYTES // (1024 * 1024)))
            try:
                blob = zf.read(info)
            except Exception:
                raise ValueError("Không giải nén được '%s' trong '%s' "
                                 "(zip có mật khẩu hoặc bị hỏng?)." % (base, zip_name))
            out.append((base, blob))
    return out


def _collect_uploads(uploads):
    """Trả về ([(tên, bytes)], [tên bị bỏ qua]) từ danh sách tệp tải lên."""
    files, skipped = [], []
    for f in uploads:
        if not f or not f.filename:
            continue
        name = os.path.basename(f.filename.replace("\\", "/")).strip()
        low = name.lower()
        if low.endswith(ALLOWED_EXT):
            files.append((name, f.read()))
        elif low.endswith(".zip"):
            found = _expand_zip(name, f.read())
            if not found:
                skipped.append("%s (không có .xlsx bên trong)" % name)
            files.extend(found)
        else:
            raise ValueError("Tệp '%s' không phải .xlsx hoặc .zip." % name)
    return files, skipped


def _output_name(name):
    try:
        out = muc_luc_ho_so.output_name_for(name)
    except Exception:
        out = None
    if not out:
        out = os.path.splitext(name)[0] + ".docx"
    return out


def _unique_names(names):
    """Tránh trùng tên tệp kết quả: 'a.docx', 'a (2).docx', ..."""
    seen, result = {}, []
    for nm in names:
        key = nm.lower()
        if key in seen:
            seen[key] += 1
            root, ext = os.path.splitext(nm)
            nm = "%s (%d)%s" % (root, seen[key], ext)
        else:
            seen[key] = 1
        result.append(nm)
    return result


def _get_job(job_id):
    with JOBS_LOCK:
        return JOBS.get(job_id)


def _public_view(job):
    with JOBS_LOCK:
        items = [dict(it) for it in job["items"]]
        ok = sum(1 for it in items if it["status"] == "ok")
        err = sum(1 for it in items if it["status"] == "error")
        total = job["total"]
        return {
            "ok": True,
            "job_id": job["id"],
            "status": job["status"],
            "total": total,
            "done": job["done"],
            "current": job["current"],
            "ok_count": ok,
            "error_count": err,
            "percent": round(job["done"] * 100.0 / total, 1) if total else 100.0,
            "options": job["options"],
            "items": items,
        }


def _attachment(data, name, mimetype):
    """Trả tệp tải về; Content-Disposition mang tên UTF-8 (filename*=UTF-8'')."""
    resp = send_file(io.BytesIO(data), mimetype=mimetype,
                     as_attachment=True, download_name=name)
    ascii_name = name.replace("Đ", "D").replace("đ", "d")
    ascii_name = unicodedata.normalize("NFKD", ascii_name)
    ascii_name = ascii_name.encode("ascii", "ignore").decode("ascii")
    ascii_name = re.sub(r"[^A-Za-z0-9._ ()-]+", "_", ascii_name).strip() or "download"
    resp.headers["Content-Disposition"] = (
        "attachment; filename=\"%s\"; filename*=UTF-8''%s"
        % (ascii_name, quote(name, safe="")))
    resp.headers["Cache-Control"] = "no-store"
    return resp


def _json_404(msg):
    return jsonify(ok=False, error=msg), 404


# ----------------------------------------------------------------------------
# Luồng nền: chuyển đổi tuần tự từng tệp
# ----------------------------------------------------------------------------
def _process_job(job):
    options = job["options"]
    for i, item in enumerate(job["items"]):
        with JOBS_LOCK:
            item["status"] = "running"
            job["current"] = item["name"]
            blob = job["inputs"][i]
            job["inputs"][i] = None            # giải phóng RAM đầu vào
        try:
            data = muc_luc_ho_so.convert_workbook(blob, options)
            if not isinstance(data, (bytes, bytearray)):
                data = data.getvalue() if hasattr(data, "getvalue") else bytes(data)
            data = bytes(data)
            if not data.startswith(b"PK"):
                raise RuntimeError("kết quả trả về không phải tệp .docx")
            with JOBS_LOCK:
                job["outputs"][i] = data
                item["status"] = "ok"
                item["size"] = len(data)
                item["error"] = None
            log.info("OK   %s -> %s (%d bytes)", item["name"], item["out_name"], len(data))
        except muc_luc_ho_so.ConvertError as e:
            msg = str(e).strip() or "Tệp Excel không đúng mẫu Mục lục hồ sơ."
            with JOBS_LOCK:
                item["status"] = "error"
                item["error"] = msg
            log.warning("LỖI  %s: %s", item["name"], msg)
        except Exception as e:  # không để một tệp hỏng làm dừng cả đợt
            log.exception("Lỗi không mong đợi khi chuyển '%s'", item["name"])
            detail = ("%s: %s" % (type(e).__name__, e))[:200]
            with JOBS_LOCK:
                item["status"] = "error"
                item["error"] = ("Lỗi không xác định khi chuyển đổi tệp này "
                                 "(%s). Xem log máy chủ." % detail)
        with JOBS_LOCK:
            job["done"] = i + 1
    with JOBS_LOCK:
        job["status"] = "done"
        job["current"] = ""
        job["finished"] = datetime.datetime.now()
    log.info("Job %s xong: %d tệp.", job["id"], job["total"])


# ----------------------------------------------------------------------------
# Routes
# ----------------------------------------------------------------------------
@app.route("/")
def index():
    return Response(PAGE, mimetype="text/html; charset=utf-8")


@app.route("/health")
def health():
    return jsonify(ok=True)


@app.route("/api/jobs", methods=["POST"])
def create_job():
    """Nhận nhiều tệp Excel (và/hoặc .zip), tạo job nền, trả job_id ngay."""
    cleanup()
    try:
        uploads = request.files.getlist("files[]") + request.files.getlist("files")
        files, skipped = _collect_uploads(uploads)
        if not files:
            raise ValueError("Vui lòng chọn ít nhất một tệp Excel (.xlsx) "
                             "— hoặc tệp .zip có chứa .xlsx.")
        if len(files) > MAX_FILES_PER_JOB:
            raise ValueError("Mỗi lượt tối đa %d tệp (đang chọn %d). "
                             "Hãy chia thành nhiều lượt." % (MAX_FILES_PER_JOB, len(files)))
        options = _parse_options(request.form)
    except ValueError as e:
        return jsonify(ok=False, error=str(e)), 400

    names = [n for n, _ in files]
    out_names = _unique_names([_output_name(n) for n in names])
    job_id = uuid.uuid4().hex
    job = {
        "id": job_id,
        "created": datetime.datetime.now(),
        "finished": None,
        "status": "running",
        "total": len(files),
        "done": 0,
        "current": "",
        "options": options,
        "items": [{"name": n, "out_name": o, "status": "pending",
                   "error": None, "size": None}
                  for n, o in zip(names, out_names)],
        "inputs": [b for _, b in files],
        "outputs": [None] * len(files),
    }
    with JOBS_LOCK:
        JOBS[job_id] = job
    threading.Thread(target=_process_job, args=(job,), daemon=True).start()
    log.info("Job %s: %d tệp, tùy chọn %s", job_id, len(files), options)
    return jsonify(ok=True, job_id=job_id, total=len(files),
                   skipped=skipped, items=[dict(it) for it in job["items"]])


@app.route("/api/jobs/<job_id>")
def job_status(job_id):
    job = _get_job(job_id)
    if not job:
        return _json_404("Không tìm thấy công việc (có thể đã quá 2 giờ và bị xóa).")
    return jsonify(_public_view(job))


@app.route("/api/jobs/<job_id>/files/<int:idx>")
def job_file(job_id, idx):
    job = _get_job(job_id)
    if not job:
        return _json_404("Không tìm thấy công việc (có thể đã quá 2 giờ và bị xóa).")
    with JOBS_LOCK:
        if idx < 0 or idx >= job["total"]:
            return _json_404("Không có tệp số %d trong công việc này." % idx)
        data = job["outputs"][idx]
        item = dict(job["items"][idx])
    if data is None:
        if item["status"] == "error":
            return _json_404("Tệp '%s' chuyển đổi bị lỗi: %s" % (item["name"], item["error"]))
        return _json_404("Tệp '%s' chưa được chuyển đổi xong." % item["name"])
    return _attachment(data, item["out_name"], DOCX_MIME)


@app.route("/api/jobs/<job_id>/zip")
def job_zip(job_id):
    job = _get_job(job_id)
    if not job:
        return _json_404("Không tìm thấy công việc (có thể đã quá 2 giờ và bị xóa).")
    with JOBS_LOCK:
        pairs = [(it["out_name"], job["outputs"][i])
                 for i, it in enumerate(job["items"]) if job["outputs"][i] is not None]
    if not pairs:
        return _json_404("Chưa có tệp kết quả nào để đóng gói.")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nm, data in pairs:
            z.writestr(nm, data)
    name = "muc_luc_ho_so_%s.zip" % datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return _attachment(buf.getvalue(), name, ZIP_MIME)


@app.errorhandler(RequestEntityTooLarge)
def _too_large(_e):
    return jsonify(ok=False, error="Dung lượng tải lên vượt quá giới hạn 500 MB. "
                                   "Hãy chia thành nhiều lượt nhỏ hơn."), 413


@app.errorhandler(NotFound)
def _not_found(e):
    if request.path.startswith("/api/"):
        return jsonify(ok=False, error="Không tìm thấy đường dẫn hoặc tài nguyên."), 404
    return e


# ----------------------------------------------------------------------------
# Giao diện (HTML/CSS/JS nhúng — không dùng tài nguyên ngoài, chạy được offline)
# ----------------------------------------------------------------------------
PAGE = r"""<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>In dự án Tam Đại — Mục lục hồ sơ Excel → Word</title>
<style>
  :root{
    --bg1:#0f172a; --bg2:#1e293b;
    --card:#ffffff; --ink:#0f172a; --muted:#64748b; --line:#e2e8f0;
    --brand:#0d9488; --brand2:#059669; --ok:#16a34a; --bad:#dc2626;
    --shadow:0 10px 40px rgba(2,6,23,.12);
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",
       Roboto,Helvetica,Arial,"Apple Color Emoji",sans-serif;color:var(--ink);
       background:radial-gradient(1200px 600px at 10% -10%,#134e4a 0%,transparent 60%),
                  radial-gradient(1000px 500px at 100% 0%,#047857 0%,transparent 55%),
                  linear-gradient(180deg,var(--bg1),var(--bg2));
       min-height:100vh;padding:32px 16px;}
  .wrap{max-width:1020px;margin:0 auto;}
  header{color:#fff;text-align:center;margin-bottom:26px;}
  header .badge{display:inline-flex;align-items:center;gap:8px;
       background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);
       padding:6px 14px;border-radius:999px;font-size:13px;backdrop-filter:blur(6px);}
  header h1{font-size:30px;margin:14px 0 6px;font-weight:800;letter-spacing:-.02em;}
  header p{margin:0 auto;color:#a7f3d0;font-size:15px;max-width:760px;line-height:1.5}
  .card{background:var(--card);border-radius:20px;box-shadow:var(--shadow);padding:26px;}

  .drop{border:2px dashed #cbd5e1;border-radius:16px;padding:24px;text-align:center;
        cursor:pointer;transition:.18s;background:#f8fafc;position:relative;display:block}
  .drop:hover{border-color:var(--brand);background:#f0fdfa;}
  .drop.drag{border-color:var(--brand);background:#ccfbf1;transform:translateY(-2px);}
  .drop .ico{width:48px;height:48px;border-radius:12px;display:grid;place-items:center;
        margin:0 auto 10px;font-size:24px;color:#fff;
        background:linear-gradient(135deg,var(--brand),var(--brand2));}
  .drop .t{font-weight:700;font-size:15px}
  .drop .s{color:var(--muted);font-size:13px;margin-top:4px}
  .drop .fname{margin-top:10px;font-size:13px;font-weight:700;color:var(--brand);}
  .drop.has{border-style:solid;border-color:#99f6e4;background:#fff;}
  .drop.off{pointer-events:none;opacity:.7}
  input[type=file]{display:none}

  .files{margin-top:14px;display:none;flex-direction:column;gap:8px;
         max-height:46vh;overflow:auto;padding-right:2px}
  .frow{display:flex;align-items:center;gap:10px;background:#f8fafc;
        border:1px solid var(--line);border-radius:12px;padding:9px 12px;font-size:14px}
  .frow .no{color:#94a3b8;font-size:12px;min-width:26px;text-align:right}
  .frow .meta{flex:1;min-width:0}
  .frow .nm{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:600}
  .frow .sub{font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .frow .er{font-size:12px;color:#b91c1c;margin-top:2px;white-space:normal;line-height:1.4}
  .frow .st{font-size:12px;font-weight:700;padding:3px 9px;border-radius:999px;white-space:nowrap}
  .frow .st.wait{background:#e2e8f0;color:#475569}
  .frow .st.run{background:#ccfbf1;color:#0f766e}
  .frow .st.ok{background:#dcfce7;color:#15803d}
  .frow .st.err{background:#fee2e2;color:#b91c1c}
  .frow .rm{cursor:pointer;color:#94a3b8;padding:0 4px;font-size:15px}
  .frow .rm:hover{color:var(--bad)}
  .frow a.dl{font-size:13px;font-weight:700;color:var(--brand2);text-decoration:none;
        border:1px solid #a7f3d0;background:#ecfdf5;padding:5px 10px;border-radius:9px;white-space:nowrap}
  .frow a.dl:hover{background:#d1fae5}
  .listhead{display:flex;justify-content:space-between;align-items:center;margin-top:16px;
        font-size:13px;color:var(--muted)}
  .listhead b{color:var(--ink)}

  .opts{display:grid;grid-template-columns:1.4fr 1fr 1fr;gap:14px;margin-top:18px}
  @media (max-width:820px){.opts{grid-template-columns:1fr}}
  .opt{background:#f8fafc;border:1px solid var(--line);border-radius:14px;padding:12px 14px}
  .opt .lb{font-size:12px;font-weight:700;color:#475569;text-transform:uppercase;
        letter-spacing:.03em;margin-bottom:8px}
  .radios{display:flex;flex-direction:column;gap:6px}
  .radios label{display:flex;gap:9px;align-items:flex-start;cursor:pointer;font-size:13px;
        padding:7px 9px;border-radius:10px;border:1px solid transparent;background:#fff;line-height:1.35}
  .radios label:hover{border-color:#99f6e4}
  .radios input{margin:2px 0 0;accent-color:var(--brand)}
  .radios label.on{border-color:var(--brand);background:#f0fdfa}
  .radios .d{color:var(--muted);font-size:12px;display:block}
  .seg{display:flex;flex-wrap:wrap;gap:6px}
  .seg label{position:relative;cursor:pointer}
  .seg input{position:absolute;opacity:0;pointer-events:none}
  .seg span{display:inline-block;padding:8px 13px;border-radius:9px;border:1px solid var(--line);
        background:#fff;font-size:13px;font-weight:600;color:#334155;transition:.15s;user-select:none}
  .seg input:checked+span{background:var(--brand);border-color:var(--brand);color:#fff}
  .seg input:focus-visible+span{box-shadow:0 0 0 3px rgba(13,148,136,.25)}
  .opts.off{opacity:.6;pointer-events:none}

  .toolbar{display:flex;flex-wrap:wrap;align-items:center;gap:12px;margin-top:18px}
  .toolbar .hint{margin:0;flex:1;min-width:200px}
  button{font:inherit;border:0;border-radius:12px;padding:13px 22px;font-weight:700;
         cursor:pointer;transition:.15s;display:inline-flex;align-items:center;gap:8px;}
  button:disabled{opacity:.55;cursor:not-allowed}
  .btn-primary{background:linear-gradient(135deg,var(--brand),var(--brand2));color:#fff;
         box-shadow:0 8px 20px rgba(5,150,105,.35);}
  .btn-primary:hover:not(:disabled){transform:translateY(-1px)}
  .btn-ghost{background:#f1f5f9;color:#334155}
  .btn-ghost:hover:not(:disabled){background:#e2e8f0}
  .hint{color:var(--muted);font-size:13px;margin-top:14px;line-height:1.6}
  .spin{width:16px;height:16px;border:2px solid rgba(255,255,255,.4);
        border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;display:none}
  @keyframes sp{to{transform:rotate(360deg)}}

  .progress{margin-top:20px;display:none}
  .progress .lab{display:flex;justify-content:space-between;gap:12px;font-size:13px;
        color:var(--muted);margin-bottom:6px}
  .progress .lab span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .progress .lab b{color:var(--ink);white-space:nowrap}
  .bar{height:14px;background:#e2e8f0;border-radius:999px;overflow:hidden}
  .bar .fill{height:100%;width:0;border-radius:999px;transition:width .3s ease;
        background:linear-gradient(90deg,var(--brand),var(--brand2));
        background-size:200% 100%;animation:flow 1.4s linear infinite}
  .bar .fill.up{background:linear-gradient(90deg,#38bdf8,#2563eb);background-size:200% 100%}
  .bar .fill.stop{animation:none}
  @keyframes flow{to{background-position:200% 0}}

  .result{margin-top:18px;background:linear-gradient(135deg,#ecfdf5,#f0fdfa);
        border:1px solid #bbf7d0;border-radius:16px;padding:18px;display:none;
        align-items:center;gap:16px;flex-wrap:wrap}
  .result.warn{background:linear-gradient(135deg,#fffbeb,#fef3c7);border-color:#fde68a}
  .result.fail{background:linear-gradient(135deg,#fef2f2,#fee2e2);border-color:#fecaca}
  .result .ic{width:44px;height:44px;border-radius:12px;background:var(--ok);color:#fff;
        display:grid;place-items:center;font-size:22px;flex:0 0 auto}
  .result.warn .ic{background:#d97706}
  .result.fail .ic{background:var(--bad)}
  .result .meta{flex:1;min-width:200px}
  .result .meta .a{font-weight:700;font-size:15px}
  .result .meta .b{color:var(--muted);font-size:13px;margin-top:2px}
  .result .acts{display:flex;gap:10px;flex-wrap:wrap}
  .result a{text-decoration:none}

  .toast{position:fixed;left:50%;bottom:26px;transform:translateX(-50%) translateY(20px);
        background:#0f172a;color:#fff;padding:12px 18px;border-radius:12px;
        box-shadow:var(--shadow);opacity:0;transition:.25s;pointer-events:none;z-index:9;
        font-size:14px;max-width:90vw}
  .toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
  .toast.bad{background:#7f1d1d}
  footer{color:#6ee7b7;text-align:center;font-size:13px;margin-top:24px;line-height:1.6}
  footer code{color:#d1fae5}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <span class="badge">🗂️ Lưu trữ • Dự án Tam Đại</span>
    <h1>In dự án Tam Đại — Mục lục hồ sơ Excel → Word</h1>
    <p>Chọn nhiều tệp Excel theo mẫu "Mục lục hồ sơ" — mỗi tệp cho ra một tệp Word khổ ngang
       gồm Tờ bìa, Bảng chữ viết tắt, Phụ lục tra cứu, Bảng thống kê hồ sơ và Chứng từ kết thúc.</p>
  </header>

  <div class="card">
    <label class="drop" id="drop">
      <div class="ico">📊</div>
      <div class="t">Chọn tệp Excel (.xlsx) — có thể chọn nhiều, hoặc một tệp .zip chứa các .xlsx</div>
      <div class="s">Kéo thả vào đây hoặc bấm để chọn từ máy</div>
      <div class="fname" id="dropName"></div>
      <input type="file" id="fileInput" accept=".xlsx,.xlsm,.zip" multiple>
    </label>

    <div class="listhead" id="listHead" style="display:none">
      <span id="listCount"></span><span id="listNote"></span>
    </div>
    <div class="files" id="list"></div>

    <div class="opts" id="opts">
      <div class="opt">
        <div class="lb">Đánh số trang</div>
        <div class="radios" id="pnGroup">
          <label class="on"><input type="radio" name="pn" value="2" checked>
            <span>Trang 2 mang số 2<span class="d">Bìa là trang 1, không in số</span></span></label>
          <label><input type="radio" name="pn" value="1">
            <span>Trang 2 mang số 1<span class="d">Không tính bìa</span></span></label>
        </div>
      </div>
      <div class="opt">
        <div class="lb">Khổ giấy</div>
        <div class="seg">
          <label><input type="radio" name="ps" value="A4" checked><span>A4 ngang</span></label>
          <label><input type="radio" name="ps" value="Letter"><span>Letter ngang</span></label>
        </div>
      </div>
      <div class="opt">
        <div class="lb">Cỡ chữ bảng thống kê</div>
        <div class="seg">
          <label><input type="radio" name="fs" value="8.5" checked><span>8.5</span></label>
          <label><input type="radio" name="fs" value="9"><span>9</span></label>
          <label><input type="radio" name="fs" value="10"><span>10</span></label>
          <label><input type="radio" name="fs" value="11"><span>11</span></label>
        </div>
      </div>
    </div>

    <div class="toolbar" id="toolbar">
      <div class="hint">Tệp Word giữ nguyên tên tệp Excel (đuôi .docx). Trang bìa không in số trang;
        số trang in ở giữa đầu trang từ trang 2. Sheet "Biên mục" (mục lục văn bản) được bỏ qua.</div>
      <button class="btn-ghost" id="btnClear">🗑️ Xóa danh sách</button>
      <button class="btn-primary" id="btnRun" disabled>
        <span class="spin" id="spin"></span>⚙️ Chuyển đổi
      </button>
    </div>

    <div class="progress" id="progress">
      <div class="lab"><span id="progText">Đang chuẩn bị…</span><b id="progPct">0%</b></div>
      <div class="bar"><div class="fill" id="progFill"></div></div>
    </div>

    <div class="result" id="result">
      <div class="ic" id="resIcon">✓</div>
      <div class="meta">
        <div class="a" id="resTitle"></div>
        <div class="b" id="resInfo"></div>
      </div>
      <div class="acts">
        <a id="zipLink" href="#" style="display:none"><button class="btn-primary">⬇️ Tải tất cả (.zip)</button></a>
        <button class="btn-ghost" id="btnReset">🔄 Làm lại</button>
      </div>
    </div>
  </div>

  <footer>Khổ ngang • Times New Roman • mỗi tệp Excel → một tệp Word • tối đa 500 MB / lượt •
    kết quả lưu tạm trên máy chủ 2 giờ</footer>
</div>

<div class="toast" id="toast"></div>

<script>
const $=s=>document.querySelector(s);
// Tiền tố đường dẫn: chạy độc lập -> "", khi gộp mount tại /tamdai -> "/tamdai".
const BASE=location.pathname.replace(/\/+$/,"");
const POLL_MS=700;
const fileInput=$("#fileInput"), drop=$("#drop"), listEl=$("#list");
let selected=[];            // File[] đã chọn, chưa gửi
let busy=false;             // đang tải lên / đang xử lý
let jobId=null, pollTimer=null, rowsBuilt=false;
let toastTimer;

function toast(msg,bad){
  const t=$("#toast"); t.textContent=msg; t.className="toast show"+(bad?" bad":"");
  clearTimeout(toastTimer); toastTimer=setTimeout(()=>t.className="toast",bad?5000:3500);
}
function esc(s){return (s??"").toString().replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));}
function fmtSize(b){
  if(b==null) return "";
  if(b<1024) return b+" B";
  if(b<1048576) return (b/1024).toFixed(0)+" KB";
  return (b/1048576).toFixed(1)+" MB";
}
const isXlsx=n=>/\.(xlsx|xlsm)$/i.test(n), isZip=n=>/\.zip$/i.test(n);
const val=n=>{const el=document.querySelector('input[name="'+n+'"]:checked');return el?el.value:"";};

/* ---------- chọn tệp ---------- */
function addFiles(list){
  const skipped=[];
  for(const f of list){
    if(!(isXlsx(f.name)||isZip(f.name))){ skipped.push(f.name); continue; }
    if(selected.some(x=>x.name===f.name&&x.size===f.size&&x.lastModified===f.lastModified)) continue;
    selected.push(f);
  }
  if(skipped.length) toast("Bỏ qua "+skipped.length+" tệp không phải .xlsx/.zip: "
      +skipped.slice(0,3).join(", ")+(skipped.length>3?"…":""),true);
  renderSelected();
}
function renderSelected(){
  rowsBuilt=false;
  const n=selected.length, head=$("#listHead");
  $("#btnRun").disabled=!n||busy;
  if(!n){
    listEl.style.display="none"; listEl.innerHTML=""; head.style.display="none";
    drop.classList.remove("has"); $("#dropName").textContent=""; return;
  }
  const bytes=selected.reduce((s,f)=>s+f.size,0), zips=selected.filter(f=>isZip(f.name)).length;
  drop.classList.add("has");
  $("#dropName").textContent="Đã chọn "+n+" tệp ("+fmtSize(bytes)+")";
  head.style.display="flex";
  $("#listCount").innerHTML="<b>"+n+"</b> tệp đã chọn"+(zips?" · "+zips+" tệp .zip sẽ được giải nén trên máy chủ":"");
  $("#listNote").textContent="";
  listEl.style.display="flex";
  listEl.innerHTML=selected.map((f,i)=>`<div class="frow">
      <span class="no">${i+1}</span><span>${isZip(f.name)?"🗜️":"📄"}</span>
      <div class="meta"><div class="nm" title="${esc(f.name)}">${esc(f.name)}</div>
        <div class="sub">${fmtSize(f.size)}</div></div>
      <span class="st wait">chờ</span>
      <span class="rm" data-rm="${i}" title="Bỏ tệp này">✕</span></div>`).join("");
}
listEl.addEventListener("click",e=>{
  const rm=e.target.getAttribute("data-rm");
  if(rm!==null&&!busy){ selected.splice(+rm,1); renderSelected(); }
});
fileInput.addEventListener("change",()=>{ addFiles(fileInput.files); fileInput.value=""; });
["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();if(!busy)drop.classList.add("drag");}));
["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("drag");}));
drop.addEventListener("drop",e=>{ if(!busy&&e.dataTransfer.files.length) addFiles(e.dataTransfer.files); });
// tránh trình duyệt mở tệp khi thả ra ngoài vùng kéo thả
["dragover","drop"].forEach(ev=>document.addEventListener(ev,e=>e.preventDefault()));
$("#btnClear").addEventListener("click",()=>{ if(busy) return; selected=[]; renderSelected(); });
$("#pnGroup").addEventListener("change",()=>{
  $("#pnGroup").querySelectorAll("label").forEach(l=>l.classList.toggle("on",l.querySelector("input").checked));
});

/* ---------- tiến độ ---------- */
function showProgress(p,text,upload){
  $("#progress").style.display="block";
  const f=$("#progFill"); f.style.width=Math.max(2,Math.min(100,p))+"%";
  f.classList.toggle("up",!!upload);
  $("#progPct").textContent=Math.round(p)+"%";
  if(text!==undefined) $("#progText").textContent=text;
}
function setBusy(b){
  busy=b;
  $("#btnRun").disabled=b||!selected.length;
  $("#spin").style.display=b?"inline-block":"none";
  $("#btnClear").disabled=b;
  $("#opts").classList.toggle("off",b);
  drop.classList.toggle("off",b);
  fileInput.disabled=b;
}

/* ---------- gửi lên & theo dõi ---------- */
function upload(fd){
  return new Promise((resolve,reject)=>{
    const xhr=new XMLHttpRequest();
    xhr.open("POST",BASE+"/api/jobs");
    xhr.upload.onprogress=e=>{
      if(e.lengthComputable){ const p=Math.round(e.loaded/e.total*100);
        showProgress(p,p<100?"Đang tải lên máy chủ… "+fmtSize(e.loaded)+" / "+fmtSize(e.total):"Đã tải lên, máy chủ đang chuẩn bị…",true); }
    };
    xhr.onload=()=>{
      try{ resolve(JSON.parse(xhr.responseText)); }
      catch(_){ reject(new Error("máy chủ trả về lỗi HTTP "+xhr.status)); }
    };
    xhr.onerror=()=>reject(new Error("mất kết nối (tệp quá lớn hoặc máy chủ không phản hồi)"));
    xhr.send(fd);
  });
}
async function startJob(){
  if(busy||!selected.length) return;
  setBusy(true);
  $("#result").style.display="none";
  listEl.querySelectorAll(".rm").forEach(el=>el.remove());
  const fd=new FormData();
  selected.forEach(f=>fd.append("files[]",f,f.name));
  fd.append("page_number_start",val("pn"));
  fd.append("page_size",val("ps"));
  fd.append("table_font_pt",val("fs"));
  showProgress(0,"Đang tải lên máy chủ…",true);
  let d;
  try{ d=await upload(fd); }
  catch(e){ toast("Không gửi được tệp lên máy chủ: "+e.message,true); setBusy(false); $("#progress").style.display="none"; return; }
  if(!d.ok){ toast(d.error||"Máy chủ từ chối yêu cầu.",true); setBusy(false); $("#progress").style.display="none"; return; }
  jobId=d.job_id; rowsBuilt=false;
  if(d.skipped&&d.skipped.length) toast("Bỏ qua: "+d.skipped.join(", "),true);
  showProgress(0,"Đang bắt đầu chuyển đổi "+d.total+" tệp…");
  poll();
}
async function poll(){
  let d;
  try{
    const r=await fetch(BASE+"/api/jobs/"+jobId,{cache:"no-store"});
    if(r.status===404){ toast("Công việc không còn tồn tại trên máy chủ.",true); finish(null); return; }
    d=await r.json();
  }catch(_){ pollTimer=setTimeout(poll,POLL_MS*2); return; }
  renderJob(d);
  if(d.status==="done"){ finish(d); return; }
  pollTimer=setTimeout(poll,POLL_MS);
}
function renderJob(d){
  const running=d.status!=="done";
  if(running) showProgress(d.percent,"Đang xử lý tệp "+Math.min(d.done+1,d.total)+"/"+d.total+(d.current?": "+d.current:""));
  else showProgress(100,"Hoàn tất "+d.ok_count+"/"+d.total+" tệp");
  if(!rowsBuilt){
    listEl.style.display="flex";
    listEl.innerHTML=d.items.map((it,i)=>`<div class="frow" id="row${i}">
        <span class="no">${i+1}</span><span>📄</span>
        <div class="meta"><div class="nm" title="${esc(it.name)}">${esc(it.name)}</div>
          <div class="sub" id="sub${i}"></div><div class="er" id="er${i}"></div></div>
        <span class="st wait" id="st${i}">chờ</span>
        <a class="dl" id="dl${i}" style="display:none" download>⬇️ Tải về</a></div>`).join("");
    $("#listHead").style.display="flex";
    $("#listCount").innerHTML="<b>"+d.total+"</b> tệp Excel";
    rowsBuilt=true;
  }
  let runRow=null;
  d.items.forEach((it,i)=>{
    const st=$("#st"+i), sub=$("#sub"+i), er=$("#er"+i), dl=$("#dl"+i);
    if(it.status==="ok"){
      st.className="st ok"; st.textContent="✓ xong";
      sub.textContent="→ "+it.out_name+" · "+fmtSize(it.size); er.textContent="";
      dl.href=BASE+"/api/jobs/"+d.job_id+"/files/"+i; dl.style.display="inline-block";
    }else if(it.status==="error"){
      st.className="st err"; st.textContent="lỗi"; sub.textContent="→ "+it.out_name;
      er.textContent=it.error||"Không chuyển đổi được."; dl.style.display="none";
    }else if(it.status==="running"){
      st.className="st run"; st.textContent="đang xử lý…"; sub.textContent="→ "+it.out_name; runRow=i;
    }else{ st.className="st wait"; st.textContent="chờ"; sub.textContent="→ "+it.out_name; }
  });
  $("#listNote").textContent=(d.ok_count?d.ok_count+" xong":"")+(d.error_count?(d.ok_count?" · ":"")+d.error_count+" lỗi":"");
  if(runRow!==null&&running){ const row=$("#row"+runRow); if(row) row.scrollIntoView({block:"nearest"}); }
}
function finish(d){
  clearTimeout(pollTimer); pollTimer=null;
  setBusy(false);
  $("#btnRun").disabled=true;          // đợt này đã xong; bấm "Làm lại" để chọn tệp mới
  $("#toolbar").style.display="none";
  $("#progFill").classList.add("stop");
  const box=$("#result"), zip=$("#zipLink");
  box.className="result";
  if(!d){
    box.classList.add("fail"); $("#resIcon").textContent="!";
    $("#resTitle").textContent="Không theo dõi được công việc";
    $("#resInfo").textContent="Hãy bấm Làm lại và thử lại."; zip.style.display="none";
  }else if(d.ok_count===0){
    box.classList.add("fail"); $("#resIcon").textContent="✕";
    $("#resTitle").textContent="Không chuyển đổi được tệp nào";
    $("#resInfo").textContent=d.error_count+" tệp lỗi — xem thông báo ở từng dòng bên trên.";
    zip.style.display="none"; toast("Không tạo được tệp Word nào",true);
  }else{
    if(d.error_count){ box.classList.add("warn"); $("#resIcon").textContent="!"; }
    else $("#resIcon").textContent="✓";
    $("#resTitle").textContent="Đã tạo "+d.ok_count+" tệp Word"+(d.error_count?" · "+d.error_count+" tệp lỗi":"");
    $("#resInfo").textContent="Tải từng tệp ở mỗi dòng hoặc tải tất cả trong một tệp .zip. Kết quả lưu tạm 2 giờ.";
    zip.href=BASE+"/api/jobs/"+d.job_id+"/zip"; zip.style.display="inline-block";
    toast("Đã tạo xong "+d.ok_count+" tệp Word ✔");
  }
  box.style.display="flex";
  box.scrollIntoView({behavior:"smooth",block:"nearest"});
}
function resetAll(){
  clearTimeout(pollTimer); pollTimer=null; jobId=null; rowsBuilt=false;
  selected=[]; fileInput.value="";
  setBusy(false);
  $("#toolbar").style.display="flex";
  $("#progress").style.display="none"; $("#progFill").classList.remove("stop"); $("#progFill").style.width="0";
  $("#result").style.display="none";
  renderSelected();
  window.scrollTo({top:0,behavior:"smooth"});
}
$("#btnRun").addEventListener("click",startJob);
$("#btnReset").addEventListener("click",resetAll);
window.addEventListener("beforeunload",e=>{ if(busy){ e.preventDefault(); e.returnValue=""; } });
</script>
</body>
</html>
"""

def _lan_ip():
    """Địa chỉ IP trong mạng LAN (không tra DNS nên không bị treo khi khởi động)."""
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))      # không gửi gói tin nào, chỉ chọn giao diện mạng
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()


if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5039"))
    print("Truy cập:  http://127.0.0.1:%d   |   http://%s:%d" % (port, _lan_ip(), port))
    app.run(host=host, port=port, debug=False, threaded=True)
