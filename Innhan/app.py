# -*- coding: utf-8 -*-
"""
Web app (Flask) IN NHÃN HỘP lưu trữ (.docx khổ A4 dọc) từ Excel.

- Chọn một hoặc nhiều tệp Excel (.xlsx). MỖI TỆP Excel -> MỘT tệp Word
  (gộp mọi sheet, 4 nhãn/trang). Nhiều tệp -> đóng gói .zip.

Chạy (không cần venv):
    python3 app.py
Mở:  http://127.0.0.1:5050
"""

import io
import os
import sys
import uuid
import zipfile
import datetime
import importlib.util
import subprocess
import threading


def ensure_deps():
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

from flask import (Flask, request, jsonify, send_file,  # noqa: E402
                   render_template_string, abort)

import in_nhan  # noqa: E402  (bản hợp nhất: khuôn sample.docx + co bề ngang nội dung)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

JOBS = {}
JOBS_LOCK = threading.Lock()


def cleanup():
    now = datetime.datetime.now()
    with JOBS_LOCK:
        for k in list(JOBS):
            if (now - JOBS[k]["created"]).total_seconds() > 3600:
                JOBS.pop(k, None)


@app.route("/")
def index():
    return render_template_string(PAGE)


def _process_job(job_id, files):
    job = JOBS[job_id]
    total = len(files)
    outputs = []

    for i, (name, blob) in enumerate(files):
        with JOBS_LOCK:
            job["file_index"] = i + 1
            job["current"] = name
            job["current_done"] = 0
            job["current_total"] = 0

        def cb(done, tot, _i=i):
            with JOBS_LOCK:
                job["current_done"] = done
                job["current_total"] = tot
                frac = (done / tot) if tot else 0
                job["percent"] = round((_i + frac) / total * 100, 1)

        try:
            out_bytes, n = in_nhan.generate_from_bytes(blob, progress=cb)
            base = os.path.splitext(os.path.basename(name))[0]
            docx_name = "%s - Nhãn.docx" % base
            outputs.append((docx_name, out_bytes))
            with JOBS_LOCK:
                job["results"].append({"name": name, "ok": True,
                                       "nhan": n, "out": docx_name})
        except Exception as e:
            with JOBS_LOCK:
                job["results"].append({"name": name, "ok": False, "error": str(e)})
        with JOBS_LOCK:
            job["percent"] = round((i + 1) / total * 100, 1)

    with JOBS_LOCK:
        if len(outputs) == 1:
            job["download_name"] = outputs[0][0]
            job["download_bytes"] = outputs[0][1]
            job["download_kind"] = "docx"
        elif len(outputs) > 1:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                seen = {}
                for nm, data in outputs:
                    if nm in seen:
                        seen[nm] += 1
                        root, ext = os.path.splitext(nm)
                        nm = "%s (%d)%s" % (root, seen[nm], ext)
                    else:
                        seen[nm] = 0
                    z.writestr(nm, data)
            job["download_name"] = "Nhan_hop_%d_tep.zip" % len(outputs)
            job["download_bytes"] = buf.getvalue()
            job["download_kind"] = "zip"
        job["status"] = "done"
        job["percent"] = 100.0
        job["ok_count"] = len(outputs)


@app.route("/jobs", methods=["POST"])
def create_job():
    cleanup()
    try:
        uploads = request.files.getlist("data")
        files = []
        for f in uploads:
            if not (f and f.filename):
                continue
            if not f.filename.lower().endswith((".xlsx", ".xlsm")):
                raise ValueError("Tệp '%s' không phải .xlsx" % f.filename)
            files.append((f.filename, f.read()))
        if not files:
            raise ValueError("Vui lòng chọn ít nhất một tệp Excel (.xlsx).")
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 400

    job_id = uuid.uuid4().hex
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "running", "total_files": len(files), "file_index": 0,
            "current": "", "current_done": 0, "current_total": 0,
            "percent": 0.0, "results": [], "ok_count": 0,
            "download_name": None, "download_bytes": None, "download_kind": None,
            "created": datetime.datetime.now(),
        }
    t = threading.Thread(target=_process_job, args=(job_id, files), daemon=True)
    t.start()
    return jsonify(ok=True, job_id=job_id, total_files=len(files))


@app.route("/jobs/<job_id>")
def job_status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify(ok=False, error="Không tìm thấy công việc."), 404
    with JOBS_LOCK:
        return jsonify(
            ok=True, status=job["status"], percent=job["percent"],
            total_files=job["total_files"], file_index=job["file_index"],
            current=job["current"], current_done=job["current_done"],
            current_total=job["current_total"], results=job["results"],
            ok_count=job["ok_count"],
            has_download=job["download_bytes"] is not None,
            download_name=job["download_name"], download_kind=job["download_kind"],
        )


@app.route("/jobs/<job_id>/download")
def job_download(job_id):
    job = JOBS.get(job_id)
    if not job or job["download_bytes"] is None:
        abort(404)
    kind = job["download_kind"]
    mime = ("application/zip" if kind == "zip" else
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document")
    return send_file(io.BytesIO(job["download_bytes"]), mimetype=mime,
                     as_attachment=True, download_name=job["download_name"])


# --------------------------- Giao diện (HTML) ---------------------------------
PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>In nhãn hộp lưu trữ</title>
<style>
  :root{
    --bg1:#0f172a; --bg2:#1e293b; --card:#fff; --ink:#0f172a; --muted:#64748b;
    --line:#e2e8f0; --brand:#059669; --brand2:#0d9488; --ok:#16a34a; --bad:#dc2626;
    --shadow:0 10px 40px rgba(2,6,23,.12);
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,
       Helvetica,Arial,sans-serif;color:var(--ink);
       background:radial-gradient(1200px 600px at 10% -10%,#065f46 0%,transparent 60%),
                  radial-gradient(1000px 500px at 100% 0%,#0f766e 0%,transparent 55%),
                  linear-gradient(180deg,var(--bg1),var(--bg2));
       min-height:100vh;padding:32px 16px;}
  .wrap{max-width:900px;margin:0 auto;}
  header{color:#fff;text-align:center;margin-bottom:26px;}
  header .badge{display:inline-flex;align-items:center;gap:8px;
       background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);
       padding:6px 14px;border-radius:999px;font-size:13px;}
  header h1{font-size:30px;margin:14px 0 6px;font-weight:800;}
  header p{margin:0;color:#a7f3d0;font-size:15px;}
  .card{background:var(--card);border-radius:20px;box-shadow:var(--shadow);padding:26px;}
  .drop{border:2px dashed #cbd5e1;border-radius:16px;padding:22px;text-align:center;
        cursor:pointer;transition:.18s;background:#f8fafc;}
  .drop:hover{border-color:var(--brand);background:#f0fdf4;}
  .drop.drag{border-color:var(--brand);background:#dcfce7;transform:translateY(-2px);}
  .drop .ico{width:46px;height:46px;border-radius:12px;display:grid;place-items:center;
        margin:0 auto 10px;font-size:22px;color:#fff;
        background:linear-gradient(135deg,var(--brand),var(--brand2));}
  .drop .t{font-weight:700;font-size:15px}
  .drop .s{color:var(--muted);font-size:13px;margin-top:4px}
  .drop .fname{margin-top:10px;font-size:13px;font-weight:600;color:var(--brand);word-break:break-all;}
  .drop.has{border-style:solid;border-color:#a7f3d0;background:#fff;}
  input[type=file]{display:none}
  button{font:inherit;border:0;border-radius:12px;padding:13px 22px;font-weight:700;
         cursor:pointer;transition:.15s;display:inline-flex;align-items:center;gap:8px;}
  button:disabled{opacity:.55;cursor:not-allowed}
  .btn-primary{background:linear-gradient(135deg,var(--brand),var(--brand2));color:#fff;
         box-shadow:0 8px 20px rgba(5,150,105,.35);}
  .btn-primary:hover:not(:disabled){transform:translateY(-1px)}
  .btn-ghost{background:#f1f5f9;color:#334155}
  .hint{color:var(--muted);font-size:13px;margin-top:14px;line-height:1.6}
  .files{margin-top:14px;display:none;flex-direction:column;gap:8px}
  .frow{display:flex;align-items:center;gap:10px;background:#f8fafc;
        border:1px solid var(--line);border-radius:12px;padding:9px 12px;font-size:14px}
  .frow .nm{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:600}
  .frow .st{font-size:12px;font-weight:700;padding:2px 9px;border-radius:999px;white-space:nowrap}
  .frow .st.wait{background:#e2e8f0;color:#475569}
  .frow .st.run{background:#d1fae5;color:#047857}
  .frow .st.ok{background:#dcfce7;color:#15803d}
  .frow .st.err{background:#fee2e2;color:#b91c1c}
  .toolbar{display:flex;gap:12px;margin-top:18px;flex-wrap:wrap}
  .toolbar #btnGen{margin-left:auto}
  .progress{margin-top:20px;display:none}
  .progress .lab{display:flex;justify-content:space-between;font-size:13px;color:var(--muted);margin-bottom:6px}
  .progress .lab b{color:var(--ink)}
  .bar{height:14px;background:#e2e8f0;border-radius:999px;overflow:hidden}
  .bar .fill{height:100%;width:0;border-radius:999px;transition:width .25s ease;
        background:linear-gradient(90deg,var(--brand),var(--brand2));}
  .spin{width:16px;height:16px;border:2px solid rgba(255,255,255,.4);
        border-top-color:#fff;border-radius:50%;animation:sp .7s linear infinite;display:none}
  @keyframes sp{to{transform:rotate(360deg)}}
  .dlcard{margin-top:18px;background:linear-gradient(135deg,#ecfdf5,#f0fdfa);
        border:1px solid #bbf7d0;border-radius:16px;padding:18px;display:none;
        align-items:center;gap:16px;flex-wrap:wrap}
  .dlcard .ic{width:44px;height:44px;border-radius:12px;background:var(--ok);color:#fff;
        display:grid;place-items:center;font-size:22px}
  .dlcard .meta{flex:1;min-width:160px}
  .dlcard .meta .a{font-weight:700}
  .dlcard .meta .b{color:var(--muted);font-size:13px}
  .toast{position:fixed;left:50%;bottom:26px;transform:translateX(-50%) translateY(20px);
        background:#0f172a;color:#fff;padding:12px 18px;border-radius:12px;
        box-shadow:var(--shadow);opacity:0;transition:.25s;pointer-events:none;z-index:9;font-size:14px;}
  .toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
  .toast.bad{background:#7f1d1d}
  footer{color:#6ee7b7;text-align:center;font-size:13px;margin-top:24px}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <span class="badge">🏷️ Lưu trữ • In nhãn hộp</span>
    <h1>In nhãn hộp lưu trữ</h1>
    <p>Chọn tệp Excel dữ liệu nhãn (.xlsx) — xuất Word khổ A4 dọc, 4 nhãn/trang.</p>
  </header>

  <div class="card">
    <label class="drop" id="dropData">
      <div class="ico">🏷️</div>
      <div class="t">Chọn tệp Excel (.xlsx) — có thể chọn nhiều</div>
      <div class="s">Kéo thả vào đây hoặc bấm để chọn từ máy</div>
      <div class="fname" id="nameData"></div>
      <input type="file" id="fileData" accept=".xlsx,.xlsm" multiple>
    </label>

    <div class="files" id="files"></div>

    <div class="toolbar">
      <button class="btn-primary" id="btnGen"><span class="spin" id="spinGen"></span>🏷️ Tạo nhãn &amp; tải về</button>
      <button class="btn-ghost" id="btnClear">🗑️ Xóa danh sách</button>
    </div>

    <div class="hint">
      Cấu trúc Excel mỗi sheet (dòng 1 là tiêu đề): <b>A</b>=Hộp số · <b>B</b>=Từ hồ sơ số ·
      <b>C</b>=Đến hồ sơ số · <b>D</b>=Nội dung (mã mục/nhóm, mỗi dòng một mục) ·
      <b>E</b>=Tiêu đề (Phông lưu trữ…). Mỗi dòng dữ liệu = một nhãn.
    </div>

    <div class="progress" id="progress">
      <div class="lab"><span id="progText">Đang chuẩn bị…</span><b id="progPct">0%</b></div>
      <div class="bar"><div class="fill" id="progFill"></div></div>
    </div>

    <div class="dlcard" id="dlcard">
      <div class="ic">✓</div>
      <div class="meta"><div class="a" id="dlName"></div><div class="b" id="dlInfo"></div></div>
      <a id="dlLink"><button class="btn-primary">⬇️ Tải kết quả</button></a>
    </div>
  </div>

  <footer>A4 dọc • 4 nhãn/trang • mỗi tệp Excel một tệp Word</footer>
</div>

<div class="toast" id="toast"></div>

<script>
const $=s=>document.querySelector(s);
// Tiền tố đường dẫn: chạy độc lập -> "", khi gộp mount tại /nhan -> "/nhan".
const BASE=location.pathname.replace(/\/+$/,"");
const fileInput=$("#fileData"), drop=$("#dropData");
let selected=[], statuses=[], busy=false, toastTimer;
function toast(msg,bad){const t=$("#toast");t.textContent=msg;t.className="toast show"+(bad?" bad":"");
  clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.className="toast",3500);}
function esc(s){return (s??"").toString().replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));}
function addFiles(list){for(const f of list){
    if(!/\.xlsx?$/i.test(f.name)){toast("Bỏ qua (không phải .xlsx): "+f.name,true);continue;}
    if(!selected.some(x=>x.name===f.name&&x.size===f.size)){selected.push(f);statuses.push({cls:"wait",txt:"chờ"});}}
  renderFiles();}
function renderFiles(){const box=$("#files");
  if(!selected.length){box.style.display="none";box.innerHTML="";drop.classList.remove("has");$("#nameData").textContent="";return;}
  drop.classList.add("has");$("#nameData").textContent="Đã chọn "+selected.length+" tệp";box.style.display="flex";
  box.innerHTML=selected.map((f,i)=>{const s=statuses[i]||{cls:"wait",txt:"chờ"};
    return `<div class="frow"><span>📄</span><span class="nm">${esc(f.name)}</span>
      <span class="st ${s.cls}" id="st${i}">${esc(s.txt)}</span>
      ${busy?"":`<span style="cursor:pointer;color:#94a3b8" data-rm="${i}">✕</span>`}</div>`;}).join("");}
$("#files").addEventListener("click",e=>{const rm=e.target.getAttribute("data-rm");
  if(rm!==null&&!busy){selected.splice(+rm,1);statuses.splice(+rm,1);renderFiles();}});
fileInput.addEventListener("change",()=>{addFiles(fileInput.files);fileInput.value="";});
["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("drag");}));
["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("drag");}));
drop.addEventListener("drop",e=>{if(e.dataTransfer.files.length)addFiles(e.dataTransfer.files);});
$("#btnClear").addEventListener("click",()=>{if(busy)return;selected=[];statuses=[];renderFiles();
  $("#progress").style.display="none";$("#dlcard").style.display="none";});
function setStatus(i,cls,txt){statuses[i]={cls,txt};const el=$("#st"+i);if(el){el.className="st "+cls;el.textContent=txt;}}
function setProgress(p,text){$("#progress").style.display="block";$("#progFill").style.width=Math.max(2,p)+"%";
  $("#progPct").textContent=Math.round(p)+"%";if(text!==undefined)$("#progText").textContent=text;}
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
async function poll(jobId){while(true){let d;
  try{d=await (await fetch(BASE+"/jobs/"+jobId)).json();}catch(_){await sleep(600);continue;}
  if(!d.ok){toast(d.error||"Mất công việc",true);return false;}
  setProgress(d.percent);const done=d.results.length;
  selected.forEach((f,i)=>{if(i<done){const r=d.results[i];
      if(r.ok)setStatus(i,"ok",r.nhan+" nhãn ✓");else setStatus(i,"err","lỗi");}
    else if(i===done&&d.status==="running"){setStatus(i,"run","đang xử lý…");}else setStatus(i,"wait","chờ");});
  if(d.status==="running"){const ct=d.current_total?(" — "+d.current_done+"/"+d.current_total+" nhãn"):"";
    setProgress(d.percent,"Đang xử lý tệp "+d.file_index+"/"+d.total_files+": "+(d.current||"")+ct);}
  if(d.status==="done"){setProgress(100,"Hoàn tất "+d.ok_count+" tệp Word");
    const errs=d.results.filter(r=>!r.ok);
    if(d.has_download){const hs=d.results.filter(r=>r.ok).reduce((s,r)=>s+(r.nhan||0),0);
      $("#dlName").textContent=d.download_name;
      $("#dlInfo").textContent=d.ok_count+" tệp Word • "+hs+" nhãn"+(d.download_kind==="zip"?" • đóng gói ZIP":"")
        +(errs.length?(" • "+errs.length+" tệp lỗi"):"");
      $("#dlLink").href=BASE+"/jobs/"+jobId+"/download";$("#dlcard").style.display="flex";
      toast("Đã tạo xong "+hs+" nhãn ✔");}else{toast("Không tạo được tệp nào",true);}
    if(errs.length){toast("Tệp lỗi: "+errs.map(r=>r.name).join(", "),true);}
    return true;}
  await sleep(400);}}
$("#btnGen").addEventListener("click",async()=>{if(busy)return;
  if(!selected.length){toast("Vui lòng chọn ít nhất một tệp Excel",true);return;}
  busy=true;statuses=selected.map(()=>({cls:"wait",txt:"chờ"}));
  const btn=$("#btnGen"),sp=$("#spinGen");btn.disabled=true;sp.style.display="inline-block";$("#dlcard").style.display="none";
  renderFiles();setProgress(0,"Đang tải lên & chuẩn bị…");
  try{const fd=new FormData();selected.forEach(f=>fd.append("data",f));
    const res=await fetch(BASE+"/jobs",{method:"POST",body:fd});const d=await res.json();
    if(!d.ok){toast(d.error,true);}else{const finished=await poll(d.job_id);
      if(finished){selected=[];statuses=[];fileInput.value="";}}}
  catch(e){toast("Lỗi: "+e,true);}
  finally{busy=false;btn.disabled=false;sp.style.display="none";renderFiles();}});
</script>
</body>
</html>
"""

if __name__ == "__main__":
    import socket
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        ip = "127.0.0.1"
    print("Truy cập:  http://127.0.0.1:5051   |   http://%s:5051" % ip)
    app.run(host="0.0.0.0", port=5051, debug=False, threaded=True)
