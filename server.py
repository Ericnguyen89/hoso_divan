# -*- coding: utf-8 -*-
"""
Máy chủ HỢP NHẤT: gộp 4 chương trình Flask vào MỘT trang duy nhất, hiển thị
bằng 4 tab. Bấm vào từng tab để chuyển giữa các chương trình.

    /            -> trang vỏ (shell) có 4 tab
    /bia/...     -> 📚 Tạo bìa hồ sơ            (app.py ở thư mục gốc)
    /mucluc/...  -> 📑 Mục lục văn bản, tài liệu (Phanmem_muclucdang/app.py)
    /nhan/...    -> 🏷️ In nhãn hộp lưu trữ       (Innhan/app.py)
    /hsxa/...    -> 📂 Tạo bìa hồ sơ xã          (HSxa/app.py)

Mỗi chương trình con vẫn giữ NGUYÊN toàn bộ logic và vẫn chạy độc lập được
như cũ (python3 app.py trong từng thư mục).

Chạy:
    python3 server.py
Mở:  http://127.0.0.1:8000

Triển khai (gunicorn):
    gunicorn --workers 1 --threads 8 --bind 0.0.0.0:8000 server:application
"""

import os
import sys
import importlib.util

from flask import Flask, render_template_string
from werkzeug.middleware.dispatcher import DispatcherMiddleware

ROOT = os.path.dirname(os.path.abspath(__file__))
MUCLUC_DIR = os.path.join(ROOT, "Phanmem_muclucdang")
NHAN_DIR = os.path.join(ROOT, "Innhan")
HSXA_DIR = os.path.join(ROOT, "HSxa")

# Đưa từng thư mục vào sys.path để mỗi app con nạp được module riêng của nó
# (tao_bia / muc_luc / in_nhan) và tìm đúng tệp mẫu .docx cạnh nó.
for d in (ROOT, MUCLUC_DIR, NHAN_DIR, HSXA_DIR):
    if d not in sys.path:
        sys.path.insert(0, d)


def _load_flask_app(path, module_name, shadow=None):
    """Nạp một tệp app.py và trả về đối tượng Flask `app` bên trong.

    Dùng tên module riêng biệt để các tệp app.py (trùng tên) không đè lên nhau.
    Nhờ import theo đường dẫn thật, biến __file__ của mỗi app trỏ đúng thư mục
    của nó, nên phần dò tệp mẫu .docx (theo __file__) vẫn chính xác.

    shadow: dict {tên_module: đường_dẫn.py} cho các module con mà app.py tự
    "import tên_module" (vd bia_app và hsxa_app đều "import tao_bia" nhưng
    mỗi thư mục có một bản tao_bia.py khác nhau). Nếu không xử lý, bản nạp
    trước sẽ bị cache trong sys.modules và bản sau "import" nhầm phải nó.
    Ở đây ta nạp đúng bản theo đường dẫn, đặt tạm vào sys.modules trong lúc
    exec_module rồi khôi phục lại như cũ ngay sau đó.
    """
    saved = {}
    if shadow:
        for name, subpath in shadow.items():
            saved[name] = sys.modules.get(name)
            sub_spec = importlib.util.spec_from_file_location(
                module_name + "__" + name, subpath)
            sub_module = importlib.util.module_from_spec(sub_spec)
            sub_spec.loader.exec_module(sub_module)
            sys.modules[name] = sub_module

    try:
        spec = importlib.util.spec_from_file_location(module_name, path)
        module = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = module
        spec.loader.exec_module(module)
        return module.app
    finally:
        for name, prev in saved.items():
            if prev is None:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = prev


bia_app = _load_flask_app(os.path.join(ROOT, "app.py"), "bia_app_module",
                           shadow={"tao_bia": os.path.join(ROOT, "tao_bia.py")})
mucluc_app = _load_flask_app(os.path.join(MUCLUC_DIR, "app.py"), "mucluc_app_module")
nhan_app = _load_flask_app(os.path.join(NHAN_DIR, "app.py"), "nhan_app_module")
hsxa_app = _load_flask_app(os.path.join(HSXA_DIR, "app.py"), "hsxa_app_module",
                            shadow={"tao_bia": os.path.join(HSXA_DIR, "tao_bia.py")})


# --------------------------- Trang vỏ (4 tab) --------------------------------
shell = Flask("shell")

SHELL_PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Bộ công cụ lưu trữ</title>
<style>
  *{box-sizing:border-box}
  html,body{height:100%}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,
       Helvetica,Arial,sans-serif;color:#e2e8f0;
       background:linear-gradient(180deg,#0f172a,#1e293b);
       display:flex;flex-direction:column;height:100vh;overflow:hidden;}

  .topbar{flex:0 0 auto;padding:14px 18px 0;}
  .brand{display:flex;align-items:center;gap:10px;margin-bottom:12px;}
  .brand .dot{width:30px;height:30px;border-radius:9px;display:grid;place-items:center;
       font-size:17px;background:linear-gradient(135deg,#4f46e5,#7c3aed);}
  .brand .t{font-weight:800;font-size:17px;letter-spacing:-.01em;color:#fff}
  .brand .s{color:#94a3b8;font-size:13px;margin-left:2px}

  .tabs{display:flex;gap:8px;flex-wrap:wrap;}
  .tab{appearance:none;border:1px solid rgba(148,163,184,.25);cursor:pointer;
       background:rgba(255,255,255,.04);color:#cbd5e1;font:inherit;font-weight:700;
       padding:11px 18px;border-radius:12px 12px 0 0;border-bottom:none;
       display:inline-flex;align-items:center;gap:9px;font-size:14px;
       transition:.15s;position:relative;top:1px;}
  .tab:hover{background:rgba(255,255,255,.09);color:#fff}
  .tab .ico{font-size:16px}
  .tab.active{color:#fff;background:#fff;}
  .tab.active.bia{background:linear-gradient(180deg,#eef2ff,#fff);color:#4338ca;
       box-shadow:0 -3px 0 #4f46e5 inset;}
  .tab.active.mucluc{background:linear-gradient(180deg,#e0f2fe,#fff);color:#0369a1;
       box-shadow:0 -3px 0 #0ea5e9 inset;}
  .tab.active.nhan{background:linear-gradient(180deg,#dcfce7,#fff);color:#047857;
       box-shadow:0 -3px 0 #059669 inset;}
  .tab.active.hsxa{background:linear-gradient(180deg,#fef3c7,#fff);color:#b45309;
       box-shadow:0 -3px 0 #d97706 inset;}

  .stage{flex:1 1 auto;background:#fff;position:relative;overflow:hidden;}
  .frame{position:absolute;inset:0;width:100%;height:100%;border:0;display:none;}
  .frame.active{display:block;}
  .loader{position:absolute;inset:0;display:none;place-items:center;
       color:#64748b;font-size:14px;background:#fff;}
  .loader.show{display:grid;}
</style>
</head>
<body>
  <div class="topbar">
    <div class="brand">
      <span class="dot">🗄️</span>
      <span class="t">Bộ công cụ lưu trữ</span>
      <span class="s">— 4 chương trình trong một trang</span>
    </div>
    <div class="tabs" id="tabs">
      <button class="tab bia active"    data-key="bia"    data-src="/bia/"><span class="ico">📚</span>Tạo bìa hồ sơ</button>
      <button class="tab mucluc"        data-key="mucluc" data-src="/mucluc/"><span class="ico">📑</span>Mục lục văn bản</button>
      <button class="tab nhan"          data-key="nhan"   data-src="/nhan/"><span class="ico">🏷️</span>In nhãn hộp</button>
      <button class="tab hsxa"          data-key="hsxa"   data-src="/hsxa/"><span class="ico">📂</span>Tạo bìa hồ sơ xã</button>
    </div>
  </div>

  <div class="stage" id="stage">
    <div class="loader show" id="loader">Đang tải…</div>
  </div>

<script>
  const tabs = [...document.querySelectorAll(".tab")];
  const stage = document.getElementById("stage");
  const loader = document.getElementById("loader");
  const frames = {};   // key -> iframe (nạp một lần rồi giữ lại)

  function activate(key){
    tabs.forEach(t => t.classList.toggle("active", t.dataset.key === key));
    Object.entries(frames).forEach(([k,f]) => f.classList.toggle("active", k === key));

    if(!frames[key]){
      const tab = tabs.find(t => t.dataset.key === key);
      const f = document.createElement("iframe");
      f.className = "frame active";
      f.src = tab.dataset.src;
      loader.classList.add("show");
      f.addEventListener("load", () => loader.classList.remove("show"), {once:true});
      stage.appendChild(f);
      frames[key] = f;
    } else {
      loader.classList.remove("show");
    }
  }

  tabs.forEach(t => t.addEventListener("click", () => activate(t.dataset.key)));
  activate("bia");   // tab mặc định
</script>
</body>
</html>
"""


@shell.route("/")
def home():
    return render_template_string(SHELL_PAGE)


# App WSGI hợp nhất: mặc định là trang vỏ, 4 tiền tố dẫn tới 4 app con.
application = DispatcherMiddleware(shell, {
    "/bia": bia_app,
    "/mucluc": mucluc_app,
    "/nhan": nhan_app,
    "/hsxa": hsxa_app,
})


if __name__ == "__main__":
    import socket
    from werkzeug.serving import run_simple

    port = int(os.environ.get("PORT", "8000"))
    try:
        ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        ip = "127.0.0.1"
    print("Truy cập:  http://127.0.0.1:%d   |   http://%s:%d" % (port, ip, port))
    run_simple("0.0.0.0", port, application, threaded=True, use_reloader=False)
