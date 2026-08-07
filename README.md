# Bộ công cụ lưu trữ hồ sơ (Di Văn)

Gộp **5 chương trình** xử lý hồ sơ lưu trữ vào **một trang web duy nhất với 5 tab**
(`server.py`). Mỗi chương trình đọc **Excel (.xlsx)** rồi xuất ra **Word (.docx)**.
Giao diện kéo-thả nhiều tệp, xử lý **tuần tự** với **thanh tiến độ realtime** (không đơ UI);
4 chương trình đầu đóng gói **.zip** khi có nhiều kết quả.

| Tab | Chương trình | Nguồn | Kết quả |
|-----|--------------|-------|---------|
| 📂 | **Tạo bìa hồ sơ xã** *(mặc định)* | `HSxa/app.py` + `HSxa/tao_bia.py` | Bìa A4, mỗi hồ sơ một trang |
| 📚 | **Tạo bìa hồ sơ** | `app.py` + `tao_bia.py` | Bìa A4, mỗi hồ sơ một trang |
| 📑 | **Mục lục văn bản, tài liệu** | `Phanmem_muclucdang/app.py` + `muc_luc.py` | Word A4 **nằm ngang** (14 cột), mỗi sheet một tệp |
| 🏷️ | **In nhãn hộp lưu trữ xã** | `Innhan_xa/app.py` + `in_nhan.py` | Word A4 **dọc**, 4 nhãn/trang |
| 📋 | **Mục lục xã (Excel → Word)** | `Muc_luc_Xa/app_mucluc.py` | Word A4 **dọc** (7 cột), mỗi hồ sơ một trang |

## Bắt đầu nhanh

```bash
uv run server.py            # mở http://127.0.0.1:8000
PORT=5019 uv run server.py  # đổi cổng
```

`uv` tự tạo `.venv` đúng nền tảng và cài thư viện theo `pyproject.toml` + `uv.lock`
ngay lần chạy đầu — không cần kích hoạt môi trường ảo thủ công.
Chưa có `uv`? Xem https://docs.astral.sh/uv/ hoặc chạy `./install.sh` (tự cài `uv` giúp bạn).

Trang chủ (`/`) là vỏ chứa 5 tab; mỗi tab nạp một chương trình con qua đường dẫn
`/hsxa/`, `/bia/`, `/mucluc/`, `/nhan/`, `/mlxa/` trong **cùng một tiến trình Flask**.

## Năm chương trình & cấu trúc Excel

### 📂 Tạo bìa hồ sơ xã — tab `/hsxa/`

Sinh **bìa hồ sơ** (.docx, A4 dọc) theo khuôn `HSxa/Bìa HS.docx` — mỗi dòng dữ liệu là một
bìa, mỗi bìa một trang; giữ nguyên font, khung viền và bố cục của tệp mẫu.
Khác bản 📚 ở chỗ **Phông số / Mục lục số đọc theo từng dòng** (cột A/B); dòng nào bỏ trống
thì lấy giá trị mặc định nhập ở giao diện. Đọc **sheet đầu**, tự bỏ qua dòng tiêu đề.

| Cột | Nội dung | Cột | Nội dung |
|-----|----------|-----|----------|
| A | Phông số | E | Ngày tháng bắt đầu-kết thúc (vd `16/02/2006-19/02/2008`) |
| B | Mục lục số | F | Số tờ |
| C | Hồ sơ số | G | Thời hạn bảo quản (vd `Vĩnh viễn`) |
| D | Trích yếu nội dung | H | Tên phông (dòng đầu bìa, vd `ỦY BAN NHÂN DÂN XÃ …`) |

### 📚 Tạo bìa hồ sơ — tab `/bia/`

Sinh **bìa hồ sơ** (.docx, A4 dọc) theo tệp mẫu `Sample.docx` ở thư mục gốc — mỗi dòng dữ
liệu là một bìa, mỗi bìa một trang. **Phông số / Mục lục số** nhập một lần ở giao diện
(mặc định `001` / `01`). Đọc **sheet đầu**, tự bỏ qua dòng tiêu đề, chuẩn hóa ngày về `dd/mm/yyyy`.

| Cột | Nội dung |
|-----|----------|
| A | Hồ sơ số |
| B | Nội dung / Tiêu đề hồ sơ |
| C | Thời gian (bắt đầu-kết thúc) |
| D | Số trang |
| E | Số TL (số tờ) |
| F | Thời hạn bảo quản |
| G | Dòng phông / đảng bộ (đầu bìa) |

### 📑 Mục lục văn bản, tài liệu — tab `/mucluc/`

Mỗi **sheet** trong mỗi tệp Excel → **một tệp Word A4 nằm ngang** (14 cột), tiện in từng
phông/đơn vị. Nhập **Phông số / Mục lục số** để ghép **Mã hồ sơ** =
`Phông số - Mục lục số - Hồ sơ số` (vd `A47.35.21-001-01-0001`).

| Cột | Nội dung | Cột | Nội dung |
|-----|----------|-----|----------|
| A | Hồ sơ số (gộp nhóm) | I | Loại bản |
| B | Số thứ tự | J | Trang số |
| C | Số, ký hiệu | K | Số trang |
| D | Ngày tháng | L | Từ khoá |
| E | Tên loại và trích yếu | M | Ghi chú |
| F | Tác giả | N | Số lượng tệp |
| G | Người ký | O | Tên tệp tài liệu |
| H | Độ mật | | |

### 🏷️ In nhãn hộp lưu trữ xã — tab `/nhan/`

Mỗi **tệp** Excel → **một tệp Word A4 dọc** (gộp mọi sheet), 4 nhãn/trang, dựng theo khuôn
`Innhan_xa/Sample.docx`. Mỗi dòng dữ liệu là một nhãn.

Chương trình **đọc theo TÊN CỘT** (không phân biệt hoa/thường và khoảng trắng), nên thứ tự
cột đặt thế nào cũng được và thừa cột cũng không sao:

| Tên cột | Nội dung |
|---------|----------|
| Số hộp | Số hộp (tự đệm 0 thành 2 chữ số) |
| Từ Hồ sơ | Hồ sơ số đầu (tự đệm 0 thành 4 chữ số) |
| Đến Hồ sơ | Hồ sơ số cuối |
| Số lượng hồ sơ | Số ĐVBQ — bỏ trống thì tự suy ra từ Từ/Đến |
| Tên phông | Tiêu đề nhãn |
| Năm | Năm (chỉ lấy phần năm nếu Excel hiểu nhầm thành ngày) |

Nếu sheet **không có dòng tiêu đề** nhận dạng được, chương trình đọc theo thứ tự cột
**A-F** đúng như bảng trên.

### 📋 Mục lục xã (Excel → Word) — tab `/mlxa/`

Mỗi **hồ sơ** (nhóm theo cột A) → **một trang Word A4 dọc**, đầu trang là
`MỤC LỤC VĂN BẢN` + `Hồ sơ số: …`, bên dưới là bảng 7 cột. Các hồ sơ được sắp xếp theo số
và ngăn nhau bằng dấu ngắt trang; cột A để trống sẽ tự nối tiếp hồ sơ ở dòng trên.

Khác 4 chương trình kia: **không đóng gói .zip** — mỗi tệp Excel được chuyển xong là trình
duyệt tải ngay tệp `.docx` cùng tên về máy. Toàn bộ xử lý diễn ra **trong RAM**, server
không ghi tệp nào lên đĩa. Giới hạn **16 MB/tệp tải lên**.

| Cột | Nội dung |
|-----|----------|
| A | Hồ sơ số (gộp nhóm — mỗi hồ sơ một trang) |
| B | Số TT |
| C | Số, ký hiệu tài liệu |
| D | Ngày tháng văn bản |
| E | Tác giả tài liệu |
| F | Tên loại và trích yếu nội dung văn bản |
| G | Tờ số |
| H | Ghi chú |

Tệp Excel cần **ít nhất 8 cột** và **dòng 1 là tiêu đề** (luôn bị bỏ qua).

## Chạy

```bash
# Toàn bộ 5 tab (khuyến nghị)
uv run server.py

# Sản xuất — gunicorn phục vụ cả 5 tab
uv run --group deploy gunicorn --workers 1 --threads 8 --bind 0.0.0.0:5019 server:application
```

Mở http://127.0.0.1:8000 (dev) hoặc cổng đã cấu hình. Lắng nghe trên `0.0.0.0` nên dùng được trong LAN.

Chạy **riêng** từng chương trình (vẫn hoạt động độc lập như cũ):

```bash
uv run HSxa/app.py                     # 📂 -> http://127.0.0.1:5019
uv run app.py                          # 📚 -> http://127.0.0.1:5019
uv run Phanmem_muclucdang/app.py       # 📑 -> http://127.0.0.1:5050
uv run Innhan_xa/app.py                # 🏷️ -> http://127.0.0.1:5050
uv run Muc_luc_Xa/app_mucluc.py        # 📋 -> http://127.0.0.1:5000
```

> Cổng của 📂/📚 trùng nhau (5019) và 📑/🏷️ cũng trùng nhau (5050) — **không chạy đồng thời
> hai chương trình cùng cổng**. Chỉ 📋 nhận biến `PORT` để đổi cổng; muốn chạy tất cả cùng lúc
> thì dùng `server.py`.

Dòng lệnh (CLI) không cần web:

```bash
uv run tao_bia.py --mau Sample.docx --data "du_lieu.xlsx" --out ket_qua.docx
uv run HSxa/tao_bia.py --mau "HSxa/Bìa HS.docx" --data "du_lieu.xlsx" \
        --out ket_qua.docx --phong-so 001 --muc-luc 01
uv run Phanmem_muclucdang/muc_luc.py "du_lieu.xlsx" --phong 001 --muc-luc 01 --outdir .
```

## Cài đặt

```bash
# Tự động (máy chủ/triển khai): cài uv nếu thiếu, uv sync, tạo service & chạy
chmod +x install.sh
./install.sh                # tạo .venv + thư viện + gunicorn, tạo service
./install.sh --no-service   # chỉ cài môi trường, không tạo service

# Thủ công
uv sync --group deploy      # .venv + thư viện (+ gunicorn cho triển khai)
```

Hoặc dùng pip truyền thống với `requirements.txt` (do uv xuất ra, luôn đồng bộ `uv.lock`):

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt gunicorn
```

Thư viện cần có: `flask`, `python-docx`, `openpyxl`, `lxml`, `pandas` (📋 dùng pandas để đọc Excel).

`install.sh` tạo service tự khởi động:
- **Linux**: `systemd` (`hoso.service`).
- **macOS**: `launchd` (`com.hoso.bia`) — nếu đặt dự án trong `~/Documents` (bị macOS bảo vệ
  TCC) launchd có thể bị chặn; hãy đặt ngoài `~/Documents`, cấp *Full Disk Access*, hoặc
  chạy gunicorn thủ công.

Cấu hình qua biến môi trường: `PORT` (mặc định 5019 khi triển khai / 8000 khi `uv run server.py`),
`HOST` (0.0.0.0), `WORKERS` (giữ = 1), `THREADS` (8), `BIA_TEMPLATE_DIR` (thư mục chứa tệp mẫu bìa,
dùng cho 📂 và 📚), `TEM_TEMPLATE` (tệp khuôn nhãn của 🏷️).

## Ghi chú kỹ thuật

> **Vì sao `--workers 1`?** Bốn chương trình 📂 📚 📑 🏷️ giữ tệp kết quả **tạm trong RAM** của
> tiến trình theo mã công việc. Nhiều worker sẽ khiến lượt tải rơi vào tiến trình khác (mất kết
> quả). Cần chịu tải cao hơn thì tăng `--threads`, **đừng** tăng worker.
> (📋 trả tệp thẳng trong một request nên không bị ràng buộc này, nhưng vẫn dùng chung tiến trình.)

- `server.py` dùng `DispatcherMiddleware` mount 5 app dưới `/hsxa`, `/bia`, `/mucluc`, `/nhan`,
  `/mlxa`. Phần JavaScript mỗi app gọi API theo tiền tố `BASE` lấy từ `location.pathname`, nên
  chạy đúng cả khi gộp lẫn khi chạy độc lập.
- Bốn tệp `app.py` trùng tên được nạp bằng `importlib` theo **đường dẫn thật** với tên module
  riêng, nên không đè lên nhau. Hai bản `tao_bia.py` (gốc và `HSxa/`) được "che" (shadow) tạm
  trong `sys.modules` lúc nạp để mỗi app dùng đúng bản của mình.
- Tệp mẫu `.docx`: 📚 dùng `Sample.docx` ở thư mục gốc; 📂 dùng `HSxa/Bìa HS.docx` (tự bỏ qua
  các tệp `Bia_ho_so*`, `Sample*`, `~$*`); cả hai đổi được bằng `BIA_TEMPLATE_DIR`.
  🏷️ dùng `Innhan_xa/Sample.docx` (đổi bằng `TEM_TEMPLATE`). 📑 và 📋 dựng bằng code, không cần mẫu.
- `deploy/nginx/` đặt `client_max_body_size 50m` và `proxy_read_timeout 300s` — đủ cho giới hạn
  16 MB của 📋 và các lượt xử lý dài.

## Tệp chính

- `server.py` — máy chủ hợp nhất (`server:application`), trang 5 tab.
- `pyproject.toml` / `uv.lock` / `.python-version` — khai báo & khoá thư viện cho uv.
- `HSxa/app.py` + `HSxa/tao_bia.py` — 📂 Tạo bìa hồ sơ xã (web + CLI).
- `app.py` + `tao_bia.py` — 📚 Tạo bìa hồ sơ (web + CLI).
- `Phanmem_muclucdang/app.py` + `muc_luc.py` — 📑 Mục lục văn bản (web + CLI).
- `Innhan_xa/app.py` + `in_nhan.py` — 🏷️ In nhãn hộp xã.
- `Muc_luc_Xa/app_mucluc.py` — 📋 Mục lục xã (một tệp, tự chứa).
- `install.sh` — cài đặt & tạo service; `deploy/nginx/` — cấu hình reverse proxy.
