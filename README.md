# Bộ công cụ lưu trữ hồ sơ (Di Văn)

Gộp **3 chương trình** xử lý hồ sơ lưu trữ vào **một trang web duy nhất với 3 tab**
(`server.py`). Mỗi chương trình đọc **Excel (.xlsx)** hàng loạt rồi xuất ra **Word (.docx)**;
nhiều kết quả được đóng gói **.zip** để tải một lần. Giao diện kéo-thả nhiều tệp, xử lý
**tuần tự** với **thanh tiến độ realtime** (không đơ UI).

| Tab | Chương trình | Nguồn | Kết quả |
|-----|--------------|-------|---------|
| 📚 | **Tạo bìa hồ sơ** | `app.py` + `tao_bia.py` | Bìa A4, mỗi hồ sơ một trang |
| 📑 | **Mục lục văn bản, tài liệu** | `Phanmem_muclucdang/app.py` + `muc_luc.py` | Word A4 **nằm ngang**, mỗi sheet một tệp |
| 🏷️ | **In nhãn hộp lưu trữ** | `Innhan/app.py` + `in_nhan.py` | Word A4 **dọc**, 4 nhãn/trang |

## Bắt đầu nhanh

```bash
uv run server.py            # mở http://127.0.0.1:8000
PORT=5019 uv run server.py  # đổi cổng
```

`uv` tự tạo `.venv` đúng nền tảng và cài thư viện theo `pyproject.toml` + `uv.lock`
ngay lần chạy đầu — không cần kích hoạt môi trường ảo thủ công.
Chưa có `uv`? Xem https://docs.astral.sh/uv/ hoặc chạy `./install.sh` (tự cài `uv` giúp bạn).

Trang chủ (`/`) là vỏ chứa 3 tab; mỗi tab nạp một chương trình con qua đường dẫn
`/bia/`, `/mucluc/`, `/nhan/` trong **cùng một tiến trình Flask**.

## Ba chương trình & cấu trúc Excel

### 📚 Tạo bìa hồ sơ  — tab `/bia/`

Sinh **bìa hồ sơ** (.docx, A4 dọc) — mỗi dòng dữ liệu là một bìa, mỗi bìa một trang;
giữ nguyên font, khung viền và bố cục của tệp mẫu. Nhập thêm **Phông số / Mục lục số**.
Đọc **sheet đầu**, tự bỏ qua dòng tiêu đề, chuẩn hóa ngày về `dd/mm/yyyy`.

| Cột | Nội dung |
|-----|----------|
| A | Hồ sơ số |
| B | Nội dung / Tiêu đề hồ sơ |
| C | Thời gian (bắt đầu-kết thúc) |
| D | Số trang |
| E | Số TL |
| F | Thời hạn bảo quản |
| G | Dòng phông (đầu bìa) |

### 📑 Mục lục văn bản, tài liệu  — tab `/mucluc/`

Mỗi **sheet** trong mỗi tệp Excel → **một tệp Word A4 nằm ngang** (14 cột), tiện in từng
phông/đơn vị. Nhập **Phông số / Mục lục số** để ghép **Mã hồ sơ** =
`Phông số - Mục lục số - Hồ sơ số` (vd `A47.35.21-001-01-0001`).

| Cột | Nội dung | Cột | Nội dung |
|-----|----------|-----|----------|
| A | Hồ sơ số | H | Độ mật |
| B | Số thứ tự | I | Loại bản |
| C | Số, ký hiệu | J | Trang số |
| D | Ngày tháng | K | Số trang |
| E | Tên loại và trích yếu | L | Từ khoá |
| F | Tác giả | M | Ghi chú |
| G | Người ký | N | Số lượng tệp |
| | | O | Tên tệp tài liệu |

### 🏷️ In nhãn hộp lưu trữ  — tab `/nhan/`

Mỗi **tệp** Excel → **một tệp Word A4 dọc** (gộp mọi sheet), 4 nhãn/trang, dựng theo khuôn
`Innhan/sample.docx`. Mỗi dòng dữ liệu là một nhãn; dòng 1 là tiêu đề.

| Cột | Nội dung |
|-----|----------|
| A | Hộp số |
| B | Từ hồ sơ số |
| C | Đến hồ sơ số |
| D | Nội dung (mã mục/nhóm, mỗi dòng một mục) |
| E | Tiêu đề (Phông lưu trữ…) |

## Chạy

```bash
# Toàn bộ 3 tab (khuyến nghị)
uv run server.py

# Sản xuất — gunicorn phục vụ cả 3 tab
uv run --group deploy gunicorn --workers 1 --threads 8 --bind 0.0.0.0:5019 server:application
```

Mở http://127.0.0.1:8000 (dev) hoặc cổng đã cấu hình. Lắng nghe trên `0.0.0.0` nên dùng được trong LAN.

Chạy **riêng** từng chương trình (vẫn hoạt động độc lập như cũ):

```bash
uv run app.py                     # 📚 -> http://127.0.0.1:5019
uv run Phanmem_muclucdang/app.py  # 📑 -> http://127.0.0.1:5050
uv run Innhan/app.py              # 🏷️ -> http://127.0.0.1:5051
```

Dòng lệnh (CLI) tạo bìa không cần web:

```bash
uv run tao_bia.py --mau Sample.docx --data "du_lieu.xlsx" --out ket_qua.docx
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

`install.sh` tạo service tự khởi động:
- **Linux**: `systemd` (`hoso.service`).
- **macOS**: `launchd` (`com.hoso.bia`) — nếu đặt dự án trong `~/Documents` (bị macOS bảo vệ
  TCC) launchd có thể bị chặn; hãy đặt ngoài `~/Documents`, cấp *Full Disk Access*, hoặc
  chạy gunicorn thủ công.

Cấu hình qua biến môi trường: `PORT` (mặc định 5019 khi triển khai / 8000 khi `uv run server.py`),
`HOST` (0.0.0.0), `WORKERS` (giữ = 1), `THREADS` (8), `BIA_TEMPLATE_DIR` (thư mục chứa tệp mẫu bìa).

## Ghi chú kỹ thuật

> **Vì sao `--workers 1`?** Cả ba chương trình giữ tệp kết quả **tạm trong RAM** của tiến
> trình theo mã công việc. Nhiều worker sẽ khiến lượt tải rơi vào tiến trình khác (mất kết
> quả). Cần chịu tải cao hơn thì tăng `--threads`, **đừng** tăng worker.

- `server.py` dùng `DispatcherMiddleware` mount 3 app dưới `/bia`, `/mucluc`, `/nhan`. Phần
  JavaScript mỗi app gọi API theo tiền tố `BASE` lấy từ `location.pathname`, nên chạy đúng cả
  khi gộp lẫn khi chạy độc lập.
- Tệp mẫu `.docx`: **bìa** dùng `Sample.docx` ở thư mục gốc (đổi bằng `BIA_TEMPLATE_DIR`);
  **nhãn** dùng `Innhan/sample.docx`; **mục lục** dựng bằng code (không cần tệp mẫu).

## Tệp chính

- `server.py` — máy chủ hợp nhất (`server:application`), trang 3 tab.
- `pyproject.toml` / `uv.lock` / `.python-version` — khai báo & khoá thư viện cho uv.
- `app.py` + `tao_bia.py` — 📚 Tạo bìa hồ sơ (web + CLI).
- `Phanmem_muclucdang/app.py` + `muc_luc.py` — 📑 Mục lục văn bản.
- `Innhan/app.py` + `in_nhan.py` — 🏷️ In nhãn hộp.
- `install.sh` — cài đặt & tạo service; `deploy/nginx/` — cấu hình reverse proxy.
