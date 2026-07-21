# Hồ sơ — Tạo bìa hồ sơ lưu trữ tự động

Sinh **bìa hồ sơ lưu trữ** (.docx, khổ A4) hàng loạt từ một tệp dữ liệu **Excel (.xlsx)**.
Giữ nguyên font, khung viền và bố cục của tệp mẫu; mỗi dòng dữ liệu là một bìa, mỗi bìa một trang.

## Tính năng

- Đọc Excel, tự bỏ qua dòng tiêu đề, chuẩn hóa ngày về `dd/mm/yyyy` (thiếu số 0 thì tự thêm,
  chỉ tháng/năm → `mm/yyyy`, dạng dính liền `ddmmyyyy` → `dd/mm/yyyy`).
- Điền dữ liệu theo **nội dung** (không phụ thuộc số thứ tự đoạn) nên bền với thay đổi mẫu.
- **Khối dưới** (Phông số, ngày, số trang…) được neo cố định ở **đáy trang** bằng một text box
  canh đáy riêng — tiêu đề dài bao nhiêu cũng không làm xê dịch.
- Giao diện web hiện đại: kéo-thả **nhiều** tệp Excel → xử lý **tuần tự** với
  **thanh tiến độ realtime** (không đơ UI) → tải về `.docx` (1 tệp) hoặc `.zip` (nhiều tệp).

## Cấu trúc cột Excel (sheet đầu)

| Cột | Nội dung |
|-----|----------|
| A | Hồ sơ số |
| B | Nội dung / Tiêu đề hồ sơ |
| C | Thời gian (bắt đầu-kết thúc) |
| D | Số trang |
| E | Số TL |
| F | Thời hạn bảo quản |
| G | Dòng phông (đầu bìa) |

## Chạy toàn bộ (3 tab) với uv — khuyến nghị

Dự án gộp **3 chương trình** vào một trang duy nhất với 3 tab (`server.py`):
📚 Tạo bìa hồ sơ · 📑 Mục lục văn bản · 🏷️ In nhãn hộp.

```bash
uv run server.py
```

`uv` tự tạo `.venv` đúng nền tảng và cài thư viện theo `pyproject.toml` + `uv.lock`
ngay lần chạy đầu (chưa có uv? xem https://docs.astral.sh/uv/ hoặc chạy `install.sh`).
Sau đó mở http://127.0.0.1:8000 (đổi cổng bằng `PORT=5019 uv run server.py`).

### Chạy riêng từng chương trình (tuỳ chọn)

Mỗi app vẫn chạy độc lập được như cũ:

```bash
uv run app.py                       # 📚 bìa hồ sơ      -> http://127.0.0.1:5019
uv run Phanmem_muclucdang/app.py    # 📑 mục lục văn bản -> http://127.0.0.1:5050
uv run Innhan/app.py                # 🏷️ in nhãn hộp     -> http://127.0.0.1:5051
```

## Cài đặt nhanh (tự động) — cho máy chủ/triển khai

```bash
chmod +x install.sh
./install.sh                # tạo .venv, cài thư viện + gunicorn, tạo service & chạy
./install.sh --no-service   # chỉ cài thư viện, không tạo service
```

Script tự tạo service:
- **Linux**: `systemd` (`hoso.service`) — tự khởi động cùng máy.
- **macOS**: `launchd` (`com.hoso.bia`). Lưu ý: nếu đặt dự án trong `~/Documents`
  (bị macOS bảo vệ TCC), launchd có thể bị chặn — hãy đặt dự án ngoài `~/Documents`
  hoặc cấp *Full Disk Access*, hoặc chạy thủ công bằng gunicorn.

Cấu hình qua biến môi trường: `PORT` (mặc định 5019), `HOST` (0.0.0.0),
`WORKERS` (giữ = 1), `THREADS` (8).

## Cài đặt thủ công

```bash
uv sync --group deploy      # tạo .venv + cài thư viện (+ gunicorn cho triển khai)
```

Hoặc dùng pip truyền thống với `requirements.txt` (do uv xuất ra, luôn đồng bộ với `uv.lock`):

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt gunicorn
```

Đặt tệp mẫu `.docx` vào thư mục dự án (mặc định dùng `Sample.docx`).
Có thể đổi thư mục mẫu bằng biến môi trường `BIA_TEMPLATE_DIR`.

## Chạy web (trang 3 tab hợp nhất)

```bash
# Phát triển
uv run server.py

# Sản xuất (gunicorn, 1 worker + nhiều thread) — phục vụ cả 3 tab
uv run --group deploy gunicorn --workers 1 --threads 8 --bind 0.0.0.0:5019 server:application
```

Mở: http://127.0.0.1:8000 (dev) hoặc cổng đã cấu hình. Lắng nghe trên `0.0.0.0`, dùng được trong LAN.

> **Vì sao `--workers 1`?** File kết quả được giữ tạm trong RAM của tiến trình theo
> token tải về. Nhiều worker sẽ khiến lượt tải rơi vào tiến trình khác (mất token).
> Cần chịu tải cao hơn thì tăng `--threads`, đừng tăng worker.

## Chạy dòng lệnh (CLI)

```bash
uv run tao_bia.py --mau Sample.docx --data "du_lieu.xlsx" --out ket_qua.docx
```

## Tệp chính

- `server.py` — **máy chủ hợp nhất**: gộp 3 app thành một trang 3 tab (`server:application`).
- `pyproject.toml` / `uv.lock` — khai báo & khoá phiên bản thư viện cho uv.
- `app.py` + `tao_bia.py` — 📚 Tạo bìa hồ sơ (web + CLI).
- `Phanmem_muclucdang/app.py` + `muc_luc.py` — 📑 Mục lục văn bản.
- `Innhan/app.py` + `in_nhan.py` — 🏷️ In nhãn hộp.
