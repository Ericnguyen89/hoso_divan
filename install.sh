#!/usr/bin/env bash
#
# Cài đặt tự động ứng dụng "Hoso - Tạo bìa hồ sơ".
#   - Tạo môi trường ảo Python (.venv) và cài thư viện cần thiết + gunicorn
#   - Tạo service tự khởi động:
#       * Linux  -> systemd  (hoso.service)
#       * macOS  -> launchd  (com.hoso.bia)
#
# Dùng:
#   chmod +x install.sh
#   ./install.sh                # cài + tạo service + chạy
#   ./install.sh --no-service   # chỉ cài thư viện, không tạo service
#
set -euo pipefail

# ----------------------------- Cấu hình --------------------------------------
PORT="${PORT:-5019}"
HOST="${HOST:-0.0.0.0}"
WORKERS="${WORKERS:-1}"        # GIỮ =1: file kết quả lưu trong RAM của 1 tiến trình
THREADS="${THREADS:-8}"
SERVICE_NAME="hoso"
APP_MODULE="server:application"   # app WSGI hợp nhất (3 tab) trong server.py

# Thư mục chứa script này (= thư mục dự án)
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$APP_DIR/.venv"
PY_BIN="$VENV/bin/python"
GUNICORN="$VENV/bin/gunicorn"

MAKE_SERVICE=1
[ "${1:-}" = "--no-service" ] && MAKE_SERVICE=0

say()  { printf "\033[1;36m==>\033[0m %s\n" "$*"; }
ok()   { printf "\033[1;32m✓\033[0m %s\n" "$*"; }
warn() { printf "\033[1;33m!\033[0m %s\n" "$*"; }
die()  { printf "\033[1;31m✗ %s\033[0m\n" "$*" >&2; exit 1; }

# --------------------------- 1) uv + môi trường ------------------------------
# Cài uv nếu chưa có (trình quản lý gói/môi trường Python siêu nhanh của Astral).
UV="$(command -v uv || true)"
[ -z "$UV" ] && [ -x "$HOME/.local/bin/uv" ] && UV="$HOME/.local/bin/uv"
if [ -z "$UV" ]; then
  say "Chưa có uv — đang cài đặt uv…"
  if command -v curl >/dev/null 2>&1; then
    curl -LsSf https://astral.sh/uv/install.sh | sh
  elif command -v wget >/dev/null 2>&1; then
    wget -qO- https://astral.sh/uv/install.sh | sh
  else
    die "Cần curl hoặc wget để cài uv. Hãy cài uv thủ công: https://docs.astral.sh/uv/"
  fi
  export PATH="$HOME/.local/bin:$PATH"
  UV="$(command -v uv || echo "$HOME/.local/bin/uv")"
fi
[ -x "$UV" ] || command -v uv >/dev/null 2>&1 || die "Không tìm thấy uv sau khi cài."
say "Dùng uv: $("$UV" --version)"

# uv tự tạo .venv (đúng nền tảng), cài thư viện theo pyproject.toml + uv.lock,
# kèm nhóm 'deploy' để có gunicorn. Chạy trong thư mục dự án.
say "Tạo môi trường & cài thư viện (uv sync --group deploy)"
( cd "$APP_DIR" && "$UV" sync --group deploy )
ok "Đã tạo .venv + cài thư viện + gunicorn"

# Kiểm tra import nhanh
"$PY_BIN" - <<'PY'
import flask, docx, openpyxl, lxml  # noqa
print("   thư viện OK")
PY

if [ "$MAKE_SERVICE" -eq 0 ]; then
  ok "Hoàn tất (bỏ qua tạo service)."
  echo "Chạy nhanh (dev):  uv run server.py"
  echo "Chạy sản xuất  :  $GUNICORN --workers $WORKERS --threads $THREADS --bind $HOST:$PORT $APP_MODULE"
  echo "               (hoặc: uv run --group deploy gunicorn --bind $HOST:$PORT $APP_MODULE)"
  exit 0
fi

EXEC="$GUNICORN --workers $WORKERS --threads $THREADS --timeout 120 --bind $HOST:$PORT $APP_MODULE"
OS="$(uname -s)"

# ------------------------------ 2) Service -----------------------------------
if [ "$OS" = "Linux" ] && command -v systemctl >/dev/null 2>&1; then
  say "Tạo systemd service: $SERVICE_NAME.service"
  RUN_USER="${SUDO_USER:-$USER}"
  UNIT="/etc/systemd/system/$SERVICE_NAME.service"
  SUDO=""; [ "$(id -u)" -ne 0 ] && SUDO="sudo"

  $SUDO tee "$UNIT" >/dev/null <<EOF
[Unit]
Description=Hoso - Tao bia ho so luu tru (Flask + gunicorn)
After=network.target

[Service]
Type=simple
User=$RUN_USER
WorkingDirectory=$APP_DIR
ExecStart=$EXEC
Restart=always
RestartSec=3
Environment=BIA_TEMPLATE_DIR=$APP_DIR

[Install]
WantedBy=multi-user.target
EOF

  $SUDO systemctl daemon-reload
  $SUDO systemctl enable "$SERVICE_NAME" >/dev/null 2>&1 || true
  $SUDO systemctl restart "$SERVICE_NAME"
  sleep 1
  ok "Service đang chạy."
  echo
  echo "  Trạng thái :  sudo systemctl status $SERVICE_NAME"
  echo "  Nhật ký    :  sudo journalctl -u $SERVICE_NAME -f"
  echo "  Khởi động  :  sudo systemctl start $SERVICE_NAME"
  echo "  Dừng       :  sudo systemctl stop $SERVICE_NAME"

elif [ "$OS" = "Darwin" ]; then
  say "Tạo launchd service (macOS): com.$SERVICE_NAME.bia"
  PLIST="$HOME/Library/LaunchAgents/com.$SERVICE_NAME.bia.plist"
  mkdir -p "$HOME/Library/LaunchAgents"
  cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>com.$SERVICE_NAME.bia</string>
  <key>WorkingDirectory</key><string>$APP_DIR</string>
  <key>EnvironmentVariables</key>
  <dict><key>BIA_TEMPLATE_DIR</key><string>$APP_DIR</string></dict>
  <key>ProgramArguments</key>
  <array>
    <string>$GUNICORN</string>
    <string>--workers</string><string>$WORKERS</string>
    <string>--threads</string><string>$THREADS</string>
    <string>--timeout</string><string>120</string>
    <string>--bind</string><string>$HOST:$PORT</string>
    <string>$APP_MODULE</string>
  </array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>StandardOutPath</key><string>$APP_DIR/hoso.log</string>
  <key>StandardErrorPath</key><string>$APP_DIR/hoso.err.log</string>
</dict>
</plist>
EOF
  launchctl unload "$PLIST" 2>/dev/null || true
  launchctl load "$PLIST"
  sleep 1
  ok "Service đã nạp (tự chạy khi đăng nhập)."
  echo
  echo "  Nhật ký   :  tail -f $APP_DIR/hoso.log"
  echo "  Dừng      :  launchctl unload $PLIST"
  echo "  Chạy lại  :  launchctl load $PLIST"
else
  warn "Không phát hiện systemd/launchd. Bỏ qua tạo service."
  echo "Chạy thủ công:  $EXEC"
fi

echo
IP="$( (command -v hostname >/dev/null && hostname -I 2>/dev/null | awk '{print $1}') || true )"
[ -z "$IP" ] && IP="$( (ipconfig getifaddr en0 2>/dev/null) || true )"
ok "Hoàn tất! Truy cập:"
echo "    http://127.0.0.1:$PORT"
[ -n "$IP" ] && echo "    http://$IP:$PORT   (trong mạng LAN)"
