#!/data/data/com.termux/files/usr/bin/bash
# ============================================================
#  setup_v9.sh — Setup otomatis untuk v9.py di Termux
#  Jalankan: bash setup_v9.sh
# ============================================================

set -e  # berhenti jika ada error

CYAN='\033[0;36m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # No Color

log()  { echo -e "${CYAN}[INFO]${NC} $1"; }
ok()   { echo -e "${GREEN}[ OK ]${NC} $1"; }
warn() { echo -e "${YELLOW}[WARN]${NC} $1"; }
err()  { echo -e "${RED}[ERR ]${NC} $1"; }

echo ""
echo -e "${CYAN}============================================================${NC}"
echo -e "${CYAN}   Setup v9.py — PDF Scan → Excel Converter (Termux)       ${NC}"
echo -e "${CYAN}============================================================${NC}"
echo ""

# ------------------------------------------------------------
# 1. Update & upgrade Termux packages
# ------------------------------------------------------------
log "Update package list Termux..."
pkg update -y && pkg upgrade -y
ok "Termux packages updated."

# ------------------------------------------------------------
# 2. Install package sistem yang dibutuhkan
# ------------------------------------------------------------
log "Menginstall package sistem: python, poppler, tesseract, mariadb..."
pkg install -y python poppler tesseract tesseract-lang mariadb
ok "Package sistem terinstall."

# ------------------------------------------------------------
# 3. Setup storage (akses ke shared storage Android)
# ------------------------------------------------------------
log "Setup storage Termux..."
if [ ! -d "$HOME/storage" ]; then
    termux-setup-storage
    sleep 3
    ok "Storage setup selesai."
else
    ok "Storage sudah pernah di-setup sebelumnya."
fi

# Buat folder kerja
WORKDIR="$HOME/storage/shared/python"
log "Membuat folder kerja: $WORKDIR"
mkdir -p "$WORKDIR"
ok "Folder $WORKDIR siap."

# ------------------------------------------------------------
# 4. Install Python packages
# ------------------------------------------------------------
log "Upgrade pip..."
pip install --upgrade pip --break-system-packages -q

log "Menginstall Python packages (pytesseract, Pillow, openpyxl, pdf2image, mysql-connector-python)..."
pip install \
    pytesseract \
    Pillow \
    openpyxl \
    pdf2image \
    mysql-connector-python \
    --break-system-packages -q
ok "Python packages terinstall."

# ------------------------------------------------------------
# 5. Setup MariaDB (MySQL)
# ------------------------------------------------------------
log "Menyiapkan MariaDB..."

# Install DB jika belum ada
if [ ! -d "$PREFIX/var/lib/mysql" ]; then
    log "Inisialisasi database MariaDB..."
    mysql_install_db
    ok "Database MariaDB diinisialisasi."
else
    ok "MariaDB sudah pernah diinisialisasi."
fi

# Start mysqld di background jika belum jalan
if ! pgrep -x mysqld > /dev/null; then
    log "Menjalankan mysqld..."
    mysqld_safe --datadir="$PREFIX/var/lib/mysql" &>/dev/null &
    sleep 5
    ok "mysqld berjalan."
else
    ok "mysqld sudah berjalan."
fi

# Buat database & tabel tve8.produk jika belum ada
log "Membuat database tve8 dan tabel produk (jika belum ada)..."
mysql -u root --connect-timeout=10 <<'SQL'
CREATE DATABASE IF NOT EXISTS tve8;
USE tve8;
CREATE TABLE IF NOT EXISTS produk (
    id  INT AUTO_INCREMENT PRIMARY KEY,
    plu VARCHAR(8) NOT NULL UNIQUE
);
SQL
ok "Database tve8 dan tabel produk siap."

# ------------------------------------------------------------
# 6. Salin v9.py ke folder kerja (jika ada di folder saat ini)
# ------------------------------------------------------------
if [ -f "v9.py" ]; then
    log "Menyalin v9.py ke $WORKDIR..."
    cp v9.py "$WORKDIR/v9.py"
    ok "v9.py disalin ke $WORKDIR."
else
    warn "v9.py tidak ditemukan di folder ini. Salin manual ke: $WORKDIR"
fi

# ------------------------------------------------------------
# 7. Verifikasi instalasi
# ------------------------------------------------------------
echo ""
log "Verifikasi instalasi..."

python - <<'PYCHECK'
errors = []
mods = [
    ("pytesseract", "pytesseract"),
    ("PIL",         "Pillow"),
    ("openpyxl",    "openpyxl"),
    ("pdf2image",   "pdf2image"),
    ("mysql.connector", "mysql-connector-python"),
]
for mod, pkg in mods:
    try:
        __import__(mod)
        print(f"  \033[0;32m✓\033[0m {pkg}")
    except ImportError:
        print(f"  \033[0;31m✗\033[0m {pkg} — GAGAL")
        errors.append(pkg)

if errors:
    print(f"\n\033[0;31m[ERR] Package berikut gagal: {', '.join(errors)}\033[0m")
    exit(1)
PYCHECK

ok "Semua Python package OK."

# ------------------------------------------------------------
# Selesai
# ------------------------------------------------------------
echo ""
echo -e "${GREEN}============================================================${NC}"
echo -e "${GREEN}   Setup selesai!                                          ${NC}"
echo -e "${GREEN}============================================================${NC}"
echo ""
echo -e "  Folder kerja  : ${CYAN}$WORKDIR${NC}"
echo -e "  Cara pakai    :"
echo -e "    ${YELLOW}cd $WORKDIR${NC}"
echo -e "    ${YELLOW}python v9.py file.pdf${NC}"
echo -e "    ${YELLOW}python v9.py *.pdf --no-db   # tanpa MySQL${NC}"
echo ""
echo -e "  ${YELLOW}Catatan:${NC} mysqld perlu dijalankan ulang setiap sesi baru:"
echo -e "    ${YELLOW}mysqld_safe &${NC}   (tunggu 5 detik)"
echo ""
