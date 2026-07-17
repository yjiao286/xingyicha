#!/usr/bin/env bash
# ── 星易查 便携 Linux 包构建脚本 ──
# 在任意平台（含 macOS）运行，产出离线可运行的 Linux 应用目录 + tar.gz。
# 用法: ./build/package.sh [x64|arm64]
set -euo pipefail

ARCH="${1:-x64}"
PY_VERSION="3.11.15"
PBS_TAG="20260623"
APP_NAME="xingyicha-linux-${ARCH}"
VERSION="1.0"

case "$ARCH" in
  x64)   PY_TRIPLET="x86_64-unknown-linux-gnu";  PIP_PLAT="manylinux_2_28_x86_64";  DEB_ARCH="amd64" ;;
  arm64) PY_TRIPLET="aarch64-unknown-linux-gnu"; PIP_PLAT="manylinux_2_28_aarch64"; DEB_ARCH="arm64" ;;
  *) echo "用法: $0 [x64|arm64]"; exit 1 ;;
esac

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
BUILD="$ROOT/build"
CACHE="$BUILD/cache/$ARCH"
DIST="$ROOT/dist"
PKG="$DIST/$APP_NAME"

PY_URL="https://github.com/astral-sh/python-build-standalone/releases/download/${PBS_TAG}/cpython-${PY_VERSION}%2B${PBS_TAG}-${PY_TRIPLET}-install_only.tar.gz"
PY_TARBALL="$CACHE/cpython-${ARCH}.tar.gz"
WHEELHOUSE="$CACHE/wheelhouse"
SITE_DIR_REL="python/lib/python3.11/site-packages"

log() { printf '\n\033[1;36m▶ %s\033[0m\n' "$*"; }
ok()  { printf '  \033[32m✓\033[0m %s\n' "$*"; }

mkdir -p "$CACHE" "$DIST"

# 若目录已存在则重建
rm -rf "$PKG"
mkdir -p "$PKG" "$PKG/app" "$PKG/bin"

# ── 1. 可重定位 CPython ──────────────────────────────────────────
log "下载可重定位 CPython ${PY_VERSION} (${ARCH})"
if [ ! -f "$PY_TARBALL" ]; then
  curl -fL --http1.1 -C - --retry 8 --retry-delay 2 \
    --speed-time 30 --speed-limit 5000 -o "$PY_TARBALL" "$PY_URL"
fi
ok "已缓存: $(basename "$PY_TARBALL")"
# install_only 包解压后顶层目录名为 python/
log "解压 Python"
tmp_extract="$(mktemp -d)"
tar -xzf "$PY_TARBALL" -C "$tmp_extract"
# 找到顶层目录并移动到 PKG/python
py_top=""
for d in "$tmp_extract"/*/; do
  if [ -x "${d}bin/python3" ]; then py_top="${d%/}"; break; fi
done
if [ -z "$py_top" ]; then echo "✗ 未在压缩包中找到 bin/python3"; exit 1; fi
rm -rf "$PKG/python"
mv "$py_top" "$PKG/python"
rm -rf "$tmp_extract"
ok "Python 就位: $PKG/python/bin/python3"

# ── 2. 下载 Linux 预编译 wheel ───────────────────────────────────
log "下载 Linux 依赖 wheel (${PIP_PLAT})"
mkdir -p "$WHEELHOUSE"
# 仅在为空或缺少时重新下载（wheelhouse 非空即复用）
if [ -z "$(ls -A "$WHEELHOUSE" 2>/dev/null)" ]; then
  python3 -m pip download \
    --platform "$PIP_PLAT" \
    --python-version 311 \
    --abi cp311 \
    --implementation cp \
    --only-binary :all: \
    --dest "$WHEELHOUSE" \
    -r "$ROOT/requirements.txt"
fi
ok "wheelhouse: $(ls "$WHEELHOUSE" | wc -l | tr -d ' ') 个 wheel"

# ── 3. 解压 wheel 到 site-packages ────────────────────────────────
log "安装依赖到便携 Python 的 site-packages"
SITE="$PKG/$SITE_DIR_REL"
mkdir -p "$SITE"
for w in "$WHEELHOUSE"/*.whl; do
  python3 - "$w" "$SITE" <<'PYEOF'
import sys, zipfile
zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])
PYEOF
done
ok "已展开到: $SITE"

# ── 4. 应用代码 ──────────────────────────────────────────────────
log "拷贝应用代码"
cp "$ROOT/app.py" "$PKG/app/app.py"
cp -R "$ROOT/templates" "$PKG/app/templates"
cp -R "$ROOT/static" "$PKG/app/static"
ok "app.py + templates + static"

# ── 5. antiword（.doc 支持，尽力而为）────────────────────────────
log "获取 antiword (${DEB_ARCH}) for .doc 支持"
if python3 "$ROOT/build/_extract_deb.py" "$DEB_ARCH" "$PKG/bin/antiword" "$CACHE"; then
  chmod +x "$PKG/bin/antiword"
  ok "antiword 就位 (.doc 可用)"
else
  echo "  ⚠ 未能获取 antiword：.doc 不可用（.docx/.pdf 正常）"
fi

# ── 6. 运行时文件 ─────────────────────────────────────────────────
log "写入启动器 / 桌面项 / 自检 / 说明"
cp "$ROOT/build/templates/run.sh" "$PKG/run.sh"
cp "$ROOT/build/templates/install.sh" "$PKG/install.sh"
cp "$ROOT/build/templates/xingyicha.desktop" "$PKG/xingyicha.desktop"
cp "$ROOT/build/templates/verify.sh" "$PKG/verify.sh"
cp "$ROOT/build/templates/README.txt" "$PKG/README.txt"
# 全部设可执行（含 install.sh 与 .desktop，避免麒麟双击进编辑器）
chmod +x "$PKG/run.sh" "$PKG/install.sh" "$PKG/verify.sh" "$PKG/xingyicha.desktop"
# 去除 macOS 噪声文件
find "$PKG" -name '.DS_Store' -delete 2>/dev/null || true
ok "运行时文件已就位"

# ── 7. 静态校验 ──────────────────────────────────────────────────
log "静态校验：关键 import 是否齐备"
python3 - "$PKG/app/app.py" "$SITE" "$PKG/python/lib/python3.11" <<'PYEOF'
import sys, os, re
app_src = open(sys.argv[1], encoding='utf-8').read()
site = sys.argv[2]
stdlib = sys.argv[3]

available = set()
BUILTINS = {"sys", "builtins", "_thread", "_io", "posix", "nt", "_warnings",
            "_signal", "_weakref", "_abc", "_collections_abc", "codecs",
            "_frozen_importlib", "_frozen_importlib_external", "time", "math", "_json", "array", "binascii", "fcntl", "select", "unicodedata", "zlib", "ssl", "hashlib"}

# vendored 第三方（site-packages）
if os.path.isdir(site):
    for d in os.listdir(site):
        full = os.path.join(site, d)
        if d.endswith('.dist-info'):
            tl = os.path.join(full, 'top_level.txt')
            if os.path.exists(tl):
                for line in open(tl, encoding='utf-8', errors='replace'):
                    line = line.strip()
                    if line:
                        available.add(line.lower())
        elif os.path.isdir(full) or d.endswith('.so'):
            available.add(d.split('.')[0].lower())

# 便携 Python 自带标准库（lib/python3.11/）
if os.path.isdir(stdlib):
    for d in os.listdir(stdlib):
        full = os.path.join(stdlib, d)
        if d.endswith('.py'):
            available.add(d[:-3].lower())
        elif os.path.isdir(full):
            available.add(d.lower())
        elif d.endswith('.so'):
            available.add(d.split('.')[0].lower())
    # C 扩展标准库模块
    dynload = os.path.join(stdlib, 'lib-dynload')
    if os.path.isdir(dynload):
        for d in os.listdir(dynload):
            if d.endswith('.so'):
                available.add(d.split('.')[0].lower())
available |= BUILTINS

imports = re.findall(r'^\s*(?:import|from)\s+([a-zA-Z0-9_]+)', app_src, re.M)
missing = [p for p in sorted(set(imports))
           if p.lower() not in available and p.lower().replace('_', '-') not in available]
if missing:
    print("  ✗ 缺失 import:", ", ".join(missing))
    sys.exit(1)
print("  ✓ 全部顶层 import 可解析 (stdlib + vendored 第三方)")
PYEOF

# 校验 wheel 平台标签均为 linux（pure-python 为 py3-none-any 也通过）
bad=0
for w in "$WHEELHOUSE"/*.whl; do
  bn="$(basename "$w")"
  case "$bn" in
    *linux*|*py3-none-any*|*cp311-none-any*) : ;;
    *) echo "  ⚠ 非 Linux wheel: $bn"; bad=1 ;;
  esac
done
[ $bad -eq 0 ] && ok "所有 wheel 平台标签兼容 Linux"

# ── 8. 打包 tar.gz ───────────────────────────────────────────────
log "打包 ${APP_NAME}-${VERSION}.tar.gz"
tar -czf "$DIST/${APP_NAME}-${VERSION}.tar.gz" -C "$DIST" "$APP_NAME"
ok "产物: dist/${APP_NAME}-${VERSION}.tar.gz ($(du -h "$DIST/${APP_NAME}-${VERSION}.tar.gz" | cut -f1))"

echo ""
echo "完成。解压后运行 ./run.sh 即可（目标机无需联网）。"
