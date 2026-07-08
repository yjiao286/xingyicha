#!/usr/bin/env bash
# ── 星易查 桌面快捷方式安装/启动器 (银河麒麟等) ──
# 双击本脚本即可启动应用；首次运行会把应用快捷方式安装到开始菜单/桌面。
set -u
HERE="$(cd "$(dirname "$0")" && pwd)"
RUN="$HERE/run.sh"
DESKTOP="$HOME/.local/share/applications/xingyicha.desktop"
ICON_DIR="$HOME/.local/share/icons/hicolor/256x256/apps"
ICON_PNG="$ICON_DIR/xingyicha.png"

[ -x "$RUN" ] || chmod +x "$RUN" 2>/dev/null

# 生成内置图标（简单的星形 SVG → 不依赖任何外部图片文件）
mkdir -p "$ICON_DIR" 2>/dev/null
# 用 python 把一段极简 PNG 写出来更稳；若失败则退化为 text-editor 图标
"$HERE/python/bin/python3" - "$ICON_PNG" <<'PYEOF' 2>/dev/null || true
import sys, base64, zlib, struct
# 32x32 纯色方块（星易查蓝 #2563eb）作为图标占位，保证桌面能显示图标
def png_bytes():
    W=H=32
    rgba=b'\x25\x63\xeb\xff'*W
    raw=b''.join(b'\x00'+rgba for _ in range(H))
    def chunk(t,d):
        c=t+d
        return struct.pack('>I',len(d))+c+struct.pack('>I',zlib.crc32(c)&0xffffffff)
    sig=b'\x89PNG\r\n\x1a\n'
    ihdr=struct.pack('>IIBBBBB',W,H,8,6,0,0,0)
    return sig+chunk(b'IHDR',ihdr)+chunk(b'IDAT',zlib.compress(raw))+chunk(b'IEND',b'')
open(sys.argv[1],'wb').write(png_bytes())
PYEOF

mkdir -p "$(dirname "$DESKTOP")" 2>/dev/null
cat > "$DESKTOP" <<EOF
[Desktop Entry]
Type=Application
Version=1.0
Name=星易查
Name[zh_CN]=星易查
GenericName=围串标风险识别系统
Comment=多维度分析投标文件，检测围标串标行为
Exec=bash -lc 'cd "$HERE" && ./run.sh'
Path=$HERE
Icon=$([ -f "$ICON_PNG" ] && echo "$ICON_PNG" || echo "text-editor")
Terminal=true
Categories=Office;Utility;
StartupNotify=true
EOF
chmod +x "$DESKTOP" 2>/dev/null

# 同步到桌面（若有标准桌面目录）
DESKTOP_DIR=$(xdg-user-dir DESKTOP 2>/dev/null || echo "$HOME/Desktop")
if [ -d "$DESKTOP_DIR" ]; then
  cp "$DESKTOP" "$DESKTOP_DIR/xingyicha.desktop" 2>/dev/null && chmod +x "$DESKTOP_DIR/xingyicha.desktop" 2>/dev/null
fi

# 尝试刷新桌面图标缓存（非关键）
update-desktop-database "$HOME/.local/share/applications" 2>/dev/null || true
gtk-update-icon-cache -f "$ICON_DIR" 2>/dev/null || true

echo "✅ 已创建桌面快捷方式。"
echo "  - 开始菜单：星易查"
[ -d "$DESKTOP_DIR" ] && echo "  - 桌面：xingyicha.desktop（首次右键→“允许启动”/"信任"）"
echo ""
echo "正在启动应用..."
exec bash "$RUN"
