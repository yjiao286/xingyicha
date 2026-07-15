#!/usr/bin/env bash
# ── 星易查 离线自检 ──
# 在目标 Linux 机器上运行，验证便携包完整性（无需联网）。
set -u
HERE="$(cd "$(dirname "$0")" && pwd)"
PY="$HERE/python/bin/python3"

if [ ! -x "$PY" ]; then echo "✗ 找不到内置 Python: $PY"; exit 1; fi

export PYTHONPATH="$HERE/python/lib/python3.11/site-packages${PYTHONPATH:+:$PYTHONPATH}"

echo "① 导入关键依赖..."
"$PY" -c "import flask,docx,pypdf,lxml,werkzeug,jinja2,olefile; print('   OK: flask / python-docx / pypdf / lxml / olefile ...')" \
  || { echo "   ✗ 依赖导入失败"; exit 1; }

echo "② app 模块自检..."
"$PY" "$HERE/app/app.py" --check \
  || { echo "   ✗ app 自检失败"; exit 1; }

echo "③ 外部工具 (.doc 支持)..."
if [ -x "$HERE/bin/antiword" ]; then
  echo "   OK: antiword 已就位 → .doc 可解析"
else
  echo "   ⚠ antiword 缺失 → .doc 不可用（.docx / .pdf 正常）"
fi

echo ""
echo "✅ 自检通过。运行 ./run.sh 启动并打开浏览器。"
