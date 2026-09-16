#!/bin/sh
# Install the sensitive-data guards into .git/hooks.
#
# Hooks live in .git/, which is not versioned, so this script is what makes
# them reproducible: run it once per clone. See tools/check_sensitive.py for
# what the guards actually reject and why they exist.
set -e

root=$(git rev-parse --show-toplevel)
hooks="$root/.git/hooks"
mkdir -p "$hooks"

write_hook() {
    name="$1"; shift
    cat > "$hooks/$name" <<EOF
#!/bin/sh
# Installed by tools/install-hooks.sh — edit tools/check_sensitive.py instead.
root=\$(git rev-parse --show-toplevel)
if [ -x "\$root/venv/bin/python3" ]; then
    py="\$root/venv/bin/python3"
elif command -v python3 >/dev/null 2>&1; then
    py=python3
else
    echo "check_sensitive: 未找到 python3，跳过敏感信息检查" >&2
    exit 0
fi
exec "\$py" "\$root/tools/check_sensitive.py" $*
EOF
    chmod +x "$hooks/$name"
    echo "  installed .git/hooks/$name"
}

write_hook pre-commit --staged
write_hook commit-msg --message '"$1"'

echo "完成。手动全量自检：tools/check_sensitive.py --all"
