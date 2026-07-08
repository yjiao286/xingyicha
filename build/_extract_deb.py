#!/usr/bin/env python3
"""从 Debian .deb 中抽取 antiword 二进制（跨平台，无需 dpkg）。

用法: _extract_deb.py <deb_arch> <out_bin> <cache_dir>
  deb_arch : amd64 | arm64
  out_bin  : 输出 antiword 文件路径
  cache_dir: .deb 缓存目录

成功抽出则退出 0；失败退出非 0（构建脚本据此降级）。
"""
import io
import os
import re
import sys
import tarfile
import urllib.request

DEB_ARCH = sys.argv[1]
OUT_BIN = sys.argv[2]
CACHE = sys.argv[3]
os.makedirs(CACHE, exist_ok=True)

POOL = "http://ftp.debian.org/debian/pool/main/a/antiword/"


def http_get(url, dest=None, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": "xingyicha-build/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read()
    if dest:
        with open(dest, "wb") as f:
            f.write(data)
    return data


def find_deb():
    listing = http_get(POOL).decode("latin1", "replace")
    pat = r'href="(antiword_[^"]*_{}\.deb)"'.format(re.escape(DEB_ARCH))
    names = re.findall(pat, listing)
    if not names:
        return None
    # 取文件名最长的（通常是最高版本）
    names.sort()
    return names[-1]


def parse_ar(data):
    """简易 AR 归档解析 → {name: bytes}"""
    assert data[:8] == b"!<arch>\n", "not an AR archive"
    members = {}
    pos = 8
    while pos + 60 <= len(data):
        hdr = data[pos:pos + 60]
        name = hdr[0:16].decode("latin1").strip().rstrip("/").rstrip()
        size = int(hdr[48:58].decode("latin1").strip())
        pos += 60
        body = data[pos:pos + size]
        # GNU ar long-name (#1/XX) 在此包中不会出现（文件名都很短）
        if name and not name.isdigit():
            members[name] = body
        pos += size + (size & 1)  # 偶对齐填充
    return members


def main():
    deb_name = find_deb()
    if not deb_name:
        print("  ✗ 在 Debian pool 未找到 antiword _{}_ 的 .deb".format(DEB_ARCH), file=sys.stderr)
        return 1
    deb_path = os.path.join(CACHE, deb_name)
    if not os.path.exists(deb_path):
        print("  → 下载 {}".format(deb_name))
        http_get(POOL + deb_name, deb_path)
    else:
        print("  → 复用缓存 {}".format(deb_name))
    with open(deb_path, "rb") as f:
        members = parse_ar(f.read())
    # 找 data.tar.*
    data_key = next((k for k in members if k.startswith("data.tar")), None)
    if not data_key:
        print("  ✗ .deb 中无 data.tar.*", file=sys.stderr)
        return 1
    raw = members[data_key]
    tf = None
    last_err = None
    # 依次尝试透明解压（含 Python 3.14 的 zstd）与显式模式
    for mode in ("r", "r:*", "r:gz", "r:xz", "r:zstd", "r:bz2"):
        try:
            tf = tarfile.open(fileobj=io.BytesIO(raw), mode=mode)
            break
        except Exception as e:  # noqa: BLE001
            last_err = e
            tf = None
    if tf is None:
        print("  ✗ 无法读取 {}: {}".format(data_key, last_err), file=sys.stderr)
        return 1
    for m in tf.getmembers():
        if m.isfile() and m.name.rstrip("/").endswith("bin/antiword"):
            data = tf.extractfile(m).read()
            with open(OUT_BIN, "wb") as f:
                f.write(data)
            print("  ✓ 抽取 antiword ({}, {} 字节)".format(m.name, len(data)))
            return 0
    print("  ✗ data.tar 中未找到 antiword 二进制", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
