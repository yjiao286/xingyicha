#!/usr/bin/env python3
"""Reject commits that carry real personal / client / competition data.

Why this exists: on 2026-09-16 one working session leaked, twice, into this
repo — once as real person and company names written into comments, docstrings
and test fixtures, once as a company registration number inside a commit
MESSAGE describing a finding. Both were spotted after the fact and needed
history rewritten. Reviewing by eye is not a control; it does not survive a
busy afternoon. This is the mechanical version.

Two layers, deliberately different in kind:

  * TERMS — exact substrings: names, companies, project and competition
    vocabulary that came out of the analysis corpus, plus every concrete
    value this repo has previously leaked. Zero false positives, but only
    catches what is already known to be sensitive.

  * SHAPES — phone numbers, ID numbers, credit codes and street addresses
    caught by pattern. This is what catches a leak nobody has seen before, so
    it has to tolerate the synthetic fixtures the tests are built on: those
    are enumerated in ALLOWED_* below.

Usage:
    tools/check_sensitive.py --staged          # pre-commit
    tools/check_sensitive.py --message FILE    # commit-msg
    tools/check_sensitive.py --all             # audit the whole worktree
Exit status 1 when something matches.
"""
import os
import re
import subprocess
import sys

# ── Exact terms ──────────────────────────────────────────────────────────
# The term list is NOT versioned, and that is deliberate: a file listing the
# real names and client vocabulary it is meant to catch is itself the leak.
# It lives in .git/sensitive-terms.txt, which git never tracks. A fresh clone
# starts with only the SHAPES below, which need no term list at all.
#
# Format: one term per line, '#' starts a comment. See
# tools/sensitive-terms.example.txt.
def _load_terms():
    try:
        git_dir = subprocess.run(['git', 'rev-parse', '--git-dir'],
                                 capture_output=True, text=True).stdout.strip()
    except OSError:
        return []
    if not git_dir:
        return []
    path = os.path.join(git_dir, 'sensitive-terms.txt')
    try:
        with open(path, 'r', encoding='utf-8') as f:
            lines = f.read().splitlines()
    except OSError:
        return []
    return [ln.strip() for ln in lines
            if ln.strip() and not ln.lstrip().startswith('#')]


BLOCKED_TERMS = _load_terms()

# ── Value shapes ─────────────────────────────────────────────────────────
# Values the test suite legitimately uses. Anything shaped like these but NOT
# listed here is treated as real. Keep this list tight: an over-wide allowlist
# is a hole, an over-narrow one makes the hook unbearable and it gets --no-verify'd.
ALLOWED_PHONES = {
    '13800000000', '13800000001', '13800138000', '13801234567',
    '13900000001', '13900000002', '13900000003', '13900000004',
    '13900000009', '13911111111', '13911112222', '13912345678',
}
ALLOWED_IDS = {
    '110101199001011234', '110101199003070000',
    '320101199001011234', '32012319900101123X',
}
ALLOWED_LANDLINES = {
    '010-12345678', '010-77775555', '010-88886666',
    '021-12345678', '0755-12345678',
}
ALLOWED_ADDRESSES = {
    '北京市海淀区某路某号', '北京市海淀区中关村大街1号',
    '北京市海淀区上地十街10号', '北京市朝阳区建国路88号院6号',
}
ALLOWED_EMAILS = {'a@b.com', 'zhang@example.com', 'jiaoy1@chinasatnet.com.cn'}

SHAPES = [
    ('手机号', re.compile(r'(?<!\d)1[3-9]\d{9}(?!\d)'), ALLOWED_PHONES),
    ('座机号', re.compile(r'(?<![\d-])0\d{2,3}-\d{7,8}(?![\d-])'), ALLOWED_LANDLINES),
    ('身份证号', re.compile(r'(?<![\dXx])\d{17}[\dXx](?![\dXx])'), ALLOWED_IDS),
    # 统一社会信用代码 — 18 chars, starts with 9. No synthetic ones exist here.
    ('统一社会信用代码', re.compile(r'(?<![0-9A-Z])9[0-9A-Z]{17}(?![0-9A-Z])'), set()),
    ('住址',
     re.compile(r'[一-鿿]{2,}(?:市|区|县)[一-鿿0-9]{2,30}号'), ALLOWED_ADDRESSES),
    ('邮箱',
     re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'), ALLOWED_EMAILS),
]

BINARY_SUFFIXES = ('.png', '.jpg', '.jpeg', '.gif', '.ico', '.icns', '.dylib',
                   '.so', '.exe', '.zip', '.gz', '.dmg', '.pdf', '.docx',
                   '.xlsx', '.pyc', '.onnx')


def _iter_terms(text):
    for term in BLOCKED_TERMS:
        if term in text:
            yield term


def _iter_shapes(text):
    for label, pattern, allowed in SHAPES:
        for m in pattern.finditer(text):
            value = m.group(0)
            if value not in allowed:
                yield label, value


def _lines_with(text, needle, limit=3):
    hits = []
    for n, line in enumerate(text.splitlines(), 1):
        if needle in line:
            hits.append((n, line.strip()[:160]))
            if len(hits) >= limit:
                break
    return hits


def check_text(label, text, report):
    for term in _iter_terms(text):
        report.append((label, '敏感词', term, _lines_with(text, term)))
    for kind, value in _iter_shapes(text):
        report.append((label, kind, value, _lines_with(text, value)))


def staged_blobs():
    """(path, text) for every staged text file."""
    out = subprocess.run(
        ['git', 'diff', '--cached', '--name-only', '--diff-filter=ACMR'],
        capture_output=True, text=True).stdout.split('\n')
    for path in out:
        path = path.strip()
        if not path or path.lower().endswith(BINARY_SUFFIXES):
            continue
        blob = subprocess.run(['git', 'show', f':{path}'],
                              capture_output=True).stdout
        try:
            yield path, blob.decode('utf-8')
        except UnicodeDecodeError:
            continue    # not text; nothing a reviewer could read anyway


def tracked_texts():
    out = subprocess.run(['git', 'ls-files'], capture_output=True,
                         text=True).stdout.split('\n')
    for path in out:
        path = path.strip()
        if not path or path.lower().endswith(BINARY_SUFFIXES):
            continue
        try:
            with open(path, 'r', encoding='utf-8') as f:
                yield path, f.read()
        except (OSError, UnicodeDecodeError):
            continue


def main(argv):
    mode = argv[1] if len(argv) > 1 else '--staged'
    pairs = []
    if mode == '--staged':
        pairs = list(staged_blobs())
    elif mode == '--all':
        pairs = list(tracked_texts())
    elif mode == '--message':
        if len(argv) < 3:
            return 0
        with open(argv[2], 'r', encoding='utf-8') as f:
            pairs = [('<commit message>', f.read())]
    else:
        print(f'unknown mode: {mode}', file=sys.stderr)
        return 2

    report = []
    for label, text in pairs:
        check_text(label, text, report)

    if not report:
        return 0

    print('\n提交被拒绝：检测到疑似真实主体 / 赛事 / 个人信息。\n', file=sys.stderr)
    for label, kind, value, lines in report:
        print(f'  [{kind}] {value}   ← {label}', file=sys.stderr)
        for n, line in lines:
            print(f'        {n}: {line}', file=sys.stderr)
    print('\n请改为占位值（甲公司/乙公司、张三/李四、明显合成的号码）。'
          '\n若确属误报，把该值加入 tools/check_sensitive.py 的 ALLOWED_* 列表。'
          '\n确需绕过：git commit --no-verify\n', file=sys.stderr)
    return 1


if __name__ == '__main__':
    sys.exit(main(sys.argv))
