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
    tools/check_sensitive.py --revs A..B       # explicit revision range
    tools/check_sensitive.py --pre-push        # pre-push (ref list on stdin)
Exit status 1 when something matches, 2 on tool failure (fail-closed).

Portability notes (all learned the hard way):
  * path listings run with core.quotepath=false and -z — with the default
    quotepath=true, non-ASCII (Chinese) paths come back as quoted octal
    escapes, the follow-up git show/open() fails, and the file silently
    scanned as empty text with exit 0 (the worst possible outcome for the
    files most likely to carry Chinese PII).
  * git outputs are decoded as UTF-8 explicitly — bare `text=True` picks the
    locale codec (cp936 on Chinese Windows), which either mojibakes the
    commit message beyond term matching or raises UnicodeDecodeError.
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
        r = subprocess.run(['git', 'rev-parse', '--git-dir'],
                           capture_output=True, text=True,
                           encoding='utf-8', errors='replace')
        git_dir = r.stdout.strip()
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
ALLOWED_EMAILS = {'a@b.com', 'zhang@example.com', 'jiaoy1@chinasatnet.com.cn',
                  'git@github.com'}   # SSH clone URL, not a person's mailbox

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
    """(path, text) for every staged text file.

    Paths are listed with core.quotepath=false and NUL separation: with the
    default quotepath=true, non-ASCII (Chinese) paths come back as quoted
    octal escapes, `git show :path` then fails, and the file used to be
    silently skipped with a clean exit 0 — exactly the files most likely to
    carry Chinese PII were never scanned. A failed read now fails CLOSED.
    """
    r = subprocess.run(
        ['git', '-c', 'core.quotepath=false', 'diff', '--cached',
         '--name-only', '--diff-filter=ACMR', '-z'],
        capture_output=True)
    if r.returncode != 0:
        print('check_sensitive: git diff 失败：'
              + r.stderr.decode('utf-8', 'replace').strip(), file=sys.stderr)
        sys.exit(2)
    names = r.stdout.decode('utf-8', 'replace').split('\0')
    for path in names:
        path = path.strip()
        if not path or path.lower().endswith(BINARY_SUFFIXES):
            continue
        show = subprocess.run(['git', 'show', ':' + path],
                              capture_output=True)
        if show.returncode != 0:
            # fail-closed: 不可读的已暂存文件必须拦截，不能当作空文本放行
            print(f'check_sensitive: 无法读取暂存文件 {path!r}（git show 失败），'
                  f'为避免漏扫本次按失败处理。', file=sys.stderr)
            sys.exit(2)
        try:
            yield path, show.stdout.decode('utf-8')
        except UnicodeDecodeError:
            continue    # not text; nothing a reviewer could read anyway


def tracked_texts():
    # -c core.quotepath=false -z：中文路径原样、按 NUL 分割（见 staged_blobs 注释）
    r = subprocess.run(
        ['git', '-c', 'core.quotepath=false', 'ls-files', '-z'],
        capture_output=True)
    if r.returncode != 0:
        print('check_sensitive: git ls-files 失败：'
              + r.stderr.decode('utf-8', 'replace').strip(), file=sys.stderr)
        sys.exit(2)
    names = r.stdout.decode('utf-8', 'replace').split('\0')
    for path in names:
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
    # 先切到仓库根目录：在子目录里 `git ls-files` 只列当前子树，
    # --all 会静默地只审计仓库的一个切片。
    try:
        top = subprocess.run(['git', 'rev-parse', '--show-toplevel'],
                             capture_output=True, text=True,
                             encoding='utf-8', errors='replace').stdout.strip()
        if top:
            os.chdir(top)
    except OSError:
        pass
    pairs = []
    if mode == '--staged':
        pairs = list(staged_blobs())
    elif mode == '--all':
        pairs = list(tracked_texts())
    elif mode == '--revs':
        # pre-push: everything about to be published. Covers the gap that
        # `git commit --no-verify` leaves open — the commit hooks are advisory
        # to anyone who knows the flag, and by push time the only way back is
        # rewriting published history.
        revs = argv[2:]
        if revs:
            out = subprocess.run(
                ['git', 'log', '--format=%H%x01%s%x01%b%x02'] + revs,
                capture_output=True, text=True,
                encoding='utf-8', errors='replace').stdout
            for rec in out.split('\x02'):
                if not rec.strip():
                    continue
                parts = rec.split('\x01')
                if len(parts) < 3:
                    continue
                pairs.append((f'commit {parts[0][:9]} message',
                              parts[1] + '\n' + parts[2]))
            names = subprocess.run(
                ['git', 'diff', '--name-only', '--diff-filter=ACMR'] + revs,
                capture_output=True, text=True,
                encoding='utf-8', errors='replace').stdout.split('\n')
            for path in names:
                path = path.strip()
                if not path or path.lower().endswith(BINARY_SUFFIXES):
                    continue
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        pairs.append((path, f.read()))
                except (OSError, UnicodeDecodeError):
                    continue
    elif mode == '--pre-push':
        # pre-push 钩子专用：git 把待推送引用以
        # '<local_ref> <local_oid> <remote_ref> <remote_oid>' 逐行写到 stdin
        # （不是 argv —— 从 argv 取 revision 的钩子在这里永远拿不到参数，
        # 会形成“看似有防护实则恒放行”的假控制）。OID 只认完整十六进制，
        # 拼出的 revision 再交给 --revs 的同一套扫描。
        for line in sys.stdin.read().splitlines():
            fields = line.split()
            if len(fields) < 4:
                continue
            loid, roid = fields[1], fields[3]
            _hex = re.compile(r'[0-9a-f]{40}(?:[0-9a-f]{24})?$')
            if not (_hex.fullmatch(loid) and _hex.fullmatch(roid)):
                continue
            if set(loid) == {'0'}:
                continue            # 删除远端引用：没有新内容
            range_revs = ([loid, '--not', '--remotes']
                          if set(roid) == {'0'}
                          else [roid + '..' + loid])
            out = subprocess.run(
                ['git', 'log', '--format=%H%x01%s%x01%b%x02'] + range_revs,
                capture_output=True, text=True,
                encoding='utf-8', errors='replace').stdout
            for rec in out.split('\x02'):
                if not rec.strip():
                    continue
                parts = rec.split('\x01')
                if len(parts) < 3:
                    continue
                pairs.append((f'commit {parts[0][:9]} message',
                              parts[1] + '\n' + parts[2]))
            diff = subprocess.run(
                ['git', 'diff', '--unified=0', '--diff-filter=ACMR'] + range_revs,
                capture_output=True, text=True,
                encoding='utf-8', errors='replace').stdout
            # 只扫新增行（--unified=0 已去掉上下文）：范围内的既有内容
            # 早已公开，新增行才是即将发布的东西。
            added = [ln[1:] for ln in diff.splitlines()
                     if ln.startswith('+') and not ln.startswith('+++')]
            if added:
                pairs.append(('<待推送范围新增行>', '\n'.join(added)))
    elif mode == '--message':
        if len(argv) < 3:
            return 0
        # git 提交信息是 UTF-8，但中文 Windows 的编辑器仍可能写 GBK；
        # 硬解码失败会带裸 traceback 阻塞每一次合法提交。
        with open(argv[2], 'rb') as f:
            raw = f.read()
        text = None
        for enc in ('utf-8', 'gb18030'):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        pairs = [('<commit message>',
                  text if text is not None
                  else raw.decode('utf-8', errors='replace'))]
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
