import os
import sys
import re
import json
import base64
import zipfile
import subprocess
import tempfile
import hashlib
import shutil
import logging
import uuid
import time
from io import BytesIO
from pathlib import Path
from datetime import datetime
from collections import defaultdict

from flask import Flask, request, jsonify, send_file, render_template, Response
# werkzeug is imported for Response/send_file; filename sanitization uses
# custom _sanitize_filename() which preserves CJK non-ASCII chars that
# werkzeug.utils.secure_filename would strip.
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from pypdf import PdfReader
import olefile
import threading

# ── Frozen (PyInstaller) detection ──────────────────────────────
# When bundled as a desktop exe, templates/static live inside the bundle
# (_MEIPASS) and user data must NOT be written next to the exe (Program Files
# is read-only) — it goes to %LOCALAPPDATA%\星易查 instead. Non-frozen
# behaviour is unchanged.
IS_FROZEN = bool(getattr(sys, 'frozen', False))
if IS_FROZEN:
    # PyInstaller console exe: stdout/stderr may default to the legacy ANSI
    # codepage (cp1252 etc.), where printing Chinese crashes with
    # UnicodeEncodeError. Force UTF-8 with replacement chars instead.
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
if IS_FROZEN:
    _BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
    # _MEIPASS: onedir -> .../星易查/_internal, onefile -> temp extraction dir
    _RESOURCE_DIR = getattr(sys, '_MEIPASS', _BASE_DIR)
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    _RESOURCE_DIR = _BASE_DIR

app = Flask(
    __name__,
    template_folder=os.path.join(_RESOURCE_DIR, 'templates'),
    static_folder=os.path.join(_RESOURCE_DIR, 'static'),
)
# Secret key from environment (fallback to a per-process random key so a missing
# env var never leaves the session signer predictable). Never commit a real key.
app.secret_key = os.environ.get('SECRET_KEY') or os.urandom(24)
# Upload body limit: UNLIMITED by default (real-world bids reach 300MB+ and
# total uploads can be several GB). Set MAX_CONTENT_LENGTH_MB to impose a cap
# (e.g. behind a reverse proxy that needs a matching client_max_body_size).
# MAX_FILE_SIZE_MB / MAX_TOTAL_SIZE_MB only warn in the UI, they never block.
_max_body_mb = os.environ.get('MAX_CONTENT_LENGTH_MB')
if _max_body_mb:
    _max_body_mb = int(_max_body_mb)
    app.config['MAX_CONTENT_LENGTH'] = _max_body_mb * 1024 * 1024


@app.errorhandler(413)
def _too_large(e):
    """Return a JSON error for oversized uploads instead of werkzeug's HTML
    page, so the frontend can surface an actionable message."""
    limit = f'{_max_body_mb}MB' if _max_body_mb else '当前限制'
    return jsonify({
        'error': f'上传文件过大：单次上传总大小限制为 {limit}。'
                 f'请压缩文件（如将大PDF拆分或转为文字版），或通过环境变量 '
                 f'MAX_CONTENT_LENGTH_MB 调整上限后重启服务。'
    }), 413

# Structured logging (replaces bare print/traceback usage for diagnostics).
logging.basicConfig(
    level=os.environ.get('LOG_LEVEL', 'INFO'),
    format='%(asctime)s %(levelname)s [%(name)s] %(message)s',
)
logger = logging.getLogger('xingyicha')

# Desktop (frozen) build: keep uploads/history in the per-user data dir so
# they survive reinstalls and work even when the app bundle itself is
# read-only (Program Files, /Applications, AppImage mount point). Env vars
# (UPLOAD_FOLDER / HISTORY_DIR) still take precedence everywhere.
if IS_FROZEN:
    if sys.platform == 'darwin':
        # ~/Library/Application Support/星易查 (standard macOS location)
        _DATA_DIR = os.path.join(os.path.expanduser('~'), 'Library',
                                 'Application Support', '星易查')
    elif os.name == 'nt':
        # %LOCALAPPDATA%\星易查
        _DATA_DIR = os.path.join(os.environ.get('LOCALAPPDATA') or _BASE_DIR, '星易查')
    else:
        # $XDG_DATA_HOME/星易查, fallback ~/.local/share/星易查 (AppImage
        # mount point under /tmp is read-only, so never write next to the exe)
        _DATA_DIR = os.path.join(
            os.environ.get('XDG_DATA_HOME')
            or os.path.join(os.path.expanduser('~'), '.local', 'share'),
            '星易查')
else:
    _DATA_DIR = _BASE_DIR

UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER') or (
    os.path.join(_DATA_DIR, 'uploads') if IS_FROZEN
    else tempfile.mkdtemp(prefix='bid_uploads_')
)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HISTORY_DIR = os.environ.get('HISTORY_DIR') or os.path.join(_DATA_DIR, 'history')
os.makedirs(HISTORY_DIR, exist_ok=True)

# ── Cancellation ─────────────────────────────────────────────────
class AnalysisCancelled(Exception):
    """Raised inside extraction/analysis threads when the user cancels.

    Propagates up through run_full_analysis / extract_text_with_tables to the
    stream generator, which emits a 'cancelled' event instead of an error."""


# request_id -> threading.Event, registered by /api/analyze_stream and set by
# POST /api/cancel. Guarded by _CANCEL_LOCK.
_CANCEL_EVENTS = {}
_CANCEL_LOCK = threading.Lock()


def _check_cancelled(cancel_event):
    if cancel_event is not None and cancel_event.is_set():
        raise AnalysisCancelled()


# ── Config constants ──────────────────────────────────────────────
MAX_FILE_SIZE_MB = 300        # warn if any single file exceeds this
MAX_TOTAL_SIZE_MB = 500       # warn if all files combined exceed this
MAX_PDF_PAGES = 0            # max pages to process per PDF (0=unlimited, default on)
MAX_EMPTY_PAGE_STREAK = 50    # consecutive empty pages → early stop

# Centralized analysis thresholds (overridable via env for industry tuning).
# Most are used inline for performance; gathered here as documentation and
# to support future env-driven configuration.
CONFIG = {
    # ── Text similarity ──
    'seg_min_len': 15,              # minimum match length (chars)
    'seg_min_meaningful': 5,        # minimum CJK+letter+digit count
    'subst_score_high': 0.6,        # >= this → substantial (collusion evidence)
    'subst_score_suspicious': 0.3,  # >= this → suspicious (shown but demoted)
    'subst_len_penalty_threshold': 200,  # chars; shorter segments get a length penalty
    # ── Pricing ──
    'min_price_val': 100,           # ignore price-like values below this
    'min_large_price': 5000,        # higher threshold for global-search prices
    'min_table_total': 50000,       # plausible cross-bidder total floor
    # ── Personnel ──
    'overlap_threshold': 0.5,       # >= this fraction → "high overlap" finding
    # ── Verdict scoring ──
    'verdict_high': 50,
    'verdict_medium': 15,
    'verdict_low': 0,
}

# ── Helpers ─────────────────────────────────────────────────────
def sanitize_text(text):
    """Remove control characters that break JSON serialization"""
    if not isinstance(text, str):
        return text
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)

def _find_tool(*names):
    """Find first available command-line tool.

    Returns the executable's full path (shutil.which resolves names on PATH;
    on Windows LibreOffice is never on PATH so we probe the standard install
    locations - the full path also works in subprocess calls on every
    platform)."""
    candidates = list(names)
    if os.name == 'nt':
        for progdir in (os.environ.get('PROGRAMFILES', r'C:\Program Files'),
                        os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)')):
            candidates.append(os.path.join(progdir, 'LibreOffice', 'program', 'soffice.exe'))
    for name in candidates:
        path = shutil.which(name)
        if path:
            return path
    return None


# ── Match normalization (shared by similarity / template index / frontend) ──
# Characters dropped entirely before comparison: all Unicode whitespace plus
# zero-width / invisible characters that PDF extraction and OCR frequently
# insert (soft hyphen, ZWSP/ZWNJ/ZWJ, BOM). MUST stay in sync with the
# _SKIP_RE regex in static/js/main.js.
_SKIP_EXTRA = '­​‌‍﻿'


def _is_skip_char(ch):
    return ch.isspace() or ch in _SKIP_EXTRA


# 1:1 character folding applied to comparison text (never to displayed text):
# full-width ASCII (！-～) → half-width, CJK punctuation → ASCII counterparts,
# A-Z → a-z. Lets segments that differ only in punctuation style, case or
# full/half-width forms still match and still get recognized as the same
# template content. MUST stay in sync with _foldChar() in static/js/main.js.
_FOLD_MAP = {}
for _o in range(0xFF01, 0xFF5F):          # ！(FF01) .. ～(FF5E) → ! .. ~
    _dst = chr(_o - 0xFEE0)
    if 'A' <= _dst <= 'Z':                # fold full-width letters to lowercase
        _dst = _dst.lower()
    _FOLD_MAP[_o] = _dst
for _c in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
    _FOLD_MAP[ord(_c)] = _c.lower()
for _src, _dst in {
    '，': ',', '。': '.', '、': ',', '；': ';', '：': ':', '？': '?', '！': '!',
    '‘': "'", '’': "'", '‚': "'", '“': '"', '”': '"', '„': '"',
    '（': '(', '）': ')', '【': '[', '】': ']', '〔': '[', '〕': ']',
    '《': '<', '》': '>', '〈': '<', '〉': '>', '『': '[', '』': ']',
    '—': '-', '–': '-', '―': '-', '−': '-', '…': '.', '·': '.',
    '～': '~', '％': '%', '＋': '+', '×': 'x', '÷': '/', '￥': '¥',
}.items():
    _FOLD_MAP[ord(_src)] = _dst
_FOLD_TABLE = str.maketrans(_FOLD_MAP)
del _o, _c, _src, _dst


# TOC leader dots (目录点线): '四、授权委托书 ....... 7' — 2+ consecutive dots
# (half/full-width, incl. 省略号 …) are filler, never content. They must not
# participate in similarity matching, otherwise unrelated TOC lines pair up
# across documents ('...7' vs '...79'). Replaced with equal-length spaces so
# the 1:1 char↔position mapping stays intact.
_TOC_DOTS_RE = re.compile(r'[.．…]{2,}')


def _strip_toc_dots(text):
    if not text or '.' not in text and '．' not in text and '…' not in text:
        return text
    return _TOC_DOTS_RE.sub(lambda m: ' ' * len(m.group()), text)


def _normalize_for_match(text):
    """Normalize text for comparison: drop whitespace/invisible chars, then
    fold case / full-width / punctuation variants. The mapping is 1:1 on
    surviving characters, so normalized length == non-skipped char count."""
    if not text:
        return ''
    text = _strip_toc_dots(text)
    out = []
    for ch in text:
        if not _is_skip_char(ch):
            out.append(ch)
    return ''.join(out).translate(_FOLD_TABLE)


# Validation regex for history ids (as generated by the history-save code:
# digits, hex, underscores). Used to reject path-traversal ids on the history
# endpoints before any filesystem join.
_VALID_HISTORY_RE = re.compile(r'^[A-Za-z0-9_\-]+$')


def _sanitize_filename(fname):
    """Sanitize a user-supplied filename: strip directory components, null bytes,
    control chars, and leading dots/dashes. Preserves CJK and other non-ASCII
    (unlike werkzeug's secure_filename which strips them).
    """
    if not fname:
        return ''
    # Take only the basename — discard any directory path.
    # Also reject bare '.' / '..' which basename preserves.
    base = os.path.basename(str(fname))
    if base in ('.', '..'):
        return ''
    # Replace path separators (including backslash) and null bytes
    base = base.replace('\\', '_').replace('/', '_').replace('\x00', '')
    # Strip leading dots/dashes (hidden files, e.g. '.bashrc')
    base = base.lstrip('.-')
    # Remove ASCII control chars but keep everything else (CJK, accents, symbols)
    base = re.sub(r'[\x00-\x1f\x7f]', '', base)
    # Collapse multiple underscores/dashes
    base = re.sub(r'_{2,}', '_', base)
    # Truncate to a reasonable max length
    if len(base) > 200:
        name, ext = os.path.splitext(base)
        base = name[:200 - len(ext)] + ext
    return base


def _safe_save(file_storage, prefix=''):
    """Save a Flask FileStorage to UPLOAD_FOLDER under a collision-safe name.

    Strips path components, sanitizes the filename, and prepends a short uuid so
    that concurrent uploads of the same original name can never overwrite each
    other. Returns (saved_path, safe_name) or (None, None) if the filename is
    invalid / has an unsupported extension.
    """
    fname = file_storage.filename or ''
    base = _sanitize_filename(fname)
    if not base or not base.lower().endswith(('.docx', '.doc', '.pdf', '.txt', '.xlsx')):
        return None, None
    safe_name = f'{prefix}{uuid.uuid4().hex[:8]}_{base}'
    fpath = os.path.join(UPLOAD_FOLDER, safe_name)
    file_storage.save(fpath)
    return fpath, safe_name


def _is_within_upload_folder(path):
    """Return True if the real path of `path` is inside UPLOAD_FOLDER.

    Guards against path traversal (../etc/passwd) on user-supplied filenames.
    """
    try:
        rp = os.path.realpath(path)
        return rp.startswith(os.path.realpath(UPLOAD_FOLDER) + os.sep)
    except Exception:
        return False

# ── .doc Conversion ─────────────────────────────────────────────
_DOC_CONVERTER = _find_tool('libreoffice', 'soffice', 'antiword', 'catdoc')


def _doc_converter_kind():
    """Classify the resolved _DOC_CONVERTER path: 'antiword' | 'catdoc' |
    'libreoffice' | None. _DOC_CONVERTER may be a full path (Windows
    soffice.exe), so matching on the basename keeps the old command-name
    semantics working."""
    if not _DOC_CONVERTER:
        return None
    base = os.path.basename(_DOC_CONVERTER).lower()
    if base.startswith('antiword'):
        return 'antiword'
    if base.startswith('catdoc'):
        return 'catdoc'
    return 'libreoffice'

# Cache: {filepath: (mtime, docx_path)} so a .doc is converted at most once per
# process (LibreOffice takes ~2-5s per launch). Without this, extract_doc_text_raw()
# and the extract_text_with_tables() fallback both launch LibreOffice for the same
# file, doubling the cost.
_DOCX_CACHE = {}
_DOCX_CACHE_LOCK = threading.Lock()

def convert_doc_to_docx(filepath):
    """Convert .doc to .docx using LibreOffice. Returns path to .docx or None.

    Results are cached per source file (keyed by path+mtime) so repeated
    extraction attempts do not relaunch LibreOffice.
    """
    try:
        mtime = os.path.getmtime(filepath)
    except OSError:
        mtime = 0

    with _DOCX_CACHE_LOCK:
        cached = _DOCX_CACHE.get(filepath)
        if cached and cached[0] == mtime:
            return cached[1]

    outdir = tempfile.mkdtemp(prefix='doc_conv_')
    docx_path = None
    try:
        subprocess.run(
            [_DOC_CONVERTER, '--headless', '--convert-to', 'docx', '--outdir', outdir, filepath],
            capture_output=True, timeout=60, check=True
        )
        for f in os.listdir(outdir):
            if f.endswith('.docx'):
                docx_path = os.path.join(outdir, f)
                break
    except subprocess.TimeoutExpired:
        logger.warning('LibreOffice .doc->.docx conversion timed out for %s', filepath)
    except Exception as e:
        logger.warning('.doc->.docx conversion failed for %s: %s', filepath, e)
    finally:
        # On failure clean the temp dir now; on success keep it (docx_path lives
        # inside) and cache for reuse.
        if docx_path is None:
            shutil.rmtree(outdir, ignore_errors=True)

    if docx_path:
        with _DOCX_CACHE_LOCK:
            _DOCX_CACHE[filepath] = (mtime, docx_path)
    return docx_path

def extract_doc_text_raw(filepath):
    """Extract text from .doc using antiword, catdoc or LibreOffice"""
    tool = _DOC_CONVERTER
    if not tool:
        return None
    kind = _doc_converter_kind()

    try:
        if kind in ('antiword', 'catdoc'):
            result = subprocess.run([tool, filepath], capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                return result.stdout.strip()
        else:
            # Convert .doc -> .docx first (cached: no repeat LibreOffice launch)
            docx_path = convert_doc_to_docx(filepath)
            if docx_path:
                doc = Document(docx_path)
                lines = [p.text for p in doc.paragraphs]
                return '\n'.join(lines)
    except Exception as e:
        logger.warning('.doc text extraction failed for %s: %s', filepath, e)
    return None


# ── Metadata Extraction ─────────────────────────────────────────
def extract_pdf_metadata(filepath):
    """Extract metadata from PDF files"""
    meta = {}
    try:
        reader = PdfReader(filepath)
        info = reader.metadata or {}
        meta['creator'] = info.get('/Author', info.get('/Creator', ''))
        meta['last_modified_by'] = info.get('/Producer', '')
        meta['created'] = _pdf_date(info.get('/CreationDate', ''))
        meta['modified'] = _pdf_date(info.get('/ModDate', ''))
        meta['application'] = info.get('/Creator', '')
        meta['company'] = ''
        meta['revision'] = ''
        meta['template'] = ''
        meta['total_edit_time'] = ''
        meta['pages'] = str(len(reader.pages))
        meta['words'] = ''
        meta['pdf_title'] = info.get('/Title', '')
        meta['pdf_subject'] = info.get('/Subject', '')
    except Exception as e:
        meta['_error'] = str(e)
    return meta

def _pdf_date(date_str):
    """Convert PDF date format to ISO-like string"""
    if not date_str:
        return ''
    # D:20250308032700+08'00' -> 2025-03-08T03:27:00+08:00
    match = re.match(r'D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', str(date_str))
    if match:
        return f'{match[1]}-{match[2]}-{match[3]}T{match[4]}:{match[5]}:{match[6]}Z'
    return str(date_str)[:50]

def _clean_doc_prop(value, codepage=1200):
    """Take the text before the first control char, decoding bytes if needed.

    olefile.get_metadata() reads each OLE2 property at the correct offset but
    appends trailing bytes (later properties / padding) to CJK string properties
    instead of stopping at the length prefix. The genuine value is the prefix
    before the first NUL/control char. For single-byte codepages (e.g. 936/GBK,
    used by older MS Office .doc), olefile returns raw bytes that we must
    decode with the stream's codepage.
    """
    if not value:
        return ''
    if isinstance(value, bytes):
        if 874 <= codepage <= 1258:
            enc = 'cp%d' % codepage
        elif codepage == 936:
            enc = 'gbk'
        else:
            enc = 'gbk'
        try:
            value = value.decode(enc, errors='replace')
        except Exception:
            value = value.decode('latin-1', errors='ignore')
    s = str(value)
    m = re.search(r'[\x00-\x1f]', s)
    if m:
        s = s[:m.start()]
    return s.strip()


def _extract_doc_kso_props(ole):
    """Extract WPS custom strong-evidence fields (KSO*/ICV) from a .doc.

    These live as user-defined properties inside the DocumentSummaryInformation
    stream, which olefile does not parse. The stream layout (UTF-16-LE) lists
    all property names contiguously, then all values contiguously, so we split
    on control chars, locate the known field names, and pair them with the
    value tokens that follow the name block.

    Robustness: positional index pairing assumes names and values are in the
    same order. To avoid false positives when an interleaved/extra property
    shifts the index, every candidate value is validated against the expected
    format for its field; a mismatched pair is dropped rather than emitted.
    """
    result = {}
    stream_name = '\x05DocumentSummaryInformation'
    if not ole.exists(stream_name):
        return result
    try:
        raw = ole.openstream(stream_name).read()
    except Exception as e:
        logger.debug('KSO .doc stream open failed: %s', e)
        return result
    text = raw.decode('utf-16-le', errors='ignore')
    text = re.sub(r'[\x00-\x1f]+', '|', text)
    tokens = [t.strip() for t in text.split('|') if t.strip()]

    KNOWN = {'KSOProductBuildVer', 'KSOTemplateDocerSaveRecord', 'ICV'}
    field_pos = [(i, t) for i, t in enumerate(tokens) if t in KNOWN]
    if not field_pos:
        return result

    field_names = [t for _, t in field_pos]
    block_end = field_pos[-1][0] + 1
    # Value tokens: everything after the contiguous name block, noise filtered
    value_tokens = []
    for t in tokens[block_end:]:
        if len(t) < 2:
            continue
        if re.fullmatch(r'[一-鿿]{1,2}', t):          # single CJK noise
            continue
        if re.search(r'[-￿]', t) and not re.search(r'[\x20-\x7e]', t):
            continue
        value_tokens.append(t)

    for idx, fname in enumerate(field_names):
        # Try the positionally-paired value first, then scan the next few value
        # tokens for the first one matching this field's expected format. This
        # tolerates a one-off index shift caused by an extra/missing property.
        candidates = value_tokens[idx:]
        for v in candidates[:3]:
            if _looks_like_kso_value(fname, v):
                result[fname] = v
                break
        # If no candidate matched, leave the field absent rather than emit a
        # low-confidence value that could drive a false collusion verdict.
    return result


def _looks_like_kso_value(field, value):
    """Validate a KSO/ICV value against the expected format for its field.

    Conservative: returns True when the value is plausible for the field, False
    when it clearly isn't (e.g. a version field getting a hex hash). Used to
    avoid positional mispairing emitting wrong strong-evidence matches.
    """
    if not value or len(value) > 80:
        return False
    # All KSO/ICV values are ASCII-ish (version strings, build numbers, hex
    # hashes). A value containing CJK ideographs is a mispair.
    if re.search(r'[一-鿿]', value):
        return False
    if field == 'KSOProductBuildVer':
        # e.g. "12.1.0.17120" / "12.1.0.17120" / build numbers: digits, dots,
        # optional letters. Must contain at least one digit.
        return bool(re.search(r'\d', value)) and bool(re.fullmatch(r'[A-Za-z0-9._\-]+', value))
    if field in ('KSOTemplateDocerSaveRecord', 'ICV'):
        # Hardware/user ID records and ICV are long hex or hex+alnum strings.
        # Require >= 6 chars dominated by hex/alnum (not a short version token).
        return len(value) >= 6 and bool(re.fullmatch(r'[A-Za-z0-9_\-]+', value))
    return True


def extract_doc_metadata(filepath):
    """Extract metadata from a legacy .doc (OLE2) file via olefile.

    SummaryInformation fields (author/template/last_saved_by/application/
    revision/create_time/last_saved_time/pages/words) come from
    olefile.get_metadata() with values truncated at the first control char
    (olefile appends trailing bytes to CJK string properties). WPS custom
    strong-evidence fields (KSOProductBuildVer/KSOTemplateDocerSaveRecord/ICV)
    are read from the DocumentSummaryInformation stream directly, since
    olefile does not parse user-defined properties.

    Maps onto the same meta keys as the .docx/.pdf extractors so cross-
    comparison (creator/last_modified_by/application/template/KSO*/ICV +
    created/modified) works uniformly across formats.
    """
    meta = {}
    try:
        ole = olefile.OleFileIO(filepath)
        try:
            m = ole.get_metadata()
            codepage = getattr(m, 'codepage', 1200) or 1200
            # Time fields (reliable)
            ct = getattr(m, 'create_time', None)
            lt = getattr(m, 'last_saved_time', None)
            if ct:
                meta['created'] = ct.strftime('%Y-%m-%dT%H:%M:%SZ')
            if lt:
                meta['modified'] = lt.strftime('%Y-%m-%dT%H:%M:%SZ')
            # Numeric fields (reliable)
            if getattr(m, 'num_pages', None):
                meta['pages'] = str(m.num_pages)
            if getattr(m, 'num_words', None):
                meta['words'] = str(m.num_words)
            # String fields from SummaryInformation (clean trailing noise)
            for attr, key in (('author', 'creator'),
                              ('last_saved_by', 'last_modified_by'),
                              ('creating_application', 'application'),
                              ('template', 'template'),
                              ('revision_number', 'revision'),
                              ('company', 'company')):
                val = _clean_doc_prop(getattr(m, attr, None), codepage)
                if val:
                    meta[key] = val
            # WPS strong-evidence fields from DocumentSummaryInformation
            meta.update(_extract_doc_kso_props(ole))
        finally:
            ole.close()
    except Exception as e:
        meta['_error'] = str(e)
    return meta

def get_file_type(filepath):
    """Detect file type: 'docx', 'doc', 'pdf', 'xlsx', or 'txt'"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.pdf':
        return 'pdf'
    if ext == '.doc':
        return 'doc'
    if ext == '.txt':
        return 'txt'
    if ext == '.xlsx':
        return 'xlsx'
    return 'docx'

def extract_metadata(filepath):
    """Extract metadata from .docx/.xlsx (OOXML zip), .doc, or .pdf.

    .xlsx shares the same docProps/core.xml + app.xml structure as .docx, so
    the OOXML zip path below reads spreadsheet metadata (creator, last
    modified by, KSO fields) for cross-comparison with no extra code."""
    ftype = get_file_type(filepath)
    if ftype == 'pdf':
        return extract_pdf_metadata(filepath)
    if ftype == 'doc':
        return extract_doc_metadata(filepath)
    if ftype == 'txt':
        # Plain text files carry no document metadata; return an empty dict so
        # metadata comparison correctly skips them and the data-sufficiency
        # check reports "无法判断" for the metadata dimension.
        return {}

    meta = {}
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            # Read docProps/core.xml
            if 'docProps/core.xml' in z.namelist():
                from xml.etree.ElementTree import parse
                core = parse(z.open('docProps/core.xml'))
                ns = {
                    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                    'dc': 'http://purl.org/dc/elements/1.1/',
                    'dcterms': 'http://purl.org/dc/terms/',
                    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                }
                def _t(tag, ns_name='dc'):
                    return f'{{{ns.get(ns_name, "")}}}{tag}'

                meta['creator'] = _text(core, _t('creator'))
                meta['last_modified_by'] = _text(core, _t('lastModifiedBy', 'cp'))
                meta['created'] = _text(core, _t('created', 'dcterms'))
                meta['modified'] = _text(core, _t('modified', 'dcterms'))
                meta['revision'] = _text(core, _t('revision', 'cp'))

            # Read docProps/app.xml
            if 'docProps/app.xml' in z.namelist():
                app_xml = parse(z.open('docProps/app.xml'))
                ns2 = {'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}
                def _t2(tag): return f'{{{ns2["ep"]}}}{tag}'
                meta['template'] = _text(app_xml, _t2('Template'))
                meta['total_edit_time'] = _text(app_xml, _t2('TotalTime'))
                meta['pages'] = _text(app_xml, _t2('Pages'))
                meta['words'] = _text(app_xml, _t2('Words'))
                meta['application'] = _text(app_xml, _t2('Application'))
                meta['company'] = _text(app_xml, _t2('Company'))

            # Try to extract KSO custom XML (WPS metadata)
            kso_data = _extract_kso_metadata(z)
            meta.update(kso_data)

    except Exception as e:
        meta['_error'] = str(e)

    return meta

def _text(tree, tag):
    el = tree.find(tag)
    return el.text if el is not None and el.text else ''

def _extract_kso_metadata(z):
    """Extract WPS KSO custom metadata from docx docProps/custom.xml"""
    result = {}
    try:
        if 'docProps/custom.xml' in z.namelist():
            from xml.etree.ElementTree import parse
            custom = parse(z.open('docProps/custom.xml'))
            ns_c = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
            for prop in custom.iter(f'{{{ns_c}}}property'):
                name_attr = prop.get('name', '')
                if not name_attr:
                    continue
                # Extract value from vt:lpwstr or vt:lpstr
                value = None
                for child in prop:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    if tag in ('lpwstr', 'lpstr', 'i4', 'r8', 'bool'):
                        value = child.text
                        break
                if value:
                    result[name_attr] = value
    except Exception as e:
        # KSO metadata is strong-evidence (same WPS hardware/user ID) but
        # optional; never let a malformed custom.xml abort extraction.
        logger.debug('KSO metadata extraction skipped: %s', e)
    return result


# ── Text Extraction ─────────────────────────────────────────────
def _read_text_file(filepath):
    """Read a plain .txt file with encoding auto-detection.

    OCR tools exporting bid documents to .txt commonly produce UTF-8 (with or
    without BOM) or GBK/GB18030 (Windows Chinese). Try UTF-8 first, then fall
    back to GB18030 (a superset of GBK/GB2312), finally decode leniently so a
    partially garbled file still yields text instead of aborting extraction.
    """
    with open(filepath, 'rb') as f:
        raw = f.read()
    for enc in ('utf-8-sig', 'gb18030'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def _read_xlsx_text(filepath, max_rows=5000, max_sheets=30, cancel_event=None):
    """Read a .xlsx workbook into pipe-table text.

    Each sheet becomes a header line '【工作表】name' followed by rows joined
    with ' | ' - the same convention the docx table path uses, so the pipe
    table parsers (price tables, personnel tables) work unchanged.
    data_only=True returns cached formula results; formula cells without a
    cached value emit nothing.
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        logger.warning('openpyxl not installed; cannot extract .xlsx %s', filepath)
        return ''
    try:
        wb = load_workbook(filepath, read_only=True, data_only=True)
    except Exception as e:
        logger.warning('.xlsx load failed for %s: %s', filepath, e)
        return ''
    lines = []
    try:
        for ws in wb.worksheets[:max_sheets]:
            _check_cancelled(cancel_event)
            lines.append(f'【工作表】{ws.title}')
            count = 0
            for row in ws.iter_rows(values_only=True):
                _check_cancelled(cancel_event)
                cells = ['' if v is None else str(v).strip() for v in row]
                if not any(cells):
                    continue
                lines.append(' | '.join(cells))
                count += 1
                if count >= max_rows:
                    break
    finally:
        wb.close()
    return '\n'.join(lines)


# ── OCR fallback for scanned PDFs (lazy, optional deps) ─────────
# RapidOCR (onnxruntime, ~15MB models) + PyMuPDF page rendering. Both are
# optional: when absent, behavior is unchanged (scanned pages yield no text).
# OCR is UNBOUNDED by default (0 = unlimited); set OCR_TIME_BUDGET /
# OCR_MAX_PAGES to cap per-file cost on slow hosts. ANALYSIS_TIMEOUT bounds
# the whole analysis, so very large scanned documents can still time out.
OCR_TIME_BUDGET = float(os.environ.get('OCR_TIME_BUDGET', 0))   # seconds/file, 0=unlimited
OCR_MAX_PAGES = int(os.environ.get('OCR_MAX_PAGES', 0))          # pages/file, 0=unlimited
_ocr_fn = None
_ocr_checked = False


def _get_ocr_fn():
    """Lazy-init the OCR callable (fitz_page, dpi=200) -> str, or None.

    The page (from the caller's already-open PyMuPDF document) is rendered
    straight to a numpy array — a PNG encode/decode round trip here cost
    ~0.2-0.5s per 200-DPI A4 page.

    Dependencies (pymupdf, rapidocr_onnxruntime) are imported lazily so the
    base deployment keeps its minimal footprint; absence degrades gracefully.
    """
    global _ocr_fn, _ocr_checked
    if _ocr_checked:
        return _ocr_fn
    _ocr_checked = True
    try:
        import fitz  # capability check: pages are rendered by the caller's doc
        import numpy as np
        import cv2
        from rapidocr_onnxruntime import RapidOCR
        engine = RapidOCR()

        def _ocr_page(page, dpi=200):
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n)
            # fitz samples are RGB; the engine consumes cv2-style BGR
            if pix.n == 3:
                img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
            elif pix.n == 1:
                img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
            result, _ = engine(img)
            if not result:
                return ''
            return '\n'.join(item[1] for item in result)

        _ocr_fn = _ocr_page
        logger.info('OCR fallback ready (rapidocr + pymupdf), budget %.0fs / %d pages per file',
                    OCR_TIME_BUDGET, OCR_MAX_PAGES)
    except Exception as e:
        logger.info('OCR deps unavailable, scanned PDFs will yield no text: %s', e)
    return _ocr_fn


def extract_text(filepath):
    """Extract all text from a .docx file"""
    doc = Document(filepath)
    paragraphs = []
    for p in doc.paragraphs:
        text = p.text
        if text.strip():
            paragraphs.append(text)
    return '\n'.join(paragraphs)


# ── PDF table extraction via PyMuPDF (optional structural channel) ──
# pypdf's plain text layer flattens tables into space-separated runs, which
# the heuristic PDF parsers can only partially recover. When pymupdf is
# importable (it is already an optional OCR dependency), page.find_tables()
# recovers cell structure; rows are emitted as ' | '-joined lines — the same
# convention docx/xlsx use — so the pipe-table personnel/price parsers work
# on PDFs unchanged. Only pages whose text mentions table-relevant keywords
# are processed, bounding the cost on large documents.
_PDF_TABLE_TRIGGER = re.compile(
    r'(?:报价|价格|费用|金额|总价|合计|一览|开标|单价|费率|'
    r'姓名|人员|职务|职称|电话|身份证|授权|账号)')
MAX_PDF_TABLE_PAGES = 300


def _fitz_page_tables_as_pipes(page):
    """Extract tables from one PyMuPDF page as pipe-separated lines."""
    rows_out = []
    try:
        finder = page.find_tables()
    except Exception:
        return ''
    for tbl in getattr(finder, 'tables', []):
        try:
            rows = tbl.extract()
        except Exception:
            continue
        for row in rows:
            cells = ['' if c is None else str(c).replace('\n', ' ').strip()
                     for c in row]
            if any(cells):
                rows_out.append(' | '.join(cells))
    return '\n'.join(rows_out)


def extract_text_with_tables(filepath, max_pages=MAX_PDF_PAGES, on_progress=None,
                             cancel_event=None):
    """Extract text including tables from .docx, .doc, .pdf, or .txt.

    max_pages: max PDF pages to process (0 = unlimited, default MAX_PDF_PAGES).
    on_progress: optional callback(phase, current, total, has_text, detail).
                 phase values: 'pdf_page', 'pdf_done', 'pdf_early_stop',
                 'pdf_ocr_start', 'pdf_ocr' (scanned-page OCR fallback).
    """
    ftype = get_file_type(filepath)

    if ftype == 'txt':
        return _read_text_file(filepath)

    if ftype == 'xlsx':
        return _read_xlsx_text(filepath, cancel_event=cancel_event)

    if ftype == 'pdf':
        reader = PdfReader(filepath)
        total_pages = len(reader.pages)
        fname = os.path.basename(filepath)
        lines = []
        pages_with_text = 0
        empty_streak = 0

        # OCR state (only initialized on the first page that needs it)
        ocr_fn = None
        ocr_started = False
        ocr_pages_done = 0
        ocr_time_spent = 0.0

        # Shared PyMuPDF document: opened lazily at most once and reused by
        # both the scanned-page OCR fallback and the structural table channel
        # (reopening the file per OCR page re-parsed the whole xref each time).
        # None when PyMuPDF is unavailable or the open fails.
        fitz_doc = None
        fitz_open_tried = False

        def _open_fitz():
            nonlocal fitz_doc, fitz_open_tried
            if not fitz_open_tried:
                fitz_open_tried = True
                try:
                    import fitz
                    fitz_doc = fitz.open(filepath)
                except Exception as e:
                    logger.info('PyMuPDF unavailable, scanned-page OCR and PDF '
                                'table extraction stay disabled: %s', e)
            return fitz_doc

        table_pages_done = 0

        for i, page in enumerate(reader.pages):
            # Cancellation check (per page — an OCR page itself is ~1-3s)
            _check_cancelled(cancel_event)
            # Page limit check
            if max_pages > 0 and i >= max_pages:
                if on_progress:
                    on_progress('pdf_early_stop', i, total_pages, bool(lines),
                              f'"{fname}" 页数过多，已截断处理前{max_pages}页（共{total_pages}页），建议压缩或使用文字版PDF')
                break

            text = page.extract_text()
            has_text = bool(text and text.strip())

            # ── OCR fallback: page has no embedded text (scanned/image page) ──
            # max_pages == 0 means unlimited (see docstring), so OCR must not
            # be disabled by the `i < max_pages` page-limit guard.
            if not has_text and (max_pages == 0 or i < max_pages):
                if not ocr_started:
                    ocr_fn = _get_ocr_fn()
                    ocr_started = True
                    if ocr_fn and on_progress:
                        on_progress('pdf_ocr_start', i + 1, total_pages, False,
                                    f'"{fname}" 含无文字页面，正在启用OCR识别扫描件内容…')
                # 0 = unlimited for both budgets
                budget_ok = (OCR_TIME_BUDGET == 0 or ocr_time_spent < OCR_TIME_BUDGET)
                pages_ok = (OCR_MAX_PAGES == 0 or ocr_pages_done < OCR_MAX_PAGES)
                if ocr_fn and budget_ok and pages_ok:
                    t0 = time.time()
                    try:
                        doc = _open_fitz()
                        ocr_text = ocr_fn(doc[i]) if doc is not None else ''
                    except Exception as e:
                        logger.warning('OCR failed on %s page %d: %s', fname, i + 1, e)
                        ocr_text = ''
                    ocr_time_spent += time.time() - t0
                    ocr_pages_done += 1
                    if ocr_text.strip():
                        text = ocr_text
                        has_text = True
                        if on_progress and ocr_pages_done % 5 == 0:
                            budget_note = ('（达到时间预算，剩余页面跳过）'
                                           if OCR_TIME_BUDGET > 0 and ocr_time_spent >= OCR_TIME_BUDGET else '')
                            on_progress('pdf_ocr', i + 1, total_pages, True,
                                        f'"{fname}" OCR识别中：已完成 {ocr_pages_done} 页{budget_note}')
                    elif OCR_MAX_PAGES > 0 and ocr_pages_done == OCR_MAX_PAGES and on_progress:
                        on_progress('pdf_ocr', i + 1, total_pages, True,
                                    f'"{fname}" OCR已达页数上限 {OCR_MAX_PAGES} 页，剩余图片页跳过')

            # ── Structural table channel (PyMuPDF) ──
            # Runs only on text-bearing pages that mention table-relevant
            # keywords, so cost stays bounded on long documents. The recovered
            # pipe rows are appended to the page text; pypdf's flattened run
            # stays too (duplicate content is harmless — same convention as
            # the docx path, which appends tables after paragraphs).
            tbl_text = ''
            if has_text and table_pages_done < MAX_PDF_TABLE_PAGES \
                    and _PDF_TABLE_TRIGGER.search(text):
                doc = _open_fitz()
                if doc is not None:
                    _check_cancelled(cancel_event)
                    try:
                        tbl_text = _fitz_page_tables_as_pipes(doc[i])
                    except Exception as e:
                        logger.warning('PDF table extraction failed on %s '
                                       'page %d: %s', fname, i + 1, e)
                    if tbl_text:
                        table_pages_done += 1

            if has_text or tbl_text:
                lines.append(text + ('\n' + tbl_text if tbl_text else ''))
                pages_with_text += 1
                empty_streak = 0
            else:
                empty_streak += 1

            # Early termination: after sampling enough pages with zero text
            if i >= 50 and empty_streak >= MAX_EMPTY_PAGE_STREAK:
                if on_progress:
                    on_progress('pdf_early_stop', i + 1, total_pages, bool(lines),
                              f'"{fname}" 连续{MAX_EMPTY_PAGE_STREAK}页无文字，疑似全图片扫描件，跳过剩余{total_pages - i - 1}页')
                break

            # Progress callback every 20 pages or on last page
            if on_progress and (i % 20 == 0 or i == total_pages - 1):
                on_progress('pdf_page', i + 1, total_pages, True, None)

        # Final callback
        if on_progress:
            if not lines and total_pages > 0:
                if ocr_started and ocr_fn:
                    on_progress('pdf_done', total_pages, total_pages, False,
                              f'"{fname}" 文字与OCR均未提取到内容，可能为空白或图片质量过低的扫描件')
                else:
                    on_progress('pdf_done', total_pages, total_pages, False,
                              f'"{fname}" 未提取到任何文字，可能为全图片扫描件。请上传可复制文字版PDF或Word文件。')
            elif pages_with_text > 0:
                suffix = f'（其中OCR识别 {ocr_pages_done} 页）' if ocr_pages_done else ''
                on_progress('pdf_done', total_pages, total_pages, True,
                          f'"{fname}" 提取完成：{pages_with_text}页有文字{suffix}')

        if fitz_doc is not None:
            fitz_doc.close()
        return '\n'.join(lines)

    if ftype == 'doc':
        text = extract_doc_text_raw(filepath)
        if text:
            return text
        # Fallback: try LibreOffice conversion
        docx_path = convert_doc_to_docx(filepath)
        if docx_path:
            return extract_text_with_tables(docx_path)
        return ''

    doc = Document(filepath)
    lines = []
    for p in doc.paragraphs:
        _check_cancelled(cancel_event)
        lines.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            _check_cancelled(cancel_event)
            row_text = ' | '.join(cell.text for cell in row.cells)
            lines.append(row_text)
    return '\n'.join(lines)


# ── Personnel Extraction ────────────────────────────────────────
def _find_personnel_sections(text):
    """Identify personnel-related sections in bid text by chapter markers.
    Returns list of {type, text, start, end}."""
    section_markers = {
        'auth_letter': [
            '法定代表人授权委托书', '法定代表人授权书', '授权委托书',
            '法人授权书', '法人代表授权书', '法人授权委托书'
        ],
        'legal_rep_proof': [
            '法定代表人身份证明', '法定代表人证明', '法人代表证明',
            '单位负责人证明', '法定代表人资格证明'
        ],
        'personnel_table': [
            '项目管理机构', '项目组成员', '主要人员', '项目成员',
            '拟投入人员', '拟派人员', '项目团队', '组织机构',
            '人员配备', '人员配置', '岗位人员', '主要管理人员',
            '人员一览表', '主要人员一览', '项目人员', '关键人员',
            '人员简历', '劳动力计划', '技术人员情况', '管理人员情况',
            '人员与分工'
        ],
        'qualification': [
            '投标人基本情况表', '资格审查资料', '投标人资格',
            '企业基本情况', '公司简介', '单位简介'
        ],
        'signature_page': [
            '签字盖章', '签章', '签字或盖章', '盖章签字',
            '法定代表人或其委托代理人', '投标人（盖单位章）',
            '（单位公章）', '（盖章）'
        ],
        'cover_letter': [
            '投标函', '投标书', '投标文件', '报价函'
        ],
    }

    found = []
    for section_type, markers in section_markers.items():
        for marker in markers:
            idx = text.find(marker)
            while idx >= 0:
                start = max(0, idx - 200)
                end = min(idx + 5000, len(text))
                for next_marker in [
                    '\n一、', '\n二、', '\n三、', '\n四、', '\n五、',
                    '\n1.', '\n2.', '\n3.', '\n4.', '\n5.',
                    '\n六、', '\n七、', '\n八、',
                ]:
                    ep = text.find(next_marker, idx + 10)
                    if ep > idx and ep < end:
                        end = ep
                sec_text = text[start:end]
                found.append({'type': section_type, 'text': sec_text, 'start': start, 'end': end})
                idx = text.find(marker, idx + len(marker))
    return found


# Chinese surname dictionary for name validation. A 2-4 char pure-CJK string
# whose first character is not a known surname (and has no compound-surname
# prefix) is almost always a table fragment, not a person. The full 百家姓
# plus common extended surnames keeps false negatives near zero; the check
# can only REJECT candidates, never accept new ones, so precision strictly
# improves. Minority transliterated names (阿不来提·买买提) skip this check.
_CHINESE_SURNAMES = set(
    '赵钱孙李周吴郑王冯陈蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章云苏潘葛奚范彭郎'
    '鲁韦昌马苗凤花方俞任袁柳鲍史唐费廉岑薛雷贺倪汤殷罗毕安常乐于时傅皮卞齐康伍余元卜顾孟平黄和穆'
    '萧尹姚邵汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪舒屈项祝董梁杜阮蓝闵席季麻强贾路娄危江童颜郭'
    '梅盛林刁钟徐邱骆高夏蔡田樊胡凌霍虞万支柯管卢莫经房裘缪干解应宗丁宣邓郁杭洪包诸左石崔吉钮龚'
    '程嵇邢滑裴陆荣翁荀羊於惠甄曲家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧隗山谷车侯宓蓬全郗班仰秋'
    '仲伊宫宁仇栾暴甘钭厉戎祖武符刘景詹束龙叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔'
    '阴胥能苍双闻莘党翟谭贡劳逄姬申扶堵冉宰郦雍却璩桑桂濮牛寿通边扈燕冀郏浦尚农温别庄晏柴瞿阎充'
    '慕连茹习宦艾鱼容向古易慎戈廖庾终暨居衡步都耿满弘匡国文寇广禄阙东欧殳沃利蔚越夔隆师巩厍聂晁'
    '勾敖融冷訾辛阚那简饶空曾毋沙乜养鞠须丰巢关蒯相查后荆红游竺权逯盖益桓公'
    '覃佘区冼招植苟代阿牟漆付兰单岳滕肖郝闫'
)
_COMPOUND_SURNAMES = (
    '欧阳', '上官', '司马', '诸葛', '夏侯', '皇甫', '尉迟', '公孙', '长孙',
    '慕容', '司徒', '司空', '端木', '独孤', '南宫', '万俟', '闻人', '东郭',
    '宇文', '呼延', '鲜于', '赫连', '澹台', '淳于', '太叔', '申屠', '公冶',
    '宗政', '濮阳', '钟离', '令狐', '轩辕', '百里', '第五',
)


def _is_person_name(name):
    """Validate that a string looks like a person name (Chinese or foreign).

    Chinese names: 2-4 CJK characters, screened against company/city keywords.
    Foreign names: ASCII letters + spaces/hyphens/periods, 2-40 chars, screened
    against the same company keyword list (lowercased).
    """
    if not name:
        return False
    name_stripped = name.strip()
    name_lower = name_stripped.lower()

    # Company name indicators — reject these (defined before both paths so
    # the foreign-name branch can reference them).
    company_keywords = [
        '公司', '集团', '有限', '责任', '股份', '进出口', '科技',
        '技术', '工程', '实业', '贸易', '企业', '中心', '研究院',
        '北京', '上海', '深圳', '广州', '成都', '武汉', '南京',
        '西安', '杭州', '苏州', '东莞', '佛山', '无锡', '宁波',
        '温州', '南通', '长沙', '郑州', '济南', '青岛', '大连',
        '厦门', '合肥', '福州', '南宁', '昆明', '贵阳', '海口',
        '哈尔滨', '长春', '沈阳', '太原', '石家庄', '兰州', '乌鲁木齐',
        '呼和浩特', '银川', '西宁', '拉萨', '南昌', '珠海', '惠州',
        '中山', '江门', '肇庆', '汕头', '天津', '重庆',
        # SOE/industry indicators commonly found in company names
        '航天', '星网', '移动', '联通', '电信', '石油', '石化',
        '电力', '核电', '钢铁', '中核', '中铁', '中建', '中交',
        '中化', '中粮', '中船', '中车', '中航',
        # Professional/status terms not found in person names
        '执业', '评估师',
    ]

    # ── Foreign / mixed-name path ──
    # Allow ASCII-alpha names with spacing/hyphens (e.g. "John Smith", "Jean-Luc").
    # Length relaxed to 40: full names like "Jean-Luc Rémond-Martinez" are real.
    if re.search(r'[A-Za-z]', name_stripped):
        if not re.fullmatch(r'[A-Za-zÀ-ÿ\-.\sĀ-ſ]{2,40}', name_stripped):
            return False
        for kw in company_keywords:
            if kw.lower() in name_lower:
                return False
        return True

    # ── Chinese name path ──
    # PDF single-char blocks insert spaces INSIDE a name ('张 三'); collapse
    # before validation so spaced names validate and dedup correctly. The
    # foreign-name path above keeps its spaces.
    name_stripped = re.sub(r'\s+', '', name_stripped)
    name_lower = name_stripped.lower()
    _FUNCTION_CHARS = set('对的了是为在与和就被就以从把向由因所给见')
    # Placeholder text that looks like a name label (CV form fields, etc.)
    _PLACEHOLDER_LABELS = (
        '姓名', '职务', '签字', '盖章', '授权', '电话', '地址', '传真',
        '牵头', '负责', '联系', '经办', '复核', '审核', '批准', '执行',
        '毕业学校', '毕业院校', '主要经历', '工作经历', '学习经历',
        '所学专业', '修读专业', '教育背景', '出生年月', '出生日期',
        '身份证号', '证件号码', '联系电话', '手机号码', '联系方式',
        '电子邮箱', '邮箱地址', '通讯地址', '联系地址', '家庭住址',
        '资格证书', '执业资格', '政治面貌', '婚姻状况', '健康状况',
        '外语水平', '语言能力', '技术职称', '专业职称', '现任职务',
        '担任职务', '相关工作', '相关年限', '工作年限', '从业时间',
        '专业领域', '研究方向', '计算机', '外语语种', '熟练程度',
        # Personal-info form fields (PDF tables split into 性别/民族/学历… cells)
        '性别', '民族', '籍贯', '学历', '学位', '年龄', '身高', '体重',
        '户籍', '党派', '血型', '婚姻', '家庭住址', '现住址',
        # Function words that are not names but pass 2-char CJK validation
        '本人', '我方', '我们', '该人', '此人', '对方', '甲方', '乙方', '丙方',
        # 职称/级别词 — PDF "姓名 职称 分工" tables put these in the 职称 column
        # ('张伟 中级 项目负责人'); they must never be captured as names.
        '中级', '高级', '初级', '正高', '副高', '教授', '副教授', '讲师', '助教',
        '研究员', '副研究员', '工程师', '技师', '助理', '总工', '高工',
        # Table column labels / role words that look like names
        '投标人名称', '项目名称', '公司名称', '企业名称', '单位名称',
        '招标人', '投标人', '采购人', '供应商', '供应商名称',
        '负责人', '联系人', '工程师', '技术员', '管理员', '组长', '组员',
        '主任', '主管', '专员', '代表', '代理人', '经办人', '签字人',
        '法定代表人', '授权代表', '项目负责人', '技术负责人', '项目经理',
    )

    # Duty-label suffix rule: any CJK string ENDING in a duty/role word —
    # simplified or traditional ('任中职务', '现任职务', '任何岗位', '任職務'…) —
    # is a form/table column label, never a person name, even when it starts
    # with a real surname like 任. Observed: '任中职务 项目经理' rows fed the
    # column label into all_persons as a project_manager.
    if re.search(r'(?:职务|職務|岗位|崗位|职称|職稱|角色|职责|職責)$', name_stripped):
        return False
    # Collective-word rule: section headers and table captions
    # ('项目团队', '关键人员', '管理人员', '项目成员', '项目分工', '项目部'…)
    # start with real surnames (项/关/管) and would pass every check below,
    # but no genuine person name contains a collective/organizational word.
    if re.search(r'项目|項目|人员|人員|团队|團隊|成员|成員|机构|機構|分工|部门|部門|简历|簡歷|配备|配備|配置|一览|一覽', name_stripped):
        return False

    # Minority-ethnic names use a middle-dot separator (· U+00B7, •, ・),
    # e.g. 阿不来提·买买提, 迪力夏提·阿不都·热合曼. Validate each CJK
    # segment independently with the same filters as plain names.
    if re.search(r'[·•・]', name_stripped):
        parts = [p for p in re.split(r'[·•・]', name_stripped) if p]
        if not 2 <= len(parts) <= 3:
            return False
        for p in parts:
            if p in _PLACEHOLDER_LABELS or not 2 <= len(p) <= 4 \
                    or not re.match(r'^[一-鿿]+$', p):
                return False
            for kw in company_keywords:
                if kw in p:
                    return False
            if p[0] in _FUNCTION_CHARS:
                return False
        return True

    if len(name_stripped) < 2 or len(name_stripped) > 4:
        return False
    # Reject placeholder text that looks like a name label
    if name_stripped in _PLACEHOLDER_LABELS:
        return False
    for kw in company_keywords:
        if kw in name_lower:
            return False
    # Must consist of Chinese characters only
    if not re.match(r'^[一-鿿]+$', name_lower):
        return False
    # Surname check: first char must be a known (possibly compound) surname.
    # Transliterated minority names reach the dotted path above and skip this.
    if name_stripped[0] not in _CHINESE_SURNAMES \
            and not name_stripped.startswith(_COMPOUND_SURNAMES):
        return False
    # Reject names starting with function/grammar characters
    # (these are prepositions, particles, etc. — never start a Chinese person name)
    if name_lower[0] in _FUNCTION_CHARS:
        return False
    return True


def _extract_from_auth_section(section_text, info):
    """Extract legal rep, authorized rep from authorization letter section."""
    # Pattern 0: "我张三（姓名）系四川某某电子科技有限公司（供应商名称）的法定代表人"
    # company_name is captured even when the name fails validation — the two
    # facts are independent, and a rare-surname miss must not also lose the
    # company (observed: 兰某某 rejected → company disappeared too).
    m = re.search(r'(?:本人\s*)?我?\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})\s*[（(]姓名[）)]\s*系\s*(.{1,40}?)\s*[（(]供应商名称[）)]\s*的法定代表人', section_text)
    if m:
        if not info.get('company_name'):
            info['company_name'] = _clean_company(m.group(2).strip())
        if _is_person_name(m.group(1).strip()):
            info['legal_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.95})

    # Pattern 0b: "（兰某某）系（北京某某航天技术有限公司）的法定代表人"
    # 法定代表人资格证明书 form: name AND company in unlabeled parentheses.
    m = re.search(r'[（(]\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})\s*[）)]\s*系\s*[（(]?\s*(.{2,40}?)\s*[）)]?\s*的法定代表人', section_text)
    if m:
        if not info.get('company_name'):
            info['company_name'] = _clean_company(m.group(2).strip())
        if not info['legal_rep'] and _is_person_name(m.group(1).strip()):
            info['legal_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.92})

    # Pattern 1: "姓名：XXX 职务：XXX 系 XXX 的法定代表人"
    if not info['legal_rep']:
        m = re.search(r'姓名[：:]\s*([^\s]{2,10})\s*[\s\S]{0,100}?系\s*(.{1,30}?)\s*的法定代表人', section_text)
        if m:
            cand = m.group(1).strip().rstrip('：:')
            # Validate the captured name: PDF form layouts leave blank-field
            # fragments like '姓名：性别：男' where the colon survives.
            if _is_person_name(cand):
                info['legal_rep'] = cand
                if not info.get('company_name'):
                    info['company_name'] = _clean_company(m.group(2).strip())
                info['all_persons'].append({'name': cand, 'role': 'legal_rep', 'confidence': 0.90})

    # Pattern 2: "本人 XXX 系 XXX 的法定代表人"
    if not info['legal_rep']:
        m = re.search(r'(?:本人\s*)?([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})\s*(?:[（(]姓名[）)])?\s*系\s*(.{1,30}?)\s*的法定代表人', section_text)
        if m:
            if not info.get('company_name'):
                info['company_name'] = _clean_company(m.group(2).strip())
            if _is_person_name(m.group(1).strip()):
                info['legal_rep'] = m.group(1).strip()
                info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.85})

    # Pattern 3: "（王戈、董事长）代表本公司授权（赵凯、销售经理）"
    # NOTE: Some documents insert company name between the auth clause and agent name:
    #   "（王戈、董事长）代表本公司授权（北京东方中科...）的在下面签字的（赵凯、销售经理）"
    m = re.search(r'[（(]([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})[、，].{0,6}?[）)]\s*代表本公司授权\s*[（(]([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})', section_text)
    if m:
        if not info['legal_rep']:
            info['legal_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.85})
        if not info['authorized_rep']:
            name = m.group(2).strip()
            if _is_person_name(name):
                info['authorized_rep'] = name
                info['all_persons'].append({'name': name, 'role': 'authorized_rep', 'confidence': 0.85})
            else:
                # Captured group is likely a company name fragment.
                # Try to find the actual person name after the company: "（XXX、role）为本公司"
                post_match = section_text[m.end():m.end() + 200]
                m2 = re.search(r'[（(]([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})[、，].{0,6}?[）)]\s*(?:为本公司的合法代理人|为代理人)', post_match)
                if m2:
                    name2 = m2.group(1).strip()
                    if _is_person_name(name2):
                        info['authorized_rep'] = name2
                        info['all_persons'].append({'name': name2, 'role': 'authorized_rep', 'confidence': 0.85})

    # Pattern 4: "签字代表（赵凯、销售经理）" — common in bid letters
    if not info['authorized_rep']:
        m = re.search(r'签字代表[（(]([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})[、，]', section_text)
        if m:
            name = m.group(1).strip()
            if _is_person_name(name):
                info['authorized_rep'] = name
                info['all_persons'].append({'name': name, 'role': 'authorized_rep', 'confidence': 0.90})

    # Pattern 5: "现委托 XXX（姓名）为我方代理人"
    if not info['authorized_rep']:
        m = re.search(r'(?:现委托|委托)\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})\s*[（(]姓名[）)]', section_text)
        if m:
            info['authorized_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['authorized_rep'], 'role': 'authorized_rep', 'confidence': 0.90})

    # Pattern 6: "代理人：XXX" or "授权代表：XXX" — with person name validation
    if not info['authorized_rep']:
        m = re.search(r'(?:授权委托代理人|委托代理人|代理人|授权代表|被授权人|受托人|签字代表|投标代表)[：:]\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿])', section_text)
        if m:
            name = m.group(1).strip()
            if _is_person_name(name):
                info['authorized_rep'] = name
                info['all_persons'].append({'name': name, 'role': 'authorized_rep', 'confidence': 0.80})

    # Pattern 7: "法定代表人：XXX"
    if not info['legal_rep']:
        m = re.search(r'(?:法定代表人|单位负责人|法人代表)[：:]\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿])', section_text)
        if m:
            name = m.group(1).strip()
            if _is_person_name(name):
                info['legal_rep'] = name
                info['all_persons'].append({'name': name, 'role': 'legal_rep', 'confidence': 0.80})

    # Pattern 8: "兹委托 XXX（同志）为我(方/公司)…代理人" / "现授权/特授权"
    if not info['authorized_rep']:
        m = re.search(r'(?:兹委托|现委托|兹授权|现授权|特授权|特此委托)\s*'
                      r'([一-鿿]{2,4}(?:[ 	]*[·•・][ 	]*[一-鿿]{2,4}){0,2})\s*(?:同志)?\s*'
                      r'(?:为|作为)[^。]{0,40}?代理\s*人', section_text)
        if not m:
            # Spaced-name retry ('兹委托 李 明 同志为我方代理人'). The trailing
            # guard is safe HERE only because the strict pass above already
            # handled zero-separator '现委托李四为…' forms, where 为 sits right
            # after the name and the guard would reject the correct capture.
            m = re.search(r'(?:兹委托|现委托|兹授权|现授权|特授权|特此委托)\s*'
                          r'([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿])\s*'
                          r'(?:同志)?\s*'
                          r'(?:为|作为)[^。]{0,40}?代理\s*人', section_text)
        if m and _is_person_name(m.group(1).strip()):
            info['authorized_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['authorized_rep'], 'role': 'authorized_rep', 'confidence': 0.85})

    # Pattern 9: "委托：XXX" / "代理人 XXX（签字）" — bare agent label without colon
    if not info['authorized_rep']:
        m = re.search(r'(?:委托|代理人)\s+([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})\s*[（(]?(?:签字|签章|盖章|姓名)', section_text)
        if m and _is_person_name(m.group(1).strip()):
            info['authorized_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['authorized_rep'], 'role': 'authorized_rep', 'confidence': 0.75})


# Known role/title strings that should NOT be treated as person names
# (shared by section-scoped and pipe-table personnel extraction).
_ROLE_KEYWORDS = [
    '项目经理', '项目负责人', '技术负责人', '技术总监', '总工程师',
    '安全员', '质量员', '施工员', '材料员', '资料员', '造价员', '预算员',
    '安全负责人', '项目副经理', '商务经理', '财务负责人', '设计负责人',
    '质量负责人', '现场负责人', '合同经理', '采购经理', '施工经理',
    '测试负责人', '运维负责人', '集成负责人', '实施负责人',
    '法定代表人', '授权代表', '代理人', '被授权人', '受托人',
    '团队成员', '项目成员', '组员', '组长',
    # Additional role/title keywords that appear as table cell values
    '核心人员', '核心团队成员', '项目核心人员',
    '总协调人', '总负责人', '总协调助理',
    '工作组组长', '小组组长', '评估助理',
]


def _extract_from_personnel_table(section_text, info):
    """Extract project members from personnel/team tables."""
    # 职称/级别词（表格"职称"列的值，如"张伟 中级 项目负责人"）— 绝不能当姓名
    _TITLE_WORDS = r'(?:高级|中级|初级|正高级|副高级|教授|副教授|讲师|助教|研究员|副研究员|工程师|高级工程师|助理工程师|技师|高级技师|助理)?'
    patterns = [
        # '姓名：...' and '职务：...' often sit on different lines in resume
        # forms; the span may cross newlines but never another 姓名 label
        # (which would pair this row's name with the NEXT row's role).
        r'姓名[：:]\s*([一-鿿]{2,4}(?:[ 	]*[·•・][ 	]*[一-鿿]{2,4}){0,2})\s*(?:(?!姓名[：:])[\s\S]){0,60}?(?:职务|岗位|角色|职称)[：:]\s*([一-鿿]{2,10})',
        # Single space or tab between name and role is common in both docx
        # and PDF extraction ('王强 项目经理'); require the role keyword
        # so a bare space-separated line cannot be a false positive.
        # PDF "序号 姓名 职称 分工" tables produce '张伟 中级 项目负责人' —
        # allow one title word between name and role so 张伟 is captured
        # instead of the 职称 column value 中级.
        # Flattened rows may also carry a duty-LABEL column between name and
        # role ('陈刚 任中职务 项目经理'): skip one label word so the real
        # name before it is captured instead of the label itself (which
        # '任中职务' starts with surname 任 and would otherwise look like one).
        # The colon/zero-space tolerance covers '陈刚 任中职务：项目经理';
        # traditional label forms (任職務) are skipped the same way.
        r'([一-鿿]{2,4}(?:[ 	]*[·•・][ 	]*[一-鿿]{2,4}){0,2})\s+(?:[一-鿿]{0,2}(?:职务|職務|岗位|崗位|职称|職稱|角色|职责|職責)[：:]?\s*)?' + _TITLE_WORDS + r'\s*(项目经理|项目负责人|技术负责人|技术总监|总工程师|安全员|质量员|施工员|材料员|资料员|造价员|预算员)',
        r'(项目经理|项目负责人|技术负责人|技术总监|总工程师)[：:]\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿])',
        # Role-then-name pairs must stay on ONE line: with '\s+' the role at
        # the end of one table row grabbed the next row's name
        # ('王强 项目经理\n李勇 施工员' made 李勇 a project_manager).
        r'(项目经理|项目负责人|技术负责人|安全负责人)[ \t]+([一-鿿]{2,4}(?:[ 	]*[·•・][ 	]*[一-鿿]{2,4}){0,2})(?![一-鿿])',
        # Reversed label order: "职务：项目经理 ... 姓名：张三" (role first)
        r'(?:职务|岗位|职称)[：:]\s*([一-鿿]{2,10})\s*[\s\S]{0,60}?姓名[：:]\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿])',
        # Bare "姓名：张三" without a role label (name-only tables, resumes).
        # The lookahead rejects the next label ('姓名：性别：男' → '性别' is
        # followed by a colon and never captured).
        r'姓名[：:]\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿]|：|:)',
    ]
    # 'role name' immediately before a match start means the pairing belongs
    # to pattern 2 above ('项目经理 王强 技术负责人 李四'): re-pairing that
    # name with the FOLLOWING role would double-tag the person. A post-filter
    # instead of a lookbehind so any run of spaces/tabs is covered.
    _ROLE_BEFORE_NAME = re.compile(
        r'(?:项目经理|项目负责人|技术负责人|安全负责人|技术总监|总工程师)[ \t]+$')
    for pi, pat in enumerate(patterns):
        for m in re.finditer(pat, section_text):
            if pi == 1 and _ROLE_BEFORE_NAME.search(section_text, 0, m.start()):
                continue
            groups = m.groups()
            if len(groups) == 2:
                # Determine which is name and which is role.
                # Priority 1: known role keywords
                if groups[0] in _ROLE_KEYWORDS:
                    role_str, name = groups[0], groups[1]
                elif groups[1] in _ROLE_KEYWORDS:
                    name, role_str = groups[0], groups[1]
                # Priority 2: length heuristic (name 2-4 chars, role can be longer)
                elif len(groups[0]) <= 4 and re.match(r'^[一-鿿]+$', groups[0]):
                    name, role_str = groups[0], groups[1]
                else:
                    name, role_str = groups[1], groups[0]
                role = _infer_role_label(role_str)
            else:
                name = groups[0]
                role = 'team_member'

            name = name.strip()
            # Reject if the "name" is actually a known role/title string
            if name in _ROLE_KEYWORDS:
                continue
            # Validate the name looks like an actual person name (not column label, company name, etc.)
            if not _is_person_name(name):
                continue
            # Same (name, role) pair can be matched by several patterns on the
            # same row ('杨帆 任中职务 项目经理' + '...项目经理\n杨帆'); keep one.
            if any(p['name'] == name and p['role'] == role for p in info['all_persons']):
                continue
            info['all_persons'].append({'name': name, 'role': role, 'confidence': 0.80})


def _parse_personnel_pipe_table(text, info):
    """Extract project members from docx/xlsx pipe-separated tables.

    docx table rows are appended at the END of the extracted text (after all
    paragraphs), so the section-scoped extractors never see them. This parser
    scans pipe-table regions for headers like 姓名|职务|职称|电话 and maps
    columns explicitly instead of relying on inline sentence patterns.
    """
    _NAME_HDR = re.compile(r'(?:姓\s*名|人员姓名|成员姓名|拟投入.{0,6}人员|人员名称|主要人员|关键人员|项目人员)')
    _ROLE_HDR = re.compile(r'(?:职\s*务|岗\s*位|职\s*称|角\s*色|担任职务|项?目?角色)')
    _PHONE_HDR = re.compile(r'(?:联?系?电话|手\s*机|移动电?话|联系方式)')
    _ID_HDR = re.compile(r'(?:身份证|证件号码|身份证明)')
    _CERT_HDR = re.compile(r'(?:证书|资格|执业|注册)')

    lines = text.split('\n')
    n = len(lines)
    i = 0
    while i < n:
        if '|' not in lines[i]:
            i += 1
            continue
        # Region = consecutive pipe lines (blank lines tolerated inside)
        start = i
        while i < n and ('|' in lines[i] or not lines[i].strip()):
            i += 1
        region = [ln for ln in lines[start:i] if '|' in ln]
        if len(region) < 2:
            continue

        # Find a header row: has a name column AND a role/cert/contact column
        for hdr in region:
            parts = [p.strip() for p in hdr.split('|')]
            name_cols = [ci for ci, p in enumerate(parts) if _NAME_HDR.search(p)]
            role_cols = [ci for ci, p in enumerate(parts) if _ROLE_HDR.search(p)]
            if not name_cols:
                continue
            if not (role_cols or any(_PHONE_HDR.search(p) or _ID_HDR.search(p)
                                    or _CERT_HDR.search(p) for p in parts)):
                continue
            name_col = name_cols[0]
            role_col = role_cols[0] if role_cols else None
            phone_col = next((ci for ci, p in enumerate(parts) if _PHONE_HDR.search(p)), None)
            id_col = next((ci for ci, p in enumerate(parts) if _ID_HDR.search(p)), None)

            hdr_idx = region.index(hdr)
            last_name = None  # merged-cell continuation: vertical merge leaves
            # empty name cells that inherit the row above (docx merged cells).
            for row in region[hdr_idx + 1:]:
                cells = [p.strip() for p in row.split('|')]
                if len(cells) <= name_col:
                    continue
                name = cells[name_col]
                # skip summary / continuation rows
                if name.startswith(('合计', '小计', '总计', '备注', '注：')):
                    last_name = None
                    continue
                if not name:
                    if last_name:
                        name = last_name  # inherit from previous row (merged cell)
                    else:
                        continue
                if not _is_person_name(name):
                    last_name = None
                    continue
                last_name = name
                role_str = cells[role_col] if role_col is not None and role_col < len(cells) else ''
                role = _infer_role_label(role_str) if role_str else 'team_member'
                info['all_persons'].append({'name': name, 'role': role, 'confidence': 0.75})
                if phone_col is not None and phone_col < len(cells):
                    for num in _iter_mobiles(cells[phone_col]):
                        _append_unique(info['phones'], num)
                        if not info.get('phone'):
                            info['phone'] = num
                if id_col is not None and id_col < len(cells):
                    im = re.search(r'\d{17}[\dXx]', cells[id_col].replace(' ', ''))
                    if im:
                        _append_unique(info['id_numbers'], im.group(0))
                        if not info.get('id_number'):
                            info['id_number'] = im.group(0)
            break  # one header per region is enough


def _append_unique(lst, value, cap=30):
    """Append value to lst if not already present (bounded)."""
    if value and value not in lst and len(lst) < cap:
        lst.append(value)


def _extract_from_signature_page(section_text, info):
    """Extract signatory names from signature/seal pages."""
    m = re.search(r'法定代表人或其委托代理人[：:][（(]?\s*(?:签字|签章|盖章|签名)\s*[）)]?\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿])', section_text)
    if m and _is_person_name(m.group(1).strip()):
        info['all_persons'].append({'name': m.group(1).strip(), 'role': 'signatory', 'confidence': 0.75})

    m = re.search(r'投标人[：:]\s*[（(]?(?:盖章|公章|单位章)[）)]?\s*(.{2,40}?)(?:\n|$)', section_text)
    if m and not info.get('company_name'):
        company = m.group(1).strip()
        # Reject form-layout captures like '法定代表人或授权代表: （签字' where
        # the 投标人 line is followed by the signature field, not the company.
        if len(company) >= 4 and not re.match(r'^[\s（(）)]+$', company) \
                and not re.search(r'(?:法定代表|授权代表|签字|盖章|单位章|公章|委托|日期|年月|^_{2,})', company):
            info['company_name'] = _clean_company(company)


def _extract_from_cover(section_text, info):
    """Extract company name and authorized rep from cover/bid letter."""
    if not info.get('company_name'):
        m = re.search(r'(?:投标人|供应商|申请.?|报价.?)[：:]\s*(.{2,40}?)(?:\n|$)', section_text)
        if m:
            company = m.group(1).strip()
            # Blank form fills like '投标人：____（盖单位章' are not companies
            if len(company) >= 4 and not re.search(r'(?:法定代表|授权代表|签字|盖章|单位章|公章|委托|日期|年月|^_{2,})', company):
                info['company_name'] = _clean_company(company)

    # Extract authorized rep from "签字代表（name、role）" in cover/bid letter
    if not info.get('authorized_rep'):
        m = re.search(r'签字代表[（(]([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})[、，]', section_text)
        if m:
            name = m.group(1).strip()
            if _is_person_name(name):
                info['authorized_rep'] = name
                info['all_persons'].append({'name': name, 'role': 'authorized_rep', 'confidence': 0.90})

    # Also try "签字代表：XXX" format
    if not info.get('authorized_rep'):
        m = re.search(r'签字代表[：:]\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿])', section_text)
        if m:
            name = m.group(1).strip()
            if _is_person_name(name):
                info['authorized_rep'] = name
                info['all_persons'].append({'name': name, 'role': 'authorized_rep', 'confidence': 0.85})

    # "授权代表：XXX" / "代理人：XXX" often appears in the bid letter itself
    # (not only in the authorization-letter section), e.g. the signature line
    # "授权代表（签字）：张三". Same label family as _extract_from_auth_section.
    if not info.get('authorized_rep'):
        m = re.search(r'(?:授权委托代理人|委托代理人|代理人|授权代表|被授权人|受托人|投标代表)[（(]?\s*(?:签字|签章|盖章|签名)?\s*[）)]?[：:]\s*([一-鿿](?:[ 	]*[一-鿿]){1,3}(?:[ 	]*[·•・][ 	]*[一-鿿](?:[ 	]*[一-鿿]){1,3}){0,2})(?![一-鿿])', section_text)
        if m:
            name = m.group(1).strip()
            if _is_person_name(name):
                info['authorized_rep'] = name
                info['all_persons'].append({'name': name, 'role': 'authorized_rep', 'confidence': 0.80})


def _infer_role_label(role_str):
    """Map Chinese role strings to standardized role labels."""
    role_str = role_str.strip()
    mapping = {
        '项目经理': 'project_manager', '项目负责人': 'project_manager',
        '项目副经理': 'project_manager', '商务经理': 'project_manager',
        '财务负责人': 'project_manager', '设计负责人': 'project_manager',
        '技术负责人': 'tech_lead', '技术总监': 'tech_lead', '总工程师': 'tech_lead',
        '安全负责人': 'tech_lead',
        '安全员': 'team_member', '质量员': 'team_member', '施工员': 'team_member',
        '材料员': 'team_member', '资料员': 'team_member', '造价员': 'team_member',
        '预算员': 'team_member',
    }
    for cn, en in mapping.items():
        if cn in role_str:
            return en
    return 'team_member'


def _clean_company(name):
    """Clean company name from parenthetical annotations / label fragments.

    PDF tabs and table cells can split a label across a boundary, leaving a
    detached open parenthesis (e.g. '北京某某大学 （投标人名称' where the
    closing '）' was consumed by a wider capture). Strip both complete labels
    and unterminated trailing fragments, plus the '（盖单位章）' style at the
    end of a cover line and PDF form blank-fill underscores.
    """
    if not name:
        return name
    name = re.sub(r'[（(]\s*(?:投标人名称|供应商名称|单位名称|公司名称|企业名称|单位负责人|'
                  r'盖单位章|盖章|公章|单位章|签章|签字|全称)\s*[）)]', '', name)
    # Detached label fragment ('（投标人名称' / '（盖单位章' with no close paren).
    name = re.sub(r'[（(][^）)]*$', '', name)
    name = re.sub(r'^[（(]|[）)]$', '', name).strip()
    # Leading label + form fill ('投标人：___北京某某大学___（盖单位章）').
    name = re.sub(r'^(?:投标人|供应商|招标人)(?:名称)?\s*[（(：:＿_\s]*|'
                  r'^(?:单位|公司|企业)名称\s*[（(：:＿_\s]*', '', name)
    name = re.sub(r'^[_\s＿]+', '', name).strip()
    name = re.sub(r'[_\s＿]+$', '', name).strip()
    return name


def _cleanup_name(info, key):
    """Clean up extracted person name."""
    val = info.get(key)
    if not val:
        return
    val = re.sub(r'^(?:本人\s*)+', '', val).strip()
    val = re.sub(r'^我(?=[一-鿿])', '', val)
    val = re.sub(r'\s*[（(](?:姓名|签字|盖章|单位负责人|法定代表人)[）)]\s*$', '', val)
    val = re.sub(r'^\s*[（(](?:姓名|签字|盖章|单位负责人|法定代表人)[）)]\s*', '', val)
    # A name may still carry extra whitespace from a PDF (single-char blocks).
    val = re.sub(r'\s+', '', val)
    if len(val) < 2 or any(w in val for w in ['注册', '签字', '盖章', '地址', '电话', '投标人', '姓名', '职务', '授权', '同志']):
        info[key] = None
    else:
        info[key] = val


def _clean_phone(v):
    """Strip junk around a captured phone number (leading '_' form-label
    fill, trailing punctuation) — defensive for OCR/form-wide captures."""
    if not v:
        return v
    v = re.sub(r'^[^\d]+', '', str(v).strip())
    v = re.sub(r'[^0-9\-]+$', '', v).strip()
    # PDF column padding inserts spaces inside a landline ('010 - 1234 5678');
    # keep the dash structure but remove the padding.
    v = re.sub(r'(?<=\d)[ \t]+(?=\d)', '', v)
    v = re.sub(r'(?<=\d)[ \t]*-[ \t]*(?=\d)', '-', v)
    return v


# ── PDF line-wrap gluing ──
# PDF text extraction frequently breaks a CJK word in the middle of a line
# (e.g. 法定代\n表人, 委托代\n理人, 投标\n报价). Downstream section detection and
# label-anchored regexes need the keyword contiguous, so re-join internal
# whitespace for a set of sensitive phrases. Whitespace BETWEEN distinct lines
# (not inside one of these phrases) is preserved, so section boundaries and
# table rows stay intact.
_GLUE_PHRASES = [
    # 人员 / 文档结构 / 角色
    '法定代表人身份证明', '法定代表人证明', '法定代表人', '单位负责人',
    '授权委托书', '授权委托', '委托书', '授权书', '授权代表', '授权代表人',
    '委托代理人', '被授权人', '受托人', '代理人', '签字代表', '投标代表',
    '签字', '签章', '盖章', '公章', '单位章', '盖单位章',
    '项目负责人', '项目经理', '技术负责人', '采购负责人', '项目副经理',
    '联系电话', '联系方式', '身份证号', '身份证',
    '供应商名称', '投标人名称', '单位名称', '公司名称', '企业名称',
    '招标人名称', '项目名称', '人员配备表', '技术标', '商务标', '资信标',
    '投标文件', '投标人',
    # 报价（同源拆分，防御性覆盖）
    '开标一览表', '报价一览表', '投标报价表', '分项报价表', '报价明细表',
    '分项明细表', '投标总价', '含税总价', '不含税总价', '投标函',
    '人民币', '人民币大写', '报价函', '报价单',
]


def _glue_phrases(text):
    """Re-join sensitive phrases that PDF extraction split across a line break.

    For every known phrase present in the text, remove any whitespace that has
    crept in BETWEEN its characters ('法\\n定\\n代\\n表\\n人' -> '法定代表人').
    A keyword already contiguous is a no-op; words absent from the text are
    skipped so the loop stays cheap on large documents.
    """
    if not text or len(text) < 2:
        return text
    flat = re.sub(r'\s+', '', text)
    for kw in _GLUE_PHRASES:
        if len(kw) < 2 or kw not in flat:
            continue
        # '\s*' between each character joins across any whitespace-only gap.
        pat = re.compile(r'\s*'.join(re.escape(c) for c in kw))
        text = pat.sub(kw, text)
    return text


def _normalize_cjk_whitespace(text):
    """Tighten whitespace PDF/OCR extraction introduces around CJK text.

    Covers the '投标人 ： 张三' / '（ 盖单位章 ）' style — spaces pushed between
    a label and its punctuation (PDF column padding, OCR spacing). Unicode
    space variants (\xa0, ideographic 　) become plain, and runs of 2+
    spaces collapse to one so pipe-table rows ('a | b') stay aligned.
    """
    text = text.replace(' ', ' ').replace('　', ' ')
    # Spaces before/after CJK punctuation: '投标人 ： 张三' -> '投标人： 张三'.
    text = re.sub(r'\s+([：:，,。；;、）)])', r'\1', text)
    # Spaces immediately after an opening paren: '（ 姓名）' -> '（姓名）'.
    text = re.sub(r'([（(])\s+', r'\1', text)
    text = re.sub(r' {2,}', ' ', text)
    return text


# Surname characters for the name-split re-join — a curated COMMON-surname
# string (the full 百家姓 includes rare single-char surnames like 国, which
# in '国 联系' style column-separated text would wrongly merge into a name).
# Only newline breaks are joined, never plain spaces: column padding would
# otherwise swallow the next label ('王强 联系电话').
_SURNAMES = '王李张刘陈杨黄赵周吴徐孙马胡朱郭何罗高林郑梁谢唐宋韩冯于董萧程曹袁邓许傅沈曾彭吕苏卢蒋蔡贾丁魏薛叶阎余潘杜戴夏钟汪田任姜范方石姚谭廖邹熊金陆郝孔白崔康毛邱秦江史顾侯邵孟龙万段雷钱汤尹易常武乔贺赖龚文'


def _join_split_names(text):
    """Re-join a person name split at a LINE BREAK right after the surname.

    '张某某' -> '张某某'. Two passes because re.sub does not rescan a
    replacement, so a name whose pieces re-join twice settles on the second
    pass. Deliberately newline-only: space-separated '张某某 联系' must not
    be re-glued (column padding would swallow label fragments).
    """
    pat = re.compile(rf'([{_SURNAMES}])\s*\n\s*([一-鿿]{{1,2}})')
    for _ in range(2):
        nxt = pat.sub(r'\1\2', text)
        if nxt == text:
            break
        text = nxt
    return text


# ── Invisible / control character defense ──
# PDF text layers emit zero-width spaces (\u200b), word joiners (\u2060),
# BOMs (\ufeff), soft hyphens (\u00ad) and exotic breaks (\r, \f, \v,
# \u2028/\u2029). '\s' does NOT match the zero-width family, so a single
# \u200b inside 法定代表人 defeats both the glue pass and every
# label-anchored regex. Normalize them before anything else runs.
def _strip_invisible(text):
    if not text:
        return text
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    text = re.sub(r'[\f\v\u2028\u2029]', '\n', text)
    return re.sub(r'[\u200b\u200c\u200d\u2060\ufeff\u00ad]', '', text)


# ── OCR near-form glyph repair (label words only) ──
# Scanned documents systematically confuse a few glyph pairs INSIDE labels:
# 人/入 (法定代表入), 话/活 (联系电活), 币/巾 (人民巾), 系/糸 (联糸).
# Every pattern below is a string that never occurs in legitimate Chinese,
# so a global replace cannot corrupt real document content. 身分证 is a
# common variant spelling rather than OCR noise, fixed the same way.
_OCR_LABEL_FIXES = (
    ('法定代表入', '法定代表人'),
    ('人民巾', '人民币'),
    ('委托入', '委托人'),
    ('电活', '电话'),
    ('联糸', '联系'),
    ('身分证', '身份证'),
)


def _fix_ocr_label_confusions(text):
    if not text:
        return text
    for bad, good in _OCR_LABEL_FIXES:
        if bad in text:
            text = text.replace(bad, good)
    return text


# ── Mobile number form coverage ──
# Bare '13912345678', country-code prefixed '+8613912345678' (the plain
# pattern's (?<!\d) lookbehind alone would reject it), and dash/space
# grouped '139-1234-5678' / '139 1234 5678' from formatted PDF cells.
_MOBILE_PATTERNS = (
    re.compile(r'(?<!\d)(?:\+?86)?1[3-9]\d{9}(?!\d)'),
    re.compile(r'(?<![\d-])(?:\+?86[- \t]?)?1[3-9]\d[- \t]?\d{4}[- \t]?\d{4}(?![\d-])'),
)


def _iter_mobiles(src):
    """Yield normalized 11-digit mobile numbers across PDF/OCR layouts."""
    out, seen = [], set()
    for pat in _MOBILE_PATTERNS:
        for match in pat.finditer(src or ''):
            digits = re.sub(r'\D', '', match.group(0))
            if len(digits) == 13 and digits.startswith('86'):
                digits = digits[2:]
            if re.fullmatch(r'1[3-9]\d{9}', digits) and digits not in seen:
                seen.add(digits)
                out.append(digits)
    return out


def extract_personnel(text):
    """Extract personnel information from bid text using chapter-scoped extraction.

    Only searches within specific sections: authorization letter, personnel table,
    qualification review, signature page, cover/bid letter.
    """
    info = {
        'legal_rep': None,
        'authorized_rep': None,
        'company_name': None,
        'id_number': None,
        'phone': None,
        'address': None,
        'response_date': None,
        'all_persons': [],
        # Multi-value contact pools for cross-file matching (a bid document
        # usually lists several phones / IDs / emails across volumes; matching
        # ANY shared value is far stronger evidence than the first one only).
        'phones': [],
        'id_numbers': [],
        'emails': [],
        # Bank account pool: a shared settlement account across bidders is
        # strong collusion evidence (资金往来同一账户). Label-anchored to 账号
        # to avoid harvesting project codes / contract numbers.
        'bank_accounts': [],
        'contacts': {'phone': None, 'email': None, 'address': None}
    }
    if not text:
        return info

    # Zero-width chars / exotic breaks would defeat every regex below; OCR
    # glyph confusions inside labels are repaired before any label matching.
    text = _strip_invisible(text)
    text = _fix_ocr_label_confusions(text)

    # Full-width digits break the ASCII-digit regexes ('１３９…' phones)
    text = _fw_digits_to_ascii(text)

    # Re-join keyword phrases that PDF extraction split across a line break
    # (法\n定\n代\n表\n人, 委托代\n理人 …). Kept as a first pass; the explicit
    # hardcoded re.subs below remain as a targeted fallback.
    text = _glue_phrases(text)

    # Normalize line breaks within key phrases
    text = re.sub(r'法定代\s*\n\s*表人', '法定代表人', text)
    text = re.sub(r'法定\s*\n\s*代表人', '法定代表人', text)
    text = re.sub(r'法\s*\n\s*定代表人', '法定代表人', text)
    text = re.sub(r'授权委\s*\n\s*托书', '授权委托书', text)
    text = re.sub(r'供应\s*\n\s*商名称', '供应商名称', text)
    # Fix common name splits: "张某某" → "张某某" (only when first char is a
    # surname; the source module-level constant covers the full 百家姓). Two
    # passes so a name broken at two whitespace boundaries settles on the 2nd.
    text = _join_split_names(text)

    # Tighten whitespace inserted between labels and CJK punctuation
    # ('投标人 ： 张三' / '（ 盖单位章 ）'), after the glues above so the
    # label-anchored regexes and section detection match.
    text = _normalize_cjk_whitespace(text)

    # ── Section Detection ──
    sections = _find_personnel_sections(text)

    # ── 1. Authorization Letter Section ──
    auth_sections = [s for s in sections if s['type'] in ('auth_letter', 'legal_rep_proof')]
    for sec in auth_sections:
        _extract_from_auth_section(sec['text'], info)

    # ── 2. Personnel Table Section ──
    personnel_sections = [s for s in sections if s['type'] in ('personnel_table', 'qualification')]
    for sec in personnel_sections:
        _extract_from_personnel_table(sec['text'], info)

    # ── 3. Signature Page Section ──
    sig_sections = [s for s in sections if s['type'] == 'signature_page']
    for sec in sig_sections:
        _extract_from_signature_page(sec['text'], info)

    # ── 4. Cover / Bid Letter Section ──
    cover_sections = [s for s in sections if s['type'] == 'cover_letter']
    for sec in cover_sections:
        _extract_from_cover(sec['text'], info)

    # ── 5. Pipe-table personnel (docx/xlsx tables, appended at text end) ──
    _parse_personnel_pipe_table(text, info)

    # ── 6. Section-less / empty-result fallback ──
    # Documents without any recognized section marker (short response letters,
    # OCR output, unusual structures) would otherwise yield zero personnel.
    # The label-anchored patterns are specific enough to run on the full text
    # as a last resort — also when section detection fired but the scoped
    # extraction STILL produced nothing (headers split across page breaks,
    # tables mis-detected as sections). Guarded on all of company/legal/auth
    # being empty, so a partial result is never overwritten or duplicated.
    if not (auth_sections or personnel_sections or sig_sections or cover_sections) \
            or (not info.get('company_name') and not info['legal_rep'] and not info['authorized_rep']):
        _extract_from_auth_section(text, info)
        _extract_from_cover(text, info)

    # ── Global extraction (section-scoped) ──
    for sec in auth_sections + sig_sections:
        m = re.search(r'身份证号[码字]?[：:]\s*(\d{17}[\dXx])', sec['text'])
        if m and not info.get('id_number'):
            info['id_number'] = m.group(1).strip()

    # ── Multi-value contact pools (phones / ID numbers / emails) ──
    # Scan the whole document: table columns and signature blocks often carry
    # contact info without section headers. IDs tolerate internal whitespace
    # ('3201 23 19…') which PDF extraction frequently inserts.
    for src in (text, _ocr_digit_normalize(text)):
        for num in _iter_mobiles(src):
            _append_unique(info['phones'], num)
    for m in re.finditer(r'(?<!\d)\d{6}(?:18|19|20)\d{2}'
                         r'(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])'
                         r'\d{3}[\dXx](?![\dXx])', text.replace(' ', '')):
        _append_unique(info['id_numbers'], m.group(0))
    for m in re.finditer(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', text):
        _append_unique(info['emails'], m.group(0).lower())
    # Bank accounts: label-anchored '账号：1100923…' (9-25 digits). Whitespace
    # tolerated inside the number (PDF cell wrapping). Exclude tender-side
    # payment instructions (保证金汇入账号, 标书款/工本费收款账户): those are
    # reprinted in EVERY bidder's document and would falsely pair all files.
    # The bidder's own 基本户 account (开户行：X / 账号：Y in qualification
    # pages) is the collusion-relevant signal and has no such context.
    for m in re.finditer(r'[账帐]\s*户?\s*号[：:\s]+\d(?:[\d\s]{7,27}\d)?', text):
        prefix = text[max(0, m.start() - 40):m.start()]
        if re.search(r'(?:保证金|投标保证|汇[入至款]|缴[纳付交]|招标|代理|'
                     r'工本费|标书[款费]|平台|收费)', prefix):
            continue
        acct = re.sub(r'\D', '', m.group(0))
        if 9 <= len(acct) <= 25:
            _append_unique(info['bank_accounts'], acct)
    if info['id_numbers'] and not info.get('id_number'):
        info['id_number'] = info['id_numbers'][0]
    if info['phones'] and not info.get('phone'):
        info['phone'] = info['phones'][0]

    # Cover/bid-letter block (投标函) commonly carries the bidder's own
    # 地址/电话/传真 lines ("投标人：X（盖单位章）…电话：010-12345678") —
    # include cover_sections so a landline there is not missed. The capture
    # tolerates the PDF column padding inside a landline ('010 - 1234 5678').
    for sec in auth_sections + personnel_sections + sig_sections + cover_sections:
        m = re.search(r'(?:电话|手机|联系电话|联系方式|移动电话|手机号码|电话号码)[：:]\s*(\d[\d\- \t]{6,19})', sec['text'])
        if m:
            phone = _clean_phone(m.group(1))
            if not info.get('phone'):
                info['phone'] = phone
            if not info['contacts'].get('phone'):
                info['contacts']['phone'] = phone

    # Section-less documents (short response letters, OCR dumps) never enter
    # the loop above — run the same label-anchored scan on the full text so a
    # landline there is not lost, mirroring the step-6 name-rescan fallback.
    if not (auth_sections or personnel_sections or sig_sections or cover_sections):
        m = re.search(r'(?:电话|手机|联系电话|联系方式|移动电话|手机号码|电话号码)[：:]\s*(\d[\d\- \t]{6,19})', text)
        if m:
            phone = _clean_phone(m.group(1))
            if not info.get('phone'):
                info['phone'] = phone
            if not info['contacts'].get('phone'):
                info['contacts']['phone'] = phone

    for sec in auth_sections:
        m = re.search(r'地址[：:]\s*(.{8,80})', sec['text'])
        if m and not info.get('address'):
            addr = m.group(1).strip()
            # Cover blocks put the next field on the SAME line — cut the
            # address at the first following label so it does not swallow
            # '电话：…' / '联系人：…' fragments.
            addr = re.split(r'\s*(?:电话|手机|联系电话|联系方式|传真|邮编|邮政编码|联系人|'
                            r'电子邮箱|邮箱|开户行|开户银行|账号|账户|网址)', addr, 1)[0][:100]
            if len(addr) >= 8:
                info['address'] = addr
                info['contacts']['address'] = addr

    # Response date (can be anywhere near top of document)
    m = re.search(r'(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', text[:800])
    if m:
        info['response_date'] = m.group(1).strip()

    # Cleanup names
    _cleanup_name(info, 'legal_rep')
    _cleanup_name(info, 'authorized_rep')

    # Deduplicate all_persons
    seen = set()
    unique_persons = []
    for p in info['all_persons']:
        # CJK names: strip PDF-inserted internal spaces ('张 三') so the same
        # person matches across files; foreign names keep their spaces.
        if p['name'] and not re.search(r'[A-Za-z]', p['name']):
            p['name'] = re.sub(r'\s+', '', p['name'])
        key = (p['name'], p['role'])
        if key not in seen:
            seen.add(key)
            unique_persons.append(p)
    info['all_persons'] = unique_persons

    return info


def _demote_environmental_pool_values(all_personnel, group_names,
                                      min_groups=3, ratio=0.8):
    """Drop contact-pool values that appear in nearly EVERY document.

    Tender-side data (招标代理联系电话, 保证金收款账号, platform notification
    emails) is reprinted inside every bidder's document. A value shared by
    ALL bidders cannot discriminate collusion — it only generates pairwise
    false positives ("联系电话相同" for every file pair). A value must appear
    in >= min_groups documents AND >= ratio of all documents to be treated
    as environmental.

    Filters all_personnel in place; returns the removed values (for logging
    and tests). With fewer than min_groups documents nothing is removed:
    with 2 files a shared phone is still genuine evidence.
    """
    n = len(group_names)
    if n < min_groups:
        return set()
    removed = set()
    for pool_key in ('phones', 'id_numbers', 'emails', 'bank_accounts'):
        counts = {}
        for gn in group_names:
            for v in set(all_personnel.get(gn, {}).get(pool_key) or []):
                counts[v] = counts.get(v, 0) + 1
        env = {v for v, c in counts.items() if c >= min_groups and c / n >= ratio}
        if not env:
            continue
        for gn in group_names:
            pool = all_personnel.get(gn, {}).get(pool_key)
            if pool:
                all_personnel[gn][pool_key] = [v for v in pool if v not in env]
        removed |= env
    if removed:
        logger.info('环境噪声降权：%d 项跨全部标书普遍出现的联系方式已从交叉比对剔除', len(removed))
    return removed


# ── Price Extraction ────────────────────────────────────────────
# Chinese uppercase numerals (used by both _parse_amount and other price helpers).
# Includes formal (壹贰叁...), informal (一二三...) and the place-value units
# (拾佰仟万亿). 角/分 are sub-yuan fractions kept for completeness.
_CN_NUM = {'零': 0, '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5,
           '陆': 6, '柒': 7, '捌': 8, '玖': 9, '一': 1,
           '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
           '七': 7, '八': 8, '九': 9}
# Place-value units: 拾/拾 = x10, 佰/百 = x100, 仟/千 = x1000,
# 万 = x10000 (also a section multiplier), 亿 = x100000000 (section multiplier).
_CN_UNIT = {'拾': 10, '十': 10, '佰': 100, '百': 100, '仟': 1000, '千': 1000}
_CN_SECTION = {'万': 10000, '亿': 100000000}
_CN_FRAC = {'角': 0.1, '分': 0.01, '毛': 0.1}


def _cn_to_number(text):
    """Parse a Chinese-numeral amount string into a float.

    Handles full uppercase forms such as '壹仟贰佰叁拾肆万伍仟陆佰柒拾捌元整',
    informal forms ('一千二百三十四万'), and mixed forms with 元/角/分/整 suffixes.
    Returns None if the string contains no parseable Chinese numeral.

    Algorithm: section-based accumulation. Within each section (delimited by
    万/亿), digits (零-玖) are summed and units (拾/佰/仟) multiply the current
    digit then fold into the section sum. A 万 section multiplies its section
    sum by 1e4 and adds to the grand total; an 亿 section multiplies the grand
    total (plus its own section) by 1e8 -- correctly handling nesting like
    '壹亿贰仟万' = 1e8 + 2e3*1e4. 角/分 are sub-yuan fractions.
    """
    if not text:
        return None
    # Strip '整/正' (no numeric value) but KEEP 元/圆 as a section terminator
    # so the integer part before 角/分 is folded correctly.
    s = re.sub(r'[整正]', '', text)

    total = 0.0      # grand total across 亿/万 sections
    section = 0.0    # current section value (before 万/亿 multiplier)
    num = 0.0        # current pending digit within a section
    frac_val = 0.0   # accumulated 角/分 sub-yuan value
    had_digit = False

    for ch in s:
        if ch in _CN_NUM:
            num += _CN_NUM[ch]
            had_digit = True
        elif ch in _CN_UNIT:
            unit = _CN_UNIT[ch]
            # 拾 without a leading digit means 1 (拾贰 = 12, i.e. 1*10 + 2)
            if num == 0:
                num = 1
            section += num * unit
            num = 0
        elif ch in _CN_SECTION:
            mult = _CN_SECTION[ch]
            section += num
            num = 0
            if mult == 100000000:  # 亿 multiplies everything seen so far
                total = (total + section) * mult
            else:                   # 万 multiplies only its own section
                total += section * mult
            section = 0
        elif ch in ('元', '圆'):
            # End of integer part: fold the pending section into the total.
            total += section + num
            section = 0
            num = 0
        elif ch in _CN_FRAC:
            # 角/分 use the immediately preceding pending digit (default 0).
            frac_val += (num if num else 0) * _CN_FRAC[ch]
            num = 0
        # other characters are ignored
    total += section + num

    if not had_digit and frac_val == 0:
        return None
    return total + frac_val


def _parse_amount(s):
    """Parse a price string to float.

    Supports: '1,234,567.89', '123.45万元', '1.2亿元',
    pure Chinese uppercase '壹佰贰拾叁万' / '壹仟贰佰叁拾肆万伍仟陆佰柒拾捌元整',
    and informal '一千二百三十四'. Falls back to the first Arabic numeral found.
    """
    if not s:
        return 0.0
    s = str(s).replace(',', '').replace('，', '').strip()

    # Handle full-width digits up front so they participate in both paths.
    _FW = '０１２３４５６７８９'
    s = s.translate(str.maketrans(_FW, '0123456789'))

    # OCR digit-letter confusions in otherwise-numeric strings
    # ('￥1O9,800元' → 109800). Only fires when the string is composed
    # solely of digits/letters/separators, so real words are untouched.
    if re.search(r'\d', s) and re.fullmatch(r'[\dOoIl|.\s]+', s):
        s = s.translate(str.maketrans('OoIl|', '00111'))

    # Thousands thin-space groups from PDF ('1 261 819.76'). A space is only
    # joined when followed by exactly 3 digits (plus optional decimal tail),
    # so two space-separated column numbers ('1838529 5002800') never merge.
    s = re.sub(r'(?<=\d)[ \t](?=\d{3}(?:\.\d+)?(?!\d))', '', s)

    has_cn = any(ch in _CN_NUM or ch in _CN_UNIT or ch in _CN_SECTION for ch in s)

    if has_cn:
        # First try a true Chinese-numeral parse (covers pure-大写 amounts,
        # which are extremely common in bid letters and were previously lost).
        cn_val = _cn_to_number(s)
        if cn_val is not None and cn_val > 0:
            return float(cn_val)
        # Fallback: an Arabic numeral embedded in a Chinese-numeral context
        # (e.g. '1.2万', '叁万' where the parser could not assemble).
        m = re.search(r'([\d]+\.?\d*)', s)
        if m:
            val = float(m.group(1))
            if '亿' in s:
                val *= 100000000
            elif '万' in s:
                val *= 10000
            return val

    # Unit multiplier detection
    wan = 1.0
    if '亿元' in s:
        wan = 100000000
        s = s.replace('亿元', '')
    elif s.endswith('亿') or '亿 ' in s or ' 亿' in s:
        wan = 100000000
        s = s.replace('亿', '')
    elif '万元' in s:
        wan = 10000
        s = s.replace('万元', '')
    elif s.endswith('万') or '万 ' in s or ' 万' in s:
        wan = 10000
        s = s.replace('万', '')

    # Extract first numeric value (handle price ranges: take first value)
    m = re.search(r'([\d]+\.?\d*)', s)
    return float(m.group(1)) * wan if m else 0.0

# ── Structured Price Extraction ─────────────────────────────────
# Arabic amount with optional 万元/亿/元 suffix captured into the group, so
# '12.5万元' / '126181976.30元' keep their magnitude through _parse_amount.
# (A bare [\d,]+\.?\d* alternation would stop before the unit and lose x10000.)
# '(?: \d{3})*' tolerates thin-space thousands groups ('1 261 819.76') that
# PDF extraction emits instead of commas — the group shape (exactly 3 digits
# after each space) keeps two space-separated column numbers apart.
_AMT_ARABIC = r'[\d,]+(?: \d{3})*\.?\d*\s*(?:万|亿)?\s*元?'
# Full Chinese uppercase / informal numeral string (incl. 元/角/分/整).
_AMT_CN = r'[壹贰叁肆伍陆柒捌玖拾佰仟万亿零一二三四五六七八九十百千元整角分圆]+'
_AMT = r'(?:' + _AMT_ARABIC + r'|' + _AMT_CN + r')'
# Contexts that are amounts but NOT the bidder's own price. Shared by the
# price channels and post-validation so an amount is judged consistently:
# bid bonds / deposits, reference contract amounts, document prices, capital,
# and tender-side ceilings (最高限价/控制价/预算) that every bidder reprints
# with the SAME value — capturing one of those as "the bid price" would make
# all bidders look identical.
_NON_BID_AMOUNT_CTX = (r'(?:保证金|投标保证|押金|投标保函|银行保函|履约保证|'
                       r'合同金额|合同价款|签约合同价|中标金额|结算金额|成交金额|'
                       r'售价|注册[资]*金|出资|暂列金额|'
                       r'最高限价|招标控制价|控制价|拦标价|暂估价|预算[金额财]*[额为]?|'
                       r'违约金|赔偿金|罚金|代理服务费|中标服务费|交易服务费|'
                       r'平台使用费|工本费|标书款|手续费|佣金)')

# Bank payment-voucher markers. Unlike the label list above these never sit
# on the amount's own line (PDF extraction puts each table cell on its own
# line), so they need a WINDOW check around the amount, not a line-prefix one.
_PAYMENT_VOUCHER_CTX = re.compile(
    r'(?:账号|开户行|开户银行|电汇|汇款|转账|缴款|制单|回单|贷记|银行流水|'
    r'收款人|付款人|出票|票据|凭证|结算凭证|业务回执|承兑|汇票|支票|本票|网银)')


def _amount_in_non_bid_context(search_text, pos):
    """True when the amount at `pos` is in a non-bid-price context.

    Judged on the CURRENT LINE's prefix only: a tender-ceiling line
    ('最高限价：￥1,000,000元') frequently sits right above the real bid
    price, and a wide look-behind window would let its label bleed into the
    real price's context and exclude BOTH lines. Label and amount are on the
    same visual line in essentially all extracted layouts (PDF/docx/plain).
    """
    line_start = search_text.rfind('\n', 0, pos) + 1
    return bool(re.search(_NON_BID_AMOUNT_CTX, search_text[line_start:pos]))


def _amount_in_payment_voucher(search_text, pos, window=160):
    """True when the amount at `pos` sits inside a bank payment voucher.

    Bid-bond transfer slips print the amount as '金额 / 人民币：22,000.00 /
    人民币：贰万贰仟元整 / 用途 / 制单日期…' — the banking words live on
    NEIGHBOURING lines (one cell per line in PDF extraction), so the
    line-prefix check above never sees them. A bounded window around the
    amount catches them; the marker words do not occur near real bid totals.
    """
    lo = max(0, pos - window)
    hi = min(len(search_text), pos + window)
    return bool(_PAYMENT_VOUCHER_CTX.search(search_text, lo, hi))


def _amount_is_unit_rate(search_text, end):
    """True when the amount ending at `end` is a UNIT RATE, not a total.

    Service bids print per-month / per-person rates ('22000元/月',
    '340000/人月', '1200元/次') in the same summary tables; the amount
    patterns drop the denominator and the rate would pose as the bid total.
    Any slash directly after the amount (or 元 + 人月/每…) marks a rate.
    """
    tail = search_text[end:end + 12]
    return bool(re.match(r'[ \t]*(?:元[ \t]*)?[/／]|'
                         r'[ \t]*元[ \t]*(?:人月|每)', tail))


def _is_yyyymmdd(v):
    """True when the number parses as a calendar date YYYYMMDD (1990-2099).

    Digit runs like 20250826 appear in flattened tables as 签订/生效 dates
    and must never be read as amounts."""
    s = str(int(v))
    if len(s) != 8:
        return False
    y, mo, d = int(s[:4]), int(s[4:6]), int(s[6:8])
    return 1990 <= y <= 2099 and 1 <= mo <= 12 and 1 <= d <= 31


def _fw_digits_to_ascii(s):
    """Translate full-width digits (０-９), ：，％＠ to ASCII in a text copy.

    PDF text layers occasionally emit full-width digits ('小写：１２３４５'),
    which the ASCII-digit regexes would silently skip. ＠ joins so a
    full-width email address still matches the ASCII email pattern."""
    return s.translate(str.maketrans('０１２３４５６７８９：，％％＠', '0123456789:,%%@'))


def _ocr_digit_normalize(s):
    """Best-effort repair of OCR digit confusions in digit-heavy strings:
    collapse spaces between digits ('139 1234 5678') and map common
    letter/digit confusions (O/Q→0, l/I/|→1) only when directly between
    digits. Used for contact-pool scanning; the primary text is never
    rewritten with this."""
    s = re.sub(r'(?<=\d)[ \t](?=\d)', '', s)
    s = re.sub(r'(?<=\d)[OoQ](?=\d)', '0', s)
    s = re.sub(r'(?<=\d)[lI|](?=\d)', '1', s)
    return s


def extract_prices(text):
    """Extract structured pricing using multi-channel pipeline with confidence scoring.
    Returns dict with: totalPriceInTax, totalPrice, taxRate, revenue, cost,
    subItemPrice[], costDetails[]"""
    result = {
        'totalPriceInTax': None,
        'totalPrice': None,
        'taxRate': None,
        'bidRate': None,
        'revenue': None,
        'cost': None,
        'subItemPrice': [],
        'costDetails': [],
        # Provenance notes for the UI / report: consistency fixes applied,
        # low-confidence provenance, sub-item sum mismatches. Populated by
        # _validate_price_extraction.
        'warnings': []
    }
    if not text:
        return result
    text = _strip_invisible(text)
    text = _fix_ocr_label_confusions(text)
    text = _fw_digits_to_ascii(text)
    text = _glue_phrases(text)
    text = _normalize_cjk_whitespace(text)

    # Track whether the accepted price came from Chinese numerals only (no
    # Arabic digits in the source). Validation then skips the "value must
    # appear in text" check, which pure-大写 documents can never satisfy.
    def _mark_cn_only(group):
        if group and not re.search(r'\d', group) \
                and re.search(r'[壹贰叁肆伍陆柒捌玖拾佰仟万亿零一二三四五六七八九十百千]', group):
            result['_cn_only'] = True

    # Full-text fallback hits need extra validation even when a bid section
    # exists (a 合同金额/售价 line may precede the real price in the document).
    # NOTE: provenance is an explicit flag, NOT `search_text is text` — a
    # full-range str slice (section == whole text, common for short docs)
    # returns the SAME object in CPython, which would mislabel a bid-section
    # hit as a global one.
    def _mark_global():
        result['_from_global'] = True

    # ── Channel 0: Bid summary section FIRST (开标一览表/投标报价表) ──
    # This is the MOST RELIABLE source for total price. Run it before global
    # text search to avoid matching bid bonds, deposits, or other small amounts.
    bid_section = _find_bid_summary_section(text)

    # ── Bond-amount echo set ──
    # The bid bond value echoes through the document: labeled once in the
    # summary table ('投标保证金 2.2万元' or '投标保证金（大写）：贰万元整') and
    # again, UNLABELED, on the bank transfer slip ('金额 / 人民币:22,000.00 /
    # 人民币:贰万贰仟元整'). Collect every explicitly bond-labeled amount (Arabic
    # or Chinese-numeral form) so unlabeled echoes can be skipped anywhere.
    bond_amounts = set()
    for bm in re.finditer(
            r'(?:投标保证金|履约保证金|投标保函|保证金|押金)[\s\S]{0,30}?'
            r'([\d,]+(?:\.\d{1,2})?\s*万元?|[壹贰叁肆伍陆柒捌玖拾佰仟万亿零]{2,20})',
            text):
        val = _parse_amount(bm.group(1))
        if val >= 100:
            bond_amounts.add(round(val, 2))

    # ── Channel 1: Symbol-based (￥/¥/CNY/RMB) ── confidence: 0.95
    # Search bid_section first (most reliable), then fall back to the full
    # text: a bid-summary keyword occurrence may start AFTER the actual price
    # line (e.g. '投标函' matched mid-letter), so narrowing the search to the
    # section alone would lose the price.
    symbol_patterns = [
        # 'RMB￥：126181976.30元' - combined prefix+symbol, optional colon
        r'(?:CNY|RMB)\s*[￥¥]?\s*[：:]?\s*(' + _AMT_ARABIC + r')',
        # '￥12.5万元' keeps its 万元 magnitude via _AMT_ARABIC suffix
        r'[￥¥]\s*[：:]?\s*(' + _AMT + r')',
        # 人民币 and its amount must stay on ONE line: with '\s+' the label
        # at the end of '2000万元人民币' grabs the next line's '2014年1月20日'
        # (a founding date), and the later-nulled value blocks all lower
        # channels from retrying.
        r'人民币[：:]?[ \t]*(' + _AMT + r')',
        r'(?:CNY|RMB)\s*(' + _AMT_ARABIC + r')',
        r'USD\s*([\d,]+\.?\d*)',
    ]
    sources = ([(bid_section, False)] if bid_section else []) + [(text, True)]
    for search_text, _from_global_hit in sources:
        if result['totalPriceInTax'] is not None:
            break
        for pat in symbol_patterns:
            if result['totalPriceInTax'] is not None:
                break
            # Iterate ALL matches: an excluded context (保证金/最高限价…)
            # earlier in the text must not hide the real price further down.
            for m in re.finditer(pat, search_text):
                # Skip non-bid contexts: bid bonds ('投标保证金...￥:500000元'),
                # reference contract amounts ('合同金额：RMB2080000'), doc prices
                # ('招标文件售价：人民币1000元'), tender ceilings ('最高限价…')
                if _amount_in_non_bid_context(search_text, m.start()):
                    continue
                # Skip bank payment vouchers (bond transfer slips) and any
                # echo of an explicitly bond-labeled amount.
                if _amount_in_payment_voucher(search_text, m.start()):
                    continue
                # A digit run directly followed by 年 is a date year
                # ('2014年1月20日'), never an amount.
                if re.match(r'[ \t]*年', search_text[m.end():]):
                    continue
                # Unit rates ('22000元/月', '340000/人月') are not totals.
                if _amount_is_unit_rate(search_text, m.end()):
                    continue
                val = _parse_amount(m.group(1))
                if any(abs(val - b) < 1 for b in bond_amounts):
                    continue
                if val >= 100:
                    result['totalPriceInTax'] = val
                    result['totalPrice'] = val
                    _mark_cn_only(m.group(1))
                    if _from_global_hit:
                        _mark_global()
                    break

    # ── Channel 2: Label-based (标签通道) ── confidence: 0.90
    if result['totalPriceInTax'] is None:
        label_patterns = [
            # Same-line 人民币 prefix (see Channel 1 note on the newline grab)
            r'人民币[：:]?[ \t]*(' + _AMT + r')',
            r'小写[（(]?\s*[:：]?\s*[）)]?\s*(' + _AMT_ARABIC + r')',
            r'(?:投标总价|投标总报价|总报价|报价金额|投标报价|项目总价|投标总金额|总金额|'
            r'最终报价|首轮报价|响应报价|谈判报价|含税总报价|投标金额|响应文件总价|'
            r'首次报价|最终投标报价|投标总价格)[：:\s]+(' + _AMT + r')',
            # Label with unit in parentheses: "投标总价（元）：123000" /
            # "合计（万元）：89.3" — the 万元 suffix keeps magnitude via _AMT.
            r'(?:投标总价|投标总报价|总报价|报价金额|投标报价|项目总价|总金额|总价|总计|合计)'
            r'[（(]\s*(?:万元?|元)\s*[）)]\s*[：:]?\s*(' + _AMT + r')',
            r'(?:总价|总计|合计)[：:]\s*(' + _AMT + r')',
            r'(?:金额|报价)[（(]元[）)][：:]\s*([\d,]+\.?\d*)',
        ]
        for search_text, _from_global_hit in sources:
            if result['totalPriceInTax'] is not None:
                break
            for pat in label_patterns:
                if result['totalPriceInTax'] is not None:
                    break
                for m in re.finditer(pat, search_text):
                    if _amount_in_non_bid_context(search_text, m.start()):
                        continue
                    if _amount_in_payment_voucher(search_text, m.start()):
                        continue
                    if re.match(r'[ \t]*年', search_text[m.end():]):
                        continue
                    if _amount_is_unit_rate(search_text, m.end()):
                        continue
                    val = _parse_amount(m.group(1))
                    if any(abs(val - b) < 1 for b in bond_amounts):
                        continue
                    if val >= 100:
                        result['totalPriceInTax'] = val
                        result['totalPrice'] = val
                        _mark_cn_only(m.group(1))
                        if _from_global_hit:
                            _mark_global()
                        break

    # ── Channel 3: 大写/小写 pair ── confidence: 0.88
    if result['totalPriceInTax'] is None:
        for search_text, _from_global_hit in sources:
            if result['totalPriceInTax'] is not None:
                break
            m = re.search(r'大写[：:]?\s*[（(]?\s*' + _AMT_CN + r'\s*[）)]?[\s\S]{0,100}?'
                          r'小写[：:]?\s*[（(]?\s*([\d,]+\.?\d*)', search_text)
            if not m:
                # Reversed order: 小写 first, 大写 after
                m = re.search(r'小写[：:]?\s*[（(]?\s*([\d,]+\.?\d*)\s*[）)]?[\s\S]{0,100}?'
                              r'大写[：:]?\s*[（(]?\s*' + _AMT_CN, search_text)
            if m:
                val = _parse_amount(m.group(1))
                if val >= 100:
                    result['totalPriceInTax'] = val
                    result['totalPrice'] = val
                    if _from_global_hit:
                        _mark_global()
                    break

    # ── Channel 3b: standalone 大写 amount (no 小写 pair) ── confidence: 0.82
    # '人民币（大写）：壹亿贰仟陆佰壹拾捌万壹仟玖佰柒拾陆元叁角' - parse the
    # Chinese numerals directly. Must precede a lone arabic-number fallback so
    # the magnitude survives even when the 小写 line was lost in extraction.
    if result['totalPriceInTax'] is None:
        for search_text, _from_global_hit in sources:
            if result['totalPriceInTax'] is not None:
                break
            m = re.search(r'(?:大写|人民币\s*[（(]\s*大写\s*[）)])[：:]?\s*[（(]?\s*(' + _AMT_CN + r')', search_text)
            if m:
                val = _parse_amount(m.group(1))
                if val >= 1000:
                    result['totalPriceInTax'] = val
                    result['totalPrice'] = val
                    # Pure-CN provenance: validation must NOT require the
                    # Arabic digits to appear in the text (they don't exist
                    # in a 大写-only document).
                    result['_cn_only'] = True
                    if _from_global_hit:
                        _mark_global()
                    break

    # ── Channel 4: Table-based (detailed search in bid section) ── confidence: 0.85
    if bid_section and result['totalPriceInTax'] is None:
        for pat in [
            r'(?:CNY|RMB)\s*[￥¥]?\s*[：:]?\s*(' + _AMT_ARABIC + r')',
            r'[￥¥]\s*[：:]?\s*(' + _AMT + r')',
            r'人民币[：:\s]+(' + _AMT + r')',
            r'小写[：:]?\s*(' + _AMT_ARABIC + r')',
            r'(?:总价|总计|合计|报价)[：:]?\s*(' + _AMT + r')',
        ]:
            if result['totalPriceInTax'] is not None:
                break
            for m in re.finditer(pat, bid_section):
                # Same non-bid exclusion as channels 1/2 — without it a
                # 最高限价 line inside the bid section would be captured here.
                if _amount_in_non_bid_context(bid_section, m.start()):
                    continue
                val = _parse_amount(m.group(1))
                if val >= 100:
                    result['totalPriceInTax'] = val
                    result['totalPrice'] = val
                    _mark_cn_only(m.group(1))
                    break

    # ── Channel 5: Docx pipe-separated total (总计 | 893000) ── confidence: 0.83
    if result['totalPriceInTax'] is None:
        # 合计 row with | separators: look for a 合计 row where one of the last
        # columns has a large number (>= 10000, to avoid matching small sub-totals)
        for m in re.finditer(r'合计\s*\|.+', text):
            row = m.group()
            # Extract all pipe parts; parse each with _parse_amount so cells
            # like '89.3万元' keep their magnitude (plain [\d,]+ would read 89.3).
            parts = [p.strip() for p in row.split('|')]
            nums = []
            for p in parts:
                if not re.search(r'\d', p):
                    continue
                val = _parse_amount(p)
                if val > 0:
                    nums.append(val)
            # The total summary row has large numbers (>= 10000) in the last columns
            large_nums = [n for n in nums if n >= 10000]
            if len(large_nums) >= 2:
                # Last two large numbers are typically 不含税总价 and 含税总价
                result['totalPrice'] = large_nums[-2]
                result['totalPriceInTax'] = large_nums[-1]
                break
            elif len(large_nums) == 1 and large_nums[0] >= 50000:
                result['totalPrice'] = large_nums[0]
                result['totalPriceInTax'] = large_nums[0]
                break

        # Fallback: "合计 NNN" space-separated (common in PDF tables without colons).
        # PDF text layers often render empty cells as '\' or '/', so the row
        # '合计 \ 1838529 5002800 \' must be accepted: separators [\/ ] between
        # 合计 and the first number, and an optional second price column
        # (不含税 / 含税 pair).
        if result['totalPriceInTax'] is None:
            _SEP = r'(?:\s|\\|/|、)*'
            for m in re.finditer(r'^合\s*计\s*' + _SEP + r'(\d{5,12}(?:\.\d{1,2})?)'
                                 r'(?:' + _SEP + r'(\d{5,12}(?:\.\d{1,2})?))?'
                                 r'(?:' + _SEP + r'(\d{1,2})(?=\s|$))?', text, re.MULTILINE):
                val = _parse_amount(m.group(1))
                if val >= 50000:
                    v2 = float(m.group(2).replace(',', '')) if m.group(2) else None
                    if v2 is not None and v2 >= 50000:
                        lo, hi = min(val, v2), max(val, v2)
                        if hi / max(lo, 1) <= 1.5:
                            # plausibly 不含税 + 含税 pair
                            result['totalPrice'] = lo
                            result['totalPriceInTax'] = hi
                        else:
                            result['totalPrice'] = val
                            result['totalPriceInTax'] = v2
                    else:
                        result['totalPriceInTax'] = val
                        result['totalPrice'] = val
                    if m.group(3) and 1 <= int(m.group(3)) <= 30:
                        result['taxRate'] = m.group(3) + '%'
                    break

        # Fallback: simple "合计 NNN" anywhere (no colons, just space)
        if result['totalPriceInTax'] is None:
            m = re.search(r'(?:合计|总计)\s+(\d{5,12}(?:\.\d{2})?)', text)
            if m:
                val = _parse_amount(m.group(1))
                if val >= 50000:
                    result['totalPriceInTax'] = val
                    result['totalPrice'] = val

        # Fallback: simple "总计 | number" pattern
        if result['totalPrice'] is None:
            m = re.search(r'总计\s*\|\s*(\d{4,10}(?:\.\d{2})?)', text)
            if m:
                val = _parse_amount(m.group(1))
                if val >= 100:
                    result['totalPrice'] = val
                    result['totalPriceInTax'] = val

        # Fallback: "不含税总价：" / "含税总价：" labels
        if result['totalPrice'] is None:
            m = re.search(r'不含税总价[：:]\s*([\d,]+\.?\d*)', text)
            if m:
                val = _parse_amount(m.group(1))
                if val >= 100:
                    result['totalPrice'] = val
        if result['totalPriceInTax'] is None:
            m = re.search(r'含税总价[：:]\s*([\d,]+\.?\d*)', text)
            if m:
                val = _parse_amount(m.group(1))
                if val >= 100:
                    result['totalPriceInTax'] = val

        # Everything in Channel 5 is a full-text search; mark for validation.
        if result.get('totalPriceInTax') is not None:
            result['_from_global'] = True

    # ── Tax rate decomposition ──
    _extract_tax_decomposition(text, bid_section if bid_section else text, result)

    # ── Derive 不含税 from 含税 + 税率 when tax-exclusive is missing ──
    tp_val = result.get('totalPrice')
    tpit_val = result.get('totalPriceInTax')
    tax_str = result.get('taxRate')
    if tpit_val is not None and tax_str is not None:
        m = re.search(r'(\d+(?:\.\d+)?)', tax_str)
        if m:
            rate = float(m.group(1)) / 100.0
            if rate > 0:
                derived_tp = round(tpit_val / (1 + rate), 2)
                if tp_val is None or tp_val == tpit_val:
                    result['totalPrice'] = derived_tp

    # ── Rate-based quotes (费率/下浮率) ──
    # Service bids (asset evaluation, supervision, agency) often quote a rate
    # instead of an amount: '下浮率：12.5%', '报价费率 1.8‰'. Extracted as an
    # informational field so price analysis can still cluster rate-based bids.
    if result.get('bidRate') is None:
        search_text = bid_section if bid_section else text
        for pat in [
            # label filler must not eat the digits: [^…\d]* stops at numbers
            r'(?:下浮率|下浮比例|下浮幅度)[（(]?[^）)：:\d]*[）)]?\s*[：:]?\s*([\d.]+)\s*([%％‰])',
            r'(?:投标|报价|评审)?费率[（(]?[^）)：:\d]*[）)]?\s*[：:]?\s*([\d.]+)\s*([%％‰])',
            r'投标报价\s*下浮\s*([\d.]+)\s*([%％])',
            # Percent-quote format: "投标报价（%）：98.5" (price as % of 控制价)
            r'投标报价\s*[（(]\s*[%％]\s*[）)]\s*[：:]?\s*([\d.]+)',
            # Chinese-style discount wording: "下浮 6 个百分点"
            r'(?:下浮|费率)\s*([\d.]+)\s*个?百分点',
        ]:
            m = re.search(pat, search_text)
            if m:
                val = float(m.group(1))
                unit = m.group(2) if m.lastindex and m.lastindex >= 2 else None
                if not 0 < val < 100:
                    continue
                if unit is None:
                    result['bidRate'] = m.group(1) + '%'
                else:
                    result['bidRate'] = m.group(1) + ('%' if unit in ('%', '％') else '‰')
                break

    # ── Always try to extract subItemPrice and costDetails ──
    _extract_structured_items(text, result)

    # ── Post-extraction validation: sanity-check and clean up ──
    _validate_price_extraction(text, result, bid_section)

    result.pop('_cn_only', None)     # internal provenance flags, not API data
    result.pop('_from_global', None)
    return result


def _find_bid_summary_section(text):
    """Find the bid summary / price overview section in text.
    Returns a section that actually contains price data (currency + numbers)."""
    keywords = ['开标一览表', '开标一览', '投标报价表', '报价一览表', '报价总表', '投标总价',
                '报价汇总表', '价格汇总表', '报价单', '最高限价', '投标函附录', '唱标单',
                '开标记录表', '投标一览表', '投标函']
    for kw in keywords:
        idx = text.find(kw)
        while idx >= 0:
            # Skip TOC entries (preceded by dots)
            prefix = text[max(0, idx - 40):idx]
            if re.search(r'\.{3,}', prefix):
                idx = text.find(kw, idx + 1)
                continue

            # Skip if this occurrence is inside a compact inline list item
            # like "l．开标一览表" or "1．开标一览表" (within a paragraph).
            # Do NOT skip proper section headers like "二、 开标一览表".
            line_start = text.rfind('\n', 0, idx)
            if line_start < 0:
                line_start = 0
            line_prefix = text[line_start:idx].strip()
            # Only skip if preceded by a single ASCII digit/letter + full-width dot
            # (compact inline list), not a proper section header (Chinese number + 、)
            if re.search(r'^(?:[a-zA-Z\d]|[一二三四五六七八九十]{1,2})[．.]\s*$', line_prefix):
                idx = text.find(kw, idx + 1)
                continue

            # Skip references like "《开标一览表》" — the keyword is inside book-title
            # marks, not a section header (e.g., "愿意以《开标一览表》中的投标报价").
            if line_prefix.endswith('《') or line_prefix.endswith('〈'):
                idx = text.find(kw, idx + 1)
                continue

            # Same for curly-quoted references: '等于"开标一览表"中的投标总价'
            # mentions the table in a note; it is not a section header either.
            if line_prefix.endswith('“') or line_prefix.endswith('‘'):
                idx = text.find(kw, idx + 1)
                continue

            # Skip TOC entries WITHOUT dot leaders and other verbal mentions:
            # many extractors drop the dots, leaving '开标一览表\n三、分项报价表
            # \n四、…\n11\n12' where every line is a title or page number. A REAL
            # summary table shows a price signal within ~400 chars of its header
            # (小写：/大写：/￥/CN-numeral amounts/5+ digit numbers).
            if not re.search(r'(?:小写|大写)\s*[：:]|￥|¥|\d{5,}|\d[，,]\d{3}|'
                             r'[壹贰叁肆伍陆柒捌玖][佰仟万亿]', text[idx:idx + 400]):
                idx = text.find(kw, idx + 1)
                continue

            # Skip inline list items like "（2）开标一览表；" or "1）开标一览表；"
            # These are bid-letter content listings, not actual section headers.
            if re.search(r'[（(]\d+[）)]\s*$', line_prefix):
                idx = text.find(kw, idx + 1)
                continue
            # Also skip when the keyword is followed by "；" or "。" (still in a list)
            after_kw = text[idx + len(kw):idx + len(kw) + 5].strip()
            if after_kw.startswith('；') or after_kw.startswith('。'):
                idx = text.find(kw, idx + 1)
                continue

            # Find end: next major section or 3000 chars
            end = min(idx + 3000, len(text))
            for end_kw in ['投标分项报价表', '投标分项报价', '法定代表人身份证明',
                           '技术方案', '项目概况', '资格证明文件']:
                ep = text.find(end_kw, idx + 30)
                if ep > idx and ep < end:
                    end = ep

            section = text[idx:end]
            # Validate: section must contain a currency indicator with numbers.
            # Accept RMB/CNY/￥/¥ symbols, standalone 元, number+万 patterns,
            # or large plain numbers (≥5 digits, typical for full-unit prices like "1228000").
            _CN_AMT = r'[壹贰叁肆伍陆柒捌玖拾佰仟万亿零一二三四五六七八九十百千元整角分圆]+'
            if re.search(r'(?:人民币|CNY|RMB|￥|¥|元)\s*(?:[\d,]+\.?\d*|' + _CN_AMT + r')', section) or \
               re.search(_CN_AMT + r'\s*元', section) or \
               re.search(r'[\d,]+\.?\d*\s*万', section) or \
               re.search(r'\d{1,3}(?: \d{3})+', section) or \
               re.search(r'(?<!\d)[\d,]{5,}(?![\d,])', section):
                return section

            # Otherwise, skip this occurrence and try next
            idx = text.find(kw, idx + 1)

    return None


def _extract_tax_decomposition(text, section, result):
    """Extract pre-tax / tax / post-tax breakdown."""
    # Skip if both prices AND tax rate are already set from a reliable channel
    # (prevents overwriting good data with random number triplets from full text)
    tp = result.get('totalPrice')
    tpit = result.get('totalPriceInTax')
    if tp is not None and tpit is not None and tp >= 50000 and tpit >= 50000 \
            and result.get('taxRate') is not None:
        return  # Already complete, don't risk overwriting
    # Each pattern is (regex, p1_group, p2_group, tax_group) where:
    #   p1 = 不含税总价, p2 = 含税总价, tax = 税率
    # Patterns 0-3: profit format (price1, tax_rate, price2)
    # Pattern 4:   万 format (price1+万, price2+万, tax_rate)
    # Pattern 5:   plain-number format (price1, price2, tax_rate)
    patterns = [
        (r'(?:￥|¥)?(\d{5,10}(?:\.\d{2})?)\s+(\d{1,2})\s*[%％]\s*(?:￥|¥)?(\d{5,10}(?:\.\d{2})?)', 1, 3, 2),
        (r'人民币[：:\s]*(\d{5,10}(?:\.\d{2})?)\s*元?\s+(\d{1,2})\s*[%％]?\s+人民币[：:\s]*(\d{5,10}(?:\.\d{2})?)', 1, 3, 2),
        (r'([\d.]+\s*万)\s+([\d.]+\s*万)\s+(\d{1,2})', 1, 2, 3),
        (r'小写[：:]\s*(\d{5,10}(?:\.\d{2})?)\s*元[\s\S]{0,80}?(\d{1,2})\s*[%％][\s\S]{0,80}?小写[：:]\s*(\d{5,10}(?:\.\d{2})?)', 1, 3, 2),
        # Plain numeric prices: two ≥5-digit numbers + 1-2 digit tax rate
        # e.g. "1228000 1264840 3" (不含税 含税 税率)
        (r'(?<!\d)(\d{5,10})\s+(\d{5,10})\s+(\d{1,2})(?!\d)', 1, 2, 3),
    ]
    for pat, pi1, pi2, pi_tax in patterns:
        m = re.search(pat, section)
        if m:
            v1 = _parse_amount(m.group(pi1))
            v2 = _parse_amount(m.group(pi2))
            if v1 >= 100 and v2 >= 100:
                # For the loose plain-number pattern (no currency markers), a
                # pair of unrelated 5-10 digit numbers (project code + budget,
                # two serial numbers, etc.) can slip through. Real pre-tax /
                # post-tax prices differ only by the tax rate (1-17%), so their
                # ratio must fall in ~[0.8, 1.2]; reject otherwise.
                is_plain_pair = pat.startswith(r'(?<!\d)')
                if is_plain_pair:
                    hi, lo = max(v1, v2), min(v1, v2)
                    if lo == 0 or hi / lo > 1.25:
                        continue
                    # Digit runs that parse as YYYYMMDD are table dates
                    # ('签订 20250826 / 生效 20250829 / 3'), not a price pair.
                    if _is_yyyymmdd(v1) or _is_yyyymmdd(v2):
                        continue
                result['totalPrice'] = v1
                result['totalPriceInTax'] = v2
                result['taxRate'] = str(int(m.group(pi_tax))) + '%'
                return

    # ── Standalone tax rate extraction (table formats without % sign) ──
    # Only search for tax rate if at least one price is already extracted.
    # A standalone tax rate without a price is useless and likely a false positive.
    if result.get('taxRate') is None and (result.get('totalPrice') or result.get('totalPriceInTax')):
        # Pattern A: "税率（%） ... N" in table header followed by data row
        m = re.search(r'税率\s*[（(]\s*%[）)]?\s*.{0,100}?(\d{1,2})(?:\s|$)', section)
        if m and 1 <= int(m.group(1)) <= 30:
            result['taxRate'] = m.group(1) + '%'
    if result.get('taxRate') is None:
        # Pattern B: "小写：price ... N 增值税" — tax rate before 增值税 in table
        if result.get('totalPriceInTax'):
            price_val = int(result['totalPriceInTax'])
            m = re.search(rf'{re.escape(str(price_val))}[\s\S]{{0,100}}?(\d{{1,2}})\s*(?:增值税|专用|普通|发票)', section)
            if m and 1 <= int(m.group(1)) <= 30:
                result['taxRate'] = m.group(1) + '%'
    if result.get('taxRate') is None and (result.get('totalPrice') or result.get('totalPriceInTax')):
        # Pattern C: standalone "6 %" or "6%" in table cell
        m = re.search(r'(?<!\d)(\d{1,2})\s*[%％](?!\d)', section)
        if m and 1 <= int(m.group(1)) <= 30:
            result['taxRate'] = m.group(1) + '%'


def _extract_structured_items(text, result):
    """Extract sub-item pricing and cost details using enhanced section discovery
    and generic cost line detection."""

    # ── Enhanced section discovery ──
    # Order matters: longer keywords first to avoid partial matches
    section_keywords = [
        '报价明细表', '报价一览表', '分项报价表', '经费总表', '费用明细表',
        '分项报价', '报价明细', '价格表', '开标一览',
        '报价清单', '费用明细', '价格清单', '投标报价', '价格构成',
        '设备清单', '费用清单', '报价构成', '价格明细', '成本明细',
        '项目报价', '费用构成', '费用表'
    ]

    bid_section = None
    # Build alternation pattern from keywords
    kw_pattern = '|'.join(re.escape(kw) for kw in section_keywords)
    for m in re.finditer(
        r'(?:^|\n)(?:[一二三四五六七八九十\d]+[、.。]\s*|\d+(?:\.\d+)+\s*|\d+\s+)?(' +
        kw_pattern + r')[^\n]*\n', text
    ):
        pos = m.start()
        # Skip if preceded by dots (TOC entry)
        prefix = text[max(0, pos - 30):pos]
        if re.search(r'\.{3,}', prefix):
            continue
        # Find next major section boundary
        next_pos = len(text)
        for end_marker in ['\n三、', '\n四、', '\n五、', '\n六、', '\n七、',
                           '\n3.', '\n4.', '\n5.', '\n6.', '\n7.']:
            ep = text.find(end_marker, pos + 10)
            if ep > pos and ep < next_pos:
                next_pos = ep
        # Also stop at section headers (number + 5+ Chinese chars, no price amounts)
        for sm in re.finditer(r'\n(\d{1,2})\s+[一-鿿]{5,}', text):
            if sm.start() > pos + 20 and sm.start() < next_pos:
                line_end = text.find('\n', sm.end())
                line = text[sm.start()+1:line_end if line_end > sm.start() else sm.end()+80]
                if len(line) < 60 and not re.search(r'\d{4,}', line):
                    next_pos = sm.start()
                    break
        bid_section = text[pos:next_pos]
        break

    if bid_section:
        # Detect docx pipe-separated tables (vs PDF space-separated)
        pipe_lines = len(re.findall(r'\n[^|\n]+\|[^|\n]+\|[^\n]+', bid_section))
        if pipe_lines >= 2:
            _parse_docx_bid_table(bid_section, result)
        else:
            _parse_pdf_bid_table(bid_section, result)

    # ── Smart docx table scan: find pricing-relevant pipe tables anywhere in text ──
    # Score each pipe table region by pricing relevance and parse only the best ones
    if not result.get('subItemPrice'):
        _scan_docx_tables_for_pricing(text, result)

    # ── Generic cost line extraction ──
    # Normalize cost names for dedup and noise filtering
    def _norm_cost_name(name):
        n = name.strip()
        n = re.sub(r'^本项目', '', n)
        n = re.sub(r'为$', '', n)
        # Filter category headers (not real cost items)
        if re.search(r'(?:项目预计成本|预计成本|小计|合计|总价|总计)', n):
            return None
        return n.strip()

    # Pattern A: "本项目材料费为   340,000.00元" (defense industry descriptive format)
    desc_cost_pat = re.compile(
        r'本项目\s*([一-鿿]{2,20}(?:费|成本|支出|收益|利润))\s*为\s*([\d,]+\.?\d*)\s*元',
        re.MULTILINE
    )
    seen_names = set()
    for m in desc_cost_pat.finditer(text):
        raw_name = m.group(1).strip()
        name = _norm_cost_name(raw_name)
        if not name or name in seen_names:
            continue
        val = float(m.group(2).replace(',', ''))
        if val == 0:
            seen_names.add(name)
            continue
        if val >= 100:
            seen_names.add(name)
            if '收益' in name or '利润' in name:
                if result['revenue'] is None:
                    result['revenue'] = val
            else:
                result['costDetails'].append({
                    'priceName': name,
                    'totalPrice': val,
                    'unit': None, 'count': None, 'unitPrice': None,
                    'tax': None, 'totalPriceInTax': val,
                    'extras': {}, 'details': []
                })

    # Pattern B: Docx pipe-separated cost lines: "材料费 | 340000"
    pipe_cost_pat = re.compile(
        r'(?:^|\n)\s*(?:[（(][一二三四五六七八九十\d]+[）)]\s*)?'
        r'(\d+\.\d+\s*)?'   # optional "1.1 " prefix
        r'([一-鿿]{2,20}(?:费|成本|支出|收益|利润|不可预见))\s*\|\s*(\d{3,}(?:\.\d{2})?)',
        re.MULTILINE
    )
    for m in pipe_cost_pat.finditer(text):
        raw_name = m.group(2).strip() if m.group(2) else m.group(1).strip()
        name = _norm_cost_name(raw_name)
        if not name or name in seen_names:
            continue
        val = float(m.group(3).replace(',', ''))
        if val == 0:
            seen_names.add(name)
            continue
        seen_names.add(name)
        if '收益' in name or '利润' in name:
            if result['revenue'] is None:
                result['revenue'] = val
        else:
            result['costDetails'].append({
                'priceName': name,
                'totalPrice': val,
                'unit': None, 'count': None, 'unitPrice': None,
                'tax': None, 'totalPriceInTax': val,
                'extras': {}, 'details': []
            })

    # Pattern C: Whitespace-separated cost lines (original pattern, kept as fallback)
    # Tolerates flattened-cell shapes: 序号 prefixes ('1. ' / '（二）'), an
    # optional colon between name and value ('材料费：340,000.00'), thousands
    # commas, and 万元 magnitudes.
    cost_pattern = re.compile(
        r'(?:^|\n)\s*(?:[（(][一二三四五六七八九十\d]+[）)]\s*|\d+[.、]\s*)?'
        r'([一-鿿]{2,20}(?:费|成本|支出|投入|工资|薪酬|酬金|折旧|摊销|租赁|租金|'
        r'维护|保养|检测|试验|测试|设计|开发|研制|采购|运输|差旅|会议|培训|办公|印刷|'
        r'咨询|审计|评估|保险|税费|利息|手续费|管理|服务|劳务|材料|设备|仪器|软件|'
        r'许可|专利|著作|技术|咨询|外协|加工|燃料|动力|事务|不可预见|预备|风险|'
        r'收益|利润|税金|公积金|基金)[一-鿿]{0,6})\s*[：:]?\s*(\d[\d,]*(?:\.\d{1,2})?)\s*(万元?)?',
        re.MULTILINE
    )
    for m in cost_pattern.finditer(text):
        raw_name = m.group(1).strip()
        name = _norm_cost_name(raw_name)
        if not name or name in seen_names:
            continue
        digits = m.group(2)
        val = float(digits.replace(',', ''))
        if m.group(3):  # 万元 magnitude suffix — keep the real scale
            val *= 10000 if m.group(3).startswith('万') else 1
        elif len(digits.replace(',', '').replace('.', '')) < 4:
            # Bare numbers keep the old \d{4,} guard (skip small values/years)
            seen_names.add(name)
            continue
        # Bare 8-digit runs that parse as YYYYMMDD are table dates, not costs
        if not m.group(3) and _is_yyyymmdd(val):
            seen_names.add(name)
            continue
        # Skip past-performance tables: a 业绩一览表 lists HISTORICAL contract
        # amounts ('合同金额(元)' header, project-name fragments like
        # '站产品采购' split across cells) — never this bid's cost items.
        if re.search(r'(?:合同金额|合同价款|签约合同价|业绩一览|类似项目|项目业绩|业绩证明)',
                     text[max(0, m.start() - 300):m.start()]):
            seen_names.add(name)
            continue
        # ── Additional filtering for Pattern C (global text search) ──
        # Reject numbers that look like years (1990-2030)
        if 1990 <= val <= 2030 or (500 <= val <= 999 and val == int(val)):
            seen_names.add(name)
            continue
        # Reject names that look like organization/regulatory bodies
        if re.search(r'(?:委员会|协会|公司|有限|集团|事务所|证监会|财政部)', name):
            seen_names.add(name)
            continue
        # Reject names ending with person/professional indicators
        if re.search(r'(?:师|人|员|专家)$', name):
            seen_names.add(name)
            continue
        # Reject names that are clearly not cost items
        if re.search(r'(?:成立于|于$|成为|自$|见后附|次会议|授权)', name):
            seen_names.add(name)
            continue
        # Validate the number appears near a currency indicator (元/万) within
        # 20 chars — an explicit 万元 suffix captured above already proves it.
        if not m.group(3):
            match_end = m.end()
            post_context = text[match_end:match_end + 30]
            if not re.search(r'(?:元|万|万元|CNY|RMB|￥|¥)', post_context):
                seen_names.add(name)
                continue
        if val >= 100:
            seen_names.add(name)
            if '收益' in name or '利润' in name:
                if result['revenue'] is None:
                    result['revenue'] = val
            else:
                result['costDetails'].append({
                    'priceName': name,
                    'totalPrice': val,
                    'unit': None, 'count': None, 'unitPrice': None,
                    'tax': None, 'totalPriceInTax': val,
                    'extras': {}, 'details': []
                })

    # Recalculate total cost from cost details if not already set
    if result['costDetails'] and result['cost'] is None:
        result['cost'] = sum(item['totalPrice'] for item in result['costDetails'])


def _validate_price_extraction(text, result, bid_section):
    """Post-extraction sanity checks. Clears values that fail validation to prevent
    showing garbage data (e.g. project history amounts) as bid prices."""
    tp = result.get('totalPrice')
    tpit = result.get('totalPriceInTax')
    tax_str = result.get('taxRate')
    cost = result.get('cost')

    # ── Helper: check if a numeric value appears near price-indicator keywords ──
    def _near_price_context(value, window=120):
        if value is None:
            return True  # nothing to validate
        # Find the value in text (as int, to avoid matching substrings)
        val_int = int(value)
        # Search for the value in various formats.
        # Iterate over ALL occurrences (not just text.find's first hit): the
        # same number often appears earlier in a non-price context (bid bond,
        # project code) and only later near the real bid price. Using only the
        # first occurrence would wrongly reject a valid price.
        # Also search 万/亿-suffixed forms: a value extracted from '￥12.5万元'
        # has no plain Arabic '125000' in the text to match against.
        formats = [str(val_int), f'{val_int:,}', f'{val_int:.2f}', f'{val_int:.1f}']
        # Thin-space thousands groups ('1 234 567') emitted by PDF extraction
        # instead of commas — without this form the validator rejects a value
        # that genuinely appears in the text.
        comma_fmt = f'{val_int:,}'
        if ',' in comma_fmt:
            formats.append(comma_fmt.replace(',', ' '))
        if val_int >= 10000:
            formats += [f'{val_int / 10000:g}万', f'{val_int / 10000:g}万元']
        if val_int >= 100000000:
            formats += [f'{val_int / 100000000:g}亿', f'{val_int / 100000000:g}亿元']
        for fmt in formats:
            search_from = 0
            while True:
                idx = text.find(fmt, search_from)
                if idx < 0:
                    break
                search_from = idx + len(fmt)
                ctx_start = max(0, idx - window)
                ctx_end = min(len(text), idx + len(fmt) + window)
                ctx = text[ctx_start:ctx_end]
                # Exclude non-price contexts BEFORE checking for price indicators
                # "出资额为人民币XXX" / "注册资金XXX万元" / "合同金额：RMB2080000" /
                # "招标文件售价：人民币1000元" / "最高限价：￥XXX" → NOT the bid price
                prefix = text[max(0, idx - 30):idx]
                if re.search(_NON_BID_AMOUNT_CTX +
                             r'\s*[额为]?\s*[：:]?\s*(?:人民币|RMB|CNY|￥|¥)?\s*$', prefix):
                    continue
                if re.search(r'(?:万元|万)\s*$', prefix):
                    continue
                # Must contain a price-indicator keyword nearby
                if re.search(r'(?:元|人民币|CNY|RMB|￥|¥|报价[总金]|投标[总报]|'
                             r'金额|总价|开标|一览表|大写|小写|合计|总计)', ctx):
                    return True
        return False

    # ── Validate totalPrice vs totalPriceInTax consistency ──
    if tp is not None and tpit is not None:
        # Extract tax rate as float for calculation
        tax_rate = None
        if tax_str:
            m = re.search(r'(\d+(?:\.\d+)?)', tax_str)
            if m:
                tax_rate = float(m.group(1)) / 100.0

        # Rule 1: Both prices must be in similar magnitude (ratio between 0.1 and 10)
        ratio = tpit / tp if tp > 0 else float('inf')
        if ratio < 0.05 or ratio > 20:
            # Wildly different magnitudes — likely from different sources
            # Check which one is near price context; keep only that one
            tp_ok = _near_price_context(tp)
            tpit_ok = _near_price_context(tpit)
            if tp_ok and not tpit_ok:
                result['totalPriceInTax'] = None
            elif tpit_ok and not tp_ok:
                result['totalPrice'] = None
            else:
                # Neither or both near context — clear both to be safe
                result['totalPrice'] = None
                result['totalPriceInTax'] = None
                result['taxRate'] = None
        elif tax_rate is not None:
            # Rule 2: If tax rate is present, verify price relationship.
            # If the two prices are mutually consistent but disagree with the
            # (often loosely-extracted) tax rate, the tax rate is the suspect
            # element -- clear ONLY it, not the correctly-extracted prices.
            expected_tpit = tp * (1 + tax_rate)
            if abs(tpit - expected_tpit) / expected_tpit > 0.15:
                # Prices don't match the stated tax rate -> distrust the rate.
                result['taxRate'] = None

    # ── Validate individual prices against text context ──
    # Global-search hits need extra scrutiny — either when no bid section
    # exists, or when a bid section exists but the price came from the
    # full-text fallback (e.g. a 合同金额 line before the real price).
    if bid_section is None or result.get('_from_global'):
        # Higher threshold for global search: real bid prices are >= 5000
        if tpit is not None and tpit < 5000:
            result['totalPriceInTax'] = None
        if tp is not None and tp < 5000:
            result['totalPrice'] = None
        # If both prices were cleared, tax rate alone is meaningless
        if result['totalPrice'] is None and result['totalPriceInTax'] is None:
            result['taxRate'] = None

        # Prices sourced purely from Chinese numerals have no Arabic form in
        # the text; skip the near-context requirement for them.
        cn_only = result.get('_cn_only')
        if tpit is not None and not cn_only and not _near_price_context(tpit, window=200):
            tp_val = result['totalPrice']
            if tp_val is None or not _near_price_context(tp_val, window=200):
                result['totalPriceInTax'] = None
        if tp is not None and not cn_only and not _near_price_context(tp, window=200):
            tpit_val = result['totalPriceInTax']
            if tpit_val is None or not _near_price_context(tpit_val, window=200):
                result['totalPrice'] = None

    # If one price was cleared but the other remains, ensure tax rate is consistent
    if (result['totalPrice'] is None) != (result['totalPriceInTax'] is None):
        # Only one price remains — tax rate is meaningless
        result['taxRate'] = None

    # ── Validate cost against total price ──
    cost_val = result.get('cost')
    tp_val = result.get('totalPrice') or result.get('totalPriceInTax')
    if cost_val is not None and tp_val is not None:
        if cost_val > tp_val * 5 or cost_val < 100:
            # Cost far exceeds price or is trivially small
            result['cost'] = None

    # ── 大写/小写 cross-consistency ──
    # When both forms are printed and disagree (>1%), one of them is wrong.
    # 小写 (Arabic) is authoritative — CN-numeral typos (壹/贰 mixups, dropped
    # 万) are far more common than digit typos. If the CN value currently
    # holds the slot, switch to the Arabic value and record a note. Scanning
    # the pair directly (not the current value's position) also catches
    # CN-derived totals whose Arabic form never appears elsewhere in text.
    cur = result.get('totalPriceInTax')
    if cur is not None:
        # The 大写 label is optional in p1: layouts like '投标总价：壹佰万元整
        # （小写：980000元）' pair a bare CN amount with its 小写 form. Paren
        # placement varies ('大写：壹佰万' / '（大写）：壹佰万' / '（小写)：98万').
        # Matches are validated before use: a lone '一' (e.g. '开标一览表')
        # must not pair up as a 大写 amount.
        # CN amount must END with a unit char (元/整/万…) so a lone '一'
        # ('开标一览表') can never consume the pair span as a fake 大写.
        _CN_UNIT_END = r'(?<=[元整角分圆万亿])'
        for pat in (
            r'(?:大写[）)]?\s*[：:]?\s*[（(]?\s*)?(' + _AMT_CN + r')' + _CN_UNIT_END +
            r'\s*[）)]?\s*[\s\S]{0,120}?'
            r'小写[）)]?\s*[：:]?\s*[（(]?\s*([\d,]+\.?\d*)',
            r'小写[）)]?\s*[：:]?\s*[（(]?\s*([\d,]+\.?\d*)\s*[）)]?\s*[\s\S]{0,120}?'
            r'大写[）)]?\s*[：:]?\s*[（(]?\s*(' + _AMT_CN + r')' + _CN_UNIT_END,
        ):
            matched = False
            for m in re.finditer(pat, text):
                if pat.startswith('小写'):
                    ar_val, cn_str = _parse_amount(m.group(1)), m.group(2)
                else:
                    cn_str, ar_val = m.group(1), _parse_amount(m.group(2))
                cn_val = _parse_amount(cn_str)
                if cn_val <= 1000 or ar_val <= 1000 or len(cn_str) < 2 \
                        or not re.search(r'[元圆整角分万亿拾佰仟]', cn_str):
                    continue  # junk pair — try the next match
                matched = True
                tol = max(cn_val, ar_val) * 0.01
                if abs(cn_val - ar_val) > tol:
                    if abs(cur - cn_val) <= tol:
                        # current value came from the (wrong) 大写 → adopt 小写
                        result['totalPriceInTax'] = ar_val
                        if result.get('totalPrice') is not None \
                                and abs(result['totalPrice'] - cn_val) <= tol:
                            scale = ar_val / cn_val
                            result['totalPrice'] = round(result['totalPrice'] * scale, 2)
                        result['warnings'].append(
                            f'大写金额({cn_val:,.0f})与小写({ar_val:,.0f})不一致，已采用小写值，建议人工复核')
                    else:
                        result['warnings'].append(
                            f'检测到大写({cn_val:,.0f})与小写({ar_val:,.0f})金额不一致，请人工复核')
                break
            if matched:
                break

    # ── Sub-item sum vs total ──
    sub_sum = sum(it.get('totalPrice') or 0 for it in result.get('subItemPrice') or [])
    total_now = result.get('totalPriceInTax') or result.get('totalPrice')
    if sub_sum > 0 and total_now:
        diff = abs(sub_sum - total_now) / total_now
        if diff > 0.05:
            result['warnings'].append(
                f'分项合计({sub_sum:,.0f})与总价({total_now:,.0f})差异 {diff:.0%}，'
                f'报价或分项可能提取不完整')

    # ── Provenance note: total came from full-text fallback ──
    if result.get('_from_global') and result.get('totalPriceInTax') is not None:
        result['warnings'].append('总价来自全文兜底匹配（未定位到报价章节），置信度较低，建议人工复核')

    # ── Filter suspicious cost details ──
    if result.get('costDetails'):
        filtered = []
        for item in result['costDetails']:
            name = item.get('priceName', '')
            val = item.get('totalPrice', 0)
            # Reject items that look like certificate/reference numbers
            if re.search(r'(?:证书|编号|注册|登记|代码|序列)', name):
                continue
            # Reject items that are trivially small
            if val < 100:
                continue
            # Reject items with suspicious names (too long, contains date patterns)
            if len(name) > 15 or re.search(r'\d{4}', name):
                continue
            # Reject items with names that look like project titles
            if re.search(r'(?:项目|公司|有限|集团|评估项目)', name):
                continue
            filtered.append(item)
        if filtered != result['costDetails']:
            result['costDetails'] = filtered
            # Recalculate cost from filtered details
            if filtered:
                result['cost'] = sum(item['totalPrice'] for item in filtered)
            else:
                result['cost'] = None


def _parse_pdf_bid_table(section, result):
    """Parse a pricing table section with dynamic column detection.
    Automatically identifies column types regardless of ordering."""
    clean = re.sub(r'[.]{3,}\s*\d*', '', section)

    lines = clean.split('\n')

    # Find header row — look for 序号 + column name keywords
    col_keywords = ['序号', '名称', '型号', '规格', '数量', '单价', '总价', '税率', '备注', '厂家']
    header_line = -1
    for i, line in enumerate(lines):
        hits = sum(1 for kw in col_keywords if kw in line)
        if hits >= 3:
            header_line = i
            break

    if header_line < 0:
        # Fallback: search relaxed
        for i, line in enumerate(lines):
            if re.search(r'序\s*号', line) and re.search(r'(?:名称|型号|产品|服务)', line):
                header_line = i
                break

    if header_line < 0:
        return

    # ── Column type inference from header ──
    col_order = _infer_columns(lines[header_line])

    # ── Data row parsing ──
    data_end = None
    for i in range(header_line + 1, len(lines)):
        s = lines[i].strip()
        if not s:
            continue
        if any(s.startswith(kw) for kw in ['合计', '总价', '小计', '总计', '注：', '备注：']):
            data_end = i
            break
        if re.match(r'^[三四五六七八九十]、', s):
            data_end = i
            break
    if data_end is None:
        data_end = len(lines)

    # Collect and merge data lines
    raw_rows = []
    for i in range(header_line + 1, data_end):
        s = lines[i].strip()
        if not s or re.match(r'^\d{1,3}$', s):
            continue
        raw_rows.append(s)

    # Merge wrapped names: a line with no amounts merges into the next line with amounts.
    # Also merge numeric-only continuation lines into the previous data line:
    # PDF text layers frequently wrap a table row so the later price columns
    # land on the following line ('设备A 2 5000' + '5300 13%').
    merged_rows = []
    pending_name = []
    _NUM_ONLY = re.compile(r'^[\d\s.,%％‰\\/万元元（）()-]+$')
    for s in raw_rows:
        has_amounts = bool(re.search(r'(\d{4,}|[\d.]+\s*万)', s))
        if has_amounts:
            if pending_name:
                merged_rows.append((''.join(pending_name), s))
                pending_name = []
            else:
                merged_rows.append(('', s))
        elif _NUM_ONLY.match(s) and merged_rows and not pending_name:
            # numeric continuation of the previous data row
            prev_name, prev_data = merged_rows[-1]
            merged_rows[-1] = (prev_name, prev_data + ' ' + s.strip())
        else:
            pending_name.append(s)

    # Process each data row
    items = []
    prev_name = None
    for name_part, data in merged_rows:
        name = name_part.strip() if name_part.strip() else ''
        name = re.sub(r'^\d+\s*', '', name).strip()
        # When name is on the same line as data (not wrapped), extract it from data
        if not name:
            rest = re.sub(r'^\d+\s*', '', data).strip()
            # Find first digit position (count field) in rest
            first_digit = re.search(r'\d', rest)
            if first_digit:
                prefix = rest[:first_digit.start()].strip()
                # prefix is "name mfr" — split on last whitespace to separate
                parts = prefix.rsplit(None, 1)
                name = parts[0].strip() if parts else prefix
        if not name and prev_name:
            name = prev_name
        elif name:
            prev_name = name

        if len(name) < 2:
            continue

        # Extract numbers (strip leading row number first)
        data = re.sub(r'^\d+\s*', '', data).strip()
        uses_wan = '万' in data
        if uses_wan:
            nums_parsed = []
            for m in re.finditer(r'([\d.]+)\s*(万)?', data):
                v = float(m.group(1))
                if m.group(2):
                    v *= 10000
                nums_parsed.append(v)
        else:
            nums = re.findall(r'(\d+(?:\.\d+)?)', data)
            nums_parsed = [float(n) for n in nums]

        if len(nums_parsed) < 2:
            continue

        # Separate small values (count, tax rate) from large values (prices)
        smalls = [v for v in nums_parsed if v < 100]
        larges = [v for v in nums_parsed if v >= 100]

        if len(larges) < 2:
            continue

        # ── Column mapping using inferred order ──
        item = {'priceName': name, 'unit': '项', 'extras': {}, 'details': []}

        _assign_columns(item, nums_parsed, smalls, larges, col_order, uses_wan)

        # Manufacturer extraction
        mfr_match = re.match(r'([^\d]+?)\s+(?=\d)', data)
        if mfr_match:
            mfr = mfr_match.group(1).strip()
            mfr = re.sub(r'^[/\-\s]+', '', mfr)
            if mfr and mfr != name and mfr not in ('/', '--', '-'):
                item['extras']['厂家/型号'] = mfr

        if item.get('totalPrice'):
            items.append(item)

    if items:
        result['subItemPrice'] = _filter_price_items(items)
        _extract_summary_total(section, result)


def _infer_columns(header_line):
    """Infer column types and order from header text.
    Returns list of column type strings sorted by position."""
    col_map = []
    detectors = [
        (r'序\s*号', 'seq'),
        (r'(?:分项\s*)?名\s*称|产品|服务|项目|内容', 'name'),
        (r'型号|规格|厂家|制造商|品牌', 'model'),
        (r'数\s*量', 'count'),
        (r'单价\s*[（(]?\s*不含税\s*[）)]?|不含税\s*单价', 'unit_price_ex'),
        (r'单价\s*[（(]?\s*含税\s*[）)]?|含税\s*单价', 'unit_price_in'),
        (r'总价\s*[（(]?\s*不含税\s*[）)]?|不含税\s*总价', 'total_ex'),
        (r'总价\s*[（(]?\s*含税\s*[）)]?|含税\s*总价', 'total_in'),
        (r'税\s*率', 'tax_rate'),
        (r'备\s*注', 'remark'),
    ]
    for pattern, col_type in detectors:
        m = re.search(pattern, header_line)
        if m:
            col_map.append((col_type, m.start()))
    col_map.sort(key=lambda x: x[1])

    # If no explicit ex/in split, use generic unit_price/total labels
    has_explicit = any(c[0] in ('unit_price_ex', 'unit_price_in', 'total_ex', 'total_in') for c in col_map)
    if not has_explicit:
        col_map = [(t if t not in ('unit_price_ex', 'unit_price_in') else 'unit_price', pos)
                   for t, pos in col_map]

    return [c[0] for c in col_map]


def _assign_columns(item, all_nums, smalls, larges, col_order, uses_wan):
    """Assign extracted numbers to item fields based on inferred column order."""
    # Count: first small integer (1-999)
    count = 1
    for v in smalls:
        if 1 <= v <= 999 and v == int(v):
            count = int(v)
            break
    item['count'] = count

    # Tax rate: value in 1-30 range
    for v in reversed(smalls):
        if 1 <= v <= 30:
            item['tax'] = str(int(v)) + '%'
            break

    # Price columns: map large values to correct fields
    # Common patterns:
    # [unit_ex, unit_in, total_ex, total_in] (4 large values)
    # [unit_ex, total_ex, total_in] (3 large values)
    # [unit_ex, total_ex] (2 large values)

    has_ex_in_split = any(col in col_order for col in ['unit_price_ex', 'unit_price_in', 'total_ex', 'total_in'])

    if has_ex_in_split and len(larges) >= 4:
        # Pattern: [unit_ex, total_ex, unit_in, total_in]
        item['unitPrice'] = larges[0]
        item['totalPrice'] = larges[1]
        item['totalPriceInTax'] = larges[3]
    elif len(larges) >= 4:
        # Pattern without explicit split: [unit, total, ...] — use safe defaults
        item['unitPrice'] = larges[0]
        item['totalPrice'] = larges[1]
        item['totalPriceInTax'] = larges[-1]
    elif len(larges) == 3:
        # Pattern: [unit_ex, total_ex, total_in]
        item['unitPrice'] = larges[0]
        item['totalPrice'] = larges[1]
        item['totalPriceInTax'] = larges[2]
    elif len(larges) == 2:
        # Pattern: [unit_ex, total_ex]
        item['unitPrice'] = larges[0]
        item['totalPrice'] = larges[1]
        item['totalPriceInTax'] = larges[1]


def _extract_summary_total(section, result):
    """Extract total/summary line from pricing table section."""
    # Docx pipe format: look for "总计 | 893000" or "合计 | ... | 893000元 | 1009090元"
    for kw in ['总计', '合计', '总价', '小计']:
        # Pattern: "总计 | 893000" (simple pipe row)
        m = re.search(kw + r'\s*\|\s*(\d{4,12}(?:\.\d{2})?)', section)
        if m:
            val = float(m.group(1))
            if val >= 10000:
                result['totalPrice'] = val
                if result['totalPriceInTax'] is None:
                    result['totalPriceInTax'] = val
                return
        # Pattern: "合计 | ... (many cols) ... | 893000元 | 1009090元" (wide pipe row)
        # Search for rows starting with kw and having 2+ large numbers near the end
        for row_m in re.finditer(kw + r'\s*\|.+', section):
            row = row_m.group()
            parts = [p.strip() for p in row.split('|')]
            nums = []
            for p in parts:
                nm = re.search(r'([\d,]+\.?\d+)', p.replace(',', '').replace('，', ''))
                if nm:
                    nums.append(float(nm.group(1)))
            large = [n for n in nums if n >= 50000]
            if len(large) >= 2:
                result['totalPrice'] = large[-2]
                result['totalPriceInTax'] = large[-1]
                return
            elif len(large) == 1 and large[0] >= 100000:
                result['totalPrice'] = large[0]
                result['totalPriceInTax'] = large[0]
                return
        # Standard ws-separated
        m = re.search(kw + r'\s+([\d.]+)\s*万', section)
        if m:
            result['totalPrice'] = _parse_amount(m.group(1) + '万')
            return
        m = re.search(kw + r'\s+(\d{5,12}(?:\.\d{2})?)', section)
        if m:
            result['totalPrice'] = float(m.group(1))
            return


def _scan_docx_tables_for_pricing(text, result):
    """Scan full text for pipe-separated tables and parse only the most pricing-relevant ones.
    Filters out personnel, project history, tech spec tables by scoring header keywords."""
    # Split text into pipe-table regions (consecutive lines with |)
    lines = text.split('\n')
    regions = []
    region_start = -1
    for i, line in enumerate(lines):
        has_pipe = '|' in line
        if has_pipe and region_start < 0:
            region_start = i
        elif not has_pipe and region_start >= 0:
            if i - region_start >= 3:  # at least 3 pipe lines
                regions.append((region_start, i))
            region_start = -1
    if region_start >= 0 and len(lines) - region_start >= 3:
        regions.append((region_start, len(lines)))

    # Merge nearby regions (gap <= 6 non-pipe lines) to handle split multi-line headers
    merged = []
    for start, end in regions:
        if merged and start - merged[-1][1] <= 6:
            merged[-1] = (merged[-1][0], end)
        else:
            merged.append((start, end))
    regions = merged

    if not regions:
        return

    # Score each region for pricing relevance
    PRICE_COL_KW = ['单价', '总价', '税率', '金额', '价格', '报价', '不含税', '含税']
    NON_PRICE_KW = ['出差事由', '合同金额', '项目名称', '职务', '岗位', '职称',
                    '联系人', '联系电话', '项目经理', '指标要求', '功能要求']
    scored = []
    for start, end in regions:
        region_text = '\n'.join(lines[start:end])
        score = 0
        # Bonus for pricing column headers
        for kw in PRICE_COL_KW:
            if kw in region_text:
                score += 3
        # Penalty for non-price table keywords
        for kw in NON_PRICE_KW:
            if kw in region_text:
                score -= 2
        # Bonus for having rows with large numbers (>= 10000)
        large_count = len(re.findall(r'\b\d{5,}(?:\.\d{2})?\b', region_text))
        score += min(large_count, 10)
        scored.append((score, region_text))

    # Parse all qualifying regions (score >= 5), accumulating items from EACH table
    all_sub_items = []
    for score, region_text in sorted(scored, key=lambda x: -x[0]):
        if score >= 5:
            temp_result = {'subItemPrice': [], 'totalPrice': None, 'totalPriceInTax': None}
            _parse_docx_bid_table(region_text, temp_result)
            if temp_result.get('subItemPrice'):
                all_sub_items.extend(temp_result['subItemPrice'])
            # Capture totals from the first region that has BOTH values
            if (temp_result.get('totalPrice') and temp_result.get('totalPriceInTax')
                    and not result.get('totalPrice')):
                result['totalPrice'] = temp_result['totalPrice']
                result['totalPriceInTax'] = temp_result['totalPriceInTax']

    if all_sub_items:
        result['subItemPrice'] = all_sub_items



def _parse_docx_bid_table(section, result):
    """Parse a docx pipe-separated (|) pricing table.
    Tries multiple header candidates; keeps the one producing most items."""
    lines = section.split('\n')
    HEADER_KW = r'(?:序号|名称|分项|数量|单位|单价|总价|税率|型号|规格|厂家|备注|产品|服务)'
    PRICE_KW = r'(?:单价|总价|税率|不含税|含税|金额)'

    candidates = []
    for i, line in enumerate(lines):
        if '|' not in line: continue
        parts = [p.strip() for p in line.split('|')]
        hits = sum(1 for p in parts if re.search(HEADER_KW, p))
        if hits >= 2:
            price_score = sum(1 for p in parts if re.search(PRICE_KW, p))
            candidates.append((price_score, hits, i))

    if not candidates: return
    candidates.sort(key=lambda x: (-x[0], -x[1]))

    # Parse from ALL qualifying headers (multiple tables in one section)
    all_items = []
    seen = set()
    for _, _, hdr_idx in candidates[:10]:
        items = _parse_docx_rows(lines, hdr_idx, HEADER_KW)
        for item in items:
            key = (item['priceName'], item.get('totalPrice'))
            if key not in seen:
                seen.add(key)
                all_items.append(item)

    if all_items:
        result['subItemPrice'] = _filter_price_items(all_items)
    _extract_summary_total(section, result)


def _parse_docx_rows(lines, hdr_idx, HEADER_KW):
    """Parse rows from a docx pipe table starting at hdr_idx. Returns item list."""
    items = []

    # Merge forward header continuations
    hdr_end = hdr_idx
    for j in range(hdr_idx + 1, min(hdr_idx + 4, len(lines))):
        nxt = lines[j].strip()
        if not nxt or '|' not in nxt: continue
        nxt_p = [p.strip() for p in nxt.split('|')]
        if sum(1 for p in nxt_p if re.search(HEADER_KW, p)) >= 1 and not any(re.search(r'\d{4,}', p) for p in nxt_p):
            hdr_end = j
        else: break

    # Backward merge
    hdr_start = hdr_idx
    for j in range(hdr_idx - 1, max(hdr_idx - 3, -1), -1):
        prev = lines[j].strip()
        if not prev or '|' not in prev: continue
        prev_p = [p.strip() for p in prev.split('|')]
        if sum(1 for p in prev_p if re.search(HEADER_KW, p)) >= 1 and not any(re.search(r'\d{4,}', p) for p in prev_p):
            hdr_start = j
        else: break

    # Build merged header for name column detection
    merged = ' | '.join(lines[i] for i in range(hdr_start, hdr_end + 1))
    merged_parts = [p.strip() for p in merged.split('|')]
    name_col = 1
    for ci, h in enumerate(merged_parts):
        if re.search(r'(?:分项\s*)?名\s*称|分项|产品|服务|项目|内容|元器件', h):
            name_col = ci; break

    prev_name = None
    for i in range(hdr_end + 1, len(lines)):
        line = lines[i].strip()
        if not line or '|' not in line: continue
        if any(line.startswith(kw) for kw in ['合计', '总价', '小计', '总计', '注：']): continue
        if re.match(r'^[三四五六七八九十]、', line): break

        parts = [p.strip() for p in line.split('|')]
        name = parts[name_col].strip() if len(parts) > name_col else ''
        # Patterns that indicate a manufacturer/company name rather than a price item
        _MFG_NAME_RE = re.compile(
            r'(?:有限公司|有限责任|公司|集团|大学|学院|研究所|'
            r'(?:科技|技术|电子|光电|仪器|测量|通信|半导体|激光)'
            r'.{0,5}(?:中国|日本|美国|德国|英国|法国|意大利|加拿大|澳大利亚|'
            r'马来西亚|新加坡|韩国|越南|印度|泰国|台湾|香港|澳门|'
            r'北京|上海|深圳|广州|成都|武汉|南京|杭州|西安))')
        if not name or _MFG_NAME_RE.search(name):
            # Try alternative columns for the real item name (skip seq col 0)
            for alt_ci in range(1, len(parts)):
                if alt_ci == name_col:
                    continue
                alt = parts[alt_ci].strip()
                if alt and not _MFG_NAME_RE.search(alt) and len(alt) >= 2:
                    # Only use if it doesn't look like a pure number/spec
                    if not re.match(r'^[\d.,\s]+$', alt) and not re.search(r'^\d{4,}', alt):
                        name = alt
                        break
        name = re.sub(r'^\d+(?:\.\d+)?\s*', '', name).strip()
        if _MFG_NAME_RE.search(name): continue

        # Skip all-same-value group headers
        unique = set(p.strip() for p in parts if p.strip())
        if len(unique) <= 2 and len(parts) >= 4 and all(not re.search(r'\d{4,}', p) for p in parts): continue

        if not name and prev_name: name = prev_name
        elif name and len(name) >= 2: prev_name = name
        if len(name) < 2: continue

        # Value extraction with context-aware classification
        # Detect spec/non-price patterns in cells (resolution, IP ratings, model numbers)
        _SPEC_PAT = re.compile(r'(?:分辨率|IP\d|dB|MHz|GHz|mm|cm|kg|g\b|V\b|A\b|W\b|'
                               r'像素|英寸|寸|比特|波特|bps|kbps|℃|℉|'
                               r'规格|型号|品牌|厂家|制造商)')
        all_nums = []
        for pi, p in enumerate(parts):
            pct = re.search(r'(\d{1,2})\s*[%％]', p)
            if pct:
                all_nums.append(('tax', float(pct.group(1)), pi))
                continue
            nm = re.search(r'([\d,]+\.?\d+)', p.replace(',','').replace('，',''))
            if not nm:
                continue
            v = float(nm.group(1))
            # Skip spec-related numbers (resolution, IP ratings, etc.) unless marked with 元
            is_spec = bool(_SPEC_PAT.search(p))
            has_yuan = '元' in p
            is_pure_num = re.match(r'^\s*[\d,]+\.?\d*\s*(?:元)?\s*$', p) is not None

            if is_spec and not has_yuan:
                # Spec numbers: classify as count if small, ignore otherwise
                if v < 1000 and v == int(v):
                    all_nums.append(('count', v, pi))
                # else: ignore (false price from spec text)
            elif v >= 100 or has_yuan:
                all_nums.append(('price', v, pi))
            else:
                all_nums.append(('count', v, pi))

        if len(all_nums) < 2: continue
        counts = [(v, pi) for t, v, pi in all_nums if t == 'count' and 1 <= v <= 999 and v == int(v)]
        taxes = [(v, pi) for t, v, pi in all_nums if t == 'tax']
        prices = [(v, pi) for t, v, pi in all_nums if t == 'price' and v >= 100]

        # Prefer count from pure-number cells over seq numbers (1, 2, 3...)
        if counts:
            pure_counts = [(v, pi) for v, pi in counts
                          if re.match(r'^\s*[\d,]+\s*$', parts[pi].strip())]
            if pure_counts:
                counts = pure_counts
            # Sort by position: prefer counts that appear AFTER the first column
            counts.sort(key=lambda x: x[1])
            # Heuristic: if there's a value >= 10 and a tiny value (1-9), prefer the larger
            big = [(v, pi) for v, pi in counts if v >= 10]
            if big:
                counts = big
            # Skip common tax-rate values masquerading as count (13, 6, 9, 3, 17)
            counts = [(v, pi) for v, pi in counts if v not in (3, 6, 9, 13, 17)]
        if len(prices) < 2: continue

        sp = sorted(prices, key=lambda x: x[0])
        n = len(sp)
        item = {
            'priceName': name, 'unit': '项',
            'count': int(counts[0][0]) if counts else 1,
            'tax': (str(int(taxes[0][0])) + '%') if taxes else None,
            'unitPrice': sp[0][0],
            # n>=4: [unit, total_ex, ..., total_in] -> 不含税=次大(sp[-2])
            # n==3: [unit, total_ex, total_in]     -> 不含税=中间值(sp[1])
            # n==2: [unit, total]                  -> 不含税=较小值(sp[0])
            'totalPrice': sp[-2][0] if n >= 4 else (sp[1][0] if n >= 3 else sp[0][0]),
            'totalPriceInTax': sp[-1][0] if n >= 3 else sp[1][0],
            'extras': {}, 'details': []
        }
        if item['totalPrice'] and item['totalPrice'] >= 100:
            items.append(item)
    return items


def _filter_price_items(items):
    """Remove non-price items (personnel, projects, tech specs, manufacturers)."""
    BAD = re.compile(
        r'(?:经理|工程师|工人|主任|主管|专员|总监|总裁|董事长|秘书|助理|'
        r'合同|协议|订单|项目\s*名称|供应商|投标人|采购人|'
        r'灵敏度|dB|MHz|GHz|指标\s*要求|功能\s*要求|'
        r'验收测试|测试评审|联通测试|差旅|交通|住宿|会议内容|出差|'
        r'^其他$|^无$|^备注$|^说明$|^小计$|'
        r'硬件费用|软件费用|其他费用|'
        # Manufacturer/company names misidentified as price items
        r'(?:有限公司|有限责任|公司|集团|大学|学院|研究所|'
        r'(?:科技|技术|电子|光电|仪器|测量|通信|网络|半导体|激光)'
        r'.{0,5}(?:中国|日本|美国|德国|英国|法国|意大利|加拿大|澳大利亚|'
        r'马来西亚|新加坡|韩国|越南|印度|泰国|台湾|香港|澳门|'
        r'北京|上海|深圳|广州|成都|武汉|南京|杭州|西安)))')
    valid = []
    for item in items:
        if BAD.search(item.get('priceName', '')): continue
        up = item.get('unitPrice') or 0
        tp = item.get('totalPrice') or 1
        if up > tp * 1.5: continue
        valid.append(item)
    return valid

# ── Text Similarity ─────────────────────────────────────────────
def _build_normalized_map(text):
    """Return (norm_text, positions) where norm_text has whitespace and
    invisible characters dropped and case/punctuation folded (see
    _FOLD_TABLE), and positions[i] is the original index of norm_text[i]."""
    text = _strip_toc_dots(text)  # TOC leader dots never participate in matching
    norm_chars = []
    positions = []
    for i, ch in enumerate(text):
        if not _is_skip_char(ch):
            norm_chars.append(ch)
            positions.append(i)
    return ''.join(norm_chars).translate(_FOLD_TABLE), positions


def _count_meaningful_chars(segment):
    """Count CJK ideographs, ASCII letters, and digits — the character
    classes that signal substantive content (not boilerplate punctuation)."""
    return len(re.findall(r'[一-鿿A-Za-z0-9]', segment))


def find_common_segments(text1, text2, min_len=15):
    """Find substrings >= min_len chars that appear in both texts.

    Uses a k-gram (min_len-gram) hash index with bidirectional maximal
    extension -- O(n+m) expected -- instead of full-text SequenceMatcher.
    SequenceMatcher's get_matching_blocks() is O(n*m) on the long,
    highly-repetitive documents produced by OCR (hundreds of thousands of
    normalised characters with thousands of matching regions), where it can
    run for minutes or never finish. The k-gram index finds the same maximal
    exact common substrings in a fraction of a second.

    Matches are whitespace-insensitive (comparison runs on the
    _build_normalized_map output). The minimum-meaningful-chars check uses a
    combined CJK + ASCII-letter + digit count so English/mixed technical
    documents get the same recall as Chinese-only ones.

    Returns list of (pos1, pos2, length, segment, ctx1, ctx2) where pos1/pos2
    are raw indices into the original texts, length is the normalised length,
    and segment/ctx are raw-text slices.
    """
    norm1, pos1_map = _build_normalized_map(text1)
    norm2, pos2_map = _build_normalized_map(text2)

    L1, L2 = len(norm1), len(norm2)
    if L1 < min_len or L2 < min_len:
        return []

    k = min_len
    # Index every k-gram of the shorter text by starting position to keep the
    # index small; scan the other text. Swap roles so norm1 is always the
    # indexed (shorter) text and norm2 the scanned (longer) one, then map
    # positions back through the (possibly swapped) pos maps and raw texts.
    swapped = L1 > L2
    if swapped:
        norm1, norm2 = norm2, norm1
        pos1_map, pos2_map = pos2_map, pos1_map
        text1, text2 = text2, text1
        L1, L2 = L2, L1
    index = {}
    for i in range(L1 - k + 1):
        index.setdefault(norm1[i:i + k], []).append(i)

    min_meaningful = max(5, k // 2)
    # Bytearrays mark normalised positions already consumed by an emitted
    # match, so we never report overlapping/duplicate segments (the original
    # approximated this with a 10-char raw-position proximity check).
    claimed1 = bytearray(L1)
    claimed2 = bytearray(L2)
    results = []
    # Cap candidates examined per k-gram to bound work on popular boilerplate
    # grams; real bid prose rarely repeats a 15-gram more than a few dozen
    # times, so this cap is almost never reached.
    MAX_CANDS = 96

    i2 = 0
    while i2 <= L2 - k:
        gram = norm2[i2:i2 + k]
        cands = index.get(gram)
        if not cands:
            i2 += 1
            continue
        # Among the unclaimed candidate starts, pick the one yielding the
        # longest maximal extension at this norm2 position.
        best_len = 0
        best_a = -1
        best_b = -1
        checked = 0
        for i1 in cands:
            if checked >= MAX_CANDS:
                break
            checked += 1
            if claimed1[i1]:
                continue
            # Extend forward while normalised chars match.
            f = 0
            while (i1 + k + f < L1 and i2 + k + f < L2
                   and norm1[i1 + k + f] == norm2[i2 + k + f]):
                f += 1
            # Extend backward while normalised chars match.
            b = 0
            while (i1 - b - 1 >= 0 and i2 - b - 1 >= 0
                   and norm1[i1 - b - 1] == norm2[i2 - b - 1]):
                b += 1
            total = k + f + b
            if total > best_len:
                best_len = total
                best_a = i1 - b
                best_b = i2 - b
        if best_len < k:
            i2 += 1
            continue
        # Skip if either end of the maximal match already sits inside a
        # previously emitted segment (overlap). The norm2 end cannot overlap
        # because we advance past emitted ranges, but check defensively.
        if claimed1[best_a] or claimed2[best_b]:
            i2 += 1
            continue
        seg_norm = norm1[best_a:best_a + best_len]
        if _count_meaningful_chars(seg_norm) < min_meaningful:
            i2 = best_b + best_len
            continue
        # Claim both normalised ranges so later matches cannot overlap.
        claimed1[best_a:best_a + best_len] = b'\x01' * best_len
        claimed2[best_b:best_b + best_len] = b'\x01' * best_len
        # Map back to raw positions. The match spans skipped
        # whitespace/invisible chars in the raw text, so map the last
        # normalised char back to its raw index for the true tail.
        raw_a = pos1_map[best_a]
        raw_b = pos2_map[best_b]
        raw_a_end = pos1_map[best_a + best_len - 1] + 1
        raw_b_end = pos2_map[best_b + best_len - 1] + 1
        seg_raw = text1[raw_a:raw_a_end]
        ctx_before = 100
        ctx_after = 100
        ctx1 = text1[max(0, raw_a - ctx_before):raw_a_end + ctx_after]
        ctx2 = text2[max(0, raw_b - ctx_before):raw_b_end + ctx_after]
        results.append((raw_a, raw_b, best_len, seg_raw, ctx1, ctx2))
        # Advance past this match in the scanned text.
        i2 = best_b + best_len

    return results


def is_template_content(text):
    """Check if text is likely a standard template/bid instruction phrase.
    Expanded to cover all lengths and common bid document boilerplate patterns."""
    text_stripped = text.strip()

    # ── Category 1: Signature / seal / date boilerplate (any length) ──
    signature_markers = [
        '供应商名称', '法定代表人或授权代表签字', '法定代表人签字',
        '授权代表签字', '法定代表人（签字）', '授权委托人（签字）',
        '法定代表人或其委托代理人', '法定代表人或委托代理人',
        '项目编号', '项目名称', '注：',
        '公章', '供应商全称', '盖单位章', '签字或盖章',
        '（单位公章）', '（盖章）', '签字或印章',
        '法定代表人盖章', '委托代理人签字',
        '日期：', '年 月 日', '年月日',
        '供应商（公章）', '供应商：（盖章）',
        '投标人名称', '投标人（盖章）', '投标人全称',
        '投标人地址', '投标人电话', '投标人传真',
        '联系人：', '联系电话：', '传真：',
        '开户银行：', '账号：', '银行账号：',
        '纳税人识别号：', '统一社会信用代码：',
    ]
    for marker in signature_markers:
        if marker in text_stripped:
            return True

    # ── Category 2: Bid announcement fixed phrases ──
    announcement_phrases = [
        '投标人须知', '投标人须知前附表', '投标人须知正文',
        '招标文件的获取', '招标文件获取方式', '招标文件获取时间',
        '投标文件的递交', '投标文件递交截止', '投标文件递交地点',
        '投标截止时间', '开标时间', '开标地点',
        '发布公告的媒介', '本招标公告在', '本次招标公告在',
        '投标保证金', '投标保证金的金额', '投标保证金的形式',
        '评标办法', '评标委员会', '评标办法前附表',
        '资格审查办法', '资格审查方式', '资格后审', '资格预审',
        '踏勘现场', '不组织踏勘现场', '招标代理机构',
        '电子投标文件', '电子招标投标', '电子招标文件',
        '招标条件', '项目概况与招标范围',
        '投标人资格要求', '投标人应具备', '本次招标不接受联合体',
        '本次招标接受联合体', '联合体投标',
    ]
    for phrase in announcement_phrases:
        if phrase in text_stripped:
            return True

    # ── Category 3: Legal / standard clause phrases ──
    legal_phrases = [
        '根据《中华人民共和国招标投标法》',
        '依据《中华人民共和国招标投标法》',
        '根据《中华人民共和国政府采购法》',
        '符合《政府采购法》',
        '根据《中华人民共和国招标投标法实施条例》',
        '依据《招标投标法实施条例》',
        '信用中国', '中国政府采购网', '失信被执行人',
        '重大税收违法案件当事人', '政府采购严重违法失信行为记录名单',
        '信用信息查询', '信用记录查询',
        '行贿犯罪档案查询', '无行贿犯罪记录',
        '本招标项目', '招标项目', '招标编号',
    ]
    for phrase in legal_phrases:
        if phrase in text_stripped:
            return True

    # ── Category 4: Standard declaration / formal language ──
    declaration_phrases = [
        '我公司郑重承诺', '我单位郑重承诺', '本公司郑重声明',
        '具有独立承担民事责任的能力',
        '具有良好的商业信誉和健全的财务会计制度',
        '具有履行合同所必需的设备和专业技术能力',
        '有依法缴纳税收和社会保障资金的良好记录',
        '近三年内在经营活动中没有重大违法记录',
        '在参加政府采购活动前三年内',
        '法律、行政法规规定的其他条件',
        '具有独立法人资格', '独立承担民事责任',
        '不是联合体投标', '非联合体投标',
        '单位负责人为同一人或者存在直接控股',
        '管理关系的不同供应商',
        '为本项目提供整体设计、规范编制',
        '不得同时参加本项目',
        '中小企业声明函', '残疾人福利性单位声明函',
        '监狱企业证明文件',
    ]
    for phrase in declaration_phrases:
        if phrase in text_stripped:
            return True

    # ── Category 5: Generic numbering / project info (regex) ──
    if re.search(r'(项目|采购|招标|工程|标段)\s*(编号|代码|名称)[：:]', text_stripped):
        return True
    if re.search(r'(包号|标段|包件|分包)\s*[：:]\s*', text_stripped):
        return True
    if re.search(r'^[第].{1,4}[章节条款]', text_stripped):
        return True

    # ── Category 6: TOC / separator / page number ──
    if re.match(r'^\s*(目\s*录|目录|TOC|Table of Contents)\s*$', text_stripped, re.IGNORECASE):
        return True
    if re.match(r'^\s*[0-9IVX]+\s*$', text_stripped):  # Pure page number / roman numeral
        return True
    if re.match(r'^[-=＿.]{5,}$', text_stripped):  # Separator line
        return True

    # ── Category 7: Pure boilerplate density check ──
    # If text is long enough but dominated by boilerplate language patterns
    if len(text_stripped) >= 40:
        boilerplate_keywords = [
            '应当', '必须', '不得', '严禁', '应具备', '须具备',
            '承诺', '保证', '保证其', '确保', '遵守',
            '递交', '送达', '提交', '受理', '备案',
            '规定', '要求', '条件', '资格',
        ]
        bp_count = sum(text_stripped.count(kw) for kw in boilerplate_keywords)
        bp_density = bp_count / max(len(text_stripped), 1)
        # Very high boilerplate density with no concrete data → template
        if bp_density > 0.06 and not re.search(r'\d{2,}', text_stripped):
            return True

    return False


def _build_global_template_index(texts_dict):
    """Build a global index of text segments that appear across many bid files.
    Segments appearing in >= max(3, ceil(N*50%)) files are considered template.

    Uses sliding window to discover common segments without relying on pre-defined rules.

    Returns:
        set of normalized segment strings that are global templates
    """
    N = len(texts_dict)
    # Global co-occurrence filtering only makes sense for >= 3 files.
    # With N=2 every shared segment appears in 100% of files, so the
    # filter would remove exactly the pairwise evidence we want to
    # compare (per the design spec: "N=2 时两两重复仍保留比对"). For
    # N>=3, a segment in >= max(3, ceil(N*50%)) files is treated as
    # field-wide boilerplate.
    if N < 3:
        return set()

    threshold = max(3, int(N * 0.5 + 0.999))  # ceil(N * 50%)
    if threshold > N:
        threshold = N

    window_size = 100
    step = 50

    # {normalized_segment: set of filenames containing it}
    segment_files = {}

    for fname, text in texts_dict.items():
        if not text:
            continue
        # Track what we've already indexed from this file to avoid duplicates
        seen_in_file = set()
        for start in range(0, max(0, len(text) - window_size + 1), step):
            segment = text[start:start + window_size]
            # Normalize: collapse whitespace, unify punctuation
            norm = _normalize_for_match(segment)
            if len(norm) < 30:  # Too short to be meaningful
                continue
            if norm in seen_in_file:
                continue
            seen_in_file.add(norm)
            if norm not in segment_files:
                segment_files[norm] = set()
            segment_files[norm].add(fname)

    # Collect segments that appear in >= threshold files
    global_templates = set()
    for norm_seg, files in segment_files.items():
        if len(files) >= threshold:
            global_templates.add(norm_seg)

    return global_templates


def _is_in_global_template(segment, global_templates):
    """Check if a segment (or any substantial part) matches a global template."""
    if not global_templates or not segment:
        return False
    seg_norm = _normalize_for_match(segment)
    if len(seg_norm) < 12:
        return False
    # Direct match (O(1) set lookup — covers the vast majority of hits)
    if seg_norm in global_templates:
        return True
    # Substring containment scan — capped so a very large template set (5000+
    # entries from many large files) does not cause a second of latency per
    # segment on the fallback path.
    _MAX_SCAN = 1000
    count = 0
    for tmpl in global_templates:
        if len(seg_norm) >= 20 and (seg_norm in tmpl or tmpl in seg_norm):
            return True
        count += 1
        if count >= _MAX_SCAN:
            break
    return False


def _score_substantiality(text):
    """Score how 'substantial' a text segment is (0-1).
    High score = likely real collusion content (technical, specific, concrete).
    Low score = likely template/boilerplate even if not caught by rule filters.

    Dimensions:
        - Length (30%): longer segments are more likely substantial
        - Technical term density (30%): model numbers, tech jargon
        - Numeric specificity (25%): amounts, percentages, dates, version numbers
        - Boilerplate language penalty (15%): inverse score for formal phrasing
    """
    if not text or len(text) < 15:
        return 0.0

    score = 0.0

    # ── Dimension 1: Length (30%) ──
    len_score = min(1.0, len(text) / 200.0)
    score += len_score * 0.30

    # ── Dimension 2: Technical / business term density (30%) ──
    tech_patterns = [
        r'[A-Z]{2,}[-–][0-9]{2,}',        # ISO-9001, GB-2020
        r'[A-Z][A-Z0-9\-]{3,}',             # Technical model numbers
        r'[0-9]+[×xX][0-9]+',              # Dimensions: 100x200
        r'[0-9]+(\.[0-9]+)?[mMkK]?[WwVvAaHhZz]',  # Units: 220V, 5kW
        r'(毫米|厘米|米|千米|克|千克|吨|升|毫升|平方米|立方米|公顷)',  # Chinese units
        r'(台|套|件|个|组|批|项|次|人|天|月|年)',  # Counting units
    ]
    tech_chars = 0
    for pat in tech_patterns:
        for m in re.finditer(pat, text):
            tech_chars += m.end() - m.start()
    tech_density = min(1.0, tech_chars / max(len(text), 1) / 0.15)
    score += tech_density * 0.30

    # ── Dimension 3: Numeric specificity (25%) ──
    numeric_patterns = [
        r'\d{2,}\.\d{2,}',                   # Decimal amounts
        r'\d{1,3}(,\d{3})+(\.\d+)?',         # Formatted numbers
        r'[¥￥]\s*\d[\d,.]*',                # Currency amounts
        r'\d+\.\d+%',                         # Percentages
        r'20\d{2}[-/年]\d{1,2}[-/月]\d{1,2}', # Dates
        r'[0-9]{4,}',                         # Large numbers (amounts, codes)
    ]
    numeric_hits = 0
    for pat in numeric_patterns:
        numeric_hits += len(re.findall(pat, text))
    numeric_score = min(1.0, numeric_hits / max(len(text) / 50, 1))
    score += numeric_score * 0.25

    # ── Dimension 4: Boilerplate language penalty (15%) ──
    boilerplate_kw = [
        '应当', '必须', '不得', '严禁', '遵守', '执行',
        '保证', '承诺', '确保', '承担', '履行', '提供',
        '规定', '要求', '条件', '标准', '规范',
        '递交', '送达', '提交',
    ]
    bp_count = sum(text.count(kw) for kw in boilerplate_kw)
    bp_density = bp_count / max(len(text), 1)
    # High BP density + low specificity → penalty
    has_concrete = bool(re.search(r'\d{2,}', text)) or bool(re.search(r'[A-Z]{2,}', text))
    if has_concrete:
        bp_penalty = min(1.0, bp_density / 0.10) * 0.4  # Reduced penalty if has concrete data
    else:
        bp_penalty = min(1.0, bp_density / 0.06)  # Full penalty
    score += (1.0 - bp_penalty) * 0.15

    return round(score, 3)


def classify_abnormal_reason(text):
    """Classify why a text match is abnormal"""
    reasons = []
    if re.search(r'[A-Z][A-Z0-9\-]+', text):  # Contains technical model numbers
        reasons.append('包含具体技术型号/参数')
    if any(w in text for w in ['健壮性', '更好的前所未有', '改进的联系']):
        reasons.append('含有非标准机器翻译痕迹')
    if re.search(r'\d+\.\d+\.\d+', text):
        reasons.append('包含不规范编号标记')
    if len(text) > 80 and not is_template_content(text):
        reasons.append('长段落逐字相同，排除独立编制的可能性')
    return reasons if reasons else ['内容异常一致']

def _is_in_reference(segment, ref_texts):
    """Check if a text segment appears in any reference document.
    Uses normalized comparison (whitespace-insensitive) for accuracy."""
    if not ref_texts or not segment:
        return False
    seg_norm = _normalize_for_match(segment)
    if len(seg_norm) < 12:
        return False
    for rt in ref_texts:
        if seg_norm in _normalize_for_match(rt):
            return True
    return False

def _split_paragraphs(text):
    """Split text into (raw, normalized) paragraphs of meaningful size.
    Splits on newlines; keeps chunks whose normalized length is 40-500 so
    near-duplicate comparison runs on substantive prose, not table rows."""
    if not text:
        return []
    out = []
    for p in re.split(r'\n+', text):
        p = p.strip()
        if not p:
            continue
        n = _normalize_for_match(p)
        if 40 <= len(n) <= 500 and not is_template_content(p):
            out.append((p, n))
    return out


def _find_near_duplicate_paragraphs(text1, text2, min_ratio=0.80, max_ratio=0.98):
    """Find paragraph pairs that are near-duplicates (high but not exact
    similarity) across two texts.

    Catches collusion where one bidder lightly edited the other's text:
    small edits break difflib's exact match blocks, so the paragraph
    escapes find_common_segments() even though it is ~90% identical.

    Bounded: paragraphs are 40-500 normalized chars, capped to 120/file
    (longest first), blocked by length bucket, then cheap quick_ratio()
    pre-filters before the full ratio() confirmation. Returns list of
    {seg1, seg2, ratio}.
    """
    import difflib
    paras1 = _split_paragraphs(text1)
    paras2 = _split_paragraphs(text2)
    if not paras1 or not paras2:
        return []
    # Keep the 120 longest substantial paragraphs per file.
    paras1 = sorted(paras1, key=lambda x: len(x[1]), reverse=True)[:120]
    paras2 = sorted(paras2, key=lambda x: len(x[1]), reverse=True)[:120]

    # Bucket file2 paragraphs by length//25 for length blocking.
    buckets = {}
    for idx, (raw, norm) in enumerate(paras2):
        buckets.setdefault(len(norm) // 25, []).append((idx, raw, norm))

    results = []
    for raw1, norm1 in paras1:
        b = len(norm1) // 25
        candidates = []
        for db in (b - 1, b, b + 1):
            candidates.extend(buckets.get(db, []))
        best = None
        checked = 0
        for _idx, raw2, norm2 in candidates:
            if checked >= 40:
                break
            ln1, ln2 = len(norm1), len(norm2)
            if abs(ln1 - ln2) > max(ln1, ln2) * 0.3:
                continue
            checked += 1
            sm = difflib.SequenceMatcher(None, norm1, norm2, autojunk=False)
            if sm.quick_ratio() < min_ratio:
                continue
            ratio = sm.ratio()
            if min_ratio <= ratio < max_ratio:
                if best is None or ratio > best[0]:
                    best = (ratio, raw2)
        if best is not None:
            results.append({'seg1': raw1, 'seg2': best[1], 'ratio': round(best[0], 3)})
        if len(results) >= 30:
            break
    return results


def text_similarity_analysis(texts_dict, ref_texts_list=None, on_progress=None,
                             cancel_event=None):
    """Full text similarity analysis across all uploaded files.
    ref_texts_list: list of text strings from reference/template documents to exclude.
    on_progress: optional callback(percent, detail) fired before each document
                 pair is compared, so the caller can stream per-pair progress.
    cancel_event: optional threading.Event checked between document pairs.

    Three-level classification:
      - substantial_abnormal: score >= 0.6, real collusion content (affects conclusion)
      - suspicious_template: score 0.3-0.6, ambiguous (shown but demoted)
      - template: score < 0.3 or caught by filters (fully excluded)
    """
    filenames = list(texts_dict.keys())

    # ── Step 0: Build global template index ──
    global_templates = _build_global_template_index(texts_dict)
    global_template_count = 0
    rule_template_count = 0

    results = {
        'total_pairs': 0,
        'pair_results': [],
        'findings': [],
        'all_abnormal': [],            # Retained for backward compatibility
        'substantial_abnormal': [],    # NEW: score >= 0.6, affects conclusion
        'suspicious_template': [],     # NEW: score 0.3-0.6, shown but demoted
        'template_matches': 0,
        'global_template_count': 0,
        'rule_template_count': 0,
    }

    ref_texts = ref_texts_list or []

    total_pairs = len(filenames) * (len(filenames) - 1) // 2
    pair_idx = 0
    for i in range(len(filenames)):
        for j in range(i+1, len(filenames)):
            _check_cancelled(cancel_event)
            pair_idx += 1
            if on_progress and total_pairs > 0:
                pct = 42 + int((pair_idx - 1) / total_pairs * 38)
                on_progress(pct, f'正在比对 {filenames[i]} 与 {filenames[j]}（{pair_idx}/{total_pairs}）')
            results['total_pairs'] += 1
            t1, t2 = texts_dict[filenames[i]], texts_dict[filenames[j]]

            segments = find_common_segments(t1, t2, min_len=15)
            pair_result = {
                'file1': filenames[i],
                'file2': filenames[j],
                'total_matches': len(segments),
                'matches': [],
                'abnormal_count': 0,
                'template_count': 0,
                'substantial_count': 0,
                'suspicious_count': 0,
            }

            for idx, (pos1, pos2, length, seg_text, ctx1, ctx2) in enumerate(segments):
                # ── Filter 1: Reference document match ──
                in_ref = _is_in_reference(seg_text, ref_texts)
                if in_ref:
                    pair_result['template_count'] += 1
                    results['template_matches'] += 1
                    pair_result['matches'].append({
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'abnormal': False, 'risk_level': 'template',
                        'reasons': ['招标文件/模板内容 — 非异常一致'],
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })
                    continue

                # ── Filter 2: Rule-based template detection ──
                if is_template_content(seg_text):
                    pair_result['template_count'] += 1
                    results['template_matches'] += 1
                    results['rule_template_count'] += 1
                    pair_result['matches'].append({
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'abnormal': False, 'risk_level': 'template',
                        'reasons': ['格式模板内容'],
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })
                    continue

                # ── Filter 3: Global cross-file concurrence (Type A) ──
                if _is_in_global_template(seg_text, global_templates):
                    pair_result['template_count'] += 1
                    results['template_matches'] += 1
                    global_template_count += 1
                    pair_result['matches'].append({
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'abnormal': False, 'risk_level': 'template',
                        'reasons': ['全局模板内容 — 多份文件共现'],
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })
                    continue

                # ── Step: Substantiality scoring ──
                score = _score_substantiality(seg_text)
                reasons = classify_abnormal_reason(seg_text)

                if score >= 0.6:
                    # Real collusion-level content
                    risk_level = 'substantial'
                    is_abnormal = True
                    pair_result['substantial_count'] += 1
                    match_entry = {
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'abnormal': True, 'risk_level': 'substantial',
                        'score': score, 'reasons': reasons,
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    }
                    pair_result['matches'].append(match_entry)
                    pair_result['abnormal_count'] += 1
                    results['all_abnormal'].append({
                        'pair': f'{filenames[i]} vs {filenames[j]}',
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'reasons': reasons, 'risk_level': 'substantial',
                        'score': score,
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })
                    results['substantial_abnormal'].append({
                        'pair': f'{filenames[i]} vs {filenames[j]}',
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'reasons': reasons, 'score': score,
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })

                elif score >= 0.3:
                    # Ambiguous — suspicious but not conclusive
                    risk_level = 'suspicious'
                    pair_result['suspicious_count'] += 1
                    match_entry = {
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'abnormal': False, 'risk_level': 'suspicious',
                        'score': score, 'reasons': reasons + ['[已降级] 段落实质性评分偏低，可能为模板套话'],
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    }
                    pair_result['matches'].append(match_entry)
                    results['suspicious_template'].append({
                        'pair': f'{filenames[i]} vs {filenames[j]}',
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'reasons': reasons, 'score': score,
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })

                else:
                    # Low score — treat as template
                    risk_level = 'template'
                    pair_result['template_count'] += 1
                    results['template_matches'] += 1
                    match_entry = {
                        'index': idx + 1, 'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'abnormal': False, 'risk_level': 'template',
                        'score': score,
                        'reasons': ['[自动过滤] 实质性评分过低，判定为模板内容'],
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    }
                    pair_result['matches'].append(match_entry)

            # ── Near-duplicate paragraph detection ──
            # Catches lightly-edited collusion that escapes the exact
            # SequenceMatcher blocks above. Reported as substantial
            # evidence (feeds clause 4-a via substantial_abnormal).
            nd_idx = len(segments)
            for nd in _find_near_duplicate_paragraphs(t1, t2):
                nd_idx += 1
                nd_seg1 = nd['seg1']
                nd_seg2 = nd['seg2']
                nd_len = len(_normalize_for_match(nd_seg1))
                nd_reasons = [f'高度近似段落（相似度{int(nd["ratio"]*100)}%，仅少量字词差异），排除独立编制可能']
                nd_score = round(max(0.6, nd['ratio']), 3)
                pair_result['matches'].append({
                    'index': nd_idx, 'length': nd_len,
                    'text': sanitize_text(nd_seg1[:300]),
                    'abnormal': True, 'risk_level': 'substantial',
                    'near_duplicate': True, 'score': nd_score,
                    'reasons': nd_reasons,
                    'pos1': 0, 'pos2': 0,
                    'ctx1': sanitize_text(nd_seg1[:400]),
                    'ctx2': sanitize_text(nd_seg2[:400])
                })
                pair_result['substantial_count'] += 1
                pair_result['abnormal_count'] += 1
                pair_result['total_matches'] += 1
                nd_entry = {
                    'pair': f'{filenames[i]} vs {filenames[j]}',
                    'index': nd_idx, 'length': nd_len,
                    'text': sanitize_text(nd_seg1[:300]),
                    'reasons': nd_reasons, 'risk_level': 'substantial',
                    'near_duplicate': True, 'score': nd_score,
                    'pos1': 0, 'pos2': 0,
                    'ctx1': sanitize_text(nd_seg1[:400]),
                    'ctx2': sanitize_text(nd_seg2[:400])
                }
                results['all_abnormal'].append(nd_entry)
                results['substantial_abnormal'].append(nd_entry)

            results['pair_results'].append(pair_result)

    results['global_template_count'] = global_template_count

    # ── Generate findings ──
    total_abnormal = sum(p['abnormal_count'] for p in results['pair_results'])
    total_template = results['template_matches']
    total_substantial = len(results['substantial_abnormal'])
    total_suspicious = len(results['suspicious_template'])
    total_global = results['global_template_count']
    total_rule = results['rule_template_count']

    filter_parts = []
    if total_global > 0:
        filter_parts.append(f'{total_global} 处全局共现模板')
    if total_rule > 0:
        filter_parts.append(f'{total_rule} 处规则库模板')
    if total_template > 0:
        filter_parts.append(f'{total_template - total_global - total_rule} 处其他模板')
    if filter_parts:
        results['findings'].append(f'模板过滤: {"、".join(filter_parts)} 已排除')

    if total_substantial > 0:
        results['findings'].append(f'共发现 {total_substantial} 处可能高风险异常文本段落（高风险）')
    if total_suspicious > 0:
        results['findings'].append(f'共 {total_suspicious} 处疑似模板段落（已降级，不参与判定）')
    if total_substantial == 0 and total_suspicious == 0:
        results['findings'].append('未发现可能高风险异常文本段落（所有匹配均为模板内容）')

    if any('机器翻译' in str(r.get('reasons', [])) for r in results['substantial_abnormal']):
        results['findings'].append('存在相同的不规范翻译表述（机器翻译痕迹），排除独立编制可能')
    if any('技术型号' in str(r.get('reasons', [])) for r in results['substantial_abnormal']):
        results['findings'].append('技术方案中具体型号/参数选择一致，不属于通用技术规范')

    return results


# ── Document Structure ──────────────────────────────────────────
def extract_structure(text):
    """Extract document TOC/structure headings.

    Recognises Chinese-numeral '一、', decimal '1.', '1.1', parenthesised
    '(1)', chapter-style '第X章', section-style 'Section X / 附录X', and
    letter-numbered appendixes.
    """
    patterns = [
        r'第[一二三四五六七八九十\d]+[章节篇部分]',
        r'[一二三四五六七八九十]+[、，.]',
        r'\d+(?:\.\d+)+',             # 1.1, 1.1.1
        r'\d+[\.\s]+',                # standalone number + dot/space
        r'[（(]\d+[）)]',             # (1) or （1）
        r'[（(][A-Fa-f][）)]',        # (A) appendix
        r'Section\s+\d+',
        r'Appendix\s+[A-Fa-f]?',
        r'附录\s*[A-Fa-f\d]*',
    ]
    combined = '|'.join(f'({p})' for p in patterns)
    # Match heading pattern at line start, followed by optional heading text
    heading_re = re.compile(
        rf'^[ \t]*(?:{combined})\s*.+',
        re.MULTILINE | re.IGNORECASE
    )
    return [m.group().strip() for m in heading_re.finditer(text)][:80]


# ── Comprehensive Analysis ──────────────────────────────────────

def _build_sub_item_comparison(all_prices, filenames):
    """Build fuzzy-merged sub-item pricing comparison table."""
    # Normalize names
    def _norm_name(name):
        """Aggressive normalization for item name comparison."""
        n = name.strip()
        # Remove parenthesized/bracketed content and quotes
        n = re.sub(r'[（(][^）)]*[）)]', '', n)
        n = re.sub(r'[【\[《<][^】\]》>]*[】\]》>]', '', n)
        n = re.sub(r'["“”‘’ ]', '', n)
        # Remove punctuation and whitespace
        n = re.sub(r'[、，。；：！？\s\-–—/\\|,\.;:!?]+', '', n)
        # Full-width to half-width
        n = n.replace('０', '0').replace('１', '1').replace('２', '2').replace('３', '3').replace('４', '4')
        n = n.replace('５', '5').replace('６', '6').replace('７', '7').replace('８', '8').replace('９', '9')
        n = n.replace('Ａ', 'A').replace('Ｂ', 'B').replace('Ｃ', 'C').replace('Ｄ', 'D')
        # Common suffixes (only strip standalone suffixes, not content)
        n = re.sub(r'(及配套代码|及配套成果|及配套)$', '', n)
        n = re.sub(r'等$', '', n)
        return n.strip()

    all_items = []
    # Collect from subItemPrice (PDF table format - 分项报价)
    for fn in filenames:
        for item in all_prices.get(fn, {}).get('subItemPrice', []):
            all_items.append({
                'file': fn, 'type': '分项报价',
                'name': item.get('priceName', ''),
                'norm': _norm_name(item.get('priceName', '')),
                'count': item.get('count'), 'unitPrice': item.get('unitPrice'),
                'totalPrice': item.get('totalPrice'), 'totalPriceInTax': item.get('totalPriceInTax'),
                'tax': item.get('tax'), 'extras': item.get('extras', {})
            })
    if not all_items:
        return []

    # LCS clustering
    def _lcs_len(a, b):
        # Cap to prevent OOM when a malformed extraction produces a 5000-char
        # "name" (e.g. whole paragraph fed as priceName).  200 chars is well
        # above realistic item names and limits the DP matrix to 40k cells.
        a = a[:200]
        b = b[:200]
        m, n = len(a), len(b)
        dp = [[0]*(n+1) for _ in range(m+1)]
        best = 0
        for i in range(1, m+1):
            for j in range(1, n+1):
                if a[i-1] == b[j-1]:
                    dp[i][j] = dp[i-1][j-1] + 1
                    best = max(best, dp[i][j])
        return best

    # Jaccard similarity on 2-grams for fuzzy name matching
    def _jaccard_2gram(a, b):
        if not a or not b:
            return 0.0
        sa = set(a[i:i+2] for i in range(len(a)-1))
        sb = set(b[i:i+2] for i in range(len(b)-1))
        if not sa or not sb:
            return 0.0
        return len(sa & sb) / len(sa | sb)

    clusters = []
    used = set()
    for i, item_i in enumerate(all_items):
        if i in used: continue
        cluster = [item_i]
        used.add(i)
        best_name = item_i['name']
        for j, item_j in enumerate(all_items):
            if j in used: continue
            # Match: exact same name, or fuzzy match via LCS / Jaccard 2-gram
            same_name = item_i['norm'] == item_j['norm']
            min_len = 4
            lcs_val = _lcs_len(item_i['norm'], item_j['norm']) if len(item_i['norm']) >= min_len and len(item_j['norm']) >= min_len else 0
            jaccard_val = _jaccard_2gram(item_i['norm'], item_j['norm']) if len(item_i['norm']) >= min_len and len(item_j['norm']) >= min_len else 0
            long_match = len(item_i['norm']) >= min_len and len(item_j['norm']) >= min_len and (lcs_val >= 6 or jaccard_val >= 0.55)
            if same_name or long_match:
                cluster.append(item_j)
                used.add(j)
                if len(item_j['name']) > len(best_name):
                    best_name = item_j['name']
        clusters.append({'name': best_name, 'items': cluster})

    # Generate findings for each cluster
    for c in clusters:
        findings = []
        items = c['items']
        # 4. Only one bidder has this item (checked before the < 2 guard so it
        # actually fires; the old placement after `continue` was unreachable).
        if len(items) == 1:
            findings.append(f'仅{os.path.basename(items[0]["file"])}有此分项')
        if len(items) < 2:
            c['findings'] = findings
            continue

        prices_incl = [it['totalPriceInTax'] for it in items if it.get('totalPriceInTax')]
        prices_excl = [it['totalPrice'] for it in items if it.get('totalPrice')]
        files_involved = [it['file'] for it in items]

        # 1. All prices identical → suspicious
        if len(set(prices_excl)) == 1 and len(prices_excl) >= 2:
            findings.append(f'{len(prices_excl)}家供应商不含税报价完全一致({prices_excl[0]:,.0f}元)，可疑')
        elif len(set(prices_incl)) == 1 and len(prices_incl) >= 2:
            findings.append(f'{len(prices_incl)}家供应商含税报价完全一致({prices_incl[0]:,.0f}元)，可疑')

        # 2. Price differences analysis
        if len(prices_excl) >= 2:
            pmin, pmax = min(prices_excl), max(prices_excl)
            if pmax > 0:
                diff_pct = (pmax - pmin) / pmax * 100
                if diff_pct < 2:
                    files_with_price = [(it['file'], it['totalPrice']) for it in items if it.get('totalPrice')]
                    detail = ' | '.join(f'{os.path.basename(f)}: {p:,.0f}元' for f, p in files_with_price)
                    findings.append(f'不含税报价差异仅{diff_pct:.1f}%（{detail}），高度接近')
                elif diff_pct < 10:
                    files_with_price = [(it['file'], it['totalPrice']) for it in items if it.get('totalPrice')]
                    detail = ' | '.join(f'{os.path.basename(f)}: {p:,.0f}元' for f, p in files_with_price)
                    findings.append(f'不含税报价差异{diff_pct:.1f}%（{detail}）')

        # 3. Sequential pattern detection
        if len(prices_excl) >= 3:
            sorted_prices = sorted(prices_excl)
            gaps = [sorted_prices[i+1] - sorted_prices[i] for i in range(len(sorted_prices)-1)]
            if len(set(gaps)) == 1:
                findings.append(f'报价呈等差数列（公差{gaps[0]:,.0f}元），存在规律性差异')

        c['findings'] = findings

    return clusters




# ── Project name extraction（报告命名用；跨文档取共识，失败返回 ''）──
# 标签正则容忍 OCR/排版造成的字间空白；含"招标/采购/分包"前缀形态
_PROJECT_LABEL_RE = re.compile(
    r'(?:招\s*标|采\s*购|分\s*包)?项\s*目\s*名\s*称|(?:工\s*程|标\s*段)\s*名\s*称|项\s*目\s*名'
)
# 值中出现即视为误抓的词（真实项目名不会包含这些标签/栏目词）
_PROJECT_JUNK_WORDS = (
    '招标编号', '标段编号', '招标代理', '招标人', '投标人', '联系方式', '联系电话',
    '联系人', '开标', '评标办法', '资格预审', '招标文件', '投标文件', '响应文件',
    '磋商文件', '谈判文件', '目录', '日期',
)


def _clean_project_value(v):
    """清理捕获到的项目名值：剥引号书名号/填空下划线/截断后续标签。"""
    if not v:
        return ''
    v = str(v).split('|')[0]
    # 末尾循环剥离包裹符/冒号/填空（如 '：《XX工程》' 需多层剥离）
    strip_chars = ':：;；,，。..、· 《》"\'“”‘’　 \t'
    prev = None
    while prev != v:
        prev = v
        v = v.strip(strip_chars)
    v = re.sub(r'^[为是]+', '', v)            # "项目名称为XXX" 的"为"
    v = re.split(r'\t| {2,}|　+', v)[0]       # 封面同行多列
    m2 = _PROJECT_LABEL_RE.search(v)
    if m2 and m2.start() > 0:                 # 同行尾随下一个标签
        v = v[:m2.start()]
    v = re.sub(r'[_＿]{2,}', '', v).strip()   # 填空下划线
    prev = None
    while prev != v:
        prev = v
        v = v.strip(strip_chars)
    return v.strip()


def _valid_project_value(v):
    """项目名有效性过滤：长度/纯数字或日期值/标签词残留/点线与冒号残留。"""
    if not v or not (4 <= len(v) <= 60):
        return False
    if re.fullmatch(r'[\d\s.,，。:：\-—_/\\()（）年月日时分秒]+', v):
        return False                          # 纯数字/日期值（2026年01月01日 等）
    if any(w in v for w in _PROJECT_JUNK_WORDS):
        return False
    if re.search(r'[.。]{4,}|[_＿]{3,}', v):   # 目录点线/空白填充未剔除干净
        return False
    if ':' in v or '：' in v:                 # 值中再出现冒号多为标签串行
        return False
    return True


def _project_candidates_from_text(text, limit=5):
    """从单份文档提取项目名称候选 [(归一化key, 展示值)]，按出现顺序去重。"""
    cands = []
    seen = set()
    if not text:
        return cands
    lines = text.splitlines()
    for li, raw_line in enumerate(lines):
        line = raw_line.strip()
        if not line or len(line) > 300:
            continue
        got = ''
        if '|' in line:
            # 管道表行："项目名称 | XXX工程 | 招标编号 | ..."
            cells = [c.strip() for c in line.split('|')]
            for i in range(len(cells) - 1):
                cell_key = re.sub(r'[\s:：]+', '', cells[i])
                if _PROJECT_LABEL_RE.fullmatch(cell_key):
                    got = cells[i + 1]
                    break
        else:
            m = _PROJECT_LABEL_RE.search(line)
            if not m:
                continue
            rest = line[m.end():]
            if rest and not re.match(r'^[\s:：为是]', rest):
                continue    # 标签后无冒号/"为"直接接正文，像一般表述而非赋值
            if rest:
                got = rest
            else:
                # 标签独占一行（封面常见），值在下一非空行
                nxt = lines[li + 1].strip() if li + 1 < len(lines) else ''
                if nxt and len(nxt) <= 80 and not _PROJECT_LABEL_RE.search(nxt):
                    got = nxt
                else:
                    continue
        got = _clean_project_value(got)
        if not _valid_project_value(got):
            continue
        key = re.sub(r'[\s　]', '', got)
        if key not in seen:
            seen.add(key)
            cands.append((key, got))
            if len(cands) >= limit:
                break
    return cands


def _extract_project_name(texts_by_file):
    """跨文档投票提取项目名：≥2 份一致的候选优先，否则取票数最高者；失败返回 ''。"""
    votes = defaultdict(int)
    display = {}
    for text in texts_by_file.values():
        for key, raw in _project_candidates_from_text(text):
            votes[key] += 1
            if key not in display or len(raw) > len(display[key]):
                display[key] = raw
    if not votes:
        return ''
    best_key = max(votes, key=lambda k: (votes[k], len(display[k])))
    return display[best_key]


def run_full_analysis(filepaths, ref_filepaths=None, group_map=None, group_texts=None, on_progress=None,
                      cancel_event=None):
    """Run all analysis modules and return structured results.
    ref_filepaths: optional reference/template document paths.
    group_map: {group_name: [filepath, ...]} for multi-volume merging.
    group_texts: {group_name: combined_text} pre-merged texts.
    on_progress: callback(step, label, percent, detail) for streaming progress.
    cancel_event: optional threading.Event; when set, raises AnalysisCancelled
    at the next checkpoint (per PDF page, per xlsx row, between phases).
    """
    def _progress(step, label, percent, detail=''):
        if on_progress:
            on_progress(step, label, percent, detail)
    filenames = [os.path.basename(fp) for fp in filepaths]

    # Determine display names: use group names if available
    if group_map and len(group_map) < len(filenames):
        display_names = list(group_map.keys())
        # Map each original filename to its group
        file_to_group = {}
        for g, paths in group_map.items():
            for p in paths:
                file_to_group[os.path.basename(p)] = g
    else:
        display_names = filenames
        file_to_group = {fn: fn for fn in filenames}
        group_map = {fn: [fp] for fn, fp in zip(filenames, filepaths)}

    # Ensure group_texts is populated for every display name. Callers that
    # already extracted text (streaming endpoint) pass it in to avoid a costly
    # 100% duplicate extraction; only re-extract what is missing.
    if group_texts is None:
        group_texts = {}
    for gn in display_names:
        if gn not in group_texts:
            paths = group_map.get(gn, [])
            combined = ''
            for p in paths:
                try:
                    combined += extract_text_with_tables(
                        p, max_pages=MAX_PDF_PAGES, cancel_event=cancel_event) + '\n'
                except AnalysisCancelled:
                    raise
                except Exception:
                    pass
            group_texts[gn] = combined

    # 1. Metadata — per original file
    all_meta = {}
    for fp, fn in zip(filepaths, filenames):
        all_meta[fn] = extract_metadata(fp)

    # 2. Text — use merged per-group
    all_text = {}
    for group_name in display_names:
        all_text[group_name] = group_texts.get(group_name, '')

    _progress('text', '文本提取完成', 28, f'已提取 {len(display_names)} 份标书的文本')

    # 2b. Text extraction — reference documents
    ref_texts = []
    ref_filenames = []
    if ref_filepaths:
        for rfp in ref_filepaths:
            try:
                rt = extract_text_with_tables(
                    rfp, max_pages=MAX_PDF_PAGES, cancel_event=cancel_event)
                if rt:
                    ref_texts.append(rt)
                    ref_filenames.append(os.path.basename(rfp))
            except AnalysisCancelled:
                raise
            except Exception:
                pass

    # Use display_names (group names) for all comparison outputs
    out_names = display_names

    # 3. Personnel — extract from merged text per group
    all_personnel = {}
    for gn in out_names:
        all_personnel[gn] = extract_personnel(all_text.get(gn, ''))

    # 4. Pricing
    all_prices = {}
    for gn in out_names:
        all_prices[gn] = extract_prices(all_text.get(gn, ''))

    # 5. Text similarity — now with reference text filtering
    _progress('similarity', '文本相似度分析', 42,
              f'正在比对 {len(display_names)} 份标书的文本相似段落，大文件可能耗时…')
    similarity = text_similarity_analysis(
        all_text, ref_texts,
        on_progress=lambda pct, d='': _progress('similarity', '文本相似度分析', pct, d),
        cancel_event=cancel_event)

    # 6. Structure
    all_structure = {}
    for fn, text in all_text.items():
        all_structure[fn] = extract_structure(text)

    # 6b. 项目名称 — 跨文档共识提取（写入结果，供报告命名与展示）
    project_name = _extract_project_name(all_text)

    # Build per-group metadata (use first file's metadata)
    # Build per-group metadata: start with the first volume, then fill in
    # empty fields from subsequent volumes so multi-volume bids don't lose
    # creator/modifier ID evidence that only appears in later volumes.
    group_meta = {}
    _META_FIELDS = ('creator', 'last_modified_by', 'application', 'template',
                    'company', 'KSOTemplateDocerSaveRecord', 'KSOProductBuildVer',
                    'ICV', 'revision', 'total_edit_time')
    for gn in out_names:
        paths = group_map.get(gn, [])
        if not paths:
            continue
        meta = dict(all_meta.get(os.path.basename(paths[0]), {}))
        for p in paths[1:]:
            extra = all_meta.get(os.path.basename(p), {})
            for f in _META_FIELDS:
                if not meta.get(f) and extra.get(f):
                    meta[f] = extra[f]
        group_meta[gn] = meta

    _progress('metadata', '元数据交叉比对', 80, f'交叉比对 {len(out_names)} 份标书的创建者、修改者、编辑程序等')

    # ── Helper: filter out software/application names from metadata matching ──
    def _is_software_name(val):
        """Check if a metadata value is a software/application name, not a person."""
        if not val:
            return False
        v = str(val).strip()
        # Known software vendor/product patterns
        software_patterns = [
            r'Microsoft[®\s]*\b(Word|Office|Excel|PowerPoint|Windows)',  # Microsoft products
            r'\bWPS\b', r'Kingsoft', r'金山(?:WPS|Office|软件|办公|文档|文字|表格|演示)',
            r'Adobe[®\s]', r'Adobe\s+(Acrobat|PDF|Photoshop|Illustrator)',
            r'LibreOffice', r'OpenOffice', r'Apache\s+OpenOffice',
            r'Apple\s+(Pages|Numbers|Keynote)',
            r'iText', r'pdf\s*kit', r'wkhtmltopdf', r'FPDF', r'TCPDF',
            r'Aspose\.', r'Spire\.',
            r'打印机', r'Printer', r'Scanner',
            r'Foxit', r'Nitro\s+PDF',
            r'Ghostscript',
            r'®', r'™',  # Trademark symbols strongly suggest software
        ]
        for pat in software_patterns:
            if re.search(pat, v, re.IGNORECASE):
                return True
        # Common generic application values
        generic_apps = [
            'Microsoft Word', 'Microsoft Office', 'Microsoft Excel',
            'WPS Office', 'WPS 文字', 'WPS 表格',
            'Adobe Acrobat', 'Adobe PDF',
        ]
        v_lower = v.lower()
        for ga in generic_apps:
            if ga.lower() in v_lower:
                return True
        return False

    def _is_default_template(m):
        """Filter circumstantial metadata matches that are generic defaults
        rather than collusion signals: empty templates, Word's default
        Normal.dotm, or ubiquitous WPS build numbers. Such a single match
        alone must not push clause (一) to the "中" (suspicious) band."""
        field = (m.get('field') or '')
        value = str(m.get('value') or '').strip()
        if not value:
            return True
        # Word's default global template
        if value.lower() in ('normal.dotm', 'normal.dot', 'default', '默认'):
            return True
        # Ubiquitous template field values that every WPS/Word doc shares
        if field == '模板' and value.lower() in ('normal.dotm', 'normal.dot'):
            return True
        # WPS product build version alone (shared by every WPS install of
        # that release) only counts when paired with another signal, never
        # as the sole circumstantial item. It is excluded here so a lone
        # KSO build number does not reach the "弱" tier either.
        if field == 'WPS版本号' and re.match(r'^[\d.\-]+$', value):
            return True
        return False

    # ── Compile metadata cross-comparison (all group pairs) ──
    meta_matches = []
    if len(out_names) >= 2:
        compare_fields = [
            ('creator', '创建者'),
            ('last_modified_by', '最后保存者'),
            ('application', '编辑程序'),
            ('template', '模板'),
            ('KSOProductBuildVer', 'WPS版本号'),
            ('KSOTemplateDocerSaveRecord', 'WPS保存记录(硬件ID+用户ID)'),
            ('ICV', 'ICV'),
        ]
        all_pairs_done = set()
        for i in range(len(out_names)):
            for j in range(i+1, len(out_names)):
                gi, gj = out_names[i], out_names[j]
                mi, mj = group_meta.get(gi, {}), group_meta.get(gj, {})
                pair_label = f'{gi} ↔ {gj}'
                for key, label in compare_fields:
                    vi = mi.get(key, '')
                    vj = mj.get(key, '')
                    if vi and vj and vi == vj:
                        is_soft = _is_software_name(vi) if key in ('application', 'last_modified_by', 'creator') else False
                        dedup_key = f'{label}|{vi}'
                        if dedup_key not in all_pairs_done:
                            all_pairs_done.add(dedup_key)
                            meta_matches.append({
                                'field': label,
                                'value': str(vi)[:200],
                                'pair': pair_label,
                                'verdict': '软件名称一致（不计分）' if is_soft else '完全一致',
                                'severity': 'info' if is_soft else (
                                    'high' if key in ('KSOTemplateDocerSaveRecord', 'last_modified_by') else 'medium'
                                )
                            })

    # Time analysis
    time_findings = []
    for gn in out_names:
        for fp in group_map.get(gn, []):
            fn = os.path.basename(fp)
            m = all_meta.get(fn, {})
            if m.get('created') and m.get('modified'):
                time_findings.append(f'{gn}: 创建={m["created"]}, 修改={m["modified"]}')
                break

    # ── Compile personnel cross-comparison (all group pairs) ──
    personnel_matches = []
    personnel_dedup = set()

    # Collect all persons per file for cross-file matching
    all_persons_map = {}
    for gn in out_names:
        persons = all_personnel.get(gn, {}).get('all_persons', [])
        all_persons_map[gn] = persons

    # Environmental noise demotion BEFORE pairwise matching (see helper docstring)
    _demote_environmental_pool_values(all_personnel, out_names)

    if len(out_names) >= 2:
        for i in range(len(out_names)):
            for j in range(i+1, len(out_names)):
                gi, gj = out_names[i], out_names[j]
                pi, pj = all_personnel[gi], all_personnel[gj]
                mi, mj = group_meta.get(gi, {}), group_meta.get(gj, {})

                # ── Layer 1: Exact name match (shared personnel across files) ──
                # Normalize minority-name separators (·/•/・) so the same
                # person spelled with different middle dots still matches.
                def _nkey(n):
                    return re.sub(r'[•・]', '·', n).strip()

                names_i = {}
                for p in all_persons_map[gi]:
                    names_i.setdefault(_nkey(p['name']), p)
                names_j = {}
                for p in all_persons_map[gj]:
                    names_j.setdefault(_nkey(p['name']), p)
                shared_names = set(names_i.keys()) & set(names_j.keys())

                for name_key in shared_names:
                    name = names_i[name_key]['name']
                    role_i = names_i[name_key].get('role', 'other')
                    role_j = names_j[name_key].get('role', 'other')
                    if role_i == role_j:
                        key = f'same_person|{name}|{gi}|{gj}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            personnel_matches.append({
                                'type': '人员重叠（同角色）',
                                'detail': f'"{name}"（{role_i}）同时出现在 {gi} 和 {gj} 中',
                                'severity': 'high',
                                'role': role_i
                            })
                    else:
                        key = f'same_person_diff_role|{name}|{gi}|{gj}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            personnel_matches.append({
                                'type': '人员重叠（不同角色）',
                                'detail': f'"{name}"在{gi}中为{role_i}，在{gj}中为{role_j}',
                                'severity': 'medium'
                            })

                # ── Layer 2: Phone cross-match (ANY shared phone in pools) ──
                phones_i = pi.get('phones') or ([pi.get('phone')] if pi.get('phone') else [])
                phones_j = pj.get('phones') or ([pj.get('phone')] if pj.get('phone') else [])
                shared_phones = set(phones_i) & set(phones_j)
                for phone in shared_phones:
                    key = f'same_phone|{phone}'
                    if key not in personnel_dedup:
                        personnel_dedup.add(key)
                        personnel_matches.append({
                            'type': '联系电话相同',
                            'detail': f'{gi} 和 {gj} 联系电话均为 {phone}',
                            'severity': 'high'
                        })

                # ── Layer 3: ID number cross-match (ANY shared ID) ──
                ids_i = pi.get('id_numbers') or ([pi.get('id_number')] if pi.get('id_number') else [])
                ids_j = pj.get('id_numbers') or ([pj.get('id_number')] if pj.get('id_number') else [])
                shared_ids = set(ids_i) & set(ids_j)
                for idv in shared_ids:
                    key = f'same_id|{idv}'
                    if key not in personnel_dedup:
                        personnel_dedup.add(key)
                        personnel_matches.append({
                            'type': '身份证号相同',
                            'detail': f'{gi} 和 {gj} 出现同一身份证号 {idv[:6]}****',
                            'severity': 'critical'
                        })

                # ── Layer 3b: Email cross-match (ANY shared email) ──
                emails_i = set(pi.get('emails') or [])
                emails_j = set(pj.get('emails') or [])
                for em in emails_i & emails_j:
                    key = f'same_email|{em}'
                    if key not in personnel_dedup:
                        personnel_dedup.add(key)
                        personnel_matches.append({
                            'type': '邮箱相同',
                            'detail': f'{gi} 和 {gj} 均出现邮箱 {em}',
                            'severity': 'medium'
                        })

                # ── Layer 3c: Bank account cross-match (ANY shared account) ──
                # Same settlement account receiving/quoted by two bidders is
                # direct evidence of 资金关联 (围标团伙常用同一账户走账).
                accts_i = set(pi.get('bank_accounts') or [])
                accts_j = set(pj.get('bank_accounts') or [])
                for acct in accts_i & accts_j:
                    key = f'same_bank_acct|{acct}'
                    if key not in personnel_dedup:
                        personnel_dedup.add(key)
                        personnel_matches.append({
                            'type': '银行账号相同',
                            'detail': f'{gi} 和 {gj} 出现同一银行账号 {acct[:4]}******{acct[-4:]}',
                            'severity': 'high'
                        })

                # ── Layer 4: Auth rep vs document creator cross-match ──
                if pi.get('authorized_rep') and mj.get('creator'):
                    if pi['authorized_rep'] == mj['creator']:
                        key = f'auth_creator|{pi["authorized_rep"]}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            personnel_matches.append({
                                'type': '授权代表与创建者交叉',
                                'detail': f'{gi}的授权代表"{pi["authorized_rep"]}" = {gj}的文档创建者',
                                'severity': 'high'
                            })
                if pj.get('authorized_rep') and mi.get('creator'):
                    if pj['authorized_rep'] == mi['creator']:
                        key = f'auth_creator|{pj["authorized_rep"]}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            personnel_matches.append({
                                'type': '授权代表与创建者交叉',
                                'detail': f'{gj}的授权代表"{pj["authorized_rep"]}" = {gi}的文档创建者',
                                'severity': 'high'
                            })

                # ── Layer 4b: Authorized representative name identical ──
                # Independent trigger for clause (二): same person handling
                # bidding for different bidders, without relying on the
                # "auth-rep == creator" cross-match substring.
                ai = pi.get('authorized_rep') or ''
                aj = pj.get('authorized_rep') or ''
                if ai and aj and ai == aj and not _is_software_name(ai):
                    key = f'auth_rep_name|{ai}'
                    if key not in personnel_dedup:
                        personnel_dedup.add(key)
                        personnel_matches.append({
                            'type': '授权代表姓名相同',
                            'detail': f'{gi} 和 {gj} 的授权代表均为"{ai}"',
                            'severity': 'high'
                        })

                # ── Layer 5: Same last modifier ──
                if mi.get('last_modified_by') and mj.get('last_modified_by'):
                    if mi['last_modified_by'] == mj['last_modified_by']:
                        key = f'same_modifier|{mi["last_modified_by"]}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            is_soft = _is_software_name(mi['last_modified_by'])
                            personnel_matches.append({
                                'type': '最后修改人为同一人（软件名，不计分）' if is_soft else '最后修改人为同一人',
                                'detail': f'"{mi["last_modified_by"]}"同时为 {gi} 和 {gj} 的最后修改人',
                                'severity': 'info' if is_soft else 'high'
                            })

                # ── Layer 6: Personnel overlap rate ──
                if len(names_i) >= 2 and len(names_j) >= 2:
                    overlap = len(shared_names)
                    total = min(len(names_i), len(names_j))
                    overlap_rate = overlap / total if total > 0 else 0
                    if overlap_rate >= 0.5:
                        key = f'high_overlap|{gi}|{gj}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            personnel_matches.append({
                                'type': '人员高度重叠',
                                'detail': f'{gi} 和 {gj} 提取到的人员重叠率 {overlap_rate:.0%}（{overlap}/{total}）',
                                'severity': 'medium'
                            })

        # Phone anomaly (per file, not pairwise)
        for fn, p in all_personnel.items():
            phone = p.get('phone', '')
            if phone and len(phone) < 7:
                personnel_matches.append({
                    'type': '联系电话异常',
                    'detail': f'{fn}的联系电话为"{phone}"，不是有效电话号码格式',
                    'severity': 'medium'
                })

    _progress('personnel', '人员交叉比对', 86, f'交叉比对法定代表人、授权代表、项目成员等，发现 {len(personnel_matches)} 处异常')

    # ── Text similarity findings (summary before pricing) ──
    total_abnormal = sum(p['abnormal_count'] for p in similarity['pair_results'])

    # ── Compile pricing comparison (structured format) ──
    price_compare = {}
    price_risk_findings = []
    price_no_data_findings = []

    # Top-level comparison fields (scalar values)
    scalar_fields = [
        ('totalPriceInTax', '含税总价'),
        ('totalPrice', '不含税总价'),
        ('taxRate', '税率'),
        ('revenue', '预计收益'),
        ('cost', '预计成本'),
    ]
    for key, label in scalar_fields:
        entry = {}
        all_same = True
        first_val = None
        for gn in out_names:
            val = all_prices.get(gn, {}).get(key)
            entry[gn] = val
            # Only seed the comparison sentinel with a real (non-None) value;
            # otherwise a leading None would let [None, X, X] be reported as
            # "all same" even though one file has no data.
            if first_val is None and val is not None:
                first_val = val
            elif val is not None and val != first_val:
                all_same = False
        if any(v is not None for v in entry.values()):
            entry['_same_all'] = all_same
            price_compare[label] = entry

    # Cost detail comparison — only standard cost categories
    all_cost_names = set()
    for prices in all_prices.values():
        for item in prices.get('costDetails', []):
            all_cost_names.add(item.get('priceName', ''))
    for name in sorted(all_cost_names):
        if not name:
            continue
        entry = {}
        all_same = True
        first_val = None
        for gn in out_names:
            items = all_prices.get(gn, {}).get('costDetails', [])
            val = next((it['totalPrice'] for it in items if it.get('priceName') == name), None)
            entry[gn] = val
            # Seed only on a real value (see scalar_fields note above).
            if first_val is None and val is not None:
                first_val = val
            elif val is not None and val != first_val:
                all_same = False
        if any(v is not None for v in entry.values()):
            entry['_same_all'] = all_same
            price_compare[name] = entry
    # Pairwise findings
    if len(out_names) >= 2:
        for i in range(len(out_names)):
            for j in range(i+1, len(out_names)):
                pi = all_prices[out_names[i]]
                pj = all_prices[out_names[j]]
                # Compare totals
                if pi.get('totalPrice') and pj.get('totalPrice') and pi['totalPrice'] == pj['totalPrice']:
                    price_risk_findings.append(f'{out_names[i]} 和 {out_names[j]} 不含税总价一致: {pi["totalPrice"]:,.0f}元')
                if pi.get('totalPriceInTax') and pj.get('totalPriceInTax') and pi['totalPriceInTax'] == pj['totalPriceInTax']:
                    price_risk_findings.append(f'{out_names[i]} 和 {out_names[j]} 含税总价一致: {pi["totalPriceInTax"]:,.0f}元')

        for gn, prices in all_prices.items():
            has_total = prices.get('totalPrice') is not None or prices.get('totalPriceInTax') is not None
            has_details = prices.get('costDetails') or prices.get('subItemPrice')
            if not has_total and not has_details:
                price_no_data_findings.append(f'{gn}未提取到任何报价/成本信息')
            elif not has_total and has_details:
                price_no_data_findings.append(f'{gn}仅提取到成本明细，未提取到总价')

    _progress('pricing', '报价分析', 92, f'比较含税总价、不含税总价、分项单价等')

    # ── Sub-item comparison: compute once, feed strong regularity
    #    findings (identical sub-prices / arithmetic progression /
    #    near-identical) into clause 4-b. Without this the等差数列
    #    detection in _build_sub_item_comparison was dead code w.r.t.
    #    the verdict. '仅X有此分项' coverage notes are excluded. ──
    sub_item_compare = _build_sub_item_comparison(all_prices, out_names)
    _STRONG_PRICE_KEYS = ('完全一致', '等差数列', '高度接近')
    for _c in sub_item_compare:
        for _f in _c.get('findings', []):
            if _f and any(_k in _f for _k in _STRONG_PRICE_KEYS) and _f not in price_risk_findings:
                price_risk_findings.append(_f)

    # ── Compile verdict ──
    # Project-management roles: a single shared name in one of these roles
    # is itself evidence for clause (三). Ordinary line roles (安全员/质量员
    # etc.) are excluded as they are frequently outsourced/coincidental.
    # NOTE: these are the ENGLISH role labels used in personnel_matches
    # (from _infer_role_label), NOT the Chinese strings. Using Chinese here
    # would silently disable this clause's same-name trigger entirely.
    PM_ROLES = {'project_manager', 'tech_lead'}
    # Clause (二) triggers: same person handling bidding affairs for
    # different bidders. The "auth-rep == creator" cross-match is included
    # here too: if bidder A's authorized representative is the document
    # creator of bidder B, that one person is both preparing B's bid (一)
    # and handling A's bidding (二) -- a single fact breaching both
    # clauses, not a double-count of one weak signal.
    CLAUSE2_TYPES = {'授权代表姓名相同', '联系电话相同', '身份证号相同', '邮箱相同', '授权代表与创建者交叉'}
    clauses = [
        {
            'clause': '第（一）项',
            'description': '不同投标人的投标文件由同一单位或者个人编制',
            'satisfied': any(m['field'] == 'WPS保存记录(硬件ID+用户ID)' for m in meta_matches) or
                         any(m['type'] == '最后修改人为同一人' for m in personnel_matches) or
                         any(m['type'] == '授权代表与创建者交叉' for m in personnel_matches),
            'evidence': []
        },
        {
            'clause': '第（二）项',
            'description': '不同投标人委托同一单位或者个人办理投标事宜',
            'satisfied': any(m.get('type') in CLAUSE2_TYPES for m in personnel_matches),
            'evidence': []
        },
        {
            'clause': '第（三）项',
            'description': '不同投标人的投标文件载明的项目管理成员为同一人',
            'satisfied': any(m['type'] == '人员高度重叠' for m in personnel_matches) or
                         any(m.get('type') == '人员重叠（同角色）' and m.get('role') in PM_ROLES
                             for m in personnel_matches),
            'evidence': []
        },
        {
            'clause': '第（四）项-a',
            'description': '投标文件异常一致',
            'satisfied': len(similarity.get('substantial_abnormal', [])) > 0,
            'evidence': []
        },
        {
            'clause': '第（四）项-b',
            'description': '投标报价异常一致或呈规律性差异',
            'satisfied': True if len(price_risk_findings) > 0 else (None if len(price_no_data_findings) > 0 else False),
            'evidence': []
        },
    ]

    # ── Data-sufficiency: detect dimensions whose input bucket is empty ──
    # A clause whose source data was entirely absent returns "无法判断"
    # (not "无"), so an all-empty analysis reaches the "数据不足" branch
    # instead of falsely reporting "未发现明显围标串标异常".
    _META_DATA_FIELDS = ('creator', 'last_modified_by', 'application',
                         'template', 'KSOTemplateDocerSaveRecord',
                         'KSOProductBuildVer', 'ICV')
    meta_bucket_empty = not any(
        any(gm.get(f) for f in _META_DATA_FIELDS)
        for gm in group_meta.values()
    )
    personnel_bucket_empty = not any(
        (all_personnel.get(gn, {}) or {}).get('all_persons')
        or (all_personnel.get(gn, {}) or {}).get('legal_rep')
        or (all_personnel.get(gn, {}) or {}).get('authorized_rep')
        for gn in out_names
    )
    text_bucket_empty = not any((all_text.get(gn) or '').strip() for gn in out_names)

    # Populate evidence
    for c in clauses:
        if c['clause'] == '第（一）项':
            if any(m['field'] == 'WPS保存记录(硬件ID+用户ID)' for m in meta_matches):
                c['evidence'].append('WPS硬件ID和用户ID完全一致，同一台设备同一账号编辑')
            if any(m['type'] == '最后修改人为同一人' for m in personnel_matches):
                c['evidence'].append('两份标书最后修改人为同一人')
            if any(m['type'] == '授权代表与创建者交叉' for m in personnel_matches):
                c['evidence'].append('一方授权代表为另一方标书创建者')
            if c['satisfied']:
                c['evidence_level'] = '强'
            elif meta_bucket_empty:
                c['evidence_level'] = '无法判断'
                c['evidence'].append('未提取到任何文档元数据，无法进行元数据比对')
            else:
                # 间接证据：过滤软件名与默认模板(如 Normal.dotm/通病WPS版本号)后的非info一致项
                circumstantial = [m for m in meta_matches
                                 if m.get('severity') != 'info' and not _is_default_template(m)]
                if len(circumstantial) >= 2:
                    c['evidence_level'] = '中'
                    c['evidence'].append(f'存在 {len(circumstantial)} 项元数据一致（模板/程序/版本等间接证据）')
                elif len(circumstantial) == 1:
                    c['evidence_level'] = '弱'
                    c['evidence'].append(f'存在 1 项元数据一致（{circumstantial[0]["field"]}），间接证据较弱')
                else:
                    c['evidence_level'] = '无'
        elif c['clause'] == '第（二）项':
            if c['satisfied']:
                c['evidence_level'] = '强'
                # 收集授权代表/电话/身份证相同等独立硬证据
                for m in personnel_matches:
                    if m.get('type') in CLAUSE2_TYPES:
                        c['evidence'].append(m.get('detail', ''))
            elif personnel_bucket_empty:
                c['evidence_level'] = '无法判断'
                c['evidence'].append('未提取到任何人员信息，无法进行人员比对')
            else:
                c['evidence_level'] = '无'
        elif c['clause'] == '第（三）项':
            if c['satisfied']:
                c['evidence_level'] = '强'
                for m in personnel_matches:
                    if m['type'] == '人员高度重叠' or (
                        m.get('type') == '人员重叠（同角色）' and m.get('role') in PM_ROLES
                    ):
                        c['evidence'].append(m.get('detail', ''))
            else:
                c['evidence_level'] = '无法判断'
                if any((all_personnel.get(gn, {}) or {}).get('all_persons') for gn in out_names):
                    c['evidence'].append('已提取项目团队人员，但未发现同名项目管理成员，重叠率未达判定阈值')
                else:
                    c['evidence'].append('标书中未明确列出项目团队成员信息，无法判断')
        elif c['clause'] == '第（四）项-a':
            substantial_count = len(similarity.get('substantial_abnormal', []))
            suspicious_count = len(similarity.get('suspicious_template', []))
            template_count = similarity.get('template_matches', 0)
            if substantial_count > 0:
                c['evidence'].append(f'共发现 {substantial_count} 处可能高风险异常文本段落（高风险）')
            if suspicious_count > 0:
                c['evidence'].append(f'共 {suspicious_count} 处疑似模板段落（已降级，不参与判定）')
            if template_count > 0:
                c['evidence'].append(f'共 {template_count} 处模板内容已过滤排除')
            if c['satisfied']:
                c['evidence_level'] = '强'
            elif suspicious_count > 0:
                c['evidence_level'] = '中'
            elif text_bucket_empty:
                c['evidence_level'] = '无法判断'
                c['evidence'].append('未提取到任何文本内容，无法进行相似度比对')
            else:
                c['evidence_level'] = '无'
        elif c['clause'] == '第（四）项-b':
            c['evidence'] = price_risk_findings + price_no_data_findings
            if c['satisfied'] is True:
                c['evidence_level'] = '强'
            elif c['satisfied'] is None:
                c['evidence_level'] = '无法判断'
            else:
                c['evidence_level'] = '无'

    num_bids = len(out_names)
    num_word = {2: '两份', 3: '三份', 4: '四份', 5: '五份', 6: '六份', 7: '七份', 8: '八份', 9: '九份', 10: '十份'}
    bid_word = num_word.get(num_bids, f'{num_bids}份')

    # ── Weighted scoring ──
    # 权重分布避免两极分化，中间分数段（25-50）由多项硬证据叠加产生
    # 第（一）项命中即达高度嫌疑线，其他硬证据叠加推高置信度
    clause_weights = {
        '第（一）项': 50,       # 硬证据: WPS ID、授权代表=创建者交叉、最后修改人同一
        '第（二）项': 25,       # 硬证据: 授权代表重叠，同一人办理投标
        '第（三）项': 15,       # 硬证据: 项目管理人员姓名重叠
        '第（四）项-a': 5,      # 软证据: 文本相似度
        '第（四）项-b': 4,      # 软证据: 报价规律
    }
    level_score_map = {'强': 1.0, '中': 0.3, '弱': 0.15, '无法判断': 0, '无': 0}

    total_score = 0
    max_score = 100
    all_uncertain = True
    for c in clauses:
        weight = clause_weights.get(c['clause'], 0)
        level = c.get('evidence_level', '无')
        multiplier = level_score_map.get(level, 0)
        c['_weight'] = weight
        c['_score'] = round(weight * multiplier, 1)
        total_score += c['_score']
        if level != '无法判断':
            all_uncertain = False

    # 软证据协同加分：文本异常一致 + 报价规律差异同时出现时额外+1
    clause_4a = next((c for c in clauses if c['clause'] == '第（四）项-a'), None)
    clause_4b = next((c for c in clauses if c['clause'] == '第（四）项-b'), None)
    synergy_bonus = 0
    if clause_4a and clause_4b and clause_4a.get('evidence_level') == '强' and clause_4b.get('evidence_level') == '强':
        synergy_bonus = 1
        total_score += synergy_bonus

    # 硬证据协同加分：同一人编制(一) + 同一人办理投标(二) 同时为"强" -> +5
    # 既证明文档同源、又证明投标事宜同人，双重确认使置信度显著提升
    # （如：A的授权代表同时是B的文档创建者，单一事实同时触犯两条）
    clause_1 = next((c for c in clauses if c['clause'] == '第（一）项'), None)
    clause_2 = next((c for c in clauses if c['clause'] == '第（二）项'), None)
    hard_synergy_bonus = 0
    if clause_1 and clause_2 and clause_1.get('evidence_level') == '强' and clause_2.get('evidence_level') == '强':
        hard_synergy_bonus = 5
        total_score += hard_synergy_bonus

    # 协同加分可能使总分超过100（如全维度强+硬协同），封顶于满分
    total_score = min(max_score, round(total_score, 1))

    # ── Three-tier conclusion ──
    if num_bids < 2:
        # 围标串标判定本质上是多份标书的交叉比对；单份文件无法判定
        conclusion = '仅上传1份标书，无法进行交叉比对，请至少上传2份标书'
        conclusion_level = 'uncertain'
    elif all_uncertain:
        conclusion = '数据不足，无法做出完整判定'
        conclusion_level = 'uncertain'
    elif total_score >= 50:
        conclusion = f'{bid_word}标书存在围标串标高度嫌疑'
        conclusion_level = 'high'
    elif total_score >= 15:
        conclusion = f'{bid_word}标书存在可疑情形，建议进一步核查'
        conclusion_level = 'medium'
    else:
        conclusion = '未发现明显围标串标异常'
        conclusion_level = 'low'

    # Also fix old hardcoded "两份" in findings
    for i, f_text in enumerate(time_findings):
        if '份标书' in f_text and '两份' in f_text:
            time_findings[i] = f_text.replace('两份标书', f'{bid_word}标书')

    _progress('verdict', '综合判定', 96, f'依据《招标投标法实施条例》第四十条判定：{conclusion}')

    return {
        'metadata': {
            'files': [{'name': fn, **all_meta[fn]} for fn in filenames],
            'matches': meta_matches,
            'findings': time_findings + list(filter(None, [
                'KSOProductBuildVer一致: 同一WPS版本' if any(m['field'] == 'KSOProductBuildVer' for m in meta_matches) else '',
                '最后保存者一致' if any(m['field'] == '最后保存者' for m in meta_matches) else '',
            ]))
        },
        'personnel': {
            'files': [{'name': gn, **all_personnel[gn]} for gn in out_names],
            'cross_matches': personnel_matches,
            'findings': [m['detail'] for m in personnel_matches]
        },
        'text_similarity': similarity,
        'pricing': {
            'files': [{'name': gn, **all_prices[gn]} for gn in out_names],
            'comparison': price_compare,
            'subItemCompare': sub_item_compare,
            'findings': price_risk_findings + price_no_data_findings
        },
        'structure': {gn: all_structure.get(gn, [])[:60] for gn in out_names},
        'ref_docs': ref_filenames,
        'project_name': project_name,
        'verdict': {
            'clauses': clauses,
            'conclusion': conclusion,
            'conclusion_level': conclusion_level,
            'score': total_score,
            'max_score': max_score,
            'synergy_bonus': synergy_bonus,
            'hard_synergy_bonus': hard_synergy_bonus,
            'scoring_rule': {
                'weights': clause_weights,
                'level_multipliers': level_score_map,
                'synergy_bonus': 1,
                'hard_synergy_bonus': 5,
                'thresholds': {'high': 50, 'medium': 15, 'low': 0}
            }
        }
    }


# ── Report Generation ────────────────────────────────────────────

# 中文角色标签（与前端 ROLE_LABELS 保持一致）
_REPORT_ROLE_LABELS = {
    'legal_rep': '法定代表人', 'authorized_rep': '授权代表',
    'project_manager': '项目经理', 'tech_lead': '技术负责人',
    'bid_contact': '投标联系人', 'team_member': '团队成员',
    'signatory': '签署人', 'other': '其他人员',
}
_SEVERITY_ORDER = {'critical': 0, 'high': 1, 'medium': 2, 'info': 3}
_SEVERITY_TEXT = {'critical': '致命', 'high': '严重', 'medium': '一般', 'info': '信息'}
_LEVEL_TEXT = {'high': '高度嫌疑', 'medium': '可疑', 'low': '无明显异常', 'uncertain': '无法判断'}

# 报告结论突出显示用色
_RED = RGBColor(0xC0, 0x00, 0x00)
_GREEN = RGBColor(0x1E, 0x7E, 0x34)
_GRAY = RGBColor(0x80, 0x80, 0x80)
_ORANGE = RGBColor(0xD9, 0x77, 0x06)
_LEVEL_COLORS = {'high': _RED, 'medium': _ORANGE, 'low': _GREEN, 'uncertain': _GRAY}
_SAT_COLORS = {True: _RED, False: _GREEN, None: _GRAY}               # 条款判定: 满足/不满足/无法判断
_SAT_TXT_COLORS = {'满足': _RED, '不满足': _GREEN, '无法判断': _GRAY}  # 判定汇总表内文字

# 发现类文本分级（按关键词归类着色；顺序即优先级）
_RISK_FINDING_KW = ('一致', '等差', '规律', '高度接近', '高风险', '不规范翻译', '型号')
_GRAY_FINDING_KW = ('未提取到', '仅提取到', '模板过滤', '已排除', '已降级', '不计分', '疑似模板')

# 按结论等级给出的总体处置建议
_REPORT_ADVICE = {
    'high': [
        '建议招标人或招标代理机构暂缓确定中标候选人/中标结果，先行复核本报告列出的各项证据；',
        '对报告中的硬证据（WPS保存记录一致、最后修改人同一、授权代表交叉、联系电话/身份证号相同等）逐项人工核实并固定原始文件；',
        '视项目监管权限向相应行政监督部门报告线索，并移交本报告及原始投标文件；',
        '必要时依法提请对涉案单位投标保证金缴纳、资金往来账户进一步调查。',
    ],
    'medium': [
        '建议对本报告命中条款的证据进行人工复核，重点核查元数据一致项与人员交叉发现；',
        '可要求相关投标人对异常一致内容、人员重叠情况作出书面澄清说明；',
        '结合投标保证金缴纳账户、投标文件相互混装等本系统未自动检测的情形补充人工查验；',
        '如复核后证据坐实，按"高度嫌疑"情形处置。',
    ],
    'low': [
        '本次分析未发现明显围标串标异常，可按正常流程推进评审工作；',
        '建议留存本报告及分析数据备查；如后续获得新证据可重新分析。',
    ],
    'uncertain': [
        '当前数据不足以做出完整判定（如仅上传1份标书，或未提取到有效文本/报价信息）；',
        '建议补齐全部投标文件（.docx/.doc/.pdf/.txt/.xlsx）及招标文件/模板后重新分析。',
    ],
}

# 按实际命中条款给出的专项处置建议（仅对"满足"的条款输出）
_REPORT_CLAUSE_ADVICE = {
    '第（一）项': '投标文件由同一单位或个人编制：封存投标文件原件，核查文档创建者、最后保存者及WPS硬件记录所指向的实际编制人，必要时调取投标单位的授权与用印台账比对；',
    '第（二）项': '委托同一单位或个人办理投标事宜：约谈相关投标人的授权代表，核验其劳动关系、社保缴纳单位与身份证明，确认是否存在同一人员或中介代办的情形；',
    '第（三）项': '项目管理成员为同一人：要求相关投标人提供拟投入项目管理人员的劳动合同与社保缴纳记录，核实人员是否真实在编在岗、是否存在同时受聘于多家投标人的情形；',
    '第（四）项-a': '投标文件异常一致：要求投标人对技术方案等异常一致内容作出书面澄清，提交独立编制过程的证明材料；',
    '第（四）项-b': '投标报价异常一致或呈规律性差异：复核各投标人报价编制依据与成本构成，逐项比对分项报价明细，核查是否存在事先合意抬价、压价或轮流中标的迹象；',
}

# 附加线索处置建议：银行账号交叉（对应条例第四十条第六项的核查方向，系统不自动判定）
_REPORT_BANK_ADVICE = ('出现同一银行账号：结合投标保证金缴纳凭证核查资金来源与流向，'
                       '确认是否存在保证金从同一单位或个人账户转出的情形（实施条例第四十条第六项）；')

_REPORT_NOTES = [
    '本报告由围串标风险识别系统基于文件元数据、人员信息、文本相似度与报价规律的自动化比对生成，仅供招标评审与监管核查参考；',
    '报告结论不构成对围标串标行为的最终认定，最终认定应以行政监督部门调查或司法机关裁判为准；',
    '元数据可能因文件流转、格式转换等原因失真，关键证据建议以原始文件复核为准；',
    '文本相似度比对已自动排除招标文件/模板等正常一致内容，个别模板性表述仍可能残留，请结合上下文判断。',
]

# 《招标投标法实施条例》第四十条（2019年修订版条文）
_REGULATION_ARTICLE_40 = [
    '（一）不同投标人的投标文件由同一单位或者个人编制；',
    '（二）不同投标人委托同一单位或者个人办理投标事宜；',
    '（三）不同投标人的投标文件载明的项目管理成员为同一人；',
    '（四）不同投标人的投标文件异常一致或者投标报价呈规律性差异；',
    '（五）不同投标人的投标文件相互混装；',
    '（六）不同投标人的投标保证金从同一单位或者个人的账户转出。',
]


def _fmt_money(v):
    """Format a money value for report tables/lines; tolerant of str/None."""
    if v is None or v == '':
        return '-'
    if isinstance(v, (int, float)):
        return f'{v:,.0f}元'
    return str(v)


def _set_style_east_font(style, name):
    """Set the East-Asian font of a style (w:eastAsia; Word 需单独指定中文字体)."""
    rpr = style.element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn('w:eastAsia'), name)


def _setup_report_styles(doc):
    """标准公文字体: 正文宋体(西文 Times New Roman)小四, 标题黑体加粗黑色."""
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    _set_style_east_font(normal, '宋体')
    for name, size in (('Title', 22), ('Heading 1', 16), ('Heading 2', 14), ('Heading 3', 12)):
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        st.font.name = 'Times New Roman'
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0, 0, 0)
        _set_style_east_font(st, '黑体')


def _conclusion_para(doc, text, color=None, size=None, underline=False, style=None):
    """加粗结论行：可选颜色/字号增大/下划线突出（不用高亮底纹）。"""
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    run.font.bold = True
    if color is not None:
        run.font.color.rgb = color
    if size is not None:
        run.font.size = Pt(size)
    if underline:
        run.font.underline = True
    return p


def _emphasize_cell(cell, color=None, bold=True):
    """Bold/color existing runs of a table cell (conclusion columns)."""
    for para in cell.paragraphs:
        for run in para.runs:
            if bold:
                run.font.bold = True
            if color is not None:
                run.font.color.rgb = color


def _finding_para(doc, text, style='List Bullet'):
    """发现类结论行：无风险绿、数据缺失与模板信息灰、风险发现红（加粗着色，不用底纹）。"""
    s = str(text)
    p = doc.add_paragraph(style=style)
    run = p.add_run(s)
    run.font.bold = True
    if '未发现' in s:
        run.font.color.rgb = _GREEN
    elif any(k in s for k in _GRAY_FINDING_KW):
        run.font.color.rgb = _GRAY
    elif any(k in s for k in _RISK_FINDING_KW):
        run.font.color.rgb = _RED
    return p


def _severity_runs(p, sev_key):
    """按严重度给【标签】+正文两个 run 着色：致命/严重红、一般橙、信息灰（整行加粗）。"""
    tag, body = p.runs[0], p.runs[1]
    tag.font.bold = True
    body.font.bold = True
    if sev_key in ('critical', 'high'):
        color = _RED
    elif sev_key == 'medium':
        color = _ORANGE
    else:
        color = _GRAY
    tag.font.color.rgb = color
    body.font.color.rgb = color


def _report_table(doc, header, rows):
    """Bordered docx table with bold header; plain table if style missing."""
    table = doc.add_table(rows=1, cols=len(header))
    try:
        table.style = 'Table Grid'
    except Exception:
        pass
    for i, text in enumerate(header):
        cell = table.rows[0].cells[i]
        cell.text = str(text)
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
    for row in rows:
        cells = table.add_row().cells
        for i, text in enumerate(row):
            cells[i].text = '-' if text is None or text == '' else str(text)
    # 表格统一五号字
    for trow in table.rows:
        for cell in trow.cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(10.5)
    return table


def _sanitize_filename_component(s, max_len=40):
    """Filename-safe component: strip Windows-forbidden chars/controls, cap length."""
    s = re.sub(r'[\\/:*?"<>|\r\n\t\x00-\x1f]', '', str(s))
    s = re.sub(r'\s+', ' ', s).strip().strip('. ').strip()
    return s[:max_len].strip()


def generate_report_docx(analysis):
    """Generate a comprehensive .docx report from analysis results.

    All field access is defensive (.get with defaults) so the report can be
    regenerated both from live results and from history records lightened by
    _prepare_history_data (which truncates match bodies and nulls per-file
    fields outside _HISTORY_KEEP).
    """
    doc = Document()
    _setup_report_styles(doc)

    meta_section = analysis.get('metadata') or {}
    personnel_section = analysis.get('personnel') or {}
    similarity_section = analysis.get('text_similarity') or {}
    pricing_section = analysis.get('pricing') or {}
    verdict_section = analysis.get('verdict') or {}

    meta_files = meta_section.get('files') or []
    personnel_files = personnel_section.get('files') or []
    pricing_files = pricing_section.get('files') or []
    meta_matches = meta_section.get('matches') or []
    cross_matches = personnel_section.get('cross_matches') or []
    pair_results = similarity_section.get('pair_results') or []
    clauses = verdict_section.get('clauses') or []
    scoring = verdict_section.get('scoring_rule') or {}
    ref_docs = analysis.get('ref_docs') or []

    score = verdict_section.get('score') or 0
    max_score = verdict_section.get('max_score') or 100
    conclusion_level = verdict_section.get('conclusion_level') or 'uncertain'
    conclusion = verdict_section.get('conclusion') or '-'
    synergy_bonus = verdict_section.get('synergy_bonus') or 0
    hard_synergy_bonus = verdict_section.get('hard_synergy_bonus') or 0
    level_text = _LEVEL_TEXT.get(conclusion_level, conclusion_level)

    # ── Title block ──
    title = doc.add_heading('围串标风险识别分析报告', level=0)
    title.alignment = 1  # center

    num_files = len(meta_files) or len(personnel_files) or len(pricing_files)
    doc.add_paragraph(f'生成时间: {datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")}')
    doc.add_paragraph('分析依据: 《中华人民共和国招标投标法实施条例》第四十条')
    doc.add_paragraph(f'分析文件: {num_files} 份，交叉比对 {len(pair_results)} 组')
    if ref_docs:
        doc.add_paragraph(
            f'模板扣除: 已上传 {len(ref_docs)} 份招标文件/模板，用于排除正常一致内容（{ "、".join(str(r) for r in ref_docs[:5]) }）')
    doc.add_paragraph('─' * 60)

    # ── 1. Executive overview ──
    doc.add_heading('一、分析概览', level=1)
    _level_color = _LEVEL_COLORS.get(conclusion_level)
    _conclusion_para(doc, f'综合风险评分: {score} / {max_score} 分（风险等级: {level_text}）',
                     color=_level_color, size=14)
    _conclusion_para(doc, f'判定结论: {conclusion}', color=_level_color, size=14, underline=True)

    satisfied_n = sum(1 for c in clauses if c.get('satisfied') is True)
    uncertain_n = sum(1 for c in clauses if c.get('satisfied') is None)
    not_n = sum(1 for c in clauses if c.get('satisfied') is False)
    bonus_parts = []
    if synergy_bonus:
        bonus_parts.append(f'软证据协同 +{synergy_bonus}分')
    if hard_synergy_bonus:
        bonus_parts.append(f'硬证据协同 +{hard_synergy_bonus}分')
    doc.add_paragraph(
        f'条款命中: 满足 {satisfied_n} 项 · 无法判断 {uncertain_n} 项 · 不满足 {not_n} 项'
        f'；协同加分: {"，".join(bonus_parts) if bonus_parts else "无"}')

    if meta_files or personnel_files:
        doc.add_heading('分析文件清单', level=2)
        file_rows = []
        for idx, f in enumerate(meta_files or personnel_files, 1):
            file_rows.append([
                idx,
                f.get('name', '-'),
                f.get('creator') or '-',
                f.get('last_modified_by') or '-',
                f.get('modified') or '-',
            ])
        _report_table(doc, ['序号', '文件名', '创建者', '最后保存者', '修改时间'], file_rows)

    doc.add_heading('各维度检查结果统计', level=2)
    total_substantial = sum(p.get('substantial_count') or 0 for p in pair_results)
    total_suspicious = sum(p.get('suspicious_count') or 0 for p in pair_results)
    severe_matches = [m for m in cross_matches if m.get('severity') in ('critical', 'high')]
    stat_rows = [
        ['元数据分析', f'一致项 {len(meta_matches)} 处（其中严重 {sum(1 for m in meta_matches if m.get("severity") == "high")} 处）'],
        ['人员及联系信息', f'交叉异常 {len(cross_matches)} 处（其中致命/严重 {len(severe_matches)} 处）'],
        ['文本相似度', f'高风险异常段落 {total_substantial} 处，疑似模板段落 {total_suspicious} 处'],
        ['报价分析', f'风险发现 {len(pricing_section.get("findings") or [])} 条'],
    ]
    _report_table(doc, ['检查维度', '结果'], stat_rows)

    # Key risks: hard evidence first (personnel/meta), then text and pricing
    key_risks = []
    for m in severe_matches:
        key_risks.append(('人员', m.get('detail', '')))
    for m in meta_matches:
        if m.get('severity') == 'high':
            pair = f'（{m.get("pair")}）' if m.get('pair') else ''
            key_risks.append(('元数据', f'{m.get("field")}: {m.get("value")}{pair}'))
    for e in (similarity_section.get('substantial_abnormal') or [])[:3]:
        key_risks.append(('文本', f'{e.get("pair")} 异常一致: {str(e.get("text", ""))[:80]}'))
    for f_text in (pricing_section.get('findings') or []):
        s = str(f_text)
        if any(k in s for k in ('一致', '等差', '高度接近')):
            key_risks.append(('报价', s))
    if key_risks:
        doc.add_heading('主要风险点摘要', level=2)
        for dim, text in key_risks[:10]:
            _conclusion_para(doc, f'[{dim}] {text}', color=_RED, style='List Bullet')

    # ── 2. Metadata ──
    doc.add_heading('二、元数据分析', level=1)
    for f in meta_files:
        doc.add_heading(f'文件: {f.get("name", "-")}', level=2)
        meta_fields = [
            ('创建者', f.get('creator', '-')),
            ('最后保存者', f.get('last_modified_by', '-')),
            ('创建时间', f.get('created', '-')),
            ('修改时间', f.get('modified', '-')),
            ('修订次数', f.get('revision', '-')),
            ('编辑时长(分钟)', f.get('total_edit_time', '-')),
            ('页数', f.get('pages', '-')),
            ('字数', f.get('words', '-')),
            ('应用程序', f.get('application', '-')),
            ('模板', f.get('template', '-')),
            ('WPS版本', f.get('KSOProductBuildVer', '-')),
            ('WPS保存记录', f.get('KSOTemplateDocerSaveRecord', '-')),
            ('ICV', f.get('ICV', '-')),
        ]
        for label, value in meta_fields:
            if value and value != '-':
                doc.add_paragraph(f'{label}: {str(value)[:200]}')

    if meta_matches:
        doc.add_heading('元数据一致项（按严重程度排序）', level=2)
        for m in sorted(meta_matches, key=lambda m: _SEVERITY_ORDER.get(m.get('severity'), 3)):
            sev = _SEVERITY_TEXT.get(m.get('severity'), '信息')
            pair = f'（{m.get("pair")}）' if m.get('pair') else ''
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(f'【{sev}】')
            p.add_run(f'{m.get("field")}: {m.get("value")}{pair} — {m.get("verdict", "")}')
            _severity_runs(p, m.get('severity'))
    else:
        _conclusion_para(doc, '未发现元数据一致项。', color=_GREEN)
    for f_text in meta_section.get('findings') or []:
        if f_text and str(f_text).strip():
            _finding_para(doc, f_text)

    # ── 3. Personnel ──
    doc.add_heading('三、人员及联系信息分析', level=1)
    for f in personnel_files:
        doc.add_heading(f'文件: {f.get("name", "-")}', level=2)
        # Multi-value contact pools (all phones / IDs / emails collected)
        _phones = '、'.join(f.get('phones') or []) or f.get('phone') or '-'
        _ids = '、'.join(f.get('id_numbers') or []) or f.get('id_number') or '-'
        _emails = '、'.join(f.get('emails') or []) or '-'
        p_fields = [
            ('公司名称', f.get('company_name', '-')),
            ('法定代表人', f.get('legal_rep', '-')),
            ('授权代表', f.get('authorized_rep', '-')),
            ('身份证号', _ids),
            ('联系电话', _phones),
            ('邮箱', _emails),
            ('地址', f.get('address', '-')),
            ('响应/投标日期', f.get('response_date', '-')),
        ]
        for label, value in p_fields:
            if value and value != '-':
                doc.add_paragraph(f'{label}: {value}')

        banks = f.get('bank_accounts') or []
        if banks:
            doc.add_paragraph('银行账号: ' + '、'.join(str(b) for b in banks[:5]))

        persons = f.get('all_persons') or []
        if persons:
            names = [
                f'{p.get("name", "")}({_REPORT_ROLE_LABELS.get(p.get("role"), p.get("role") or "人员")})'
                for p in persons[:20]
            ]
            more = f' 等共{len(persons)}人' if len(persons) > 20 else ''
            doc.add_paragraph('人员名单: ' + '、'.join(names) + more)

    if cross_matches:
        doc.add_heading('人员交叉发现（按严重程度排序）', level=2)
        for m in sorted(cross_matches, key=lambda m: _SEVERITY_ORDER.get(m.get('severity'), 3)):
            sev = _SEVERITY_TEXT.get(m.get('severity'), '信息')
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(f'【{sev}】')
            p.add_run(f'{m.get("type", "")} — {m.get("detail", "")}')
            _severity_runs(p, m.get('severity'))
    else:
        _conclusion_para(doc, '未发现人员交叉异常。', color=_GREEN)

    # ── 4. Text similarity ──
    doc.add_heading('四、文本相似度分析', level=1)
    if pair_results:
        for pr in pair_results:
            doc.add_heading(f'{pr.get("file1", "?")} vs {pr.get("file2", "?")}', level=2)
            doc.add_paragraph(f'总匹配段落数: {pr.get("total_matches", 0)}')
            doc.add_paragraph(
                f'异常一致段落数: {pr.get("abnormal_count", 0)}'
                f'（高风险 {pr.get("substantial_count", 0)}，疑似模板 {pr.get("suspicious_count", 0)}），'
                f'已过滤模板段落 {pr.get("template_count", 0)}')

            abnormal_matches = [m for m in (pr.get('matches') or []) if m.get('abnormal')]
            if abnormal_matches:
                doc.add_heading('异常一致段落详情:', level=3)
                for m in abnormal_matches[:30]:  # Limit to top 30
                    m_score = m.get('score')
                    score_str = f'，实质性评分 {m_score:.0%}' if isinstance(m_score, (int, float)) else ''
                    nd = '【高度近似】' if m.get('near_duplicate') else ''
                    reasons = '；'.join(str(r) for r in (m.get('reasons') or []))
                    reason_str = f'（{reasons}）' if reasons else ''
                    doc.add_paragraph(
                        f'{nd}第{m.get("index")}项 ({m.get("length")}字{score_str}): '
                        f'{str(m.get("text", ""))[:150]}...{reason_str}',
                        style='List Bullet'
                    )
    else:
        _conclusion_para(doc, '未进行文本相似度比对（文件数不足或未提取到有效文本）。', color=_GRAY)

    for f_text in similarity_section.get('findings') or []:
        if f_text and str(f_text).strip():
            _finding_para(doc, f_text)

    # ── 5. Pricing ──
    doc.add_heading('五、报价分析', level=1)
    if pricing_files:
        doc.add_heading('各文件报价汇总', level=2)
        price_rows = []
        for f in pricing_files:
            price_rows.append([
                f.get('name', '-'),
                _fmt_money(f.get('totalPriceInTax')),
                _fmt_money(f.get('totalPrice')),
                f.get('taxRate') or '-',
                f.get('bidRate') or '-',
            ])
        _report_table(doc, ['文件', '含税总价', '不含税总价', '税率', '费率/下浮率'], price_rows)

        warn_lines = []
        for f in pricing_files:
            for w in (f.get('warnings') or []):
                if w:
                    warn_lines.append(f'{f.get("name", "")}: {w}')
        if warn_lines:
            doc.add_heading('报价提取警示', level=2)
            for w in warn_lines[:10]:
                doc.add_paragraph(str(w), style='List Bullet')

    # Rate-based quotes (费率/下浮率) — service bids with no amount prices
    rate_lines = [f'{f.get("name")}={f.get("bidRate")}'
                  for f in pricing_files if f.get('bidRate')]
    if rate_lines:
        doc.add_paragraph('费率/下浮率报价: ' + ' | '.join(rate_lines))

    comparison = pricing_section.get('comparison') or {}
    if comparison:
        doc.add_heading('总价/成本项比对', level=2)
        for key, entry in comparison.items():
            files = [k for k in entry.keys() if k != '_same_all']
            parts = []
            for fn in files:
                val = entry[fn]
                if val is None:
                    parts.append(f'{fn}=-')
                elif isinstance(val, (int, float)):
                    parts.append(f'{fn}={val:,.0f}元')
                else:
                    parts.append(f'{fn}={str(val)}')
            mark = ' ← 全一致' if entry.get('_same_all') else ''
            if mark:
                p = doc.add_paragraph()
                p.add_run(f'{key}: {" | ".join(parts)}').font.bold = True
                mk = p.add_run(mark)
                mk.font.bold = True
                mk.font.color.rgb = _RED
            else:
                doc.add_paragraph(f'{key}: {" | ".join(parts)}')

    pricing_findings = pricing_section.get('findings') or []
    if not pricing_findings:
        _conclusion_para(doc, '未发现报价异常。', color=_GREEN)
    for f_text in pricing_findings:
        if f_text and str(f_text).strip():
            _finding_para(doc, f_text)

    # ── Sub-item comparison ──
    sub_items = pricing_section.get('subItemCompare') or []
    if sub_items:
        doc.add_heading('分项报价比对', level=2)
        for c in sub_items:
            if not c.get('items') or len(c['items']) < 2:
                continue
            doc.add_paragraph(f'{c["name"]}', style='List Bullet')
            for it in c['items']:
                parts = [os.path.basename(it.get('file', ''))]
                if it.get('totalPriceInTax'):
                    parts.append(f'含税:{it["totalPriceInTax"]:,.0f}元')
                if it.get('totalPrice'):
                    parts.append(f'不含税:{it["totalPrice"]:,.0f}元')
                if it.get('unitPrice'):
                    parts.append(f'单价:{it["unitPrice"]:,.0f}')
                if it.get('count') and it['count'] != 1:
                    parts.append(f'数量:{it["count"]}')
                doc.add_paragraph(' | '.join(parts), style='List Bullet')
            for f_text in c.get('findings', []):
                _finding_para(doc, f'→ {f_text}')

    # ── 6. Verdict & scoring ──
    doc.add_heading('六、综合判定与评分', level=1)
    if clauses:
        doc.add_heading('条款判定汇总', level=2)
        clause_rows = []
        for c in clauses:
            sat = c.get('satisfied')
            sat_str = '满足' if sat is True else ('不满足' if sat is False else '无法判断')
            clause_rows.append([
                c.get('clause', '-'),
                c.get('description', ''),
                sat_str,
                c.get('evidence_level', '无'),
                c.get('_weight', '-'),
                c.get('_score', '-'),
            ])
        clause_table = _report_table(
            doc, ['条款', '条款内容', '判定', '证据强度', '权重(分)', '得分'], clause_rows)
        # 判定/证据强度列按结论着色突出
        for trow in clause_table.rows[1:]:
            _emphasize_cell(trow.cells[2], color=_SAT_TXT_COLORS.get(trow.cells[2].text.strip()))
            if trow.cells[3].text.strip() == '强':
                _emphasize_cell(trow.cells[3], color=_RED)

        base_score = round(sum(c.get('_score') or 0 for c in clauses), 1)
        p = doc.add_paragraph()
        r = p.add_run(
            f'评分构成: 条款得分合计 {base_score:g} 分'
            + (f' + 软证据协同 {synergy_bonus} 分' if synergy_bonus else '')
            + (f' + 硬证据协同 {hard_synergy_bonus} 分' if hard_synergy_bonus else '')
            + f' = 总分 {score} 分（满分 {max_score} 分）'
        )
        r.font.bold = True

        doc.add_heading('条款证据明细', level=2)
        for c in clauses:
            sat = c.get('satisfied')
            sat_str = '满足' if sat is True else ('不满足' if sat is False else '无法判断')
            lvl = c.get('evidence_level')
            lvl_str = f'［证据强度: {lvl}］' if lvl else ''
            h = doc.add_heading('', level=3)
            h.add_run(f'{c.get("clause", "")} - {c.get("description", "")} ［')
            sat_run = h.add_run(sat_str)
            sat_run.font.bold = True
            sat_color = _SAT_COLORS.get(sat)
            if sat_color is not None:
                sat_run.font.color.rgb = sat_color
            h.add_run(f'］{lvl_str}')
            for e in c.get('evidence') or []:
                if isinstance(e, str):
                    doc.add_paragraph(f'- {e}')

    if scoring:
        doc.add_heading('评分规则', level=2)
        weights = scoring.get('weights', {})
        thresholds = scoring.get('thresholds', {})
        doc.add_paragraph(
            f'阈值: ≥{thresholds.get("high",50)}分→高度嫌疑  '
            f'{thresholds.get("medium",15)}–{thresholds.get("high",50)-1}分→可疑  '
            f'<{thresholds.get("medium",15)}分→无明显异常'
        )
        doc.add_paragraph('条款权重:', style='List Bullet')
        for clause, weight in weights.items():
            doc.add_paragraph(f'{clause}: {weight}分', style='List Bullet')
        doc.add_paragraph(
            '证据强度系数: 强=权重×1.0  中=权重×0.3  弱=权重×0.15  无法判断/无=0',
            style='List Bullet'
        )
        doc.add_paragraph(
            '软证据协同加分: 第（四）项-a 与 -b 同为"强"时 +1分',
            style='List Bullet'
        )
        doc.add_paragraph(
            '硬证据协同加分: 第（一）项 与 第（二）项 同为"强"时 +5分（文档同源+投标事宜同人双重确认）',
            style='List Bullet'
        )
        doc.add_paragraph('总分上限100分', style='List Bullet')

    # ── 7. Conclusion & advice ──
    doc.add_heading('七、结论与处置建议', level=1)
    doc.add_heading('最终结论', level=2)
    _conclusion_para(doc, conclusion, color=_level_color, size=14, underline=True)
    _conclusion_para(doc, f'综合风险评分: {score}/{max_score} 分，风险等级: {level_text}',
                     color=_level_color, size=14)

    doc.add_heading('处置建议', level=2)
    doc.add_paragraph('总体建议:', style='List Bullet')
    for a in _REPORT_ADVICE.get(conclusion_level, _REPORT_ADVICE['uncertain']):
        doc.add_paragraph(a, style='List Bullet 2')

    # 专项建议：按本次实际命中的条款与线索生成，未命中的不输出
    clause_advice = [
        _REPORT_CLAUSE_ADVICE[c['clause']]
        for c in clauses
        if c.get('satisfied') is True and c.get('clause') in _REPORT_CLAUSE_ADVICE
    ]
    if any(m.get('type') == '银行账号相同' for m in cross_matches):
        clause_advice.append(_REPORT_BANK_ADVICE)
    if clause_advice:
        doc.add_paragraph('针对本次命中情形的专项建议:', style='List Bullet')
        for a in clause_advice:
            doc.add_paragraph(a, style='List Bullet 2')

    doc.add_heading('报告使用说明', level=2)
    for note in _REPORT_NOTES:
        doc.add_paragraph(note, style='List Bullet')

    # ── Appendix: regulation text ──
    doc.add_heading('附录: 判定依据（《中华人民共和国招标投标法实施条例》第四十条）', level=1)
    doc.add_paragraph('有下列情形之一的，视为投标人相互串通投标:')
    for item in _REGULATION_ARTICLE_40:
        doc.add_paragraph(item, style='List Bullet')
    doc.add_paragraph(
        '说明: 本系统自动检测第（一）至（四）项；第（五）项（投标文件相互混装）需人工查验，'
        '第（六）项（投标保证金从同一账户转出）可结合本报告"三、人员及联系信息分析"中的银行账号交叉结果人工判断。'
    )

    # Footer disclaimer
    try:
        footer_para = doc.sections[0].footer.paragraphs[0]
        footer_para.text = '本报告由围串标风险识别系统自动生成，仅供参考'
        footer_para.alignment = 1  # center
    except Exception:
        pass

    # Save to BytesIO
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ── Routes ───────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def upload():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': '未选择文件'}), 400

    saved = []
    for f in files:
        # _safe_save strips path components, sanitizes the name, and prepends a
        # short uuid so concurrent uploads of the same original name cannot
        # overwrite each other.
        fpath, safe_name = _safe_save(f)
        if fpath:
            saved.append(safe_name)

    return jsonify({'files': saved, 'count': len(saved)})

@app.route('/api/analyze', methods=['POST'])
def analyze():
    data = request.get_json() or {}
    filenames = data.get('files', [])
    if not filenames:
        return jsonify({'error': '未指定文件'}), 400

    filepaths = []
    for fn in filenames:
        # Reject path traversal: only basenames, and the resolved path must
        # remain inside UPLOAD_FOLDER. A generic error is returned so the
        # endpoint cannot be used to probe the filesystem for other files.
        base = os.path.basename(str(fn))
        fpath = os.path.join(UPLOAD_FOLDER, base)
        if base != str(fn) or not _is_within_upload_folder(fpath) or not os.path.exists(fpath):
            return jsonify({'error': '文件不存在或无权访问'}), 404
        filepaths.append(fpath)

    try:
        results = run_full_analysis(filepaths)
        return jsonify(results)
    except Exception as e:
        logger.exception('analyze failed')
        return jsonify({'error': str(e)}), 500

@app.route('/api/report', methods=['POST'])
def report():
    data = request.get_json() or {}
    analysis_data = data.get('analysis')
    if not analysis_data:
        return jsonify({'error': '缺少分析数据'}), 400

    # Validate the analysis payload has the top-level sections the report
    # generator dereferences, so malformed/truncated input (e.g. a history
    # entry stripped for storage) yields a clear 400 instead of a 500 KeyError.
    required = ('metadata', 'personnel', 'text_similarity', 'pricing', 'verdict')
    missing = [k for k in required if k not in analysis_data]
    if missing:
        return jsonify({'error': f'分析数据结构不完整，缺少: {", ".join(missing)}'}), 400

    try:
        buf = generate_report_docx(analysis_data)
        # 文件名 = 标题 + 项目名 + 判定等级 + 时间戳；项目名做安全过滤，
        # 各段均可缺失（旧历史记录无 project_name / verdict 残缺时退化）
        parts = ['围串标风险识别分析报告']
        proj = _sanitize_filename_component(analysis_data.get('project_name') or '')
        if proj:
            parts.append(proj)
        level_text = _LEVEL_TEXT.get((analysis_data.get('verdict') or {}).get('conclusion_level'), '')
        if level_text:
            parts.append(level_text)
        parts.append(datetime.now().strftime('%Y%m%d_%H%M%S'))
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='_'.join(parts) + '.docx'
        )
    except Exception as e:
        logger.exception('report generation failed')
        return jsonify({'error': str(e)}), 500

# ── History persistence helpers (shared by analyze_stream and
#    single_upload_and_analyze) ──
# History entries are lightened for storage: text-similarity match bodies are
# truncated to snippets, and per-file detail is reduced to the keep lists
# below. When a new extraction field is added to the frontend/report, add it
# to the corresponding keep list here or it silently disappears from history
# and from report regeneration on loaded records.
_HISTORY_KEEP = {
    'metadata': ('name', 'creator', 'last_modified_by', 'created', 'modified',
                 'application', 'template', 'revision', 'total_edit_time',
                 'pages', 'words', 'company',
                 'KSOProductBuildVer', 'KSOTemplateDocerSaveRecord', 'ICV'),
    'personnel': ('name', 'legal_rep', 'authorized_rep', 'id_number', 'phone',
                  'address', 'response_date', 'company_name',
                  'phones', 'id_numbers', 'emails', 'bank_accounts',
                  'all_persons'),
    'pricing': ('name', 'totalPriceInTax', 'totalPrice', 'taxRate',
                'bidRate', 'revenue', 'cost', 'warnings'),
}


def _prepare_history_data(results):
    """Deep-copy analysis results and lighten heavy fields for history storage.

    Text-similarity match bodies are truncated to fixed lengths (the frontend
    renders a 200-char snippet + a locate modal). Per-file fields are reduced
    to _HISTORY_KEEP; other keys are nulled with their type preserved
    (str -> '' else None), matching how the frontend guards missing data.
    Keys starting with '_' are always preserved (they carry provenance flags).
    """
    history_results = json.loads(json.dumps(results, ensure_ascii=False))

    # Text-similarity matches: keep only lightweight fields + snippets
    for pr in history_results.get('text_similarity', {}).get('pair_results', []):
        light_matches = []
        for m in pr.get('matches', []):
            light_matches.append({
                'index': m.get('index'), 'length': m.get('length'),
                'text': m.get('text', '')[:200],
                'abnormal': m.get('abnormal'),
                'risk_level': m.get('risk_level'),
                'score': m.get('score'),
                'reasons': m.get('reasons', []),
                'near_duplicate': m.get('near_duplicate', False),
                'ctx1': m.get('ctx1', '')[:400],
                'ctx2': m.get('ctx2', '')[:400],
            })
        pr['matches'] = light_matches

    # Per-file fields: sub-item extras (厂家/型号) stay — small and rendered
    # by the frontend; unknown/heavy keys are nulled instead.
    for section in ('metadata', 'personnel', 'pricing'):
        for f in history_results.get(section, {}).get('files', []):
            keep = _HISTORY_KEEP[section]
            for k in list(f.keys()):
                if k not in keep and not k.startswith('_'):
                    f[k] = '' if isinstance(f[k], str) else None
    return history_results


def _write_history_entry(history_results, saved, saved_refs):
    """Serialize a lightened analysis record to HISTORY_DIR; returns its id."""
    history_id = datetime.now().strftime('%Y%m%d_%H%M%S_') + uuid.uuid4().hex[:12]
    # history_id is fully server-generated; assert its shape before it reaches
    # the filesystem so no caller-supplied component can ever become part of
    # the path. A containment check on the resolved path backs the regex up.
    history_filename = history_id + '.json'
    # 校验直接覆盖参与路径拼接的完整文件名（含 .json 后缀），
    # 保证最终写路径的文件名成分已被白名单正则完全约束。
    if not re.fullmatch(r'\d{8}_\d{6}_[0-9a-f]{12}\.json', history_filename):
        raise ValueError(f'unexpected history filename: {history_filename!r}')
    history_dir = Path(HISTORY_DIR).resolve()
    history_path = history_dir / history_filename
    # 包含性检查：resolve 后必须仍位于 HISTORY_DIR 之下
    try:
        contained = os.path.commonpath(
            [str(history_dir), str(history_path.resolve())]) == str(history_dir)
    except ValueError:            # 不同盘符等不可比较情形一律视为逃逸
        contained = False
    if not contained:
        raise ValueError('history path escaped HISTORY_DIR')
    history_entry = {
        'id': history_id,
        'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'bid_count': len(saved),
        'ref_count': len(saved_refs),
        'bid_files': [os.path.basename(s) for s in saved],
        'ref_files': [os.path.basename(s) for s in saved_refs],
        'verdict': history_results['verdict']['conclusion'],
        'abnormal_matches': sum(
            p.get('abnormal_count', 0)
            for p in history_results['text_similarity']['pair_results']),
        'template_matches': history_results['text_similarity'].get('template_matches', 0),
        'total_pairs': history_results['text_similarity'].get('total_pairs', 0),
        'data': history_results
    }
    history_path.write_text(
        json.dumps(history_entry, ensure_ascii=False), encoding='utf-8')
    return history_id


@app.route('/api/analyze_stream', methods=['POST'])
def analyze_stream():
    """Upload + analyze with streaming NDJSON progress events.
    Returns application/x-ndjson stream: each line is a JSON event.
    Final event has type='result' with the full analysis data.
    """
    files = request.files.getlist('files')
    ref_files = request.files.getlist('ref_files')

    if not files:
        return jsonify({'error': '未选择标书文件'}), 400

    saved = []
    for f in files:
        # _safe_save prepends a uuid so concurrent same-name uploads cannot
        # overwrite each other.
        fpath, _ = _safe_save(f)
        if fpath:
            saved.append(fpath)

    if len(saved) < 2:
        return jsonify({'error': '请至少上传2份标书文件(.docx/.doc/.pdf/.txt/.xlsx)'}), 400

    saved_refs = []
    for f in ref_files:
        fpath, _ = _safe_save(f, prefix='ref_')
        if fpath:
            saved_refs.append(fpath)

    # ── File size warnings ──
    size_warnings = []
    for fp in saved + saved_refs:
        fsize_mb = os.path.getsize(fp) / (1024 * 1024)
        if fsize_mb > MAX_FILE_SIZE_MB:
            size_warnings.append(
                f'"{os.path.basename(fp)}" 文件较大（{fsize_mb:.0f}MB），建议压缩后上传。'
                f'大文件处理可能较慢，请耐心等待。'
            )
    total_size = sum(os.path.getsize(fp) for fp in saved) / (1024 * 1024)
    if total_size > MAX_TOTAL_SIZE_MB:
        size_warnings.append(
            f'所有标书文件合计 {total_size:.0f}MB，处理可能需要较长时间。'
            f'建议上传可复制文字版PDF或Word文件以加速分析。'
        )

    file_groups = request.form.getlist('file_groups')
    group_map = {}
    if file_groups:
        for fp, group in zip(saved, file_groups):
            group = group.strip() or os.path.basename(fp)
            group_map.setdefault(group, []).append(fp)
    else:
        # Default: each file is its own group (no explicit grouping).
        # Matches the fallback inside run_full_analysis so the threaded
        # Phase 0 extraction actually has files to process.
        for fp in saved:
            group_map.setdefault(os.path.basename(fp), []).append(fp)

    import queue
    progress_queue = queue.Queue()

    def on_progress(step, label, percent, detail):
        progress_queue.put({'type': 'progress', 'step': step, 'label': label,
                           'percent': percent, 'detail': detail})

    # Collect extraction warnings to send as a separate event
    extraction_warnings = []

    def _on_extract_progress(phase, current, total, has_text, detail):
        """Callback for text extraction progress → sent as streaming events."""
        # Only collect "no text" final events as warnings (not every progress detail)
        if detail and phase == 'pdf_done' and not has_text:
            extraction_warnings.append(detail)
        progress_queue.put({
            'type': 'extract',
            'phase': phase,
            'current': current,
            'total': total,
            'hasText': has_text,
            'detail': detail or ''
        })

    def generate():
        import threading

        # ── Cancellation registration ──
        # The client reads request_id from the first event and calls
        # POST /api/cancel to stop OCR/extraction/analysis mid-stream.
        rid = uuid.uuid4().hex[:12]
        cancel_event = threading.Event()
        cancelled_holder = []
        with _CANCEL_LOCK:
            _CANCEL_EVENTS[rid] = cancel_event

        def _cleanup_cancel():
            with _CANCEL_LOCK:
                _CANCEL_EVENTS.pop(rid, None)

        # ── Send size warnings first ──
        for w in size_warnings:
            yield json.dumps({'type': 'warning', 'code': 'large_file', 'message': w},
                           ensure_ascii=False) + '\n'
        yield json.dumps({'type': 'ready', 'request_id': rid}, ensure_ascii=False) + '\n'

        # ── Phase 0: Text extraction (threaded so per-page PDF progress
        #    drains to the client in real time, not all at once after
        #    extraction finishes) ──
        group_texts = {}
        extraction_errors = []
        extraction_done = threading.Event()

        def extract_all():
            try:
                total_groups = len(group_map)
                for fi, (group, paths) in enumerate(group_map.items()):
                    combined = ''
                    for p in paths:
                        base = os.path.basename(p)
                        try:
                            progress_queue.put({
                                'type': 'extract', 'phase': 'start',
                                'file': base, 'group': group,
                                'fileIndex': fi + 1,
                                'totalFiles': total_groups
                            })
                            combined += extract_text_with_tables(
                                p, max_pages=MAX_PDF_PAGES,
                                on_progress=_on_extract_progress,
                                cancel_event=cancel_event
                            ) + '\n'
                        except AnalysisCancelled:
                            cancelled_holder.append(True)
                            return
                        except Exception as e:
                            msg = f'文件 {base} 文字提取失败: {e}'
                            extraction_errors.append(msg)
                            logger.warning('text extraction failed for %s: %s', base, e)
                            progress_queue.put({
                                'type': 'warning', 'code': 'extraction_failed',
                                'message': msg
                            })
                    group_texts[group] = combined
            finally:
                extraction_done.set()

        ext_thread = threading.Thread(target=extract_all, daemon=True)
        ext_thread.start()

        # Drain extraction progress events in real time — per-page PDF
        # updates now reach the client while pages are being read instead
        # of being buffered until extraction completes.
        while ext_thread.is_alive() or not progress_queue.empty():
            try:
                event = progress_queue.get(timeout=0.2)
                yield json.dumps(event, ensure_ascii=False) + '\n'
            except queue.Empty:
                pass

        ext_thread.join(timeout=5)
        if ext_thread.is_alive():
            logger.error('text extraction thread did not finish within timeout')
            yield json.dumps({'type': 'error',
                              'message': '文字提取超时，请减少文件页数或压缩后重试'},
                             ensure_ascii=False) + '\n'
            _cleanup_cancel()
            return

        if cancelled_holder:
            yield json.dumps({'type': 'cancelled',
                              'message': '分析已停止（文字提取阶段）'},
                             ensure_ascii=False) + '\n'
            _cleanup_cancel()
            return

        if extraction_warnings:
            progress_queue.put({
                'type': 'warning',
                'code': 'no_text_or_truncated',
                'messages': extraction_warnings[:10]
            })
        # Flush any remaining events (warnings, last progress ticks) before
        # starting Phase 1 so the client's bar is at the right position.
        while not progress_queue.empty():
            try:
                event = progress_queue.get(timeout=0.1)
                yield json.dumps(event, ensure_ascii=False) + '\n'
            except queue.Empty:
                break

        # ── Phase 1: Full analysis ──
        results_holder = []
        error_holder = []

        def run():
            try:
                results_holder.append(run_full_analysis(
                    saved, saved_refs if saved_refs else None,
                    group_map=group_map, group_texts=group_texts,
                    on_progress=on_progress, cancel_event=cancel_event
                ))
            except AnalysisCancelled:
                cancelled_holder.append(True)
            except Exception as e:
                logger.exception('analysis thread failed')
                error_holder.append(str(e))

        # daemon=True so a worker-killed request never leaves a zombie thread
        # holding file handles/memory; a deadline on the drain loop guarantees
        # the streaming response always terminates (gunicorn's worker timeout
        # only kills the process, not the analysis thread).
        thread = threading.Thread(target=run, daemon=True)
        thread.start()
        # < gunicorn timeout so we can flush an error before the worker kills
        # us. Default 1h: unbounded OCR on large scanned documents takes
        # minutes per file (763-page scanned bids measured ~10-30min).
        ANALYSIS_TIMEOUT = int(os.environ.get('ANALYSIS_TIMEOUT', 3600))
        deadline = time.monotonic() + ANALYSIS_TIMEOUT
        while thread.is_alive() or not progress_queue.empty():
            try:
                event = progress_queue.get(timeout=0.1)
                yield json.dumps(event, ensure_ascii=False) + '\n'
            except queue.Empty:
                pass
            if cancelled_holder:
                break
            if thread.is_alive() and time.monotonic() > deadline:
                break

        thread.join(timeout=5)
        if cancelled_holder:
            # User pressed stop — end the stream cleanly (not as an error).
            yield json.dumps({'type': 'cancelled', 'message': '分析已停止'},
                             ensure_ascii=False) + '\n'
            _cleanup_cancel()
            return

        if thread.is_alive():
            # Timed out: emit an error rather than hanging the client forever.
            logger.error('analysis thread did not finish within timeout')
            yield json.dumps({'type': 'error',
                              'message': '分析超时，请减少文件数量或文件大小后重试'},
                             ensure_ascii=False) + '\n'
            _cleanup_cancel()
            return

        if error_holder:
            yield json.dumps({'type': 'error', 'message': error_holder[0]}, ensure_ascii=False) + '\n'
            _cleanup_cancel()
            return

        results = results_holder[0]
        results['_filenames'] = [os.path.basename(s) for s in saved]
        results['_ref_filenames'] = [os.path.basename(s) for s in saved_refs]
        results['_groups'] = {g: [os.path.basename(p) for p in paths] for g, paths in group_map.items()}
        results['_group_order'] = list(group_map.keys())

        # Save to history
        try:
            history_id = _write_history_entry(
                _prepare_history_data(results), saved, saved_refs)
            results['_history_id'] = history_id
        except Exception as e:
            # History persistence is best-effort: a failed save must not
            # discard the analysis result, but the user should know it won't
            # appear in history.
            logger.warning('history save failed: %s', e)
            results['_history_save_error'] = True

        yield json.dumps({'type': 'result', 'data': results}, ensure_ascii=False, default=str) + '\n'
        _cleanup_cancel()

    return Response(generate(), mimetype='application/x-ndjson')


@app.route('/api/cancel', methods=['POST'])
def cancel_analysis():
    """Cancel a running /api/analyze_stream analysis (OCR/extraction/analysis).

    The client sends the request_id it received in the stream's first
    'ready' event; the backend sets the registered cancel Event, and the
    analysis thread raises AnalysisCancelled at its next checkpoint.
    """
    data = request.get_json(silent=True) or {}
    rid = str(data.get('id', '')).strip()
    if not rid:
        return jsonify({'ok': False, 'error': '缺少 request_id'}), 400
    with _CANCEL_LOCK:
        ev = _CANCEL_EVENTS.get(rid)
        if ev is None:
            return jsonify({'ok': False, 'error': '未找到进行中的分析'}), 404
        ev.set()
    return jsonify({'ok': True})


@app.route('/api/single_upload_and_analyze', methods=['POST'])
def single_upload_and_analyze():
    """Upload and analyze in one call.
    'files' = bid documents (required, >=2)
    'ref_files' = reference/template documents (optional, 招标文件/技术要求)
    """
    files = request.files.getlist('files')
    ref_files = request.files.getlist('ref_files')

    if not files:
        return jsonify({'error': '未选择标书文件'}), 400

    saved = []
    for f in files:
        # _safe_save prepends a uuid so concurrent same-name uploads cannot
        # overwrite each other.
        fpath, _ = _safe_save(f)
        if fpath:
            saved.append(fpath)

    if len(saved) < 2:
        return jsonify({'error': '请至少上传2份标书文件(.docx/.doc/.pdf/.txt/.xlsx)'}), 400

    saved_refs = []
    for f in ref_files:
        fpath, _ = _safe_save(f, prefix='ref_')
        if fpath:
            saved_refs.append(fpath)

    # ── File size warnings ──
    size_warnings = []
    for fp in saved + saved_refs:
        fsize_mb = os.path.getsize(fp) / (1024 * 1024)
        if fsize_mb > MAX_FILE_SIZE_MB:
            size_warnings.append(
                f'"{os.path.basename(fp)}" 文件较大（{fsize_mb:.0f}MB），处理可能较慢。'
                f'建议上传可复制文字版PDF或Word文件。'
            )
    total_size = sum(os.path.getsize(fp) for fp in saved) / (1024 * 1024)
    if total_size > MAX_TOTAL_SIZE_MB:
        size_warnings.append(
            f'所有标书文件合计 {total_size:.0f}MB，处理可能需要较长时间。'
        )

    try:
        # Get file groups — merge multi-volume files into single bidder
        file_groups = request.form.getlist('file_groups')
        group_map = {}  # group_name -> [filepath, ...]
        for fp, group in zip(saved, file_groups):
            group = group.strip() or os.path.basename(fp)
            group_map.setdefault(group, []).append(fp)

        # Merge files by group: create combined filepaths for analysis
        # Use the first file of each group as the primary, merge text internally
        group_texts = {}  # group_name -> combined_text
        extraction_warnings = []
        for group, paths in group_map.items():
            combined = ''
            for p in paths:
                try:
                    combined += extract_text_with_tables(p, max_pages=MAX_PDF_PAGES) + '\n'
                except Exception:
                    pass
            group_texts[group] = combined
            # Check if combined text is empty (possible all-image PDF)
            if not combined.strip():
                extraction_warnings.append(
                    f'"{group}"未提取到任何文字内容，可能为全图片扫描件/图形文件，'
                    f'建议上传可复制文字版PDF或Word文件'
                )

        results = run_full_analysis(saved, saved_refs if saved_refs else None,
                                    group_map=group_map, group_texts=group_texts)
        results['_filenames'] = [os.path.basename(s) for s in saved]
        results['_ref_filenames'] = [os.path.basename(s) for s in saved_refs]
        results['_groups'] = {g: [os.path.basename(p) for p in paths] for g, paths in group_map.items()}
        results['_group_order'] = list(group_map.keys())
        if extraction_warnings:
            results['_warnings'] = extraction_warnings
        if size_warnings:
            results.setdefault('_warnings', []).extend(size_warnings)

        # Save to history — strip heavy data, keep only counts & indices
        try:
            history_id = _write_history_entry(
                _prepare_history_data(results), saved, saved_refs)
            results['_history_id'] = history_id
        except Exception as e:
            logger.warning('history save failed in single_upload_and_analyze: %s', e)
            results['_history_save_error'] = True

        return jsonify(results)
    except Exception as e:
        logger.exception('single_upload_and_analyze failed')
        return jsonify({'error': str(e)}), 500


def _verdict_level(entry):
    """Conclusion level of a history entry, preferring the authoritative
    verdict.conclusion_level and falling back to parsing the conclusion text
    (legacy entries) the same way the frontend history list does."""
    data = entry.get('data') or {}
    level = (data.get('verdict') or {}).get('conclusion_level')
    if level in ('high', 'medium', 'low', 'uncertain'):
        return level
    text = entry.get('verdict') or ''
    if '高度嫌疑' in text:
        return 'high'
    if '可疑' in text or '核查' in text:
        return 'medium'
    if '数据不足' in text or '无法' in text:
        return 'uncertain'
    return 'low'


@app.route('/api/stats', methods=['GET'])
def get_stats():
    """Aggregate all history entries for the statistics page.

    One pass over HISTORY_DIR (same fault tolerance as list_history):
    corrupt or legacy entries are skipped instead of failing the whole page.
    """
    DIMENSIONS = ('metadata', 'personnel', 'similarity', 'pricing')
    stats = {
        'total_analyses': 0,
        'total_documents': 0,
        'total_pairs': 0,
        'total_abnormal_matches': 0,
        'total_template_matches': 0,
        'verdict_counts': {'high': 0, 'medium': 0, 'low': 0, 'uncertain': 0},
        'avg_score': None,
        'timeline': [],            # per-day {date, count, avg_score, high}
        'scores_over_time': [],    # per-analysis {time, score, level, id}
        'dimension_hits': {d: 0 for d in DIMENSIONS},    # analyses touching d
        'dimension_totals': {d: 0 for d in DIMENSIONS},  # accumulated findings
        'recent': [],
    }
    loaded = []
    try:
        fnames = sorted(f for f in os.listdir(HISTORY_DIR) if f.endswith('.json'))
    except OSError:
        fnames = []
    for fname in fnames:
        try:
            with open(os.path.join(HISTORY_DIR, fname), 'r', encoding='utf-8') as f:
                entry = json.load(f)
        except Exception:
            continue
        if not entry.get('id'):
            continue
        loaded.append(entry)

    scores = []
    by_day = {}
    for entry in loaded:  # filenames sort by timestamp -> chronological order
        level = _verdict_level(entry)
        data = entry.get('data') or {}
        verdict = data.get('verdict') or {}
        score = verdict.get('score')
        score = float(score) if isinstance(score, (int, float)) else None

        stats['total_analyses'] += 1
        stats['total_documents'] += entry.get('bid_count') or 0
        stats['total_pairs'] += entry.get('total_pairs') or 0
        stats['total_abnormal_matches'] += entry.get('abnormal_matches') or 0
        stats['total_template_matches'] += entry.get('template_matches') or 0
        stats['verdict_counts'][level] += 1
        if score is not None:
            scores.append(score)

        ts = data.get('text_similarity') or {}
        sub = ts.get('substantial_abnormal')
        sub_count = len(sub) if isinstance(sub, list) else (sub or 0)
        hits = {
            'metadata': sum(1 for m in (data.get('metadata') or {}).get('matches', [])
                            if m.get('severity') != 'info'),
            'personnel': len((data.get('personnel') or {}).get('cross_matches', [])),
            'similarity': sub_count,
            'pricing': len((data.get('pricing') or {}).get('findings', [])),
        }
        for dim, cnt in hits.items():
            if cnt > 0:
                stats['dimension_hits'][dim] += 1
                stats['dimension_totals'][dim] += cnt

        time_str = entry.get('time') or ''
        stats['scores_over_time'].append({
            'time': time_str, 'score': score, 'level': level, 'id': entry['id'],
        })
        day = time_str.split(' ')[0] if time_str else '未知'
        slot = by_day.setdefault(day, {'count': 0, 'scores': [], 'high': 0})
        slot['count'] += 1
        if score is not None:
            slot['scores'].append(score)
        if level == 'high':
            slot['high'] += 1

    if scores:
        stats['avg_score'] = round(sum(scores) / len(scores), 1)
    stats['timeline'] = [
        {
            'date': day,
            'count': slot['count'],
            'avg_score': round(sum(slot['scores']) / len(slot['scores']), 1)
            if slot['scores'] else None,
            'high': slot['high'],
        }
        for day, slot in by_day.items()
    ]
    stats['recent'] = [
        {
            'id': e.get('id'),
            'time': e.get('time'),
            'bid_count': e.get('bid_count', 0),
            'verdict': e.get('verdict', ''),
            'level': _verdict_level(e),
            'score': ((e.get('data') or {}).get('verdict') or {}).get('score'),
        }
        for e in reversed(loaded[-8:])  # newest first
    ]
    return jsonify(stats)


@app.route('/api/history', methods=['GET'])
def list_history():
    """List all saved analysis history entries."""
    entries = []
    for fname in sorted(os.listdir(HISTORY_DIR), reverse=True):
        if not fname.endswith('.json'):
            continue
        try:
            with open(os.path.join(HISTORY_DIR, fname), 'r', encoding='utf-8') as f:
                entry = json.load(f)
            eid = entry.get('id')
            # Skip entries with missing/null IDs or where the stored id
            # doesn't match the filename (e.g. leftover test files). This
            # keeps the list clean and every entry deletable.
            if not eid or not isinstance(eid, str):
                continue
            if fname != f'{eid}.json':
                continue
            entries.append({
                'id': eid,
                'time': entry.get('time'),
                'bid_count': entry.get('bid_count', 0),
                'ref_count': entry.get('ref_count', 0),
                'bid_files': entry.get('bid_files', []),
                'ref_files': entry.get('ref_files', []),
                'verdict': entry.get('verdict', ''),
                'abnormal_matches': entry.get('abnormal_matches', 0),
                'template_matches': entry.get('template_matches', 0),
                'total_pairs': entry.get('total_pairs', 0),
            })
        except Exception:
            pass
    return jsonify(entries)

@app.route('/api/history/<history_id>', methods=['GET'])
def get_history(history_id):
    """Retrieve a specific analysis from history."""
    # Reject ids that don't match the alphanumeric+underscore format we
    # produce — prevents path traversal with e.g. '../../etc/passwd'.
    if not _VALID_HISTORY_RE.match(history_id):
        return jsonify({'error': '无效的记录ID'}), 400
    fpath = os.path.join(HISTORY_DIR, f'{history_id}.json')
    if not os.path.exists(fpath):
        return jsonify({'error': '记录不存在'}), 404
    try:
        with open(fpath, 'r', encoding='utf-8') as f:
            entry = json.load(f)
        data = entry.get('data', {})
        data['_history_id'] = history_id
        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/history/<history_id>', methods=['DELETE'])
def delete_history(history_id):
    """Delete a history entry."""
    if not _VALID_HISTORY_RE.match(history_id):
        return jsonify({'error': '无效的记录ID'}), 400
    fpath = os.path.join(HISTORY_DIR, f'{history_id}.json')
    if os.path.exists(fpath):
        os.remove(fpath)
        return jsonify({'ok': True})
    # Fallback: some legacy files may have a filename that diverged from
    # their stored id.  Walk the history dir and delete any file whose
    # internal 'id' field matches the requested history_id.
    try:
        for fname in os.listdir(HISTORY_DIR):
            if not fname.endswith('.json'):
                continue
            candidate = os.path.join(HISTORY_DIR, fname)
            try:
                with open(candidate, 'r', encoding='utf-8') as f:
                    entry = json.load(f)
                if entry.get('id') == history_id:
                    os.remove(candidate)
                    return jsonify({'ok': True})
            except Exception:
                pass
    except OSError:
        pass
    return jsonify({'error': '记录不存在'}), 404

@app.route('/api/ping')
def api_ping():
    """Liveness marker for the desktop single-instance probe (_probe_own_instance)."""
    return jsonify({'app': 'xingyicha', 'ok': True})


@app.after_request
def _nocache(response):
    """Disable caching for static assets during development only.

    In production (gunicorn, debug=False) static assets are served with
    versioned query strings (?v=4) so aggressive caching is safe and desirable.
    """
    if request.path.startswith('/static/') and app.debug:
        response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    return response


def _open_page(url):
    print(f'  正在打开页面: {url}')
    import webbrowser
    webbrowser.open(url)


def _probe_own_instance(preferred_port, span=10):
    """Desktop single-instance probe: if another 星易查 server already answers
    on 127.0.0.1 ports [preferred_port, preferred_port+span), return its URL.

    A second launch of the frozen app becomes a page-opener for the running
    instance instead of a duplicate server on a drifted port — both would
    share the history dir and interleave writes. Loopback refusals are
    instant, so the common (free-port) case costs one failed connect."""
    import urllib.request
    for cand in range(preferred_port, preferred_port + span):
        try:
            with urllib.request.urlopen(
                    f'http://127.0.0.1:{cand}/api/ping', timeout=0.5) as resp:
                body = resp.read().decode('utf-8', 'replace')
            if resp.status == 200 and 'xingyicha' in body:
                return f'http://127.0.0.1:{cand}'
        except Exception:
            continue
    return None


def _run_macos_gui(url, host, port):
    """Frozen macOS entry: own the main thread with a minimal Cocoa app.

    A bare console server in a .app bundle cannot answer Dock activation
    (reopen) or quit Apple Events — the first launch pops the page, then
    clicking the Dock icon does nothing and further attempts report 'the
    application is not open anymore'. Running NSApplication on the main
    thread gives the process a real event loop:
      - Dock icon click / Finder relaunch -> reopen -> (re)open the page
      - menu-bar status item: 打开页面 / 退出
    waitress moves to a daemon worker thread (it is a thread pool anyway)
    and dies with the process when the user quits. Returns True once the
    user quit the GUI; False when AppKit is unavailable so the caller can
    fall back to the plain console server."""
    try:
        from AppKit import (NSApplication, NSMenu, NSMenuItem, NSStatusBar,
                            NSVariableStatusItemLength)
        from Foundation import NSObject
    except ImportError:
        return False

    class _Delegate(NSObject):
        def applicationShouldHandleReopen_hasVisibleWindows_(self, sender, has_windows):
            _open_page(self.page_url)
            return True

        def openPage_(self, sender):
            _open_page(self.page_url)

        def quit_(self, sender):
            NSApplication.sharedApplication().terminate_(sender)

    delegate = _Delegate.alloc().init()
    delegate.page_url = url

    nsapp = NSApplication.sharedApplication()
    nsapp.setDelegate_(delegate)

    status_item = NSStatusBar.systemStatusBar().statusItemWithLength_(
        NSVariableStatusItemLength)
    status_item.button().setTitle_('星')
    menu = NSMenu.alloc().init()
    item_open = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
        '打开页面', 'openPage:', '')
    item_open.setTarget_(delegate)
    menu.addItem_(item_open)
    menu.addItem_(NSMenuItem.separatorItem())
    item_quit = NSMenuItem.alloc().initWithTitle_action_keyEquivalent_(
        '退出星易查', 'quit:', '')
    item_quit.setTarget_(delegate)
    menu.addItem_(item_quit)
    status_item.setMenu_(menu)

    from waitress import serve

    def _serve():
        try:
            serve(app, host=host, port=port, threads=8)
        except Exception as exc:
            print(f'  服务器线程异常退出: {exc}')

    threading.Thread(target=_serve, daemon=True).start()
    nsapp.run()
    return True


if __name__ == '__main__':
    if '--check' in sys.argv:
        # 离线自检：验证关键依赖可正常导入（用于便携包目标机校验）
        import flask  # noqa: F401
        import docx  # noqa: F401
        import pypdf  # noqa: F401
        import lxml  # noqa: F401
        import olefile  # noqa: F401
        import werkzeug  # noqa: F401
        import jinja2  # noqa: F401
        from docx import Document
        from pypdf import PdfReader
        print('OK: 所有关键模块可正常导入')
        sys.exit(0)
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get('PORT', 5001))
    debug = os.environ.get('DEBUG', '0') == '1'
    # Desktop (frozen) build defaults to loopback: binding 0.0.0.0 would pop the
    # Windows firewall prompt on first launch. Set HOST=0.0.0.0 to share on LAN.
    host = os.environ.get('HOST') or ('127.0.0.1' if IS_FROZEN else '0.0.0.0')

    # Desktop single-instance guard: a running instance owns the preferred
    # port range, so just reopen its page instead of starting a duplicate.
    if IS_FROZEN and not debug:
        existing_url = _probe_own_instance(port)
        if existing_url:
            print(f'  星易查已在运行: {existing_url}')
            print('  正在打开页面...')
            _open_page(existing_url)
            sys.exit(0)

    def _pick_free_port(preferred):
        """Return `preferred` if bindable, else preferred+1..+9 (double-click
        relaunch while a stale instance holds the port should still work).
        No SO_REUSEADDR: on Windows it allows hijacking a port another
        process is actively listening on, defeating the probe."""
        import socket
        for cand in range(preferred, preferred + 10):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                try:
                    s.bind((host, cand))
                    return cand
                except OSError:
                    continue
        return preferred

    if not debug:
        requested_port = port
        port = _pick_free_port(port)
        if port != requested_port:
            print(f'  端口 {requested_port} 已被占用，自动改用 {port}')
    url = f'http://{"127.0.0.1" if host in ("127.0.0.1", "localhost") else host}:{port}'
    print('=' * 60)
    print('  星易查 - 围串标风险识别分析系统')
    print(f'  访问地址: {url}')
    print('=' * 60)

    # Desktop build: pop the default browser once the server is up. Delayed so
    # the listener exists first.
    if IS_FROZEN and not debug:
        threading.Timer(1.5, lambda: _open_page(url)).start()

    if debug:
        app.run(debug=debug, host=host, port=port, threaded=True)
    elif IS_FROZEN and sys.platform == 'darwin' and _run_macos_gui(url, host, port):
        pass  # Cocoa event loop owned the main thread; returned on user quit
    else:
        try:
            # waitress: production-grade pure-Python WSGI server, the only
            # option on Windows (gunicorn is Unix-only). Long analyses rely on
            # threaded request handling, same as the Flask dev server.
            from waitress import serve
            print('  服务器: waitress (threads=8)')
            serve(app, host=host, port=port, threads=8)
        except ImportError:
            app.run(debug=debug, host=host, port=port, threaded=True)
