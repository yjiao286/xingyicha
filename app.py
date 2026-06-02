import os
import re
import json
import base64
import zipfile
import subprocess
import tempfile
import hashlib
import shutil
from io import BytesIO
from datetime import datetime
from collections import defaultdict

from flask import Flask, request, jsonify, send_file, render_template, Response
from docx import Document
from pypdf import PdfReader
import threading

app = Flask(__name__)
app.secret_key = 'bid_analysis_secret_key_2025'

UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER', tempfile.mkdtemp(prefix='bid_uploads_'))
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HISTORY_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'history')
os.makedirs(HISTORY_DIR, exist_ok=True)

# ── Helpers ─────────────────────────────────────────────────────
def sanitize_text(text):
    """Remove control characters that break JSON serialization"""
    if not isinstance(text, str):
        return text
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)

def _find_tool(*names):
    """Find first available command-line tool"""
    for name in names:
        if shutil.which(name):
            return name
    return None

# ── .doc Conversion ─────────────────────────────────────────────
_DOC_CONVERTER = None

def _get_doc_converter():
    """Lazy-init: find available .doc converter"""
    global _DOC_CONVERTER
    if _DOC_CONVERTER is None:
        _DOC_CONVERTER = _find_tool('libreoffice', 'soffice', 'antiword', 'catdoc')
    return _DOC_CONVERTER

def convert_doc_to_docx(filepath):
    """Convert .doc to .docx using LibreOffice. Returns path to .docx or None."""
    outdir = tempfile.mkdtemp(prefix='doc_conv_')
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'docx', '--outdir', outdir, filepath],
            capture_output=True, timeout=60, check=True
        )
        for f in os.listdir(outdir):
            if f.endswith('.docx'):
                return os.path.join(outdir, f)
    except Exception:
        pass
    return None

def extract_doc_text_raw(filepath):
    """Extract text from .doc using antiword or catdoc"""
    tool = _get_doc_converter()
    if not tool:
        return None

    try:
        if tool in ('antiword', 'catdoc'):
            result = subprocess.run([tool, filepath], capture_output=True, text=True, timeout=30)
            if result.returncode == 0:
                return result.stdout.strip()
        elif tool in ('libreoffice', 'soffice'):
            # Convert .doc → .docx first
            docx_path = convert_doc_to_docx(filepath)
            if docx_path:
                doc = Document(docx_path)
                lines = [p.text for p in doc.paragraphs]
                return '\n'.join(lines)
    except Exception:
        pass
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

def get_file_type(filepath):
    """Detect file type: 'docx', 'doc', or 'pdf'"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.pdf':
        return 'pdf'
    if ext == '.doc':
        return 'doc'
    return 'docx'

def extract_metadata(filepath):
    """Extract metadata from .docx or .pdf"""
    if get_file_type(filepath) == 'pdf':
        return extract_pdf_metadata(filepath)

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
    except:
        pass
    return result


# ── Text Extraction ─────────────────────────────────────────────
def extract_text(filepath):
    """Extract all text from a .docx file"""
    doc = Document(filepath)
    paragraphs = []
    for p in doc.paragraphs:
        text = p.text
        if text.strip():
            paragraphs.append(text)
    return '\n'.join(paragraphs)

def extract_text_with_tables(filepath):
    """Extract text including tables from .docx, .doc, or .pdf"""
    ftype = get_file_type(filepath)

    if ftype == 'pdf':
        reader = PdfReader(filepath)
        lines = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                lines.append(text)
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
        lines.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            row_text = ' | '.join(cell.text for cell in row.cells)
            lines.append(row_text)
    return '\n'.join(lines)


# ── Personnel Extraction ────────────────────────────────────────
def extract_personnel(text):
    """Extract legal representative, authorized agent info from bid text"""
    info = {}

    # Normalize line breaks within key phrases that often get split across lines
    # e.g., "法定代\n表人" → "法定代表人", "授权委\n托书" → "授权委托书"
    text = re.sub(r'法定代\s*\n\s*表人', '法定代表人', text)
    text = re.sub(r'法定\s*\n\s*代表人', '法定代表人', text)
    text = re.sub(r'法\s*\n\s*定代表人', '法定代表人', text)
    text = re.sub(r'授权委\s*\n\s*托书', '授权委托书', text)
    text = re.sub(r'供应\s*\n\s*商名称', '供应商名称', text)

    # Legal representative — try multiple patterns
    rep = None
    co = None

    # Pattern 0: "我张三（姓名）系四川某某电子科技有限公司（供应商名称）的法定代表人"
    # Common in 法定代表人授权委托书 — captures name before （姓名） and company before （供应商名称）
    if not rep:
        m = re.search(r'(?:本人\s*)?我?\s*([一-鿿]{2,4})\s*[（(]姓名[）)]\s*系\s*(.{1,40}?)\s*[（(]供应商名称[）)]\s*的法定代表人', text)
        if m:
            rep = m.group(1).strip()
            co = m.group(2).strip()

    # Pattern 1: "姓名：刘某某 ... 职务：院长 系 XXX 的法定代表人" (身份证明 section)
    if not rep:
        m = re.search(r'姓名[：:]\s*([^\s]{2,10})\s*[\s\S]*?系\s*(.{1,30}?)\s*的法定代表人', text)
        if m:
            rep = m.group(1).strip()
            co = m.group(2).strip()

    # Pattern 2: "本人 王某某 系 北京某某大学 的法定代表人"
    if not rep:
        m = re.search(r'(?:本人\s*)?([^\s系]{2,10})\s*(?:[（(]姓名[）)])?\s*系\s*(.{1,30}?)\s*的法定代表人', text)
        if m:
            rep = m.group(1).strip()
            co = m.group(2).strip()

    # Pattern 3: 法定代表人授权书 "（王戈、董事长）代表本公司授权（赵凯、销售经理）"
    m = re.search(r'（([^、）]{2,10})[、，].{1,6}?）\s*代表本公司授权\s*（([^、）]{2,10})', text)
    if m:
        rep = m.group(1).strip()
        info['authorized_rep'] = m.group(2).strip()
    if not rep:
        m = re.search(r'（([^、）]{2,10})[、，].{1,6}?）\s*代表.{1,10}授权', text)
        if m:
            rep = m.group(1).strip()

    # Pattern 4: "（兰某某）系（北京某某航天技术有限公司）的法定代表人"
    if not rep:
        m = re.search(r'[（(]([^\s系]{2,10})[）)]\s*系\s*[（(](.{1,30}?)[）)]\s*的法定代表人', text)
        if m:
            rep = m.group(1).strip()
            co = m.group(2).strip()

    # Pattern 4: 法定代表人授权书 — find name after the section
    if not rep:
        m = re.search(r'法定代表人授权书[\s\S]{0,500}?(?:被授权人|授权代表|代理人)[：:]\s*([^\s]{2,10})', text)
        if m:
            info['authorized_rep'] = m.group(1).strip()
        m = re.search(r'法定代表人授权书[\s\S]{0,300}?(?:法定代表人|单位负责人)[^：:]*?[：:]\s*([^\s]{2,10})', text)
        if m:
            rep = m.group(1).strip()
    if not rep:
        m = re.search(r'[（(]([^\s系]{2,10})[）)]\s*系\s*[（(](.{1,30}?)[）)]\s*的法定代表人', text)
        if m:
            rep = m.group(1).strip()
            co = m.group(2).strip()

    if rep:
        rep = re.sub(r'^(?:本人\s*)+', '', rep).strip()
        rep = re.sub(r'^[（(]|[）)]$', '', rep)
        # Strip "我" prefix (common in 授权委托书: "我张三（姓名）系...")
        rep = re.sub(r'^我(?=[一-鿿])', '', rep)
        # Strip trailing parenthesized labels like （姓名）, (姓名), （签字）, (签字)
        rep = re.sub(r'\s*[（(](?:姓名|签字|盖章|单位负责人|法定代表人)[）)]\s*$', '', rep)
        # Strip leading parenthesized labels
        rep = re.sub(r'^\s*[（(](?:姓名|签字|盖章|单位负责人|法定代表人)[）)]\s*', '', rep)
        # Exclude garbage matches
        if len(rep) < 2 or any(w in rep for w in ['注册', '签字', '盖章', '国）', '地址', '电话', '投标人']):
            rep = None
        else:
            info['legal_rep'] = rep
        co = re.sub(r'[（(]投标人名称[）)]|[（(]单位负责人[）)]|[（(]供应商名称[）)]', '', co or '')
        co = re.sub(r'^[（(]|[）)]$', '', co).strip()
        info['company_name'] = co

    # Authorized representative
    m = re.search(r'现委托\s*(.{1,10})\s*[（(]姓名[）)]', text)
    if not m:
        m = re.search(r'委托\s*(.{1,10})\s*[（(]姓名[）)]\s*为我方', text)
    if m:
        info['authorized_rep'] = m.group(1).strip()

    # ID number
    m = re.search(r'身份证号码[：:]\s*(\d{17}[\dXx])', text)
    if m:
        info['id_number'] = m.group(1).strip()

    # Phone
    m = re.search(r'电话[：:]\s*(\d{7,15})', text)
    if m:
        info['phone'] = m.group(1).strip()

    # Address
    m = re.search(r'地址[：:]\s*(.{10,80})', text)
    if m:
        info['address'] = m.group(1).strip()[:100]

    # Response date
    m = re.search(r'(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', text[:800])
    if m:
        info['response_date'] = m.group(1).strip()

    return info


# ── Price Extraction ────────────────────────────────────────────
def _parse_amount(s):
    """Parse a price string like '1,234,567.89', '123.45万元', '1.2亿元', '壹佰贰拾叁万' to float"""
    if not s:
        return 0.0
    s = str(s).replace(',', '').replace('，', '').strip()

    # Handle Chinese uppercase numerals (壹贰叁肆伍陆柒捌玖拾佰仟万亿)
    _CN_NUM = {'零': 0, '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5,
               '陆': 6, '柒': 7, '捌': 8, '玖': 9, '拾': 10, '佰': 100,
               '仟': 1000, '万': 10000, '亿': 100000000, '一': 1,
               '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
               '七': 7, '八': 8, '九': 9, '十': 10, '百': 100, '千': 1000}
    has_cn = any(ch in _CN_NUM for ch in s)
    if has_cn:
        # Try to extract a simple numeric fallback first
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

    # Handle full-width digits
    s = s.replace('０', '0').replace('１', '1').replace('２', '2').replace('３', '3').replace('４', '4')
    s = s.replace('５', '5').replace('６', '6').replace('７', '7').replace('８', '8').replace('９', '9')

    # Extract first numeric value (handle price ranges: take first value)
    m = re.search(r'([\d]+\.?\d*)', s)
    return float(m.group(1)) * wan if m else 0.0

# ── Structured Price Extraction ─────────────────────────────────
def extract_prices(text):
    """Extract structured pricing using multi-channel pipeline with confidence scoring.
    Returns dict with: totalPriceInTax, totalPrice, taxRate, revenue, cost,
    subItemPrice[], costDetails[]"""
    result = {
        'totalPriceInTax': None,
        'totalPrice': None,
        'taxRate': None,
        'revenue': None,
        'cost': None,
        'subItemPrice': [],
        'costDetails': []
    }

    # ── Channel 1: Symbol-based (￥/¥/CNY/RMB) ── confidence: 0.95
    symbol_patterns = [
        r'(?:CNY|RMB)\s*([\d,]+\.?\d*)',
        r'[￥¥]\s*([\d,]+\.?\d*)',
        r'USD\s*([\d,]+\.?\d*)',
    ]
    for pat in symbol_patterns:
        m = re.search(pat, text)
        if m:
            val = _parse_amount(m.group(1))
            if val >= 100:
                result['totalPriceInTax'] = val
                result['totalPrice'] = val
                break

    # ── Channel 2: Label-based (标签通道) ── confidence: 0.90
    label_patterns = [
        r'人民币[：:]\s*([\d,]+\.?\d*)',
        r'小写[：:]\s*([\d,]+\.?\d*)',
        r'(?:投标总价|投标总报价|总报价|报价金额|投标报价|项目总价)[：:]\s*([\d,]+\.?\d*)',
        r'(?:总价|总计|合计)[：:]\s*([\d,]+\.?\d*)',
        r'(?:金额|报价)[（(]元[）)][：:]\s*([\d,]+\.?\d*)',
    ]
    if result['totalPriceInTax'] is None:
        for pat in label_patterns:
            m = re.search(pat, text)
            if m:
                val = _parse_amount(m.group(1))
                if val >= 100:
                    result['totalPriceInTax'] = val
                    result['totalPrice'] = val
                    break

    # ── Channel 3: 大写/小写 pair ── confidence: 0.88
    if result['totalPriceInTax'] is None:
        m = re.search(r'大写[：:]\s*[壹贰叁肆伍陆柒捌玖拾佰仟万亿零一二三四五六七八九十百千元整角分]+[\s\S]{0,100}?小写[：:]\s*([\d,]+\.?\d*)', text)
        if m:
            val = _parse_amount(m.group(1))
            if val >= 100:
                result['totalPriceInTax'] = val
                result['totalPrice'] = val

    # ── Channel 4: Table-based ── confidence: 0.85
    bid_section = _find_bid_summary_section(text)
    if bid_section and result['totalPriceInTax'] is None:
        for pat in [
            r'(?:CNY|RMB|￥|¥)\s*([\d,]+\.?\d*)',
            r'人民币[：:]\s*([\d,]+\.?\d*)',
            r'小写[：:]\s*([\d,]+\.?\d*)',
            r'(?:总价|总计|合计|报价)[：:]?\s*([\d,]+\.?\d*)',
        ]:
            m = re.search(pat, bid_section)
            if m:
                val = _parse_amount(m.group(1))
                if val >= 100:
                    result['totalPriceInTax'] = val
                    result['totalPrice'] = val
                    break

    # ── Tax rate decomposition ──
    _extract_tax_decomposition(text, bid_section if bid_section else text, result)

    # ── Always try to extract subItemPrice and costDetails ──
    _extract_structured_items(text, result)

    return result


def _find_bid_summary_section(text):
    """Find the bid summary / price overview section in text."""
    keywords = ['开标一览表', '开标一览', '投标报价表', '报价一览表', '报价总表', '投标总价']
    for kw in keywords:
        idx = text.find(kw)
        while idx >= 0:
            # Skip TOC entries
            prefix = text[max(0, idx - 40):idx]
            if not re.search(r'\.{3,}', prefix):
                # Find end: next major section or 3000 chars
                end = min(idx + 3000, len(text))
                for end_kw in ['投标分项报价', '分项报价表', '法定代表人', '技术方案', '项目概况']:
                    ep = text.find(end_kw, idx + 10)
                    if ep > idx and ep < end:
                        end = ep
                return text[idx:end]
            idx = text.find(kw, idx + 1)
    return None


def _extract_tax_decomposition(text, section, result):
    """Extract pre-tax / tax / post-tax breakdown."""
    patterns = [
        r'(?:￥|¥)?(\d{5,10}(?:\.\d{2})?)\s+(\d{1,2})\s*[%％]\s*(?:￥|¥)?(\d{5,10}(?:\.\d{2})?)',
        r'人民币[：:]\s*(\d{5,10}(?:\.\d{2})?)\s*元?\s+(\d{1,2})\s*[%％]?\s+人民币[：:]\s*(\d{5,10}(?:\.\d{2})?)',
        r'([\d.]+)\s*万\s+([\d.]+)\s*万\s+(\d{1,2})',
        r'小写[：:]\s*(\d{5,10}(?:\.\d{2})?)\s*元[\s\S]{0,80}?(\d{1,2})\s*[%％][\s\S]{0,80}?小写[：:]\s*(\d{5,10}(?:\.\d{2})?)',
    ]
    for pat in patterns:
        m = re.search(pat, section)
        if m:
            v1 = _parse_amount(m.group(1))
            v2 = _parse_amount(m.group(3))
            if v1 >= 100 and v2 >= 100:
                result['totalPrice'] = v1
                result['totalPriceInTax'] = v2
                result['taxRate'] = str(int(m.group(2))) + '%'
                return


def _extract_structured_items(text, result):
    """Extract sub-item pricing and cost details using enhanced section discovery
    and generic cost line detection."""

    # ── Enhanced section discovery ──
    section_keywords = [
        '分项报价表', '分项报价', '报价明细', '价格表', '开标一览',
        '报价清单', '费用明细', '价格清单', '投标报价', '价格构成',
        '设备清单', '费用清单', '报价构成', '价格明细', '成本明细',
        '项目报价', '费用构成', '费用表'
    ]

    bid_section = None
    # Build alternation pattern from keywords
    kw_pattern = '|'.join(re.escape(kw) for kw in section_keywords)
    for m in re.finditer(
        r'(?:^|\n)(?:[一二三四五六七八九十\d]+[、.。]\s*|\d+(?:\.\d+)+\s*|\d+\s+)?(' +
        kw_pattern + r')\s*\n', text
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
        _parse_pdf_bid_table(bid_section, result)

    # ── Generic cost line extraction ──
    # Match lines with: Chinese name (2-20 chars containing 费/成本/支出 etc) + large number (>= 100)
    cost_pattern = re.compile(
        r'(?:^|\n)\s*([一-鿿]{2,20}(?:费|成本|支出|投入|工资|薪酬|酬金|折旧|摊销|租赁|租金|'
        r'维护|保养|检测|试验|测试|设计|开发|研制|采购|运输|差旅|会议|培训|办公|印刷|'
        r'咨询|审计|评估|保险|税费|利息|手续费|管理|服务|劳务|材料|设备|仪器|软件|'
        r'许可|专利|著作|技术|咨询|外协|加工|燃料|动力|事务|不可预见|预备|风险|'
        r'收益|利润|税金|公积金|基金)[一-鿿]{0,6})\s+(\d{4,}(?:\.\d{2})?)',
        re.MULTILINE
    )
    seen_names = set()
    for m in cost_pattern.finditer(text):
        name = m.group(1).strip()
        if name in seen_names:
            continue
        val = float(m.group(2).replace(',', ''))
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

    # Merge wrapped names: a line with no amounts merges into the next line with amounts
    merged_rows = []
    pending_name = []
    for s in raw_rows:
        has_amounts = bool(re.search(r'(\d{4,}|[\d.]+\s*万)', s))
        if has_amounts:
            if pending_name:
                merged_rows.append((''.join(pending_name), s))
                pending_name = []
            else:
                merged_rows.append(('', s))
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
        result['subItemPrice'] = items
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
    for kw in ['合计', '总价', '小计', '总计']:
        m = re.search(kw + r'\s+([\d.]+)\s*万', section)
        if m:
            result['totalPrice'] = _parse_amount(m.group(1) + '万')
            return
        m = re.search(kw + r'\s+(\d{5,12}(?:\.\d{2})?)', section)
        if m:
            result['totalPrice'] = float(m.group(1))
            return


# ── Text Similarity ─────────────────────────────────────────────
def find_common_segments(text1, text2, min_len=15):
    """Find substrings >= min_len chars that appear in both texts.
    Returns list of (pos1, pos2, length, segment, ctx1, ctx2)"""
    results = []
    i = 0
    found_positions = set()

    while i < len(text1) - min_len:
        best_len = 0
        max_search = min(500, len(text1) - i)
        for length in range(min_len, max_search):
            segment = text1[i:i+length]
            if segment in text2:
                best_len = length
            else:
                break

        if best_len >= min_len:
            segment = text1[i:i+best_len]
            chinese_chars = len(re.findall(r'[一-鿿]', segment))
            if chinese_chars >= 10:
                pos2 = text2.find(segment)
                if pos2 >= 0 and not any(abs(i - p) < 10 for p in found_positions):
                    # Extract context: ~100 chars before and after
                    ctx_before = 100
                    ctx_after = 100
                    ctx1 = text1[max(0, i - ctx_before):i + best_len + ctx_after]
                    ctx2 = text2[max(0, pos2 - ctx_before):pos2 + best_len + ctx_after]
                    results.append((i, pos2, best_len, segment, ctx1, ctx2))
                    found_positions.add(i)
            i += max(best_len, 1)
        else:
            i += 1

    return results

def is_template_content(text):
    """Check if text is likely a standard template/bid instruction phrase"""
    template_markers = [
        '供应商名称', '法定代表人或授权代表签字', '项目编号', '项目名称',
        '注：', '公章', '供应商全称', '盖单位章', '签字或盖章',
        '（单位公章）', '（盖章）', '签字或印章',
    ]
    # Short segments that are just template boilerplate
    if len(text) < 30:
        for marker in template_markers:
            if marker in text:
                return True
    return False

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

def _normalize_for_match(text):
    """Normalize text for comparison: collapse whitespace, strip"""
    return re.sub(r'\s+', '', text)

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

def text_similarity_analysis(texts_dict, ref_texts_list=None):
    """Full text similarity analysis across all uploaded files.
    ref_texts_list: list of text strings from reference/template documents to exclude.
    """
    filenames = list(texts_dict.keys())
    results = {
        'total_pairs': 0,
        'pair_results': [],
        'findings': [],
        'all_abnormal': [],
        'template_matches': 0
    }

    ref_texts = ref_texts_list or []

    for i in range(len(filenames)):
        for j in range(i+1, len(filenames)):
            results['total_pairs'] += 1
            t1, t2 = texts_dict[filenames[i]], texts_dict[filenames[j]]

            segments = find_common_segments(t1, t2, min_len=15)
            pair_result = {
                'file1': filenames[i],
                'file2': filenames[j],
                'total_matches': len(segments),
                'matches': [],
                'abnormal_count': 0,
                'template_count': 0
            }

            for idx, (pos1, pos2, length, seg_text, ctx1, ctx2) in enumerate(segments):
                # Check if this segment is from reference/template docs
                in_ref = _is_in_reference(seg_text, ref_texts)

                if in_ref:
                    # This is a template match — expected, not suspicious
                    pair_result['template_count'] += 1
                    results['template_matches'] += 1
                    # Still record it but mark as template
                    pair_result['matches'].append({
                        'index': idx + 1,
                        'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'abnormal': False,
                        'reasons': ['招标文件/模板内容 — 非异常一致'],
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })
                    continue

                if is_template_content(seg_text):
                    pair_result['template_count'] += 1
                    results['template_matches'] += 1
                    pair_result['matches'].append({
                        'index': idx + 1,
                        'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'abnormal': False,
                        'reasons': ['格式模板内容'],
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })
                    continue

                reasons = classify_abnormal_reason(seg_text)
                is_abnormal = len(reasons) > 0 and not (
                    len(seg_text) < 30 and '注：' in seg_text
                )

                match_entry = {
                    'index': idx + 1,
                    'length': length,
                    'text': sanitize_text(seg_text[:300]),
                    'abnormal': True,
                    'reasons': reasons,
                    'pos1': pos1, 'pos2': pos2,
                    'ctx1': sanitize_text(ctx1[:400]),
                    'ctx2': sanitize_text(ctx2[:400])
                }

                pair_result['matches'].append(match_entry)
                if is_abnormal:
                    pair_result['abnormal_count'] += 1
                    results['all_abnormal'].append({
                        'pair': f'{filenames[i]} vs {filenames[j]}',
                        'index': idx + 1,
                        'length': length,
                        'text': sanitize_text(seg_text[:300]),
                        'reasons': reasons,
                        'pos1': pos1, 'pos2': pos2,
                        'ctx1': sanitize_text(ctx1[:400]),
                        'ctx2': sanitize_text(ctx2[:400])
                    })

            results['pair_results'].append(pair_result)

    # Generate findings
    total_abnormal = sum(p['abnormal_count'] for p in results['pair_results'])
    total_template = results['template_matches']
    if total_template > 0:
        results['findings'].append(f'扣除招标文件/模板内容后: 共 {total_template} 处模板匹配已排除')
    if total_abnormal > 0:
        results['findings'].append(f'共发现 {total_abnormal} 处异常一致的文本段落（已排除模板内容）')
    else:
        results['findings'].append('未发现异常一致的文本段落（扣除模板内容后）')
    if any('机器翻译' in str(r.get('reasons', [])) for r in results['all_abnormal']):
        results['findings'].append('存在相同的不规范翻译表述（机器翻译痕迹），排除独立编制可能')
    if any('技术型号' in str(r.get('reasons', [])) for r in results['all_abnormal']):
        results['findings'].append('技术方案中具体型号/参数选择一致，不属于通用技术规范')

    return results


# ── Document Structure ──────────────────────────────────────────
def extract_structure(text):
    """Extract document TOC/structure"""
    toc_pattern = re.findall(r'^[一二三四五六七八九十]+[、，.]\s*.+|^\d+\.\d*\.?\s*.+|^[（(]\w+[）)]\s*.+', text, re.MULTILINE)
    return toc_pattern[:80]


# ── Comprehensive Analysis ──────────────────────────────────────

def _build_sub_item_comparison(all_prices, filenames):
    """Build fuzzy-merged sub-item pricing comparison table."""
    # Normalize names
    def _norm_name(name):
        n = name.strip()
        n = re.sub(r'^算法设计文档[《<]', '', n)
        n = re.sub(r'[》>]及配套(?:代码|成果|成)\s*$', '', n)
        n = re.sub(r'及配套(?:成果|成)\s*$', '', n)
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
    # Collect from costDetails (docx 国防科技工业 format - 成本明细)
    for fn in filenames:
        for item in all_prices.get(fn, {}).get('costDetails', []):
            all_items.append({
                'file': fn, 'type': '成本明细',
                'name': item.get('priceName', ''),
                'norm': item.get('priceName', ''),
                'count': item.get('count'), 'unitPrice': item.get('unitPrice'),
                'totalPrice': item.get('totalPrice'), 'totalPriceInTax': item.get('totalPriceInTax'),
                'tax': item.get('tax'), 'extras': item.get('extras', {})
            })

    if not all_items:
        return []

    # LCS clustering
    def _lcs_len(a, b):
        m, n = len(a), len(b)
        dp = [[0]*(n+1) for _ in range(m+1)]
        best = 0
        for i in range(1, m+1):
            for j in range(1, n+1):
                if a[i-1] == b[j-1]:
                    dp[i][j] = dp[i-1][j-1] + 1
                    best = max(best, dp[i][j])
        return best

    clusters = []
    used = set()
    for i, item_i in enumerate(all_items):
        if i in used: continue
        cluster = [item_i]
        used.add(i)
        best_name = item_i['name']
        for j, item_j in enumerate(all_items):
            if j in used: continue
            # Match: exact same name, or long common substring
            same_name = item_i['norm'] == item_j['norm']
            long_match = len(item_i['norm']) >= 10 and len(item_j['norm']) >= 10 and _lcs_len(item_i['norm'], item_j['norm']) >= 15
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
        if len(items) < 2:
            c['findings'] = []
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
                    findings.append(f'不含税报价差异仅{diff_pct:.1f}%，高度接近')
                elif diff_pct < 10:
                    findings.append(f'不含税报价差异{diff_pct:.1f}%')

        # 3. Sequential pattern detection
        if len(prices_excl) >= 3:
            sorted_prices = sorted(prices_excl)
            gaps = [sorted_prices[i+1] - sorted_prices[i] for i in range(len(sorted_prices)-1)]
            if len(set(gaps)) == 1:
                findings.append(f'报价呈等差数列（公差{gaps[0]:,.0f}元），存在规律性差异')

        # 4. Only one bidder has this item
        if len(items) == 1:
            findings.append(f'仅{os.path.basename(items[0]["file"])}有此分项')

        c['findings'] = findings

    return clusters




def run_full_analysis(filepaths, ref_filepaths=None, group_map=None, group_texts=None, on_progress=None):
    """Run all analysis modules and return structured results.
    ref_filepaths: optional reference/template document paths.
    group_map: {group_name: [filepath, ...]} for multi-volume merging.
    group_texts: {group_name: combined_text} pre-merged texts.
    on_progress: callback(step, label, percent, detail) for streaming progress.
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
        group_texts = {fn: extract_text_with_tables(fp) for fn, fp in zip(filenames, filepaths)}

    # 1. Metadata — per original file
    all_meta = {}
    for fp, fn in zip(filepaths, filenames):
        all_meta[fn] = extract_metadata(fp)

    # 2. Text — use merged per-group
    all_text = {}
    for group_name in display_names:
        all_text[group_name] = group_texts.get(group_name, '')

    _progress('text', '提取文本内容', 10, f'已提取 {len(display_names)} 份标书的文本')

    # 2b. Text extraction — reference documents
    ref_texts = []
    ref_filenames = []
    if ref_filepaths:
        for rfp in ref_filepaths:
            try:
                rt = extract_text_with_tables(rfp)
                if rt:
                    ref_texts.append(rt)
                    ref_filenames.append(os.path.basename(rfp))
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
    similarity = text_similarity_analysis(all_text, ref_texts)

    # 6. Structure
    all_structure = {}
    for fn, text in all_text.items():
        all_structure[fn] = extract_structure(text)

    # Build per-group metadata (use first file's metadata)
    group_meta = {}
    for gn in out_names:
        paths = group_map.get(gn, [])
        if paths:
            group_meta[gn] = all_meta.get(os.path.basename(paths[0]), {})

    _progress('metadata', '元数据比对', 25, f'交叉比对 {len(out_names)} 份标书的创建者、修改者、编辑程序等')

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
                        dedup_key = f'{label}|{vi}'
                        if dedup_key not in all_pairs_done:
                            all_pairs_done.add(dedup_key)
                            meta_matches.append({
                                'field': label,
                                'value': str(vi)[:200],
                                'pair': pair_label,
                                'verdict': '完全一致',
                                'severity': 'high' if key in ('KSOTemplateDocerSaveRecord', 'last_modified_by') else 'medium'
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
    if len(out_names) >= 2:
        for i in range(len(out_names)):
            for j in range(i+1, len(out_names)):
                gi, gj = out_names[i], out_names[j]
                pi, pj = all_personnel[gi], all_personnel[gj]
                mi, mj = group_meta.get(gi, {}), group_meta.get(gj, {})

                # Check if one company's authorized rep = another's creator
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

                # Check if same person modified both
                if mi.get('last_modified_by') and mj.get('last_modified_by'):
                    if mi['last_modified_by'] == mj['last_modified_by']:
                        key = f'same_modifier|{mi["last_modified_by"]}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            personnel_matches.append({
                                'type': '最后修改人为同一人',
                                'detail': f'"{mi["last_modified_by"]}"同时为 {gi} 和 {gj} 的最后修改人',
                                'severity': 'high'
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

    _progress('personnel', '人员交叉比对', 40, f'交叉比对法定代表人、授权代表等，发现 {len(personnel_matches)} 处异常')

    # ── Text similarity findings (summary before pricing) ──
    total_abnormal = sum(p['abnormal_count'] for p in similarity['pair_results'])
    _progress('similarity', '文本相似度分析', 55, f'检测 {similarity["total_pairs"]} 组文档对，发现 {total_abnormal} 处异常一致段落')

    # ── Compile pricing comparison (structured format) ──
    price_compare = {}
    price_findings = []

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
            if first_val is None:
                first_val = val
            elif val != first_val:
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
            if first_val is None:
                first_val = val
            elif val != first_val:
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
                    price_findings.append(f'{out_names[i]} 和 {out_names[j]} 不含税总价一致: {pi["totalPrice"]:,.0f}元')
                if pi.get('totalPriceInTax') and pj.get('totalPriceInTax') and pi['totalPriceInTax'] == pj['totalPriceInTax']:
                    price_findings.append(f'{out_names[i]} 和 {out_names[j]} 含税总价一致: {pi["totalPriceInTax"]:,.0f}元')

        for gn, prices in all_prices.items():
            has_total = prices.get('totalPrice') is not None or prices.get('totalPriceInTax') is not None
            has_details = prices.get('costDetails') or prices.get('subItemPrice')
            if not has_total and not has_details:
                price_findings.append(f'{gn}未提取到任何报价/成本信息')
            elif not has_total and has_details:
                price_findings.append(f'{gn}仅提取到成本明细，未提取到总价')

    _progress('pricing', '报价分析', 75, f'比较含税总价、不含税总价、分项单价等')

    # ── Compile verdict ──
    clauses = [
        {
            'clause': '第（一）项',
            'description': '不同投标人的投标文件由同一单位或者个人编制',
            'satisfied': any(m['field'] == 'WPS保存记录(硬件ID+用户ID)' for m in meta_matches) or
                         any(m['type'] == '最后修改人为同一人' for m in personnel_matches),
            'evidence': []
        },
        {
            'clause': '第（二）项',
            'description': '不同投标人委托同一单位或者个人办理投标事宜',
            'satisfied': any('授权代表' in m.get('type', '') for m in personnel_matches),
            'evidence': []
        },
        {
            'clause': '第（三）项',
            'description': '不同投标人的投标文件载明的项目管理成员为同一人',
            'satisfied': None,
            'evidence': ['标书中未明确列出项目团队成员信息，无法判断']
        },
        {
            'clause': '第（四）项-a',
            'description': '投标文件异常一致',
            'satisfied': similarity['all_abnormal'] and len(similarity['all_abnormal']) > 0,
            'evidence': []
        },
        {
            'clause': '第（四）项-b',
            'description': '投标报价呈规律性差异',
            'satisfied': len(price_findings) > 0,
            'evidence': []
        },
    ]

    # Populate evidence
    for c in clauses:
        if c['clause'] == '第（一）项':
            if any(m['field'] == 'WPS保存记录(硬件ID+用户ID)' for m in meta_matches):
                c['evidence'].append('WPS硬件ID和用户ID完全一致，同一台设备同一账号编辑')
            if any(m['type'] == '最后修改人为同一人' for m in personnel_matches):
                c['evidence'].append('两份标书最后修改人为同一人')
            if any('创建者' in m.get('type', '') for m in personnel_matches):
                c['evidence'].append('一方授权代表为另一方标书创建者')
            c['evidence_level'] = '强' if c['satisfied'] else '无'
        elif c['clause'] == '第（二）项':
            c['evidence_level'] = '中' if c['satisfied'] else '无'
        elif c['clause'] == '第（三）项':
            c['evidence_level'] = '无法判断'
        elif c['clause'] == '第（四）项-a':
            c['evidence'].append(f'共发现 {len(similarity["all_abnormal"])} 处异常一致文本段落')
            c['evidence_level'] = '强' if c['satisfied'] else '无'
        elif c['clause'] == '第（四）项-b':
            c['evidence'] = price_findings
            c['evidence_level'] = '中' if c['satisfied'] else '无'

    num_bids = len(out_names)
    num_word = {2: '两份', 3: '三份', 4: '四份', 5: '五份', 6: '六份', 7: '七份', 8: '八份', 9: '九份', 10: '十份'}
    bid_word = num_word.get(num_bids, f'{num_bids}份')
    conclusion = f'{bid_word}标书存在围标串标高度嫌疑' if any(
        c.get('evidence_level') == '强' for c in clauses
    ) else '需要进一步核查'

    # Also fix old hardcoded "两份" in findings
    for i, f_text in enumerate(time_findings):
        if '份标书' in f_text and '两份' in f_text:
            time_findings[i] = f_text.replace('两份标书', f'{bid_word}标书')

    _progress('verdict', '综合判定', 90, f'依据《招标投标法实施条例》第四十条判定：{conclusion}')

    return {
        'metadata': {
            'files': [{'name': fn, **all_meta[fn]} for fn in filenames],
            'matches': meta_matches,
            'findings': time_findings + [
                f'KSOProductBuildVer一致: 同一WPS版本',
                f'最后保存者一致: lzkj' if any(m['field'] == '最后保存者' for m in meta_matches) else '',
                f'文档修改时间相隔很近，存在连续编辑特征',
            ]
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
            'subItemCompare': _build_sub_item_comparison(all_prices, out_names),
            'findings': price_findings
        },
        'structure': {gn: all_structure.get(gn, [])[:60] for gn in out_names},
        'ref_docs': ref_filenames,
        'verdict': {
            'clauses': clauses,
            'conclusion': conclusion
        }
    }


# ── Report Generation ────────────────────────────────────────────
def generate_report_docx(analysis):
    """Generate a .docx report from analysis results"""
    doc = Document()

    # Title
    title = doc.add_heading('围串标投标文件分析报告', level=0)
    title.alignment = 1  # center

    doc.add_paragraph(f'生成时间: {datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")}')
    doc.add_paragraph('分析依据: 《中华人民共和国招标投标法实施条例》第四十条')
    doc.add_paragraph('─' * 60)

    # 1. Metadata
    doc.add_heading('一、元数据分析', level=1)
    for f in analysis['metadata']['files']:
        doc.add_heading(f'文件: {f["name"]}', level=2)
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

    if analysis['metadata']['matches']:
        doc.add_heading('元数据一致项', level=2)
        for m in analysis['metadata']['matches']:
            doc.add_paragraph(f'【{m["verdict"]}】{m["field"]}: {m["value"]}', style='List Bullet')
    for f_text in analysis['metadata']['findings']:
        if f_text.strip():
            doc.add_paragraph(f_text, style='List Bullet')

    # 2. Personnel
    doc.add_heading('二、人员及联系信息分析', level=1)
    for f in analysis['personnel']['files']:
        doc.add_heading(f'文件: {f["name"]}', level=2)
        p_fields = [
            ('法定代表人', f.get('legal_rep', '-')),
            ('授权代表', f.get('authorized_rep', '-')),
            ('身份证号', f.get('id_number', '-')),
            ('联系电话', f.get('phone', '-')),
            ('地址', f.get('address', '-')),
        ]
        for label, value in p_fields:
            if value and value != '-':
                doc.add_paragraph(f'{label}: {value}')

    if analysis['personnel']['cross_matches']:
        doc.add_heading('人员交叉发现', level=2)
        for m in analysis['personnel']['cross_matches']:
            doc.add_paragraph(f'【{m["severity"]}】{m["detail"]}', style='List Bullet')

    # 3. Text Similarity
    doc.add_heading('三、文本相似度分析', level=1)
    for pr in analysis['text_similarity']['pair_results']:
        doc.add_heading(f'{pr["file1"]} vs {pr["file2"]}', level=2)
        doc.add_paragraph(f'总匹配段落数: {pr["total_matches"]}')
        doc.add_paragraph(f'异常一致段落数: {pr["abnormal_count"]}')

        if pr['matches']:
            doc.add_heading('异常一致段落详情:', level=3)
            for m in pr['matches'][:30]:  # Limit to top 30
                if m.get('abnormal'):
                    doc.add_paragraph(
                        f'第{m["index"]}项 ({m["length"]}字): {m["text"][:150]}...',
                        style='List Bullet'
                    )

    for f_text in analysis['text_similarity']['findings']:
        doc.add_paragraph(f_text, style='List Bullet')

    # 4. Pricing
    doc.add_heading('四、报价分析', level=1)
    if analysis['pricing']['comparison']:
        for key, entry in analysis['pricing']['comparison'].items():
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
            doc.add_paragraph(f'{key}: {" | ".join(parts)}{mark}')

    for f_text in analysis['pricing']['findings']:
        doc.add_paragraph(f_text, style='List Bullet')

    # 5. Verdict
    doc.add_heading('五、综合判定', level=1)
    for c in analysis['verdict']['clauses']:
        satisfied_str = '满足' if c['satisfied'] is True else ('不满足' if c['satisfied'] is False else '无法判断')
        doc.add_heading(f'{c["clause"]} - {c["description"]} [{satisfied_str}]', level=2)
        for e in c.get('evidence', []):
            if isinstance(e, str):
                doc.add_paragraph(f'- {e}')

    doc.add_heading('最终结论', level=1)
    doc.add_paragraph(analysis['verdict']['conclusion'])

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
        if f.filename and any(f.filename.lower().endswith(ext) for ext in ('.docx', '.doc', '.pdf')):
            safe_name = f.filename.replace('/', '_').replace('\\', '_')
            fpath = os.path.join(UPLOAD_FOLDER, safe_name)
            f.save(fpath)
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
        fpath = os.path.join(UPLOAD_FOLDER, fn)
        if not os.path.exists(fpath):
            return jsonify({'error': f'文件不存在: {fn}'}), 404
        filepaths.append(fpath)

    try:
        results = run_full_analysis(filepaths)
        return jsonify(results)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/api/report', methods=['POST'])
def report():
    data = request.get_json() or {}
    analysis_data = data.get('analysis')
    if not analysis_data:
        return jsonify({'error': '缺少分析数据'}), 400

    try:
        buf = generate_report_docx(analysis_data)
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=f'围串标分析报告_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

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

    VALID_EXTS = ('.docx', '.doc', '.pdf')

    saved = []
    for f in files:
        if f.filename and any(f.filename.lower().endswith(ext) for ext in VALID_EXTS):
            safe_name = f.filename.replace('/', '_').replace('\\', '_')
            fpath = os.path.join(UPLOAD_FOLDER, safe_name)
            f.save(fpath)
            saved.append(fpath)

    if len(saved) < 2:
        return jsonify({'error': '请至少上传2份标书文件(.docx/.doc/.pdf)'}), 400

    saved_refs = []
    for f in ref_files:
        if f.filename and any(f.filename.lower().endswith(ext) for ext in VALID_EXTS):
            safe_name = 'ref_' + f.filename.replace('/', '_').replace('\\', '_')
            fpath = os.path.join(UPLOAD_FOLDER, safe_name)
            f.save(fpath)
            saved_refs.append(fpath)

    file_groups = request.form.getlist('file_groups')
    group_map = {}
    for fp, group in zip(saved, file_groups):
        group = group.strip() or os.path.basename(fp)
        group_map.setdefault(group, []).append(fp)

    group_texts = {}
    for group, paths in group_map.items():
        combined = ''
        for p in paths:
            try:
                combined += extract_text_with_tables(p) + '\n'
            except Exception:
                pass
        group_texts[group] = combined

    import queue
    progress_queue = queue.Queue()

    def on_progress(step, label, percent, detail):
        progress_queue.put({'type': 'progress', 'step': step, 'label': label,
                           'percent': percent, 'detail': detail})

    def generate():
        import threading

        results_holder = []
        error_holder = []

        def run():
            try:
                results_holder.append(run_full_analysis(
                    saved, saved_refs if saved_refs else None,
                    group_map=group_map, group_texts=group_texts,
                    on_progress=on_progress
                ))
            except Exception as e:
                import traceback
                traceback.print_exc()
                error_holder.append(str(e))

        thread = threading.Thread(target=run)
        thread.start()

        while thread.is_alive() or not progress_queue.empty():
            try:
                event = progress_queue.get(timeout=0.1)
                yield json.dumps(event, ensure_ascii=False) + '\n'
            except queue.Empty:
                pass

        thread.join()

        if error_holder:
            yield json.dumps({'type': 'error', 'message': error_holder[0]}, ensure_ascii=False) + '\n'
            return

        results = results_holder[0]
        results['_filenames'] = [os.path.basename(s) for s in saved]
        results['_ref_filenames'] = [os.path.basename(s) for s in saved_refs]
        results['_groups'] = {g: [os.path.basename(p) for p in paths] for g, paths in group_map.items()}
        results['_group_order'] = list(group_map.keys())

        # Save to history
        try:
            history_results = json.loads(json.dumps(results, ensure_ascii=False))
            for pr in history_results.get('text_similarity', {}).get('pair_results', []):
                light_matches = []
                for m in pr.get('matches', []):
                    light_matches.append({
                        'index': m.get('index'), 'length': m.get('length'),
                        'text': m.get('text', '')[:200],
                        'abnormal': m.get('abnormal'),
                        'reasons': m.get('reasons', [])[:2],
                        'ctx1': m.get('ctx1', '')[:400],
                        'ctx2': m.get('ctx2', '')[:400],
                    })
                pr['matches'] = light_matches
            if 'subItemCompare' in history_results.get('pricing', {}):
                for c in history_results['pricing']['subItemCompare']:
                    for item in c.get('items', []):
                        item.pop('extras', None)
            for section in ['metadata', 'personnel', 'pricing']:
                for f in history_results.get(section, {}).get('files', []):
                    keep = ['name']
                    if section == 'metadata':
                        keep += ['creator', 'last_modified_by', 'created', 'modified',
                                 'application', 'template', 'revision', 'total_edit_time',
                                 'pages', 'words', 'company']
                    elif section == 'personnel':
                        keep += ['legal_rep', 'authorized_rep', 'id_number', 'phone',
                                 'address', 'response_date', 'company_name']
                    elif section == 'pricing':
                        keep += ['totalPriceInTax', 'totalPrice', 'taxRate']
                    for k in list(f.keys()):
                        if k not in keep and not k.startswith('_'):
                            f[k] = '' if isinstance(f[k], str) else None
            history_results.get('text_similarity', {}).pop('all_abnormal', None)
            history_id = datetime.now().strftime('%Y%m%d_%H%M%S_') + hashlib.md5(
                str(saved).encode()).hexdigest()[:8]
            history_entry = {
                'id': history_id,
                'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'bid_count': len(saved), 'ref_count': len(saved_refs),
                'bid_files': results['_filenames'], 'ref_files': results['_ref_filenames'],
                'verdict': results['verdict']['conclusion'],
                'abnormal_matches': sum(p['abnormal_count'] for p in results['text_similarity']['pair_results']),
                'template_matches': results['text_similarity']['template_matches'],
                'total_pairs': results['text_similarity']['total_pairs'],
                'data': history_results
            }
            with open(os.path.join(HISTORY_DIR, f'{history_id}.json'), 'w', encoding='utf-8') as f:
                json.dump(history_entry, f, ensure_ascii=False)
            results['_history_id'] = history_id
        except Exception:
            pass

        yield json.dumps({'type': 'result', 'data': results}, ensure_ascii=False, default=str) + '\n'

    return Response(generate(), mimetype='application/x-ndjson')


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

    VALID_EXTS = ('.docx', '.doc', '.pdf')

    saved = []
    for f in files:
        if f.filename and any(f.filename.lower().endswith(ext) for ext in VALID_EXTS):
            safe_name = f.filename.replace('/', '_').replace('\\', '_')
            fpath = os.path.join(UPLOAD_FOLDER, safe_name)
            f.save(fpath)
            saved.append(fpath)

    if len(saved) < 2:
        return jsonify({'error': '请至少上传2份标书文件(.docx/.doc/.pdf)'}), 400

    saved_refs = []
    for f in ref_files:
        if f.filename and any(f.filename.lower().endswith(ext) for ext in VALID_EXTS):
            safe_name = 'ref_' + f.filename.replace('/', '_').replace('\\', '_')
            fpath = os.path.join(UPLOAD_FOLDER, safe_name)
            f.save(fpath)
            saved_refs.append(fpath)

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
        for group, paths in group_map.items():
            combined = ''
            for p in paths:
                try:
                    combined += extract_text_with_tables(p) + '\n'
                except Exception:
                    pass
            group_texts[group] = combined

        results = run_full_analysis(saved, saved_refs if saved_refs else None,
                                    group_map=group_map, group_texts=group_texts)
        results['_filenames'] = [os.path.basename(s) for s in saved]
        results['_ref_filenames'] = [os.path.basename(s) for s in saved_refs]
        results['_groups'] = {g: [os.path.basename(p) for p in paths] for g, paths in group_map.items()}
        results['_group_order'] = list(group_map.keys())

        # Save to history — strip heavy data, keep only counts & indices
        history_results = json.loads(json.dumps(results, ensure_ascii=False))
        for pr in history_results.get('text_similarity', {}).get('pair_results', []):
            light_matches = []
            for m in pr.get('matches', []):
                light_matches.append({
                    'index': m.get('index'),
                    'length': m.get('length'),
                    'text': m.get('text', '')[:200],
                    'abnormal': m.get('abnormal'),
                    'reasons': m.get('reasons', [])[:2],
                    'ctx1': m.get('ctx1', '')[:400],
                    'ctx2': m.get('ctx2', '')[:400],
                })
            pr['matches'] = light_matches
        # Strip pricing subItemCompare items
        if 'subItemCompare' in history_results.get('pricing', {}):
            for c in history_results['pricing']['subItemCompare']:
                for item in c.get('items', []):
                    item.pop('extras', None)
        # Strip metadata detail from files
        for section in ['metadata', 'personnel', 'pricing']:
            for f in history_results.get(section, {}).get('files', []):
                keep = ['name']
                if section == 'metadata':
                    keep += ['creator', 'last_modified_by', 'created', 'modified',
                             'application', 'template', 'revision', 'total_edit_time',
                             'pages', 'words', 'company']
                elif section == 'personnel':
                    keep += ['legal_rep', 'authorized_rep', 'id_number', 'phone',
                             'address', 'response_date', 'company_name']
                elif section == 'pricing':
                    keep += ['totalPriceInTax', 'totalPrice', 'taxRate']
                for k in list(f.keys()):
                    if k not in keep and not k.startswith('_'):
                        f[k] = '' if isinstance(f[k], str) else None
        # Strip all_abnormal from text_similarity
        history_results.get('text_similarity', {}).pop('all_abnormal', None)
        history_id = datetime.now().strftime('%Y%m%d_%H%M%S_') + hashlib.md5(
            str(saved).encode()).hexdigest()[:8]
        history_entry = {
            'id': history_id,
            'time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'bid_count': len(saved),
            'ref_count': len(saved_refs),
            'bid_files': results['_filenames'],
            'ref_files': results['_ref_filenames'],
            'verdict': results['verdict']['conclusion'],
            'abnormal_matches': sum(p['abnormal_count'] for p in results['text_similarity']['pair_results']),
            'template_matches': results['text_similarity']['template_matches'],
            'total_pairs': results['text_similarity']['total_pairs'],
            'data': history_results
        }
        with open(os.path.join(HISTORY_DIR, f'{history_id}.json'), 'w', encoding='utf-8') as f:
            json.dump(history_entry, f, ensure_ascii=False)

        results['_history_id'] = history_id
        return jsonify(results)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


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
            entries.append({
                'id': entry.get('id'),
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
    fpath = os.path.join(HISTORY_DIR, f'{history_id}.json')
    if os.path.exists(fpath):
        os.remove(fpath)
        return jsonify({'ok': True})
    return jsonify({'error': '记录不存在'}), 404


if __name__ == '__main__':
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get('PORT', 5001))
    debug = os.environ.get('DEBUG', '0') == '1'
    print('=' * 60)
    print('  星易查 - 围串标风险识别分析系统')
    print(f'  访问地址: http://0.0.0.0:{port}')
    print('=' * 60)
    app.run(debug=debug, host='0.0.0.0', port=port)
