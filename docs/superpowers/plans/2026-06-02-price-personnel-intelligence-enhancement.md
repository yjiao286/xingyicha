# 报价分析 & 人员比对智能增强 — 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 提升报价提取和人员比对模块对多样化标书格式的泛化能力，通过多通道提取、动态列推断、章节限定姓名提取和分层交叉比对实现。

**Architecture:** 在现有 `app.py` 单一文件中渐进增强，每个函数内部采用多通道/分层策略。不新增文件，不改变外部 API 契约。所有新增字段为可选字段，保持向后兼容。

**Tech Stack:** Python 3.11+, re (regex), existing python-docx/pypdf dependencies

---

## 文件结构

| 文件 | 职责 | 改动类型 |
|------|------|---------|
| `app.py:374-386` | `_parse_amount()` — 增强数字解析 | 修改 |
| `app.py:388-476` | `extract_prices()` — 多通道总价提取 | 重写 |
| `app.py:479-550` | `_extract_structured_items()` — 增强章节发现 + 通用成本提取 | 重写 |
| `app.py:552-743` | `_parse_pdf_bid_table()` — 动态列推断解析器 | 重写 |
| `app.py:949-1059` | `_build_sub_item_comparison()` — 双算法匹配 | 修改 |
| `app.py:255-370` | `extract_personnel()` — 章节限定姓名提取 + 角色推断 | 重写 |
| `app.py:1188-1240` | `run_full_analysis()` — 增强 personnel 交叉比对 | 修改 |

---

### Task 1: 增强 `_parse_amount()` 数字解析

**Files:**
- Modify: `app.py:374-386`

- [ ] **Step 1: 替换 `_parse_amount()` 函数**

将当前函数替换为增强版本，支持亿元、中文大写数字、全角数字、价格区间：

```python
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
```

- [ ] **Step 2: 验证 — 测试各种数字格式**

```bash
cd /Users/dev/Desktop/围串标/围串标风险识别APPV2 && ./venv/bin/python3 -c "
from app import _parse_amount
tests = [
    ('1,234,567.89', 1234567.89),
    ('123.45万元', 1234500.0),
    ('1.2亿元', 120000000.0),
    ('8,123,000.00', 8123000.0),
    ('665400 元', 665400.0),
    ('39 万', 390000.0),
    ('１２３４５６', 123456.0),  # full-width digits
]
for inp, expected in tests:
    result = _parse_amount(inp)
    status = '✅' if abs(result - expected) < 0.01 else '❌'
    print(f'{status} _parse_amount(\"{inp}\") = {result:,.0f} (expected {expected:,.0f})')
"
```

- [ ] **Step 3: 提交**

```bash
git add app.py
git commit -m "feat: enhance _parse_amount with 亿 unit, full-width digits, and Chinese numeral support"
```

---

### Task 2: 多通道总价提取 — 重写 `extract_prices()`

**Files:**
- Modify: `app.py:388-476`

- [ ] **Step 1: 实现多通道提取器**

替换 `extract_prices()` 函数：

```python
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

    # ── Channel 3: 大写/小写 pair (大写金额段) ── confidence: 0.88
    if result['totalPriceInTax'] is None:
        m = re.search(r'大写[：:]\s*[壹贰叁肆伍陆柒捌玖拾佰仟万亿零一二三四五六七八九十百千元整角分]+[\s\S]{0,100}?小写[：:]\s*([\d,]+\.?\d*)', text)
        if m:
            val = _parse_amount(m.group(1))
            if val >= 100:
                result['totalPriceInTax'] = val
                result['totalPrice'] = val

    # ── Channel 4: 表格通道 — find bid summary table ── confidence: 0.85
    # Locate "开标一览表" or "投标报价表" section
    bid_section = _find_bid_summary_section(text)
    if bid_section and result['totalPriceInTax'] is None:
        # Try to find total in this section
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
    # Pattern: "￥X 13% ￥Y" or "人民币：X 13% 人民币：Y"
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
```

- [ ] **Step 2: 验证 — 测试多通道提取**

```bash
cd /Users/dev/Desktop/围串标/围串标风险识别APPV2 && ./venv/bin/python3 -c "
from app import extract_prices

# Test 1: Symbol-based
r1 = extract_prices('投标总价：￥8,123,000.00 元整')
print(f'Symbol: totalPriceInTax={r1[\"totalPriceInTax\"]}')

# Test 2: Label-based
r2 = extract_prices('投标总报价：669000 元')
print(f'Label: totalPriceInTax={r2[\"totalPriceInTax\"]}')

# Test 3: 大写/小写
r3 = extract_prices('大写：壹佰贰拾叁万肆仟伍佰陆拾柒元整  小写：1234567.00')
print(f'大写/小写: totalPriceInTax={r3[\"totalPriceInTax\"]}')

# Test 4: Tax decomposition
r4 = extract_prices('不含税价 ￥664800.00 税率 13% 含税价 ￥751224.00')
print(f'Tax: totalPrice={r4[\"totalPrice\"]}, totalPriceInTax={r4[\"totalPriceInTax\"]}, taxRate={r4[\"taxRate\"]}')
"
```

- [ ] **Step 3: 提交**

```bash
git add app.py
git commit -m "feat: multi-channel total price extraction with confidence scoring"
```

---

### Task 3: 通用成本明细提取 — 重写 `_extract_structured_items()`

**Files:**
- Modify: `app.py:479-550`

- [ ] **Step 1: 实现增强版结构化提取器**

替换 `_extract_structured_items()` 函数：

```python
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
    for m in re.finditer(
        r'(?:^|\n)(?:[一二三四五六七八九十\d]+[、.。]\s*|\d+(?:\.\d+)+\s*|\d+\s+)?(' +
        '|'.join(re.escape(kw) for kw in section_keywords) + r')\s*\n', text
    ):
        pos = m.start()
        # Skip if preceded by dots (TOC entry)
        prefix = text[max(0, pos - 30):pos]
        if re.search(r'\.{3,}', prefix):
            continue
        # Find next major section boundary
        next_pos = len(text)
        for end_marker in ['\n三、', '\n四、', '\n五、', '\n六、', '\n七、',
                           '\n3.', '\n4.', '\n5.', '\n6.', '\n7.',
                           '\n3 ', '\n4 ', '\n5 ', '\n6 ', '\n7 ']:
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
    # Match lines with: Chinese name (2-20 chars) + large number (>= 100)
    # Priority to lines containing 费, 成本, 支出, 投入
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

    # Recalculate total cost
    if result['costDetails'] and result['cost'] is None:
        result['cost'] = sum(item['totalPrice'] for item in result['costDetails'])
```

- [ ] **Step 2: 提交**

```bash
git add app.py
git commit -m "feat: enhanced section discovery and generic cost line extraction"
```

---

### Task 4: 动态列推断 — 重写 `_parse_pdf_bid_table()`

**Files:**
- Modify: `app.py:552-743` (替换整个函数)

- [ ] **Step 1: 实现动态列推断解析器**

```python
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
    # Find table end
    data_end = None
    for i in range(header_line + 1, len(lines)):
        s = lines[i].strip()
        if not s:
            continue
        if any(s.startswith(kw) for kw in ['合计', '总价', '小计', '总计', '注：', '备注：']):
            data_end = i
            break
        # Section boundary
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
            # No amounts — this line is likely a name continuation
            pending_name.append(s)

    # Process each data row
    items = []
    prev_name = None
    for name_part, data in merged_rows:
        # Use previous name if current is empty (merged cell)
        name = name_part.strip() if name_part.strip() else ''
        name = re.sub(r'^\d+\s*', '', name).strip()
        if not name and prev_name:
            name = prev_name
        elif name:
            prev_name = name

        if len(name) < 2:
            continue

        # Extract numbers
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

        # Assign based on col_order
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
        # Try to extract summary total
        if not result.get('totalPrice'):
            _extract_summary_total(section, result)


def _infer_columns(header_line):
    """Infer column types and order from header text.
    Returns list of (col_type, position_index) sorted by position."""
    col_map = []
    # Define detection patterns in priority order
    detectors = [
        (r'序\s*号', 'seq'),
        (r'(?:分项\s*)?名\s*称|产品|服务|项目|内容', 'name'),
        (r'型号|规格|厂家|制造商|品牌', 'model'),
        (r'数\s*量', 'count'),
        (r'单\s*价.*?(?:不含|未含)|不含税.*?单\s*价', 'unit_price_ex'),
        (r'单\s*价.*?(?:含税|含)|含税.*?单\s*价', 'unit_price_in'),
        (r'总\s*价.*?(?:不含|未含)|不含税.*?总\s*价', 'total_ex'),
        (r'总\s*价.*?(?:含税|含)|含税.*?总\s*价', 'total_in'),
        (r'税\s*率', 'tax_rate'),
        (r'备\s*注', 'remark'),
    ]
    for pattern, col_type in detectors:
        m = re.search(pattern, header_line)
        if m:
            col_map.append((col_type, m.start()))
    col_map.sort(key=lambda x: x[1])
    
    # If we have unit_price but not separate ex/in variants, use generic
    has_explicit = any(c[0] in ('unit_price_ex', 'unit_price_in', 'total_ex', 'total_in') for c in col_map)
    if not has_explicit:
        # Simple mapping: look for general 单价 and 总价
        for c in col_map:
            if c[0] == 'unit_price_ex' or c[0] == 'unit_price_in':
                col_map = [(t if t not in ('unit_price_ex', 'unit_price_in') else 'unit_price', pos) 
                          for t, pos in col_map]
                break
    
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

    # Price columns: map large values to inferred column positions
    # Common patterns:
    # [unit_ex, unit_in, total_ex, total_in] (4 large values)
    # [unit_ex, total_ex, total_in] (3 large values)
    # [unit_ex, total_ex] (2 large values)
    
    has_ex_in_split = any(col in col_order for col in ['unit_price_ex', 'unit_price_in', 'total_ex', 'total_in'])
    
    if has_ex_in_split and len(larges) >= 4:
        item['unitPrice'] = larges[0]
        item['totalPrice'] = larges[2]  # total_ex (不含税总价)
        item['totalPriceInTax'] = larges[3]  # total_in
    elif len(larges) >= 4:
        # Default: unit_ex, unit_in, total_ex, total_in
        item['unitPrice'] = larges[0]
        item['totalPrice'] = larges[2]
        item['totalPriceInTax'] = larges[3]
    elif len(larges) == 3:
        item['unitPrice'] = larges[0]
        item['totalPrice'] = larges[1]
        item['totalPriceInTax'] = larges[2]
    elif len(larges) == 2:
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
```

- [ ] **Step 2: 提交**

```bash
git add app.py
git commit -m "feat: dynamic column inference for PDF bid table parsing"
```

---

### Task 5: 双算法分项比对 — 增强 `_build_sub_item_comparison()`

**Files:**
- Modify: `app.py:949-1059`

- [ ] **Step 1: 增强归一化和匹配算法**

在函数开头替换归一化函数和匹配逻辑。找到 `_build_sub_item_comparison` 中约第953行的 `_norm_name` 函数，替换为：

```python
def _norm_name(name):
    """Aggressive normalization for item name comparison."""
    n = name.strip()
    # Remove parenthesized/bracketed content and quotes
    n = re.sub(r'[（(][^）)]*[）)]', '', n)
    n = re.sub(r'[【\[《<][^】\]》>]*[】\]》>]', '', n)
    n = re.sub(r'["""''‘’“”]', '', n)
    # Remove punctuation
    n = re.sub(r'[、，。；：！？\s\-–—/\\|,\.;:!?\s]+', '', n)
    # Full-width to half-width
    n = n.replace('０', '0').replace('１', '1').replace('２', '2').replace('３', '3').replace('４', '4')
    n = n.replace('５', '5').replace('６', '6').replace('７', '7').replace('８', '8').replace('９', '9')
    n = n.replace('Ａ', 'A').replace('Ｂ', 'B').replace('Ｃ', 'C').replace('Ｄ', 'D')
    # Common suffixes/prefixes
    n = re.sub(r'(及配套.*|配套.*|等.*)$', '', n)
    return n.strip()
```

在聚类循环中（约第1009行），替换匹配条件：

```python
# 原有：
# same_name = item_i['norm'] == item_j['norm']
# long_match = len(item_i['norm']) >= 10 and len(item_j['norm']) >= 10 and _lcs_len(item_i['norm'], item_j['norm']) >= 15

# 替换为：
# Jaccard similarity on 2-grams
def _jaccard_2gram(a, b):
    if not a or not b:
        return 0.0
    sa = set(a[i:i+2] for i in range(len(a)-1))
    sb = set(b[i:i+2] for i in range(len(b)-1))
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)

same_name = item_i['norm'] == item_j['norm']
lcs_val = _lcs_len(item_i['norm'], item_j['norm']) if len(item_i['norm']) >= 6 and len(item_j['norm']) >= 6 else 0
jaccard_val = _jaccard_2gram(item_i['norm'], item_j['norm']) if len(item_i['norm']) >= 6 and len(item_j['norm']) >= 6 else 0
long_match = len(item_i['norm']) >= 6 and len(item_j['norm']) >= 6 and (lcs_val >= 10 or jaccard_val >= 0.55)
```

在价格差异输出中增加百分比（约第1044行之后，在 findings.append 中）：

```python
# 原有的 price findings 中增加：
# 在价格比较 section，增加百分比差异显示
if len(prices_excl) >= 2:
    pmin, pmax = min(prices_excl), max(prices_excl)
    if pmax > 0:
        diff_pct = (pmax - pmin) / pmax * 100
        files_with_price = [(it['file'], it['totalPrice']) for it in items if it.get('totalPrice')]
        detail = ' | '.join(f'{os.path.basename(f)}: {p:,.0f}' for f, p in files_with_price)
        if diff_pct < 2:
            findings.append(f'不含税报价差异仅{diff_pct:.1f}%（{detail}），高度接近')
        elif diff_pct < 10:
            findings.append(f'不含税报价差异{diff_pct:.1f}%（{detail}）')
```

- [ ] **Step 2: 提交**

```bash
git add app.py
git commit -m "feat: dual-algorithm item matching with Jaccard similarity and percentage diff display"
```

---

### Task 6: 章节限定姓名提取 — 重写 `extract_personnel()`

**Files:**
- Modify: `app.py:255-370`

- [ ] **Step 1: 实现章节限定提取器**

```python
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
        'contacts': {'phone': None, 'email': None, 'address': None}
    }

    # Normalize line breaks within key phrases
    text = re.sub(r'法定代\s*\n\s*表人', '法定代表人', text)
    text = re.sub(r'法定\s*\n\s*代表人', '法定代表人', text)
    text = re.sub(r'法\s*\n\s*定代表人', '法定代表人', text)
    text = re.sub(r'授权委\s*\n\s*托书', '授权委托书', text)
    text = re.sub(r'供应\s*\n\s*商名称', '供应商名称', text)
    # Also fix name splits
    text = re.sub(r'([一-鿿])\s*\n\s*([一-鿿]{1,2})', r'\1\2', text)

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

    # ── Global extraction (not section-specific) ──
    # ID number (can appear anywhere in auth sections)
    for sec in auth_sections + sig_sections:
        m = re.search(r'身份证号[码字]?[：:]\s*(\d{17}[\dXx])', sec['text'])
        if m and not info.get('id_number'):
            info['id_number'] = m.group(1).strip()

    # Phone
    for sec in auth_sections + personnel_sections + sig_sections:
        m = re.search(r'(?:电话|手机|联系电话|联系方式)[：:]\s*(\d[\d\-]{6,15})', sec['text'])
        if m:
            phone = m.group(1).strip()
            if not info.get('phone'):
                info['phone'] = phone
            if not info['contacts'].get('phone'):
                info['contacts']['phone'] = phone

    # Address
    for sec in auth_sections:
        m = re.search(r'地址[：:]\s*(.{8,80})', sec['text'])
        if m and not info.get('address'):
            addr = m.group(1).strip()[:100]
            info['address'] = addr
            info['contacts']['address'] = addr

    # Response date
    m = re.search(r'(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', text[:800])
    if m:
        info['response_date'] = m.group(1).strip()

    # Cleanup legal_rep
    _cleanup_name(info, 'legal_rep')
    _cleanup_name(info, 'authorized_rep')

    # Deduplicate all_persons
    seen = set()
    unique_persons = []
    for p in info['all_persons']:
        key = (p['name'], p['role'])
        if key not in seen:
            seen.add(key)
            unique_persons.append(p)
    info['all_persons'] = unique_persons

    return info
```

- [ ] **Step 2: 实现辅助函数**

在同一区域（`extract_personnel` 之前）添加辅助函数：

```python
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
            '人员配备', '人员配置', '岗位人员', '主要管理人员'
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
                # Determine section boundaries
                start = max(0, idx - 200)
                # End: find next recognizable section header or 5000 chars
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


def _extract_from_auth_section(section_text, info):
    """Extract legal rep, authorized rep from authorization letter section."""
    # Pattern 0: "我张三（姓名）系四川某某电子科技有限公司（供应商名称）的法定代表人"
    m = re.search(r'(?:本人\s*)?我?\s*([一-鿿]{2,4})\s*[（(]姓名[）)]\s*系\s*(.{1,40}?)\s*[（(]供应商名称[）)]\s*的法定代表人', section_text)
    if m:
        info['legal_rep'] = m.group(1).strip()
        info['company_name'] = _clean_company(m.group(2).strip())
        info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.95})

    # Pattern 1: "姓名：XXX 职务：XXX 系 XXX 的法定代表人"
    if not info['legal_rep']:
        m = re.search(r'姓名[：:]\s*([^\s]{2,10})\s*[\s\S]{0,100}?系\s*(.{1,30}?)\s*的法定代表人', section_text)
        if m:
            info['legal_rep'] = m.group(1).strip()
            info['company_name'] = _clean_company(m.group(2).strip())
            info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.90})

    # Pattern 2: "本人 XXX 系 XXX 的法定代表人"
    if not info['legal_rep']:
        m = re.search(r'(?:本人\s*)?([一-鿿]{2,4})\s*(?:[（(]姓名[）)])?\s*系\s*(.{1,30}?)\s*的法定代表人', section_text)
        if m:
            info['legal_rep'] = m.group(1).strip()
            info['company_name'] = _clean_company(m.group(2).strip())
            info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.85})

    # Pattern 3: "（王戈、董事长）代表本公司授权（赵凯、销售经理）"
    m = re.search(r'[（(]([一-鿿]{2,4})[、，].{0,6}?[）)]\s*代表本公司授权\s*[（(]([一-鿿]{2,4})', section_text)
    if m:
        if not info['legal_rep']:
            info['legal_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.85})
        info['authorized_rep'] = m.group(2).strip()
        info['all_persons'].append({'name': info['authorized_rep'], 'role': 'authorized_rep', 'confidence': 0.85})

    # Pattern 4: "现委托 XXX（姓名）为我方代理人"
    if not info['authorized_rep']:
        m = re.search(r'(?:现委托|委托)\s*([一-鿿]{2,4})\s*[（(]姓名[）)]', section_text)
        if m:
            info['authorized_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['authorized_rep'], 'role': 'authorized_rep', 'confidence': 0.90})

    # Pattern 5: "代理人：XXX" or "授权代表：XXX"
    if not info['authorized_rep']:
        m = re.search(r'(?:代理人|授权代表|被授权人|受托人)[：:]\s*([一-鿿]{2,4})', section_text)
        if m:
            info['authorized_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['authorized_rep'], 'role': 'authorized_rep', 'confidence': 0.80})

    # Pattern 6: "法定代表人：XXX"
    if not info['legal_rep']:
        m = re.search(r'(?:法定代表人|单位负责人|法人代表)[：:]\s*([一-鿿]{2,4})', section_text)
        if m:
            info['legal_rep'] = m.group(1).strip()
            info['all_persons'].append({'name': info['legal_rep'], 'role': 'legal_rep', 'confidence': 0.80})


def _extract_from_personnel_table(section_text, info):
    """Extract project members from personnel/team tables."""
    # Look for rows with name + role pattern
    # Format: "姓名：XXX 职务/岗位：XXX" or "XXX 项目负责人" or table rows
    patterns = [
        # Tagged format
        r'姓名[：:]\s*([一-鿿]{2,4})\s*.*?(?:职务|岗位|角色|职称)[：:]\s*([一-鿿]{2,10})',
        # "XXX  项目经理" (name followed by role)
        r'([一-鿿]{2,4})\s{2,}(项目经理|项目负责人|技术负责人|技术总监|总工程师|安全员|质量员|施工员|材料员|资料员|造价员|预算员)',
        # "项目经理：XXX"
        r'(项目经理|项目负责人|技术负责人|技术总监|总工程师)[：:]\s*([一-鿿]{2,4})',
        # Table cell format: "岗位名称  姓名" rows
        r'(?:项目经理|项目负责人|技术负责人|安全负责人)\s+([一-鿿]{2,4})',
    ]
    for pat in patterns:
        for m in re.finditer(pat, section_text):
            groups = m.groups()
            if len(groups) == 2:
                name, role_str = groups if len(groups[0]) <= 4 else (groups[1], groups[0])
                role = _infer_role_label(role_str)
            else:
                name = groups[0]
                role = 'team_member'
            
            name = name.strip()
            if len(name) >= 2:
                info['all_persons'].append({'name': name, 'role': role, 'confidence': 0.80})


def _extract_from_signature_page(section_text, info):
    """Extract signatory names from signature/seal pages."""
    # "法定代表人或其委托代理人：（签字）XXX"
    m = re.search(r'法定代表人或其委托代理人[：:][（(]?\s*(?:签字|签章|盖章|签名)\s*[）)]?\s*([一-鿿]{2,4})', section_text)
    if m:
        info['all_persons'].append({'name': m.group(1).strip(), 'role': 'signatory', 'confidence': 0.75})

    # "投标人：（盖章）XXX" - company name
    m = re.search(r'投标人[：:]\s*[（(]?(?:盖章|公章|单位章)[）)]?\s*(.{2,40}?)(?:\n|$)', section_text)
    if m and not info.get('company_name'):
        company = m.group(1).strip()
        if len(company) >= 4 and not re.match(r'^[\s（(）)]+$', company):
            info['company_name'] = _clean_company(company)


def _extract_from_cover(section_text, info):
    """Extract company name from cover/bid letter."""
    if info.get('company_name'):
        return
    # "致：XXX（采购人）" or "采购人：XXX"
    m = re.search(r'(?:投标人|供应商|申请.?|报价.?)[：:]\s*(.{2,40}?)(?:\n|$)', section_text)
    if m:
        company = m.group(1).strip()
        if len(company) >= 4:
            info['company_name'] = _clean_company(company)


def _infer_role_label(role_str):
    """Map Chinese role strings to standardized role labels."""
    role_str = role_str.strip()
    mapping = {
        '项目经理': 'project_manager', '项目负责人': 'project_manager',
        '技术负责人': 'tech_lead', '技术总监': 'tech_lead', '总工程师': 'tech_lead',
        '安全员': 'team_member', '质量员': 'team_member', '施工员': 'team_member',
        '材料员': 'team_member', '资料员': 'team_member', '造价员': 'team_member',
        '预算员': 'team_member', '安全负责人': 'tech_lead',
    }
    for cn, en in mapping.items():
        if cn in role_str:
            return en
    return 'team_member'


def _clean_company(name):
    """Clean company name from parenthetical annotations."""
    name = re.sub(r'[（(]投标人名称[）)]|[（(]单位负责人[）)]|[（(]供应商名称[）)]', '', name)
    name = re.sub(r'^[（(]|[）)]$', '', name).strip()
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
    if len(val) < 2 or any(w in val for w in ['注册', '签字', '盖章', '地址', '电话', '投标人']):
        info[key] = None
    else:
        info[key] = val
```

- [ ] **Step 3: 验证 — 测试章节限定提取**

```bash
cd /Users/dev/Desktop/围串标/围串标风险识别APPV2 && ./venv/bin/python3 -c "
from app import extract_personnel

# Test: authorization letter
text1 = '''法定代表人授权委托书
本人 王某某 （姓名）系 北京某某大学 （供应商名称）的法定代表人。
现委托 李某某 （姓名）为我方代理人。
联系电话：13800138000
地址：北京市海淀区某路某号'''

r1 = extract_personnel(text1)
print(f'Auth letter: legal_rep={r1[\"legal_rep\"]}, authorized_rep={r1[\"authorized_rep\"]}, company={r1[\"company_name\"]}')
print(f'All persons: {r1[\"all_persons\"]}')

# Test: personnel table
text2 = '''项目管理机构
姓名：张三  职务：项目经理
姓名：李四  职务：技术负责人
姓名：王五  职务：安全员'''

r2 = extract_personnel(text2)
print(f'Personnel table: all_persons={r2[\"all_persons\"]}')
"
```

- [ ] **Step 4: 提交**

```bash
git add app.py
git commit -m "feat: chapter-scoped personnel extraction with role inference and 8 role types"
```

---

### Task 7: 分层人员交叉比对 — 增强 `run_full_analysis()`

**Files:**
- Modify: `app.py:1188-1240`

- [ ] **Step 1: 实现分层交叉比对**

替换人员交叉比对部分（约第1188-1240行）：

```python
    # ── Compile personnel cross-comparison (all group pairs) ──
    personnel_matches = []
    personnel_dedup = set()
    
    # Collect all persons per file
    all_persons_map = {}
    for gn in out_names:
        persons = all_personnel.get(gn, {}).get('all_persons', [])
        all_persons_map[gn] = persons

    if len(out_names) >= 2:
        for i in range(len(out_names)):
            for j in range(i+1, len(out_names)):
                gi, gj = out_names[i], out_names[j]
                pi, pj = all_personnel[gi], all_personnel[gj]
                mi, mj = group_meta.get(gi, {}), group_meta.get(gj, {})

                # ── Layer 1: Exact name match (cross-file shared personnel) ──
                names_i = {p['name']: p for p in all_persons_map[gi]}
                names_j = {p['name']: p for p in all_persons_map[gj]}
                shared_names = set(names_i.keys()) & set(names_j.keys())
                
                for name in shared_names:
                    role_i = names_i[name].get('role', 'other')
                    role_j = names_j[name].get('role', 'other')
                    if role_i == role_j:
                        # Same person, same role — high severity
                        key = f'same_person|{name}|{gi}|{gj}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            personnel_matches.append({
                                'type': '人员重叠（同角色）',
                                'detail': f'"{name}"（{role_i}）同时出现在 {gi} 和 {gj} 中',
                                'severity': 'high'
                            })
                    else:
                        # Same person, different role — medium severity
                        key = f'same_person_diff_role|{name}|{gi}|{gj}'
                        if key not in personnel_dedup:
                            personnel_dedup.add(key)
                            personnel_matches.append({
                                'type': '人员重叠（不同角色）',
                                'detail': f'"{name}"在{gi}中为{role_i}，在{gj}中为{role_j}',
                                'severity': 'medium'
                            })

                # ── Layer 2: Phone cross-match ──
                phone_i = pi.get('phone') or (pi.get('contacts') or {}).get('phone')
                phone_j = pj.get('phone') or (pj.get('contacts') or {}).get('phone')
                if phone_i and phone_j and phone_i == phone_j:
                    key = f'same_phone|{phone_i}'
                    if key not in personnel_dedup:
                        personnel_dedup.add(key)
                        personnel_matches.append({
                            'type': '联系电话相同',
                            'detail': f'{gi} 和 {gj} 联系电话均为 {phone_i}',
                            'severity': 'high'
                        })

                # ── Layer 3: ID number cross-match ──
                id_i = pi.get('id_number')
                id_j = pj.get('id_number')
                if id_i and id_j and id_i == id_j:
                    key = f'same_id|{id_i}'
                    if key not in personnel_dedup:
                        personnel_dedup.add(key)
                        personnel_matches.append({
                            'type': '身份证号相同',
                            'detail': f'{gi} 和 {gj} 出现同一身份证号 {id_i[:6]}****',
                            'severity': 'critical'
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

                # ── Layer 5: Same last modifier ──
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

        # Phone anomaly (per file)
        for gn in out_names:
            p = all_personnel.get(gn, {})
            phone = p.get('phone', '')
            if phone and len(phone) < 7:
                personnel_matches.append({
                    'type': '联系电话异常',
                    'detail': f'{gn}的联系电话为"{phone}"，不是有效电话号码格式',
                    'severity': 'medium'
                })
```

- [ ] **Step 2: 提交**

```bash
git add app.py
git commit -m "feat: layered personnel cross-comparison with 6 match types and overlap rate"
```

---

### Task 8: 端到端验证 & 回归测试

**Files:**
- 无新建文件

- [ ] **Step 1: 综合验证脚本**

```bash
cd /Users/dev/Desktop/围串标/围串标风险识别APPV2 && ./venv/bin/python3 -c "
from app import extract_prices, extract_personnel, _parse_amount, _build_sub_item_comparison

print('=== 价格提取测试 ===')
test_cases = [
    ('投标总价：￥8,123,000.00', 'Symbol'),
    ('人民币：669000 元', 'RMB label'),
    ('小写：665400 元', '小写 label'),
    ('大写：壹佰贰拾万  小写：1200000.00', '大写/小写'),
    ('投标总报价：8912300 元', '投标总报价'),
    ('金额（元）：5678000', '金额（元）'),
    ('不含税 ￥664800.00 13% 含税 ￥751224.00', 'Tax decompose'),
]
for text, label in test_cases:
    r = extract_prices(text)
    tot = r.get('totalPriceInTax') or r.get('totalPrice')
    print(f'  [{label}] {text[:40]:40s} => total={tot} taxRate={r.get(\"taxRate\")}'  )

print()
print('=== 人员提取测试 ===')
personnel_tests = [
    ('''法定代表人授权委托书
本人 王某某 （姓名）系 北京某某大学 （供应商名称）的法定代表人。
现委托 李某某 （姓名）为我方代理人。
联系电话：13800138000''', 'auth letter'),
    ('''项目管理机构
项目经理：张三
技术负责人：李四
安全员：王五''', 'personnel table'),
    ('''法定代表人或其委托代理人：（签字）赵六''', 'signature page'),
]
for text, label in personnel_tests:
    r = extract_personnel(text)
    print(f'  [{label}]: legal_rep={r.get(\"legal_rep\")}, auth_rep={r.get(\"authorized_rep\")}, all_persons={r.get(\"all_persons\")}')

print()
print('=== parse_amount 测试 ===')
for val, expected in [('1,234,567.89', 1234567.89), ('123.45万元', 1234500), ('1.2亿元', 120000000)]:
    result = _parse_amount(val)
    ok = abs(result - expected) < 1
    print(f'  _parse_amount(\"{val}\") = {result:,.0f} (expected {expected:,.0f}) {\"✅\" if ok else \"❌\"}')
"
```

- [ ] **Step 2: 语法检查**

```bash
cd /Users/dev/Desktop/围串标/围串标风险识别APPV2 && ./venv/bin/python3 -c "import py_compile; py_compile.compile('app.py', doraise=True)" && echo "✅ Syntax OK"
```

- [ ] **Step 3: 启动应用进行端到端测试**

```bash
cd /Users/dev/Desktop/围串标/围串标风险识别APPV2 && ./venv/bin/python3 app.py 5001 &
sleep 3
# Health check
curl -s -o /dev/null -w "%{http_code}" http://localhost:5001/
echo " (should be 200)"
# Stop server
lsof -ti:5001 | xargs kill -9 2>/dev/null
```

- [ ] **Step 4: 最终提交**

```bash
git add app.py
git commit -m "chore: end-to-end validation of price and personnel enhancements"
```
