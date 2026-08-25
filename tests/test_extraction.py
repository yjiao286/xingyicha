#!/usr/bin/env python3
"""Unit tests for personnel / price extraction generalization.

Run: ./venv/bin/python3 tests/test_extraction.py
No pytest dependency; plain asserts with a tiny runner.
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as m  # noqa: E402

PASS = []


def check(label, fn):
    try:
        fn()
    except AssertionError as e:
        print(f'FAIL {label}: {e}')
        raise
    PASS.append(label)
    print(f'PASS {label}')


# ── Price extraction channels ──
def t_price_standalone_cn():
    r = m.extract_prices('人民币（大写）：壹亿贰仟陆佰壹拾捌万壹仟玖佰柒拾陆元叁角')
    assert r['totalPriceInTax'] is not None and abs(r['totalPriceInTax'] - 126181976.30) < 1, r


def t_price_symbol_with_wan():
    r = m.extract_prices('投标总价：￥12.5万元')
    assert abs(r['totalPriceInTax'] - 125000) < 1, r


def t_price_cn_label():
    r = m.extract_prices('投标总价：壹佰贰拾叁万元整')
    assert abs(r['totalPriceInTax'] - 1230000) < 1, r


def t_price_fullwidth():
    r = m.extract_prices('小写：１２３４５６')
    assert r['totalPriceInTax'] == 123456, r


def t_price_bond_excluded():
    r = m.extract_prices('投标保证金一份，金额为人民币：￥500000元')
    assert r['totalPriceInTax'] is None, r


def t_price_summary_sep():
    r = m.extract_prices('合计 \\ 1838529 5002800 \\')
    assert r['totalPrice'] == 1838529 and r['totalPriceInTax'] == 5002800, r


def t_price_bidrate():
    assert m.extract_prices('下浮率：12.5%')['bidRate'] == '12.5%'
    assert m.extract_prices('报价费率 1.8‰')['bidRate'] == '1.8‰'
    assert m.extract_prices('投标报价下浮6%')['bidRate'] == '6%'


def t_price_smallcase_first():
    r = m.extract_prices('小写：123456元 大写：壹拾贰万叁仟肆佰伍拾陆元整')
    assert r['totalPriceInTax'] == 123456, r


def t_price_cn_fractions():
    r = m.extract_prices('人民币（大写）：贰佰叁拾肆万伍仟陆佰柒拾捌元玖角')
    assert abs(r['totalPriceInTax'] - 2345678.9) < 0.01, r


def t_price_reverse_pair():
    r = m.extract_prices('小写：123456\n大写：壹拾贰万叁仟肆佰伍拾陆元整')
    assert r['totalPriceInTax'] == 123456, r


def t_similarity_toc_dots_ignored():
    # TOC leader dots ('四、授权委托书 ....... 7') must never pair unrelated
    # lines across documents ('9.7 网络拥塞的感知时间不高于 5s ....... 79')
    t1 = '目录\n四、授权委托书 ....... 7\n一、投标函 ......... 3\n'
    t2 = '9.7 网络拥塞的感知时间不高于 5s ....... 79\n项目人员与分工表\n'
    res = m.text_similarity_analysis({'a': t1, 'b': t2})
    for pr in res['pair_results']:
        for mm in pr.get('matches', []):
            assert '授权委托书' not in mm.get('text', ''), mm
            assert '网络拥塞' not in mm.get('text', ''), mm
    # 真实内容仍应匹配
    t3 = '本项目采用三层架构设计，核心交换节点采用双机热备冗余部署策略\n'
    t4 = '本项目采用三层架构设计，核心交换节点采用双机热备冗余部署策略\n'
    res2 = m.text_similarity_analysis({'a': t3, 'b': t4})
    assert any('三层架构' in mm.get('text', '')
               for pr in res2['pair_results'] for mm in pr.get('matches', [])), res2


def t_price_wan_near_context():
    # 万-form value must survive the near-context validation
    r = m.extract_prices('投标总价：￥12.5万元')
    assert abs(r['totalPriceInTax'] - 125000) < 1, r


def t_price_reference_amount_excluded():
    # 合同金额 reference amounts must never win over the real bid price
    t = ('开标一览表\n投标报价\n总价：7590000元\n\n业绩：\n'
         '项目名称：某某市计量质量检测研究院\n数量：1\n合同金额：RMB2080000.00\n')
    r = m.extract_prices(t)
    assert r['totalPriceInTax'] == 7590000, r
    # 招标文件售价 (document price, not bid price)
    r2 = m.extract_prices('开标一览表\n投标报价表\n\n招标文件售价：人民币1000元\n')
    assert r2['totalPriceInTax'] is None, r2


# ── Personnel extraction ──
def t_personnel_pipe_table():
    text = ('姓名 | 职务 | 电话\n'
            '张三 | 项目经理 | 13900000001\n'
            '李四 | 技术负责人 | 13900000002\n')
    p = m.extract_personnel(text)
    names = {x['name'] for x in p['all_persons']}
    assert {'张三', '李四'} <= names, names
    assert '13900000001' in p['phones'], p['phones']
    assert '13900000002' in p['phones'], p['phones']


def t_personnel_pipe_id():
    text = ('姓名 | 职务 | 身份证号\n王五 | 安全员 | 320101199001011234\n')
    p = m.extract_personnel(text)
    assert '320101199001011234' in p['id_numbers'], p['id_numbers']
    assert p['id_number'] == '320101199001011234', p['id_number']


def t_personnel_multivalue():
    p = m.extract_personnel('身份证号：320101199001011234 电话：13912345678 邮箱：a@b.com')
    assert '13912345678' in p['phones'] and '320101199001011234' in p['id_numbers']
    assert 'a@b.com' in p['emails'], p['emails']


def t_personnel_spaced_id():
    p = m.extract_personnel('身份证号：320101 1990 01 01 1234')
    assert p['id_numbers'] and p['id_numbers'][0] == '320101199001011234', p['id_numbers']


def t_personnel_auth_variants():
    p = m.extract_personnel('授权委托代理人：赵六\n联系电话：13800000000\n')
    assert p['authorized_rep'] == '赵六', p
    assert '13800000000' in p['phones'], p['phones']


def t_personnel_blank_template_no_leak():
    # Blank authorization-letter templates ('本人 （姓名）系 （投标人名称）…')
    # must not leak 本人/性别/盖单位章 as names or companies
    t = ('四、授权委托书\n本人 （姓名）系 （投标人名称）的法定代表人（单位负责人），现委托\n'
         '投标人：________________（盖单位章）\n姓名：性别：男 身份证号：\n')
    p = m.extract_personnel(t)
    names = {x['name'] for x in p['all_persons']}
    assert '本人' not in names and '性别' not in names, names
    assert p.get('legal_rep') not in ('本人', '性别', '性别:'), p.get('legal_rep')
    assert p.get('company_name') not in ('________________（盖单位章', '（盖单位章'), p.get('company_name')


def t_personnel_title_word_not_name():
    # PDF "序号 姓名 职称 分工" tables: '张然 中级 项目负责人' — the 职称
    # column value 中级 must not become the name; 张然 must be captured.
    t = ('项目人员配置\n表 5 项目人员与分工\n序号 姓名 职称 分工\n'
         '1 张然 中级 项目负责人\n'
         '2 刘某某 教授 流资源预留协议设计\n'
         '3 潘某某 副教授 负载均衡技术设计\n')
    p = m.extract_personnel(t)
    names = [x['name'] for x in p['all_persons']]
    assert '中级' not in names and '教授' not in names and '副教授' not in names, names
    assert '张然' in names, names
    zhang = [x for x in p['all_persons'] if x['name'] == '张然']
    assert any(x['role'] == 'project_manager' for x in zhang), zhang


def t_personnel_role_not_name():
    # Role keywords in the name column must never become person entries
    p = m.extract_personnel('姓名 | 职务\n张三 | 项目经理\n项目经理 | 组长\n')
    names = {x['name'] for x in p['all_persons']}
    assert names == {'张三'}, names
    roles = {x['role'] for x in p['all_persons']}
    assert 'project_manager' in roles, roles


# ── xlsx ──
def t_xlsx_extract(tmp='/tmp/_t.xlsx'):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '报价单'
    ws.append(['序号', '名称', '数量', '单价', '总价'])
    ws.append([1, '服务器', 2, 30000, 60000])
    wb.save(tmp)
    try:
        text = m.extract_text_with_tables(tmp)
        assert '报价单' in text and '服务器' in text, text
        assert ' | ' in text
    finally:
        os.remove(tmp)


# ── Cross-match layer helpers (set-intersection pools) ──
def t_pools_intersection():
    from app import extract_personnel
    pi = extract_personnel('电话：13900000001 身份证号：320101199001011234')
    pj = extract_personnel('手机：13900000001 身份证号：320101199001011234')
    shared_phones = set(pi['phones']) & set(pj['phones'])
    shared_ids = set(pi['id_numbers']) & set(pj['id_numbers'])
    assert shared_phones == {'13900000001'}, shared_phones
    assert shared_ids == {'320101199001011234'}, shared_ids


# ── Format-generalization enhancements ──
def t_price_label_unit_paren():
    # '总价（元）：' and '合计（万元）：' label-with-unit formats
    r = m.extract_prices('开标一览表\n投标总价（元）：1230000元')
    assert r['totalPriceInTax'] == 1230000, r
    r2 = m.extract_prices('报价汇总表\n合计（万元）：89.3万元')
    assert abs(r2['totalPriceInTax'] - 893000) < 1, r2


def t_price_pipe_total_wan_cells():
    # 合计 pipe row with 万元-suffixed cells must keep magnitude
    r = m.extract_prices('序号 | 名称 | 金额\n合计 | 全部 | 89.3万元 | 91.97万元')
    assert abs(r['totalPrice'] - 893000) < 1, r
    assert abs(r['totalPriceInTax'] - 919700) < 1, r


def t_price_bidrate_extra():
    assert m.extract_prices('投标报价（%）：98.5')['bidRate'] == '98.5%', m.extract_prices('投标报价（%）：98.5')
    assert m.extract_prices('投标报价下浮 6 个百分点')['bidRate'] == '6%'


def t_personnel_zhweituo():
    p = m.extract_personnel('授权委托书\n兹委托 李某某 同志为我方代理人，负责签署投标文件。')
    assert p['authorized_rep'] == '李某某', p


def t_personnel_reversed_labels():
    p = m.extract_personnel('项目管理机构\n职务：项目经理 姓名：王某某 联系电话：13800000000')
    names = {x['name'] for x in p['all_persons']}
    assert '王某某' in names, names


def t_personnel_pipe_merged_cell():
    # docx vertically-merged name cells: empty name inherits row above
    text = ('姓名 | 职务 | 联系电话\n'
            '张三 | 项目经理 | 13900000001\n'
            ' | 技术负责人 | 13900000002\n')
    p = m.extract_personnel(text)
    names = {x['name'] for x in p['all_persons']}
    assert {'张三'} <= names, names
    assert '13900000002' in p['phones'], p['phones']


def t_personnel_pipe_multi_phone_cell():
    text = '姓名 | 职务 | 联系电话\n张三 | 项目经理 | 13900000001/13900000002\n'
    p = m.extract_personnel(text)
    assert '13900000001' in p['phones'] and '13900000002' in p['phones'], p['phones']


def t_personnel_pipe_header_alias():
    text = ('拟投入主要人员 | 职务\n'
            '李四 | 技术负责人\n')
    p = m.extract_personnel(text)
    names = {x['name'] for x in p['all_persons']}
    assert '李四' in names, names


def t_personnel_minority_dot_normalized():
    # Same minority name spelled with different middle dots must cross-match
    a = m.extract_personnel('项目管理机构\n姓名：阿不来提•买买提')
    b = m.extract_personnel('项目管理机构\n姓名：阿不来提·买买提')
    ka = {re.sub(r'[•・]', '·', x['name']) for x in a['all_persons']}
    kb = {re.sub(r'[•・]', '·', x['name']) for x in b['all_persons']}
    assert ka & kb, (ka, kb)


def t_personnel_bank_account_pool():
    p = m.extract_personnel('开户银行：中国工商银行北京分行\n账号：1100923456789000123')
    assert '1100923456789000123' in p.get('bank_accounts', []), p.get('bank_accounts')


def t_personnel_section_marker_extended():
    # New personnel-table section markers must scope extraction
    p = m.extract_personnel('人员一览表\n姓名 职务\n王某某 项目经理\n')
    names = {x['name'] for x in p['all_persons']}
    assert '王某某' in names, names


# ── Accuracy hardening: ceilings, environmental noise, cross-checks ──
def t_price_ceiling_excluded():
    # 最高限价 reprinted in every bid must never win as the bid price,
    # and must not HIDE the real price that appears further down.
    t = ('开标一览表\n最高限价（招标控制价）：￥1,000,000元\n投标报价：￥980,000元\n')
    r = m.extract_prices(t)
    assert r['totalPriceInTax'] == 980000, r


def t_price_ceiling_only_not_captured():
    r = m.extract_prices('开标一览表\n招标控制价：￥1000000元\n')
    assert r['totalPriceInTax'] is None, r


def t_bank_account_tender_side_excluded():
    # 保证金汇入账号 appears in EVERY bid → must stay out of the pool
    t = ('投标保证金请汇入以下账户\n开户银行：XX银行北京分行\n账号：1100923456789000123\n')
    p = m.extract_personnel(t)
    assert '1100923456789000123' not in p.get('bank_accounts', []), p.get('bank_accounts')


def t_environmental_pool_demotion():
    pools = {
        'A': {'phones': ['13900000001', '13800000000'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
        'B': {'phones': ['13900000002', '13800000000'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
        'C': {'phones': ['13900000003', '13800000000'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
        'D': {'phones': ['13900000004', '13800000000'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
        'E': {'phones': ['13800000001'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
    }
    removed = m._demote_environmental_pool_values(pools, list(pools))
    # 13800000000 in 4/5 groups (80%) → environmental, demoted
    assert '13800000000' in removed, removed
    assert pools['A']['phones'] == ['13900000001'], pools['A']
    assert pools['E']['phones'] == ['13800000001'], pools['E']
    # 3 of 4 (75%) is below the ratio — a genuinely shared phone between
    # 3 of 4 bidders is still collusion evidence and must NOT be demoted
    pools3 = {
        'A': {'phones': ['13900000001'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
        'B': {'phones': ['13900000001'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
        'C': {'phones': ['13900000001'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
        'D': {'phones': ['13900000009'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
    }
    assert not m._demote_environmental_pool_values(pools3, list(pools3))
    assert pools3['A']['phones'] == ['13900000001']
    # With 2 groups only, a shared phone stays (still real evidence)
    pools2 = {
        'A': {'phones': ['13900000001'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
        'B': {'phones': ['13900000001'], 'id_numbers': [], 'emails': [], 'bank_accounts': []},
    }
    assert not m._demote_environmental_pool_values(pools2, list(pools2))
    assert pools2['A']['phones'] == ['13900000001']


def t_price_cn_arabic_mismatch():
    # 大写 says 100万 but 小写 says 980000 → adopt 小写, note the conflict
    t = '开标一览表\n投标总价：壹佰万元整（小写：980000元）\n'
    r = m.extract_prices(t)
    assert r['totalPriceInTax'] == 980000, r
    assert any('大写' in w for w in r['warnings']), r['warnings']


def t_price_cn_arabic_consistent_no_warning():
    t = '开标一览表\n投标总价：壹佰万元整（小写：1000000元）\n'
    r = m.extract_prices(t)
    assert not any('大写' in w for w in r['warnings']), r['warnings']


def t_price_subitem_sum_warning():
    t = ('报价明细表\n序号 | 名称 | 数量 | 单价 | 总价\n'
         '1 | 服务器 | 2 | 30000 | 60000\n'
         '2 | 交换机 | 1 | 20000 | 20000\n'
         '合计 | | | | 980000\n')
    r = m.extract_prices(t)
    assert any('分项合计' in w for w in r['warnings']), r['warnings']


def t_price_global_fallback_warning():
    r = m.extract_prices('合计 \\ 1838529 5002800 \\')
    assert any('兜底' in w for w in r['warnings']), r['warnings']


def t_phone_spaced_and_ocr():
    p = m.extract_personnel('联系电话：139 1234 5678')
    assert '13912345678' in p['phones'], p['phones']
    p2 = m.extract_personnel('手机：138O1234567')
    assert '13801234567' in p2['phones'], p2['phones']


def t_surname_filter():
    # Non-surname table fragments must be rejected; rare/compound surnames kept
    assert not m._is_person_name('情况'), '情况 should be rejected'
    assert not m._is_person_name('规程'), '规程 should be rejected'
    assert m._is_person_name('欧阳建国')
    assert m._is_person_name('覃芳')
    assert m._is_person_name('诸葛云')


def t_parse_amount_ocr_letters():
    assert m._parse_amount('1O9,800') == 109800
    assert m._parse_amount('l2300') == 12300
    assert m._parse_amount('I00000') == 100000


def t_fitz_table_pipes():
    # PyMuPDF table rows → pipe lines (fake page object, no pymupdf needed)
    class FakeTable:
        def extract(self):
            return [['姓名', '职务'], ['张三', '项目经理'], [None, '']]
    class FakeFinder:
        tables = [FakeTable()]
    class FakePage:
        def find_tables(self):
            return FakeFinder()
    out = m._fitz_page_tables_as_pipes(FakePage())
    assert '姓名 | 职务' in out and '张三 | 项目经理' in out, out


# ── Real-world regression: 某某航天 format (2026-08-25) ──
def t_personnel_auth_letter_full():
    # 授权委托书 with labeled parens; also exercises the 兰 surname
    t = ('授权委托书\n本人  兰某某  （姓名）系  北京某某航天技术有限公司  （供应商名称）的'
         '法定代表人（单位负责人），现委托  陈某某  （姓名）为我方授权代理。\n')
    p = m.extract_personnel(t)
    assert p['legal_rep'] == '兰某某', p
    assert p['company_name'] == '北京某某航天技术有限公司', p
    assert p['authorized_rep'] == '陈某某', p


def t_personnel_cert_paren_format():
    # 法定代表人资格证明书: name AND company in unlabeled parentheses
    t = '法定代表人资格证明书\n（兰某某）系（北京某某航天技术有限公司）的法定代表人。\n特此证明\n'
    p = m.extract_personnel(t)
    assert p['legal_rep'] == '兰某某', p
    assert p['company_name'] == '北京某某航天技术有限公司', p


def t_personnel_company_survives_bad_name():
    # A rejected name must not also lose the company
    t = ('授权委托书\n本人  某某  （姓名）系  北京某某航天技术有限公司  （供应商名称）的法定代表人\n')
    p = m.extract_personnel(t)
    assert p['legal_rep'] is None, p
    assert p['company_name'] == '北京某某航天技术有限公司', p


def t_personnel_surnames_extended():
    for n in ('兰某某', '付强', '肖磊', '闫丽', '郝建', '滕飞', '岳鹏', '单芳'):
        assert m._is_person_name(n), n


# ── History persistence: _prepare_history_data keep/strip semantics ──
def _fake_results():
    return {
        '_flags': {'group_map': {}, 'ok': True},
        'verdict': {'conclusion': '两份标书存在围串标高度嫌疑'},
        'metadata': {
            'files': [{'name': 'A.docx', 'creator': '张三',
                       'KSOProductBuildVer': '12.1.0.1', 'ICV': 'abc-def',
                       'KSOTemplateDocerSaveRecord': 'r1', 'pages': 12,
                       '_error': None, 'secret_huge': 'x' * 5000}],
            'matches': [], 'findings': [],
        },
        'personnel': {
            'files': [{'name': 'A.docx', 'legal_rep': '王五',
                       'phones': ['13800000000', '13911111111'],
                       'id_numbers': ['110101199001011234'],
                       'id_number': '110101199001011234',
                       'emails': ['a@b.com'], 'bank_accounts': ['110060123456'],
                       'all_persons': [{'name': '王五', 'role': 'legal_rep',
                                        'confidence': 0.9}],
                       'contacts': {'phone': '13800000000', 'email': None,
                                    'address': None},
                       'heavy_blob': ['x' * 3000]}],
            'cross_matches': [], 'findings': [],
        },
        'pricing': {
            'files': [{'name': 'A.docx', 'totalPriceInTax': 100,
                       'totalPrice': 92, 'taxRate': '13%', 'bidRate': '99.5%',
                       'revenue': 10, 'cost': 5,
                       'warnings': ['大写/小写不一致已修正'],
                       'subItemPrice': [{'priceName': '设备', 'totalPrice': 92}]}],
            'comparison': {},
            'subItemCompare': [
                {'name': '设备',
                 'items': [{'file': 'A.docx', 'totalPriceInTax': 100,
                            'extras': {'厂家/型号': '华为S5720'}}],
                 'findings': []}],
            'findings': [],
        },
        'text_similarity': {
            'pair_results': [{
                'file1': 'A', 'file2': 'B', 'total_matches': 1,
                'matches': [{'index': 1, 'length': 500, 'text': 'x' * 300,
                             'ctx1': 'y' * 500, 'ctx2': 'z' * 500,
                             'abnormal': True, 'risk_level': 'substantial',
                             'score': 0.8, 'reasons': ['术语密度高']}],
            }],
            'template_matches': 0, 'total_pairs': 1, 'findings': [],
        },
    }


def t_history_prep_keeps_new_fields():
    r = m._prepare_history_data(_fake_results())
    pf = r['personnel']['files'][0]
    assert pf['phones'] == ['13800000000', '13911111111'], pf
    assert pf['id_numbers'] == ['110101199001011234'], pf
    assert pf['emails'] == ['a@b.com'], pf
    assert pf['bank_accounts'] == ['110060123456'], pf
    assert pf['all_persons'][0]['name'] == '王五', pf
    qf = r['pricing']['files'][0]
    assert qf['bidRate'] == '99.5%', qf
    assert qf['revenue'] == 10 and qf['cost'] == 5, qf
    assert qf['warnings'] == ['大写/小写不一致已修正'], qf
    mf = r['metadata']['files'][0]
    assert mf['KSOProductBuildVer'] == '12.1.0.1', mf
    assert mf['KSOTemplateDocerSaveRecord'] == 'r1', mf
    assert mf['ICV'] == 'abc-def', mf


def t_history_prep_keeps_extras():
    r = m._prepare_history_data(_fake_results())
    items = r['pricing']['subItemCompare'][0]['items']
    assert items[0]['extras'] == {'厂家/型号': '华为S5720'}, items


def t_history_prep_light_matches():
    r = m._prepare_history_data(_fake_results())
    match = r['text_similarity']['pair_results'][0]['matches'][0]
    assert len(match['text']) == 200, len(match['text'])
    assert len(match['ctx1']) == 400 and len(match['ctx2']) == 400
    assert match['index'] == 1 and match['score'] == 0.8
    assert match['reasons'] == ['术语密度高']


def t_history_prep_strips_unknown_and_keeps_underscore():
    r = m._prepare_history_data(_fake_results())
    pf = r['personnel']['files'][0]
    # Unused by the frontend/report: contacts is a dict -> None
    assert pf['contacts'] is None, pf['contacts']
    assert pf['heavy_blob'] is None, pf['heavy_blob']
    # '_'-prefixed keys survive; unknown str keys become '' (type preserved)
    assert r['_flags'] == {'group_map': {}, 'ok': True}
    mf = r['metadata']['files'][0]
    assert mf['_error'] is None
    assert mf['secret_huge'] == '', repr(mf['secret_huge'])


def t_history_prep_modifies_copy_only():
    src = _fake_results()
    r = m._prepare_history_data(src)
    assert src['personnel']['files'][0]['phones'] == ['13800000000', '13911111111']
    assert src['text_similarity']['pair_results'][0]['matches'][0]['text'] == 'x' * 300
    assert src['pricing']['subItemCompare'][0]['items'][0]['extras'] is not None


# ── Company-name / authorized-rep cleanliness regressions ──
def t_company_strips_detached_label():
    # '（投标人名称' had the closing paren consumed by a wider capture — the
    # detached fragment must be cleaned away, not left on the company name.
    assert m._clean_company('北京某某大学 （投标人名称') == '北京某某大学'
    assert m._clean_company('北京某某大学（盖单位章）') == '北京某某大学'
    assert m._clean_company('北京某某大学（盖单位章') == '北京某某大学'
    assert m._clean_company('___北京某某大学___（盖单位章）') == '北京某某大学'
    assert m._clean_company('北京某某航天技术有限公司') == '北京某某航天技术有限公司'


def t_auth_xianweituo_split_newline():
    # '现委托刘某某为我方代理\n人' — 代理人 split across a PDF line wrap must
    # still be recognized (pattern 8 tolerates the newline).
    info = {'legal_rep': None, 'authorized_rep': None, 'company_name': None, 'all_persons': []}
    m._extract_from_auth_section(
        '本人王某某系北京某某大学的法定代表人（单位负责人），现委托刘某某为我方代理\n人。'
        '代理人根据授权，以我方名义签署、澄明确认。', info)
    assert info['authorized_rep'] == '刘某某', info
    assert info['legal_rep'] == '王某某', info
    assert info['company_name'] == '北京某某大学', info


def t_history_prep_keeps_name():
    r = m._prepare_history_data(_fake_results())
    assert r['personnel']['files'][0]['name'] == 'A.docx', r['personnel']['files'][0]['name']
    assert r['pricing']['files'][0]['name'] == 'A.docx'
    assert r['metadata']['files'][0]['name'] == 'A.docx'


def _glue(text):
    return m._glue_phrases(text)


def t_glue_phrases_multi_position():
    # 法定\n代\n表\n人 split across several line breaks -> one keyword.
    assert _glue('本人王某某系北京交\n通大学的法定\n代\n表\n人（单位负责人）') == \
        '本人王某某系北京交\n通大学的法定代表人（单位负责人）'
    # 委托代\n理人
    assert _glue('现委托刘某某为我方委托代\n理人。代理人行使签署权。') == \
        '现委托刘某某为我方委托代理人。代理人行使签署权。'
    # absent keyword is untouched
    assert _glue('这是一段普通文字，没有关键词') == '这是一段普通文字，没有关键词'


def t_glue_cleans_auth_extraction():
    # A whole-keyword word wrap inside the auth letter must not lose the agent.
    info = {'legal_rep': None, 'authorized_rep': None, 'company_name': None, 'all_persons': []}
    m._extract_from_auth_section(
        '本人王某某系北京某某大学的法定代表人（单位负责人），现委托刘某某为我方代理\n人。', info)
    assert info['authorized_rep'] == '刘某某', info
    assert info['company_name'] == '北京某某大学', info


def t_cjk_ws_normalized():
    assert m._normalize_cjk_whitespace('投标人 ： 张三（ 盖单位章 ）') == '投标人： 张三（盖单位章）'
    assert m._normalize_cjk_whitespace('电话 ：　010-51683081') == '电话： 010-51683081'


def t_join_split_names():
    # 换行拆词拼接
    assert m._join_split_names('本人王\n稼琼系的法人') == '本人王某某系的法人'
    # 空格列间距不得拼接（否则 '国 联系' / '建国 联系' 吞掉标签）
    assert m._join_split_names('王 稼琼') == '王 稼琼'
    assert m._join_split_names('职务：项目经理 姓名：王某某 联系电话：13800000000') == \
        '职务：项目经理 姓名：王某某 联系电话：13800000000'


def t_glue_new_phrases():
    assert m._glue_phrases('投标\n文件：开标一览表') == '投标文件：开标一览表'
    assert m._glue_phrases('授权\n委托\n书') == '授权委托书'


def t_company_prefix_strip():
    assert m._clean_company('投标人：北京某某大学') == '北京某某大学'
    assert m._clean_company('单位名称：北京某某大学') == '北京某某大学'
    assert m._clean_company('企业名称　某省未来网络创新研究院') == '某省未来网络创新研究院'


def t_clean_phone_junk():
    assert m._clean_phone('_021-12345678') == '021-12345678'
    assert m._clean_phone('010-51683081；') == '010-51683081'


def main():
    for name, fn in sorted(globals().items()):
        if name.startswith('t_') and callable(fn):
            check(name, fn)
    print(f'\n{len(PASS)}/{len(PASS)} tests passed')


if __name__ == '__main__':
    main()
