#!/usr/bin/env python3
"""Unit tests for personnel / price extraction generalization.

Run: ./venv/bin/python3 tests/test_extraction.py
No pytest dependency; plain asserts with a tiny runner.
"""
import os
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


def main():
    for name, fn in sorted(globals().items()):
        if name.startswith('t_') and callable(fn):
            check(name, fn)
    print(f'\n{len(PASS)}/{len(PASS)} tests passed')


if __name__ == '__main__':
    main()
