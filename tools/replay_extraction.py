#!/usr/bin/env python3
"""Offline extraction replay harness.

Walks a corpus directory, runs text extraction + personnel extraction +
price extraction on every supported file, and dumps a per-file JSON summary.
Used to compare extraction quality before/after changes:

    ./venv/bin/python3 tools/replay_extraction.py <corpus_dir> <out_json>

Each summary records:
  - text length extracted (0 => total extraction failure, e.g. scanned PDF)
  - personnel: legal_rep / authorized_rep / #all_persons / has phone
  - prices: totalPrice / totalPriceInTax / taxRate / #subItemPrice / #costDetails
"""
import os
import sys
import json
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app as appmod  # noqa: E402

SUPPORTED = ('.docx', '.doc', '.pdf', '.txt', '.xlsx')


def summarize_file(fp):
    rec = {'file': os.path.relpath(fp), 'error': None}
    try:
        text = appmod.extract_text_with_tables(fp, max_pages=300)
    except Exception as e:
        text = ''
        rec['error'] = f'text: {e}'
    rec['text_len'] = len(text or '')

    try:
        p = appmod.extract_personnel(text or '')
        rec['personnel'] = {
            'legal_rep': p.get('legal_rep'),
            'authorized_rep': p.get('authorized_rep'),
            'company_name': p.get('company_name'),
            'num_persons': len(p.get('all_persons', [])),
            'has_phone': bool(p.get('phone') or (p.get('contacts') or {}).get('phone')),
            'persons': sorted({x['name'] for x in p.get('all_persons', [])}),
        }
    except Exception as e:
        rec['personnel'] = {'error': str(e)}

    try:
        pr = appmod.extract_prices(text or '')
        rec['prices'] = {
            'totalPrice': pr.get('totalPrice'),
            'totalPriceInTax': pr.get('totalPriceInTax'),
            'taxRate': pr.get('taxRate'),
            'num_sub': len(pr.get('subItemPrice', [])),
            'num_cost': len(pr.get('costDetails', [])),
            'cost': pr.get('cost'),
            'revenue': pr.get('revenue'),
        }
    except Exception as e:
        rec['prices'] = {'error': str(e)}
    return rec


def main(corpus_dir, out_json):
    records = []
    for root, dirs, files in os.walk(corpus_dir):
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        for fn in sorted(files):
            if fn.startswith('.') or not fn.lower().endswith(SUPPORTED):
                continue
            fp = os.path.join(root, fn)
            try:
                records.append(summarize_file(fp))
            except Exception:
                records.append({'file': fp, 'error': traceback.format_exc()})

    # Aggregate metrics
    total = len(records)
    with_text = sum(1 for r in records if r.get('text_len', 0) > 100)
    with_person = sum(1 for r in records
                      if (r.get('personnel') or {}).get('num_persons', 0) > 0)
    with_rep = sum(1 for r in records
                   if (r.get('personnel') or {}).get('legal_rep')
                   or (r.get('personnel') or {}).get('authorized_rep'))
    with_price = sum(1 for r in records
                     if (r.get('prices') or {}).get('totalPriceInTax')
                     or (r.get('prices') or {}).get('totalPrice'))
    with_sub = sum(1 for r in records
                   if (r.get('prices') or {}).get('num_sub', 0) > 0)

    metrics = {
        'corpus': corpus_dir,
        'total_files': total,
        'files_with_text': with_text,
        'files_with_persons': with_person,
        'files_with_rep': with_rep,
        'files_with_price': with_price,
        'files_with_subitems': with_sub,
    }
    with open(out_json, 'w', encoding='utf-8') as f:
        json.dump({'metrics': metrics, 'records': records}, f,
                  ensure_ascii=False, indent=1)
    print(json.dumps(metrics, ensure_ascii=False, indent=2))
    print(f'details -> {out_json}')


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
