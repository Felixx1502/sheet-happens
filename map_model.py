#!/usr/bin/env python3
"""sheet-happens — map any Excel financial model into an interactive HTML graph.

Usage:
    python map_model.py MODEL.xlsx [-o out.html] [--title "My model"]
                        [--blocks-override blocks.json] [--anonymize]
                        [--max-blocks 14] [--min-block-rows 3]

Output: a single self-contained HTML file (works offline, file://).
Everything runs locally; the model never leaves your machine.
"""
import argparse, os, sys, time


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('xlsx', help='path to .xlsx/.xlsm model')
    ap.add_argument('-o', '--out', default=None, help='output HTML path')
    ap.add_argument('--title', default=None, help='map title (default: file name)')
    ap.add_argument('--blocks-override', default=None,
                    help='JSON file pinning block boundaries per sheet')
    ap.add_argument('--anonymize', action='store_true',
                    help='strip labels/formulas, pseudonymise sheet names')
    ap.add_argument('--max-blocks', type=int, default=14)
    ap.add_argument('--min-block-rows', type=int, default=3)
    args = ap.parse_args()

    if not os.path.isfile(args.xlsx):
        sys.exit(f'error: {args.xlsx} not found')
    here = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, here)
    from xlmap.extract import extract
    from xlmap.blocks import load_blocks
    from xlmap.graph import build
    from xlmap.render import render, anonymize

    t0 = time.time()
    base = os.path.splitext(os.path.basename(args.xlsx))[0]
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.xlsx)),
                                   base + ' - Model Map.html')
    title = args.title or base

    print(f'[1/4] extracting {args.xlsx} …')
    wb = extract(args.xlsx)
    if not wb.sheets:
        sys.exit('error: no worksheets found')
    print(f'[2/4] detecting blocks …')
    blocks = load_blocks(wb, args.blocks_override, args.min_block_rows, args.max_blocks)
    nb = sum(len(v) for v in blocks.values())
    print(f'      {nb} blocks across {len(wb.sheets)} sheets')
    print(f'[3/4] building graph …')
    data = build(wb, blocks)
    a = data['audit']
    print(f"      {len(data['edges'])} block links · {len(data['redges'])} item links")
    print(f"      audit: {len(a['hardcodes'])} hardcode rows · {len(a['errors'])} error cells"
          f" · {len(a['cycles'])} circular groups"
          f" · names {a['names_resolved']}/{a['names_total']} resolved ({a['names_broken']} broken)")
    if a['dropped']:
        print('      not mapped: ' + '; '.join(f"{d['what']} x{d['n']} ({d['sheet']})"
                                               for d in a['dropped'][:5]))
    if args.anonymize:
        print('      anonymizing labels & formulas …')
        anonymize(data)
        title = 'Anonymized model'
        base = 'model'
    print(f'[4/4] rendering …')
    render(data, out, title, base + os.path.splitext(args.xlsx)[1],
           os.path.join(here, 'xlmap'))
    print(f'done in {time.time()-t0:.1f}s -> {out}')
    print('NOTE: the HTML embeds labels & formulas from your model.'
          ' Share it with the same care as the model itself'
          ' (use --anonymize for a shareable skeleton).')


if __name__ == '__main__':
    main()
