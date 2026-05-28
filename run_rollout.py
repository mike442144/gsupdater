#!/usr/bin/env python3
"""One-off driver: roll out the three ROIC changes to the spreadsheets that the
safety audit (audit_roic_rollout.py) cleared — i.e. Summary max ref <= ROIC row,
so inserting rows after ROIC can't shift any Summary-referenced row.

Per spreadsheet, runs in order:
  1. add_roic_methods.py     relabel base ROIC + insert 资产法 & Greenblatt rows
  2. wrap_keystats_refs.py --item "Minority Interest"   wrap MI refs in N()
  3. wrap_keystats_refs.py --item "Effective Tax Rate %" wrap tax refs in N()

Excluded (handled separately): 地产开发, 金融, 租赁物业, 教育 — no usable ROIC
anchor / divergent Summary layout.
"""
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY = os.path.join(HERE, '.venv', 'bin', 'python')

SAFE = ['互联网', '传媒', '贸易', '纺织服装', '食品', '餐饮', '个人用品',
        '药', '建材', '家电', '汽车', '设备', '交通运输', 'SAAS']


def run(args):
    env = dict(os.environ, PATH='/tmp/py-shim:' + os.environ.get('PATH', ''))
    print(f"\n$ {' '.join(args)}", flush=True)
    p = subprocess.run([PY] + args, cwd=HERE, env=env,
                       capture_output=True, text=True)
    sys.stdout.write(p.stdout)
    if p.returncode != 0:
        sys.stdout.write("STDERR:\n" + p.stderr)
    sys.stdout.flush()
    return p.returncode


def main():
    inds = json.load(open(os.path.join(HERE, 'industry_spreadsheets.json')))
    names = sys.argv[1:] or SAFE
    for name in names:
        sid = inds[name]['spreadsheet_id']
        print(f"\n{'#'*70}\n# {name}  {sid}\n{'#'*70}", flush=True)
        run(['add_roic_methods.py', '--spreadsheet-id', sid])
        run(['wrap_keystats_refs.py', '--spreadsheet-id', sid, '--item', 'Minority Interest'])
        run(['wrap_keystats_refs.py', '--spreadsheet-id', sid, '--item', 'Effective Tax Rate %'])
    print("\nALL DONE", flush=True)


if __name__ == '__main__':
    main()
