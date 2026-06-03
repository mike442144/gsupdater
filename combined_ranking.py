#!/usr/bin/env python3
"""
Combined ranking: merge EV/EBIT + ROIC + Payout Ratio into one master ranking.

Reads from:
  - rankings.csv        (from rank_companies.py: ev_rank, roic_rank, combined)
  - payout_rankings.csv (from rank_payout_ratio.py: rank as payout_rank)

Merges by (company, industry). The combined score is the sum of all three ranks
(lower = better). Companies missing any rank are excluded from the combined list
but still shown in per-dimension rankings.

Usage:
    python combined_ranking.py                                # default inputs
    python combined_ranking.py --rankings r.csv --payout p.csv
    python combined_ranking.py --output master.csv
"""

import csv
import os
import argparse


def load_rankings(path):
    """Load rankings.csv → dict keyed by (company, industry)."""
    data = {}
    with open(path, encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            key = (row['company'].strip(), row['industry'].strip())
            data[key] = {
                'company': row['company'].strip(),
                'industry': row['industry'].strip(),
                'code': row.get('code', '').strip(),
                'ev_ebit': _float(row.get('ev_ebit')),
                'ev_rank': _int(row.get('ev_rank')),
                'roic': _float(row.get('roic_corrected') or row.get('roic')),
                'roic_rank': _int(row.get('roic_rank')),
            }
    return data


def load_payout(path):
    """Load payout_rankings.csv → dict keyed by (company, industry)."""
    data = {}
    with open(path, encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            key = (row['company'].strip(), row['industry'].strip())
            data[key] = {
                'payout': _float(row.get('agg_payout')),
                'payout_rank': _int(row.get('rank')),
            }
    return data


def _float(v):
    try:
        return float(v) if v and v.strip() else None
    except (ValueError, AttributeError):
        return None


def _int(v):
    try:
        return int(float(v)) if v and v.strip() else None
    except (ValueError, AttributeError):
        return None


def main():
    here = os.path.dirname(__file__)
    parser = argparse.ArgumentParser(
        description='Combined EV/EBIT + ROIC + Payout Ratio ranking')
    parser.add_argument('--rankings', default=os.path.join(here, 'rankings.csv'),
                        help='Input: EV/EBIT + ROIC rankings CSV')
    parser.add_argument('--payout', default=os.path.join(here, 'payout_rankings.csv'),
                        help='Input: Payout Ratio rankings CSV')
    parser.add_argument('--output', '-o', default='master_ranking.csv',
                        help='Output CSV path')
    args = parser.parse_args()

    rankings = load_rankings(args.rankings)
    payout = load_payout(args.payout)

    # Merge
    all_keys = set(rankings.keys()) | set(payout.keys())
    merged = []
    for key in all_keys:
        r = rankings.get(key, {})
        p = payout.get(key, {})
        merged.append({
            'company': r.get('company') or key[0],
            'industry': r.get('industry') or key[1],
            'code': r.get('code', ''),
            'ev_ebit': r.get('ev_ebit'),
            'ev_rank': r.get('ev_rank'),
            'roic': r.get('roic'),
            'roic_rank': r.get('roic_rank'),
            'payout': p.get('payout'),
            'payout_rank': p.get('payout_rank'),
        })

    # Combined score: sum of all three ranks (lower = better)
    # Only include companies that have ALL three ranks
    combined = []
    for m in merged:
        if m['ev_rank'] and m['roic_rank'] and m['payout_rank']:
            m['combined'] = m['ev_rank'] + m['roic_rank'] + m['payout_rank']
            combined.append(m)
    combined.sort(key=lambda x: x['combined'])
    for i, c in enumerate(combined, 1):
        c['master_rank'] = i

    # EV/EBIT only ranking
    ev_only = [m for m in merged if m['ev_rank']]
    ev_only.sort(key=lambda x: x['ev_rank'])

    # ROIC only ranking
    roic_only = [m for m in merged if m['roic_rank']]
    roic_only.sort(key=lambda x: x['roic_rank'])

    # Payout only ranking
    payout_only = [m for m in merged if m['payout_rank']]
    payout_only.sort(key=lambda x: x['payout_rank'])

    # ── Write CSV ────────────────────────────────────────────────
    fieldnames = ['master_rank', 'industry', 'code', 'company',
                  'ev_ebit', 'ev_rank', 'roic', 'roic_rank',
                  'payout', 'payout_rank', 'combined']

    output_path = os.path.join(here, args.output)
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames,
                                extrasaction='ignore')
        writer.writeheader()
        for c in combined:
            row = dict(c)
            if row.get('ev_ebit') is not None:
                row['ev_ebit'] = round(row['ev_ebit'], 1)
            if row.get('roic') is not None:
                row['roic'] = round(row['roic'], 4)
            if row.get('payout') is not None:
                row['payout'] = round(row['payout'], 4)
            writer.writerow(row)

    # ── Print: Master combined ranking ───────────────────────────
    print(f"\n{'=' * 120}")
    print(f" Master Ranking (EV/EBIT rank + ROIC rank + Payout Ratio rank)")
    print(f"{'=' * 120}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'ΣRank':>6} {'EV/EBIT':>8} {'R_EV':>5} {'ROIC':>9} {'R_ROIC':>6} "
          f"{'Payout':>8} {'R_Pay':>6}  Code")
    print(f"{'-' * 120}")
    for c in combined:
        ev_s = f"{c['ev_ebit']:.1f}" if c['ev_ebit'] else '-'
        roic_s = f"{c['roic']:.1%}" if c['roic'] else '-'
        pay_s = f"{c['payout']:.1%}" if c['payout'] else '-'
        print(f"{c['master_rank']:>3}  {c['company']:<16} {c['industry']:<8} "
              f"{c['combined']:>6} {ev_s:>8} {c['ev_rank']:>5} "
              f"{roic_s:>9} {c['roic_rank']:>6} "
              f"{pay_s:>8} {c['payout_rank']:>6}  {c['code']}")

    # Stats
    in_all = len(combined)
    in_ev = len(ev_only)
    in_roic = len(roic_only)
    in_pay = len(payout_only)
    missing_pay = [m for m in merged if m['ev_rank'] and m['roic_rank'] and not m['payout_rank']]
    missing_ev = [m for m in merged if m['payout_rank'] and not m['ev_rank']]

    print(f"\nTotal: {len(all_keys)} companies | "
          f"EV/EBIT: {in_ev} | ROIC: {in_roic} | Payout: {in_pay} | "
          f"Combined (all 3): {in_all}")

    if missing_pay:
        print(f"\nCompanies with EV/EBIT+ROIC but missing Payout ({len(missing_pay)}):")
        for m in sorted(missing_pay, key=lambda x: x['ev_rank'] + x['roic_rank']):
            print(f"  {m['company']:<16} {m['industry']:<8} "
                  f"EV_rank={m['ev_rank']}, ROIC_rank={m['roic_rank']}")

    if missing_ev:
        print(f"\nCompanies with Payout but missing EV/EBIT ({len(missing_ev)}):")
        for m in sorted(missing_ev, key=lambda x: x['payout_rank']):
            print(f"  {m['company']:<16} {m['industry']:<8} "
                  f"Payout_rank={m['payout_rank']}")

    print(f"\nCSV: {output_path}")


if __name__ == '__main__':
    main()
