#!/usr/bin/env python3
"""
Combined ranking: merge EV/EBIT + ROIC + Profit Quality + FCF Ratio + Capex Ratio + Payout Ratio
into one master ranking.

Reads from:
  - rankings.csv        (from gs_rankings.py: ev_rank, roic_rank, quality_rank, fcf_rank, capex_rank)
  - payout_rankings.csv (from gs_rankings.py: rank as payout_rank)

Merges by (company, industry). The combined score is the sum of selected ranks
(lower = better). Use --preset to choose which dimensions to include.

Usage:
    python combined_ranking.py                                # default inputs, all 6 dims
    python combined_ranking.py --rankings r.csv --payout p.csv
    python combined_ranking.py --preset classic               # original 4 dims
    python combined_ranking.py --preset fcf4                  # EV + ROIC + FCF + Payout
    python combined_ranking.py --output master.csv
"""

import csv
import os
import argparse


PRESETS = {
    'full': {
        'dims': ['ev_rank', 'roic_rank', 'quality_rank', 'fcf_rank', 'capex_rank', 'payout_rank'],
        'label': 'EV/EBIT + ROIC + Profit Quality + FCF Ratio + Capex Ratio + Payout',
    },
    'classic': {
        'dims': ['ev_rank', 'roic_rank', 'quality_rank', 'payout_rank'],
        'label': 'EV/EBIT + ROIC + Profit Quality + Payout',
    },
    'fcf4': {
        'dims': ['ev_rank', 'roic_rank', 'fcf_rank', 'payout_rank'],
        'label': 'EV/EBIT + ROIC + FCF Ratio + Payout',
    },
    'fcf5': {
        'dims': ['ev_rank', 'roic_rank', 'fcf_rank', 'capex_rank', 'payout_rank'],
        'label': 'EV/EBIT + ROIC + FCF Ratio + Capex Ratio + Payout',
    },
}


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
                'profit_quality': _float(row.get('profit_quality')),
                'quality_rank': _int(row.get('quality_rank')),
                'fcf_ratio': _float(row.get('fcf_ratio')),
                'fcf_rank': _int(row.get('fcf_rank')),
                'capex_ratio': _float(row.get('capex_ratio')),
                'capex_rank': _int(row.get('capex_rank')),
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
        description='Combined EV/EBIT + ROIC + Profit Quality + FCF Ratio + Capex Ratio + Payout ranking')
    parser.add_argument('--rankings', default=os.path.join(here, 'rankings.csv'),
                        help='Input: EV/EBIT + ROIC + Quality + FCF + Capex rankings CSV')
    parser.add_argument('--payout', default=os.path.join(here, 'payout_rankings.csv'),
                        help='Input: Payout Ratio rankings CSV')
    parser.add_argument('--preset', default='full', choices=PRESETS.keys(),
                        help='Ranking preset')
    parser.add_argument('--output', '-o', default='master_ranking.csv',
                        help='Output CSV path')
    args = parser.parse_args()

    preset = PRESETS[args.preset]
    dim_fields = preset['dims']

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
            'profit_quality': r.get('profit_quality'),
            'quality_rank': r.get('quality_rank'),
            'fcf_ratio': r.get('fcf_ratio'),
            'fcf_rank': r.get('fcf_rank'),
            'capex_ratio': r.get('capex_ratio'),
            'capex_rank': r.get('capex_rank'),
            'payout': p.get('payout'),
            'payout_rank': p.get('payout_rank'),
        })

    # Combined score: sum of preset dimension ranks (lower = better)
    # Only include companies that have ALL required ranks
    combined = []
    for m in merged:
        if all(m.get(d) for d in dim_fields):
            m['combined'] = sum(m[d] for d in dim_fields)
            combined.append(m)
    combined.sort(key=lambda x: x['combined'])
    for i, c in enumerate(combined, 1):
        c['master_rank'] = i

    # Per-dimension only rankings
    ev_only = [m for m in merged if m['ev_rank']]
    ev_only.sort(key=lambda x: x['ev_rank'])

    roic_only = [m for m in merged if m['roic_rank']]
    roic_only.sort(key=lambda x: x['roic_rank'])

    payout_only = [m for m in merged if m['payout_rank']]
    payout_only.sort(key=lambda x: x['payout_rank'])

    quality_only = [m for m in merged if m['quality_rank']]
    quality_only.sort(key=lambda x: x['quality_rank'])

    fcf_only = [m for m in merged if m['fcf_rank']]
    fcf_only.sort(key=lambda x: x['fcf_rank'])

    capex_only = [m for m in merged if m['capex_rank']]
    capex_only.sort(key=lambda x: x['capex_rank'])

    # ── Write CSV ────────────────────────────────────────────────
    fieldnames = ['master_rank', 'industry', 'code', 'company',
                  'ev_ebit', 'ev_rank', 'roic', 'roic_rank',
                  'profit_quality', 'quality_rank',
                  'fcf_ratio', 'fcf_rank',
                  'capex_ratio', 'capex_rank',
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
            if row.get('profit_quality') is not None:
                row['profit_quality'] = round(row['profit_quality'], 4)
            if row.get('fcf_ratio') is not None:
                row['fcf_ratio'] = round(row['fcf_ratio'], 4)
            if row.get('capex_ratio') is not None:
                row['capex_ratio'] = round(row['capex_ratio'], 4)
            if row.get('payout') is not None:
                row['payout'] = round(row['payout'], 4)
            writer.writerow(row)

    # ── Print: Master combined ranking ───────────────────────────
    n_dims = len(dim_fields)
    print(f"\n{'=' * 160}")
    print(f" Master Ranking [{args.preset}] ({preset['label']})")
    print(f"{'=' * 160}")
    print(f"{'#':>3}  {'Company':<16} {'Industry':<8} "
          f"{'ΣRank':>6} {'EV/EBIT':>8} {'R_EV':>5} {'ROIC':>9} {'R_ROIC':>6} "
          f"{'Qual%':>7} {'R_Qual':>6} {'FCF%':>7} {'R_FCF':>5} "
          f"{'Capex%':>7} {'R_Cpx':>5} "
          f"{'Payout':>8} {'R_Pay':>6}  Code")
    print(f"{'-' * 160}")
    for c in combined:
        ev_s = f"{c['ev_ebit']:.1f}" if c['ev_ebit'] else '-'
        roic_s = f"{c['roic']:.1%}" if c['roic'] else '-'
        qual_s = f"{c['profit_quality']:.1%}" if c['profit_quality'] else '-'
        fcf_s = f"{c['fcf_ratio']:.1%}" if c['fcf_ratio'] else '-'
        capex_s = f"{c['capex_ratio']:.1%}" if c['capex_ratio'] else '-'
        pay_s = f"{c['payout']:.1%}" if c['payout'] else '-'
        r_qual = c.get('quality_rank') or '-'
        r_fcf = c.get('fcf_rank') or '-'
        r_capex = c.get('capex_rank') or '-'
        print(f"{c['master_rank']:>3}  {c['company']:<16} {c['industry']:<8} "
              f"{c['combined']:>6} {ev_s:>8} {c['ev_rank']:>5} "
              f"{roic_s:>9} {c['roic_rank']:>6} "
              f"{qual_s:>7} {r_qual:>6} "
              f"{fcf_s:>7} {r_fcf:>5} "
              f"{capex_s:>7} {r_capex:>5} "
              f"{pay_s:>8} {c['payout_rank']:>6}  {c['code']}")

    # Stats
    in_all = len(combined)
    in_ev = len(ev_only)
    in_roic = len(roic_only)
    in_pay = len(payout_only)
    missing_pay = [m for m in merged if m['ev_rank'] and m['roic_rank'] and m['quality_rank'] and not m['payout_rank']]
    missing_ev = [m for m in merged if m['payout_rank'] and not m['ev_rank']]

    print(f"\nTotal: {len(all_keys)} companies | "
          f"EV/EBIT: {in_ev} | ROIC: {in_roic} | Profit Quality: {len(quality_only)} | "
          f"FCF Ratio: {len(fcf_only)} | Capex Ratio: {len(capex_only)} | Payout: {in_pay} | "
          f"Combined ({args.preset}, {n_dims} dims): {in_all}")

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
