#!/usr/bin/env python3
"""스냅샷 저장 + 추세/차분 계산.

스냅샷 저장: collected_vtubers.csv를 snapshots/<수집일>.csv 로 보존(이미 있으면 덮어쓰지 않음, --force로 강제).
요약 산출: snapshots/index.json 에 월별 집계(총원/규모구간/플랫폼) 누적.
차분 산출: 가장 최근 두 스냅샷 비교 → snapshots/diff_latest.json (신규/이탈/5만±진입이탈/팔로워 톱무버).

사용:
  python3 snapshot.py save [--force]        # 현재 csv를 스냅샷으로 적재 + index 갱신 + diff 갱신
  python3 snapshot.py rebuild               # snapshots/*.csv 전체로 index 재구성
"""
import csv, json, sys, os, glob, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
# 경로는 환경변수로 재정의 가능(배포 레포는 data/ 분리 레이아웃 사용)
CSV = os.environ.get('CSV_PATH', os.path.join(HERE, 'collected_vtubers.csv'))
SNAP = os.environ.get('SNAP_DIR', os.path.join(os.path.dirname(CSV), 'snapshots'))
os.makedirs(SNAP, exist_ok=True)

def load_csv(path):
    return list(csv.DictReader(open(path, encoding='utf-8-sig')))

def agg_by_name(rows):
    """사람 단위 집계. 키는 고정 ID(vtuber_id) — 개명해도 같은 사람으로 본다(없으면 이름).
    값: {n:최신 표기 이름, f:최대팔로워, p:주플랫폼, plats:set, d:데뷔, l:최근라이브}"""
    a = {}
    for r in rows:
        k = (r.get('vtuber_id') or '').strip() or r['name']
        try: f = int(float(r['followers'] or 0))
        except: f = 0
        if k not in a:
            a[k] = {'n': r['name'], 'f': 0, 'p': '', 'plats': set(), 'd': r.get('debut_date') or '', 'l': ''}
        x = a[k]; x['plats'].add(r['platform'])
        if f > x['f']: x['f'], x['p'] = f, r['platform']
        ll = (r.get('last_live_at') or '')[:10]
        if ll > x['l']: x['l'] = ll
    return a

def summarize(rows):
    a = agg_by_name(rows)
    fs = [v['f'] for v in a.values()]
    plats = {}
    for v in a.values(): plats[v['p'] or '미상'] = plats.get(v['p'] or '미상', 0) + 1
    return {
        'total_people': len(a),
        'total_channels': len(rows),
        'n_50k': sum(1 for f in fs if f >= 50000),
        'n_30k': sum(1 for f in fs if f >= 30000),
        'n_10k': sum(1 for f in fs if f >= 10000),
        'n_1k': sum(1 for f in fs if f >= 1000),
        'sum_followers': sum(fs),
        'platforms': plats,
    }

def snap_date(rows):
    ca = max((r.get('collected_at') or '') for r in rows)
    return (ca[:10] or datetime.date.today().isoformat())

def save(force=False):
    rows = load_csv(CSV)
    d = snap_date(rows)
    dst = os.path.join(SNAP, f'{d}.csv')
    if os.path.exists(dst) and not force:
        print(f'snapshot {d} already exists (use --force to overwrite)')
    else:
        with open(dst, 'w', encoding='utf-8-sig', newline='') as f:
            w = csv.DictWriter(f, fieldnames=rows[0].keys()); w.writeheader(); w.writerows(rows)
        print(f'saved snapshot {dst} ({len(rows)} rows)')
    rebuild()
    diff_latest()

def rebuild():
    idx = {}
    for p in sorted(glob.glob(os.path.join(SNAP, '*.csv'))):
        d = os.path.splitext(os.path.basename(p))[0]
        idx[d] = summarize(load_csv(p))
    json.dump(idx, open(os.path.join(SNAP, 'index.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'index rebuilt: {len(idx)} snapshots')
    return idx

def diff_latest():
    snaps = sorted(glob.glob(os.path.join(SNAP, '*.csv')))
    out = os.path.join(SNAP, 'diff_latest.json')
    if len(snaps) < 2:
        json.dump({'available': False, 'reason': '스냅샷이 2개 미만 — 비교는 다음 수집부터'}, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        print('diff: need >=2 snapshots'); return
    pa, pb = snaps[-2], snaps[-1]
    da, db = [os.path.splitext(os.path.basename(x))[0] for x in (pa, pb)]
    a, b = agg_by_name(load_csv(pa)), agg_by_name(load_csv(pb))
    na, nb = set(a), set(b)
    N = lambda d, k: d[k]['n']
    entered = sorted(N(b, k) for k in nb - na)
    left = sorted(N(a, k) for k in na - nb)
    renamed = sorted([{'from': a[k]['n'], 'to': b[k]['n']} for k in nb & na if a[k]['n'] != b[k]['n']],
                     key=lambda x: x['to'])
    up50 = [k for k in (nb & na) if b[k]['f'] >= 50000 and a[k]['f'] < 50000]
    down50 = [k for k in (nb & na) if b[k]['f'] < 50000 and a[k]['f'] >= 50000]
    movers = []
    for k in (nb & na):
        dlt = b[k]['f'] - a[k]['f']
        if a[k]['f'] >= 1000 or b[k]['f'] >= 1000:
            movers.append({'name': b[k]['n'], 'from': a[k]['f'], 'to': b[k]['f'], 'delta': dlt})
    gain = sorted(movers, key=lambda x: -x['delta'])[:15]
    drop = sorted(movers, key=lambda x: x['delta'])[:15]
    res = {'available': True, 'from': da, 'to': db,
           'entered': entered, 'left': left, 'renamed': renamed,
           'up50': [{'name': b[k]['n'], 'to': b[k]['f']} for k in up50],
           'down50': [{'name': b[k]['n'], 'to': b[k]['f']} for k in down50],
           'top_gain': gain, 'top_drop': drop}
    json.dump(res, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'diff {da}->{db}: +{len(entered)} -{len(left)} 개명{len(renamed)} | 5만진입 {len(up50)} 이탈 {len(down50)}')

if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else 'save'
    if cmd == 'save': save('--force' in sys.argv)
    elif cmd == 'rebuild': rebuild(); diff_latest()
    else: print(__doc__)
