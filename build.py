#!/usr/bin/env python3
"""배포용 통합 빌드 오케스트레이터.

순서: (1) 수집 → (2) 스냅샷 적재 → (3) 대시보드 빌드(public 모드).
GitHub Actions 와 로컬에서 동일하게 동작. 경로는 이 파일 기준 상대경로로 고정.

환경변수:
  GH_REPO      예) ysw070/vtuber-monitor  (재수집 버튼이 가리킬 Actions 레포)
  SKIP_COLLECT 1 이면 수집 단계를 건너뜀(기존 CSV로 빌드만)
"""
import os, sys, glob, subprocess, shutil

ROOT = os.path.dirname(os.path.abspath(__file__))
PIPE = os.path.join(ROOT, 'pipeline')
DATA = os.path.join(ROOT, 'data')
PUB = os.path.join(ROOT, 'public')
CSV = os.path.join(DATA, 'collected_vtubers.csv')
SNAP = os.path.join(DATA, 'snapshots')
os.makedirs(PUB, exist_ok=True)
os.makedirs(SNAP, exist_ok=True)


def run(desc, argv, env=None):
    print(f'\n=== {desc} ===')
    e = dict(os.environ); e.update(env or {})
    r = subprocess.run([sys.executable] + argv, env=e)
    return r.returncode == 0


def latest_xlsx():
    xs = glob.glob(os.path.join(DATA, '한국_버추얼아이돌_전수조사_v*.xlsx'))
    if not xs:
        xs = glob.glob(os.path.join(DATA, '*.xlsx'))
    return max(xs) if xs else None


def main():
    # 1) 수집 (실패해도 빌드는 계속 — 기존 CSV 유지)
    if os.environ.get('SKIP_COLLECT') != '1':
        ok = run('1/3 수집', [os.path.join(PIPE, 'collector.py')], {'OUT_CSV': CSV})
        if not ok:
            print('⚠ 수집 실패 — 기존 CSV 로 계속 진행')
    else:
        print('SKIP_COLLECT=1 → 수집 건너뜀')

    if not os.path.exists(CSV):
        print('✖ CSV 가 없어 빌드 불가'); return 1

    # 2) 스냅샷 적재 + index/diff 갱신
    run('2/3 스냅샷', [os.path.join(PIPE, 'snapshot.py'), 'save'],
        {'CSV_PATH': CSV, 'SNAP_DIR': SNAP})

    # 3) 대시보드 빌드 (public 모드)
    xlsx = latest_xlsx()
    if not xlsx:
        print('✖ 전수조사 xlsx 없음'); return 1
    out = os.path.join(PUB, 'index.html')
    ok = run('3/3 대시보드', [os.path.join(PIPE, 'build_dashboard.py'), xlsx, CSV, out],
             {'SITE_MODE': 'public', 'GH_REPO': os.environ.get('GH_REPO', '')})
    if not ok:
        return 1
    print(f'\n✔ 빌드 완료 → {out}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
