#!/usr/bin/env python3
"""유튜브 Data API v3로 제보 그룹(manual_groups.csv)의 채널 통계를 수집.

- 채널 URL에서 핸들(@xxxx) 또는 channel/UCxxxx ID를 추출 → channels.list 호출
- 구독자수·영상수·조회수·채널명을 data/youtube_stats.json 에 핸들/ID 키로 누적 저장
- 환경변수 YT_API_KEY 가 없으면 아무것도 하지 않고 종료(로컬·키 미설정 환경 안전)

쿼터: channels.list = 1 unit/호출. 일일 10,000 units 무료 → 사실상 충분.
실행: python pipeline/enrich_youtube.py  (레포 루트에서)
"""
import os, re, csv, json, datetime, urllib.parse, urllib.request

KEY = os.environ.get('YT_API_KEY', '').strip()
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 레포 루트(=pipeline의 상위)
MANUAL = os.path.join(ROOT, 'data', 'manual_groups.csv')
OUT = os.path.join(ROOT, 'data', 'youtube_stats.json')


def channel_query(url):
    """채널 URL → API 파라미터(dict). 핸들/ID/legacy user 지원."""
    url = url or ''
    m = re.search(r'youtube\.com/@([A-Za-z0-9._\-가-힣]+)', url)
    if m:
        return {'forHandle': '@' + m.group(1)}
    m = re.search(r'youtube\.com/channel/(UC[A-Za-z0-9_\-]+)', url)
    if m:
        return {'id': m.group(1)}
    m = re.search(r'youtube\.com/(?:c|user)/([A-Za-z0-9._\-]+)', url)
    if m:
        return {'forUsername': m.group(1)}
    return None


def fetch(params):
    q = urllib.parse.urlencode({'part': 'snippet,statistics', 'key': KEY, **params})
    req = urllib.request.Request('https://www.googleapis.com/youtube/v3/channels?' + q,
                                 headers={'Accept': 'application/json'})
    with urllib.request.urlopen(req, timeout=30) as r:
        data = json.load(r)
    items = data.get('items') or []
    if not items:
        return None
    it = items[0]
    st = it.get('statistics', {})
    hidden = st.get('hiddenSubscriberCount')
    return {
        'channel_id': it.get('id', ''),
        'title': it.get('snippet', {}).get('title', ''),
        'subscribers': None if hidden else int(st.get('subscriberCount', 0)),
        'videos': int(st.get('videoCount', 0)),
        'views': int(st.get('viewCount', 0)),
    }


def main():
    if not KEY:
        print('YT_API_KEY 없음 → 유튜브 통계 수집 건너뜀(정상)')
        return 0
    rows = list(csv.DictReader(open(MANUAL, encoding='utf-8-sig'))) if os.path.exists(MANUAL) else []
    stats = {}
    if os.path.exists(OUT):
        try:
            stats = json.load(open(OUT, encoding='utf-8'))
        except Exception:
            stats = {}
    today = datetime.date.today().isoformat()
    n = 0
    for r in rows:
        params = channel_query(r.get('channel'))
        if not params:
            continue
        key = list(params.values())[0]  # 핸들/ID를 캐시 키로
        try:
            info = fetch(params)
            if info:
                # 추세용 히스토리(월별 구독자) 누적
                hist = (stats.get(key) or {}).get('history', {})
                if info['subscribers'] is not None:
                    hist[today] = info['subscribers']
                info['history'] = hist
                info['fetched'] = today
                info['group'] = r.get('name', '')
                stats[key] = info
                n += 1
                print(f"ok {key}: 구독 {info['subscribers']} · 영상 {info['videos']}")
        except Exception as e:
            print(f'err {key}: {e}')
    json.dump(stats, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'유튜브 통계 {n}건 갱신 → {OUT}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
