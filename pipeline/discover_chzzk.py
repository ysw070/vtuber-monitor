#!/usr/bin/env python3
"""치지직 태그 기반 버튜버 자동 탐지.

배경: 중소 수집원('나의 작은 버튜버')은 버튜버가 직접 등록해야 잡히고, 기업 멤버 명단은
운영자가 아는 사람만 담긴다. 그래서 어디에도 등록되지 않은 버튜버(예: 백아란)가 빠졌다.
치지직은 방송 태그로 '지금 방송 중'인 채널을 조회할 수 있으므로, 이를 여러 번 모아
목록 밖 채널을 찾아낸다.

사용:
  python pipeline/discover_chzzk.py scan      # 태그 방송 조회 → 새 채널만 기록(팔로워는 새 채널만 조회)
  python pipeline/discover_chzzk.py refresh   # 기록된 채널 전체의 팔로워·채널명 갱신(주간 빌드에서)

출력: data/discovered_chzzk.json  {channelId: {name, followers, verified, tags, first_seen, last_seen, seen_days}}
API(공개, 키 불필요):
  - https://api.chzzk.naver.com/service/v1/tag/lives?size=50&sortType=POPULAR&tags=<태그>[&concurrentUserCount=&liveId=]
  - https://api.chzzk.naver.com/service/v1/channels/<channelId>
"""
import os, sys, json, time, datetime, urllib.parse, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'data', 'discovered_chzzk.json')
TAGS = ['버튜버', '버추얼', '버츄얼', 'VTuber', '버추얼유튜버']
MAX_PAGES = 12
UA = 'Mozilla/5.0 (vtuber-monitor; discovery)'
TODAY = datetime.date.today().isoformat()


def get_json(url):
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept': 'application/json'})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.load(r)


def channel(cid):
    c = (get_json(f'https://api.chzzk.naver.com/service/v1/channels/{cid}').get('content') or {})
    if not c.get('channelId'):
        return None
    return {'name': c.get('channelName') or '', 'followers': int(c.get('followerCount') or 0),
            'verified': bool(c.get('verifiedMark'))}


def load():
    try:
        return json.load(open(OUT, encoding='utf-8'))
    except Exception:
        return {}


def save(d):
    json.dump(dict(sorted(d.items(), key=lambda kv: -(kv[1].get('followers') or 0))),
              open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


def scan():
    d = load()
    found = {}
    for tag in TAGS:
        nxt = None
        for _ in range(MAX_PAGES):
            q = {'size': 50, 'sortType': 'POPULAR', 'tags': tag}
            if nxt:
                q.update({'concurrentUserCount': nxt['concurrentUserCount'], 'liveId': nxt['liveId']})
            try:
                c = get_json('https://api.chzzk.naver.com/service/v1/tag/lives?' + urllib.parse.urlencode(q)).get('content') or {}
            except Exception as e:
                print('  tag 조회 실패', tag, e); break
            for x in c.get('data') or []:
                ch = x.get('channel') or {}
                cid = ch.get('channelId')
                if not cid: continue
                f = found.setdefault(cid, {'name': ch.get('channelName') or '', 'tags': set()})
                f['tags'].update(t for t in (x.get('tags') or []) if t)
            nxt = (c.get('page') or {}).get('next')
            if not nxt: break
            time.sleep(0.3)
    new = 0
    for cid, f in found.items():
        rec = d.get(cid)
        if rec is None:
            rec = {'name': f['name'], 'followers': 0, 'verified': False, 'tags': [], 'first_seen': TODAY,
                   'last_seen': TODAY, 'seen_days': 1}
            try:
                info = channel(cid)
                if info: rec.update(info)
                time.sleep(0.25)
            except Exception as e:
                print('  채널 조회 실패', cid, e)
            d[cid] = rec; new += 1
        elif rec.get('last_seen') != TODAY:
            rec['last_seen'] = TODAY; rec['seen_days'] = rec.get('seen_days', 1) + 1
        rec['tags'] = sorted(set(rec.get('tags') or []) | f['tags'])[:12]
    save(d)
    big = sum(1 for cid in found if d[cid].get('first_seen') == TODAY and (d[cid].get('followers') or 0) >= 1000)
    print(f'scan: 방송 중 태그 채널 {len(found)} · 신규 기록 {new} (1천+ {big}) · 누적 {len(d)}')
    if os.environ.get('GITHUB_OUTPUT'):
        with open(os.environ['GITHUB_OUTPUT'], 'a') as g:
            g.write(f'new_big={big}\n')
    return new


def refresh():
    d = load(); ok = 0
    for cid, rec in d.items():
        try:
            info = channel(cid)
            if info:
                rec.update(info); ok += 1
            time.sleep(0.2)
        except Exception as e:
            print('  갱신 실패', cid, e)
    save(d)
    print(f'refresh: {ok}/{len(d)} 갱신')


if __name__ == '__main__':
    mode = sys.argv[1] if len(sys.argv) > 1 else 'scan'
    scan() if mode == 'scan' else refresh()
