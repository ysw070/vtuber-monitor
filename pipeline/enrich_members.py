#!/usr/bin/env python3
"""기업 소속 버튜버(그룹 멤버) 통계 수집.

배경: 주 수집원 '나의 작은 버튜버'는 중소 버튜버 전용이라 이세계아이돌·스텔라이브 같은
기업 소속 멤버가 빠져 있다. 운영자가 관리하는 멤버 명단(data/agency_members.csv)의
채널 ID로 공개 API에서 팔로워 수를 직접 가져온다.

입력:  data/agency_members.csv  (group,agency,member,status,youtube,chzzk,soop,note,source)
출력:  data/member_stats.json   {키: {member, group, agency, status, yt, chzzk, soop, fetched, history}}

수집원:
  - 치지직: https://api.chzzk.naver.com/service/v1/channels/<ID>  (공개, 키 불필요)
  - 유튜브: YouTube Data API v3 channels.list (YT_API_KEY 있을 때만, ID 50개 묶음 = 1 unit)
  - SOOP:   https://chapi.sooplive.co.kr/api/<ID>/station  → station.upd.fan_cnt(즐겨찾기 수)
실패한 항목은 이전 값을 유지한다(네트워크 오류로 수치가 0이 되지 않게).
"""
import os, csv, json, time, datetime, urllib.parse, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'data', 'agency_members.csv')
OUT = os.path.join(ROOT, 'data', 'member_stats.json')
KEY = os.environ.get('YT_API_KEY', '').strip()
UA = 'Mozilla/5.0 (vtuber-monitor; weekly stats)'
TODAY = datetime.date.today().isoformat()


def get_json(url, headers=None):
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept': 'application/json', **(headers or {})})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.load(r)


def chzzk(cid):
    d = get_json(f'https://api.chzzk.naver.com/service/v1/channels/{cid}')
    c = d.get('content') or {}
    if not c.get('channelId'):
        return None
    return {'id': cid, 'name': c.get('channelName') or '', 'followers': int(c.get('followerCount') or 0),
            'verified': bool(c.get('verifiedMark'))}


def soop(sid):
    d = get_json(f'https://chapi.sooplive.co.kr/api/{sid}/station', {'Referer': 'https://ch.sooplive.co.kr/'})
    st = d.get('station') or {}
    upd = st.get('upd') or {}
    if not upd.get('user_id'):
        return None
    return {'id': sid, 'name': st.get('user_nick') or '', 'followers': int(upd.get('fan_cnt') or 0)}


def yt_call(params):
    q = urllib.parse.urlencode({'part': 'snippet,statistics', 'key': KEY, **params})
    return get_json('https://www.googleapis.com/youtube/v3/channels?' + q).get('items') or []


def yt_item(it):
    st = it.get('statistics', {})
    return {'id': it.get('id', ''), 'title': it.get('snippet', {}).get('title', ''),
            'subs': None if st.get('hiddenSubscriberCount') else int(st.get('subscriberCount', 0)),
            'videos': int(st.get('videoCount', 0))}


def main():
    if not os.path.exists(SRC):
        print('agency_members.csv 없음 → 건너뜀'); return 0
    rows = [r for r in csv.DictReader(open(SRC, encoding='utf-8-sig')) if (r.get('member') or '').strip()]
    try:
        old = json.load(open(OUT, encoding='utf-8'))
    except Exception:
        old = {}
    out = {}
    ok_c = ok_y = ok_s = fail = 0

    # 유튜브: 핸들은 개별 조회(이전에 UC ID를 알아냈으면 재사용), ID는 50개씩 묶음
    yt_by_key = {}
    if KEY:
        ids = {}
        for r in rows:
            k = f"{r['group'].strip()}|{r['member'].strip()}"
            y = (r.get('youtube') or '').strip()
            prev = (old.get(k) or {}).get('yt') or {}
            if y.startswith('UC'):
                ids[k] = y
            elif y.startswith('@'):
                if prev.get('id') and prev.get('handle') == y:
                    ids[k] = prev['id']
                else:
                    try:
                        items = yt_call({'forHandle': y})
                        if items:
                            yt_by_key[k] = {**yt_item(items[0]), 'handle': y}; ok_y += 1
                    except Exception as e:
                        print('  yt handle 실패', y, e); fail += 1
        keys = list(ids)
        for i in range(0, len(keys), 50):
            chunk = keys[i:i + 50]
            try:
                items = {it['id']: it for it in yt_call({'id': ','.join(ids[k] for k in chunk)})}
                for k in chunk:
                    it = items.get(ids[k])
                    if it:
                        h = next(((r.get('youtube') or '').strip() for r in rows
                                  if f"{r['group'].strip()}|{r['member'].strip()}" == k), '')
                        yt_by_key[k] = {**yt_item(it), 'handle': h if h.startswith('@') else ''}; ok_y += 1
            except Exception as e:
                print('  yt batch 실패', e); fail += 1
    else:
        print('YT_API_KEY 없음 → 유튜브 수치는 이전 값 유지')

    for r in rows:
        g, m = r['group'].strip(), r['member'].strip()
        k = f'{g}|{m}'
        prev = old.get(k) or {}
        rec = {'member': m, 'group': g, 'agency': (r.get('agency') or '').strip(),
               'status': (r.get('status') or '활동').strip(), 'note': (r.get('note') or '').strip(),
               'yt': prev.get('yt') or {}, 'chzzk': prev.get('chzzk') or {}, 'soop': {},
               'fetched': prev.get('fetched', ''), 'history': prev.get('history') or {}}
        y = (r.get('youtube') or '').strip()
        if y:
            rec['yt'] = yt_by_key.get(k) or ({**rec['yt']} if rec['yt'].get('handle') == y or rec['yt'].get('id') == y
                                             else {'id': y if y.startswith('UC') else '', 'handle': y if y.startswith('@') else ''})
        else:
            rec['yt'] = {}
        cid = (r.get('chzzk') or '').strip()
        if cid:
            try:
                c = chzzk(cid)
                if c:
                    rec['chzzk'] = c; ok_c += 1
                time.sleep(0.25)
            except Exception as e:
                print('  chzzk 실패', m, e); fail += 1
                rec['chzzk'] = rec['chzzk'] if rec['chzzk'].get('id') == cid else {'id': cid}
        else:
            rec['chzzk'] = {}
        s = (r.get('soop') or '').strip()
        prev_s = prev.get('soop') or {}
        rec['soop'] = {}
        if s:
            rec['soop'] = prev_s if prev_s.get('id') == s else {'id': s}
            try:
                x = soop(s)
                if x:
                    rec['soop'] = x; ok_s += 1
                time.sleep(0.25)
            except Exception as e:
                print('  soop 실패', m, e); fail += 1
        best = max([v for v in [(rec['yt'] or {}).get('subs'), (rec['chzzk'] or {}).get('followers'),
                                (rec['soop'] or {}).get('followers')] if v] or [0])
        if best:
            rec['history'][TODAY] = best
            rec['fetched'] = TODAY
        out[k] = rec

    json.dump(out, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'members {len(out)} · chzzk {ok_c} · soop {ok_s} · youtube {ok_y} · 실패 {fail} → {OUT}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
