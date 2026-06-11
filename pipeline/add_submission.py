#!/usr/bin/env python3
"""GitHub 이슈 본문(ISSUE_BODY 환경변수)에서 그룹 제보 데이터를 파싱해
data/manual_groups.csv 에 한 줄 추가한다. group-submission.yml 워크플로에서 호출.

이슈 본문에는 폼이 심어둔 `<!--GROUPDATA:{...json...}-->` 마커가 있어야 한다.
없으면(일반 이슈) 아무것도 하지 않고 종료한다(exit 0).
"""
import os, re, json, csv

BODY = os.environ.get('ISSUE_BODY', '')
PATH = os.path.join('data', 'manual_groups.csv')
FIELDS = ['name', 'name_en', 'agency', 'year', 'category', 'status', 'channel', 'note', 'submitter']
ALLOWED_CAT = {'스트리밍/MCN형', '음반/아이돌형', '버추얼휴먼/솔로', '혼합형'}
ALLOWED_ST = {'활동중', '휴면', '확인필요'}


def fail(msg):
    print(f'::notice::{msg}')
    raise SystemExit(0)


m = re.search(r'GROUPDATA:(.*?)-->', BODY, re.S)
if not m:
    fail('GROUPDATA 마커 없음 — 제보 이슈가 아니므로 건너뜀')

try:
    d = json.loads(m.group(1).strip())
except Exception as e:
    fail(f'GROUPDATA JSON 파싱 실패: {e}')

name = str(d.get('name', '')).strip()
if not name:
    fail('그룹명이 비어 있어 건너뜀')

# 길이 제한(스팸/악용 완화)
clean = {}
for k in FIELDS:
    v = str(d.get(k, '') or '').strip().replace('\n', ' ')
    clean[k] = v[:200]
if clean['category'] not in ALLOWED_CAT:
    clean['category'] = '스트리밍/MCN형'
if clean['status'] not in ALLOWED_ST:
    clean['status'] = '활동중'

exists = os.path.exists(PATH)
rows = []
if exists:
    try:
        rows = list(csv.DictReader(open(PATH, encoding='utf-8-sig')))
    except Exception:
        rows = []

if any((r.get('name', '').strip() == name) for r in rows):
    fail(f'이미 등록된 그룹: {name}')

write_header = not exists or os.path.getsize(PATH) == 0
with open(PATH, 'a', encoding='utf-8-sig', newline='') as f:
    w = csv.DictWriter(f, fieldnames=FIELDS)
    if write_header:
        w.writeheader()
    w.writerow(clean)

print(f'::notice::제보 추가됨: {name}')
# 워크플로 후속 단계용 출력
gho = os.environ.get('GITHUB_OUTPUT')
if gho:
    with open(gho, 'a') as f:
        f.write(f'added_name={name}\n')
        f.write('added=true\n')
