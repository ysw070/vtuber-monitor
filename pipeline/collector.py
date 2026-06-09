#!/usr/bin/env python3
"""
한국 버튜버 데이터 수집 PoC — '나의 작은 버튜버' (mysmallvtuber.net)

배경:
  1차 타깃이었던 브니버스(vniverse.live)는 2026-06 기준 DNS 권한 서버가
  질의를 거부(REFUSED)하여 사이트 자체가 접속 불가 → 대체 소스로 전환.

소스:
  - 사이트: https://mysmallvtuber.net (한국 중소/개인 버튜버 집계 사이트, SvelteKit SPA)
  - API:    https://api.mysmallvtuber.net/get_allVtuberList2.php
            (프런트엔드 JS 번들에서 확인한 공개 XHR endpoint)
  - 인증:   요청 헤더 `Accept-Version` = SHA-256("mysmallvtuber" + UTC 날짜 YYYYMMDD)
            (프런트엔드가 매 요청마다 계산해 붙이는 일일 토큰. 사이트 코드 그대로 재현)

출력:
  collected_vtubers.csv (UTF-8-sig)
  - 정규화(long) 테이블: 버튜버 1명 × 플랫폼 채널 1개 = 1행
  - vtuber_id 로 그룹핑하면 버튜버 단위로 복원 가능

주기 실행 시 주의:
  - Accept-Version 토큰은 UTC 날짜 기준 → 자정(UTC) 직후 캐시된 토큰 사용 금지
  - 개인 운영 사이트이므로 호출 빈도는 하루 1~2회 이하 권장 (응답에 cached 필드 존재,
    서버측 캐시 운용 중)
  - 소속(agency) 필드는 이 소스에 없음 → affiliation 컬럼은 공란
"""

import csv
import hashlib
import json
import re
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

API_BASE = "https://api.mysmallvtuber.net"
LIST_ENDPOINT = "/get_allVtuberList2.php"
import os
OUT_CSV = Path(os.environ.get("OUT_CSV", Path(__file__).resolve().parent / "collected_vtubers.csv"))

CSV_FIELDS = [
    "vtuber_id",        # 소스 내부 고유 ID
    "name",             # 활동명 (소스 표기 그대로)
    "name_en",          # 영문명 (활동명에서 라틴 문자 부분 추출 — 휴리스틱)
    "platform",         # youtube / chzzk / soop
    "channel_id",
    "channel_url",
    "channel_name",     # 해당 플랫폼에서의 채널명
    "followers",        # 유튜브=구독자, 치지직=팔로워, SOOP=즐겨찾기
    "debut_date",       # 데뷔일 (YYYY-MM-DD)
    "affiliation",      # 소속 — 소스에 필드 없음(중소/개인 버튜버 위주 사이트), 공란
    "last_live_at",     # 최근 라이브 시각 (소스 기준)
    "last_live_title",
    "twitter_id",
    "community_link",
    "source",           # 데이터 출처
    "collected_at",     # 수집 시각 (UTC)
]


def daily_token() -> str:
    """사이트 프런트엔드와 동일한 일일 인증 토큰 생성."""
    day = datetime.now(timezone.utc).strftime("%Y%m%d")
    return hashlib.sha256(f"mysmallvtuber{day}".encode()).hexdigest()


def fetch_vtuber_list() -> dict:
    params = urllib.parse.urlencode({"hide": 0, "e": 0, "f": 0})
    url = f"{API_BASE}{LIST_ENDPOINT}?{params}"
    req = urllib.request.Request(url, headers={
        "Accept-Version": daily_token(),
        "User-Agent": "Mozilla/5.0 (vtuber-data-poc; contact: ysw070@gmail.com)",
        "Origin": "https://mysmallvtuber.net",
        "Referer": "https://mysmallvtuber.net/",
        "Accept": "application/json",
    })
    with urllib.request.urlopen(req, timeout=60) as resp:
        if resp.status != 200:
            raise RuntimeError(f"HTTP {resp.status}")
        return json.loads(resp.read().decode("utf-8"))


_LATIN_RE = re.compile(r"[A-Za-z][A-Za-z0-9 .'_-]*[A-Za-z0-9]|[A-Za-z]")


def extract_english_name(name: str) -> str:
    """활동명에 병기된 라틴 문자 표기를 추출 (없으면 공란)."""
    if not name:
        return ""
    matches = _LATIN_RE.findall(name)
    # 활동명 전체가 라틴이면 그대로, 한글+라틴 병기면 라틴 부분만
    joined = " ".join(m.strip() for m in matches).strip()
    return joined


def to_date(value: str) -> str:
    """'YYYY-MM-DD HH:MM:SS' → 'YYYY-MM-DD'. 비정상 값/플레이스홀더는 공란."""
    if not value:
        return ""
    s = str(value)
    if not s.startswith(("19", "20")):
        return ""
    d = s[:10]
    if d == "2000-01-01":  # 소스의 '데뷔일 미입력' 플레이스홀더
        return ""
    return d


def to_int(value) -> str:
    try:
        n = int(str(value))
        return str(n) if n > 0 else ""
    except (TypeError, ValueError):
        return ""


def normalize(raw: list[dict], collected_at: str) -> list[dict]:
    rows = []
    for v in raw:
        common = {
            "vtuber_id": v.get("id") or v.get("idx") or "",
            "name": (v.get("name") or "").strip(),
            "name_en": extract_english_name(v.get("name") or ""),
            "debut_date": to_date(v.get("debut_time")),
            "affiliation": "",  # 소스에 소속사 필드 없음
            "last_live_at": (v.get("last_live_time") or "")[:19],
            "last_live_title": (v.get("last_live_title") or "").strip(),
            "twitter_id": (v.get("twitter_id") or "").strip(),
            "community_link": (v.get("community_link") or "").strip(),
            "source": "mysmallvtuber.net",
            "collected_at": collected_at,
        }

        platforms = []
        for yt_id_key, yt_name_key, yt_cnt_key in (
            ("youtube_id_1", "youtube_name_1", "youtube_count_1"),
            ("youtube_id_2", "youtube_name_2", "youtube_count_2"),
        ):
            yid = (v.get(yt_id_key) or "").strip()
            if yid:
                platforms.append({
                    "platform": "youtube",
                    "channel_id": yid,
                    "channel_url": f"https://www.youtube.com/channel/{yid}",
                    "channel_name": (v.get(yt_name_key) or "").strip(),
                    "followers": to_int(v.get(yt_cnt_key)),
                })

        cid = (v.get("chzzk_id") or "").strip()
        if cid:
            platforms.append({
                "platform": "chzzk",
                "channel_id": cid,
                "channel_url": f"https://chzzk.naver.com/{cid}",
                "channel_name": (v.get("chzzk_name") or "").strip(),
                "followers": to_int(v.get("chzzk_count")),
            })

        aid = (v.get("afreeca_id") or "").strip()
        if aid:
            platforms.append({
                "platform": "soop",
                "channel_id": aid,
                "channel_url": f"https://ch.sooplive.co.kr/{aid}",
                "channel_name": (v.get("afreeca_name") or "").strip(),
                "followers": to_int(v.get("afreeca_count")),
            })

        if not platforms:  # 플랫폼 채널이 하나도 없으면 버튜버 메타만 1행
            platforms.append({
                "platform": "", "channel_id": "", "channel_url": "",
                "channel_name": "", "followers": "",
            })

        for p in platforms:
            rows.append({**common, **p})
    return rows


def main() -> int:
    collected_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    print(f"[1/3] fetching {API_BASE}{LIST_ENDPOINT} ...")
    data = fetch_vtuber_list()
    raw = data.get("all") or []
    print(f"      vtubers in response: {len(raw)} (cached={data.get('cached')})")
    if not raw:
        print("ERROR: empty response", file=sys.stderr)
        return 1

    print("[2/3] normalizing ...")
    rows = normalize(raw, collected_at)

    print(f"[3/3] writing {OUT_CSV} ...")
    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        w.writeheader()
        w.writerows(rows)

    n_vtubers = len({r["vtuber_id"] for r in rows})
    by_platform = {}
    for r in rows:
        by_platform[r["platform"] or "(none)"] = by_platform.get(r["platform"] or "(none)", 0) + 1
    print(f"done: {n_vtubers} vtubers, {len(rows)} channel rows -> {by_platform}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
