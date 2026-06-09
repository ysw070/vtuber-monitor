# 한국 버추얼 아이돌·버튜버 모니터링 — 공개 사이트

내 컴퓨터와 **완전히 무관하게** GitHub의 클라우드에서 매월 자동으로 데이터를 수집하고, 정적 웹사이트로 배포하는 구성입니다. 서버를 직접 운영하거나 비용을 낼 필요가 없습니다(공개 레포 기준 무료).

## 동작 구조

```
매월 1일 09:00(KST)  또는  "수동 수집 실행" 버튼
        │
        ▼
GitHub Actions 워크플로 (ubuntu 러너, 클라우드)
   1) collector.py   → 나의 작은 버튜버 API 수집 → data/collected_vtubers.csv
   2) snapshot.py    → data/snapshots/<날짜>.csv 보존 + index/diff 갱신(추세용)
   3) build.py       → public/index.html 재생성 (그룹목록 xlsx + 수집 데이터 + 추세)
   4) 변경분을 레포에 커밋(스냅샷 히스토리 영구 보존)
   5) public/ 을 GitHub Pages 로 배포
        │
        ▼
   https://<사용자명>.github.io/<레포명>/  (또는 연결한 도메인)
```

수집 서버가 GitHub 클라우드에서 돌기 때문에 PC가 꺼져 있어도 매월 자동 실행됩니다.

## 폴더 구성

```
deploy/
├─ .github/workflows/monthly-collect.yml   # 월간 cron + 수동 실행 워크플로
├─ build.py                                # 통합 빌드(수집→스냅샷→대시보드)
├─ requirements.txt                        # openpyxl
├─ pipeline/
│  ├─ collector.py        # 나의 작은 버튜버 API 수집
│  ├─ snapshot.py         # 스냅샷 적재 + 추세/변동 계산
│  └─ build_dashboard.py  # 대시보드 HTML 빌더(public 모드 지원)
├─ data/
│  ├─ 한국_버추얼아이돌_전수조사_v8.xlsx   # 그룹 전수목록(수동 관리 마스터)
│  ├─ collected_vtubers.csv               # 최신 수집본
│  └─ snapshots/                          # 월별 스냅샷 + index.json + diff_latest.json
└─ public/
   └─ index.html          # 배포되는 사이트(자동 생성, 직접 수정 금지)
```

---

## 1단계 — GitHub 레포 만들고 올리기

1. github.com 로그인 → 우측 상단 **+** → **New repository**
2. 이름 예: `vtuber-monitor` · **Public** 선택 · 나머지 비움 → **Create repository**
3. 이 `deploy/` 폴더의 **내용물**을 레포 루트로 올립니다. 터미널이 익숙하면:

   ```bash
   cd "deploy 폴더 경로"
   git init
   git add .
   git commit -m "초기 구성"
   git branch -M main
   git remote add origin https://github.com/<사용자명>/vtuber-monitor.git
   git push -u origin main
   ```

   터미널이 부담되면 GitHub 웹의 **uploading an existing file** 로 폴더 내용을 드래그해도 됩니다(단, `.github` 폴더가 빠지지 않게 주의 — 숨김 폴더라 누락되기 쉬움).

## 2단계 — GitHub Pages 켜기

1. 레포 → **Settings** → 좌측 **Pages**
2. **Build and deployment** → **Source** 를 **GitHub Actions** 로 선택(저장 자동)

## 3단계 — 첫 배포 실행

1. 레포 → **Actions** 탭 → 좌측 **월간 수집 및 배포** 워크플로 선택
2. 우측 **Run workflow** → **Run workflow** 클릭
3. 2~4분 후 초록 체크가 뜨면 완료. **Settings → Pages** 상단에 사이트 주소가 표시됩니다:
   `https://<사용자명>.github.io/vtuber-monitor/`

> 사이트의 **"수동 수집 실행"** 버튼은 바로 이 Actions 페이지를 엽니다. 레포 권한이 있는 사람(=나)이 거기서 Run workflow 를 누르면 즉시 재수집됩니다. 비밀키를 사이트에 노출하지 않는 안전한 방식입니다.

이후부터는 **매월 1일 오전 9시(KST)** 에 자동으로 수집·재배포됩니다. 손댈 것 없습니다.

---

## 4단계 (선택) — 내 도메인 연결하기

도메인 없이도 위 `github.io` 주소로 바로 공개됩니다. 보기 좋은 주소를 원하면 도메인을 연결하세요. 도메인은 "구매"가 아니라 **연 단위 임대**입니다.

### 등록처 추천

| 등록처 | .com 연간(대략) | 특징 |
|---|---|---|
| **Cloudflare Registrar** (추천) | 약 $10.44 (≈1.4만원) | 원가 그대로, 마크업 0. 단 Cloudflare 네임서버 사용 필수. 갱신가도 동일해 장기적으로 가장 저렴 |
| 가비아 | .com 약 23,000원 / **.kr 13,500원** | 국내 1위, 한국어 지원·결제 편리. `.kr`/`.co.kr` 쓰려면 사실상 국내 등록처 |

`.com`이면 Cloudflare가 가장 저렴하고 깔끔하며, 한국어 지원이나 `.kr` 도메인을 원하면 가비아가 편합니다. (2026년 1월부로 일부 TLD 원가 인상이 있었으니 결제 화면 금액을 최종 확인하세요.)

### 연결 방법(2가지)

도메인을 사이트에 붙이는 건 **GitHub 레포 Settings → Pages → Custom domain** 에 도메인을 입력하고, 등록처 DNS에 레코드를 추가하면 됩니다.

- **루트 도메인**(예: `myvtuber.com`): 등록처 DNS에 A 레코드 4개 추가
  `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153`
- **서브도메인**(예: `vtuber.myvtuber.com`, 더 간단·권장): CNAME 레코드 1개
  이름 `vtuber` → 값 `<사용자명>.github.io`

설정 후 Pages 화면에서 **Enforce HTTPS** 를 체크하면 무료 SSL(https)이 자동 적용됩니다(전파에 수십 분~수 시간).

> Cloudflare에서 도메인을 샀다면 DNS도 Cloudflare에서 관리하므로, 위 레코드를 Cloudflare 대시보드에 넣으면 됩니다. 실제 도메인을 정하면 그 도메인 기준으로 정확한 레코드 값을 다시 정리해 드릴게요.

---

## 운영 메모

- **그룹 전수목록 갱신**: `data/한국_버추얼아이돌_전수조사_v8.xlsx` 가 마스터입니다. 그룹 추가/상태 변경은 이 파일을 수정해 커밋하면 다음 빌드에 반영됩니다. (더 최신 `v9` 등을 올리면 빌드가 자동으로 최신 버전을 집습니다.)
- **수집 빈도**: 월 1회로 설정했습니다. 나의 작은 버튜버는 개인 운영 사이트이므로 호출 빈도를 함부로 올리지 마세요. 장기적으로는 YouTube/치지직 공식 API로 이전하는 것이 안전합니다.
- **스냅샷 2개 이상**부터 "추세·변동" 탭의 차트와 변동 리포트가 자동으로 채워집니다(현재 1개).
- **로컬 테스트**: `SKIP_COLLECT=1 python3 build.py` 로 수집 없이 빌드만 확인할 수 있습니다.

Sources: [Cloudflare Registrar](https://www.cloudflare.com/products/registrar/), [Cloudflare 가격 2026](https://tldprice.org/registrar/cloudflare), [가비아 도메인](https://domain.gabia.com/)
