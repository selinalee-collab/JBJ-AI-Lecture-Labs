---
name: stock-opinion
description: "Use this skill when the user asks for '종목 여론 조회', '여론 분석', '토론실 분석', '네이버 토론실', '커뮤니티 여론', '감성 분석', '종목 의견', '주식 여론' about a Korean stock. Navigates to Naver Finance stock discussion board (종목토론실), collects 100 post titles, performs Korean sentiment analysis, and generates an Excel report with charts. Trigger examples: 'XX 여론 조회해줘', 'XX 토론실 분석', 'XX 종목 여론', 'XX 감성 분석해줘'."
---

# stock-opinion — 네이버 종목토론실 여론 분석 Excel 자동 생성

## 목적

사용자가 한국 상장 종목의 여론 조회를 요청하면, 네이버 금융 종목토론실에서 최신 게시글 제목 100개를 수집하여 감성 분석(긍정/중립/부정)을 수행하고, 차트가 포함된 Excel 파일을 생성한다.

**수집 방법**: Playwright Python (헤드리스 Chromium) — claude-in-chrome MCP는 `finance.naver.com` 차단으로 사용 불가.

---

## Step 0 — 환경 점검 및 자동 설치

**가장 먼저 수행한다.** PowerShell에서 아래 두 줄을 실행한다.

```powershell
python -m pip install --quiet playwright openpyxl
python -m playwright install chromium --quiet
```

설치 확인:

```powershell
python -c "from playwright.sync_api import sync_playwright; import openpyxl; print('환경 OK')"
```

출력이 `환경 OK`이면 Step 1로 진행한다. 오류가 발생하면 오류 메시지를 확인하고 재설치한다.

> **최초 실행 시** chromium 다운로드로 1~2분 소요된다. 이미 설치된 경우 수 초 내 완료된다.

---

## Step 1 — 종목 코드 확인

### 주요 종목 코드 조회표

| 기업명 | 종목코드 | | 기업명 | 종목코드 |
|--------|----------|-|--------|----------|
| 삼성전자 | 005930 | | KB금융 | 105560 |
| SK하이닉스 | 000660 | | 신한지주 | 055550 |
| LG에너지솔루션 | 373220 | | 하나금융지주 | 086790 |
| 삼성바이오로직스 | 207940 | | LG전자 | 066570 |
| 현대차 | 005380 | | 현대모비스 | 012330 |
| 기아 | 000270 | | SK텔레콤 | 017670 |
| POSCO홀딩스 | 005490 | | KT | 030200 |
| LG화학 | 051910 | | 두산에너빌리티 | 034020 |
| 셀트리온 | 068270 | | 삼성SDI | 006400 |
| 카카오 | 035720 | | 네이버 | 035420 |

### 표에 없는 종목

네이버 금융 URL로 직접 확인한다:

```
https://finance.naver.com/search/searchList.naver?query={기업명}
```

검색 결과 페이지에서 첫 번째 종목 링크의 `code=` 파라미터 값을 종목코드로 사용한다.

---

## Step 2 — 게시글 100개 수집 (Playwright Python)

아래 **표준 스크래핑 스크립트**를 스크래치패드에 생성하고 실행한다.  
`STOCK_CODE`와 `COMPANY`만 변경하면 모든 종목에 재사용 가능하다.

### 스크래핑 스크립트 템플릿

스크래치패드 경로: `{scratchpad}/scrape_{ticker}.py`  
출력 JSON 경로: `{scratchpad}/{ticker}_posts.json`

```python
# -*- coding: utf-8 -*-
import asyncio, json, html, sys

# ── 설정 ──────────────────────────────────
STOCK_CODE  = "000660"        # 종목코드 (6자리)
COMPANY     = "SK하이닉스"    # 기업명
TOTAL_PAGES = 5               # 수집 페이지 수 (페이지당 20건)
OUTPUT_JSON = r"{scratchpad}\{ticker}_posts.json"
# ──────────────────────────────────────────

JS_EXTRACT = """
() => {
    const rows = document.querySelectorAll('table.type2 tr');
    const posts = [];
    rows.forEach(row => {
        const titleEl = row.querySelector('td.title a');
        if (!titleEl) return;
        const cells = row.querySelectorAll('td');
        posts.push({
            title:    titleEl.innerText.trim(),
            date:     cells[3] ? cells[3].innerText.trim() : '',
            views:    cells[4] ? cells[4].innerText.trim().replace(/,/g,'') : '0',
            likes:    cells[5] ? cells[5].innerText.trim().replace(/,/g,'') : '0',
            dislikes: cells[6] ? cells[6].innerText.trim().replace(/,/g,'') : '0'
        });
    });
    return JSON.stringify(posts);
}
"""

async def scrape():
    from playwright.async_api import async_playwright
    all_posts = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled",
                  "--no-sandbox", "--disable-dev-shm-usage"]
        )
        ctx = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="ko-KR", timezone_id="Asia/Seoul"
        )
        page = await ctx.new_page()
        for pg in range(1, TOTAL_PAGES + 1):
            url = f"https://finance.naver.com/item/board.naver?code={STOCK_CODE}&page={pg}"
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=20000)
                await page.wait_for_timeout(1200)
                raw = await page.evaluate(JS_EXTRACT)
                items = json.loads(raw)
                for item in items:
                    item["title"] = html.unescape(item["title"])
                all_posts.extend(items)
                print(f"[페이지 {pg}] {len(items)}건 수집 (누적: {len(all_posts)}건)")
            except Exception as e:
                print(f"[페이지 {pg}] 오류: {e}")
        await browser.close()
    return all_posts

posts = asyncio.run(scrape())
for i, p in enumerate(posts, 1):
    p["no"] = i

with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(posts, f, ensure_ascii=False, indent=2)

print(f"\n[완료] 총 {len(posts)}건 저장: {OUTPUT_JSON}")
for item in posts[:5]:
    print(f"  {item['no']}. {item['title']}")
```

### 실행

```powershell
python "{scratchpad}\scrape_{ticker}.py"
```

### 수집 결과 확인

실행 후 JSON 파일을 Read 도구로 읽어 100개 제목 전체를 컨텍스트에 로드한다.

---

## Step 3 — 감성 분석 (Claude 직접 수행)

JSON에서 읽은 100개 제목 전체를 한 번에 분석한다.

### 분류 기준

| 분류 | 설명 |
|------|------|
| **긍정** | 주가 상승·매수·호재 기대·실적 개선 등 긍정적 전망 |
| **부정** | 주가 하락·매도·악재 우려·손절·규제 등 부정적 전망 |
| **중립** | 단순 질문·정보 공유·관망·무관 내용 |

### 감성 점수 척도

| 분류 | 범위 | 예시 |
|------|------|------|
| 강한 긍정 | +0.7 ~ +1.0 | 급등, 신고가, 폭등, 역대급 호재, 강력 매수 |
| 약한 긍정 | +0.1 ~ +0.6 | 상승, 호재, 기대, 반등, 매수 추천 |
| 중립 | 0.0 | 오늘 거래량은?, 뉴스 공유, 단순 질문 |
| 약한 부정 | -0.1 ~ -0.6 | 하락, 우려, 조정, 관망 |
| 강한 부정 | -0.7 ~ -1.0 | 급락, 폭락, 손절, 악재, 인버스 몰빵 |

### 각 게시글 분석 항목

- `sentiment`: "긍정" / "중립" / "부정"
- `score`: -1.0 ~ +1.0 (소수점 2자리)
- `keywords`: 판단 근거 핵심 단어 최대 3개 (없으면 빈 문자열)

---

## Step 4 — Excel 생성 (Python + openpyxl)

### 출력 파일

```
파일명: {기업명}_여론분석_{YYYYMMDD}.xlsx
저장처: C:\Users\JBJ\JBJ-AI-Lecture-Labs\M07_skill\
```

### Excel 생성 스크립트 작성 원칙

1. **설정 블록** (스크립트 최상단):
   ```python
   COMPANY_NAME = "SK하이닉스"
   STOCK_CODE   = "000660"
   EXCHANGE     = "유가증권시장(KOSPI)"
   COLLECT_DATE = "2026-06-28"
   OUTPUT_DIR   = r"C:\Users\JBJ\JBJ-AI-Lecture-Labs\M07_skill"
   OUTPUT_FILE  = f"{COMPANY_NAME}_여론분석_{YYYYMMDD}.xlsx"
   ```

2. **POSTS 리스트** (설정 블록 바로 아래, 하드코딩):
   ```python
   # (no, title, date, views, likes, dislikes, sentiment, score, keywords)
   POSTS = [
       (1, "제목...", "날짜", views, likes, dislikes, "긍정", 0.70, "키워드1,키워드2"),
       ...
   ]
   ```

3. **openpyxl 자동 설치**:
   ```python
   try:
       import openpyxl
   except ImportError:
       import subprocess, sys
       subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
       import openpyxl
   ```

4. **시트별 try/except** — 한 시트 실패 시 나머지 계속 진행

5. **이모지·특수문자 print 금지** — Windows cp949 인코딩 오류 방지  
   성공 메시지: `print("[완료] 파일 저장:", os.path.abspath(out_path))`

6. 실행: `python "{scratchpad}\gen_{ticker}_excel.py"`

---

### Sheet 구성 (3개 시트)

#### Sheet 1: 요약 대시보드

**메타 정보 블록 (1~8행)**

| 항목 | 값 |
|------|----|
| 기업명 | {회사명} |
| 종목코드 | {코드} |
| 거래소 | {거래소} |
| 수집일시 | {YYYY-MM-DD} |
| 수집 게시글 수 | {N}개 |
| 분석 기준 | 네이버 금융 종목토론실 최신글 |
| 평균 감성 점수 | {avg} |
| 여론 종합 | 긍정 우세 / 팽팽 / 부정 우세 |

**감성 분포 테이블 (10~14행)**

| 감성 | 건수 | 비율(%) | 평균점수 | 평균조회수 | 평균공감수 |
|------|------|---------|---------|-----------|-----------|
| 긍정 | | | | | |
| 중립 | | | | | |
| 부정 | | | | | |
| 합계 | 100 | 100.0 | | | |

**차트 2개**

1. **파이차트** (H2): 긍정/중립/부정 비율 — 색상: 긍정=`4472C4`, 중립=`A9A9A9`, 부정=`FF0000` — 14×12cm
2. **묶음 막대차트** (H22): 감성별 평균 조회수·공감수 — 색상: 조회수=`2E75B6`, 공감수=`ED7D31` — 16×12cm

---

#### Sheet 2: 게시글 원본 데이터

| 열 | 항목 | 형식 |
|----|------|------|
| A | 번호 | 정수 (1~100) |
| B | 제목 | 텍스트, 열 너비 50 |
| C | 날짜 | 텍스트 |
| D | 조회수 | 정수 #,##0 |
| E | 공감 | 정수 |
| F | 비공감 | 정수 |
| G | 감성 | 긍정=배경`DDEEFF` / 부정=배경`FFE0E0` |
| H | 감성점수 | #,##0.00, 양수=파랑`1F3864`, 음수=빨강 |
| I | 근거키워드 | 텍스트, 회색 |

**꺾은선 차트** (K2): 번호별 감성점수 추이 — 22×12cm

---

#### Sheet 3: 키워드 분석

- **긍정 키워드 Top 15** (A1:B16): 키워드 | 등장 횟수
- **부정 키워드 Top 15** (D1:E16): 키워드 | 등장 횟수
- **긍정 수평 막대차트** (G2): 색상 `4472C4` — 16×12cm
- **부정 수평 막대차트** (G22): 색상 `FF0000` — 16×12cm

---

### 공통 서식 규칙

```
헤더 행:      배경 #1F3864, 글자 흰색, Bold, 가운데 정렬
합계/소계 행: 배경 #2E75B6, 글자 흰색, Bold
짝수 데이터행: 배경 #F2F2F2
홀수 데이터행: 배경 #FFFFFF
폰트:          맑은 고딕 10pt
행 높이:       20pt 고정
틀 고정:       1행 + A열
경계선:        thin, #BFBFBF
```

---

## Step 5 — 결과 보고

```
[여론 분석 결과 — {기업명} ({종목코드})]
수집일시: {날짜}  |  수집 건수: {N}개

감성 분포:
  긍정 {n}건 ({p}%)  ·  중립 {n}건 ({p}%)  ·  부정 {n}건 ({p}%)

평균 감성 점수: {avg} (범위 -1.0 ~ +1.0)

여론 종합: {긍정 우세 / 팽팽 / 부정 우세}

주요 긍정 키워드: {keyword1}, {keyword2}, {keyword3}
주요 부정 키워드: {keyword1}, {keyword2}, {keyword3}

저장 파일: {절대경로}
```

---

## 오류 처리

| 상황 | 대응 |
|------|------|
| playwright/openpyxl 미설치 | Step 0 pip install로 자동 해결 |
| chromium 미설치 | Step 0 `playwright install chromium`으로 자동 해결 |
| 특정 페이지 수집 실패 | 해당 페이지 스킵, 나머지 계속 — 수집 건수 명시 |
| 100건 미달 | 수집된 건수로 진행, "XX건 수집(100건 미달)" 명시 |
| Excel 시트 생성 실패 | 해당 시트 오류 출력 후 나머지 시트 계속 저장 |

---

## 주의사항

- 네이버 금융 종목토론실은 **로그인 없이** 열람 가능하다.
- 수집한 제목에 HTML 엔티티(`&amp;`, `&lt;` 등)가 포함된 경우 `html.unescape()`로 자동 정제한다.
- 감성 분석은 **제목 텍스트만** 기준으로 한다 (본문 미참조).
- 동일 날짜·종목 재조회 시 파일명에 `_2`, `_3` suffix를 붙여 덮어쓰기를 방지한다.
- claude-in-chrome MCP는 `finance.naver.com` 차단으로 사용 불가 — **Playwright Python만 사용한다.**
