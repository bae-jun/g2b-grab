# -*- coding: utf-8 -*-
"""
부산청 관할(부산·울산·경남) 시·군별 공공 공사 발주량 집계
- 나라장터 입찰공고(공사) 1년치를 받아 공사 위치(시·군)별로 건수·금액 집계
- 건설재해예방 기술지도 대상 규모(추정가격 1억 이상 120억 미만)만 별도 집계
- 결과: 로그에 표 출력 + out/region_demand.csv 저장
필요 환경변수: SERVICE_KEY (공공데이터포털 인증키)
선택: DATE_FROM, DATE_TO (YYYY-MM-DD)
"""
import os
import sys
import csv
import time
import datetime
import urllib.parse
from collections import defaultdict

import requests

BASES = [f"https://apis.data.go.kr{p}" for p in (
    "/1230000/ad/BidPublicInfoService",
    "/1230000/BidPublicInfoService05",
    "/1230000/BidPublicInfoService04",
)]
OP_LIST = "getBidPblancListInfoCnstwkPPSSrch"
OP_BSIS = "getBidPblancListInfoCnstwkBsisAmount"
HEADERS = {"User-Agent": "Mozilla/5.0"}

RGN_CODES = {"부산": "26", "울산": "31", "경남": "48"}
GN_SIGUN = ["창원", "진주", "통영", "사천", "김해", "밀양", "거제", "양산",
            "의령", "함안", "창녕", "고성", "남해", "하동", "산청", "함양",
            "거창", "합천"]
ELEC_WORDS = ("전기", "정보통신", "통신", "소방")

MIN_P, MAX_P = 100_000_000, 12_000_000_000   # 1억 이상 120억 미만


def service_key():
    k = os.environ.get("SERVICE_KEY", "").strip()
    if not k:
        print("오류: SERVICE_KEY 없음")
        sys.exit(1)
    return urllib.parse.unquote(k) if "%" in k else k


KEY = None
BASE_OK = None
DEBUG_LEFT = [8]   # 진단용 원문 출력 횟수 제한
CALLS = [0]        # API 호출 횟수 (일일 한도 보호)
MAX_CALLS = int(os.environ.get("MAX_CALLS", "700"))


def fetch_all(op, params):
    """기간 1구간 전 페이지 수집. 작동하는 base를 기억해 재사용."""
    global BASE_OK
    bases = [BASE_OK] if BASE_OK else BASES
    last = ""
    for base in bases:
        items, page = [], 1
        try:
            while True:
                q = {"serviceKey": KEY, "pageNo": page, "numOfRows": 100,
                     "type": "json", **params}
                for attempt in range(4):
                    try:
                        CALLS[0] += 1
                        r = requests.get(f"{base}/{op}", params=q,
                                         headers=HEADERS, timeout=60)
                        break
                    except Exception as e:
                        last = str(e)
                        time.sleep(2 ** attempt)
                else:
                    raise RuntimeError(last)
                if r.status_code != 200 or r.text.lstrip().startswith("<"):
                    last = f"HTTP {r.status_code}: {r.text[:200]}"
                    break
                js = r.json()
                resp = js.get("response", js)
                hdr = resp.get("header", {}) or {}
                code = str(hdr.get("resultCode", "00"))
                if code not in ("00", "0", "000"):
                    last = f"{code} {hdr.get('resultMsg', '')}"
                    if DEBUG_LEFT[0] > 0:
                        DEBUG_LEFT[0] -= 1
                        print(f"  [응답오류] {base.split('/')[-1]}/{op} "
                              f"{params}: {last} / 원문: {r.text[:300]}")
                    break
                body = resp.get("body", {}) or {}
                rows = body.get("items", [])
                if isinstance(rows, dict):
                    rows = rows.get("item", [])
                if isinstance(rows, dict):
                    rows = [rows]
                if not rows:
                    if page == 1 and DEBUG_LEFT[0] > 0:
                        DEBUG_LEFT[0] -= 1
                        print(f"  [0건] {base.split('/')[-1]}/{op} {params} "
                              f"totalCount={body.get('totalCount')} "
                              f"원문: {r.text[:300]}")
                    BASE_OK = base
                    return items
                items.extend(rows)
                total = int(body.get("totalCount", 0) or 0)
                if len(items) >= total:
                    BASE_OK = base
                    return items
                page += 1
                time.sleep(0.2)
        except Exception as e:
            last = str(e)
    print(f"  [경고] {op} {params.get('inqryBgnDt')} 실패: {last}")
    return []


def chunks(d_from, d_to, days=28):
    cur = d_from
    while cur <= d_to:
        end = min(cur + datetime.timedelta(days=days - 1), d_to)
        yield cur.strftime("%Y%m%d") + "0000", end.strftime("%Y%m%d") + "2359"
        cur = end + datetime.timedelta(days=1)


def num(v):
    s = str(v or "").replace(",", "").strip()
    try:
        return int(float(s))
    except ValueError:
        return 0


def price_of(it):
    for f in ("presmptPrce", "bssamt", "bssAmt", "asignBdgtAmt", "bdgtAmt"):
        p = num(it.get(f))
        if p:
            return p
    return 0


def sigun_of(it):
    """현장 미기재 공고용 보조 판정: 수요/공고기관명에서 지역 추정"""
    s = str(it.get("dminsttNm", "") or "") + str(it.get("ntceInsttNm", "") or "")
    if "부산" in s:
        return "부산", "(현장 미기재)"
    if "울산" in s:
        return "울산", "(현장 미기재)"
    for sg in GN_SIGUN:
        if sg in s:
            return "경남", sg
    if "경남" in s or "경상남도" in s:
        return "경남", "(현장 미기재)"
    return None, None


SIDO = {"부산광역시": "부산", "울산광역시": "울산", "경상남도": "경남"}


def site_of(it):
    """공사현장 지역명 → (광역, 시군구). 예: '경상남도 진주시' → ('경남','진주')"""
    site = str(it.get("cnstrtsiteRgnNm", "") or "").strip()
    for full, short in SIDO.items():
        if site.startswith(full):
            rest = site[len(full):].strip().split(" ")[0]
            if short == "경남":
                for sg in GN_SIGUN:
                    if rest.startswith(sg):
                        return short, sg
                return short, rest or "(시군 미기재)"
            return short, rest or "(구군 미기재)"
    if not site:
        # 현장 미기재 → 수요기관명으로 보조 판정
        return sigun_of(it)
    return None, None


def is_elec(it):
    s = " ".join(str(it.get(f, "") or "") for f in
                 ("mainCnsttyNm", "cnsttyNm", "bidNtceNm"))
    # 공고명에 '전기'가 들어간 건축공사도 있으므로 주공종을 우선 본다
    main = str(it.get("mainCnsttyNm", "") or "")
    if main:
        return any(w in main for w in ELEC_WORDS)
    return any(w in s for w in ("전기공사", "정보통신공사", "소방시설공사"))


def main():
    global KEY
    KEY = service_key()
    today = datetime.date.today()
    d_to = datetime.date.fromisoformat(
        os.environ.get("DATE_TO") or str(today.replace(day=1)
                                         - datetime.timedelta(days=1)))
    d_from = datetime.date.fromisoformat(
        os.environ.get("DATE_FROM") or str(d_to - datetime.timedelta(days=364)))
    print(f"조회기간: {d_from} ~ {d_to}")

    if os.environ.get("PROBE") == "1":
        probe(today)
        return

    print(f"API 호출 상한: {MAX_CALLS}회")
    collect_and_report(d_from, d_to)


def probe(today):
    # 진단: 지역조건 없이 최근 1일 조회가 되는지 먼저 확인
    y = today - datetime.timedelta(days=1)
    test = fetch_all(OP_LIST, {"inqryDiv": 1,
                               "inqryBgnDt": y.strftime("%Y%m%d") + "0000",
                               "inqryEndDt": today.strftime("%Y%m%d") + "2359"})
    print(f"[진단] 지역조건 없이 최근 1일 공사공고: {len(test)}건 "
          f"(사용 주소: {BASE_OK})")
    if test:
        s = test[0]
        print("[진단] 응답 필드:", ", ".join(
            f"{k}={str(v)[:20]}" for k, v in sorted(s.items())))
    if True:
        # 지역조건 조회에서 허용되는 최대 기간 확인 (호출 수 최소화)
        for days in (1, 3, 7, 14):
            b = today - datetime.timedelta(days=days)
            DEBUG_LEFT[0] = 1
            rows = fetch_all(OP_LIST, {
                "inqryDiv": 1, "inqryBgnDt": b.strftime("%Y%m%d") + "0000",
                "inqryEndDt": today.strftime("%Y%m%d") + "2359",
                "prtcptLmtRgnCd": "48"})
            print(f"[진단] 경남 지역조건 {days}일 조회: {len(rows)}건")
        print("PROBE 모드 종료")


def collect_and_report(d_from, d_to):
    # 1) 지역별 공고 수집 (공고번호 기준 최신 차수만)
    #    지역조건 조회는 약 4일까지만 허용되므로 3일 단위로 나눠 조회.
    #    호출 수가 MAX_CALLS를 넘으면 그 시점까지 받은 자료로 집계(기간 표시).
    notices = {}
    got_until = None
    stop = False
    for b, e in chunks(d_from, d_to, days=3):
        for rgn, code in RGN_CODES.items():
            if CALLS[0] >= MAX_CALLS:
                stop = True
                break
            rows = fetch_all(OP_LIST, {"inqryDiv": 1, "inqryBgnDt": b,
                                       "inqryEndDt": e,
                                       "prtcptLmtRgnCd": code})
            for it in rows:
                no = str(it.get("bidNtceNo", ""))
                if not no:
                    continue
                old = notices.get(no)
                if old is None or num(it.get("bidNtceOrd")) >= num(
                        old.get("bidNtceOrd")):
                    notices[no] = it
        if stop:
            break
        got_until = e[:8]
    print(f"수집 완료: 공고 {len(notices)}건, API 호출 {CALLS[0]}회, "
          f"실제 수집기간 {d_from} ~ {got_until}")

    if not notices:
        print("수신된 공고가 없습니다.")
        sys.exit(1)

    # 2) 집계: 공사현장 지역명(cnstrtsiteRgnNm) 기준
    agg = defaultdict(lambda: defaultdict(int))
    for no, it in notices.items():
        kind = str(it.get("ntceKindNm", "") or "")
        if "취소" in kind:
            continue
        rgn, sg = site_of(it)
        if rgn is None:
            continue                     # 부산·울산·경남 밖 현장
        p = price_of(it)
        a = agg[(rgn, sg)]
        a["전체공고"] += 1
        if p == 0:
            a["금액미확인"] += 1
            continue
        if MIN_P <= p < MAX_P:
            if is_elec(it):
                a["전기통신_대상건수"] += 1
                a["전기통신_대상금액"] += p
            else:
                a["건설_대상건수"] += 1
                a["건설_대상금액"] += p

    cols = ["전체공고", "건설_대상건수", "건설_대상금액", "전기통신_대상건수",
            "전기통신_대상금액", "금액미확인"]
    os.makedirs("out", exist_ok=True)
    path = os.path.join("out", "region_demand.csv")
    order = sorted(agg, key=lambda k: -agg[k]["건설_대상건수"])
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["광역", "시군구"] + cols)
        for k in order:
            w.writerow(list(k) + [agg[k][c] for c in cols])

    print(f"\n=== 결과: 수집기간 {d_from} ~ {got_until}, "
          "대상 = 추정가격 1억 이상 120억 미만, 금액은 억원 ===")
    print("광역 | 시군 | 전체공고 | 건설대상 건수 | 건설대상 금액(억) | "
          "전기통신대상 건수 | 금액미확인")
    for k in order:
        a = agg[k]
        print(f"RESULT | {k[0]} | {k[1]} | {a['전체공고']} | "
              f"{a['건설_대상건수']} | {a['건설_대상금액'] / 1e8:,.0f} | "
              f"{a['전기통신_대상건수']} | {a['금액미확인']}")
    print(f"\nCSV 저장: {path}")


if __name__ == "__main__":
    main()
