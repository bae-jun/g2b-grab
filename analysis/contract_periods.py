# -*- coding: utf-8 -*-
"""
나라장터 계약정보(공사)로 공사금액 구간별 공사기간·법정 지도횟수 분포 계산
PROBE=1: 주소·오퍼레이션 조합을 몇 번만 시험해 응답 구조(착공일·준공일 항목) 확인
필요 환경변수: SERVICE_KEY
"""
import os
import sys
import json
import time
import re
import datetime
import urllib.parse

import requests

RAW = os.environ.get("SERVICE_KEY", "").strip()
KEY = urllib.parse.unquote(RAW) if "%" in RAW else RAW
HEADERS = {"User-Agent": "Mozilla/5.0"}
BASES = [
    "https://apis.data.go.kr/1230000/ao/CntrctInfoService",  # 01은 폐기 확인됨
]
OPS = ["getCntrctInfoListCnstwk", "getCntrctInfoListCnstwkPPSSrch"]
CALLS = [0]


def call(base, op, params):
    CALLS[0] += 1
    q = {"serviceKey": KEY, "type": "json", "numOfRows": 5, "pageNo": 1,
         **params}
    try:
        r = requests.get(f"{base}/{op}", params=q, headers=HEADERS, timeout=60)
    except Exception as e:
        return None, f"접속오류 {e}"
    txt = r.text.strip()
    try:
        return json.loads(txt), None
    except ValueError:
        return None, f"HTTP {r.status_code} 비JSON: {txt[:300]}"


def items_of(js):
    resp = js.get("response", js)
    body = resp.get("body", {}) or {}
    rows = body.get("items", [])
    if isinstance(rows, dict):
        rows = rows.get("item", [])
    if isinstance(rows, dict):
        rows = [rows]
    return body.get("totalCount"), rows or []


def probe():
    today = datetime.date.today()
    y = today - datetime.timedelta(days=1)
    p = {"inqryDiv": 1, "inqryBgnDt": y.strftime("%Y%m%d") + "0000",
         "inqryEndDt": y.strftime("%Y%m%d") + "2359"}
    for base in BASES:
        for op in OPS:
            js, err = call(base, op, p)
            tag = f"{base.split('/')[-1]}/{op}"
            if err:
                print(f"[{tag}] {err}")
                continue
            total, rows = items_of(js)
            if not rows:
                print(f"[{tag}] 0건, 원문: "
                      f"{json.dumps(js, ensure_ascii=False)[:400]}")
                continue
            print(f"[{tag}] 성공: 전체 {total}건")
            print("  필드:", ", ".join(
                f"{k}={str(v)[:25]}" for k, v in sorted(rows[0].items())))
            time.sleep(0.3)
    print(f"PROBE 종료: 호출 {CALLS[0]}회")


GN = ["창원", "진주", "통영", "사천", "김해", "밀양", "거제", "양산", "의령",
      "함안", "창녕", "고성", "남해", "하동", "산청", "함양", "거창", "합천"]
CORE = ("창원", "함안", "의령", "창녕")              # 내서 1차 권역
BANDS = [(0, 1e8, "1억 미만"), (1e8, 2e8, "1억~2억"), (2e8, 5e8, "2억~5억"),
         (5e8, 10e8, "5억~10억"), (10e8, 20e8, "10억~20억"),
         (20e8, 50e8, "20억~50억"), (50e8, 120e8, "50억~120억"),
         (120e8, 1e15, "120억 이상")]
MAX_CALLS = int(os.environ.get("MAX_CALLS", "800"))
ROWS = [999]


def fetch_day(op, day):
    """하루치(등록일 기준) 전 페이지. 상한 도달 시 None"""
    out, page = [], 1
    d = day.strftime("%Y%m%d")
    while True:
        if CALLS[0] >= MAX_CALLS:
            return None
        CALLS[0] += 1
        q = {"serviceKey": KEY, "type": "json", "numOfRows": ROWS[0],
             "pageNo": page, "inqryDiv": 1,
             "inqryBgnDt": d + "0000", "inqryEndDt": d + "2359"}
        try:
            r = requests.get(f"{BASES[0]}/{op}", params=q, headers=HEADERS,
                             timeout=90)
            js = json.loads(r.text)
        except Exception as e:
            print(f"  [오류] {op} {d} p{page}: {e}")
            return out
        if "nkoneps.com.response.ResponseError" in js:
            msg = str(js)[:200]
            if ROWS[0] > 100:          # 999가 거부되면 100으로 낮춰 재시도
                ROWS[0] = 100
                print(f"  numOfRows 999 거부 → 100으로 변경 ({msg})")
                continue
            print(f"  [응답오류] {op} {d}: {msg}")
            return out
        total, rows = items_of(js)
        out.extend(rows)
        if not rows or len(out) >= int(total or 0):
            return out
        page += 1
        time.sleep(0.2)


def num(v):
    try:
        return float(str(v or "0").replace(",", ""))
    except ValueError:
        return 0.0


def region_of(it):
    s = f"{it.get('dminsttList', '')} {it.get('cntrctInsttNm', '')}"
    if "강원" in s:                    # 강원 고성군 제외
        return None
    if "경상남도" not in s and "경남" not in s and not any(g in s for g in GN):
        return None
    for g in GN:
        if g in s:
            return g
    return "경남(시군불명)"


def days_of(it):
    m = re.search(r"총공사\s*(\d+)\s*일", str(it.get("cntrctPrd", "")))
    if m:
        return int(m.group(1))
    m = re.search(r"(\d+)\s*일", str(it.get("cntrctPrd", "")))
    if m:
        return int(m.group(1))
    try:
        a = datetime.date.fromisoformat(str(it.get("cbgnDate"))[:10])
        b = datetime.date.fromisoformat(str(it.get("ttalCcmpltDate"))[:10])
        return (b - a).days + 1
    except ValueError:
        return None


def corp_of(it):
    s = str(it.get("corpList", ""))
    m = re.search(r"\[\d+\^[^\^]*\^[^\^]*\^([^\^\]]+)", s)
    return m.group(1).strip() if m else "(업체 불명)"


def main():
    if not KEY:
        print("오류: SERVICE_KEY 없음")
        sys.exit(1)
    if os.environ.get("PROBE") == "1":
        probe()
        return
    today = datetime.date.today()
    cn, sv, days_done = {}, {}, []
    day = today - datetime.timedelta(days=1)
    while CALLS[0] < MAX_CALLS and len(days_done) < 365:
        a = fetch_day("getCntrctInfoListCnstwk", day)
        if a is None:
            break
        b = fetch_day("getCntrctInfoListServc", day)
        for it in a:
            cn[it.get("untyCntrctNo") or it.get("cntrctRefNo")] = it
        for it in b or []:
            sv[it.get("untyCntrctNo") or it.get("cntrctRefNo")] = it
        if b is None:
            break
        days_done.append(day)
        day -= datetime.timedelta(days=1)
    if not days_done:
        print("수집 실패")
        sys.exit(1)
    print(f"수집기간(계약 등록일): {days_done[-1]} ~ {days_done[0]} "
          f"({len(days_done)}일), 호출 {CALLS[0]}회, "
          f"공사계약 {len(cn)}건, 용역계약 {len(sv)}건")

    # 1) 경남 공사: 금액구간별 공사기간·지도횟수
    stats = {lab: [] for _, _, lab in BANDS}
    core_legal = core_rule = core_n = 0
    for it in cn.values():
        if str(it.get("lngtrmCtnuDivNm", "")).strip() not in ("", "신규"):
            continue                      # 장기계속 2차 이후 계약 제외
        rg = region_of(it)
        if not rg:
            continue
        amt = num(it.get("totCntrctAmt")) or num(it.get("thtmCntrctAmt"))
        dd = days_of(it)
        if not amt or not dd:
            continue
        for lo, hi, lab in BANDS:
            if lo <= amt < hi:
                stats[lab].append((amt, dd))
        if 1e8 <= amt < 120e8:
            legal = dd // 15 if dd >= 30 else 0
            rule = amt / 6.8e7
            if rg in CORE:
                core_n += 1
                core_legal += legal
                core_rule += rule
    print("\n=== 경남 공사계약 금액구간별 (총계약금액, 부가세 포함) ===")
    print("BAND | 구간 | 건수 | 공사기간 중앙값(일) | 1개월 미만 비율 | "
          "법정 지도횟수 평균(기간÷15) | 6,800만원당 1회 평균")
    for _, _, lab in BANDS:
        v = stats[lab]
        if not v:
            print(f"BAND | {lab} | 0 | - | - | - | -")
            continue
        ds = sorted(d for _, d in v)
        med = ds[len(ds) // 2]
        short = sum(1 for d in ds if d < 30) / len(ds)
        legal = sum((d // 15 if d >= 30 else 0) for _, d in v) / len(v)
        rule = sum(a / 6.8e7 for a, _ in v) / len(v)
        print(f"BAND | {lab} | {len(v)} | {med} | {short:.0%} | "
              f"{legal:.1f} | {rule:.1f}")
    if core_n:
        print(f"CORE | 창원·함안·의령·창녕 대상(1억~120억) {core_n}건 | "
              f"법정 횟수 합 {core_legal} (평균 {core_legal / core_n:.1f}) | "
              f"6,800만원 기준 합 {core_rule:.0f} (평균 {core_rule / core_n:.1f})")

    # 2) 기술지도 용역 계약: 업체별
    kw = re.compile(r"기술\s*지도|재해\s*예방")
    by_corp_gn, by_corp_core, n_all = {}, {}, 0
    for it in sv.values():
        nm = str(it.get("cnstwkNm") or it.get("cntrctNm") or "")
        if not kw.search(nm):
            continue
        n_all += 1
        rg = region_of(it)
        if not rg:
            continue
        c = corp_of(it)
        amt = num(it.get("totCntrctAmt")) or num(it.get("thtmCntrctAmt"))
        x = by_corp_gn.setdefault(c, [0, 0.0])
        x[0] += 1
        x[1] += amt
        if rg in CORE:
            y = by_corp_core.setdefault(c, [0, 0.0])
            y[0] += 1
            y[1] += amt
    print(f"\n=== 기술지도·재해예방 용역 계약: 전국 {n_all}건 중 경남 "
          f"{sum(v[0] for v in by_corp_gn.values())}건 ===")
    for title, dct in (("경남", by_corp_gn), ("창원권", by_corp_core)):
        tot = sum(v[0] for v in dct.values()) or 1
        for c, (k, a) in sorted(dct.items(), key=lambda x: -x[1][0])[:15]:
            print(f"CORP | {title} | {c} | {k}건 ({k / tot:.0%}) | "
                  f"{a / 1e4:,.0f}만원")


if __name__ == "__main__":
    main()
