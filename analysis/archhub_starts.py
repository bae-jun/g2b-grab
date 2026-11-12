# -*- coding: utf-8 -*-
"""
국토교통부 건축HUB 건축인허가정보로 시·군별 민간 포함 건축 착공 집계
PROBE=1: 주소·파라미터 조합을 몇 번만 시험해 응답 구조와 필요한 호출 수 확인
필요 환경변수: SERVICE_KEY (공공데이터포털 인증키)
"""
import os
import sys
import json
import time
import urllib.parse

import requests

RAW = os.environ.get("SERVICE_KEY", "").strip()
KEY = urllib.parse.unquote(RAW) if "%" in RAW else RAW
HEADERS = {"User-Agent": "Mozilla/5.0"}
BASES = [
    "https://apis.data.go.kr/1613000/ArchPmsHubService",
    "https://apis.data.go.kr/1613000/ArchPmsService_v2",
]
OP = "getApBasisOulnInfo"          # 건축인허가 기본개요
CALLS = [0]

# 행정표준코드 시군구 (경남 18개 시군; 창원은 5개 구)
SIGUNGU = {
    "창원": ["48121", "48123", "48125", "48127", "48129"], "진주": ["48170"],
    "통영": ["48220"], "사천": ["48240"], "김해": ["48250"],
    "밀양": ["48270"], "거제": ["48310"], "양산": ["48330"],
    "의령": ["48720"], "함안": ["48730"], "창녕": ["48740"],
    "고성": ["48820"], "남해": ["48840"], "하동": ["48850"],
    "산청": ["48860"], "함양": ["48870"], "거창": ["48880"],
    "합천": ["48890"],
}


def call(base, params):
    CALLS[0] += 1
    q = {"serviceKey": KEY, "_type": "json", "type": "json",
         "numOfRows": 10, "pageNo": 1, **params}
    try:
        r = requests.get(f"{base}/{OP}", params=q, headers=HEADERS, timeout=60)
    except Exception as e:
        return None, f"접속오류 {e}"
    txt = r.text.strip()
    try:
        js = json.loads(txt)
    except ValueError:
        return None, f"HTTP {r.status_code} 비JSON: {txt[:300]}"
    return js, None


def summarize(js):
    resp = js.get("response", js)
    hdr = resp.get("header", {}) or {}
    body = resp.get("body", {}) or {}
    items = body.get("items", {})
    if isinstance(items, dict):
        items = items.get("item", [])
    if isinstance(items, dict):
        items = [items]
    return hdr, body.get("totalCount"), items or []


def probe():
    tests = [
        ("시군구만", {"sigunguCd": "48310"}),
        ("시군구+빈 법정동", {"sigunguCd": "48310", "bjdongCd": ""}),
        ("시군구+기간", {"sigunguCd": "48310", "startDate": "20251101",
                      "endDate": "20261031"}),
        ("시군구+법정동(거제 고현동 추정)", {"sigunguCd": "48310",
                                    "bjdongCd": "10100"}),
    ]
    for base in BASES:
        print(f"== 주소: {base.split('/')[-1]}")
        for name, p in tests:
            js, err = call(base, p)
            if err:
                print(f"  [{name}] {err}")
                continue
            hdr, total, items = summarize(js)
            if not items:
                print(f"  [{name}] 원문: {json.dumps(js, ensure_ascii=False)[:400]}")
            print(f"  [{name}] 결과코드={hdr.get('resultCode')} "
                  f"메시지={hdr.get('resultMsg')} 전체건수={total} "
                  f"이번 응답 {len(items)}건")
            if items:
                it = items[0]
                print("    필드:", ", ".join(
                    f"{k}={str(v)[:25]}" for k, v in sorted(it.items())))
            time.sleep(0.3)
    print(f"PROBE 종료: 호출 {CALLS[0]}회")


def main():
    if not KEY:
        print("오류: SERVICE_KEY 없음")
        sys.exit(1)
    if os.environ.get("PROBE") == "1":
        probe()
        return
    print("본 집계는 시험 결과 확인 후 추가 예정")


if __name__ == "__main__":
    main()
